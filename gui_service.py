#!/usr/bin/env python3
"""
GUI Service Module: Thread-safe Job Manager, Secret Redaction, and History Scanner.

Provides:
- ExtractionJobManager: Runs InstagramCommentExporter in a managed background thread.
- RedactingLogHandler: In-memory ring buffer log capture with automatic secrets redaction.
- Checkpoint & history discovery for recent accounts.
- Google Sheets connectivity testing and on-demand tab re-sync.
"""

import collections
import logging
import os
import re
import sqlite3
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl

from instagram_comments import InstagramCommentExporter
from sheets_sync import (
    DEFAULT_EXCEL_PATH,
    DEFAULT_SPREADSHEET_ID,
    GoogleAuthMissingError,
    GoogleSheetsManager,
    SyncCoordinator,
)

logger = logging.getLogger("instagram_comments.gui")

# Patterns for secret redaction
SECRET_PATTERNS = [
    # Cookies and session values
    (r"(?i)(cookie\s*[:=]\s*)(['\"][^'\"]+['\"]|[^\r\n;]+)", r"\1[REDACTED]"),
    (r"(?i)(sessionid|csrftoken|ds_user_id|mid|ig_did|rur)=[^;\s&]+", r"\1=[REDACTED]"),
    # Bearer tokens and Authorization headers
    (r"(?i)(authorization\s*[:=]\s*Bearer\s+)[A-Za-z0-9._~+/-]+", r"\1[REDACTED]"),
    (r"(?i)(Bearer\s+)[A-Za-z0-9._~+/-]{15,}", r"\1[REDACTED]"),
    # Private keys
    (r"-----BEGIN[ A-Z_-]+KEY-----[^-]+-----END[ A-Z_-]+KEY-----", "[REDACTED_PRIVATE_KEY]"),
    # Client secrets and passwords in URLs or JSON
    (r'(?i)"(client_secret|private_key|token|access_token|secret)":\s*"[^"]+"', r'"\1": "[REDACTED]"'),
]


def redact_secrets(text: str) -> str:
    """Removes sensitive credentials, cookies, and tokens from log output."""
    if not text:
        return ""
    clean = text
    for pattern, replacement in SECRET_PATTERNS:
        clean = re.sub(pattern, replacement, clean)
    return clean


class RedactingLogHandler(logging.Handler):
    """Custom logging handler that buffers redacted log records into a deque."""

    def __init__(self, log_deque: collections.deque):
        super().__init__()
        self.log_deque = log_deque
        self.formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
        )

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            safe_msg = redact_secrets(msg)
            self.log_deque.append({
                "timestamp": record.created,
                "level": record.levelname,
                "message": safe_msg,
            })
        except Exception:
            self.handleError(record)


class ExtractionJobManager:
    """
    Singleton thread-safe manager for orchestrating background extraction jobs,
    handling cancellation, progress callbacks, and real-time state snapshots.
    """

    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
        self._initialized = True
        self.lock = threading.Lock()
        self.is_running = False
        self.cancel_event: Optional[threading.Event] = None
        self.worker_thread: Optional[threading.Thread] = None

        self.logs: collections.deque = collections.deque(maxlen=400)
        self.log_handler = RedactingLogHandler(self.logs)
        root_logger = logging.getLogger("instagram_comments")
        if self.log_handler not in root_logger.handlers:
            root_logger.addHandler(self.log_handler)

        self.state: Dict[str, Any] = {
            "status": "idle",  # idle, running, completed, cancelled, error
            "stage": "idle",
            "username": "",
            "posts_discovered": 0,
            "current_post_index": 0,
            "total_posts": 0,
            "current_post_shortcode": "",
            "current_post_url": "",
            "comments_collected": 0,
            "replies_collected": 0,
            "posts_completed": 0,
            "start_time": None,
            "end_time": None,
            "elapsed_seconds": 0,
            "status_text": "Ready to extract.",
            "error_message": "",
            "excel_updated": False,
            "excel_path": DEFAULT_EXCEL_PATH,
            "sheets_synced": False,
            "sheets_pending_auth": False,
            "sheets_message": "",
            "spreadsheet_id": DEFAULT_SPREADSHEET_ID,
        }

    def start_job(
        self,
        username: str,
        limit_posts: int = 50,
        extract_comments: bool = True,
        extract_replies: bool = True,
        sync_excel: bool = True,
        sync_sheets: bool = True,
        resume: bool = True,
        delay: float = 2.0,
        max_comments_per_post: Optional[int] = None,
        max_replies_per_comment: Optional[int] = None,
        spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
        excel_path: str = DEFAULT_EXCEL_PATH,
        google_creds: Optional[str] = None,
    ) -> Tuple[bool, str]:
        """Validates configuration and starts extraction in a background worker thread."""
        with self.lock:
            if self.is_running:
                return False, "An extraction job is already in progress. Please wait or cancel it."

            clean_user = username.strip().lstrip("@")
            if not clean_user:
                return False, "Username cannot be empty."
            if not re.match(r"^[A-Za-z0-9._]{1,30}$", clean_user):
                return False, f"Invalid Instagram username '{clean_user}'. Only letters, numbers, periods, and underscores allowed."

            if not (1 <= limit_posts <= 500):
                return False, f"Post limit must be between 1 and 500 (got {limit_posts})."

            self.cancel_event = threading.Event()
            self.is_running = True
            self.logs.clear()

            self.state = {
                "status": "running",
                "stage": "initializing",
                "username": clean_user,
                "posts_discovered": 0,
                "current_post_index": 0,
                "total_posts": limit_posts,
                "current_post_shortcode": "",
                "current_post_url": "",
                "comments_collected": 0,
                "replies_collected": 0,
                "posts_completed": 0,
                "start_time": time.time(),
                "end_time": None,
                "elapsed_seconds": 0,
                "status_text": f"Initializing extraction for @{clean_user}...",
                "error_message": "",
                "excel_updated": False,
                "excel_path": excel_path,
                "sheets_synced": False,
                "sheets_pending_auth": False,
                "sheets_message": "",
                "spreadsheet_id": spreadsheet_id,
            }

            def _worker():
                try:
                    exporter = InstagramCommentExporter(
                        username=clean_user,
                        limit_posts=limit_posts,
                        extract_comments=extract_comments,
                        extract_replies=extract_replies,
                        sync_excel=sync_excel,
                        sync_sheets=sync_sheets,
                        resume=resume,
                        delay=delay,
                        max_comments_per_post=max_comments_per_post,
                        max_replies_per_comment=max_replies_per_comment,
                        spreadsheet_id=spreadsheet_id,
                        excel_path=excel_path,
                        google_creds=google_creds,
                        progress_callback=self._on_progress_update,
                        cancel_event=self.cancel_event,
                    )
                    res = exporter.run()
                    with self.lock:
                        self.is_running = False
                        self.state["end_time"] = time.time()
                        self.state["elapsed_seconds"] = int(self.state["end_time"] - (self.state["start_time"] or time.time()))
                        if res:
                            self.state["status"] = res.get("status", "completed")
                            self.state["posts_completed"] = res.get("posts_completed", self.state["posts_completed"])
                            self.state["comments_collected"] = res.get("comments_collected", self.state["comments_collected"])
                            self.state["replies_collected"] = res.get("replies_collected", self.state["replies_collected"])
                            self.state["excel_updated"] = res.get("excel_updated", False)
                            self.state["sheets_synced"] = res.get("sheets_synced", False)
                            self.state["sheets_pending_auth"] = res.get("sheets_pending_auth", False)
                            self.state["sheets_message"] = res.get("sheets_message", "")
                except Exception as e:
                    logger.error("Job worker crashed with error: %s", e, exc_info=True)
                    with self.lock:
                        self.is_running = False
                        self.state["status"] = "error"
                        self.state["stage"] = "error"
                        self.state["end_time"] = time.time()
                        self.state["elapsed_seconds"] = int(self.state["end_time"] - (self.state["start_time"] or time.time()))
                        self.state["error_message"] = str(e)
                        self.state["status_text"] = f"Error: {e}"

            self.worker_thread = threading.Thread(target=_worker, name=f"Extract-{clean_user}", daemon=True)
            self.worker_thread.start()
            return True, f"Extraction started for @{clean_user}."

    def cancel_job(self) -> Tuple[bool, str]:
        """Requests graceful cancellation of active run."""
        with self.lock:
            if not self.is_running:
                return False, "No extraction job is currently running."
            if self.cancel_event:
                self.cancel_event.set()
                self.state["status_text"] = "Stopping extraction safely (preserving checkpoints)..."
                return True, "Cancellation requested. Finishing current request and saving progress..."
            return False, "Cancel event not found."

    def _on_progress_update(self, data: Dict[str, Any]) -> None:
        """Callback invoked by InstagramCommentExporter on milestones."""
        with self.lock:
            stage = data.get("stage", self.state["stage"])
            self.state["stage"] = stage
            if "status_text" in data:
                self.state["status_text"] = data["status_text"]
            if "total_posts" in data:
                self.state["total_posts"] = data["total_posts"]
            if "posts_discovered" in data:
                self.state["posts_discovered"] = data["posts_discovered"]
            if "current_post_index" in data:
                self.state["current_post_index"] = data["current_post_index"]
            if "current_post_shortcode" in data:
                self.state["current_post_shortcode"] = data["current_post_shortcode"]
            if "current_post_url" in data:
                self.state["current_post_url"] = data["current_post_url"]
            if "comments_collected" in data:
                self.state["comments_collected"] = data["comments_collected"]
            if "replies_collected" in data:
                self.state["replies_collected"] = data["replies_collected"]
            if "posts_completed" in data:
                self.state["posts_completed"] = data["posts_completed"]
            if "excel_updated" in data:
                self.state["excel_updated"] = data["excel_updated"]
            if "sheets_synced" in data:
                self.state["sheets_synced"] = data["sheets_synced"]
            if "sheets_pending_auth" in data:
                self.state["sheets_pending_auth"] = data["sheets_pending_auth"]
            if "sheets_message" in data:
                self.state["sheets_message"] = data["sheets_message"]

            if self.state["start_time"]:
                self.state["elapsed_seconds"] = int(time.time() - self.state["start_time"])

    def get_snapshot(self) -> Dict[str, Any]:
        """Returns a thread-safe snapshot of job state and logs."""
        with self.lock:
            state_copy = dict(self.state)
            if self.is_running and state_copy["start_time"]:
                state_copy["elapsed_seconds"] = int(time.time() - state_copy["start_time"])
            state_copy["is_running"] = self.is_running
            state_copy["logs"] = list(self.logs)
            return state_copy

def get_job_manager() -> ExtractionJobManager:
    """Returns the singleton ExtractionJobManager instance."""
    return ExtractionJobManager()


def get_checkpoint_history() -> List[Dict[str, Any]]:
    """Scans .checkpoints directories across cwd and project directories and returns historical run metrics."""
    search_dirs = [
        Path(".checkpoints"),
        Path(__file__).resolve().parent / ".checkpoints",
        Path(__file__).resolve().parent.parent / ".checkpoints",
    ]
    seen_users = set()
    history = []

    all_dbs = []
    for d in search_dirs:
        if d.exists() and d.is_dir():
            for db_path in d.glob("*_checkpoint.db"):
                all_dbs.append(db_path)

    # Sort newest first
    all_dbs.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for db_path in all_dbs:
        username = db_path.stem.replace("_checkpoint", "")
        if username in seen_users:
            continue
        seen_users.add(username)
        try:
            conn = sqlite3.connect(str(db_path))
            conn.row_factory = sqlite3.Row
            posts_cnt = conn.execute("SELECT count(*) FROM posts").fetchone()[0]
            completed_cnt = conn.execute("SELECT count(*) FROM posts WHERE status = 'completed'").fetchone()[0]
            comments_cnt = conn.execute("SELECT count(*) FROM comments WHERE record_type = 'comment'").fetchone()[0]
            replies_cnt = conn.execute("SELECT count(*) FROM comments WHERE record_type = 'reply'").fetchone()[0]
            last_mod = datetime.fromtimestamp(db_path.stat().st_mtime, timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
            conn.close()

            history.append({
                "username": username,
                "db_path": str(db_path),
                "posts_total": posts_cnt,
                "posts_completed": completed_cnt,
                "comments_count": comments_cnt,
                "replies_count": replies_cnt,
                "last_modified": last_mod,
            })
        except Exception as e:
            logger.debug("Failed reading checkpoint %s: %s", db_path, e)

    return history


def resync_existing_account(
    username: str,
    spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
    excel_path: str = DEFAULT_EXCEL_PATH,
    credentials_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Re-syncs an already extracted account's checkpoint to Excel and Google Sheets."""
    candidate_dbs = [
        Path(f".checkpoints/{username}_checkpoint.db"),
        Path(__file__).resolve().parent / f".checkpoints/{username}_checkpoint.db",
        Path(__file__).resolve().parent.parent / f".checkpoints/{username}_checkpoint.db",
    ]
    db_path = next((p for p in candidate_dbs if p.exists()), None)
    if not db_path:
        raise FileNotFoundError(f"Checkpoint for @{username} not found in any checkpoint directory.")

    coord = SyncCoordinator(
        spreadsheet_id=spreadsheet_id,
        excel_path=excel_path,
        credentials_path=credentials_path,
    )
    return coord.sync(db_path, username, sync_excel=True, sync_sheets=True)


def test_google_sheets_connection(
    spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
    credentials_path: Optional[str] = None,
) -> Tuple[bool, str, List[str]]:
    """Tests Google Sheets API authentication and fetches existing worksheet tabs."""
    try:
        mgr = GoogleSheetsManager(spreadsheet_id=spreadsheet_id)
        service = mgr.get_service(credentials_path=credentials_path)
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        title = meta.get("properties", {}).get("title", "Unknown")
        sheets = [s["properties"]["title"] for s in meta.get("sheets", [])]
        return True, title, sheets
    except GoogleAuthMissingError as e:
        return False, f"Authorization Pending: {e}", []
    except Exception as e:
        return False, f"Connection Failed: {e}", []


# ==============================================================================
# MUMBAI LOCATION CHECKER JOB MANAGER
# ==============================================================================

from mumbai_checker import (
    MumbaiCheckDB,
    MumbaiCheckSync,
    MumbaiLocationChecker,
    get_available_checkpoint_accounts,
    load_commenters_from_checkpoints,
)


def extract_usernames_from_excel(file_or_path: Any, sheet_name: Optional[str] = None) -> List[str]:
    """
    Extracts Instagram usernames from an uploaded or local Excel workbook.
    Scans columns for username-like patterns or specific header names.
    """
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    target_sheet = sheet_name if sheet_name and sheet_name in wb.sheetnames else wb.sheetnames[0]
    ws = wb[target_sheet]

    usernames: List[str] = []
    seen = set()

    for row in ws.iter_rows(values_only=True):
        for cell_val in row:
            if not cell_val:
                continue
            txt = str(cell_val).strip()
            # If formatted as @username or handle
            candidate = txt.lstrip("@").strip()
            if re.match(r"^[A-Za-z0-9._]{2,30}$", candidate):
                # Avoid non-usernames like header names
                if candidate.lower() not in ("username", "author", "user", "comment", "reply", "profile", "type"):
                    if candidate.lower() not in seen:
                        seen.add(candidate.lower())
                        usernames.append(candidate)

    return usernames


def get_commenters_summary(
    source_accounts: Optional[List[str]] = None,
    recheck: bool = False,
) -> Dict[str, Any]:
    """
    Computes discovered interactions, unique accounts, already checked, and remaining.
    """
    db = MumbaiCheckDB()
    commenters = load_commenters_from_checkpoints(source_accounts, exclude_source_brand=True)
    discovered = sum(c.get("total_interactions", 1) for c in commenters.values())
    unique_count = len(commenters)

    completed_set = db.get_completed_usernames() if not recheck else set()
    already_checked = sum(1 for norm_u in commenters if norm_u in completed_set)
    remaining = unique_count - already_checked if not recheck else unique_count

    return {
        "discovered_count": discovered,
        "unique_count": unique_count,
        "already_checked_count": already_checked,
        "remaining_count": max(0, remaining),
        "commenters": commenters,
    }


class MumbaiJobManager:
    """
    Singleton thread-safe manager for orchestrating background batch
    Mumbai Location Checks on commenters & reply authors, progress tracking, and sheet syncing.
    """

    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
        self._initialized = True
        self.lock = threading.Lock()
        self.is_running = False
        self.cancel_event: Optional[threading.Event] = None
        self.worker_thread: Optional[threading.Thread] = None

        self.checker = MumbaiLocationChecker()
        self.db = self.checker.db

        self.state: Dict[str, Any] = {
            "status": "idle",  # idle, running, completed, cancelled, error
            "stage": "idle",
            "source_accounts": [],
            "discovered_count": 0,
            "deduplicated_count": 0,
            "skipped_count": 0,
            "remaining_count": 0,
            "processed_count": 0,
            "total_accounts": 0,
            "total_to_process": 0,
            "current_index": 0,
            "current_username": "",
            "failed_count": 0,
            "likely_mumbai_count": 0,
            "serves_mumbai_count": 0,
            "possibly_mumbai_count": 0,
            "outside_mumbai_count": 0,
            "not_enough_evidence_count": 0,
            "start_time": None,
            "end_time": None,
            "elapsed_seconds": 0,
            "status_text": "Ready to run Mumbai location checks.",
            "error_message": "",
            "excel_updated": False,
            "excel_path": DEFAULT_EXCEL_PATH,
            "sheets_synced": False,
            "sheets_url": "",
            "spreadsheet_id": DEFAULT_SPREADSHEET_ID,
            "results": [],
        }

    def start_batch_check(
        self,
        source_accounts: Optional[List[str]] = None,
        custom_usernames: Optional[List[str]] = None,
        recheck: bool = False,
        resume: bool = True,
        limit: int = 0,
        delay: float = 2.0,
        sync_excel: bool = True,
        sync_sheets: bool = True,
        spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
        excel_path: str = DEFAULT_EXCEL_PATH,
        usernames: Optional[List[str]] = None,  # for backward compatibility
    ) -> Tuple[bool, str]:
        """Starts batch Mumbai location checks in a managed background thread."""
        with self.lock:
            if self.is_running:
                return False, "A Mumbai check batch is already in progress."

            commenters_dict: Dict[str, Dict[str, Any]] = {}

            # Case A: Source accounts selected from scraped checkpoints
            if source_accounts:
                commenters_dict = load_commenters_from_checkpoints(
                    source_accounts=source_accounts,
                    exclude_source_brand=True,
                )
            # Case B: Custom usernames list or legacy usernames argument
            target_list = custom_usernames or usernames
            if not commenters_dict and target_list:
                seen = set()
                for u in target_list:
                    c = str(u).strip().lstrip("@")
                    norm = c.lower()
                    if norm and norm not in seen:
                        seen.add(norm)
                        commenters_dict[norm] = {
                            "checked_username": c,
                            "normalized_username": norm,
                            "source_accounts": ["custom_input"],
                            "comment_count": 1,
                            "reply_count": 0,
                            "total_interactions": 1,
                        }

            # Case C: If neither provided, default to all available checkpoints
            if not commenters_dict and not target_list:
                commenters_dict = load_commenters_from_checkpoints(
                    source_usernames=None,
                    exclude_source_brand=True,
                )

            if not commenters_dict:
                return False, "No valid Instagram accounts or commenters found to audit."

            effective_recheck = recheck or not resume
            discovered_total = sum(c.get("total_interactions", 1) for c in commenters_dict.values())
            deduped_total = len(commenters_dict)

            # Determine initial counts fast via completed_set
            completed_set = self.db.get_completed_usernames() if not effective_recheck else set()
            skipped_init = sum(1 for norm_u in commenters_dict if norm_u in completed_set)
            to_check_count = deduped_total - skipped_init

            total_to_process = min(to_check_count, limit) if limit > 0 else to_check_count

            self.is_running = True
            self.cancel_event = threading.Event()

            self.state.update({
                "status": "running",
                "stage": "starting",
                "source_accounts": source_accounts or [],
                "discovered_count": discovered_total,
                "deduplicated_count": deduped_total,
                "skipped_count": skipped_init,
                "remaining_count": total_to_process,
                "processed_count": 0,
                "total_accounts": deduped_total,
                "total_to_process": total_to_process,
                "current_index": 0,
                "current_username": "",
                "failed_count": 0,
                "likely_mumbai_count": 0,
                "serves_mumbai_count": 0,
                "possibly_mumbai_count": 0,
                "outside_mumbai_count": 0,
                "not_enough_evidence_count": 0,
                "start_time": time.time(),
                "end_time": None,
                "elapsed_seconds": 0,
                "status_text": f"Initializing batch of {total_to_process} commenter account(s)...",
                "error_message": "",
                "excel_updated": False,
                "excel_path": excel_path,
                "sheets_synced": False,
                "sheets_url": "",
                "spreadsheet_id": spreadsheet_id,
                "results": [],
            })

            def worker():
                results: List[Dict[str, Any]] = []
                try:
                    def on_progress(evt: Dict[str, Any]):
                        with self.lock:
                            stg = evt.get("stage")
                            if stg == "checking":
                                self.state["current_index"] = evt["current_index"]
                                self.state["current_username"] = evt["username"]
                                self.state["status_text"] = evt["status_text"]
                                self.state["discovered_count"] = evt.get("discovered_count", self.state["discovered_count"])
                                self.state["deduplicated_count"] = evt.get("deduplicated_count", self.state["deduplicated_count"])
                                self.state["skipped_count"] = evt.get("skipped_count", self.state["skipped_count"])
                                self.state["failed_count"] = evt.get("failed_count", self.state["failed_count"])
                            elif stg == "account_done":
                                self.state["processed_count"] = evt.get("newly_checked_count", self.state["processed_count"] + 1)
                                self.state["status_text"] = evt["status_text"]
                                self.state["failed_count"] = evt.get("failed_count", self.state["failed_count"])
                                cats = evt.get("category_counts", {})
                                if cats:
                                    self.state["likely_mumbai_count"] = cats.get("Likely Mumbai", 0)
                                    self.state["serves_mumbai_count"] = cats.get("Serves Mumbai", 0)
                                    self.state["possibly_mumbai_count"] = cats.get("Possibly Mumbai", 0)
                                    self.state["outside_mumbai_count"] = cats.get("Likely outside Mumbai", 0)
                                    self.state["not_enough_evidence_count"] = cats.get("Not enough evidence", 0)

                    batch_result = self.checker.check_commenter_batch(
                        commenters_dict,
                        recheck=effective_recheck,
                        limit=limit,
                        delay=delay,
                        progress_callback=on_progress,
                        cancel_event=self.cancel_event,
                    )
                    results = batch_result.get("results", [])

                    with self.lock:
                        self.state["results"] = results
                        self.state["discovered_count"] = batch_result.get("discovered_count", self.state["discovered_count"])
                        self.state["deduplicated_count"] = batch_result.get("deduplicated_count", self.state["deduplicated_count"])
                        self.state["skipped_count"] = batch_result.get("skipped_count", self.state["skipped_count"])
                        self.state["processed_count"] = batch_result.get("newly_checked_count", len(results))
                        self.state["failed_count"] = batch_result.get("failed_count", self.state["failed_count"])
                        if self.cancel_event and self.cancel_event.is_set():
                            self.state["status"] = "cancelled"
                            self.state["status_text"] = "Batch stopped by user. Checkpoints preserved."
                        else:
                            self.state["status"] = "completed"
                            self.state["status_text"] = f"Finished checking {len(results)} accounts ({self.state['skipped_count']} skipped)."

                    # Sync to Excel
                    if sync_excel and results:
                        try:
                            MumbaiCheckSync.sync_to_excel(results, excel_path=excel_path)
                            with self.lock:
                                self.state["excel_updated"] = True
                        except Exception as e:
                            logger.error("Failed syncing Mumbai Check to Excel: %s", e)

                    # Sync to Google Sheets
                    if sync_sheets and results:
                        try:
                            res_s = MumbaiCheckSync.sync_to_google_sheets(results, spreadsheet_id=spreadsheet_id)
                            with self.lock:
                                self.state["sheets_synced"] = True
                                self.state["sheets_url"] = res_s.get("spreadsheet_url", "")
                        except Exception as e:
                            logger.error("Failed syncing Mumbai Check to Google Sheets: %s", e)

                except Exception as e:
                    logger.exception("Unexpected error in Mumbai check worker: %s", e)
                    with self.lock:
                        self.state["status"] = "error"
                        self.state["error_message"] = str(e)
                        self.state["status_text"] = f"Error: {e}"
                finally:
                    with self.lock:
                        self.is_running = False
                        self.state["end_time"] = time.time()
                        if self.state["start_time"]:
                            self.state["elapsed_seconds"] = int(self.state["end_time"] - self.state["start_time"])

            self.worker_thread = threading.Thread(target=worker, daemon=True)
            self.worker_thread.start()
            return True, f"Started Mumbai check batch for {total_to_process} commenter(s)."

    def cancel_job(self) -> None:
        """Signals the running worker to gracefully stop."""
        with self.lock:
            if self.is_running and self.cancel_event:
                self.cancel_event.set()
                self.state["status_text"] = "Stopping batch after current account..."

    def get_snapshot(self) -> Dict[str, Any]:
        """Returns a thread-safe snapshot of job state."""
        with self.lock:
            state_copy = dict(self.state)
            if self.is_running and state_copy["start_time"]:
                state_copy["elapsed_seconds"] = int(time.time() - state_copy["start_time"])
            state_copy["is_running"] = self.is_running
            return state_copy


def get_mumbai_manager() -> MumbaiJobManager:
    """Returns the singleton MumbaiJobManager instance."""
    return MumbaiJobManager()


class LeadOutreachManager:
    """
    Thread-safe controller for MeatNDoor Lead Outreach ingestion,
    queue updates, DNC management, authorized messaging, and reporting sync.
    """
    _instance: Optional["LeadOutreachManager"] = None
    _lock = threading.Lock()

    def __new__(cls) -> "LeadOutreachManager":
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
        self.lock = threading.Lock()
        self.is_running = False
        self.worker_thread: Optional[threading.Thread] = None
        self.cancel_event: Optional[threading.Event] = None
        from lead_outreach import LeadOutreachDB, MetaMessagingClient
        self.db = LeadOutreachDB()
        self.meta_client = MetaMessagingClient(db=self.db)
        self.state: Dict[str, Any] = {
            "status": "idle",
            "is_running": False,
            "status_text": "Outreach manager ready.",
            "total_raw_found": 0,
            "leads_ingested_or_updated": 0,
            "new_leads_added": 0,
            "mumbai_eligible": 0,
            "dnc_suppressed": 0,
            "start_time": None,
            "end_time": None,
            "elapsed_seconds": 0,
            "error_message": "",
        }
        self._initialized = True

    def start_ingestion_job(self, source_accounts: Optional[List[str]] = None) -> Tuple[bool, str]:
        with self.lock:
            if self.is_running:
                return False, "An outreach ingestion job is already running."

            self.is_running = True
            self.cancel_event = threading.Event()
            self.state = {
                "status": "running",
                "is_running": True,
                "status_text": "Scanning checkpoint databases for commenters and reply authors...",
                "total_raw_found": 0,
                "leads_ingested_or_updated": 0,
                "new_leads_added": 0,
                "mumbai_eligible": 0,
                "dnc_suppressed": 0,
                "start_time": time.time(),
                "end_time": None,
                "elapsed_seconds": 0,
                "error_message": "",
            }

            def worker():
                try:
                    stats = self.db.ingest_from_checkpoints(source_accounts=source_accounts)
                    with self.lock:
                        self.state["status"] = "completed"
                        self.state["status_text"] = (
                            f"Ingestion complete: {stats['leads_ingested_or_updated']} leads processed "
                            f"({stats['new_leads_added']} new, {stats['mumbai_eligible']} Mumbai-eligible)."
                        )
                        self.state["total_raw_found"] = stats["total_raw_found"]
                        self.state["leads_ingested_or_updated"] = stats["leads_ingested_or_updated"]
                        self.state["new_leads_added"] = stats["new_leads_added"]
                        self.state["mumbai_eligible"] = stats["mumbai_eligible"]
                        self.state["dnc_suppressed"] = stats["dnc_suppressed"]
                except Exception as e:
                    logger.exception("Error in outreach ingestion worker: %s", e)
                    with self.lock:
                        self.state["status"] = "error"
                        self.state["error_message"] = str(e)
                        self.state["status_text"] = f"Ingestion error: {e}"
                finally:
                    with self.lock:
                        self.is_running = False
                        self.state["end_time"] = time.time()
                        if self.state["start_time"]:
                            self.state["elapsed_seconds"] = int(self.state["end_time"] - self.state["start_time"])

            self.worker_thread = threading.Thread(target=worker, daemon=True)
            self.worker_thread.start()
            return True, "Started lead ingestion job."

    def get_snapshot(self) -> Dict[str, Any]:
        with self.lock:
            snap = dict(self.state)
            if self.is_running and snap["start_time"]:
                snap["elapsed_seconds"] = int(time.time() - snap["start_time"])
            snap["is_running"] = self.is_running
            return snap

    def get_leads_paginated(self, **kwargs) -> Tuple[List[Dict[str, Any]], int]:
        return self.db.get_leads_paginated(**kwargs)

    def get_lead_counts_summary(self) -> Dict[str, int]:
        return self.db.get_lead_counts_summary()

    def get_lead(self, username: str) -> Optional[Dict[str, Any]]:
        return self.db.get_lead(username)

    def update_lead_status(self, username: str, new_status: str, notes: str = "", operator: str = "Operator") -> bool:
        return self.db.update_lead_status(username, new_status, notes=notes, operator=operator)

    def edit_lead_draft(self, username: str, new_draft: str, operator: str = "Operator") -> bool:
        return self.db.edit_lead_draft(username, new_draft, operator=operator)

    def add_to_dnc(self, username: str, reason: str = "Operator suppression", operator: str = "Operator") -> bool:
        return self.db.add_to_dnc(username, reason=reason, operator=operator)

    def remove_from_dnc(self, username: str, operator: str = "Operator") -> bool:
        return self.db.remove_from_dnc(username, operator=operator)

    def get_audit_log(self, limit: int = 50) -> List[Dict[str, Any]]:
        return self.db.get_audit_log(limit=limit)

    def execute_messaging_workflow(self, username: str, igsid: Optional[str] = None, mock: bool = False, operator: str = "Operator") -> Dict[str, Any]:
        return self.meta_client.execute_messaging_workflow(username=username, igsid=igsid, mock=mock, operator=operator)

    def sync_to_excel(self, excel_path: str = DEFAULT_EXCEL_PATH) -> int:
        from lead_outreach import LeadOutreachSync
        return LeadOutreachSync.sync_to_excel(excel_path=excel_path, db=self.db)

    def sync_to_google_sheets(self, spreadsheet_id: str = DEFAULT_SPREADSHEET_ID) -> Dict[str, Any]:
        from lead_outreach import LeadOutreachSync
        return LeadOutreachSync.sync_to_google_sheets(spreadsheet_id=spreadsheet_id, db=self.db)


def get_outreach_manager() -> LeadOutreachManager:
    """Returns the singleton LeadOutreachManager instance."""
    return LeadOutreachManager()


class CampaignJobManager:
    """
    Thread-safe controller for n8n Campaign Automation, execution orchestration,
    progress monitoring, and lead review queue actions.
    """
    _instance: Optional["CampaignJobManager"] = None
    _lock = threading.Lock()

    def __new__(cls) -> "CampaignJobManager":
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
        self.lock = threading.Lock()
        self.is_running = False
        self.worker_thread: Optional[threading.Thread] = None
        self.pause_event = threading.Event()
        self.cancel_event = threading.Event()
        from campaign_service import CampaignService
        from n8n_client import N8nClient
        self.service = CampaignService()
        self.n8n_client = N8nClient()
        self.state: Dict[str, Any] = {
            "status": "idle",
            "is_running": False,
            "campaign_id": "",
            "stage": "Ready",
            "start_time": None,
            "elapsed_seconds": 0,
            "error_message": "",
            "n8n_triggered": False,
            "n8n_execution_id": "",
        }
        self._initialized = True

    def start_campaign_job(
        self,
        campaign_id: str,
        name: str,
        objective: str = "Audience Research & Outreach",
        target_location: str = "Mumbai",
        source_accounts: Optional[List[str]] = None,
        max_leads: int = 500,
        min_lead_score: int = 0,
        n8n_webhook_url: Optional[str] = None,
        notes: str = "",
    ) -> Tuple[bool, str]:
        with self.lock:
            if self.is_running:
                return False, f"A campaign execution ({self.state.get('campaign_id')}) is already running."

            clean_id = re.sub(r"[^a-zA-Z0-9_-]", "_", campaign_id.strip())
            if not clean_id:
                return False, "Campaign ID cannot be empty."

            # Ensure campaign exists or create it
            existing = self.service.get_campaign(clean_id)
            if not existing:
                try:
                    self.service.create_campaign(
                        campaign_id=clean_id,
                        name=name,
                        objective=objective,
                        target_location=target_location,
                        source_accounts=source_accounts or [],
                        max_leads=max_leads,
                        min_lead_score=min_lead_score,
                        notes=notes,
                    )
                except Exception as e:
                    return False, f"Failed to initialize campaign: {e}"

            self.is_running = True
            self.pause_event.clear()
            self.cancel_event.clear()

            # Optional n8n custom client override if URL provided
            active_n8n = self.n8n_client
            if n8n_webhook_url and n8n_webhook_url.strip():
                from n8n_client import N8nClient
                active_n8n = N8nClient(webhook_url=n8n_webhook_url.strip())

            self.state = {
                "status": "running",
                "is_running": True,
                "campaign_id": clean_id,
                "stage": "Initiating Workflow...",
                "start_time": time.time(),
                "elapsed_seconds": 0,
                "error_message": "",
                "n8n_triggered": False,
                "n8n_execution_id": "",
            }

            def worker():
                n8n_exec_id = ""
                # Attempt n8n webhook notification if configured
                webhook_target = active_n8n.webhook_url
                if webhook_target:
                    try:
                        n8n_res = active_n8n.trigger_campaign({
                            "campaign_id": clean_id,
                            "name": name,
                            "objective": objective,
                            "target_location": target_location,
                            "source_accounts": source_accounts or [],
                            "max_leads": max_leads,
                            "min_lead_score": min_lead_score,
                            "trigger_action": "start",
                            "notes": notes,
                        })
                        n8n_exec_id = n8n_res.get("execution_id", "")
                        with self.lock:
                            self.state["n8n_triggered"] = True
                            self.state["n8n_execution_id"] = n8n_exec_id
                            self.state["stage"] = "n8n Webhook Accepted. Executing Pipeline..."
                    except Exception as e:
                        logger.warning("n8n webhook notification failed or offline: %s. Continuing with core pipeline.", e)
                        with self.lock:
                            self.state["error_message"] = f"n8n Note: {e}"

                # Run core pipeline with checkpoints
                try:
                    res = self.service.execute_pipeline(
                        campaign_id=clean_id,
                        trigger_source="n8n_workflow" if n8n_exec_id else "streamlit_gui",
                        pause_event=self.pause_event,
                        cancel_event=self.cancel_event,
                    )
                    with self.lock:
                        final_st = res.get("status", "completed")
                        self.state["status"] = final_st
                        self.state["stage"] = f"Finished ({final_st.title()})"
                except Exception as e:
                    logger.exception("Error executing campaign %s: %s", clean_id, e)
                    with self.lock:
                        self.state["status"] = "failed"
                        self.state["error_message"] = str(e)
                        self.state["stage"] = f"Failed: {e}"
                finally:
                    with self.lock:
                        self.is_running = False

            self.worker_thread = threading.Thread(target=worker, daemon=True)
            self.worker_thread.start()
            return True, f"Started execution for campaign '{clean_id}'."

    def pause_campaign_job(self, campaign_id: str) -> bool:
        with self.lock:
            self.pause_event.set()
            self.state["stage"] = "Pausing at next checkpoint..."
            return self.service.update_campaign_status(campaign_id, "paused", current_stage="Paused by Operator")

    def resume_campaign_job(self, campaign_id: str) -> Tuple[bool, str]:
        camp = self.service.get_campaign(campaign_id)
        if not camp:
            return False, f"Campaign '{campaign_id}' not found."
        return self.start_campaign_job(
            campaign_id=camp["campaign_id"],
            name=camp["name"],
            objective=camp["objective"],
            target_location=camp["target_location"],
            source_accounts=camp["source_accounts"],
            max_leads=camp["max_leads"],
            min_lead_score=camp["min_lead_score"],
            notes=camp["notes"],
        )

    def cancel_campaign_job(self, campaign_id: str) -> bool:
        with self.lock:
            self.cancel_event.set()
            self.state["stage"] = "Cancelling..."
            return self.service.update_campaign_status(campaign_id, "cancelled", current_stage="Cancelled by Operator")

    def get_snapshot(self) -> Dict[str, Any]:
        with self.lock:
            snap = dict(self.state)
            if self.is_running and snap["start_time"]:
                snap["elapsed_seconds"] = int(time.time() - snap["start_time"])
            snap["is_running"] = self.is_running
            return snap

    def get_campaign(self, campaign_id: str) -> Optional[Dict[str, Any]]:
        return self.service.get_campaign(campaign_id)

    def list_campaigns(self) -> List[Dict[str, Any]]:
        return self.service.list_campaigns()

    def get_campaign_leads_paginated(self, **kwargs) -> Tuple[List[Dict[str, Any]], int]:
        return self.service.get_campaign_leads_paginated(**kwargs)

    def approve_lead_draft(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        return self.service.approve_lead_draft(campaign_id, username, operator=operator)

    def edit_lead_draft(self, campaign_id: str, username: str, new_draft: str, operator: str = "Operator") -> bool:
        return self.service.edit_lead_draft(campaign_id, username, new_draft, operator=operator)

    def reject_lead_draft(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        return self.service.reject_lead_draft(campaign_id, username, operator=operator)

    def update_lead_contact_status(self, campaign_id: str, username: str, contact_status: str, notes: str = "", operator: str = "Operator") -> bool:
        return self.service.update_lead_contact_status(campaign_id, username, contact_status, notes=notes, operator=operator)

    def add_lead_to_dnc(self, campaign_id: str, username: str, reason: str = "Operator suppression", operator: str = "Operator") -> bool:
        return self.service.add_lead_to_dnc(campaign_id, username, reason=reason, operator=operator)

    def remove_lead_from_campaign(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        return self.service.remove_lead_from_campaign(campaign_id, username, operator=operator)

    def sync_campaign_to_sheets(self, campaign_id: str) -> Dict[str, Any]:
        ex_cnt = self.service.sync_campaign_to_excel(campaign_id)
        sh_res = self.service.sync_campaign_to_google_sheets(campaign_id)
        return {
            "excel_rows": ex_cnt,
            "sheets_res": sh_res,
        }


def get_campaign_manager() -> CampaignJobManager:
    """Returns singleton CampaignJobManager instance."""
    return CampaignJobManager()



