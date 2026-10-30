#!/usr/bin/env python3
"""
Instagram Public-Post Comment & Reply Exporter

Discovers up to N profile posts/Reels automatically with timeline pagination,
fetches top-level comments and child replies, deduplicates records,
maintains SQLite checkpointing, and exports data to:
1. Online Master: Google Spreadsheet (ID: 1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y)
2. Local Mirror: Multi-tab Excel workbook (exports/instagram_research.xlsx)

Strictly no CSV output. One username = One worksheet tab.
"""

import argparse
import csv
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from sheets_sync import (
    DEFAULT_EXCEL_PATH,
    DEFAULT_SPREADSHEET_ID,
    GoogleAuthMissingError,
    GoogleSheetsManager,
    SyncCoordinator,
)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("instagram_comments")


class _NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Prevents urllib from following redirects on POST requests into infinite redirect loops."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def shortcode_to_media_id(shortcode: str) -> str:
    """Converts an Instagram shortcode (base64 URL safe) to numeric media ID."""
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    media_id = 0
    for char in shortcode:
        if char not in alphabet:
            continue
        media_id = media_id * 64 + alphabet.index(char)
    return str(media_id)


def ensure_operations_configured() -> None:
    """
    Ensures that required GraphQL operations are configured in ~/.api-anything/sites/instagram.json:
    1. getCommentsPageByMediaId: for top-level comment pagination with cursor.
    2. getCommentReplies: for child reply extraction with parent_comment_id.
    """
    site_path = Path(os.path.expanduser("~/.api-anything/sites/instagram.json"))
    if not site_path.exists():
        logger.warning("instagram.json site configuration not found at %s", site_path)
        return

    try:
        with open(site_path, "r", encoding="utf-8") as f:
            site = json.load(f)

        ops = site.get("operations", [])
        has_comments_page = any(o.get("name") == "getCommentsPageByMediaId" for o in ops)
        has_comment_replies = any(o.get("name") == "getCommentReplies" for o in ops)

        if has_comments_page and has_comment_replies:
            return

        orig_op = next((o for o in ops if o.get("name") == "getCommentsByMediaId"), None)
        if not orig_op:
            return

        # Backup before modifying
        bak_path = site_path.with_name(f"instagram.json.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        shutil.copy2(site_path, bak_path)
        logger.info("Backed up instagram.json to %s", bak_path)

        # 1. Add getCommentsPageByMediaId if missing
        if not has_comments_page:
            new_op = json.loads(json.dumps(orig_op))
            new_op["name"] = "getCommentsPageByMediaId"
            new_op["description"] = "Read a page of comments for a media ID with optional cursor"
            new_op["params"] = [
                {"name": "media_id", "type": "string", "required": True, "example": "3939657355522858837"},
                {"name": "after", "type": "string", "required": False},
            ]
            if not any(s.get("param") == "after" for s in new_op.get("slots", [])):
                new_op.setdefault("slots", []).append({"param": "after", "at": ["form:variables", "json:/after"]})

            body = new_op.get("request", {}).get("body", "")
            if body:
                parsed = urllib.parse.parse_qs(body)
                if "variables" in parsed:
                    vars_dict = json.loads(parsed["variables"][0])
                    vars_dict["after"] = None
                    parsed["variables"] = [json.dumps(vars_dict)]
                    flat = {k: v[0] for k, v in parsed.items()}
                    new_op["request"]["body"] = urllib.parse.urlencode(flat)

            new_op["response"] = {
                "format": "json",
                "extract": "data.xdt_api__v1__media__media_id__comments__connection",
            }
            site["operations"].append(new_op)

        # 2. Add getCommentReplies if missing
        if not has_comment_replies:
            rep_op = json.loads(json.dumps(orig_op))
            rep_op["name"] = "getCommentReplies"
            rep_op["description"] = "Read child replies for a comment on a media ID"
            rep_op["match"] = {
                "method": "POST",
                "host": "www.instagram.com",
                "path": "/api/graphql",
                "operationName": "PolarisPostChildCommentsQuery",
            }
            rep_op["params"] = [
                {"name": "media_id", "type": "string", "required": True, "example": "3939657355522858837"},
                {"name": "parent_comment_id", "type": "string", "required": True, "example": "18059639606512216"},
                {"name": "after", "type": "string", "required": False},
            ]
            body_dict = urllib.parse.parse_qs(rep_op["request"]["body"])
            body_dict["doc_id"] = ["28027289793632076"]
            body_dict["fb_api_req_friendly_name"] = ["PolarisPostChildCommentsQuery"]
            vars_dict = {
                "media_id": "3939657355522858837",
                "parent_comment_id": "18059639606512216",
                "after": None,
                "__relay_internal__pv__PolarisIsLoggedInrelayprovider": True,
            }
            body_dict["variables"] = [json.dumps(vars_dict)]
            flat = {k: v[0] for k, v in body_dict.items()}
            rep_op["request"]["body"] = urllib.parse.urlencode(flat)

            slots = [
                {"param": "media_id", "at": ["form:variables", "json:/media_id"]},
                {"param": "parent_comment_id", "at": ["form:variables", "json:/parent_comment_id"]},
                {"param": "after", "at": ["form:variables", "json:/after"]},
            ]
            for s in orig_op.get("slots", []):
                if s.get("ref"):
                    slots.append(s)
            rep_op["slots"] = slots
            rep_op["response"] = {
                "format": "json",
                "extract": "data.xdt_api__v1__media__media_id__comments__parent_comment_id__child_comments__connection",
            }
            site["operations"].append(rep_op)

        with open(site_path, "w", encoding="utf-8") as f:
            json.dump(site, f, indent=2)

        logger.info("Configured operations in %s", site_path)
    except Exception as e:
        logger.error("Failed to ensure operations in instagram.json: %s", e)


class CheckpointManager:
    """Manages SQLite checkpoint database for tracking posts, comments, and replies progress."""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def _init_db(self) -> None:
        with self.conn:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS posts (
                    shortcode TEXT PRIMARY KEY,
                    media_id TEXT,
                    caption TEXT,
                    likes INTEGER,
                    comments_count INTEGER,
                    status TEXT, -- 'pending', 'in_progress', 'completed'
                    comments_collected INTEGER DEFAULT 0,
                    replies_collected INTEGER DEFAULT 0,
                    next_cursor TEXT,
                    has_more INTEGER DEFAULT 1,
                    post_timestamp INTEGER DEFAULT 0,
                    updated_at TEXT
                )
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS comments (
                    media_id TEXT,
                    comment_id TEXT,
                    shortcode TEXT,
                    record_type TEXT DEFAULT 'comment', -- 'comment' or 'reply'
                    parent_comment_id TEXT DEFAULT '',
                    username TEXT,
                    text TEXT,
                    created_at INTEGER,
                    likes INTEGER,
                    replies_count INTEGER DEFAULT 0,
                    replies_fetched INTEGER DEFAULT 0,
                    PRIMARY KEY (media_id, comment_id)
                )
                """
            )

            # Check and run backwards-compatible migrations on existing databases
            cur = self.conn.cursor()
            cur.execute("PRAGMA table_info(comments)")
            cols = {row[1] for row in cur.fetchall()}
            if "record_type" not in cols:
                self.conn.execute("ALTER TABLE comments ADD COLUMN record_type TEXT DEFAULT 'comment'")
            if "parent_comment_id" not in cols:
                self.conn.execute("ALTER TABLE comments ADD COLUMN parent_comment_id TEXT DEFAULT ''")
            if "replies_count" not in cols:
                self.conn.execute("ALTER TABLE comments ADD COLUMN replies_count INTEGER DEFAULT 0")
            if "replies_fetched" not in cols:
                self.conn.execute("ALTER TABLE comments ADD COLUMN replies_fetched INTEGER DEFAULT 0")

            cur.execute("PRAGMA table_info(posts)")
            post_cols = {row[1] for row in cur.fetchall()}
            if "post_timestamp" not in post_cols:
                self.conn.execute("ALTER TABLE posts ADD COLUMN post_timestamp INTEGER DEFAULT 0")
            if "replies_collected" not in post_cols:
                self.conn.execute("ALTER TABLE posts ADD COLUMN replies_collected INTEGER DEFAULT 0")

    def register_posts(self, posts: List[Dict[str, Any]]) -> None:
        """Inserts discovered posts into checkpoint if not already present."""
        now = datetime.now(timezone.utc).isoformat()
        with self.conn:
            for p in posts:
                self.conn.execute(
                    """
                    INSERT OR IGNORE INTO posts 
                    (shortcode, media_id, caption, likes, comments_count, status, comments_collected, replies_collected, next_cursor, has_more, post_timestamp, updated_at)
                    VALUES (?, ?, ?, ?, ?, 'pending', 0, 0, NULL, 1, ?, ?)
                    """,
                    (
                        p["shortcode"],
                        p["media_id"],
                        p.get("caption", ""),
                        p.get("likes", 0),
                        p.get("comments", 0),
                        p.get("timestamp", 0),
                        now,
                    ),
                )

    def get_post_status(self, shortcode: str) -> Optional[sqlite3.Row]:
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM posts WHERE shortcode = ?", (shortcode,))
        return cur.fetchone()

    def update_post_progress(
        self,
        shortcode: str,
        status: str,
        comments_collected: int,
        replies_collected: int,
        next_cursor: Optional[str],
        has_more: bool,
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self.conn:
            self.conn.execute(
                """
                UPDATE posts
                SET status = ?, comments_collected = ?, replies_collected = ?, next_cursor = ?, has_more = ?, updated_at = ?
                WHERE shortcode = ?
                """,
                (status, comments_collected, replies_collected, next_cursor, 1 if has_more else 0, now, shortcode),
            )

    def save_comments(self, comments: List[Dict[str, Any]]) -> int:
        """Saves comments/replies to SQLite DB. Returns the number of newly inserted items."""
        if not comments:
            return 0
        new_count = 0
        with self.conn:
            for c in comments:
                cur = self.conn.execute(
                    """
                    INSERT OR IGNORE INTO comments
                    (media_id, comment_id, shortcode, record_type, parent_comment_id, username, text, created_at, likes, replies_count, replies_fetched)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        c["media_id"],
                        c["comment_id"],
                        c["post_shortcode"],
                        c.get("record_type", "comment"),
                        c.get("parent_comment_id", ""),
                        c.get("comment_username", ""),
                        c.get("comment_text", ""),
                        c.get("comment_created_at"),
                        c.get("comment_likes", 0),
                        c.get("comment_replies_count", 0),
                        c.get("replies_fetched", 0),
                    ),
                )
                if cur.rowcount > 0:
                    new_count += 1
        return new_count

    def get_unfetched_parent_comments(self, media_id: str) -> List[sqlite3.Row]:
        """Returns top-level comments that have child replies not yet fetched."""
        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT * FROM comments 
            WHERE media_id = ? AND record_type = 'comment' AND replies_count > 0 AND replies_fetched = 0
            ORDER BY created_at DESC
            """,
            (media_id,),
        )
        return cur.fetchall()

    def mark_replies_fetched(self, media_id: str, comment_id: str) -> None:
        with self.conn:
            self.conn.execute(
                "UPDATE comments SET replies_fetched = 1 WHERE media_id = ? AND comment_id = ?",
                (media_id, comment_id),
            )

    def get_existing_comment_ids_for_media(self, media_id: str) -> Set[str]:
        cur = self.conn.cursor()
        cur.execute("SELECT comment_id FROM comments WHERE media_id = ?", (media_id,))
        return {row[0] for row in cur.fetchall()}

    def close(self) -> None:
        self.conn.close()


class InstagramCommentExporter:
    """Coordinates post discovery, comment & reply extraction, and Excel/Google Sheets synchronization."""

    def __init__(
        self,
        username: Optional[str] = None,
        posts_csv: Optional[str] = None,
        url: Optional[str] = None,
        shortcode: Optional[str] = None,
        excel_path: str = DEFAULT_EXCEL_PATH,
        spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
        google_creds: Optional[str] = None,
        skip_sheets: bool = False,
        limit_posts: Optional[int] = 50,
        max_comments_per_post: Optional[int] = None,
        max_replies_per_comment: Optional[int] = None,
        resume: bool = False,
        delay: float = 2.0,
        dry_run: bool = False,
        checkpoint_path: Optional[str] = None,
        extract_comments: bool = True,
        extract_replies: bool = True,
        sync_excel: bool = True,
        sync_sheets: bool = True,
        progress_callback: Optional[Any] = None,
        cancel_event: Optional[Any] = None,
    ):
        self.username = username
        self.posts_csv = posts_csv
        self.url = url
        self.shortcode = shortcode
        self.extract_comments = extract_comments
        self.extract_replies = extract_replies
        self.sync_excel = sync_excel
        self.sync_sheets = sync_sheets and (not skip_sheets)
        self.progress_callback = progress_callback
        self.cancel_event = cancel_event

        # Extract shortcode from URL if given
        if self.url and not self.shortcode:
            m = re.search(r"/(?:p|reel|tv)/([A-Za-z0-9_-]+)", self.url)
            if m:
                self.shortcode = m.group(1)
            else:
                raise ValueError(f"Could not extract shortcode from URL: {self.url}")

        self.excel_path = Path(excel_path)
        self.spreadsheet_id = spreadsheet_id
        self.google_creds = google_creds
        self.skip_sheets = skip_sheets

        self.limit_posts = limit_posts if limit_posts is not None else 50
        self.max_comments_per_post = max_comments_per_post
        self.max_replies_per_comment = max_replies_per_comment
        self.resume = resume
        self.delay = max(0.5, delay)
        self.dry_run = dry_run

        # Checkpoint path
        if checkpoint_path:
            self.checkpoint_db_path = Path(checkpoint_path)
        else:
            name = self.shortcode or self.username or (Path(posts_csv).stem if posts_csv else "comments")
            self.checkpoint_db_path = Path(f".checkpoints/{name}_checkpoint.db")

        self.checkpoint = CheckpointManager(self.checkpoint_db_path)
        self._session_cache: Optional[Dict[str, Any]] = None
        self._site_cache: Optional[Dict[str, Any]] = None
        self._live_tokens: Optional[Dict[str, str]] = None

    def _notify_progress(self, stage: str, **kwargs) -> None:
        """Emits progress event to callback if registered."""
        if self.progress_callback:
            data = {"stage": stage, "username": self.username, **kwargs}
            try:
                self.progress_callback(data)
            except Exception as e:
                logger.debug("Progress callback error: %s", e)

    def _get_live_tokens(self, force_refresh: bool = False) -> Dict[str, str]:
        """Extracts and caches active DTSG, LSD, spin, and revision tokens from Instagram."""
        if self._live_tokens and not force_refresh:
            return self._live_tokens

        if not self._session_cache:
            session_path = Path(os.path.expanduser("~/.api-anything/sessions/instagram.json"))
            if session_path.exists():
                with open(session_path, "r", encoding="utf-8") as f:
                    self._session_cache = json.load(f)
            else:
                self._session_cache = {}

        cookies = self._session_cache.get("cookies", [])
        cookie_dict = {c["name"]: c["value"] for c in cookies if "instagram.com" in c.get("domain", "")}
        cookie_str = "; ".join(f"{k}={v}" for k, v in cookie_dict.items())

        tokens: Dict[str, str] = {}
        if cookie_str:
            req = urllib.request.Request("https://www.instagram.com/")
            req.add_header(
                "User-Agent",
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
            )
            req.add_header("Cookie", cookie_str)
            req.add_header("Accept-Language", "en-US,en;q=0.9")

            opener = urllib.request.build_opener(_NoRedirectHandler)
            try:
                with opener.open(req) as resp:
                    html = resp.read().decode("utf-8", errors="ignore")
            except Exception as e:
                logger.warning("Could not fetch live tokens from Instagram homepage: %s", e)
                html = ""

            dtsg_m = re.search(r'\"DTSGInitialData\"[^\}]*\"token\":\"([^\"]+)\"', html)
            if not dtsg_m:
                dtsg_m = re.search(r'\"token\":\"([^\"]+)\"', html)
            if dtsg_m:
                tokens["dtsg"] = dtsg_m.group(1)
                tokens["jazoest"] = "2" + str(sum(ord(c) for c in tokens["dtsg"]))
                if ":" in tokens["dtsg"]:
                    tokens["av"] = tokens["dtsg"].split(":")[1]

            lsd_m = re.search(r'\"LSD\"[^\}]*\"token\":\"([^\"]+)\"', html)
            if lsd_m:
                tokens["lsd"] = lsd_m.group(1)

            spin_r_m = re.search(r'\"__spin_r\":([0-9]+)', html)
            if spin_r_m:
                tokens["spin_r"] = spin_r_m.group(1)

            spin_t_m = re.search(r'\"__spin_t\":([0-9]+)', html)
            if spin_t_m:
                tokens["spin_t"] = spin_t_m.group(1)

            hsi_m = re.search(r'\"hsi\":\"([0-9]+)\"', html)
            if hsi_m:
                tokens["hsi"] = hsi_m.group(1)

            rev_m = re.search(r'\"server_revision\":([0-9]+)', html)
            if rev_m:
                tokens["rev"] = rev_m.group(1)

        self._live_tokens = tokens
        return tokens

    def _get_credentials_and_headers(self, op_name: str) -> Tuple[Dict[str, str], Dict[str, str], str]:
        """Loads session cookies and form template values for an operation."""
        if not self._session_cache:
            session_path = Path(os.path.expanduser("~/.api-anything/sessions/instagram.json"))
            with open(session_path, "r", encoding="utf-8") as f:
                self._session_cache = json.load(f)

        if not self._site_cache:
            site_path = Path(os.path.expanduser("~/.api-anything/sites/instagram.json"))
            with open(site_path, "r", encoding="utf-8") as f:
                self._site_cache = json.load(f)

        op = next((o for o in self._site_cache["operations"] if o["name"] == op_name), None)
        if not op:
            raise RuntimeError(f"Operation '{op_name}' not found in instagram.json")

        cookies = self._session_cache.get("cookies", [])
        cookie_dict = {c["name"]: c["value"] for c in cookies if "instagram.com" in c.get("domain", "")}
        cookie_str = "; ".join(f"{k}={v}" for k, v in cookie_dict.items())

        headers = dict(op["request"]["headers"])
        headers["cookie"] = cookie_str

        vals = self._session_cache.get("values", {})
        for s in op.get("slots", []):
            ref = s.get("ref")
            if ref:
                name = ref.split(":", 1)[1]
                val = vals.get(name)
                if val:
                    at = s["at"][0]
                    if at.startswith("header:"):
                        headers[at.split(":", 1)[1]] = val

        body_dict = urllib.parse.parse_qs(op["request"].get("body", ""))
        flat_body = {k: v[0] for k, v in body_dict.items()}
        for s in op.get("slots", []):
            ref = s.get("ref")
            if ref:
                name = ref.split(":", 1)[1]
                val = vals.get(name)
                if val:
                    at = s["at"][0]
                    if at.startswith("form:"):
                        flat_body[at.split(":", 1)[1]] = val

        # Synchronize live session tokens to ensure valid GraphQL calls
        live_tokens = self._get_live_tokens()
        if live_tokens.get("dtsg"):
            flat_body["fb_dtsg"] = live_tokens["dtsg"]
            flat_body["jazoest"] = live_tokens["jazoest"]
        if live_tokens.get("lsd"):
            flat_body["lsd"] = live_tokens["lsd"]
            headers["x-fb-lsd"] = live_tokens["lsd"]
        if live_tokens.get("av"):
            flat_body["av"] = live_tokens["av"]
            flat_body["__user"] = "0"
        if live_tokens.get("spin_r"):
            flat_body["__spin_r"] = live_tokens["spin_r"]
        if live_tokens.get("spin_t"):
            flat_body["__spin_t"] = live_tokens["spin_t"]
        if live_tokens.get("hsi"):
            flat_body["__hsi"] = live_tokens["hsi"]
        if live_tokens.get("rev"):
            flat_body["__rev"] = live_tokens["rev"]

        if "csrftoken" in cookie_dict:
            headers["x-csrftoken"] = cookie_dict["csrftoken"]

        headers["user-agent"] = (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
        )
        for k in list(headers.keys()):
            if k.lower() == "accept-encoding":
                del headers[k]

        return headers, flat_body, op["request"]["url"]

    def _query_graphql(
        self, op_name: str, variables: Dict[str, Any], retry_on_token_error: bool = True
    ) -> Dict[str, Any]:
        """Directly queries Instagram's GraphQL endpoint using authenticated live session headers."""
        headers, flat_body, url = self._get_credentials_and_headers(op_name)
        flat_body["variables"] = json.dumps(variables)
        post_data = urllib.parse.urlencode(flat_body).encode("utf-8")
        req = urllib.request.Request(url, data=post_data, headers=headers)
        opener = urllib.request.build_opener(_NoRedirectHandler)

        try:
            with opener.open(req) as resp:
                raw_bytes = resp.read()
                clean_bytes = raw_bytes.strip().removeprefix(b"for (;;);")
                res = json.loads(clean_bytes.decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (301, 302, 303, 307, 308):
                loc = e.headers.get("Location", "")
                logger.error("Instagram redirected to '%s' (Authentication checkpoint or session expired).", loc)
                raise RuntimeError(
                    "Instagram session expired or rate-limited: Instagram redirected the request to a login checkpoint. "
                    "Please refresh your Instagram login session using 'api-anything login instagram --window'."
                )
            elif e.code == 429:
                logger.error("HTTP 429 Too Many Requests encountered during GraphQL query.")
                raise RuntimeError("Instagram rate limit encountered (HTTP 429). Please wait a few minutes before trying again.")
            elif e.code in (401, 403):
                logger.error("HTTP %d Auth challenge encountered during GraphQL query.", e.code)
                raise RuntimeError(f"Instagram authorization failed (HTTP {e.code}). Please refresh your Instagram session.")
            raise

        if res.get("error") == 1357004 and retry_on_token_error:
            logger.info("Encountered token mismatch (1357004). Refreshing live session tokens and retrying...")
            self._get_live_tokens(force_refresh=True)
            return self._query_graphql(op_name, variables, retry_on_token_error=False)

        if "error" in res and not res.get("data"):
            err_code = res.get("error")
            err_summary = res.get("errorSummary", "")
            err_desc = res.get("errorDescription", "")
            if err_code == 1357001 or "log in" in err_summary.lower() or "log in" in err_desc.lower():
                raise RuntimeError(
                    "Instagram session expired or invalid ('Log in to continue'). "
                    "Please refresh your Instagram login session using 'api-anything login instagram --window'."
                )
            if err_code == 1357004:
                raise RuntimeError(
                    f"Instagram security token error ({err_summary}: {err_desc}). "
                    "Please refresh your Instagram login session using 'api-anything login instagram --window'."
                )
            raise RuntimeError(f"Instagram returned error {err_code}: {err_summary} - {err_desc}")

        if "errors" in res and not res.get("data"):
            err_msgs = [err.get("message", "") for err in res.get("errors", [])]
            combined_err = "; ".join(err_msgs)
            logger.error("Instagram GraphQL returned error: %s", combined_err)
            if any("rate limit" in m.lower() for m in err_msgs):
                raise RuntimeError(
                    "Instagram GraphQL rate limit reached ('Rate limit exceeded'). Please wait a few minutes before trying again."
                )
            if any("unauthorized" in m.lower() or "logged out" in m.lower() for m in err_msgs):
                raise RuntimeError(
                    "Instagram session is expired or logged out ('Unauthorized logged out query'). Please refresh your session."
                )
            raise RuntimeError(f"Instagram GraphQL error: {combined_err}")

        return res

    def run_api_anything_call(self, op: str, params: Dict[str, Any], max_retries: int = 3) -> Dict[str, Any]:
        """Invokes api-anything call safely via subprocess with exponential retry on transient network errors."""
        cmd = ["api-anything", "call", "instagram", op, "--json", json.dumps(params)]
        for attempt in range(1, max_retries + 1):
            try:
                res = subprocess.run(cmd, capture_output=True, text=True, check=False)
            except FileNotFoundError:
                raise RuntimeError("api-anything CLI not found in PATH.")

            stdout_clean = res.stdout.strip()
            if not stdout_clean:
                err = res.stderr.strip()
                if attempt < max_retries:
                    wait_s = attempt * 3
                    logger.warning(
                        "Empty response from api-anything on attempt %d/%d for '%s'. Retrying in %ds...",
                        attempt, max_retries, op, wait_s
                    )
                    time.sleep(wait_s)
                    continue
                raise RuntimeError(f"api-anything returned empty output. Stderr: {err}")

            try:
                data = json.loads(stdout_clean)
            except json.JSONDecodeError:
                if attempt < max_retries:
                    wait_s = attempt * 3
                    logger.warning(
                        "JSON parse error on attempt %d/%d for '%s'. Retrying in %ds...",
                        attempt, max_retries, op, wait_s
                    )
                    time.sleep(wait_s)
                    continue
                raise RuntimeError(f"Failed to parse JSON response from api-anything: {stdout_clean[:200]}")

            if not data.get("ok"):
                reason = str(data.get("reason", "")).lower()
                cls = data.get("class")
                # Do not retry auth or rate-limit blocks
                if cls in ("auth", "blocked", "rate"):
                    return data

                # Check for transient network issues (ETIMEDOUT, ECONNRESET, etc.)
                transient_keywords = ["etimedout", "econnreset", "econnrefused", "timeout", "fetch failed", "socket hang up", "network"]
                if any(kw in reason for kw in transient_keywords):
                    if attempt < max_retries:
                        wait_s = attempt * 3
                        logger.warning(
                            "Transient network error (%s) on attempt %d/%d for '%s'. Retrying in %ds...",
                            data.get("reason"), attempt, max_retries, op, wait_s
                        )
                        time.sleep(wait_s)
                        continue

            return data

        return {"ok": False, "reason": "Max retries exceeded"}

    def fetch_profile_posts_page(
        self, username: str, cursor: Optional[str] = None
    ) -> Tuple[List[Dict[str, Any]], Optional[str], bool]:
        """
        Fetches one page of profile posts.
        Page 1 uses PolarisProfilePostsQuery (doc_id: 28542612348729311).
        Page 2+ uses PolarisProfilePostsTabContentQuery_connection (doc_id: 28142289515441884).
        """
        op_name = "getPosts" if cursor is None else "getProfilePostsPage"
        if cursor is None:
            variables = {
                "data": {
                    "count": 12,
                    "include_reel_media_seen_timestamp": True,
                    "include_relationship_info": True,
                    "latest_besties_reel_media": True,
                    "latest_reel_media": True,
                },
                "username": username,
                "__relay_internal__pv__PolarisMultiCaptionCarouselEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisShortDramaEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisReelsRecoDebugOverlayEnabledrelayprovider": False,
            }
        else:
            variables = {
                "after": cursor,
                "first": 12,
                "data": {},
                "username": username,
                "__relay_internal__pv__PolarisMultiCaptionCarouselEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisShortDramaEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisReelsRecoDebugOverlayEnabledrelayprovider": False,
            }

        res = self._query_graphql(op_name, variables)

        data_block = res.get("data") or {}
        conn = data_block.get("xdt_api__v1__feed__user_timeline_graphql_connection") or {}
        edges = conn.get("edges") or []
        page_info = conn.get("page_info") or {}
        has_next_page = bool(page_info.get("has_next_page", False))
        next_cursor = page_info.get("end_cursor")

        extracted_posts = []
        for edge in edges:
            node = edge.get("node", {})
            code = node.get("code")
            if not code:
                continue
            media_id = str(node.get("pk") or shortcode_to_media_id(code))
            caption_obj = node.get("caption")
            caption_text = caption_obj.get("text", "") if isinstance(caption_obj, dict) else (caption_obj or "")
            extracted_posts.append(
                {
                    "shortcode": code,
                    "media_id": media_id,
                    "post_url": f"https://www.instagram.com/p/{code}/",
                    "caption": caption_text,
                    "likes": int(node.get("like_count") or 0),
                    "comments": int(node.get("comment_count") or 0),
                    "timestamp": int(node.get("taken_at") or 0),
                }
            )

        return extracted_posts, next_cursor, has_next_page

    def discover_posts(self) -> List[Dict[str, Any]]:
        """Discovers posts from single URL/shortcode, input CSV, or by paginating profile posts up to limit."""
        posts: List[Dict[str, Any]] = []

        if self.shortcode:
            media_id = shortcode_to_media_id(self.shortcode)
            logger.info("Targeting single post: %s (media_id=%s)", self.shortcode, media_id)
            return [
                {
                    "shortcode": self.shortcode,
                    "media_id": media_id,
                    "post_url": f"https://www.instagram.com/p/{self.shortcode}/",
                    "caption": "",
                    "likes": 0,
                    "comments": 0,
                    "timestamp": 0,
                }
            ]

        if self.posts_csv and Path(self.posts_csv).exists():
            logger.info("Reading input posts from CSV: %s", self.posts_csv)
            with open(self.posts_csv, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    shortcode = row.get("shortcode", "").strip()
                    if not shortcode:
                        continue
                    media_id = row.get("media_id", "").strip() or shortcode_to_media_id(shortcode)
                    posts.append(
                        {
                            "shortcode": shortcode,
                            "media_id": media_id,
                            "post_url": row.get("post_url", f"https://www.instagram.com/p/{shortcode}/"),
                            "caption": row.get("caption", ""),
                            "likes": int(row.get("likes") or 0),
                            "comments": int(row.get("comments") or 0),
                            "timestamp": int(row.get("timestamp") or 0),
                        }
                    )
            logger.info("Loaded %d posts from %s", len(posts), self.posts_csv)
            if self.limit_posts and len(posts) > self.limit_posts:
                posts = posts[: self.limit_posts]
            return posts

        if not self.username:
            raise ValueError("Either --username, --url, --shortcode, or --posts-csv must be provided.")

        logger.info(
            "Discovering profile posts for '%s' (cap: %d posts)...",
            self.username,
            self.limit_posts,
        )

        cursor: Optional[str] = None
        has_more = True
        page_num = 1
        seen_media_ids: Set[str] = set()

        while has_more and len(posts) < self.limit_posts:
            if page_num > 1:
                time.sleep(self.delay)

            page_posts, next_cursor, has_next_page = self.fetch_profile_posts_page(self.username, cursor)
            if not page_posts:
                logger.info("No further posts returned on profile page %d.", page_num)
                break

            new_in_page = 0
            for p in page_posts:
                if p["media_id"] not in seen_media_ids:
                    seen_media_ids.add(p["media_id"])
                    posts.append(p)
                    new_in_page += 1
                    if len(posts) >= self.limit_posts:
                        break

            logger.info(
                "Profile page %d: found %d posts (%d new). Total discovered: %d / %d. has_next_page=%s",
                page_num,
                len(page_posts),
                new_in_page,
                len(posts),
                self.limit_posts,
                has_next_page,
            )

            cursor = next_cursor
            has_more = has_next_page and bool(cursor)
            page_num += 1

        logger.info("Discovered total %d unique posts for username '%s'", len(posts), self.username)
        return posts

    def fetch_comments_page(
        self, media_id: str, after_cursor: Optional[str] = None
    ) -> Tuple[List[Dict[str, Any]], Optional[str], bool]:
        """Fetches one page of top-level comments using direct GraphQL with api-anything fallback."""
        clean_media_id = str(media_id).split("_")[0]
        variables = {
            "media_id": clean_media_id,
            "__relay_internal__pv__PolarisIsLoggedInrelayprovider": True,
            "after": after_cursor,
            "first": 24,
        }

        conn: Dict[str, Any] = {}
        try:
            res = self._query_graphql("getCommentsPageByMediaId", variables)
            data = res.get("data") or {}
            conn = data.get("xdt_api__v1__media__media_id__comments__connection") or {}
        except Exception as e:
            logger.warning("Direct GraphQL comments query failed (%s); falling back to api-anything CLI...", e)
            params: Dict[str, Any] = {"media_id": clean_media_id}
            if after_cursor:
                params["after"] = after_cursor
            res_cli = self.run_api_anything_call("getCommentsPageByMediaId", params)
            if not res_cli.get("ok"):
                cls = res_cli.get("class")
                reason = res_cli.get("reason", "unknown error")
                logger.error("getCommentsPageByMediaId failed: class=%s, reason=%s", cls, reason)
                if cls == "rate":
                    logger.error("Rate limit encountered while fetching comments. Stopping safely.")
                    sys.exit(2)
                elif cls in ("auth", "blocked"):
                    logger.error("Auth wall or checkpoint challenge encountered. Stopping safely.")
                    sys.exit(3)
                raise RuntimeError(f"Comment fetch failed for media {clean_media_id}: {reason}")
            conn = res_cli.get("data", {}) or {}

        if not isinstance(conn, dict):
            return [], None, False

        edges = conn.get("edges", [])
        page_info = conn.get("page_info", {})
        has_next_page = bool(page_info.get("has_next_page", False))
        next_cursor = page_info.get("end_cursor")

        extracted_comments = []
        for edge in edges:
            node = edge.get("node", {})
            if not node:
                continue
            comment_id = str(node.get("pk") or "")
            if not comment_id:
                continue

            user = node.get("user") or {}
            extracted_comments.append(
                {
                    "media_id": clean_media_id,
                    "comment_id": comment_id,
                    "record_type": "comment",
                    "parent_comment_id": "",
                    "comment_username": user.get("username", ""),
                    "comment_text": node.get("text", "") or "",
                    "comment_created_at": node.get("created_at"),
                    "comment_likes": node.get("comment_like_count", 0),
                    "comment_replies_count": node.get("child_comment_count", 0),
                    "replies_fetched": 0,
                }
            )

        return extracted_comments, next_cursor, has_next_page

    def fetch_comment_replies(
        self, media_id: str, parent_comment_id: str, max_replies: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Fetches all accessible child replies for a comment, paginating as supported."""
        all_replies: List[Dict[str, Any]] = []
        cursor: Optional[str] = None
        has_more = True
        reply_page = 1
        clean_media_id = str(media_id).split("_")[0]

        while has_more:
            if max_replies and len(all_replies) >= max_replies:
                break

            time.sleep(self.delay)
            variables = {
                "media_id": clean_media_id,
                "parent_comment_id": str(parent_comment_id),
                "after": cursor,
                "first": 12,
                "__relay_internal__pv__PolarisIsLoggedInrelayprovider": True,
            }

            conn: Dict[str, Any] = {}
            try:
                res = self._query_graphql("getCommentReplies", variables)
                data = res.get("data") or {}
                conn = data.get("xdt_api__v1__media__media_id__comments__parent_comment_id__child_comments__connection") or {}
            except Exception as e:
                logger.warning("Direct GraphQL replies query failed (%s); falling back to api-anything CLI...", e)
                params: Dict[str, Any] = {
                    "media_id": clean_media_id,
                    "parent_comment_id": str(parent_comment_id),
                }
                if cursor:
                    params["after"] = cursor
                res_cli = self.run_api_anything_call("getCommentReplies", params)
                if not res_cli.get("ok"):
                    cls = res_cli.get("class")
                    reason = res_cli.get("reason", "unknown error")
                    logger.error("getCommentReplies failed: class=%s, reason=%s", cls, reason)
                    if cls == "rate":
                        logger.error("Rate limit encountered while fetching replies. Stopping safely.")
                        sys.exit(2)
                    elif cls in ("auth", "blocked"):
                        logger.error("Auth wall encountered while fetching replies. Stopping safely.")
                        sys.exit(3)
                    raise RuntimeError(f"Reply fetch failed for parent {parent_comment_id}: {reason}")
                conn = res_cli.get("data", {}) or {}

            if not isinstance(conn, dict):
                break

            edges = conn.get("edges", [])
            page_info = conn.get("page_info", {})
            has_next_page = bool(page_info.get("has_next_page", False))
            next_cursor = page_info.get("end_cursor")

            for edge in edges:
                node = edge.get("node", {})
                if not node:
                    continue
                reply_id = str(node.get("pk") or "")
                if not reply_id:
                    continue

                user = node.get("user") or {}
                all_replies.append(
                    {
                        "media_id": clean_media_id,
                        "comment_id": reply_id,
                        "record_type": "reply",
                        "parent_comment_id": parent_comment_id,
                        "comment_username": user.get("username", ""),
                        "comment_text": node.get("text", "") or "",
                        "comment_created_at": node.get("created_at"),
                        "comment_likes": node.get("comment_like_count", 0),
                        "comment_replies_count": 0,
                        "replies_fetched": 1,
                    }
                )
                if max_replies and len(all_replies) >= max_replies:
                    break

            cursor = next_cursor
            has_more = has_next_page and bool(cursor)
            reply_page += 1

        return all_replies

    def process_post_comments_and_replies(self, post: Dict[str, Any]) -> Tuple[int, int]:
        """Fetches top-level comments and child replies for a post, paginating and updating checkpoints."""
        shortcode = post["shortcode"]
        media_id = post["media_id"]
        status_row = self.checkpoint.get_post_status(shortcode)

        if self.resume and status_row and status_row["status"] == "completed":
            logger.info("Skipping already completed post: %s (media_id=%s)", shortcode, media_id)
            return status_row["comments_collected"], status_row["replies_collected"]

        comments_collected = status_row["comments_collected"] if (self.resume and status_row) else 0
        replies_collected = status_row["replies_collected"] if (self.resume and status_row) else 0
        cursor = status_row["next_cursor"] if (self.resume and status_row) else None
        has_more_comments = bool(status_row["has_more"]) if (self.resume and status_row) else True

        page_num = 1
        logger.info(
            "Processing post %s (media_id=%s, expected ~%d comments)...",
            shortcode,
            media_id,
            post.get("comments", 0),
        )

        # Check cancellation
        if self.cancel_event and self.cancel_event.is_set():
            return comments_collected, replies_collected

        zero_new_count = 0
        max_pages_per_post = 40

        # 1. Fetch Top-Level Comments (if enabled)
        if not self.extract_comments:
            has_more_comments = False

        while has_more_comments and page_num <= max_pages_per_post:
            if self.cancel_event and self.cancel_event.is_set():
                logger.info("Cancellation detected while fetching comments for %s", shortcode)
                break

            if self.max_comments_per_post and comments_collected >= self.max_comments_per_post:
                logger.info("Reached cap of %d comments for post %s", self.max_comments_per_post, shortcode)
                break

            time.sleep(self.delay)

            try:
                comments, next_cursor, has_next_page = self.fetch_comments_page(media_id, cursor)
            except Exception as e:
                logger.error("Error fetching comments on page %d for post %s: %s", page_num, shortcode, e)
                self.checkpoint.update_post_progress(
                    shortcode, "in_progress", comments_collected, replies_collected, cursor, True
                )
                logger.warning(
                    "Skipping remainder of post %s due to persistent network error; progress is saved in checkpoint.",
                    shortcode
                )
                break

            if not comments:
                logger.info("No comments returned on page %d for post %s.", page_num, shortcode)
                break

            for c in comments:
                c["post_shortcode"] = post["shortcode"]
                c["post_url"] = post["post_url"]
                c["post_caption"] = post.get("caption", "")
                c["post_timestamp"] = post.get("timestamp", 0)
                c["post_likes"] = post.get("likes", 0)
                c["post_comments_count"] = post.get("comments", 0)

            # Save in checkpoint DB
            newly_saved = self.checkpoint.save_comments(comments)
            comments_collected += newly_saved

            logger.info(
                "Post %s: comments page %d fetched %d comments (%d new, total %d collected). has_next_page=%s",
                shortcode,
                page_num,
                len(comments),
                newly_saved,
                comments_collected,
                has_next_page,
            )

            # Loop detection: if Instagram keeps returning comments we've already collected
            if newly_saved == 0:
                zero_new_count += 1
                if zero_new_count >= 2:
                    logger.info(
                        "Post %s: received 2 consecutive pages with 0 new comments (cursor loop detected). Finished comments for this post.",
                        shortcode
                    )
                    break
            else:
                zero_new_count = 0

            cursor = next_cursor
            has_more_comments = has_next_page and bool(cursor)
            page_num += 1

            self.checkpoint.update_post_progress(
                shortcode, "in_progress", comments_collected, replies_collected, cursor, has_more_comments
            )

        # 2. Fetch Child Replies for Comments that have replies (if enabled)
        if self.extract_replies and not (self.cancel_event and self.cancel_event.is_set()):
            unfetched_parents = self.checkpoint.get_unfetched_parent_comments(media_id)
            if unfetched_parents:
                logger.info(
                    "Fetching replies for %d parent comments with replies on post %s...",
                    len(unfetched_parents),
                    shortcode,
                )
                for parent in unfetched_parents:
                    if self.cancel_event and self.cancel_event.is_set():
                        logger.info("Cancellation detected while fetching replies for %s", shortcode)
                        break

                    p_id = parent["comment_id"]
                    expected_reps = parent["replies_count"]
                    logger.info("Fetching replies for parent comment %s (expected ~%d replies)...", p_id, expected_reps)
                    try:
                        replies = self.fetch_comment_replies(
                            media_id, p_id, max_replies=self.max_replies_per_comment
                        )
                    except Exception as e:
                        logger.error("Error fetching replies for parent comment %s on post %s: %s", p_id, shortcode, e)
                        logger.warning("Skipping replies for parent %s; continuing remaining comments...", p_id)
                        continue

                    if replies:
                        for rep in replies:
                            rep["post_shortcode"] = post["shortcode"]
                            rep["post_url"] = post["post_url"]
                            rep["post_caption"] = post.get("caption", "")
                            rep["post_timestamp"] = post.get("timestamp", 0)
                            rep["post_likes"] = post.get("likes", 0)
                            rep["post_comments_count"] = post.get("comments", 0)

                        new_reps = self.checkpoint.save_comments(replies)
                        replies_collected += new_reps
                        logger.info("Parent comment %s: fetched %d replies (%d new).", p_id, len(replies), new_reps)

                    self.checkpoint.mark_replies_fetched(media_id, p_id)
                    self.checkpoint.update_post_progress(
                        shortcode, "in_progress", comments_collected, replies_collected, None, False
                    )

        if not (self.cancel_event and self.cancel_event.is_set()):
            self.checkpoint.update_post_progress(
                shortcode, "completed", comments_collected, replies_collected, None, False
            )
        logger.info(
            "Completed post %s: total %d top-level comments and %d replies collected.",
            shortcode,
            comments_collected,
            replies_collected,
        )
        return comments_collected, replies_collected

    def sync_outputs(self) -> Dict[str, Any]:
        """Synchronizes checkpoint data to the local Excel mirror and Google Spreadsheet master."""
        target_username = self.username or (f"post_{self.shortcode}" if self.shortcode else "instagram_data")
        coord = SyncCoordinator(
            spreadsheet_id=self.spreadsheet_id,
            excel_path=str(self.excel_path),
            credentials_path=self.google_creds,
        )
        return coord.sync(
            self.checkpoint_db_path,
            target_username,
            sync_excel=self.sync_excel,
            sync_sheets=self.sync_sheets,
        )

    def run(self) -> Dict[str, Any]:
        """Executes the complete extraction and synchronization workflow."""
        ensure_operations_configured()

        if self.cancel_event and self.cancel_event.is_set():
            self._notify_progress("cancelled", status_text="Extraction cancelled before start.")
            return {"status": "cancelled", "posts_completed": 0, "total_posts": 0}

        self._notify_progress("discovering_posts", status_text=f"Discovering posts for @{self.username}...")
        posts = self.discover_posts()
        if not posts:
            logger.warning("No posts found to process.")
            self._notify_progress("completed", status_text="No posts found to process.", posts_completed=0, total_posts=0)
            return {"status": "completed", "posts_completed": 0, "total_posts": 0}

        self.checkpoint.register_posts(posts)
        self._notify_progress(
            "posts_discovered",
            total_posts=len(posts),
            posts_discovered=len(posts),
            status_text=f"Discovered {len(posts)} posts for @{self.username}.",
        )

        if self.dry_run:
            print("\n" + "=" * 65)
            print("DRY RUN SUMMARY: Planned Work")
            print("=" * 65)
            print(f"Target Username:         {self.username or 'N/A'}")
            print(f"Posts to process:        {len(posts)} (cap: {self.limit_posts})")
            print(f"Local Excel Mirror:      {self.excel_path}")
            print(f"Google Spreadsheet ID:   {self.spreadsheet_id}")
            print(f"Checkpoint DB path:      {self.checkpoint_db_path}")
            print(f"Delay per request:       {self.delay}s")
            print(f"Max comments per post:   {self.max_comments_per_post or 'Unlimited'}")
            print(f"Max replies per comment: {self.max_replies_per_comment or 'Unlimited'}")
            print("\nDiscovered Posts:")
            for i, p in enumerate(posts, 1):
                st = self.checkpoint.get_post_status(p["shortcode"])
                status_desc = st["status"] if st else "pending"
                print(
                    f"  {i}. [{p['shortcode']}] media_id={p['media_id']} | "
                    f"likes={p['likes']} | comments~={p['comments']} | status={status_desc}"
                )
            print("=" * 65)
            print("Dry run completed. No comments or replies were fetched.")
            self._notify_progress("completed", status_text="Dry run completed.", posts_completed=0, total_posts=len(posts))
            return {"status": "dry_run", "posts_completed": 0, "total_posts": len(posts)}

        total_comments = 0
        total_replies = 0
        posts_completed = 0
        is_cancelled = False
        sync_res: Dict[str, Any] = {}

        try:
            for idx, post in enumerate(posts, start=1):
                if self.cancel_event and self.cancel_event.is_set():
                    logger.info("Cancellation received before post %s. Stopping run.", post["shortcode"])
                    is_cancelled = True
                    break

                self._notify_progress(
                    "fetching_comments",
                    current_post_index=idx,
                    total_posts=len(posts),
                    current_post_shortcode=post["shortcode"],
                    current_post_url=post["post_url"],
                    comments_collected=total_comments,
                    replies_collected=total_replies,
                    status_text=f"Processing post {idx}/{len(posts)}: {post['shortcode']}...",
                )

                c_cnt, r_cnt = self.process_post_comments_and_replies(post)
                total_comments += c_cnt
                total_replies += r_cnt
                posts_completed += 1

                self._notify_progress(
                    "post_completed",
                    current_post_index=idx,
                    total_posts=len(posts),
                    current_post_shortcode=post["shortcode"],
                    comments_collected=total_comments,
                    replies_collected=total_replies,
                    status_text=f"Completed post {idx}/{len(posts)} ({post['shortcode']}). Total records: {total_comments + total_replies}.",
                )

                if self.cancel_event and self.cancel_event.is_set():
                    logger.info("Cancellation received after post %s. Stopping run.", post["shortcode"])
                    is_cancelled = True
                    break
        finally:
            self.checkpoint.close()
            # Perform synchronization to Excel and Google Sheets if enabled
            if self.sync_excel or self.sync_sheets:
                self._notify_progress("syncing", status_text="Synchronizing results to destinations...")
                logger.info("Synchronizing data to Excel mirror and Google Spreadsheet...")
                sync_res = self.sync_outputs()

        final_stage = "cancelled" if is_cancelled else "completed"
        logger.info(
            "Extraction %s! Total: %d comments, %d replies collected across %d posts.",
            final_stage,
            total_comments,
            total_replies,
            posts_completed,
        )

        final_data = {
            "status": final_stage,
            "posts_completed": posts_completed,
            "total_posts": len(posts),
            "comments_collected": total_comments,
            "replies_collected": total_replies,
            "excel_updated": sync_res.get("excel_success", False),
            "excel_path": str(self.excel_path),
            "sheets_synced": sync_res.get("sheets_success", False),
            "sheets_pending_auth": sync_res.get("sheets_pending_auth", False),
            "sheets_message": sync_res.get("sheets_message", ""),
            "spreadsheet_id": self.spreadsheet_id,
            "username": self.username,
        }
        self._notify_progress(
            final_stage,
            **final_data,
            status_text=f"Extraction {final_stage}! Total {total_comments} comments and {total_replies} replies saved across {posts_completed} posts.",
        )
        return final_data


def authenticate_google_cli(credentials_path: Optional[str] = None) -> None:
    """Runs interactive authorization for Google Sheets and tests the connection."""
    mgr = GoogleSheetsManager(spreadsheet_id=DEFAULT_SPREADSHEET_ID)
    try:
        service = mgr.get_service(credentials_path=credentials_path)
        meta = service.spreadsheets().get(spreadsheetId=DEFAULT_SPREADSHEET_ID).execute()
        title = meta.get("properties", {}).get("title", "Unknown")
        sheets = [s["properties"]["title"] for s in meta.get("sheets", [])]
        print("\n" + "=" * 65)
        print("GOOGLE SHEETS CONNECTION VERIFIED!")
        print("=" * 65)
        print(f"Spreadsheet Title: {title}")
        print(f"Spreadsheet ID:    {DEFAULT_SPREADSHEET_ID}")
        print(f"Spreadsheet URL:   https://docs.google.com/spreadsheets/d/{DEFAULT_SPREADSHEET_ID}/edit")
        print(f"Existing Tabs:     {sheets}")
        print("=" * 65 + "\n")
    except Exception as e:
        print("\n[-] Google Sheets authorization failed:")
        print(e)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Export Instagram public-post comments and replies to Google Sheets & Excel with pagination.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--username",
        type=str,
        help="Instagram public username to inspect (e.g. licious_foods, meatigo_official).",
    )
    parser.add_argument(
        "--url",
        type=str,
        help="Direct URL to an Instagram Reel or Post (e.g. https://www.instagram.com/reel/Dbf991LMdQy/).",
    )
    parser.add_argument(
        "--shortcode",
        type=str,
        help="Direct Instagram post shortcode (e.g. Dbf991LMdQy).",
    )
    parser.add_argument(
        "--posts-csv",
        type=str,
        help="Optional input CSV containing posts/shortcodes to process.",
    )
    parser.add_argument(
        "--limit-posts",
        type=int,
        default=50,
        help="Cap on number of posts to discover and process.",
    )
    parser.add_argument(
        "--max-comments-per-post",
        type=int,
        default=None,
        help="Cap on number of top-level comments to collect per post.",
    )
    parser.add_argument(
        "--max-replies-per-comment",
        type=int,
        default=None,
        help="Cap on number of child replies to collect per parent comment.",
    )
    parser.add_argument(
        "--excel-path",
        type=str,
        default=DEFAULT_EXCEL_PATH,
        help="Destination path for the local multi-tab Excel workbook.",
    )
    parser.add_argument(
        "--spreadsheet-id",
        type=str,
        default=DEFAULT_SPREADSHEET_ID,
        help="Google Spreadsheet ID for the master sheet.",
    )
    parser.add_argument(
        "--google-creds",
        type=str,
        default=None,
        help="Path to Google credentials JSON (service_account.json or credentials.json).",
    )
    parser.add_argument(
        "--auth-google",
        action="store_true",
        help="Authorize Google Sheets API interactively and test connection to spreadsheet.",
    )
    parser.add_argument(
        "--sync-only",
        action="store_true",
        help="Sync existing SQLite checkpoint data directly to Excel and Google Sheets without re-fetching.",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Resume an interrupted run using the checkpoint database.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Conservative delay in seconds between pagination requests.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Inspect discovered posts and show planned work without fetching comments.",
    )
    parser.add_argument(
        "--checkpoint",
        type=str,
        default=None,
        help="Custom path to the SQLite checkpoint database.",
    )

    args = parser.parse_args()

    if args.auth_google:
        authenticate_google_cli(credentials_path=args.google_creds)
        return

    if not args.username and not args.posts_csv and not args.url and not args.shortcode:
        args.username = "licious_foods"

    if args.sync_only:
        name = args.shortcode or args.username or (Path(args.posts_csv).stem if args.posts_csv else "comments")
        db_path = Path(args.checkpoint or f".checkpoints/{name}_checkpoint.db")
        if not db_path.exists():
            print(f"[-] Checkpoint DB not found at {db_path}. Run extraction first.")
            sys.exit(1)
        coord = SyncCoordinator(
            spreadsheet_id=args.spreadsheet_id,
            excel_path=args.excel_path,
            credentials_path=args.google_creds,
        )
        coord.sync(db_path, args.username or name)
        return

    exporter = InstagramCommentExporter(
        username=args.username,
        posts_csv=args.posts_csv,
        url=args.url,
        shortcode=args.shortcode,
        excel_path=args.excel_path,
        spreadsheet_id=args.spreadsheet_id,
        google_creds=args.google_creds,
        limit_posts=args.limit_posts,
        max_comments_per_post=args.max_comments_per_post,
        max_replies_per_comment=args.max_replies_per_comment,
        resume=args.resume,
        delay=args.delay,
        dry_run=args.dry_run,
        checkpoint_path=args.checkpoint,
    )

    exporter.run()


if __name__ == "__main__":
    main()
