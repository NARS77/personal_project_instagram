#!/usr/bin/env python3
"""
Mumbai Location Checker for Instagram Commenters & Reply Authors.

Identifies and audits Instagram users who commented or replied on posts of specified
source accounts (e.g. @licious_foods, @meatigo_official, etc.).

Features:
- Discovers commenters ('comment') and reply authors ('reply') from local checkpoint databases.
- Normalizes and deduplicates usernames across posts and profiles.
- Skips previously completed location checks (with retry on access errors, or forced --recheck).
- Collects public evidence (bio, display name, public business address, external website, post location tags).
- Classifies into:
  * Likely Mumbai
  * Possibly Mumbai
  * Serves Mumbai
  * Likely outside Mumbai
  * Not enough evidence
- Synchronizes output directly to:
  * Google Spreadsheet tab: 'Mumbai Check' (ID: 1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y)
  * Local Excel workbook tab: 'Mumbai Check' (exports/instagram_research.xlsx)
  * Local SQLite checkpoint database: .checkpoints/mumbai_check_checkpoint.db
- Zero CSV generation.
"""

import argparse
import json
import logging
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from instagram_comments import InstagramCommentExporter
from sheets_sync import (
    DEFAULT_EXCEL_PATH,
    DEFAULT_SPREADSHEET_ID,
    GoogleAuthMissingError,
    GoogleSheetsManager,
    sanitize_sheet_title,
)

logger = logging.getLogger("instagram_comments.mumbai_checker")


def get_canonical_mumbai_db_path() -> Path:
    """Finds or initializes the canonical mumbai_check_checkpoint.db path."""
    candidates = [
        Path(__file__).resolve().parent / ".checkpoints" / "mumbai_check_checkpoint.db",
        Path(".checkpoints") / "mumbai_check_checkpoint.db",
        Path(__file__).resolve().parent.parent / ".checkpoints" / "mumbai_check_checkpoint.db",
    ]
    for p in candidates:
        if p.exists() and p.is_file():
            return p.resolve()
    target = Path(__file__).resolve().parent / ".checkpoints" / "mumbai_check_checkpoint.db"
    target.parent.mkdir(parents=True, exist_ok=True)
    return target.resolve()


CHECKPOINT_DB_PATH = get_canonical_mumbai_db_path()

# Mumbai geographical keywords
MUMBAI_CORE = {
    "mumbai", "bombay", "bombaim", "mmr", "bkc", "bandra-kurla complex",
    "bandra kurla complex", "kamala mills", "lower parel", "malabar hill",
    "nariman point", "cuffe parade", "marine drive", "fort mumbai", "colaba",
    "worli", "dadar", "prabhadevi", "mahalaxmi", "byculla", "tardeo", "girgaon"
}

MUMBAI_SUBURBS = {
    "bandra", "khar", "santacruz", "vile parle", "juhu", "andheri",
    "versova", "lokhandwala", "jogeshwari", "goregaon", "malad", "kandivali",
    "borivali", "dahisar", "kurla", "chembur", "ghatkopar", "vikhroli",
    "kanjurmarg", "kanjur", "bhandup", "mulund", "sion", "matunga", "wadala",
    "powai"
}

MUMBAI_METRO_REGION = {
    "thane", "navi mumbai", "vashi", "belapur", "nerul", "kopar khairane",
    "kharghar", "panvel", "kalwa", "airoli", "mira road", "bhayandar"
}

ALL_MUMBAI_TERMS = MUMBAI_CORE | MUMBAI_SUBURBS | MUMBAI_METRO_REGION

# Prominent non-Mumbai Indian cities & international locations
OUTSIDE_LOCATIONS = {
    "bengaluru": "Bengaluru",
    "bangalore": "Bengaluru",
    "delhi": "Delhi/NCR",
    "new delhi": "New Delhi",
    "gurgaon": "Gurugram",
    "gurugram": "Gurugram",
    "noida": "Noida",
    "ghaziabad": "Ghaziabad",
    "faridabad": "Faridabad",
    "hyderabad": "Hyderabad",
    "secunderabad": "Hyderabad",
    "chennai": "Chennai",
    "madras": "Chennai",
    "kolkata": "Kolkata",
    "calcutta": "Kolkata",
    "pune": "Pune",
    "ahmedabad": "Ahmedabad",
    "jaipur": "Jaipur",
    "chandigarh": "Chandigarh",
    "lucknow": "Lucknow",
    "indore": "Indore",
    "surat": "Surat",
    "kochi": "Kochi",
    "cochin": "Kochi",
    "goa": "Goa",
    "london": "London, UK",
    "new york": "New York, USA",
    "dubai": "Dubai, UAE",
    "singapore": "Singapore",
}

SERVICE_KEYWORDS = [
    "delivering across", "delivering in", "delivery across", "delivery in",
    "ships across", "shipping across", "pan-india", "pan india", "across india",
    "all over india", "nationwide delivery", "available across", "available in",
    "serves", "serving", "branches in", "outlets in", "stores in", "order online"
]


class MumbaiCheckDB:
    """Manages persistent SQLite caching and resume capability for Mumbai checks."""

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = Path(db_path).resolve() if db_path else get_canonical_mumbai_db_path()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(self.db_path), timeout=30.0, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    def get_completed_usernames(self) -> Set[str]:
        """Returns the set of all normalized usernames that have completed checks without errors."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT normalized_username FROM mumbai_checks WHERE has_error = 0"
            ).fetchall()
            return {r[0].lower() for r in rows if r[0]}

    def _init_db(self) -> None:
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS mumbai_checks (
                    normalized_username TEXT PRIMARY KEY,
                    checked_username TEXT NOT NULL,
                    profile_url TEXT NOT NULL,
                    display_name TEXT,
                    stated_location TEXT,
                    evidence_summary TEXT NOT NULL,
                    classification TEXT NOT NULL,
                    confidence TEXT NOT NULL,
                    checked_at TEXT NOT NULL,
                    has_error INTEGER DEFAULT 0,
                    error_message TEXT DEFAULT '',
                    source_accounts TEXT DEFAULT '',
                    raw_evidence_json TEXT
                )
            """)
            # Migration checks for existing database schemas
            cols_info = conn.execute("PRAGMA table_info(mumbai_checks)").fetchall()
            existing_cols = {col[1] for col in cols_info}
            pk_col = next((col[1] for col in cols_info if col[5] == 1), None)
            if pk_col != "normalized_username":
                logger.info("Migrating legacy mumbai_checks table to normalized_username primary key schema.")
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS mumbai_checks_new (
                        normalized_username TEXT PRIMARY KEY,
                        checked_username TEXT NOT NULL,
                        profile_url TEXT NOT NULL,
                        display_name TEXT,
                        stated_location TEXT,
                        evidence_summary TEXT NOT NULL,
                        classification TEXT NOT NULL,
                        confidence TEXT NOT NULL,
                        checked_at TEXT NOT NULL,
                        has_error INTEGER DEFAULT 0,
                        error_message TEXT DEFAULT '',
                        source_accounts TEXT DEFAULT '',
                        raw_evidence_json TEXT
                    )
                """)
                norm_expr = "lower(ltrim(username, '@'))" if "username" in existing_cols else "normalized_username"
                checked_expr = "checked_username" if "checked_username" in existing_cols else ("username" if "username" in existing_cols else "''")
                purl_expr = "profile_url" if "profile_url" in existing_cols else "''"
                disp_expr = "display_name" if "display_name" in existing_cols else "''"
                stat_expr = "stated_location" if "stated_location" in existing_cols else "''"
                evid_expr = "evidence_summary" if "evidence_summary" in existing_cols else "''"
                cls_expr = "classification" if "classification" in existing_cols else "''"
                conf_expr = "confidence" if "confidence" in existing_cols else "''"
                chk_expr = "checked_at" if "checked_at" in existing_cols else "''"
                err_expr = "COALESCE(has_error, 0)" if "has_error" in existing_cols else "0"
                errmsg_expr = "COALESCE(error_message, '')" if "error_message" in existing_cols else "''"
                src_expr = "COALESCE(source_accounts, '')" if "source_accounts" in existing_cols else "''"
                raw_expr = "raw_evidence_json" if "raw_evidence_json" in existing_cols else "''"

                migrate_sql = f"""
                    INSERT OR IGNORE INTO mumbai_checks_new (
                        normalized_username, checked_username, profile_url, display_name, stated_location,
                        evidence_summary, classification, confidence, checked_at, has_error, error_message,
                        source_accounts, raw_evidence_json
                    )
                    SELECT
                        {norm_expr},
                        {checked_expr},
                        {purl_expr},
                        {disp_expr},
                        {stat_expr},
                        {evid_expr},
                        {cls_expr},
                        {conf_expr},
                        {chk_expr},
                        {err_expr},
                        {errmsg_expr},
                        {src_expr},
                        {raw_expr}
                    FROM mumbai_checks
                """
                conn.execute(migrate_sql)
                conn.execute("DROP TABLE mumbai_checks")
                conn.execute("ALTER TABLE mumbai_checks_new RENAME TO mumbai_checks")
            else:
                cols = [col[1] for col in cols_info]
                if "normalized_username" not in cols:
                    conn.execute("ALTER TABLE mumbai_checks ADD COLUMN normalized_username TEXT")
                    conn.execute("UPDATE mumbai_checks SET normalized_username = lower(ltrim(COALESCE(username, checked_username), '@'))")
                if "checked_username" not in cols:
                    conn.execute("ALTER TABLE mumbai_checks ADD COLUMN checked_username TEXT")
                    conn.execute("UPDATE mumbai_checks SET checked_username = COALESCE(username, normalized_username)")
                if "has_error" not in cols:
                    conn.execute("ALTER TABLE mumbai_checks ADD COLUMN has_error INTEGER DEFAULT 0")
                if "error_message" not in cols:
                    conn.execute("ALTER TABLE mumbai_checks ADD COLUMN error_message TEXT DEFAULT ''")
                if "source_accounts" not in cols:
                    conn.execute("ALTER TABLE mumbai_checks ADD COLUMN source_accounts TEXT DEFAULT ''")
                conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_mumbai_checks_norm ON mumbai_checks(normalized_username)")

    def get_check(self, username: str) -> Optional[Dict[str, Any]]:
        clean = username.strip().lstrip("@").lower()
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM mumbai_checks WHERE normalized_username = ? OR lower(checked_username) = ?",
                (clean, clean),
            ).fetchone()
            if row:
                d = dict(row)
                if "username" not in d:
                    d["username"] = d.get("checked_username") or d.get("normalized_username", "")
                return d
        return None

    def should_skip(self, username: str, recheck: bool = False) -> bool:
        """
        Determines if an account was previously completed and should be skipped.
        Allows retrying if recheck=True, or if the prior check ended in an access error.
        """
        if recheck:
            return False
        existing = self.get_check(username)
        if not existing:
            return False
        # If previous check had an access error, allow retry
        if existing.get("has_error", 0) == 1:
            return False
        return True

    def save_check(self, record: Dict[str, Any]) -> None:
        raw_u = record.get("checked_username") or record.get("username", "")
        norm = raw_u.strip().lstrip("@").lower()
        if not norm:
            return

        has_err = 1 if record.get("has_error") or (
            record.get("classification") == "Not enough evidence"
            and any(w in record.get("evidence_summary", "").lower() for w in ["inaccessible", "rate limit", "error", "timed out", "unverified"])
        ) else 0
        err_msg = record.get("error_message", "")

        # RULE: Do not overwrite a prior good result with a subsequent access error
        prior = self.get_check(norm)
        if prior and prior.get("has_error", 0) == 0 and has_err == 1:
            logger.info("Preserving prior successful check for @%s; ignoring new transient error.", norm)
            return

        src_acc = record.get("source_accounts", "")
        if isinstance(src_acc, (list, set)):
            src_acc = ", ".join(sorted(list(src_acc)))

        with self._get_connection() as conn:
            conn.execute("""
                INSERT INTO mumbai_checks (
                    normalized_username, checked_username, profile_url, display_name, stated_location,
                    evidence_summary, classification, confidence, checked_at, has_error, error_message,
                    source_accounts, raw_evidence_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(normalized_username) DO UPDATE SET
                    checked_username = excluded.checked_username,
                    profile_url = excluded.profile_url,
                    display_name = excluded.display_name,
                    stated_location = excluded.stated_location,
                    evidence_summary = excluded.evidence_summary,
                    classification = excluded.classification,
                    confidence = excluded.confidence,
                    checked_at = excluded.checked_at,
                    has_error = excluded.has_error,
                    error_message = excluded.error_message,
                    source_accounts = CASE
                        WHEN excluded.source_accounts != '' AND mumbai_checks.source_accounts != ''
                        THEN mumbai_checks.source_accounts || ', ' || excluded.source_accounts
                        ELSE COALESCE(NULLIF(excluded.source_accounts, ''), mumbai_checks.source_accounts)
                    END,
                    raw_evidence_json = excluded.raw_evidence_json
            """, (
                norm,
                raw_u.strip().lstrip("@"),
                record.get("profile_url", f"https://www.instagram.com/{norm}/"),
                record.get("display_name", ""),
                record.get("stated_location", "—"),
                record.get("evidence_summary", ""),
                record.get("classification", "Not enough evidence"),
                record.get("confidence", "Low"),
                record.get("checked_at", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")),
                has_err,
                err_msg,
                src_acc,
                json.dumps(record.get("raw_evidence", {})),
            ))

    def get_all_checks(self) -> List[Dict[str, Any]]:
        with self._get_connection() as conn:
            rows = conn.execute("SELECT * FROM mumbai_checks ORDER BY checked_at DESC").fetchall()
            results = []
            for r in rows:
                d = dict(r)
                if "username" not in d:
                    d["username"] = d.get("checked_username") or d.get("normalized_username", "")
                results.append(d)
            return results

    def get_checks_count(self, classification: Optional[str] = None, search: Optional[str] = None) -> int:
        query = "SELECT count(*) FROM mumbai_checks WHERE 1=1"
        params: List[Any] = []
        if classification and classification != "All Categories":
            query += " AND classification = ?"
            params.append(classification)
        if search and search.strip():
            sq = f"%{search.strip().lower()}%"
            query += """ AND (
                lower(checked_username) LIKE ? OR
                lower(normalized_username) LIKE ? OR
                lower(COALESCE(source_accounts, '')) LIKE ? OR
                lower(COALESCE(display_name, '')) LIKE ? OR
                lower(COALESCE(evidence_summary, '')) LIKE ? OR
                lower(COALESCE(stated_location, '')) LIKE ?
            )"""
            params.extend([sq, sq, sq, sq, sq, sq])
        with self._get_connection() as conn:
            return conn.execute(query, params).fetchone()[0]

    def get_checks_page(
        self,
        limit: int = 50,
        offset: int = 0,
        classification: Optional[str] = None,
        search: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        query = "SELECT * FROM mumbai_checks WHERE 1=1"
        params: List[Any] = []
        if classification and classification != "All Categories":
            query += " AND classification = ?"
            params.append(classification)
        if search and search.strip():
            sq = f"%{search.strip().lower()}%"
            query += """ AND (
                lower(checked_username) LIKE ? OR
                lower(normalized_username) LIKE ? OR
                lower(COALESCE(source_accounts, '')) LIKE ? OR
                lower(COALESCE(display_name, '')) LIKE ? OR
                lower(COALESCE(evidence_summary, '')) LIKE ? OR
                lower(COALESCE(stated_location, '')) LIKE ?
            )"""
            params.extend([sq, sq, sq, sq, sq, sq])
        query += " ORDER BY checked_at DESC LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        with self._get_connection() as conn:
            rows = conn.execute(query, params).fetchall()
            results = []
            for r in rows:
                d = dict(r)
                if "username" not in d:
                    d["username"] = d.get("checked_username") or d.get("normalized_username", "")
                results.append(d)
            return results

    def get_classification_stats(self) -> Dict[str, int]:
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT classification, count(*) FROM mumbai_checks GROUP BY classification"
            ).fetchall()
            return {row[0]: row[1] for row in rows}


# ==============================================================================
# COMMENTER DISCOVERY HELPERS
# ==============================================================================

def get_available_checkpoint_accounts(checkpoints_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """
    Discovers all extracted Instagram profiles with comments/replies saved in checkpoints.
    Searches .checkpoints, api-anything/.checkpoints, and parent dirs.
    """
    if checkpoints_dir is not None:
        search_dirs = [Path(checkpoints_dir).resolve()]
    else:
        candidate_dirs = [
            Path(".checkpoints"),
            Path(__file__).resolve().parent / ".checkpoints",
            Path(__file__).resolve().parent.parent / ".checkpoints",
        ]
        search_dirs = []
        seen_dirs = set()
        for d in candidate_dirs:
            try:
                res_d = d.resolve()
                if res_d.exists() and res_d.is_dir() and str(res_d) not in seen_dirs:
                    seen_dirs.add(str(res_d))
                    search_dirs.append(res_d)
            except Exception:
                pass

    accounts_by_name: Dict[str, Dict[str, Any]] = {}
    for d in search_dirs:
        if d.exists() and d.is_dir():
            for db_path in d.glob("*_checkpoint.db"):
                name = db_path.stem.replace("_checkpoint", "")
                if name in ("mumbai_check", "test"):
                    continue
                try:
                    conn = sqlite3.connect(str(db_path.resolve()), timeout=15.0)
                    tot_comments = conn.execute("SELECT count(*) FROM comments WHERE record_type = 'comment'").fetchone()[0]
                    tot_replies = conn.execute("SELECT count(*) FROM comments WHERE record_type = 'reply'").fetchone()[0]
                    unique_users = conn.execute(
                        "SELECT count(DISTINCT lower(username)) FROM comments WHERE username IS NOT NULL AND trim(username) != ''"
                    ).fetchone()[0]
                    conn.close()
                    total_rec = tot_comments + tot_replies
                    if total_rec > 0:
                        if name not in accounts_by_name or total_rec > accounts_by_name[name]["total_records"]:
                            accounts_by_name[name] = {
                                "source_username": name,
                                "db_path": str(db_path.resolve()),
                                "comments_count": tot_comments,
                                "replies_count": tot_replies,
                                "total_records": total_rec,
                                "unique_commenters": unique_users,
                                "last_modified": datetime.fromtimestamp(db_path.stat().st_mtime, timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                            }
                except Exception as e:
                    logger.debug("Skipping invalid checkpoint db %s: %s", db_path, e)
    accounts = list(accounts_by_name.values())
    accounts.sort(key=lambda a: a["total_records"], reverse=True)
    return accounts


def load_commenters_from_checkpoints(
    source_usernames: Optional[List[str]] = None,
    exclude_source_brand: bool = True,
    checkpoints_dir: Optional[Path] = None,
    source_accounts: Optional[List[str]] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Extracts all commenter and reply author usernames from the selected checkpoint databases.
    Deduplicates usernames across posts and profiles using case-insensitive normalization.

    Returns:
        dict: normalized_username -> {
            'checked_username': original_display_username,
            'normalized_username': lower_username,
            'source_accounts': list_of_source_profiles,
            'comment_count': count,
            'reply_count': count,
            'total_interactions': count,
        }
    """
    sources_to_filter = source_accounts if source_accounts is not None else source_usernames
    available = get_available_checkpoint_accounts(checkpoints_dir=checkpoints_dir)
    if sources_to_filter and "all" not in [s.lower() for s in sources_to_filter]:
        selected_names = {u.strip().lstrip("@").lower() for u in sources_to_filter}
        target_accounts = [a for a in available if a["source_username"].lower() in selected_names]
    else:
        target_accounts = available

    commenters: Dict[str, Dict[str, Any]] = {}

    for acc in target_accounts:
        source_name = acc["source_username"]
        db_path = acc["db_path"]
        try:
            conn = sqlite3.connect(db_path)
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT username, record_type FROM comments WHERE username IS NOT NULL AND trim(username) != ''"
            ).fetchall()
            conn.close()

            for r in rows:
                raw_u = str(r["username"]).strip().lstrip("@")
                norm = raw_u.lower()
                if not norm:
                    continue
                # Do not check source account as a commenter unless specified
                if exclude_source_brand and norm == source_name.lower():
                    continue

                if norm not in commenters:
                    commenters[norm] = {
                        "checked_username": raw_u,
                        "normalized_username": norm,
                        "source_accounts": set(),
                        "comment_count": 0,
                        "reply_count": 0,
                        "total_interactions": 0,
                    }
                commenters[norm]["source_accounts"].add(source_name)
                if r["record_type"] == "comment":
                    commenters[norm]["comment_count"] += 1
                else:
                    commenters[norm]["reply_count"] += 1
                commenters[norm]["total_interactions"] += 1

        except Exception as e:
            logger.warning("Error reading comments from %s: %s", db_path, e)

    # Convert sets to sorted lists
    for data in commenters.values():
        data["source_accounts"] = sorted(list(data["source_accounts"]))

    return commenters


# ==============================================================================
# MUMBAI LOCATION CHECKER PIPELINE
# ==============================================================================

class MumbaiLocationChecker:
    """Extracts public evidence and classifies Instagram accounts by Mumbai location."""

    def __init__(self, db: Optional[MumbaiCheckDB] = None):
        self.db = db or MumbaiCheckDB()

    def fetch_public_evidence(self, username: str) -> Dict[str, Any]:
        """
        Gathers public metadata for an Instagram account via official GraphQL operations
        and public business links.
        """
        clean_u = username.strip().lstrip("@")
        evidence: Dict[str, Any] = {
            "username": clean_u,
            "profile_url": f"https://www.instagram.com/{clean_u}/",
            "accessible": False,
            "display_name": "",
            "biography": "",
            "address_street": "",
            "city_name": "",
            "zip": "",
            "category": "",
            "external_url": "",
            "bio_links": [],
            "post_locations": [],
            "post_captions": [],
            "website_address_hits": [],
            "website_mumbai_mentions": 0,
            "error": "",
        }

        try:
            exporter = InstagramCommentExporter(username=clean_u)
            headers, flat_body, _ = exporter._get_credentials_and_headers("getPosts")
        except Exception as e:
            evidence["error"] = f"Failed to initialize exporter credentials: {e}"
            return evidence

        # 1. Fetch recent posts & user ID
        try:
            flat_body["doc_id"] = "28542612348729311"
            flat_body["fb_api_req_friendly_name"] = "PolarisProfilePostsQuery"
            flat_body["variables"] = json.dumps({
                "data": {"count": 12},
                "username": clean_u,
                "__relay_internal__pv__PolarisMultiCaptionCarouselEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisShortDramaEnabledrelayprovider": True,
                "__relay_internal__pv__PolarisReelsRecoDebugOverlayEnabledrelayprovider": False,
            })
            post_data = urllib.parse.urlencode(flat_body).encode("utf-8")
            req = urllib.request.Request("https://www.instagram.com/graphql/query", data=post_data, headers=headers)
            with urllib.request.urlopen(req, timeout=12) as resp:
                res1 = json.loads(resp.read().decode("utf-8"))

            conn = res1.get("data", {}).get("xdt_api__v1__feed__user_timeline_graphql_connection", {})
            edges = conn.get("edges", [])

            if not edges:
                err_msg = res1.get("errors", [{}])[0].get("message", "Profile inaccessible or private")
                evidence["error"] = err_msg
                return evidence

            evidence["accessible"] = True
            first_node = edges[0].get("node", {})
            user_obj = first_node.get("user", {})
            user_pk = user_obj.get("pk") or first_node.get("owner_id", {}).get("pk")
            evidence["display_name"] = user_obj.get("full_name") or ""

            # Extract post location tags and captions
            for edge in edges:
                node = edge.get("node", {})
                loc = node.get("location")
                if loc and isinstance(loc, dict) and loc.get("name"):
                    evidence["post_locations"].append(loc.get("name"))
                caption = node.get("caption")
                txt = caption.get("text", "") if isinstance(caption, dict) else str(caption or "")
                if txt:
                    evidence["post_captions"].append(txt)

            # 2. Fetch full profile details (bio, public address, external URL)
            if user_pk:
                flat_body["doc_id"] = "28036671149327607"
                flat_body["fb_api_req_friendly_name"] = "PolarisProfilePageContentQuery"
                flat_body["variables"] = json.dumps({
                    "enable_integrity_filters": True,
                    "id": str(user_pk),
                    "__relay_internal__pv__PolarisCannesGuardianExperienceEnabledrelayprovider": True,
                    "__relay_internal__pv__PolarisCASB976ProfileEnabledrelayprovider": False,
                    "__relay_internal__pv__PolarisWebSchoolsEnabledrelayprovider": False,
                    "__relay_internal__pv__PolarisRepostsConsumptionEnabledrelayprovider": False,
                    "__relay_internal__pv__PolarisShortDramaEnabledrelayprovider": True,
                })
                post_data2 = urllib.parse.urlencode(flat_body).encode("utf-8")
                req2 = urllib.request.Request("https://www.instagram.com/graphql/query", data=post_data2, headers=headers)
                with urllib.request.urlopen(req2, timeout=12) as resp2:
                    res2 = json.loads(resp2.read().decode("utf-8"))

                u2 = res2.get("data", {}).get("user", {})
                if u2:
                    evidence["display_name"] = u2.get("full_name") or evidence["display_name"]
                    evidence["biography"] = u2.get("biography") or ""
                    evidence["address_street"] = u2.get("address_street") or ""
                    evidence["city_name"] = u2.get("city_name") or ""
                    evidence["zip"] = u2.get("zip") or ""
                    evidence["category"] = u2.get("category") or ""
                    evidence["external_url"] = u2.get("external_url") or ""
                    evidence["bio_links"] = [b.get("url") for b in u2.get("bio_links", []) if b.get("url")]

        except Exception as e:
            logger.warning("Error fetching GraphQL profile for @%s: %s", clean_u, e)
            evidence["error"] = str(e)

        # 3. Optional Public Website Check (if external_url is present and not an IG link)
        ext_url = evidence.get("external_url")
        if ext_url and "instagram.com" not in ext_url:
            self._enrich_from_public_website(ext_url, evidence)

        return evidence

    def _enrich_from_public_website(self, url: str, evidence: Dict[str, Any]) -> None:
        """Inspects publicly reachable business homepage/contact text for explicitly published addresses."""
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                html = resp.read(150000).decode("utf-8", errors="ignore")
                text = re.sub(r"<[^>]+>", " ", html).lower()
                text = re.sub(r"\s+", " ", text)

                # Check for explicit address or registered office in Mumbai
                if any(k in text for k in ["registered office", "head office", "corporate office", "contact us", "store locator"]):
                    for term in ALL_MUMBAI_TERMS:
                        if re.search(rf"\b{re.escape(term)}\b", text):
                            evidence["website_mumbai_mentions"] += 1
                            evidence["website_address_hits"].append(term.title())

                # Check for Mumbai postal codes 400xxx
                pincodes = re.findall(r"\b(400\d{3})\b", text)
                if pincodes:
                    evidence["website_address_hits"].extend(pincodes[:2])

        except Exception as e:
            logger.debug("Website enrichment skipped for %s: %s", url, e)

    def classify(self, evidence: Dict[str, Any], source_accounts: str = "") -> Dict[str, Any]:
        """
        Classifies profile into:
        - Likely Mumbai
        - Possibly Mumbai
        - Serves Mumbai
        - Likely outside Mumbai
        - Not enough evidence
        """
        clean_u = evidence.get("username", "")
        if evidence.get("accessible") is False or (evidence.get("error") and not evidence.get("post_locations") and not evidence.get("location_tags") and not evidence.get("biography")):
            return {
                "checked_username": clean_u,
                "profile_url": evidence.get("profile_url", f"https://www.instagram.com/{clean_u}/"),
                "display_name": evidence.get("display_name", "—"),
                "source_accounts": source_accounts,
                "stated_location": "—",
                "evidence_summary": f"Profile inaccessible: {evidence.get('error', 'not accessible or private')}.",
                "classification": "Not enough evidence",
                "confidence": "Low",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 1,
                "error_message": evidence.get("error", "inaccessible"),
                "raw_evidence": evidence,
            }

        city = (evidence.get("city_name") or "").strip()
        street = (evidence.get("address_street") or "").strip()
        zip_code = (evidence.get("zip") or "").strip()
        bio = (evidence.get("biography") or "").lower()
        display_name = evidence.get("display_name") or "—"
        post_locations = evidence.get("post_locations") or evidence.get("location_tags") or []
        post_captions = [c.lower() for c in (evidence.get("post_captions") or evidence.get("captions") or [])]

        # Count Mumbai mentions in post location tags
        mumbai_location_tags = []
        outside_location_tags = []
        for loc in post_locations:
            loc_lower = loc.lower()
            if any(term in loc_lower for term in ALL_MUMBAI_TERMS):
                mumbai_location_tags.append(loc)
            elif any(out in loc_lower for out in OUTSIDE_LOCATIONS):
                outside_location_tags.append(loc)

        # Check explicitly stated public business address
        full_stated_addr = ", ".join(filter(bool, [street, city, zip_code]))
        stated_mumbai_match = any(term in full_stated_addr.lower() for term in ALL_MUMBAI_TERMS)
        stated_outside_match = next((name for k, name in OUTSIDE_LOCATIONS.items() if k in full_stated_addr.lower()), None)

        # Check bio for Mumbai and outside cities
        bio_mumbai_matches = [term for term in ALL_MUMBAI_TERMS if re.search(rf"\b{re.escape(term)}\b", bio)]
        bio_outside_matches = [name for k, name in OUTSIDE_LOCATIONS.items() if re.search(rf"\b{re.escape(k)}\b", bio)]

        # Check for service / delivery keywords
        has_delivery_intent = any(kw in bio for kw in SERVICE_KEYWORDS) or any(
            any(kw in cap for kw in SERVICE_KEYWORDS) for cap in post_captions[:3]
        )

        stated_location = full_stated_addr if full_stated_addr else ("—")

        # -------------------------------------------------------------
        # 1. EVALUATE: LIKELY MUMBAI
        # -------------------------------------------------------------
        # Case A: Public business address explicitly in Mumbai/MMR
        if stated_mumbai_match:
            tag_summary = f"{len(mumbai_location_tags)}/{len(post_locations)} post location tags in Mumbai" if post_locations else "no post location tags"
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": full_stated_addr,
                "evidence_summary": f"Public business address explicitly set to '{full_stated_addr}'; {tag_summary}.",
                "classification": "Likely Mumbai",
                "confidence": "High",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # Case B: Bio explicitly states headquarters/base in Mumbai (e.g. "based in Mumbai", "Mumbai • India")
        bio_base_patterns = [
            r"\bbased in\b[^\.\n]*\b(mumbai|bombay|bandra|andheri|juhu|colaba|thane)\b",
            r"\bhq\s*[:\-]\s*(mumbai|bombay)\b",
            r"\b(mumbai|bombay)\s*based\b",
            r"\bheadquartered in (mumbai|bombay)\b",
            r"\bflagship\b[^\.\n]*\b(mumbai|bandra|juhu|colaba)\b",
            r"📍\s*(mumbai|bombay|bandra|andheri|juhu|colaba|thane|powai)",
        ]
        is_explicitly_based_in_mumbai = any(re.search(pat, bio) for pat in bio_base_patterns)
        if is_explicitly_based_in_mumbai:
            stated_loc = ", ".join(bio_mumbai_matches[:2]).title()
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": stated_loc,
                "evidence_summary": f"Bio explicitly states base/location in Mumbai ({', '.join(bio_mumbai_matches[:2]).title()}); {len(mumbai_location_tags)} post tags in Mumbai.",
                "classification": "Likely Mumbai",
                "confidence": "High",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # -------------------------------------------------------------
        # 2. EVALUATE: SERVES MUMBAI
        # -------------------------------------------------------------
        delivering_to_mumbai_terms = [t for t in bio_mumbai_matches if t in ["mumbai", "thane", "mulund", "powai", "vikhroli", "bhandup", "navi mumbai"]]
        if has_delivery_intent and (delivering_to_mumbai_terms or "across india" in bio or "pan-india" in bio or "pan india" in bio):
            if delivering_to_mumbai_terms:
                detail = f"delivering across {', '.join(delivering_to_mumbai_terms).title()}"
            else:
                detail = "Pan-India delivery network covering Mumbai"
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": stated_location if stated_location != "—" else "Multi-City / Pan-India",
                "evidence_summary": f"Bio specifies delivery/service operation ({detail}).",
                "classification": "Serves Mumbai",
                "confidence": "High" if delivering_to_mumbai_terms else "Medium",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # -------------------------------------------------------------
        # 3. EVALUATE: LIKELY OUTSIDE MUMBAI
        # -------------------------------------------------------------
        # Case A: Public business address in another city
        if stated_outside_match:
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": full_stated_addr,
                "evidence_summary": f"Public address registered in {stated_outside_match} ('{full_stated_addr}'); no primary Mumbai headquarters.",
                "classification": "Likely outside Mumbai",
                "confidence": "High",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # Case B: Bio states base/HQ in another city, or Display Name / Handle specifies another city
        bio_outside_base_patterns = [
            r"\bbased in\b[^\.\n]*\b(bengaluru|bangalore|delhi|gurgaon|gurugram|noida|hyderabad|chennai|pune|kolkata)\b",
            r"\bhq\s*[:\-]\s*(bengaluru|bangalore|delhi|gurgaon|gurugram|noida|hyderabad|chennai|pune|kolkata)\b",
            r"📍\s*(bengaluru|bangalore|delhi|gurgaon|gurugram|noida|hyderabad|chennai|pune|kolkata)",
        ]
        is_outside_base = any(re.search(pat, bio) for pat in bio_outside_base_patterns)
        disp_and_handle = f"{display_name} {clean_u}".lower()
        disp_outside_matches = [name for k, name in OUTSIDE_LOCATIONS.items() if re.search(rf"\b{re.escape(k)}\b", disp_and_handle)]

        if (is_outside_base or disp_outside_matches) and not bio_mumbai_matches:
            outside_name = (bio_outside_matches[0] if bio_outside_matches else disp_outside_matches[0]) if (bio_outside_matches or disp_outside_matches) else "outside city"
            reason = "Bio indicates primary base" if is_outside_base else "Display name or username identifies location"
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": outside_name,
                "evidence_summary": f"{reason} in {outside_name}; no Mumbai primary address or base.",
                "classification": "Likely outside Mumbai",
                "confidence": "High" if is_outside_base else "Medium",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # -------------------------------------------------------------
        # 4. EVALUATE: POSSIBLY MUMBAI vs NOT ENOUGH EVIDENCE
        # -------------------------------------------------------------
        # Rule: "A single location tag must not automatically classify an account as Mumbai-based."
        # If repeated (2+) Mumbai location tags exist:
        if len(mumbai_location_tags) >= 2:
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": stated_location,
                "evidence_summary": f"Multiple post location tags in Mumbai ({len(mumbai_location_tags)} tags: {', '.join(set(mumbai_location_tags))[:50]}); no explicit bio statement.",
                "classification": "Possibly Mumbai",
                "confidence": "Medium",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # Bio mentions Mumbai in passing (1 match) without explicit base or delivery
        if bio_mumbai_matches:
            matched_terms = ", ".join(bio_mumbai_matches).title()
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": matched_terms,
                "evidence_summary": f"Bio contains reference to {matched_terms}, but without explicit base or business address verification.",
                "classification": "Possibly Mumbai",
                "confidence": "Low",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # Check if display name or handle references Mumbai/Bombay
        disp_and_handle = f"{display_name} {clean_u}".lower()
        handle_mumbai_matches = [term for term in ALL_MUMBAI_TERMS if re.search(rf"\b{re.escape(term)}\b", disp_and_handle)]
        if handle_mumbai_matches:
            matched_terms = ", ".join(handle_mumbai_matches[:2]).title()
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": matched_terms,
                "evidence_summary": f"Display name or username references '{matched_terms}'; public address or location tags not explicitly confirmed.",
                "classification": "Possibly Mumbai",
                "confidence": "Low",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # If only a single location tag exists:
        if len(mumbai_location_tags) == 1:
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": stated_location,
                "evidence_summary": f"Only a single isolated post location tag ('{mumbai_location_tags[0]}'); insufficient to verify Mumbai base.",
                "classification": "Not enough evidence",
                "confidence": "Low",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # Predominant outside post tags
        if len(outside_location_tags) >= 2 and not mumbai_location_tags:
            return {
                "checked_username": clean_u,
                "profile_url": evidence["profile_url"],
                "display_name": display_name,
                "source_accounts": source_accounts,
                "stated_location": ", ".join(sorted(list(set(outside_location_tags)))[:2]),
                "evidence_summary": f"Multiple post location tags in other regions ({', '.join(set(outside_location_tags))[:50]}); 0 Mumbai references.",
                "classification": "Likely outside Mumbai",
                "confidence": "Medium",
                "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "has_error": 0,
                "raw_evidence": evidence,
            }

        # -------------------------------------------------------------
        # 5. DEFAULT: NOT ENOUGH EVIDENCE
        # -------------------------------------------------------------
        return {
            "checked_username": clean_u,
            "profile_url": evidence["profile_url"],
            "display_name": display_name,
            "source_accounts": source_accounts,
            "stated_location": "—",
            "evidence_summary": "No public city, address, or repeated location tags found in profile or recent posts.",
            "classification": "Not enough evidence",
            "confidence": "Low",
            "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            "has_error": 0,
            "raw_evidence": evidence,
        }

    def check_account(
        self,
        username: str,
        source_accounts: str = "",
        recheck: bool = False,
    ) -> Dict[str, Any]:
        """Runs check for a single username, consulting cache if recheck=False."""
        clean_u = username.strip().lstrip("@")
        if not recheck and self.db.should_skip(clean_u, recheck=False):
            cached = self.db.get_check(clean_u)
            if cached:
                logger.info("Using cached Mumbai check for @%s (%s)", clean_u, cached["classification"])
                return cached

        evidence = self.fetch_public_evidence(clean_u)
        result = self.classify(evidence, source_accounts=source_accounts)
        self.db.save_check(result)
        return result

    def check_commenter_batch(
        self,
        commenters_dict: Dict[str, Dict[str, Any]],
        recheck: bool = False,
        limit: int = 0,
        delay: float = 2.0,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        cancel_event: Optional[Any] = None,
    ) -> Dict[str, Any]:
        """
        Executes a controlled batch check over extracted commenters:
        1. Identifies unique normalized accounts.
        2. Filters out previously completed accounts (unless recheck=True).
        3. Applies optional limit.
        4. Runs checks with rate-delay and updates database.
        5. Reports comprehensive telemetry.
        """
        discovered_interactions = sum(c.get("total_interactions", 1) for c in commenters_dict.values())
        deduped_total = len(commenters_dict)

        # Filter out already checked accounts
        to_check: List[Dict[str, Any]] = []
        skipped_count = 0

        for norm_u, c_data in commenters_dict.items():
            if not recheck and self.db.should_skip(norm_u, recheck=False):
                skipped_count += 1
            else:
                to_check.append(c_data)

        if limit > 0 and len(to_check) > limit:
            to_check = to_check[:limit]

        total_to_process = len(to_check)
        newly_checked = 0
        failed_count = 0
        results: List[Dict[str, Any]] = []

        category_counts = {
            "Likely Mumbai": 0,
            "Serves Mumbai": 0,
            "Possibly Mumbai": 0,
            "Likely outside Mumbai": 0,
            "Not enough evidence": 0,
        }

        for idx, item in enumerate(to_check, start=1):
            if cancel_event and cancel_event.is_set():
                logger.info("Batch checking cancelled by user.")
                break

            target_user = item["checked_username"]
            src_str = ", ".join(item.get("source_accounts", []))

            if progress_callback:
                progress_callback({
                    "stage": "checking",
                    "current_index": idx,
                    "total": total_to_process,
                    "username": target_user,
                    "status_text": f"Checking @{target_user} ({idx}/{total_to_process})...",
                    "discovered_count": discovered_interactions,
                    "deduplicated_count": deduped_total,
                    "skipped_count": skipped_count,
                    "newly_checked_count": newly_checked,
                    "failed_count": failed_count,
                })

            res = self.check_account(target_user, source_accounts=src_str, recheck=recheck)
            results.append(res)
            newly_checked += 1

            cls_name = res.get("classification", "Not enough evidence")
            if cls_name in category_counts:
                category_counts[cls_name] += 1
            if res.get("has_error", 0) == 1:
                failed_count += 1

            if progress_callback:
                progress_callback({
                    "stage": "account_done",
                    "current_index": idx,
                    "total": total_to_process,
                    "username": target_user,
                    "result": res,
                    "status_text": f"Completed @{target_user}: {cls_name} ({res.get('confidence')} confidence)",
                    "discovered_count": discovered_interactions,
                    "deduplicated_count": deduped_total,
                    "skipped_count": skipped_count,
                    "newly_checked_count": newly_checked,
                    "failed_count": failed_count,
                    "category_counts": category_counts,
                })

            if idx < total_to_process and delay > 0:
                time.sleep(delay)

        return {
            "discovered_count": discovered_interactions,
            "deduplicated_count": deduped_total,
            "skipped_count": skipped_count,
            "newly_checked_count": newly_checked,
            "failed_count": failed_count,
            "category_counts": category_counts,
            "results": results,
        }

    def check_batch(
        self,
        usernames: List[str],
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        cancel_event: Optional[Any] = None,
        resume: bool = True,
        recheck: bool = False,
        limit: int = 0,
        delay: float = 2.0,
    ) -> List[Dict[str, Any]]:
        """Convenience method accepting a list of usernames, delegating to check_commenter_batch."""
        commenters_dict: Dict[str, Dict[str, Any]] = {}
        for u in usernames:
            raw = str(u).strip().lstrip("@")
            norm = raw.lower()
            if not norm:
                continue
            if norm not in commenters_dict:
                commenters_dict[norm] = {
                    "checked_username": raw,
                    "normalized_username": norm,
                    "source_accounts": [],
                    "comment_count": 1,
                    "reply_count": 0,
                    "total_interactions": 1,
                }
        effective_recheck = recheck or not resume
        res = self.check_commenter_batch(
            commenters_dict,
            recheck=effective_recheck,
            limit=limit,
            delay=delay,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
        )
        return res.get("results", [])


# ==============================================================================
# SPREADSHEET & EXCEL OUTPUT SYNCHRONIZATION
# ==============================================================================

class MumbaiCheckSync:
    """Synchronizes Mumbai Check results to Google Sheets and local Excel."""

    SHEET_TITLE = "Mumbai Check"

    # Color definitions for Excel & Google Sheets
    COLOR_MAP = {
        "Likely Mumbai": {"bg_hex": "D1E7DD", "font_hex": "0F5132", "rgb": (0.82, 0.91, 0.87)},
        "Serves Mumbai": {"bg_hex": "CFF4FC", "font_hex": "055160", "rgb": (0.81, 0.96, 0.99)},
        "Possibly Mumbai": {"bg_hex": "FFF3CD", "font_hex": "664D03", "rgb": (1.00, 0.95, 0.80)},
        "Likely outside Mumbai": {"bg_hex": "E2E3E5", "font_hex": "383D41", "rgb": (0.89, 0.89, 0.90)},
        "Not enough evidence": {"bg_hex": "F8D7DA", "font_hex": "842029", "rgb": (0.97, 0.84, 0.85)},
    }

    @classmethod
    def sync_to_excel(
        cls,
        checks: Optional[List[Dict[str, Any]]] = None,
        excel_path: str = DEFAULT_EXCEL_PATH,
    ) -> int:
        """
        Creates or updates the 'Mumbai Check' worksheet in exports/instagram_research.xlsx,
        strictly preserving all existing username tabs.
        """
        target = Path(excel_path)
        target.parent.mkdir(parents=True, exist_ok=True)

        if target.exists():
            wb = openpyxl.load_workbook(str(target))
        else:
            wb = openpyxl.Workbook()
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

        # Deduplicate checks, merging with DB history
        all_checks_dict: Dict[str, Dict[str, Any]] = {}
        try:
            for c in MumbaiCheckDB().get_all_checks():
                norm = (c.get("normalized_username") or c.get("checked_username") or c.get("username", "")).lower()
                all_checks_dict[norm] = c
        except Exception:
            pass
        if checks:
            for c in checks:
                norm = (c.get("normalized_username") or c.get("checked_username") or c.get("username", "")).lower()
                all_checks_dict[norm] = c

        final_checks = sorted(all_checks_dict.values(), key=lambda x: x.get("checked_at", ""), reverse=True)

        if cls.SHEET_TITLE in wb.sheetnames:
            ws = wb[cls.SHEET_TITLE]
            ws.delete_rows(1, ws.max_row + 1)
        else:
            ws = wb.create_sheet(title=cls.SHEET_TITLE)

        now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

        # Styles
        font_title = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_title = PatternFill("solid", fgColor="1F4E78")  # Dark Blue banner
        font_hdr = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        fill_hdr = PatternFill("solid", fgColor="2F343D")   # Graphite header
        font_body = Font(name="Calibri", size=9.5, color="1B2631")
        font_link = Font(name="Calibri", size=9.5, bold=True, underline="single", color="0563C1")
        thin_border = Border(
            left=Side(style="thin", color="D5D8DC"),
            right=Side(style="thin", color="D5D8DC"),
            top=Side(style="thin", color="D5D8DC"),
            bottom=Side(style="thin", color="D5D8DC"),
        )

        # Banner Row
        ws.row_dimensions[1].height = 26
        banner_vals = [
            (1, "MUMBAI LOCATION AUDIT (COMMENTERS & REPLIERS)"),
            (2, f"Last Updated: {now_str}"),
            (3, f"Total Accounts Checked: {len(final_checks)}"),
        ]
        for col, txt in banner_vals:
            c = ws.cell(row=1, column=col, value=txt)
            c.font = font_title
            c.fill = fill_title
            c.alignment = Alignment(vertical="center", horizontal="left" if col <= 2 else "center")

        # Spacer
        ws.row_dimensions[2].height = 10

        # Table Header
        ws.row_dimensions[3].height = 22
        headers = [
            "Checked Username", "Profile URL", "Display Name", "Source Account(s)",
            "Publicly Stated Location", "Evidence Summary", "Classification",
            "Confidence", "Checked Timestamp"
        ]
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=c_idx, value=h)
            cell.font = font_hdr
            cell.fill = fill_hdr
            cell.alignment = Alignment(vertical="center", horizontal="center" if c_idx in (7, 8, 9) else "left")
            cell.border = thin_border

        # Rows
        for r_idx, item in enumerate(final_checks, start=4):
            ws.row_dimensions[r_idx].height = 20
            u = item.get("checked_username") or item.get("username", "")
            p_url = item.get("profile_url", f"https://www.instagram.com/{u}/")
            disp = item.get("display_name", "—")
            srcs = item.get("source_accounts", "—")
            loc = item.get("stated_location", "—")
            evid = item.get("evidence_summary", "")
            cls_name = item.get("classification", "Not enough evidence")
            conf = item.get("confidence", "Low")
            ts = item.get("checked_at", now_str)

            # Cells
            c1 = ws.cell(row=r_idx, column=1, value=f"@{u}")
            c2 = ws.cell(row=r_idx, column=2, value=p_url)
            c2.hyperlink = p_url
            c2.font = font_link
            c3 = ws.cell(row=r_idx, column=3, value=disp)
            c4 = ws.cell(row=r_idx, column=4, value=srcs)
            c5 = ws.cell(row=r_idx, column=5, value=loc)
            c6 = ws.cell(row=r_idx, column=6, value=evid)
            c7 = ws.cell(row=r_idx, column=7, value=cls_name)
            c8 = ws.cell(row=r_idx, column=8, value=conf)
            c9 = ws.cell(row=r_idx, column=9, value=ts)

            # Apply base styling
            for col_i in range(1, 10):
                cell = ws.cell(row=r_idx, column=col_i)
                if col_i != 2:
                    cell.font = font_body
                cell.border = thin_border
                cell.alignment = Alignment(
                    vertical="center",
                    horizontal="center" if col_i in (1, 7, 8, 9) else "left"
                )

            # Classification color badge
            color_cfg = cls.COLOR_MAP.get(cls_name, cls.COLOR_MAP["Not enough evidence"])
            c7.fill = PatternFill("solid", fgColor=color_cfg["bg_hex"])
            c7.font = Font(name="Calibri", size=9.5, bold=True, color=color_cfg["font_hex"])

        # Auto-fit column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if cell.row == 1:
                    continue
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(14, min(max_len + 3, 50))

        wb.save(str(target))
        logger.info("Successfully updated '%s' worksheet in %s (%d rows)", cls.SHEET_TITLE, target, len(final_checks))
        return len(final_checks)

    @classmethod
    def sync_to_google_sheets(
        cls,
        checks: Optional[List[Dict[str, Any]]] = None,
        spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
        credentials_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Creates or updates the 'Mumbai Check' worksheet in the master Google Spreadsheet,
        preserving all existing worksheets.
        """
        mgr = GoogleSheetsManager(spreadsheet_id=spreadsheet_id)
        service = mgr.get_service(credentials_path=credentials_path)

        # Deduplicate checks, merging with DB history
        all_checks_dict: Dict[str, Dict[str, Any]] = {}
        try:
            for c in MumbaiCheckDB().get_all_checks():
                norm = (c.get("normalized_username") or c.get("checked_username") or c.get("username", "")).lower()
                all_checks_dict[norm] = c
        except Exception:
            pass
        if checks:
            for c in checks:
                norm = (c.get("normalized_username") or c.get("checked_username") or c.get("username", "")).lower()
                all_checks_dict[norm] = c

        final_checks = sorted(all_checks_dict.values(), key=lambda x: x.get("checked_at", ""), reverse=True)

        # Get existing sheet metadata
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets", [])}

        sheet_title = cls.SHEET_TITLE
        if sheet_title not in sheet_map:
            add_req = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": sheet_title,
                                "gridProperties": {"rowCount": max(100, len(final_checks) + 15), "columnCount": 11},
                            }
                        }
                    }
                ]
            }
            res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=add_req).execute()
            sheet_id = res["replies"][0]["addSheet"]["properties"]["sheetId"]
            logger.info("Created new sheet '%s' (id: %s) in Google Spreadsheet", sheet_title, sheet_id)
        else:
            sheet_id = sheet_map[sheet_title]
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id, range=f"'{sheet_title}'!A:Z"
            ).execute()
            logger.info("Cleared existing sheet '%s' for clean update", sheet_title)

        now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

        # Build values grid
        rows: List[List[Any]] = [
            # Row 1: Banner
            ["MUMBAI LOCATION AUDIT (COMMENTERS & REPLIERS)", f"Last Updated: {now_str}", f"Total Accounts: {len(final_checks)}"],
            # Row 2: Spacer
            [],
            # Row 3: Header
            [
                "Checked Username", "Profile URL", "Display Name", "Source Account(s)",
                "Publicly Stated Location", "Evidence Summary", "Classification",
                "Confidence", "Checked Timestamp"
            ],
        ]

        for item in final_checks:
            u = item.get("checked_username") or item.get("username", "")
            p_url = item.get("profile_url", f"https://www.instagram.com/{u}/")
            disp = item.get("display_name", "—")
            srcs = item.get("source_accounts", "—")
            loc = item.get("stated_location", "—")
            evid = item.get("evidence_summary", "")
            cls_name = item.get("classification", "Not enough evidence")
            conf = item.get("confidence", "Low")
            ts = item.get("checked_at", now_str)

            hyperlink_formula = f'=HYPERLINK("{p_url}", "{p_url}")'
            rows.append([
                f"@{u}",
                hyperlink_formula,
                disp,
                srcs,
                loc,
                evid,
                cls_name,
                conf,
                ts,
            ])

        # Write data
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'!A1",
            valueInputOption="USER_ENTERED",
            body={"values": rows},
        ).execute()

        # Format header banner and column widths
        format_requests = [
            # Banner format
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 9},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
                            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True, "fontSize": 11},
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat)",
                }
            },
            # Header row format
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 9},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.18, "green": 0.20, "blue": 0.24},
                            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True, "fontSize": 10},
                            "horizontalAlignment": "LEFT",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            },
            # Auto-resize columns
            {
                "autoResizeDimensions": {
                    "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 9}
                }
            },
        ]

        try:
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": format_requests}).execute()
        except Exception as e:
            logger.warning("Formatting batch update warning: %s", e)

        return {
            "status": "success",
            "spreadsheet_id": spreadsheet_id,
            "sheet_title": sheet_title,
            "rows_synced": len(final_checks),
            "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={sheet_id}",
        }


# ==============================================================================
# CLI INTERFACE
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Audit Instagram Commenters and Repliers for Mumbai location."
    )
    parser.add_argument(
        "--source-account", "-s", action="append",
        help="Source account(s) to load commenters from (e.g. licious_foods, or 'all')."
    )
    parser.add_argument(
        "--usernames", "-u",
        help="Explicit comma-separated list of Instagram usernames to check."
    )
    parser.add_argument(
        "--recheck", action="store_true",
        help="Force rechecking of previously completed usernames."
    )
    parser.add_argument(
        "--limit", type=int, default=0,
        help="Maximum number of accounts to check in this run (0 = unlimited)."
    )
    parser.add_argument(
        "--delay", type=float, default=2.0,
        help="Delay in seconds between requests (default: 2.0)."
    )
    parser.add_argument(
        "--no-sheets", action="store_true",
        help="Skip syncing to Google Spreadsheet."
    )
    parser.add_argument(
        "--no-excel", action="store_true",
        help="Skip syncing to local Excel mirror."
    )
    parser.add_argument(
        "--list-sources", action="store_true",
        help="List available source accounts with extracted comments."
    )

    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    if args.list_sources:
        sources = get_available_checkpoint_accounts()
        print(f"\nDiscovered {len(sources)} source accounts with extracted comments:")
        for s in sources:
            print(f" - @{s['source_username']}: {s['total_records']} records ({s['comments_count']} comments, {s['replies_count']} replies) | {s['unique_commenters']} unique commenters")
        return

    checker = MumbaiLocationChecker()
    commenters_dict: Dict[str, Dict[str, Any]] = {}

    if args.usernames:
        for u in args.usernames.split(","):
            norm = u.strip().lstrip("@").lower()
            if norm:
                commenters_dict[norm] = {
                    "checked_username": u.strip().lstrip("@"),
                    "normalized_username": norm,
                    "source_accounts": ["manual_input"],
                    "comment_count": 1,
                    "reply_count": 0,
                    "total_interactions": 1,
                }
    else:
        sources = args.source_account or ["all"]
        print(f"[+] Loading commenters from source accounts: {', '.join(sources)}...")
        commenters_dict = load_commenters_from_checkpoints(sources)

    if not commenters_dict:
        print("[!] No commenters found to process.")
        return

    print(f"[+] Discovered {sum(c['total_interactions'] for c in commenters_dict.values())} comment/reply records.")
    print(f"[+] Deduplicated to {len(commenters_dict)} unique commenter/replier accounts.")

    telemetry = checker.check_commenter_batch(
        commenters_dict,
        recheck=args.recheck,
        limit=args.limit,
        delay=args.delay,
    )

    print("\n" + "=" * 60)
    print("MUMBAI LOCATION CHECK RESULTS SUMMARY")
    print("=" * 60)
    print(f"Total Discovered:        {telemetry['discovered_count']}")
    print(f"Unique Deduplicated:     {telemetry['deduplicated_count']}")
    print(f"Skipped (Already Done):  {telemetry['skipped_count']}")
    print(f"Newly Checked:           {telemetry['newly_checked_count']}")
    print(f"Failed / Inaccessible:   {telemetry['failed_count']}")
    print("-" * 60)
    for cat, count in telemetry.get("category_counts", {}).items():
        print(f"  {cat:22}: {count}")
    print("=" * 60)

    # Output sync
    if not args.no_excel:
        ex_cnt = MumbaiCheckSync.sync_to_excel()
        print(f"[+] Local Excel mirror updated ({ex_cnt} accounts) in {DEFAULT_EXCEL_PATH}")

    if not args.no_sheets:
        try:
            sh_res = MumbaiCheckSync.sync_to_google_sheets()
            print(f"[+] Google Spreadsheet updated: {sh_res['spreadsheet_url']}")
        except Exception as e:
            print(f"[!] Google Sheets sync warning: {e}")


if __name__ == "__main__":
    main()
