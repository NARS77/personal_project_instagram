#!/usr/bin/env python3
"""
Excel and Google Sheets Synchronization Module for Instagram Extractor.

Features:
- Dual-destination synchronization:
  1. Online Master: Google Spreadsheet (ID: 1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y)
  2. Local Mirror: Single multi-tab Excel workbook (exports/instagram_research.xlsx)
- One username = One worksheet tab across both destinations.
- Strict hierarchical layout:
  - Header block at the top with username and update timestamp.
  - Per post: Post shortcode, clickable URL link (shown ONCE per post block), likes, total comments reported.
  - NO post captions or hashtags.
  - Top-level comments followed directly by indented child replies.
  - Visual separation between posts.
- Graceful Google Sheets authentication handling:
  - Supports Service Account (service_account.json) or OAuth 2.0 (credentials.json / token.json).
  - If auth is missing, saves locally to Excel and SQLite checkpoint without failure,
    and provides explicit setup instructions.
"""

import json
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger("instagram_comments.sync")

DEFAULT_SPREADSHEET_ID = "1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y"
DEFAULT_EXCEL_PATH = "exports/instagram_research.xlsx"

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]


class GoogleAuthMissingError(Exception):
    """Raised when Google Sheets API credentials are not found or unconfigured."""
    pass


def format_timestamp(ts: Any) -> str:
    """Formats epoch integer or string into UTC YYYY-MM-DD HH:MM:SS string."""
    if not ts:
        return "—"
    try:
        val = int(ts)
        # Check if milliseconds
        if val > 10_000_000_000:
            val = val // 1000
        return datetime.fromtimestamp(val, timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, OSError):
        return str(ts)


def sanitize_sheet_title(name: str) -> str:
    """Sanitizes username for Excel and Google Sheets tab title (max 31 chars, no forbidden chars)."""
    # Remove chars: [ ] : * ? / \
    clean = re.sub(r"[\[\]:*?/\\]", "_", name.strip())
    clean = clean[:31] if len(clean) > 31 else clean
    return clean or "instagram_data"


class ExcelWorkbookManager:
    """Manages the local multi-tab Excel workbook (exports/instagram_research.xlsx)."""

    def __init__(self, excel_path: str = DEFAULT_EXCEL_PATH):
        self.excel_path = Path(excel_path)
        self.excel_path.parent.mkdir(parents=True, exist_ok=True)

    def export_from_db(self, db_path: Path, username: str) -> int:
        """
        Reads posts and comments from SQLite checkpoint DB and creates/updates
        the worksheet tab for username in the workbook.
        Returns the total number of records written (posts + comments + replies).
        """
        if not db_path.exists():
            raise FileNotFoundError(f"Checkpoint database not found at {db_path}")

        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        try:
            posts = conn.execute("SELECT * FROM posts ORDER BY post_timestamp DESC").fetchall()
            if not posts:
                logger.warning("No posts found in checkpoint DB %s for username '%s'", db_path, username)
                return 0

            # Load or create workbook
            if self.excel_path.exists():
                try:
                    wb = openpyxl.load_workbook(str(self.excel_path))
                except Exception as e:
                    logger.warning("Could not read existing workbook (%s), initializing new one", e)
                    wb = openpyxl.Workbook()
                    if "Sheet" in wb.sheetnames:
                        wb.remove(wb["Sheet"])
            else:
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])

            sheet_title = sanitize_sheet_title(username)
            if sheet_title in wb.sheetnames:
                ws = wb[sheet_title]
                ws.delete_rows(1, ws.max_row + 1)
            else:
                ws = wb.create_sheet(title=sheet_title)

            # Define Styles
            font_title = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            fill_title = PatternFill("solid", fgColor="1F4E78")  # Dark Blue banner

            font_post_hdr = Font(name="Calibri", size=10, bold=True, color="1F4E78")
            fill_post_hdr = PatternFill("solid", fgColor="D9E1F2")  # Soft Ice Blue

            font_link = Font(name="Calibri", size=10, bold=True, underline="single", color="0563C1")

            font_tbl_hdr = Font(name="Calibri", size=9, bold=True, color="2C3E50")
            fill_tbl_hdr = PatternFill("solid", fgColor="EAECEE")  # Soft Gray

            font_comment = Font(name="Calibri", size=9.5, color="1B2631")
            font_comment_bold = Font(name="Calibri", size=9.5, bold=True, color="1B2631")

            font_reply = Font(name="Calibri", size=9, italic=True, color="34495E")
            fill_reply = PatternFill("solid", fgColor="F7F9FB")  # Subtle light slate

            thin_border_side = Side(style="thin", color="D5D8DC")
            thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
            bottom_divider = Border(bottom=Side(style="medium", color="BDC3C7"))

            align_left = Alignment(horizontal="left", vertical="center")
            align_center = Alignment(horizontal="center", vertical="center")
            align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

            # Query overall counts
            tot_comments = conn.execute(
                "SELECT count(*) FROM comments WHERE record_type = 'comment'"
            ).fetchone()[0]
            tot_replies = conn.execute(
                "SELECT count(*) FROM comments WHERE record_type = 'reply'"
            ).fetchone()[0]

            now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

            # Row 1: Header Banner
            ws.row_dimensions[1].height = 26
            banner_cells = [
                (1, f"INSTAGRAM RESEARCH: @{username}"),
                (2, f"Last Updated: {now_str}"),
                (3, f"Posts: {len(posts)}"),
                (4, f"Comments: {tot_comments}"),
                (5, f"Replies: {tot_replies}"),
                (6, f"Total Records: {tot_comments + tot_replies}"),
            ]
            for col_idx, text in banner_cells:
                cell = ws.cell(row=1, column=col_idx, value=text)
                cell.font = font_title
                cell.fill = fill_title
                cell.alignment = align_left if col_idx <= 2 else align_center

            # Row 2: Blank spacer
            ws.row_dimensions[2].height = 10
            current_row = 3

            # Loop through each post
            for p_idx, p in enumerate(posts, start=1):
                media_id = p["media_id"]
                shortcode = p["shortcode"]
                post_likes = p["likes"] or 0
                rep_comments = p["comments_count"] or 0
                post_url = f"https://www.instagram.com/p/{shortcode}/"

                # Get comments and replies for this post
                top_comments = conn.execute(
                    "SELECT * FROM comments WHERE media_id = ? AND record_type = 'comment' ORDER BY created_at ASC",
                    (media_id,),
                ).fetchall()

                child_replies = conn.execute(
                    "SELECT * FROM comments WHERE media_id = ? AND record_type = 'reply' ORDER BY created_at ASC",
                    (media_id,),
                ).fetchall()

                # Group replies by parent_comment_id
                replies_by_parent: Dict[str, List[sqlite3.Row]] = {}
                for r in child_replies:
                    pid = r["parent_comment_id"]
                    replies_by_parent.setdefault(pid, []).append(r)

                # Post Header Row
                ws.row_dimensions[current_row].height = 24
                c1 = ws.cell(row=current_row, column=1, value=f"POST #{p_idx}: {shortcode}")
                c2 = ws.cell(row=current_row, column=2, value=post_url)
                c2.hyperlink = post_url
                c2.font = font_link
                c3 = ws.cell(row=current_row, column=3, value=f"Likes: {post_likes:,}")
                c4 = ws.cell(row=current_row, column=4, value=f"Reported Comments: {rep_comments:,}")
                c5 = ws.cell(row=current_row, column=5, value=f"Extracted Comments: {len(top_comments):,}")
                c6 = ws.cell(row=current_row, column=6, value=f"Extracted Replies: {len(child_replies):,}")

                for col in range(1, 7):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = fill_post_hdr
                    if col != 2:
                        cell.font = font_post_hdr
                    cell.alignment = align_left if col <= 2 else align_center

                current_row += 1

                # Table Header Row
                ws.row_dimensions[current_row].height = 20
                headers = [
                    "Type",
                    "Author",
                    "Comment / Reply Text",
                    "Likes",
                    "Date / Time (UTC)",
                    "Replies / Parent ID",
                ]
                for col_idx, h_text in enumerate(headers, start=1):
                    cell = ws.cell(row=current_row, column=col_idx, value=h_text)
                    cell.font = font_tbl_hdr
                    cell.fill = fill_tbl_hdr
                    cell.border = thin_border
                    cell.alignment = align_left if col_idx in (1, 2, 3) else align_center

                current_row += 1

                if not top_comments:
                    ws.row_dimensions[current_row].height = 18
                    empty_vals = ["(No comments)", "—", "No comments extracted for this post.", 0, "—", "—"]
                    for col_idx, val in enumerate(empty_vals, start=1):
                        cell = ws.cell(row=current_row, column=col_idx, value=val)
                        cell.font = font_reply
                        cell.border = thin_border
                        cell.alignment = align_left if col_idx <= 3 else align_center
                    current_row += 1
                else:
                    for c in top_comments:
                        c_id = c["comment_id"]
                        c_user = c["username"] or "unknown"
                        c_text = c["text"] or ""
                        c_likes = c["likes"] or 0
                        c_date = format_timestamp(c["created_at"])
                        c_rep_count = f"{c['replies_count']} replies"

                        ws.row_dimensions[current_row].height = 22
                        r_vals = ["Comment", f"@{c_user}", c_text, c_likes, c_date, c_rep_count]
                        for col_idx, val in enumerate(r_vals, start=1):
                            cell = ws.cell(row=current_row, column=col_idx, value=val)
                            cell.font = font_comment_bold if col_idx == 2 else font_comment
                            cell.border = thin_border
                            if col_idx == 3:
                                cell.alignment = align_wrap
                            elif col_idx in (1, 2):
                                cell.alignment = align_left
                            else:
                                cell.alignment = align_center

                        current_row += 1

                        # Child replies indented directly beneath this comment
                        parent_reps = replies_by_parent.get(c_id, [])
                        for rep in parent_reps:
                            r_user = rep["username"] or "unknown"
                            r_text = rep["text"] or ""
                            r_likes = rep["likes"] or 0
                            r_date = format_timestamp(rep["created_at"])
                            r_parent_info = f"Parent: {c_id}"

                            ws.row_dimensions[current_row].height = 20
                            rep_vals = [
                                "  ↳ Reply",
                                f"  ↳ @{r_user}",
                                f"  {r_text}",
                                r_likes,
                                r_date,
                                r_parent_info,
                            ]
                            for col_idx, val in enumerate(rep_vals, start=1):
                                cell = ws.cell(row=current_row, column=col_idx, value=val)
                                cell.font = font_reply
                                cell.fill = fill_reply
                                cell.border = thin_border
                                if col_idx == 3:
                                    cell.alignment = align_wrap
                                elif col_idx in (1, 2):
                                    cell.alignment = align_left
                                else:
                                    cell.alignment = align_center

                            current_row += 1

                # Visual Separator between posts
                ws.row_dimensions[current_row].height = 12
                for col in range(1, 7):
                    cell = ws.cell(row=current_row, column=col, value="")
                    cell.border = bottom_divider
                current_row += 1

            # Adjust column widths
            col_widths = {1: 16, 2: 24, 3: 65, 4: 12, 5: 22, 6: 24}
            for col_idx, width in col_widths.items():
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = width

            # Freeze pane at row 3 so title banner is always visible
            ws.freeze_panes = "A3"

            # Save workbook safely
            wb.save(str(self.excel_path))
            total_records = len(posts) + tot_comments + tot_replies
            logger.info(
                "Updated Excel tab '%s' in %s (%d posts, %d comments, %d replies)",
                sheet_title,
                self.excel_path,
                len(posts),
                tot_comments,
                tot_replies,
            )
            return total_records
        finally:
            conn.close()


class GoogleSheetsManager:
    """Manages synchronization with the Google Spreadsheet master."""

    def __init__(self, spreadsheet_id: str = DEFAULT_SPREADSHEET_ID):
        self.spreadsheet_id = spreadsheet_id
        self._service = None

    def get_service(self, credentials_path: Optional[str] = None):
        """
        Attempts to authenticate with Google Sheets API.
        Checks:
        1. Explicit credentials_path argument
        2. Environment variables (GOOGLE_SHEETS_CREDENTIALS, GOOGLE_APPLICATION_CREDENTIALS)
        3. Local service_account.json
        4. Local token.json (from prior OAuth login)
        5. Local credentials.json (OAuth desktop client)
        """
        if self._service:
            return self._service

        try:
            from google.oauth2 import service_account
            from google.oauth2.credentials import Credentials
            from google.auth.transport.requests import Request
            from googleapiclient.discovery import build
        except ImportError:
            raise RuntimeError(
                "Google API client libraries are not installed. "
                "Run: pip install --user --break-system-packages google-api-python-client google-auth google-auth-oauthlib"
            )

        creds = None
        candidate_paths = []
        if credentials_path:
            candidate_paths.append(credentials_path)
        for env_var in ("GOOGLE_SHEETS_CREDENTIALS", "GOOGLE_APPLICATION_CREDENTIALS"):
            val = os.getenv(env_var)
            if val:
                candidate_paths.append(val)

        script_dir = Path(__file__).resolve().parent
        parent_dir = script_dir.parent

        candidate_paths.extend([
            str(script_dir / "service_account.json"),
            str(script_dir / "service_account.json.json"),
            str(parent_dir / "service_account.json"),
            str(parent_dir / "service_account.json.json"),
            "service_account.json",
            "service_account.json.json",
            str(script_dir / "token.json"),
            str(parent_dir / "token.json"),
            "token.json",
            str(script_dir / "credentials.json"),
            str(parent_dir / "credentials.json"),
            "credentials.json",
            "client_secret.json",
            os.path.expanduser("~/.config/google/service_account.json"),
            os.path.expanduser("~/.config/google/token.json"),
            os.path.expanduser("~/.config/google/credentials.json"),
        ])
        # Also auto-detect any service account key file in project directories or Downloads
        search_dirs = [Path("."), script_dir, parent_dir]
        for d in search_dirs:
            if d.exists():
                for f in d.glob("*service_account*.json*"):
                    if f.is_file() and str(f.resolve()) not in candidate_paths:
                        candidate_paths.append(str(f.resolve()))
        for f in list(Path("~/Downloads").expanduser().glob("instagram-sheet*.json")):
            if f.is_file() and str(f.resolve()) not in candidate_paths:
                candidate_paths.append(str(f.resolve()))


        # 1. Try token.json (OAuth user credentials)
        token_file = next((p for p in candidate_paths if p.endswith("token.json") and Path(p).exists()), None)
        if token_file:
            try:
                creds = Credentials.from_authorized_user_file(token_file, GOOGLE_SCOPES)
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
            except Exception as e:
                logger.warning("Failed loading token from %s: %s", token_file, e)
                creds = None

        # 2. Try Service Account
        if not creds:
            sa_file = next((p for p in candidate_paths if "service_account" in p and Path(p).exists()), None)
            if sa_file:
                try:
                    creds = service_account.Credentials.from_service_account_file(
                        sa_file, scopes=GOOGLE_SCOPES
                    )
                    creds.refresh(Request())
                    logger.info("Loaded and verified Service Account from %s", sa_file)
                except Exception as e:
                    logger.warning("Failed loading service account from %s: %s", sa_file, e)
                    creds = None

        # 3. Try credentials.json / client_secret.json (OAuth Flow)
        if not creds:
            client_file = next((p for p in candidate_paths if ("credentials.json" in p or "client_secret" in p) and Path(p).exists()), None)
            if client_file:
                try:
                    from google_auth_oauthlib.flow import InstalledAppFlow
                    flow = InstalledAppFlow.from_client_secrets_file(client_file, GOOGLE_SCOPES)
                    creds = flow.run_local_server(port=0)
                    with open("token.json", "w", encoding="utf-8") as token_out:
                        token_out.write(creds.to_json())
                    logger.info("Saved new OAuth token to token.json")
                except Exception as e:
                    logger.warning("Failed running OAuth flow with %s: %s", client_file, e)
                    creds = None

        if not creds:
            raise GoogleAuthMissingError(
                "Google Sheets credentials were not found or not yet authorized.\n"
                "Please configure one of the following methods:\n\n"
                "METHOD 1 (Recommended for Automation - Service Account):\n"
                "  1. In Google Cloud Console, enable the Google Sheets API.\n"
                "  2. Create a Service Account and download its JSON key as 'service_account.json'.\n"
                "  3. Open your Google Sheet:\n"
                f"     https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit\n"
                "  4. Click 'Share' and add the service account's client email as an Editor.\n\n"
                "METHOD 2 (Interactive OAuth 2.0):\n"
                "  1. In Google Cloud Console, create an OAuth 2.0 Client ID (type: Desktop App).\n"
                "  2. Download the client secret JSON as 'credentials.json'.\n"
                "  3. Run: python3 instagram_comments.py --auth-google\n"
            )

        self._service = build("sheets", "v4", credentials=creds, cache_discovery=False)
        return self._service

    def export_from_db(self, db_path: Path, username: str, credentials_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Connects to Google Spreadsheet ID, verifies or creates the username worksheet tab,
        clears existing content, writes formatted post and comment rows, and formats the sheet.
        """
        service = self.get_service(credentials_path=credentials_path)

        if not db_path.exists():
            return False, f"Checkpoint database {db_path} does not exist"

        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        try:
            posts = conn.execute("SELECT * FROM posts ORDER BY post_timestamp DESC").fetchall()
            if not posts:
                return False, f"No posts found in checkpoint DB {db_path}"

            sheet_title = sanitize_sheet_title(username)

            # 1. Fetch current spreadsheet metadata
            ss_meta = service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
            sheets = ss_meta.get("sheets", [])
            sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in sheets}

            # 2. Add or Clear Sheet
            if sheet_title not in sheet_map:
                add_req = {
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": sheet_title,
                                }
                            }
                        }
                    ]
                }
                res = service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id, body=add_req
                ).execute()
                sheet_id = res["replies"][0]["addSheet"]["properties"]["sheetId"]
                logger.info("Created new sheet tab '%s' (id: %s) in Google Spreadsheet", sheet_title, sheet_id)
            else:
                sheet_id = sheet_map[sheet_title]
                service.spreadsheets().values().clear(
                    spreadsheetId=self.spreadsheet_id, range=f"'{sheet_title}'!A:Z"
                ).execute()
                logger.info("Cleared existing sheet tab '%s' for clean update", sheet_title)

            # 3. Build data grid
            tot_comments = conn.execute(
                "SELECT count(*) FROM comments WHERE record_type = 'comment'"
            ).fetchone()[0]
            tot_replies = conn.execute(
                "SELECT count(*) FROM comments WHERE record_type = 'reply'"
            ).fetchone()[0]
            now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

            rows: List[List[Any]] = []

            # Banner row
            rows.append([
                f"INSTAGRAM RESEARCH: @{username}",
                f"Last Updated: {now_str}",
                f"Posts: {len(posts)}",
                f"Comments: {tot_comments}",
                f"Replies: {tot_replies}",
                f"Total Records: {tot_comments + tot_replies}",
            ])
            # Spacer row
            rows.append([])

            for p_idx, p in enumerate(posts, start=1):
                media_id = p["media_id"]
                shortcode = p["shortcode"]
                post_likes = p["likes"] or 0
                rep_comments = p["comments_count"] or 0
                post_url = f"https://www.instagram.com/p/{shortcode}/"

                top_comments = conn.execute(
                    "SELECT * FROM comments WHERE media_id = ? AND record_type = 'comment' ORDER BY created_at ASC",
                    (media_id,),
                ).fetchall()

                child_replies = conn.execute(
                    "SELECT * FROM comments WHERE media_id = ? AND record_type = 'reply' ORDER BY created_at ASC",
                    (media_id,),
                ).fetchall()

                replies_by_parent: Dict[str, List[sqlite3.Row]] = {}
                for r in child_replies:
                    pid = r["parent_comment_id"]
                    replies_by_parent.setdefault(pid, []).append(r)

                # Post Header Row
                # In Google Sheets USER_ENTERED mode, =HYPERLINK formula renders as clickable link
                hyperlink_formula = f'=HYPERLINK("{post_url}", "{post_url}")'
                rows.append([
                    f"POST #{p_idx}: {shortcode}",
                    hyperlink_formula,
                    f"Likes: {post_likes:,}",
                    f"Reported Comments: {rep_comments:,}",
                    f"Extracted Comments: {len(top_comments):,}",
                    f"Extracted Replies: {len(child_replies):,}",
                ])

                # Table Header Row
                rows.append([
                    "Type",
                    "Author",
                    "Comment / Reply Text",
                    "Likes",
                    "Date / Time (UTC)",
                    "Replies / Parent ID",
                ])

                if not top_comments:
                    rows.append([
                        "(No comments)",
                        "—",
                        "No comments extracted for this post.",
                        0,
                        "—",
                        "—",
                    ])
                else:
                    for c in top_comments:
                        c_id = c["comment_id"]
                        c_user = c["username"] or "unknown"
                        c_text = c["text"] or ""
                        c_likes = c["likes"] or 0
                        c_date = format_timestamp(c["created_at"])
                        c_rep_count = f"{c['replies_count']} replies"

                        rows.append([
                            "Comment",
                            f"@{c_user}",
                            c_text,
                            c_likes,
                            c_date,
                            c_rep_count,
                        ])

                        parent_reps = replies_by_parent.get(c_id, [])
                        for rep in parent_reps:
                            r_user = rep["username"] or "unknown"
                            r_text = rep["text"] or ""
                            r_likes = rep["likes"] or 0
                            r_date = format_timestamp(rep["created_at"])
                            r_parent_info = f"Parent: {c_id}"

                            rows.append([
                                "  ↳ Reply",
                                f"  ↳ @{r_user}",
                                f"  {r_text}",
                                r_likes,
                                r_date,
                                r_parent_info,
                            ])

                # Separator empty row
                rows.append([])

            # 4. Write values using USER_ENTERED
            body = {"values": rows}
            write_res = service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_title}'!A1",
                valueInputOption="USER_ENTERED",
                body=body,
            ).execute()

            updated_cells = write_res.get("updatedCells", 0)

            # 5. Format Google Sheet via batchUpdate (Column widths, freeze pane)
            format_requests = [
                # Freeze top 2 rows
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "frozenRowCount": 2,
                            },
                        },
                        "fields": "gridProperties.frozenRowCount",
                    }
                },
                # Set Column Widths (A: 130, B: 180, C: 450, D: 80, E: 160, F: 180 pixels)
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1,
                        },
                        "properties": {"pixelSize": 130},
                        "fields": "pixelSize",
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 1,
                            "endIndex": 2,
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize",
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,
                            "endIndex": 3,
                        },
                        "properties": {"pixelSize": 450},
                        "fields": "pixelSize",
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 3,
                            "endIndex": 4,
                        },
                        "properties": {"pixelSize": 80},
                        "fields": "pixelSize",
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 4,
                            "endIndex": 5,
                        },
                        "properties": {"pixelSize": 160},
                        "fields": "pixelSize",
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 5,
                            "endIndex": 6,
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize",
                    }
                },
            ]

            service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": format_requests},
            ).execute()

            msg = (
                f"Successfully synced tab '{sheet_title}' to Google Spreadsheet "
                f"(ID: {self.spreadsheet_id}). Updated {updated_cells} cells."
            )
            logger.info(msg)
            return True, msg
        finally:
            conn.close()


class SyncCoordinator:
    """Coordinates local Excel mirroring and Google Sheets master synchronization."""

    def __init__(
        self,
        spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
        excel_path: str = DEFAULT_EXCEL_PATH,
        credentials_path: Optional[str] = None,
    ):
        self.spreadsheet_id = spreadsheet_id
        self.excel_path = excel_path
        self.credentials_path = credentials_path
        self.excel_mgr = ExcelWorkbookManager(excel_path=self.excel_path)
        self.sheets_mgr = GoogleSheetsManager(spreadsheet_id=self.spreadsheet_id)

    def sync(
        self,
        db_path: Path,
        username: str,
        sync_excel: bool = True,
        sync_sheets: bool = True,
    ) -> Dict[str, Any]:
        """
        Executes dual sync:
        1. Local Excel workbook update (if sync_excel=True)
        2. Google Sheets master update (if sync_sheets=True)
        Handles auth failures gracefully without raising or losing data.
        """
        results: Dict[str, Any] = {
            "excel_success": False,
            "excel_records": 0,
            "excel_path": str(self.excel_path),
            "sheets_success": False,
            "sheets_message": "",
            "sheets_pending_auth": False,
        }

        # 1. Update local Excel mirror
        if sync_excel:
            try:
                records = self.excel_mgr.export_from_db(db_path, username)
                results["excel_success"] = True
                results["excel_records"] = records
                print(f"\n[+] Local Excel mirror updated: {self.excel_path} (tab '{sanitize_sheet_title(username)}')")
            except Exception as e:
                logger.error("Failed to update local Excel workbook: %s", e)
                print(f"\n[-] Error updating local Excel workbook: {e}")
        else:
            results["excel_message"] = "Excel sync disabled by user."

        # 2. Update Google Spreadsheet master
        if sync_sheets:
            try:
                ok, msg = self.sheets_mgr.export_from_db(
                    db_path, username, credentials_path=self.credentials_path
                )
                results["sheets_success"] = ok
                results["sheets_message"] = msg
                if ok:
                    print(f"[+] Google Spreadsheet master updated: https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit")
                else:
                    print(f"[-] Google Spreadsheet sync failed: {msg}")
            except GoogleAuthMissingError as e:
                results["sheets_pending_auth"] = True
                results["sheets_message"] = str(e)
                print("\n" + "=" * 70)
                print("GOOGLE SHEETS INTEGRATION STATUS: PENDING AUTHORIZATION")
                print("=" * 70)
                print(f"Local Excel workbook was saved successfully at: {self.excel_path}")
                print(f"SQLite checkpoint state is preserved at: {db_path}")
                print("\nTo connect Google Spreadsheet (ID: " + self.spreadsheet_id + "):")
                print(str(e))
                print("=" * 70)
            except Exception as e:
                results["sheets_message"] = str(e)
                logger.error("Unexpected error syncing to Google Sheets: %s", e)
                print(f"\n[-] Google Sheets sync encountered error: {e}")
        return results


def get_lead_outreach_sync():
    """Lazy importer for LeadOutreachSync to avoid circular imports."""
    from lead_outreach import LeadOutreachSync
    return LeadOutreachSync

