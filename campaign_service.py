#!/usr/bin/env python3
"""
Campaign Service & Execution Engine for MeatNDoor Campaign Automation.

Implements:
1. Campaign configuration, persistence, and execution tracking in SQLite.
2. The 10-stage end-to-end pipeline:
   - Config validation -> Source discovery -> Deduplication -> Location qualification
   - Lead scoring (Intent + Location) -> Message drafting -> DNC & eligibility checks
   - SQLite update -> Google Sheets / Excel sync -> Completion reporting.
3. Resumable checkpointing, pause, cancel, and duplicate-start prevention.
4. Lead scoring separating explicit buying intent signals from location confidence.
5. Actionable, searchable, SQL-paginated review queue with operator audit logs.
"""

import json
import logging
import os
import re
import sqlite3
import threading
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from lead_outreach import (
    DEFAULT_EXCEL_PATH,
    DEFAULT_SPREADSHEET_ID,
    ELIGIBILITY_BUSINESS_REVIEW,
    ELIGIBILITY_LOW_CONF_REVIEW,
    ELIGIBILITY_MUMBAI_TARGET,
    ELIGIBILITY_PENDING_LOCATION,
    ELIGIBILITY_SUPPRESSED_DNC,
    ELIGIBILITY_SUPPRESSED_OUTSIDE,
    GENERIC_OUTREACH_TEMPLATE,
    MUMBAI_OUTREACH_TEMPLATE,
    STATUS_APPROVED,
    STATUS_CONTACTED,
    STATUS_DNC,
    STATUS_INTERESTED,
    STATUS_NOT_INTERESTED,
    STATUS_PENDING,
    STATUS_REJECTED,
    STATUS_REPLIED,
    LeadOutreachDB,
    get_canonical_checkpoints_dir,
    get_canonical_mumbai_db_path,
    get_canonical_outreach_db_path,
    normalize_username,
)

logger = logging.getLogger("instagram_comments.campaign_service")

# Campaign Status Constants
CAMPAIGN_STATUS_DRAFT = "draft"
CAMPAIGN_STATUS_RUNNING = "running"
CAMPAIGN_STATUS_PAUSED = "paused"
CAMPAIGN_STATUS_COMPLETED = "completed"
CAMPAIGN_STATUS_CANCELLED = "cancelled"
CAMPAIGN_STATUS_FAILED = "failed"

ALL_CAMPAIGN_STATUSES = [
    CAMPAIGN_STATUS_DRAFT,
    CAMPAIGN_STATUS_RUNNING,
    CAMPAIGN_STATUS_PAUSED,
    CAMPAIGN_STATUS_COMPLETED,
    CAMPAIGN_STATUS_CANCELLED,
    CAMPAIGN_STATUS_FAILED,
]

# High-Intent Keyword Dictionary for Lead Scoring
BUYING_INTENT_PATTERNS = [
    # Explicit order/delivery inquiries (+15)
    (r"(?i)\b(order|deliver|delivery|home delivery|doorstep|cod|dispatch|shipping)\b", 15, "Delivery/Order Inquiry"),
    # Specific meat/cut product mentions (+15)
    (r"(?i)\b(mutton|chicken|prawns|fish|meat|bacon|sausage|keema|chops|steak|curry cut|boneless|fresh)\b", 15, "Specific Meat Product Interest"),
    # Price/availability inquiries (+10)
    (r"(?i)\b(price|cost|rate|how much|available|availability|in stock|menu)\b", 10, "Price/Availability Inquiry"),
    # Location/store inquiries (+10)
    (r"(?i)\b(store|shop|outlet|bkc|bandra|andheri|powai|juhu|colaba|thane|vashi)\b", 10, "Location Inquiry"),
]


class CampaignError(Exception):
    """Base exception for campaign operations."""
    pass


class CampaignValidationError(CampaignError):
    """Raised on invalid campaign inputs or configurations."""
    pass


class CampaignDuplicateExecutionError(CampaignError):
    """Raised when attempting to start an already running campaign."""
    pass


class CampaignService:
    """
    Manages campaign state, execution workflows, and lead scoring.
    """

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = Path(db_path) if db_path else get_canonical_outreach_db_path()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.lead_outreach_db = LeadOutreachDB(db_path=self.db_path)
        self._init_campaign_tables()
        self._active_events: Dict[str, Dict[str, threading.Event]] = {}
        self._lock = threading.Lock()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _init_campaign_tables(self) -> None:
        """Creates campaign-specific tables in the SQLite database."""
        with self._get_connection() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS campaigns (
                    campaign_id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    objective TEXT NOT NULL,
                    target_location TEXT NOT NULL,
                    source_accounts TEXT NOT NULL,
                    max_leads INTEGER DEFAULT 500,
                    min_lead_score INTEGER DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'draft',
                    n8n_workflow_id TEXT DEFAULT '',
                    n8n_execution_id TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    leads_discovered INTEGER DEFAULT 0,
                    leads_processed INTEGER DEFAULT 0,
                    leads_qualified INTEGER DEFAULT 0,
                    leads_suppressed INTEGER DEFAULT 0,
                    drafts_generated INTEGER DEFAULT 0,
                    awaiting_review INTEGER DEFAULT 0,
                    authorized_messages INTEGER DEFAULT 0,
                    failed_operations INTEGER DEFAULT 0,
                    current_stage TEXT DEFAULT 'Configured',
                    last_checkpoint TEXT DEFAULT '',
                    last_error TEXT DEFAULT '',
                    start_time TEXT,
                    completion_time TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS campaign_leads (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    campaign_id TEXT NOT NULL,
                    normalized_username TEXT NOT NULL,
                    username TEXT NOT NULL,
                    profile_url TEXT NOT NULL,
                    source_accounts TEXT NOT NULL,
                    intent_evidence TEXT DEFAULT '',
                    location_category TEXT NOT NULL,
                    location_confidence TEXT NOT NULL,
                    location_score INTEGER DEFAULT 0,
                    intent_score INTEGER DEFAULT 0,
                    lead_score INTEGER DEFAULT 0,
                    suggested_message TEXT NOT NULL,
                    campaign_status TEXT NOT NULL DEFAULT 'eligible',
                    approval_status TEXT NOT NULL DEFAULT 'pending',
                    contact_status TEXT NOT NULL DEFAULT 'not_contacted',
                    dnc_status INTEGER DEFAULT 0,
                    suppression_reason TEXT DEFAULT '',
                    last_action TEXT DEFAULT 'Scored & Ingested',
                    last_action_at TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(campaign_id, normalized_username),
                    FOREIGN KEY(campaign_id) REFERENCES campaigns(campaign_id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS campaign_executions (
                    execution_id TEXT PRIMARY KEY,
                    campaign_id TEXT NOT NULL,
                    trigger_source TEXT NOT NULL,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    metrics_json TEXT NOT NULL,
                    error_message TEXT DEFAULT '',
                    started_at TEXT NOT NULL,
                    completed_at TEXT,
                    FOREIGN KEY(campaign_id) REFERENCES campaigns(campaign_id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_camp_status ON campaigns(status);
                CREATE INDEX IF NOT EXISTS idx_campleads_campid ON campaign_leads(campaign_id);
                CREATE INDEX IF NOT EXISTS idx_campleads_score ON campaign_leads(lead_score);
                CREATE INDEX IF NOT EXISTS idx_campleads_appr ON campaign_leads(approval_status);
                CREATE INDEX IF NOT EXISTS idx_campleads_contact ON campaign_leads(contact_status);
            """)

    # -------------------------------------------------------------------------
    # Campaign CRUD & Validation
    # -------------------------------------------------------------------------
    def create_campaign(
        self,
        campaign_id: str,
        name: str,
        objective: str = "Audience Research & Outreach",
        target_location: str = "Mumbai",
        source_accounts: Optional[List[str]] = None,
        max_leads: int = 500,
        min_lead_score: int = 0,
        n8n_workflow_id: str = "",
        notes: str = "",
    ) -> Dict[str, Any]:
        """Creates and validates a new campaign record."""
        clean_id = re.sub(r"[^a-zA-Z0-9_-]", "_", campaign_id.strip())
        if not clean_id:
            raise CampaignValidationError("Campaign ID cannot be empty.")
        if not name.strip():
            raise CampaignValidationError("Campaign name cannot be empty.")

        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        sources_json = json.dumps(source_accounts or [])

        with self._get_connection() as conn:
            existing = conn.execute("SELECT campaign_id FROM campaigns WHERE campaign_id = ?", (clean_id,)).fetchone()
            if existing:
                raise CampaignValidationError(f"Campaign with ID '{clean_id}' already exists.")

            conn.execute("""
                INSERT INTO campaigns (
                    campaign_id, name, objective, target_location, source_accounts,
                    max_leads, min_lead_score, status, n8n_workflow_id, notes,
                    current_stage, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                clean_id, name.strip(), objective.strip(), target_location.strip(),
                sources_json, int(max_leads), int(min_lead_score),
                CAMPAIGN_STATUS_DRAFT, n8n_workflow_id.strip(), notes.strip(),
                "Configured", now_utc, now_utc
            ))

        return self.get_campaign(clean_id)

    def get_campaign(self, campaign_id: str) -> Optional[Dict[str, Any]]:
        """Retrieves a campaign by ID."""
        with self._get_connection() as conn:
            row = conn.execute("SELECT * FROM campaigns WHERE campaign_id = ?", (campaign_id,)).fetchone()
            if not row:
                return None
            data = dict(row)
            try:
                data["source_accounts"] = json.loads(data["source_accounts"])
            except Exception:
                data["source_accounts"] = []
            return data

    def list_campaigns(self) -> List[Dict[str, Any]]:
        """Lists all campaigns ordered by creation date descending."""
        with self._get_connection() as conn:
            rows = conn.execute("SELECT * FROM campaigns ORDER BY created_at DESC").fetchall()
            result = []
            for r in rows:
                item = dict(r)
                try:
                    item["source_accounts"] = json.loads(item["source_accounts"])
                except Exception:
                    item["source_accounts"] = []
                result.append(item)
            return result

    def update_campaign_status(
        self,
        campaign_id: str,
        status: str,
        current_stage: Optional[str] = None,
        last_error: str = "",
        checkpoint: str = "",
    ) -> bool:
        """Updates campaign status, stage, and error details."""
        if status not in ALL_CAMPAIGN_STATUSES:
            raise CampaignValidationError(f"Invalid status '{status}'")

        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            updates = ["status = ?", "updated_at = ?"]
            params = [status, now_utc]

            if current_stage:
                updates.append("current_stage = ?")
                params.append(current_stage)
            if last_error:
                updates.append("last_error = ?")
                params.append(last_error)
            if checkpoint:
                updates.append("last_checkpoint = ?")
                params.append(checkpoint)
            if status == CAMPAIGN_STATUS_COMPLETED:
                updates.append("completion_time = ?")
                params.append(now_utc)
            elif status == CAMPAIGN_STATUS_RUNNING and not conn.execute(
                "SELECT start_time FROM campaigns WHERE campaign_id = ?", (campaign_id,)
            ).fetchone()["start_time"]:
                updates.append("start_time = ?")
                params.append(now_utc)

            params.append(campaign_id)
            conn.execute(f"UPDATE campaigns SET {', '.join(updates)} WHERE campaign_id = ?", params)
        return True

    # -------------------------------------------------------------------------
    # Lead Scoring Formula
    # -------------------------------------------------------------------------
    @staticmethod
    def calculate_lead_score(
        location_tier: str,
        comments_texts: List[str],
        total_comments: int = 1,
    ) -> Tuple[int, int, int, str]:
        """
        Calculates separate location score (0-50) and intent score (0-50).
        Total score = location_score + intent_score (0-100).
        Returns: (total_score, location_score, intent_score, intent_evidence_summary)
        """
        # 1. Location Score (0 - 50)
        loc_scores = {
            "Likely Mumbai": 50,
            "Serves Mumbai": 35,
            "Possibly Mumbai": 25,
            "Not enough evidence": 10,
            "Likely outside Mumbai": 0,
        }
        location_score = loc_scores.get(location_tier, 10)

        # 2. Intent Score (0 - 50)
        intent_score = 0
        evidence_hits = []
        combined_text = " ".join(comments_texts)

        for pattern, pts, label in BUYING_INTENT_PATTERNS:
            match = re.search(pattern, combined_text)
            if match:
                intent_score += pts
                evidence_hits.append(f"{label} ('{match.group(0)}')")

        # Question presence signal (+10)
        if "?" in combined_text:
            intent_score += 10
            evidence_hits.append("Asked direct question ('?')")

        # Engagement frequency (+5 for 2+, +10 for 5+)
        if total_comments >= 5:
            intent_score += 10
            evidence_hits.append(f"High engagement ({total_comments} comments)")
        elif total_comments >= 2:
            intent_score += 5
            evidence_hits.append(f"Repeat engagement ({total_comments} comments)")

        # Cap intent score at 50
        intent_score = min(50, intent_score)

        total_score = min(100, location_score + intent_score)
        evidence_str = "; ".join(evidence_hits) if evidence_hits else "General interaction"

        return total_score, location_score, intent_score, evidence_str

    # -------------------------------------------------------------------------
    # End-to-End Campaign Execution Engine
    # -------------------------------------------------------------------------
    def execute_pipeline(
        self,
        campaign_id: str,
        trigger_source: str = "manual_gui",
        checkpoint_dir: Optional[Path] = None,
        mumbai_db_path: Optional[Path] = None,
        pause_event: Optional[threading.Event] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> Dict[str, Any]:
        """
        Executes the 10-stage campaign pipeline with checkpointing.
        Idempotent and safe to pause/resume.
        """
        camp = self.get_campaign(campaign_id)
        if not camp:
            raise CampaignValidationError(f"Campaign '{campaign_id}' not found.")

        # Guard against duplicate starts
        if camp["status"] == CAMPAIGN_STATUS_RUNNING:
            raise CampaignDuplicateExecutionError(f"Campaign '{campaign_id}' is already running.")

        exec_id = f"exec_{campaign_id}_{int(time.time())}_{uuid.uuid4().hex[:6]}"
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

        with self._get_connection() as conn:
            conn.execute("""
                INSERT INTO campaign_executions (
                    execution_id, campaign_id, trigger_source, status, stage, metrics_json, started_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (exec_id, campaign_id, trigger_source, "running", "Initiated", "{}", now_utc))

        self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 1: Validation")

        cp_dir = Path(checkpoint_dir) if checkpoint_dir else get_canonical_checkpoints_dir()
        m_db_path = Path(mumbai_db_path) if mumbai_db_path else get_canonical_mumbai_db_path()

        # Check for cancel or pause
        def _check_interruption(stage_name: str) -> Optional[str]:
            if cancel_event and cancel_event.is_set():
                self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_CANCELLED, current_stage=f"Cancelled at {stage_name}", checkpoint=stage_name)
                logger.info("Campaign %s cancelled at %s", campaign_id, stage_name)
                return CAMPAIGN_STATUS_CANCELLED
            if pause_event and pause_event.is_set():
                self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_PAUSED, current_stage=f"Paused at {stage_name}", checkpoint=stage_name)
                logger.info("Campaign %s paused at %s", campaign_id, stage_name)
                return CAMPAIGN_STATUS_PAUSED
            return None

        try:
            # -----------------------------------------------------------------
            # Stage 2: Source Discovery & Lead Harvesting
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 2: Source Discovery", checkpoint="stage_2")
            if interrupted := _check_interruption("stage_2"):
                return {"status": interrupted, "campaign_id": campaign_id}

            selected_sources = set(camp["source_accounts"]) if camp["source_accounts"] else None
            db_files = list(cp_dir.glob("*_checkpoint.db"))
            target_files = []
            for f in db_files:
                if f.name in ("mumbai_check_checkpoint.db", "lead_outreach.db"):
                    continue
                stem = f.stem.replace("_checkpoint", "")
                if selected_sources:
                    norm_sources = {normalize_username(s) for s in selected_sources}
                    if normalize_username(stem) not in norm_sources:
                        continue
                target_files.append((f, stem))

            leads_raw: Dict[str, Dict[str, Any]] = {}
            total_discovered = 0

            for db_file, brand_name in target_files:
                try:
                    with sqlite3.connect(str(db_file), timeout=15.0) as sconn:
                        sconn.row_factory = sqlite3.Row
                        table_check = sconn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='comments'").fetchone()
                        if not table_check:
                            continue
                        rows = sconn.execute("SELECT username, shortcode, text, record_type FROM comments").fetchall()
                        for r in rows:
                            u = r["username"]
                            if not u:
                                continue
                            norm_u = normalize_username(u)
                            # Brand self-exclusion
                            if norm_u == normalize_username(brand_name):
                                continue

                            total_discovered += 1
                            if norm_u not in leads_raw:
                                leads_raw[norm_u] = {
                                    "username": u.strip().lstrip("@"),
                                    "normalized_username": norm_u,
                                    "source_accounts": set(),
                                    "comments_texts": [],
                                    "total_comments": 0,
                                }
                            item = leads_raw[norm_u]
                            item["source_accounts"].add(brand_name)
                            item["total_comments"] += 1
                            if r["text"]:
                                item["comments_texts"].append(r["text"].strip())
                except Exception as e:
                    logger.error("Error reading source checkpoint %s: %s", db_file, e)

            # Respect max_leads constraint
            all_discovered_leads = list(leads_raw.values())
            max_limit = camp["max_leads"] if camp["max_leads"] > 0 else len(all_discovered_leads)
            target_leads = all_discovered_leads[:max_limit]

            # -----------------------------------------------------------------
            # Stage 3 & 4: Deduplication & Mumbai Qualification
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 3/4: Qualification & Deduplication", checkpoint="stage_4")
            if interrupted := _check_interruption("stage_4"):
                return {"status": interrupted, "campaign_id": campaign_id}

            mumbai_cache: Dict[str, Dict[str, Any]] = {}
            if m_db_path.exists():
                try:
                    with sqlite3.connect(str(m_db_path), timeout=15.0) as mconn:
                        mconn.row_factory = sqlite3.Row
                        rows = mconn.execute(
                            "SELECT normalized_username, checked_username, classification, confidence, "
                            "stated_location, evidence_summary, has_error FROM mumbai_checks"
                        ).fetchall()
                        for r in rows:
                            norm = r["normalized_username"] or normalize_username(r["checked_username"])
                            mumbai_cache[norm] = dict(r)
                except Exception as e:
                    logger.warning("Could not read Mumbai cache: %s", e)

            dnc_map = self.lead_outreach_db.get_dnc_map()

            # -----------------------------------------------------------------
            # Stage 5, 6, 7: Scoring, Draft Generation & DNC Suppression
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 5/6: Scoring & Drafting", checkpoint="stage_6")
            if interrupted := _check_interruption("stage_6"):
                return {"status": interrupted, "campaign_id": campaign_id}

            qualified_count = 0
            suppressed_count = 0
            drafts_count = 0
            awaiting_review_count = 0

            leads_to_persist = []
            target_loc = camp["target_location"].strip().lower()
            min_score = camp["min_lead_score"]

            for lead_info in target_leads:
                norm_u = lead_info["normalized_username"]
                u = lead_info["username"]
                profile_url = f"https://www.instagram.com/{u}/"
                srcs_json = json.dumps(sorted(list(lead_info["source_accounts"])))

                m_data = mumbai_cache.get(norm_u, {})
                classification = m_data.get("classification", "Not enough evidence")
                confidence = m_data.get("confidence", "Low")
                has_error = bool(m_data.get("has_error", 0))

                # Error shielding: transient failure != outside Mumbai
                if has_error and classification == "Likely outside Mumbai":
                    classification = "Not enough evidence"

                # Calculate Scores
                tot_score, loc_score, intent_score, intent_evid = self.calculate_lead_score(
                    location_tier=classification,
                    comments_texts=lead_info["comments_texts"],
                    total_comments=lead_info["total_comments"],
                )

                # DNC check
                is_dnc = 1 if norm_u in dnc_map else 0
                suppression_reason = ""
                camp_status = "eligible"

                if is_dnc:
                    camp_status = "suppressed"
                    suppression_reason = f"Suppressed by DNC: {dnc_map.get(norm_u, 'Requested opt-out')}"
                    suppressed_count += 1
                elif "mumbai" in target_loc and classification == "Likely outside Mumbai":
                    camp_status = "suppressed"
                    suppression_reason = "Suppressed: Confirmed outside Mumbai target area"
                    suppressed_count += 1
                elif tot_score < min_score:
                    camp_status = "suppressed"
                    suppression_reason = f"Score ({tot_score}) below campaign minimum ({min_score})"
                    suppressed_count += 1
                else:
                    qualified_count += 1
                    awaiting_review_count += 1

                # Message Drafting
                if classification in ("Likely Mumbai", "Serves Mumbai", "Possibly Mumbai"):
                    suggested_msg = MUMBAI_OUTREACH_TEMPLATE
                else:
                    suggested_msg = GENERIC_OUTREACH_TEMPLATE
                drafts_count += 1

                leads_to_persist.append({
                    "campaign_id": campaign_id,
                    "normalized_username": norm_u,
                    "username": u,
                    "profile_url": profile_url,
                    "source_accounts": srcs_json,
                    "intent_evidence": intent_evid,
                    "location_category": classification,
                    "location_confidence": confidence,
                    "location_score": loc_score,
                    "intent_score": intent_score,
                    "lead_score": tot_score,
                    "suggested_message": suggested_msg,
                    "campaign_status": camp_status,
                    "approval_status": "pending",
                    "contact_status": "not_contacted",
                    "dnc_status": is_dnc,
                    "suppression_reason": suppression_reason,
                    "last_action": "Scored & Ingested",
                })

            # -----------------------------------------------------------------
            # Stage 8: Database Persistence
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 8: Database Persistence", checkpoint="stage_8")
            if interrupted := _check_interruption("stage_8"):
                return {"status": interrupted, "campaign_id": campaign_id}

            with self._get_connection() as conn:
                for item in leads_to_persist:
                    conn.execute("""
                        INSERT INTO campaign_leads (
                            campaign_id, normalized_username, username, profile_url,
                            source_accounts, intent_evidence, location_category,
                            location_confidence, location_score, intent_score, lead_score,
                            suggested_message, campaign_status, approval_status,
                            contact_status, dnc_status, suppression_reason, last_action,
                            last_action_at, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(campaign_id, normalized_username) DO UPDATE SET
                            source_accounts = excluded.source_accounts,
                            intent_evidence = excluded.intent_evidence,
                            location_category = excluded.location_category,
                            location_confidence = excluded.location_confidence,
                            location_score = excluded.location_score,
                            intent_score = excluded.intent_score,
                            lead_score = excluded.lead_score,
                            campaign_status = excluded.campaign_status,
                            dnc_status = excluded.dnc_status,
                            suppression_reason = excluded.suppression_reason,
                            updated_at = excluded.updated_at
                    """, (
                        item["campaign_id"], item["normalized_username"], item["username"],
                        item["profile_url"], item["source_accounts"], item["intent_evidence"],
                        item["location_category"], item["location_confidence"],
                        item["location_score"], item["intent_score"], item["lead_score"],
                        item["suggested_message"], item["campaign_status"],
                        item["approval_status"], item["contact_status"], item["dnc_status"],
                        item["suppression_reason"], item["last_action"], now_utc, now_utc, now_utc
                    ))

                # Update campaign summary metrics
                conn.execute("""
                    UPDATE campaigns SET
                        leads_discovered = ?,
                        leads_processed = ?,
                        leads_qualified = ?,
                        leads_suppressed = ?,
                        drafts_generated = ?,
                        awaiting_review = ?,
                        updated_at = ?
                    WHERE campaign_id = ?
                """, (
                    total_discovered, len(leads_to_persist), qualified_count,
                    suppressed_count, drafts_count, awaiting_review_count,
                    now_utc, campaign_id
                ))

            # -----------------------------------------------------------------
            # Stage 9: Google Sheets & Excel Sync
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_RUNNING, current_stage="Stage 9: Syncing Sheets & Excel", checkpoint="stage_9")
            if interrupted := _check_interruption("stage_9"):
                return {"status": interrupted, "campaign_id": campaign_id}

            try:
                self.sync_campaign_to_excel(campaign_id)
            except Exception as e:
                logger.error("Campaign Excel sync error: %s", e)

            try:
                self.sync_campaign_to_google_sheets(campaign_id)
            except Exception as e:
                logger.warning("Campaign Google Sheets sync note: %s", e)

            # -----------------------------------------------------------------
            # Stage 10: Completion
            # -----------------------------------------------------------------
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_COMPLETED, current_stage="Completed", checkpoint="completed")

            with self._get_connection() as conn:
                metrics_summary = {
                    "leads_discovered": total_discovered,
                    "leads_processed": len(leads_to_persist),
                    "leads_qualified": qualified_count,
                    "leads_suppressed": suppressed_count,
                    "drafts_generated": drafts_count,
                    "awaiting_review": awaiting_review_count,
                }
                conn.execute("""
                    UPDATE campaign_executions SET
                        status = 'completed', stage = 'Completed', metrics_json = ?, completed_at = ?
                    WHERE execution_id = ?
                """, (json.dumps(metrics_summary), now_utc, exec_id))

            logger.info("Campaign %s completed successfully: %s", campaign_id, metrics_summary)
            return {
                "status": "completed",
                "campaign_id": campaign_id,
                "execution_id": exec_id,
                "metrics": metrics_summary,
            }

        except Exception as e:
            err_msg = str(e)
            logger.exception("Error during campaign execution %s: %s", campaign_id, e)
            self.update_campaign_status(campaign_id, CAMPAIGN_STATUS_FAILED, current_stage="Execution Error", last_error=err_msg)
            with self._get_connection() as conn:
                conn.execute("""
                    UPDATE campaign_executions SET status = 'failed', error_message = ?, completed_at = ?
                    WHERE execution_id = ?
                """, (err_msg, now_utc, exec_id))
            raise

    # -------------------------------------------------------------------------
    # Lead Review Queue Actions
    # -------------------------------------------------------------------------
    def get_campaign_leads_paginated(
        self,
        campaign_id: str,
        limit: int = 25,
        offset: int = 0,
        search: Optional[str] = "",
        approval_status: Optional[str] = None,
        contact_status: Optional[str] = None,
        campaign_status: Optional[str] = None,
        min_score: int = 0,
        sort_by: str = "lead_score DESC",
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Queries campaign leads using SQL-level pagination and filtering."""
        where_clauses = ["campaign_id = ?"]
        params: List[Any] = [campaign_id]

        if search and search.strip():
            term = f"%{search.strip().lower()}%"
            where_clauses.append("(normalized_username LIKE ? OR intent_evidence LIKE ?)")
            params.extend([term, term])

        if approval_status and approval_status != "All":
            where_clauses.append("approval_status = ?")
            params.append(approval_status.lower())

        if contact_status and contact_status != "All":
            where_clauses.append("contact_status = ?")
            params.append(contact_status.lower())

        if campaign_status and campaign_status != "All":
            where_clauses.append("campaign_status = ?")
            params.append(campaign_status.lower())

        if min_score > 0:
            where_clauses.append("lead_score >= ?")
            params.append(min_score)

        where_sql = f"WHERE {' AND '.join(where_clauses)}"

        allowed_sorts = {
            "lead_score DESC": "lead_score DESC",
            "lead_score ASC": "lead_score ASC",
            "username ASC": "username ASC",
            "last_action_at DESC": "last_action_at DESC",
        }
        order_sql = allowed_sorts.get(sort_by, "lead_score DESC")

        with self._get_connection() as conn:
            cnt_query = f"SELECT COUNT(*) FROM campaign_leads {where_sql}"
            total_count = conn.execute(cnt_query, params).fetchone()[0]

            data_query = f"""
                SELECT * FROM campaign_leads
                {where_sql}
                ORDER BY {order_sql}
                LIMIT ? OFFSET ?
            """
            rows = conn.execute(data_query, params + [limit, offset]).fetchall()
            return [dict(r) for r in rows], total_count

    def approve_lead_draft(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        """Approves a draft message. Note: Approving a draft MUST NOT send it."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    approval_status = 'approved',
                    last_action = 'Draft Approved',
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (now_utc, now_utc, campaign_id, norm))
            # Also log in central outreach audit log
            conn.execute("""
                INSERT INTO outreach_audit_log (
                    normalized_username, action, previous_status, new_status, notes, operator, created_at
                ) VALUES (?, 'CAMPAIGN_APPROVE', 'pending', 'approved', ?, ?, ?)
            """, (norm, f"Approved for campaign {campaign_id}", operator, now_utc))
        return True

    def edit_lead_draft(self, campaign_id: str, username: str, new_draft: str, operator: str = "Operator") -> bool:
        """Saves custom edits to a suggested message draft."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    suggested_message = ?,
                    last_action = 'Draft Edited',
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (new_draft.strip(), now_utc, now_utc, campaign_id, norm))
            conn.execute("""
                INSERT INTO outreach_audit_log (
                    normalized_username, action, previous_status, new_status, notes, operator, created_at
                ) VALUES (?, 'CAMPAIGN_EDIT_DRAFT', 'draft', 'draft', ?, ?, ?)
            """, (norm, f"Custom draft updated for campaign {campaign_id}", operator, now_utc))
        return True

    def reject_lead_draft(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        """Rejects a draft message."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    approval_status = 'rejected',
                    last_action = 'Draft Rejected',
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (now_utc, now_utc, campaign_id, norm))
            conn.execute("""
                INSERT INTO outreach_audit_log (
                    normalized_username, action, previous_status, new_status, notes, operator, created_at
                ) VALUES (?, 'CAMPAIGN_REJECT', 'pending', 'rejected', ?, ?, ?)
            """, (norm, f"Rejected in campaign {campaign_id}", operator, now_utc))
        return True

    def update_lead_contact_status(
        self, campaign_id: str, username: str, contact_status: str, notes: str = "", operator: str = "Operator"
    ) -> bool:
        """Updates contact status (contacted, replied, interested, not_interested)."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    contact_status = ?,
                    last_action = ?,
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (contact_status, f"Status -> {contact_status}", now_utc, now_utc, campaign_id, norm))
            conn.execute("""
                INSERT INTO outreach_audit_log (
                    normalized_username, action, previous_status, new_status, notes, operator, created_at
                ) VALUES (?, 'CAMPAIGN_CONTACT_STATUS', 'not_contacted', ?, ?, ?, ?)
            """, (norm, contact_status, f"Campaign {campaign_id}: {notes}", operator, now_utc))
        return True

    def add_lead_to_dnc(self, campaign_id: str, username: str, reason: str = "Operator suppression", operator: str = "Operator") -> bool:
        """Adds lead to persistent Do Not Contact registry and suppresses across campaigns."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        # Add to global DNC
        self.lead_outreach_db.add_to_dnc(username, reason=reason, operator=operator)
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    dnc_status = 1,
                    campaign_status = 'suppressed',
                    suppression_reason = ?,
                    last_action = 'Added to DNC',
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (f"Suppressed by DNC: {reason}", now_utc, now_utc, campaign_id, norm))
        return True

    def remove_lead_from_campaign(self, campaign_id: str, username: str, operator: str = "Operator") -> bool:
        """Removes a lead from an active campaign."""
        norm = normalize_username(username)
        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE campaign_leads SET
                    campaign_status = 'removed',
                    last_action = 'Removed from Campaign',
                    last_action_at = ?,
                    updated_at = ?
                WHERE campaign_id = ? AND normalized_username = ?
            """, (now_utc, now_utc, campaign_id, norm))
        return True

    # -------------------------------------------------------------------------
    # Synchronization (Google Sheets & Excel Mirror)
    # -------------------------------------------------------------------------
    def sync_campaign_to_excel(self, campaign_id: str, excel_path: str = DEFAULT_EXCEL_PATH) -> int:
        """Syncs campaign leads to a dedicated 'Campaign Leads' worksheet in the Excel workbook."""
        import openpyxl
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        target = Path(excel_path)
        target.parent.mkdir(parents=True, exist_ok=True)

        if target.exists():
            wb = openpyxl.load_workbook(str(target))
        else:
            wb = openpyxl.Workbook()
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

        sheet_title = "Campaign Leads"
        if sheet_title in wb.sheetnames:
            ws = wb[sheet_title]
            ws.delete_rows(1, ws.max_row + 1)
        else:
            ws = wb.create_sheet(title=sheet_title)

        with self._get_connection() as conn:
            leads = conn.execute(
                "SELECT * FROM campaign_leads WHERE campaign_id = ? ORDER BY lead_score DESC", (campaign_id,)
            ).fetchall()

        now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

        # Header Styles
        font_title = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_title = PatternFill("solid", fgColor="1F4E78")
        font_hdr = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        fill_hdr = PatternFill("solid", fgColor="2F343D")
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
        ws.cell(row=1, column=1, value=f"CAMPAIGN LEADS: {campaign_id}").fill = fill_title
        ws.cell(row=1, column=1).font = font_title
        ws.cell(row=1, column=2, value=f"Last Updated: {now_str}").fill = fill_title
        ws.cell(row=1, column=2).font = font_title
        ws.cell(row=1, column=3, value=f"Total Leads: {len(leads)}").fill = fill_title
        ws.cell(row=1, column=3).font = font_title

        headers = [
            "Campaign ID", "Username", "Profile URL", "Direct Chat (ig.me)",
            "Lead Score", "Location Tier", "Location Conf", "Intent Evidence",
            "Campaign Status", "Approval Status", "Contact Status", "Suggested Message",
            "Last Action", "Updated UTC"
        ]
        ws.row_dimensions[3].height = 22
        for col_i, h in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_i, value=h)
            cell.font = font_hdr
            cell.fill = fill_hdr
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", horizontal="center" if col_i in (5, 9, 10, 11) else "left")

        for r_idx, r in enumerate(leads, start=4):
            ws.row_dimensions[r_idx].height = 20
            u = r["username"]
            p_url = r["profile_url"]
            dm_url = f"https://ig.me/m/{u}"

            vals = [
                campaign_id,
                f"@{u}",
                p_url,
                dm_url,
                r["lead_score"],
                r["location_category"],
                r["location_confidence"],
                r["intent_evidence"],
                r["campaign_status"],
                r["approval_status"],
                r["contact_status"],
                r["suggested_message"],
                r["last_action"],
                r["updated_at"],
            ]
            for c_idx, val in enumerate(vals, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", horizontal="center" if c_idx in (5, 9, 10, 11, 14) else "left")
                if c_idx in (3, 4):
                    cell.hyperlink = val
                    cell.font = font_link

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col if cell.row > 1)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 45))

        wb.save(str(target))
        logger.info("Synced %d campaign leads to %s -> %s", len(leads), target, sheet_title)
        return len(leads)

    def sync_campaign_to_google_sheets(
        self, campaign_id: str, spreadsheet_id: str = DEFAULT_SPREADSHEET_ID, credentials_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """Syncs campaign leads to Google Spreadsheet master under 'Campaign Leads' tab."""
        from sheets_sync import GoogleSheetsManager

        with self._get_connection() as conn:
            leads = conn.execute(
                "SELECT * FROM campaign_leads WHERE campaign_id = ? ORDER BY lead_score DESC", (campaign_id,)
            ).fetchall()

        mgr = GoogleSheetsManager(spreadsheet_id=spreadsheet_id)
        service = mgr.get_service(credentials_path=credentials_path)

        sheet_title = "Campaign Leads"
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets", [])}

        if sheet_title not in sheet_map:
            add_req = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": sheet_title,
                                "gridProperties": {"rowCount": max(100, len(leads) + 15), "columnCount": 16},
                            }
                        }
                    }
                ]
            }
            res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=add_req).execute()
            sheet_id = res["replies"][0]["addSheet"]["properties"]["sheetId"]
        else:
            sheet_id = sheet_map[sheet_title]

        now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        headers = [
            "Campaign ID", "Username", "Profile URL", "Direct Chat (ig.me)",
            "Lead Score", "Location Tier", "Location Conf", "Intent Evidence",
            "Campaign Status", "Approval Status", "Contact Status", "Suggested Message",
            "Last Action", "Updated UTC"
        ]

        values = [
            [f"CAMPAIGN LEADS: {campaign_id}", f"Last Updated: {now_str}", f"Total Leads: {len(leads)}"],
            [],
            headers,
        ]

        for r in leads:
            u = r["username"]
            p_url = r["profile_url"]
            dm_url = f"https://ig.me/m/{u}"
            values.append([
                campaign_id,
                f'=HYPERLINK("{p_url}", "@{u}")',
                p_url,
                f'=HYPERLINK("{dm_url}", "Direct Chat")',
                r["lead_score"],
                r["location_category"],
                r["location_confidence"],
                r["intent_evidence"],
                r["campaign_status"],
                r["approval_status"],
                r["contact_status"],
                r["suggested_message"],
                r["last_action"],
                r["updated_at"],
            ])

        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id, range=f"'{sheet_title}'!A1:Z{max(100, len(values) + 20)}"
        ).execute()

        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values},
        ).execute()

        fmt_requests = [
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 3}},
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 14},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
                            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 11},
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)",
                }
            },
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 14},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.18, "green": 0.20, "blue": 0.24},
                            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 10},
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)",
                }
            },
        ]
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": fmt_requests}).execute()

        return {
            "spreadsheet_id": spreadsheet_id,
            "sheet_title": sheet_title,
            "rows_synced": len(leads),
            "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={sheet_id}",
        }
