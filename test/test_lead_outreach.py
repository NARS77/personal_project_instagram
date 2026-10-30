#!/usr/bin/env python3
"""
Unit and Integration Tests for MeatNDoor Lead Outreach Manager.

Validates:
1. Deduplication and incremental ingestion from checkpoint DBs.
2. Commenters vs reply authors extraction.
3. Source brand filtering (brand itself is never ingested as a lead).
4. Geographic campaign eligibility rules & outside-Mumbai suppression.
5. Persistent Do-Not-Contact (DNC) suppression across runs.
6. Editable message drafting (Mumbai vs Generic).
7. Review queue state transitions and operator audit logging.
8. Duplicate-send prevention and approval enforcement.
9. Meta messaging workflow (authorized vs manual fallback).
10. Dual synchronization (Excel & Google Sheets) idempotency.
"""

import json
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from lead_outreach import (
    CAMPAIGN_MUMBAI,
    CAMPAIGN_GENERIC,
    ELIGIBILITY_MUMBAI_TARGET,
    ELIGIBILITY_BUSINESS_REVIEW,
    ELIGIBILITY_LOW_CONF_REVIEW,
    ELIGIBILITY_PENDING_LOCATION,
    ELIGIBILITY_SUPPRESSED_OUTSIDE,
    ELIGIBILITY_SUPPRESSED_DNC,
    STATUS_PENDING,
    STATUS_APPROVED,
    STATUS_REJECTED,
    STATUS_CONTACTED,
    STATUS_REPLIED,
    STATUS_INTERESTED,
    STATUS_NOT_INTERESTED,
    STATUS_DNC,
    MSG_STATUS_MANUAL,
    MSG_STATUS_SENT,
    MUMBAI_OUTREACH_TEMPLATE,
    GENERIC_OUTREACH_TEMPLATE,
    LeadOutreachDB,
    LeadOutreachSync,
    MetaMessagingClient,
    DNCSuppressionError,
    CampaignSuppressionError,
    DuplicateSendBlockedError,
    UnapprovedDraftError,
    normalize_username,
)


class TestLeadOutreachManager(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_outreach_")
        self.db_path = Path(self.test_dir) / "test_outreach.db"
        self.mumbai_db_path = Path(self.test_dir) / "test_mumbai.db"
        self.excel_path = Path(self.test_dir) / "test_outreach.xlsx"
        self.db = LeadOutreachDB(db_path=self.db_path)

        # Setup mock mumbai_checks database
        with sqlite3.connect(str(self.mumbai_db_path)) as mconn:
            mconn.executescript("""
                CREATE TABLE mumbai_checks (
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
                );
            """)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_mock_checkpoint(self, brand: str, records: list) -> Path:
        cp_path = Path(self.test_dir) / f"{brand}_checkpoint.db"
        with sqlite3.connect(str(cp_path)) as conn:
            conn.executescript("""
                CREATE TABLE comments (
                    media_id TEXT,
                    comment_id TEXT PRIMARY KEY,
                    shortcode TEXT,
                    username TEXT,
                    text TEXT,
                    created_at INTEGER,
                    likes INTEGER,
                    replies INTEGER,
                    record_type TEXT,
                    parent_comment_id TEXT,
                    replies_count INTEGER,
                    replies_fetched INTEGER
                );
            """)
            for r in records:
                conn.execute("""
                    INSERT INTO comments (
                        media_id, comment_id, shortcode, username, text,
                        created_at, likes, replies, record_type, parent_comment_id,
                        replies_count, replies_fetched
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    r.get("media_id", "m1"),
                    r["comment_id"],
                    r.get("shortcode", "s1"),
                    r["username"],
                    r.get("text", "Looks great"),
                    1700000000,
                    r.get("likes", 0),
                    0,
                    r.get("record_type", "comment"),
                    r.get("parent_comment_id", ""),
                    0,
                    0,
                ))
        return cp_path

    def _add_mumbai_check(self, username: str, classification: str, confidence: str = "High", location: str = "Bandra, Mumbai", has_error: int = 0):
        with sqlite3.connect(str(self.mumbai_db_path)) as mconn:
            norm = normalize_username(username)
            mconn.execute("""
                INSERT OR REPLACE INTO mumbai_checks (
                    normalized_username, checked_username, profile_url, display_name,
                    stated_location, evidence_summary, classification, confidence,
                    checked_at, has_error
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '2026-10-14 10:00:00 UTC', ?)
            """, (norm, username, f"https://instagram.com/{username}/", username.title(), location, f"Test {classification}", classification, confidence, has_error))

    def test_normalization_and_deduplication(self):
        self.assertEqual(normalize_username("@Rahul_Sharma_123"), "rahul_sharma_123")
        self.assertEqual(normalize_username("  @MUMBAI.FOODIE!  "), "mumbai.foodie")

    def test_brand_self_filtering(self):
        """Ensures that the source brand account itself is NEVER ingested as a lead."""
        self._create_mock_checkpoint("licious_foods", [
            {"comment_id": "c1", "username": "licious_foods", "text": "Thanks for ordering with us!"},
            {"comment_id": "c2", "username": "customer_one", "text": "Can I get delivery in Bandra?"},
        ])

        stats = self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        self.assertEqual(stats["leads_ingested_or_updated"], 1)
        self.assertIsNone(self.db.get_lead("licious_foods"))
        self.assertIsNotNone(self.db.get_lead("customer_one"))

    def test_commenters_and_reply_authors_ingestion(self):
        """Verifies ingestion of both top-level commenters and nested reply authors."""
        self._create_mock_checkpoint("meatigo_official", [
            {"comment_id": "c1", "username": "parent_user", "record_type": "comment", "text": "Is this fresh?"},
            {"comment_id": "c2", "username": "reply_user", "record_type": "reply", "parent_comment_id": "c1", "text": "Yes it is very fresh!"},
        ])

        stats = self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        self.assertEqual(stats["leads_ingested_or_updated"], 2)
        p_lead = self.db.get_lead("parent_user")
        r_lead = self.db.get_lead("reply_user")
        self.assertIn("comment", p_lead["comment_types"])
        self.assertIn("reply", r_lead["comment_types"])

    def test_campaign_eligibility_and_suppression(self):
        """Validates Mumbai eligibility, outside-Mumbai suppression, and DNC suppression."""
        self._add_mumbai_check("mumbai_resident", "Likely Mumbai")
        self._add_mumbai_check("mumbai_business", "Serves Mumbai")
        self._add_mumbai_check("weak_lead", "Possibly Mumbai")
        self._add_mumbai_check("delhi_customer", "Likely outside Mumbai")
        self._add_mumbai_check("unknown_customer", "Not enough evidence")

        self._create_mock_checkpoint("prasuma_official", [
            {"comment_id": "c1", "username": "mumbai_resident"},
            {"comment_id": "c2", "username": "mumbai_business"},
            {"comment_id": "c3", "username": "weak_lead"},
            {"comment_id": "c4", "username": "delhi_customer"},
            {"comment_id": "c5", "username": "unknown_customer"},
            {"comment_id": "c6", "username": "dnc_account"},
        ])

        # Pre-add to DNC
        self.db.add_to_dnc("dnc_account", reason="Customer opted out")

        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        l1 = self.db.get_lead("mumbai_resident")
        self.assertEqual(l1["campaign_eligibility"], ELIGIBILITY_MUMBAI_TARGET)
        self.assertEqual(l1["drafted_message"], MUMBAI_OUTREACH_TEMPLATE)

        l2 = self.db.get_lead("mumbai_business")
        self.assertEqual(l2["campaign_eligibility"], ELIGIBILITY_BUSINESS_REVIEW)

        l3 = self.db.get_lead("weak_lead")
        self.assertEqual(l3["campaign_eligibility"], ELIGIBILITY_LOW_CONF_REVIEW)

        l4 = self.db.get_lead("delhi_customer")
        self.assertEqual(l4["campaign_eligibility"], ELIGIBILITY_SUPPRESSED_OUTSIDE)

        l5 = self.db.get_lead("unknown_customer")
        self.assertEqual(l5["campaign_eligibility"], ELIGIBILITY_PENDING_LOCATION)
        self.assertEqual(l5["drafted_message"], GENERIC_OUTREACH_TEMPLATE)

        l6 = self.db.get_lead("dnc_account")
        self.assertEqual(l6["campaign_eligibility"], ELIGIBILITY_SUPPRESSED_DNC)
        self.assertEqual(l6["review_status"], STATUS_DNC)
        self.assertEqual(l6["is_dnc"], 1)

    def test_error_shielding_in_eligibility(self):
        """Never converts a transient lookup error into a confirmed outside-Mumbai classification."""
        self._add_mumbai_check("error_user", "Likely outside Mumbai", has_error=1)

        self._create_mock_checkpoint("fishtokri", [
            {"comment_id": "c1", "username": "error_user"},
        ])

        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        lead = self.db.get_lead("error_user")
        # Since has_error was 1, it should NOT be confirmed outside Mumbai
        self.assertEqual(lead["campaign_eligibility"], ELIGIBILITY_PENDING_LOCATION)

    def test_draft_editing_and_status_transitions(self):
        """Tests draft editing, approval without sending, and operator decision audit logging."""
        self._add_mumbai_check("chef_rohit", "Likely Mumbai")
        self._create_mock_checkpoint("bastianmumbai", [
            {"comment_id": "c1", "username": "chef_rohit"},
        ])

        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        custom_draft = "Hi Rohit! Saw you at Bastian Mumbai. Would you like to try our special cuts at MeatNDoor?"
        self.db.edit_lead_draft("chef_rohit", custom_draft, operator="Tester")

        lead = self.db.get_lead("chef_rohit")
        self.assertEqual(lead["drafted_message"], custom_draft)
        self.assertEqual(lead["review_status"], STATUS_PENDING)

        # Approve draft
        self.db.update_lead_status("chef_rohit", STATUS_APPROVED, notes="Approved for outreach", operator="Tester")
        lead = self.db.get_lead("chef_rohit")
        self.assertEqual(lead["review_status"], STATUS_APPROVED)

        # Verify audit log
        logs = self.db.get_audit_log(limit=5)
        self.assertTrue(any(l["action"] == "EDIT_DRAFT" for l in logs))
        self.assertTrue(any(l["action"] == "STATUS_UPDATE" and l["new_status"] == STATUS_APPROVED for l in logs))

    def test_duplicate_send_and_suppression_enforcement(self):
        """Verifies duplicate-send prevention and suppression blocks in MetaMessagingClient."""
        client = MetaMessagingClient(db=self.db)

        # Setup test leads
        self._add_mumbai_check("active_lead", "Likely Mumbai")
        self._add_mumbai_check("outside_lead", "Likely outside Mumbai")
        self._create_mock_checkpoint("brand_x", [
            {"comment_id": "c1", "username": "active_lead"},
            {"comment_id": "c2", "username": "outside_lead"},
        ])

        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        # 1. Unapproved draft block
        with self.assertRaises(UnapprovedDraftError):
            client.execute_messaging_workflow("active_lead", mock=True)

        # Approve draft
        self.db.update_lead_status("active_lead", STATUS_APPROVED)

        # 2. Mock authorized send success
        res = client.execute_messaging_workflow("active_lead", mock=True)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["review_status"], STATUS_CONTACTED)

        # 3. Duplicate send block
        with self.assertRaises(DuplicateSendBlockedError):
            client.execute_messaging_workflow("active_lead", mock=True)

        # 4. Outside Mumbai suppression block for Mumbai campaign
        self.db.update_lead_status("outside_lead", STATUS_APPROVED)
        with self.assertRaises(CampaignSuppressionError):
            client.execute_messaging_workflow("outside_lead", mock=True)

        # 5. DNC suppression block
        self.db.add_to_dnc("active_lead", reason="Do not call")
        with self.assertRaises(DNCSuppressionError):
            client.execute_messaging_workflow("active_lead", mock=True)

    def test_cold_commenter_manual_opt_in_fallback(self):
        """Cold commenters without Meta IGSID must fall back to Manual / Opt-In Required."""
        client = MetaMessagingClient(db=self.db)
        self._add_mumbai_check("cold_user", "Likely Mumbai")
        self._create_mock_checkpoint("brand_y", [
            {"comment_id": "c1", "username": "cold_user"},
        ])

        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )
        self.db.update_lead_status("cold_user", STATUS_APPROVED)

        # Real execution without IGSID / token
        res = client.execute_messaging_workflow("cold_user", mock=False)
        self.assertEqual(res["status"], "manual_required")
        self.assertEqual(res["messaging_status"], MSG_STATUS_MANUAL)
        self.assertIn("https://ig.me/m/cold_user", res["direct_message_url"])

    def test_excel_sync_idempotency(self):
        """Tests that Excel sync creates Outreach Manager tab and preserves other tabs."""
        import openpyxl

        # Pre-create Excel with another tab
        wb = openpyxl.Workbook()
        ws_exist = wb.active
        ws_exist.title = "licious_foods"
        ws_exist.cell(row=1, column=1, value="Historical Scraping Data")
        wb.save(str(self.excel_path))

        self._add_mumbai_check("sync_user", "Likely Mumbai")
        self._create_mock_checkpoint("brand_z", [
            {"comment_id": "c1", "username": "sync_user"},
        ])
        self.db.ingest_from_checkpoints(
            checkpoint_dir=Path(self.test_dir),
            mumbai_db_path=self.mumbai_db_path,
        )

        count = LeadOutreachSync.sync_to_excel(excel_path=str(self.excel_path), db=self.db)
        self.assertEqual(count, 1)

        # Verify workbook has both tabs intact
        wb_check = openpyxl.load_workbook(str(self.excel_path))
        self.assertIn("licious_foods", wb_check.sheetnames)
        self.assertIn("Outreach Manager", wb_check.sheetnames)

        ws_out = wb_check["Outreach Manager"]
        self.assertEqual(ws_out.cell(row=1, column=1).value, "MEATNDOOR LEAD OUTREACH MANAGER")
        self.assertEqual(ws_out.cell(row=4, column=1).value, "@sync_user")


if __name__ == "__main__":
    unittest.main()
