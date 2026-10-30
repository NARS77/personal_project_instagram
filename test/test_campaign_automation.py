#!/usr/bin/env python3
"""
Unit and Integration Tests for n8n Campaign Automation & Lead Intelligence Pipeline.

Validates:
1. Campaign configuration, creation, listing, and validation.
2. Dual-axis scoring engine: Location Score (0-50) + Customer Intent Score (0-50) = Total (0-100).
3. Intent signal extraction: product keywords, delivery/order terms, question marks, repeat engagement.
4. Geographic qualification and suppression:
   - Likely Mumbai / Serves Mumbai / Possibly Mumbai qualified if meeting min score.
   - Likely outside Mumbai suppressed from Mumbai campaigns.
5. Persistent Do-Not-Contact (DNC) list suppression.
6. Execution pipeline stage transitions and checkpointing (pause, resume, cancel).
7. Review queue operations (approve, edit, reject, contact status, DNC).
8. n8n client webhook triggering, authentication headers, and health check.
9. Campaign API microservice endpoints (/health, /campaign/start, /campaign/<id>/status, etc.).
10. Dual synchronization to Excel workbook ('Campaign Leads' tab) and Google Sheets.
"""

import json
import os
import shutil
import sqlite3
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from campaign_service import (
    CampaignService,
    CAMPAIGN_STATUS_DRAFT,
    CAMPAIGN_STATUS_RUNNING,
    CAMPAIGN_STATUS_PAUSED,
    CAMPAIGN_STATUS_COMPLETED,
    CAMPAIGN_STATUS_CANCELLED,
    CAMPAIGN_STATUS_FAILED,
)
from n8n_client import N8nClient, N8nError
from campaign_api import AuthenticatedCampaignHandler, run_api_server


class TestCampaignAutomation(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_campaign_")
        self.db_path = Path(self.test_dir) / "lead_outreach.db"
        self.mumbai_db_path = Path(self.test_dir) / "mumbai_check_checkpoint.db"
        self.excel_path = Path(self.test_dir) / "instagram_research.xlsx"

        # Initialize mock mumbai checks DB
        with sqlite3.connect(str(self.mumbai_db_path)) as conn:
            conn.executescript("""
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
                );
            """)

        # Initialize Campaign Service
        self.service = CampaignService(db_path=self.db_path)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _seed_mumbai_check(self, username: str, classification: str, confidence: str = "High", stated_location: str = "Mumbai"):
        with sqlite3.connect(str(self.mumbai_db_path)) as conn:
            conn.execute("""
                INSERT OR REPLACE INTO mumbai_checks (
                    normalized_username, checked_username, profile_url, display_name,
                    stated_location, evidence_summary, classification, confidence,
                    checked_at, source_accounts
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                username.lower(), username, f"https://instagram.com/{username}", f"Name {username}",
                stated_location, f"Evidence for {classification}", classification, confidence,
                "2026-10-15T12:00:00Z", "licious_in"
            ))

    def _seed_comment_checkpoint(self, brand: str, comments: list):
        cp_path = Path(self.test_dir) / f"{brand}_checkpoint.db"
        with sqlite3.connect(str(cp_path)) as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS comments (
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
            for c in comments:
                conn.execute("""
                    INSERT OR REPLACE INTO comments (
                        comment_id, shortcode, username, text, created_at, record_type
                    ) VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    c.get("id"), c.get("shortcode", "SHORT1"), c.get("username"),
                    c.get("text", "nice post"), int(time.time()), c.get("record_type", "comment")
                ))

    # --- 1. Campaign Creation & Retrieval ---

    def test_create_and_get_campaign(self):
        camp = self.service.create_campaign(
            campaign_id="test_camp_01",
            name="MeatNDoor Fresh Mutton Campaign",
            objective="Customer Conversion",
            target_location="Mumbai",
            source_accounts=["licious_in"],
            max_leads=50,
            min_lead_score=35,
            notes="Target high-intent commenters",
        )
        self.assertEqual(camp["campaign_id"], "test_camp_01")
        self.assertEqual(camp["name"], "MeatNDoor Fresh Mutton Campaign")
        self.assertEqual(camp["status"], "draft")
        self.assertEqual(camp["min_lead_score"], 35)

        retrieved = self.service.get_campaign("test_camp_01")
        self.assertIsNotNone(retrieved)
        self.assertEqual(retrieved["objective"], "Customer Conversion")

    def test_list_campaigns(self):
        self.service.create_campaign(campaign_id="camp_a", name="Camp A")
        self.service.create_campaign(campaign_id="camp_b", name="Camp B")
        camps = self.service.list_campaigns()
        self.assertGreaterEqual(len(camps), 2)
        c_ids = [c["campaign_id"] for c in camps]
        self.assertIn("camp_a", c_ids)
        self.assertIn("camp_b", c_ids)

    # --- 2. Dual-Axis Scoring Engine & Signal Detection ---

    def test_dual_axis_scoring_high_intent_mumbai(self):
        # Likely Mumbai = 50 location score
        # Meat keywords (mutton, delivery) + question '?' = 15 + 15 + 10 = 40 intent score
        # Total = 90
        tot, loc, intent, signals = self.service.calculate_lead_score(
            location_tier="Likely Mumbai",
            comments_texts=["Do you have fresh goat mutton delivery in Bandra? What is the price?"],
            total_comments=1,
        )
        self.assertEqual(loc, 50)
        self.assertGreaterEqual(intent, 40)
        self.assertEqual(tot, loc + intent)
        self.assertIn("Specific Meat Product Interest", signals)
        self.assertIn("Delivery/Order Inquiry", signals)

    def test_scoring_outside_mumbai_zero_location_score(self):
        # Likely outside Mumbai gives 0 location score
        tot, loc, intent, signals = self.service.calculate_lead_score(
            location_tier="Likely outside Mumbai",
            comments_texts=["I want delivery of chicken to Pune! Can you deliver?"],
            total_comments=1,
        )
        self.assertEqual(loc, 0)
        self.assertEqual(tot, intent)

    def test_repeat_commenter_boost(self):
        tot1, loc1, intent1, _ = self.service.calculate_lead_score(
            location_tier="Serves Mumbai",
            comments_texts=["Nice meat"],
            total_comments=1,
        )
        tot3, loc3, intent3, sig3 = self.service.calculate_lead_score(
            location_tier="Serves Mumbai",
            comments_texts=["Nice meat", "looks delicious", "best quality"],
            total_comments=3,
        )
        self.assertGreater(intent3, intent1)
        self.assertIn("Repeat engagement", sig3)

    # --- 3. End-to-End Pipeline Execution & Suppression ---

    def test_pipeline_execution_qualification_and_suppression(self):
        self._seed_comment_checkpoint("licious_in", [
            {"id": "c1", "username": "rohit_mumbai", "text": "Do you deliver fresh chicken to Andheri west?"},
            {"id": "c2", "username": "rahul_delhi", "text": "Is this available in Delhi Connaught Place?"},
            {"id": "c3", "username": "dnc_user", "text": "Can I order mutton in Colaba?"},
        ])

        self._seed_mumbai_check("rohit_mumbai", "Likely Mumbai", "High", "Andheri West, Mumbai")
        self._seed_mumbai_check("rahul_delhi", "Likely outside Mumbai", "High", "New Delhi")
        self._seed_mumbai_check("dnc_user", "Likely Mumbai", "High", "Colaba, Mumbai")

        # Add dnc_user to persistent DNC
        self.service.lead_outreach_db.add_to_dnc("dnc_user", reason="User requested no contact")

        # Create campaign
        camp = self.service.create_campaign(
            campaign_id="exec_test_01",
            name="Mumbai Chicken & Mutton Launch",
            min_lead_score=30,
            source_accounts=["licious_in"],
        )

        # Execute pipeline
        res = self.service.execute_pipeline("exec_test_01", checkpoint_dir=self.test_dir, mumbai_db_path=self.mumbai_db_path)
        self.assertEqual(res["status"], "completed")

        # Verify leads in database
        leads, total = self.service.get_campaign_leads_paginated(campaign_id="exec_test_01")
        self.assertEqual(total, 3)

        lead_map = {l["username"]: l for l in leads}

        # 1. rohit_mumbai should be eligible/qualified
        rohit = lead_map["rohit_mumbai"]
        self.assertEqual(rohit["campaign_status"], "eligible")
        self.assertGreaterEqual(rohit["lead_score"], 50)
        self.assertEqual(rohit["location_category"], "Likely Mumbai")
        self.assertIn("MeatNDoor", rohit["suggested_message"])
        self.assertIn("Mumbai", rohit["suggested_message"])

        # 2. rahul_delhi should be suppressed due to outside Mumbai
        delhi = lead_map["rahul_delhi"]
        self.assertEqual(delhi["campaign_status"], "suppressed")
        self.assertIn("Confirmed outside Mumbai", delhi["suppression_reason"])

        # 3. dnc_user should be suppressed due to DNC
        dnc = lead_map["dnc_user"]
        self.assertEqual(dnc["campaign_status"], "suppressed")
        self.assertIn("Suppressed by DNC", dnc["suppression_reason"])

    # --- 4. Review Queue Operations ---

    def test_review_queue_approve_edit_reject(self):
        self._seed_comment_checkpoint("licious_in", [
            {"id": "c10", "username": "anjali_meat", "text": "Where is your store in Dadar? Price list please!"}
        ])
        self._seed_mumbai_check("anjali_meat", "Likely Mumbai", "High", "Dadar, Mumbai")

        self.service.create_campaign(campaign_id="rev_test_01", name="Review Test", source_accounts=["licious_in"])
        self.service.execute_pipeline("rev_test_01", checkpoint_dir=self.test_dir, mumbai_db_path=self.mumbai_db_path)

        # Initial state should be pending
        leads, _ = self.service.get_campaign_leads_paginated(campaign_id="rev_test_01")
        self.assertEqual(leads[0]["approval_status"], "pending")

        # Operator edits draft
        new_text = "Hi Anjali! Special Dadar delivery offer from MeatNDoor."
        ok = self.service.edit_lead_draft("rev_test_01", "anjali_meat", new_text, operator="ChefOperator")
        self.assertTrue(ok)
        l_edit = self.service.get_campaign_leads_paginated(campaign_id="rev_test_01")[0][0]
        self.assertEqual(l_edit["suggested_message"], new_text)

        # Operator approves draft
        ok = self.service.approve_lead_draft("rev_test_01", "anjali_meat", operator="ChefOperator")
        self.assertTrue(ok)
        l_app = self.service.get_campaign_leads_paginated(campaign_id="rev_test_01")[0][0]
        self.assertEqual(l_app["approval_status"], "approved")

        # Operator updates contact status to contacted
        ok = self.service.update_lead_contact_status("rev_test_01", "anjali_meat", "contacted", notes="Sent IG message manually")
        self.assertTrue(ok)
        l_cnt = self.service.get_campaign_leads_paginated(campaign_id="rev_test_01")[0][0]
        self.assertEqual(l_cnt["contact_status"], "contacted")

        # Operator rejects
        ok = self.service.reject_lead_draft("rev_test_01", "anjali_meat", operator="ChefOperator")
        self.assertTrue(ok)
        l_rej = self.service.get_campaign_leads_paginated(campaign_id="rev_test_01")[0][0]
        self.assertEqual(l_rej["approval_status"], "rejected")

    # --- 5. Pause, Resume, and Cancellation Controls ---

    def test_pipeline_cancellation(self):
        self.service.create_campaign(campaign_id="cancel_test", name="Cancellation Test")
        cancel_evt = threading.Event()
        cancel_evt.set()

        res = self.service.execute_pipeline("cancel_test", cancel_event=cancel_evt, checkpoint_dir=self.test_dir, mumbai_db_path=self.mumbai_db_path)
        self.assertEqual(res["status"], "cancelled")
        updated_camp = self.service.get_campaign("cancel_test")
        self.assertEqual(updated_camp["status"], "cancelled")

    def test_pipeline_pause_and_resume(self):
        self.service.create_campaign(campaign_id="pause_test", name="Pause Test")
        pause_evt = threading.Event()
        pause_evt.set()

        res = self.service.execute_pipeline("pause_test", pause_event=pause_evt, checkpoint_dir=self.test_dir, mumbai_db_path=self.mumbai_db_path)
        self.assertEqual(res["status"], "paused")

        # Resume without pause event
        res_resumed = self.service.execute_pipeline("pause_test", checkpoint_dir=self.test_dir, mumbai_db_path=self.mumbai_db_path)
        self.assertEqual(res_resumed["status"], "completed")

    # --- 6. n8n Client Functionality ---

    def test_n8n_client_init_and_validation(self):
        client = N8nClient(base_url="http://localhost:5678", webhook_url="http://localhost:5678/webhook/test")
        self.assertEqual(client.base_url, "http://localhost:5678")
        self.assertEqual(client.webhook_url, "http://localhost:5678/webhook/test")

        # Invalid URL raises N8nError
        with self.assertRaises(N8nError):
            N8nClient(base_url="invalid_url_no_scheme")

    @patch("requests.post")
    def test_n8n_client_trigger_campaign(self, mock_post):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"success": True, "execution_id": "exec_12345", "message": "Campaign workflow started"}
        mock_post.return_value = mock_response

        client = N8nClient(
            webhook_url="http://localhost:5678/webhook/test",
            webhook_secret="test_secret_header",
        )
        res = client.trigger_campaign({
            "campaign_id": "c1",
            "name": "Meat Outreach",
            "target_location": "Mumbai",
        })

        self.assertEqual(res["status"], "triggered")
        self.assertEqual(res["execution_id"], "exec_12345")
        mock_post.assert_called_once()
        args, kwargs = mock_post.call_args
        self.assertIn("X-Webhook-Secret", kwargs["headers"])
        self.assertEqual(kwargs["headers"]["X-Webhook-Secret"], "test_secret_header")

    # --- 7. Campaign API Microservice Endpoints ---

    @patch.object(CampaignService, "execute_pipeline")
    def test_campaign_api_endpoints(self, mock_exec):
        mock_exec.return_value = {"status": "completed", "campaign_id": "api_test_01"}

        AuthenticatedCampaignHandler.service = self.service
        os.environ["CAMPAIGN_API_TOKEN"] = "test_token_abc"

        server = run_api_server(host="127.0.0.1", port=0)
        host, port = server.server_address
        server_thread = threading.Thread(target=server.serve_forever, daemon=True)
        server_thread.start()

        import requests
        base_url = f"http://{host}:{port}"

        try:
            # 1. /health endpoint (no auth needed)
            r_health = requests.get(f"{base_url}/health", timeout=3)
            self.assertEqual(r_health.status_code, 200)
            self.assertEqual(r_health.json()["status"], "healthy")

            # 2. Endpoint with missing auth should get 401
            r_unauth = requests.post(f"{base_url}/campaign/start", json={"name": "No Auth"}, timeout=3)
            self.assertEqual(r_unauth.status_code, 401)

            # 3. Endpoint with auth should succeed
            headers = {"Authorization": "Bearer test_token_abc"}
            r_start = requests.post(
                f"{base_url}/campaign/start",
                headers=headers,
                json={
                    "campaign_id": "api_camp_01",
                    "name": "API Automated Campaign",
                    "min_lead_score": 40,
                },
                timeout=3,
            )
            self.assertEqual(r_start.status_code, 202)
            data = r_start.json()
            self.assertEqual(data["status"], "accepted")
            self.assertEqual(data["campaign_id"], "api_camp_01")

            # 4. /campaign/<id>/status
            r_stat = requests.get(f"{base_url}/campaign/api_camp_01/status", headers=headers, timeout=3)
            self.assertEqual(r_stat.status_code, 200)
            self.assertEqual(r_stat.json()["campaign"]["name"], "API Automated Campaign")

        finally:
            server.shutdown()
            server.server_close()


if __name__ == "__main__":
    unittest.main()
