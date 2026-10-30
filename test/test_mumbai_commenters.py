#!/usr/bin/env python3
"""
Unit and integration tests for Mumbai Location Checker Commenter Processing.

Verifies:
1. Target list contains commenters and repliers, NOT the source account itself.
2. Reply authors are included alongside top-level commenters.
3. Multiple comments across posts by the same user are deduplicated to 1 check.
4. Accounts already marked completed in SQLite are skipped by default.
5. The recheck flag forces rechecking of completed accounts.
6. A failed check / access error does not overwrite an existing completed result.
"""

import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from typing import Any, Dict

from mumbai_checker import (
    MumbaiCheckDB,
    MumbaiLocationChecker,
    load_commenters_from_checkpoints,
)


class TestMumbaiCommenters(unittest.TestCase):
    def setUp(self):
        # Create a temporary directory for isolated test databases
        self.test_dir = tempfile.mkdtemp(prefix="mumbai_test_")
        self.checkpoints_dir = Path(self.test_dir) / ".checkpoints"
        self.checkpoints_dir.mkdir(parents=True, exist_ok=True)
        self.mumbai_db_path = self.checkpoints_dir / "mumbai_check_checkpoint.db"
        self.db = MumbaiCheckDB(db_path=self.mumbai_db_path)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_mock_checkpoint(self, source_name: str, records: list) -> Path:
        """Helper to create a mock brand checkpoint database with comments & replies."""
        db_path = self.checkpoints_dir / f"{source_name}_checkpoint.db"
        conn = sqlite3.connect(str(db_path))
        conn.execute("""
            CREATE TABLE IF NOT EXISTS comments (
                id TEXT PRIMARY KEY,
                post_id TEXT,
                username TEXT,
                record_type TEXT,
                text TEXT,
                timestamp TEXT
            )
        """)
        for r in records:
            conn.execute(
                "INSERT INTO comments (id, post_id, username, record_type, text, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
                (
                    r.get("id"),
                    r.get("post_id", "p1"),
                    r.get("username"),
                    r.get("record_type", "comment"),
                    r.get("text", "mock text"),
                    r.get("timestamp", "2026-10-12T10:00:00Z"),
                ),
            )
        conn.commit()
        conn.close()
        return db_path

    def test_1_source_account_excluded_commenters_included(self):
        """
        Verify: When comments from a source account (e.g. 'licious_foods') are processed,
        the target list contains the commenters and repliers, NOT 'licious_foods' itself.
        """
        self._create_mock_checkpoint(
            source_name="licious_foods",
            records=[
                {"id": "c1", "username": "rahul_mumbai", "record_type": "comment"},
                {"id": "c2", "username": "licious_foods", "record_type": "comment"},  # Brand self-comment
                {"id": "c3", "username": "anita_sharma", "record_type": "comment"},
                {"id": "r1", "username": "licious_foods", "record_type": "reply"},    # Brand official reply
                {"id": "r2", "username": "vikram_foodie", "record_type": "reply"},
            ],
        )

        commenters = load_commenters_from_checkpoints(
            source_usernames=["licious_foods"],
            exclude_source_brand=True,
            checkpoints_dir=self.checkpoints_dir,
        )

        # Assert brand profile is excluded
        self.assertNotIn("licious_foods", commenters)

        # Assert commenters are included
        self.assertIn("rahul_mumbai", commenters)
        self.assertIn("anita_sharma", commenters)
        self.assertIn("vikram_foodie", commenters)
        self.assertEqual(len(commenters), 3)

    def test_2_reply_authors_included_alongside_top_level_commenters(self):
        """
        Verify: Reply authors ('reply') are included alongside top-level commenters ('comment').
        """
        self._create_mock_checkpoint(
            source_name="meatigo_official",
            records=[
                {"id": "c10", "username": "top_commenter_1", "record_type": "comment"},
                {"id": "r20", "username": "reply_author_1", "record_type": "reply"},
                {"id": "r21", "username": "reply_author_2", "record_type": "reply"},
            ],
        )

        commenters = load_commenters_from_checkpoints(
            source_usernames=["meatigo_official"],
            exclude_source_brand=True,
            checkpoints_dir=self.checkpoints_dir,
        )

        self.assertIn("top_commenter_1", commenters)
        self.assertIn("reply_author_1", commenters)
        self.assertIn("reply_author_2", commenters)

        self.assertEqual(commenters["top_commenter_1"]["comment_count"], 1)
        self.assertEqual(commenters["top_commenter_1"]["reply_count"], 0)
        self.assertEqual(commenters["reply_author_1"]["comment_count"], 0)
        self.assertEqual(commenters["reply_author_1"]["reply_count"], 1)

    def test_3_deduplication_across_multiple_posts_and_cases(self):
        """
        Verify: A user who commented multiple times across different posts and replies,
        with mixed case and '@' prefixes, is checked only once.
        """
        self._create_mock_checkpoint(
            source_name="prasuma_official",
            records=[
                {"id": "c1", "post_id": "post_A", "username": "FoodLover_99", "record_type": "comment"},
                {"id": "c2", "post_id": "post_B", "username": "@foodlover_99", "record_type": "comment"},
                {"id": "r1", "post_id": "post_C", "username": " FOODLOVER_99 ", "record_type": "reply"},
                {"id": "c3", "post_id": "post_D", "username": "another_user", "record_type": "comment"},
            ],
        )

        commenters = load_commenters_from_checkpoints(
            source_usernames=["prasuma_official"],
            exclude_source_brand=True,
            checkpoints_dir=self.checkpoints_dir,
        )

        # Must have exactly 2 unique users
        self.assertEqual(len(commenters), 2)
        self.assertIn("foodlover_99", commenters)
        self.assertIn("another_user", commenters)

        # Interaction totals must be aggregated properly
        user_data = commenters["foodlover_99"]
        self.assertEqual(user_data["total_interactions"], 3)
        self.assertEqual(user_data["comment_count"], 2)
        self.assertEqual(user_data["reply_count"], 1)
        self.assertEqual(user_data["checked_username"], "FoodLover_99")  # Preserves display casing

    def test_4_skip_already_completed_accounts_by_default(self):
        """
        Verify: An account already marked completed in SQLite is skipped on a subsequent run.
        """
        # Save a completed check in database
        self.db.save_check({
            "checked_username": "mumbai_foodie_chef",
            "display_name": "Chef Rohan",
            "profile_url": "https://www.instagram.com/mumbai_foodie_chef/",
            "stated_location": "Bandra, Mumbai",
            "evidence_summary": "Bio states Bandra, Mumbai.",
            "classification": "Likely Mumbai",
            "confidence": "High",
            "has_error": 0,
        })

        # Should skip by default when recheck is False
        self.assertTrue(self.db.should_skip("mumbai_foodie_chef", recheck=False))
        self.assertTrue(self.db.should_skip("@MUMBAI_FOODIE_CHEF", recheck=False))

        # Unchecked user should not be skipped
        self.assertFalse(self.db.should_skip("new_unknown_user", recheck=False))

    def test_5_recheck_flag_forces_rechecking_completed_accounts(self):
        """
        Verify: The --recheck flag (recheck=True) forces rechecking of completed accounts.
        """
        self.db.save_check({
            "checked_username": "existing_account",
            "display_name": "Existing Account",
            "classification": "Serves Mumbai",
            "confidence": "Medium",
            "has_error": 0,
        })

        # When recheck=False -> skips
        self.assertTrue(self.db.should_skip("existing_account", recheck=False))

        # When recheck=True -> does NOT skip
        self.assertFalse(self.db.should_skip("existing_account", recheck=True))

    def test_6_subsequent_error_does_not_overwrite_completed_result(self):
        """
        Verify: A subsequent access error or failure does not overwrite a good completed result.
        """
        # Initial successful check
        initial_check = {
            "checked_username": "stable_mumbai_user",
            "display_name": "Stable User",
            "profile_url": "https://www.instagram.com/stable_mumbai_user/",
            "stated_location": "Andheri West, Mumbai",
            "evidence_summary": "Explicit address in Andheri West, Mumbai verified.",
            "classification": "Likely Mumbai",
            "confidence": "High",
            "has_error": 0,
            "error_message": "",
        }
        self.db.save_check(initial_check)

        saved_1 = self.db.get_check("stable_mumbai_user")
        self.assertEqual(saved_1["classification"], "Likely Mumbai")
        self.assertEqual(saved_1["has_error"], 0)

        # Later transient network/rate limit error
        error_check = {
            "checked_username": "stable_mumbai_user",
            "display_name": "",
            "profile_url": "https://www.instagram.com/stable_mumbai_user/",
            "stated_location": "—",
            "evidence_summary": "Profile request timed out / rate limited.",
            "classification": "Not enough evidence",
            "confidence": "Low",
            "has_error": 1,
            "error_message": "ETIMEDOUT: Connection timed out",
        }
        self.db.save_check(error_check)

        # Verify prior clean result is preserved intact
        preserved = self.db.get_check("stable_mumbai_user")
        self.assertEqual(preserved["classification"], "Likely Mumbai")
        self.assertEqual(preserved["has_error"], 0)
        self.assertEqual(preserved["stated_location"], "Andheri West, Mumbai")
        self.assertIn("Explicit address in Andheri West", preserved["evidence_summary"])

    def test_7_batch_limit_skipping_and_telemetry(self):
        """
        Verify: check_commenter_batch honors batch limit, skips already checked,
        and accurately returns discovered, deduplicated, skipped, and newly checked counts.
        """
        checker = MumbaiLocationChecker(db=self.db)

        # Pre-seed 1 user as already checked
        self.db.save_check({
            "checked_username": "user_already_done",
            "display_name": "Done User",
            "classification": "Serves Mumbai",
            "confidence": "Medium",
            "has_error": 0,
        })

        # 4 unique commenters with varying interactions
        mock_commenters = {
            "user_already_done": {
                "checked_username": "user_already_done",
                "normalized_username": "user_already_done",
                "source_accounts": ["brand_x"],
                "total_interactions": 3,
            },
            "user_new_1": {
                "checked_username": "user_new_1",
                "normalized_username": "user_new_1",
                "source_accounts": ["brand_x"],
                "total_interactions": 2,
            },
            "user_new_2": {
                "checked_username": "user_new_2",
                "normalized_username": "user_new_2",
                "source_accounts": ["brand_x"],
                "total_interactions": 1,
            },
            "user_new_3": {
                "checked_username": "user_new_3",
                "normalized_username": "user_new_3",
                "source_accounts": ["brand_x"],
                "total_interactions": 1,
            },
        }

        # Mock fetch_public_evidence to avoid network calls
        original_fetch = checker.fetch_public_evidence
        checker.fetch_public_evidence = lambda u: {
            "username": u,
            "profile_url": f"https://www.instagram.com/{u}/",
            "full_name": f"Mock {u}",
            "biography": "Based in Mumbai" if u == "user_new_1" else "Foodie in Delhi",
            "external_url": "",
            "business_address": {},
            "location_tags": ["Mumbai"] if u == "user_new_1" else ["Delhi"],
            "posts_checked": 1,
            "has_error": False,
        }

        try:
            # Batch with limit=2 (should skip user_already_done, and process user_new_1, user_new_2)
            progress_events = []
            res = checker.check_commenter_batch(
                commenters_dict=mock_commenters,
                recheck=False,
                limit=2,
                delay=0.0,
                progress_callback=lambda evt: progress_events.append(evt),
            )

            self.assertEqual(res["discovered_count"], 7)  # 3 + 2 + 1 + 1
            self.assertEqual(res["deduplicated_count"], 4)
            self.assertEqual(res["skipped_count"], 1)  # user_already_done skipped
            self.assertEqual(res["newly_checked_count"], 2)  # limited to 2
            self.assertEqual(len(res["results"]), 2)
            self.assertTrue(len(progress_events) >= 2)
        finally:
            checker.fetch_public_evidence = original_fetch

    def test_8_classify_outside_location_tags(self):
        """
        Verify: Accounts with multiple non-Mumbai location tags are classified
        as 'Likely outside Mumbai' without subscripting errors.
        """
        checker = MumbaiLocationChecker(db=self.db)
        evidence = {
            "username": "bengaluru_resident",
            "profile_url": "https://www.instagram.com/bengaluru_resident/",
            "full_name": "Bangalore Foodie",
            "biography": "Love exploring food",
            "external_url": "",
            "business_address": {},
            "location_tags": ["Bengaluru", "Indiranagar, Bangalore"],
            "posts_checked": 2,
            "has_error": False,
        }
        res = checker.classify(evidence, source_accounts="licious_foods")
        self.assertEqual(res["classification"], "Likely outside Mumbai")
        self.assertIn("Bengaluru", res["stated_location"])


if __name__ == "__main__":
    unittest.main()

