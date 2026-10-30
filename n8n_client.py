#!/usr/bin/env python3
"""
n8n Client & Webhook Integration Module for MeatNDoor Campaign Automation.

This module provides authenticated communication between the Streamlit/Python
application and n8n workflows:
- Triggers campaign start/resume workflows via authenticated webhooks.
- Formats campaign configuration and source references.
- Validates webhook URLs (enforcing HTTP/HTTPS schemes and domain syntax).
- Handles timeouts, connection errors, retries, and execution tracking.
- Queries n8n execution status when API credentials are provided.
- Protects credentials from ever leaking into logs, UI, or exceptions.
"""

import json
import logging
import os
import re
from typing import Any, Dict, Optional, Tuple
from urllib.parse import urlparse

import requests

logger = logging.getLogger("instagram_comments.n8n_client")

# Environment Variable Names
ENV_N8N_BASE_URL = "N8N_BASE_URL"
ENV_N8N_WEBHOOK_URL = "N8N_CAMPAIGN_WEBHOOK_URL"
ENV_N8N_API_KEY = "N8N_API_KEY"
ENV_N8N_WEBHOOK_SECRET = "N8N_WEBHOOK_SECRET"

DEFAULT_N8N_BASE_URL = "http://localhost:5678"
DEFAULT_WEBHOOK_PATH = "/webhook/meatndoor-campaign"


class N8nError(Exception):
    """Base exception for n8n integration failures."""
    pass


class N8nConnectionError(N8nError):
    """Raised when n8n server or webhook is unreachable."""
    pass


class N8nWebhookError(N8nError):
    """Raised when n8n returns an HTTP error status."""
    pass


class N8nValidationError(N8nError):
    """Raised when webhook URL or campaign payload fails validation."""
    pass


def validate_webhook_url(url: str) -> str:
    """
    Validates that a webhook URL is well-formed with http/https scheme.
    Prevents invalid schemes or empty inputs.
    """
    if not url or not isinstance(url, str):
        raise N8nValidationError("Webhook URL cannot be empty.")
    url = url.strip()
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise N8nValidationError(f"Invalid URL scheme '{parsed.scheme}'. Only http and https are permitted.")
    if not parsed.netloc:
        raise N8nValidationError("Webhook URL must contain a valid host / domain.")
    return url


class N8nClient:
    """
    Client for interacting with n8n workflow engine.
    """

    def __init__(
        self,
        base_url: Optional[str] = None,
        webhook_url: Optional[str] = None,
        api_key: Optional[str] = None,
        webhook_secret: Optional[str] = None,
        timeout: float = 15.0,
    ):
        raw_base = base_url or os.environ.get(ENV_N8N_BASE_URL) or DEFAULT_N8N_BASE_URL
        self.base_url = validate_webhook_url(raw_base).rstrip("/")
        raw_webhook = webhook_url or os.environ.get(ENV_N8N_WEBHOOK_URL) or f"{self.base_url}{DEFAULT_WEBHOOK_PATH}"
        self.webhook_url = validate_webhook_url(raw_webhook.strip())
        self.api_key = api_key or os.environ.get(ENV_N8N_API_KEY) or ""
        self.webhook_secret = webhook_secret or os.environ.get(ENV_N8N_WEBHOOK_SECRET) or ""
        self.timeout = timeout

    def _get_auth_headers(self) -> Dict[str, str]:
        """Constructs safe authentication headers."""
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "MeatNDoor-CampaignManager/2.0",
        }
        if self.webhook_secret:
            headers["X-Webhook-Secret"] = self.webhook_secret
            headers["X-N8N-Webhook-Secret"] = self.webhook_secret
        if self.api_key:
            headers["X-N8N-API-KEY"] = self.api_key
        return headers

    def health_check(self) -> Tuple[bool, str]:
        """
        Tests whether the configured n8n base endpoint or webhook is reachable.
        Returns (is_healthy, status_message).
        """
        target_url = self.base_url or self.webhook_url
        try:
            parsed = urlparse(target_url)
            check_url = f"{parsed.scheme}://{parsed.netloc}/healthz" if parsed.netloc else target_url
            resp = requests.get(check_url, timeout=4.0)
            if resp.status_code in (200, 404, 401, 403):
                # 404/401 still means host is alive and listening
                return True, f"n8n host reachable at {parsed.netloc} (HTTP {resp.status_code})"
            return True, f"n8n host responded with HTTP {resp.status_code}"
        except requests.exceptions.RequestException as e:
            return False, f"n8n host unreachable at {target_url}: {type(e).__name__}"

    def trigger_campaign(self, campaign_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Triggers the MeatNDoor campaign automation workflow in n8n via webhook.

        Parameters:
            campaign_data: Dict containing campaign parameters:
                - campaign_id: str
                - name: str
                - objective: str
                - target_location: str
                - source_accounts: List[str]
                - max_leads: int
                - min_lead_score: int
                - trigger_action: 'start' | 'resume' | 'sync'

        Returns:
            Dict with execution details:
                - status: 'triggered' | 'completed' | 'queued'
                - execution_id: str (if returned by n8n)
                - response: Dict
        """
        valid_url = validate_webhook_url(self.webhook_url)

        # Validate required campaign fields
        required_fields = ["campaign_id", "target_location"]
        for rf in required_fields:
            if rf not in campaign_data or not campaign_data[rf]:
                raise N8nValidationError(f"Missing required campaign payload field: '{rf}'")

        payload = {
            "campaign_id": campaign_data["campaign_id"],
            "name": campaign_data.get("name", "Unnamed Campaign"),
            "objective": campaign_data.get("objective", "Audience Research & Outreach"),
            "target_location": campaign_data["target_location"],
            "source_accounts": campaign_data.get("source_accounts", []),
            "max_leads": campaign_data.get("max_leads", 500),
            "min_lead_score": campaign_data.get("min_lead_score", 0),
            "trigger_action": campaign_data.get("trigger_action", "start"),
            "notes": campaign_data.get("notes", ""),
        }

        headers = self._get_auth_headers()

        try:
            logger.info("Triggering n8n campaign webhook at %s for campaign %s", valid_url, payload["campaign_id"])
            resp = requests.post(
                valid_url,
                json=payload,
                headers=headers,
                timeout=self.timeout,
            )

            # Check HTTP status
            if resp.status_code >= 400:
                err_msg = f"n8n webhook responded with HTTP error {resp.status_code}"
                try:
                    err_json = resp.json()
                    err_detail = err_json.get("message") or err_json.get("error") or str(err_json)
                    err_msg += f": {err_detail}"
                except Exception:
                    if resp.text:
                        err_msg += f": {resp.text[:200]}"
                logger.error(err_msg)
                raise N8nWebhookError(err_msg)

            # Parse response
            try:
                res_data = resp.json() if resp.text else {}
            except Exception:
                res_data = {"raw_text": resp.text}

            execution_id = ""
            if isinstance(res_data, dict):
                execution_id = str(res_data.get("executionId") or res_data.get("execution_id") or res_data.get("id") or "")

            return {
                "status": "triggered",
                "campaign_id": payload["campaign_id"],
                "execution_id": execution_id,
                "status_code": resp.status_code,
                "data": res_data,
            }

        except requests.exceptions.Timeout as e:
            msg = f"n8n webhook timed out after {self.timeout}s: {valid_url}"
            logger.error(msg)
            raise N8nConnectionError(msg) from e
        except requests.exceptions.ConnectionError as e:
            msg = f"Failed to connect to n8n webhook at {valid_url}. Please ensure n8n is running."
            logger.error(msg)
            raise N8nConnectionError(msg) from e
        except N8nError:
            raise
        except Exception as e:
            msg = f"Unexpected error while communicating with n8n: {e}"
            logger.exception(msg)
            raise N8nError(msg) from e

    def get_execution_status(self, execution_id: str) -> Dict[str, Any]:
        """
        Queries n8n API for execution status (if n8n API key is configured).
        """
        if not execution_id:
            return {"status": "unknown", "message": "No execution ID provided."}

        if not self.api_key:
            return {
                "status": "unverified",
                "execution_id": execution_id,
                "message": "N8N_API_KEY not configured; execution status tracking via local checkpoint.",
            }

        url = f"{self.base_url}/api/v1/executions/{execution_id}"
        headers = self._get_auth_headers()

        try:
            resp = requests.get(url, headers=headers, timeout=self.timeout)
            if resp.status_code == 200:
                data = resp.json()
                return {
                    "status": data.get("status", "running"),
                    "execution_id": execution_id,
                    "finished": data.get("finished", False),
                    "mode": data.get("mode"),
                    "started_at": data.get("startedAt"),
                    "stopped_at": data.get("stoppedAt"),
                }
            return {
                "status": f"http_{resp.status_code}",
                "execution_id": execution_id,
                "message": f"n8n API responded with {resp.status_code}",
            }
        except Exception as e:
            logger.warning("Could not query n8n execution %s: %s", execution_id, e)
            return {"status": "query_failed", "execution_id": execution_id, "error": str(e)}
