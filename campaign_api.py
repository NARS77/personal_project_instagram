#!/usr/bin/env python3
"""
Minimal Authenticated API Service for n8n Campaign Automation Integration.

Provides an authenticated HTTP boundary for n8n to communicate with the
MeatNDoor Python application:
- Endpoints:
  - GET  /health -> Service liveness & readiness check.
  - POST /campaign/start -> Trigger/execute campaign pipeline.
  - GET  /campaign/<id>/status -> Query campaign execution progress & metrics.
  - POST /campaign/<id>/pause -> Pause running campaign.
  - POST /campaign/<id>/resume -> Resume paused campaign from checkpoint.
  - POST /campaign/<id>/cancel -> Cancel running campaign.
  - POST /campaign/<id>/sync -> Trigger sync to Google Sheets & Excel.
- Authentication: Enforces 'X-API-Key' or 'Authorization: Bearer <token>'.
- Uses standard library ThreadingHTTPServer (zero external web framework dependencies).
"""

import json
import logging
import os
import re
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, Optional
from urllib.parse import parse_qs, urlparse

from campaign_service import CampaignError, CampaignService

logger = logging.getLogger("instagram_comments.campaign_api")

DEFAULT_API_PORT = int(os.environ.get("CAMPAIGN_API_PORT", 8000))
DEFAULT_API_HOST = os.environ.get("CAMPAIGN_API_HOST", "127.0.0.1")
ENV_API_TOKEN = "CAMPAIGN_API_TOKEN"


class AuthenticatedCampaignHandler(BaseHTTPRequestHandler):
    """HTTP Request handler with API token authentication."""

    service = CampaignService()

    def _send_json(self, status_code: int, data: Dict[str, Any]) -> None:
        payload = json.dumps(data).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _is_authenticated(self) -> bool:
        expected_token = os.environ.get(ENV_API_TOKEN) or os.environ.get("N8N_WEBHOOK_SECRET")
        if not expected_token:
            # If no secret is configured, allow local loopback requests
            client_ip = self.client_address[0]
            return client_ip in ("127.0.0.1", "::1", "localhost")

        # Check Authorization header or X-API-Key
        auth_header = self.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            token = auth_header[7:].strip()
            if token == expected_token:
                return True

        api_key_header = self.headers.get("X-API-Key") or self.headers.get("X-N8N-Webhook-Secret")
        if api_key_header and api_key_header.strip() == expected_token:
            return True

        return False

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        # Health endpoint requires no auth
        if path == "/health":
            self._send_json(200, {
                "status": "healthy",
                "service": "MeatNDoor Campaign API",
                "version": "2.0",
            })
            return

        if not self._is_authenticated():
            self._send_json(401, {"error": "Unauthorized: Invalid or missing API token."})
            return

        # GET /campaign/<id>/status
        status_match = re.match(r"^/campaign/([^/]+)/status$", path)
        if status_match:
            camp_id = status_match.group(1)
            camp = self.service.get_campaign(camp_id)
            if not camp:
                self._send_json(404, {"error": f"Campaign '{camp_id}' not found."})
                return
            self._send_json(200, {"status": "ok", "campaign": camp})
            return

        self._send_json(404, {"error": f"Endpoint not found: {path}"})

    def do_POST(self) -> None:
        if not self._is_authenticated():
            self._send_json(401, {"error": "Unauthorized: Invalid or missing API token."})
            return

        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        # Read JSON body
        content_len = int(self.headers.get("Content-Length", 0))
        body_data: Dict[str, Any] = {}
        if content_len > 0:
            try:
                body_bytes = self.rfile.read(content_len)
                body_data = json.loads(body_bytes.decode("utf-8"))
            except Exception as e:
                self._send_json(400, {"error": f"Invalid JSON body: {e}"})
                return

        # POST /campaign/start
        if path == "/campaign/start":
            camp_id = body_data.get("campaign_id")
            if not camp_id:
                self._send_json(400, {"error": "Missing 'campaign_id' in request body."})
                return

            existing = self.service.get_campaign(camp_id)
            if not existing:
                try:
                    self.service.create_campaign(
                        campaign_id=camp_id,
                        name=body_data.get("name", camp_id),
                        objective=body_data.get("objective", "Audience Research"),
                        target_location=body_data.get("target_location", "Mumbai"),
                        source_accounts=body_data.get("source_accounts", []),
                        max_leads=body_data.get("max_leads", 500),
                        min_lead_score=body_data.get("min_lead_score", 0),
                        notes=body_data.get("notes", "Triggered via API"),
                    )
                except Exception as e:
                    self._send_json(400, {"error": str(e)})
                    return

            # Execute pipeline synchronously or in background thread
            async_mode = body_data.get("async", True)
            if async_mode:
                t = threading.Thread(
                    target=self.service.execute_pipeline,
                    kwargs={"campaign_id": camp_id, "trigger_source": "api_webhook"},
                    daemon=True,
                )
                t.start()
                self._send_json(202, {
                    "status": "accepted",
                    "campaign_id": camp_id,
                    "message": f"Campaign '{camp_id}' execution started in background.",
                })
            else:
                try:
                    res = self.service.execute_pipeline(campaign_id=camp_id, trigger_source="api_sync")
                    self._send_json(200, res)
                except Exception as e:
                    self._send_json(500, {"error": str(e)})
            return

        # POST /campaign/<id>/pause
        pause_match = re.match(r"^/campaign/([^/]+)/pause$", path)
        if pause_match:
            camp_id = pause_match.group(1)
            try:
                self.service.update_campaign_status(camp_id, "paused", current_stage="Paused via API")
                self._send_json(200, {"status": "paused", "campaign_id": camp_id})
            except Exception as e:
                self._send_json(500, {"error": str(e)})
            return

        # POST /campaign/<id>/cancel
        cancel_match = re.match(r"^/campaign/([^/]+)/cancel$", path)
        if cancel_match:
            camp_id = cancel_match.group(1)
            try:
                self.service.update_campaign_status(camp_id, "cancelled", current_stage="Cancelled via API")
                self._send_json(200, {"status": "cancelled", "campaign_id": camp_id})
            except Exception as e:
                self._send_json(500, {"error": str(e)})
            return

        # POST /campaign/<id>/sync
        sync_match = re.match(r"^/campaign/([^/]+)/sync$", path)
        if sync_match:
            camp_id = sync_match.group(1)
            try:
                ex_cnt = self.service.sync_campaign_to_excel(camp_id)
                sh_res = self.service.sync_campaign_to_google_sheets(camp_id)
                self._send_json(200, {
                    "status": "synced",
                    "campaign_id": camp_id,
                    "excel_rows": ex_cnt,
                    "sheets_url": sh_res.get("spreadsheet_url", ""),
                })
            except Exception as e:
                self._send_json(500, {"error": str(e)})
            return

        self._send_json(404, {"error": f"Endpoint not found: {path}"})

    def log_message(self, format: str, *args: Any) -> None:
        """Suppresses raw console clutter while logging securely."""
        logger.debug("%s - - [%s] %s", self.client_address[0], self.log_date_time_string(), format % args)


def run_api_server(host: str = DEFAULT_API_HOST, port: int = DEFAULT_API_PORT) -> ThreadingHTTPServer:
    """Initializes and returns running ThreadingHTTPServer instance."""
    server = ThreadingHTTPServer((host, port), AuthenticatedCampaignHandler)
    logger.info("Started Campaign API Server on http://%s:%d", host, port)
    return server


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="MeatNDoor Campaign Automation API Server")
    parser.add_argument("--host", default=DEFAULT_API_HOST, help="Host to bind to")
    parser.add_argument("--port", type=int, default=DEFAULT_API_PORT, help="Port to listen on")
    args = parser.parse_args()

    srv = run_api_server(host=args.host, port=args.port)
    print(f"[✓] Campaign API running on http://{args.host}:{args.port}")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Campaign API...")
        srv.server_close()
