# Features Implemented: Instagram Intelligence & Location Checker Suite

> **Project Root**: `/home/nars/projects/api-anything-api`  
> **Core Package**: `api-anything/`  
> **Target Google Spreadsheet ID**: [`1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y`](https://docs.google.com/spreadsheets/d/1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y/edit)  
> **Target Local Excel Mirror**: `exports/instagram_research.xlsx`  
> **Last Updated**: October 14, 2026  

---

## Document Purpose & Maintenance Rule

This document maintains the canonical, up-to-date registry of all implemented features across the platform. 

> [!IMPORTANT]
> **Incremental Maintenance Rule**:  
> Every time a new feature, engine, or pipeline is developed and verified in this codebase, a corresponding section and changelog entry **must be appended to this document**.

---

## 1. System Overview & Verified Baseline

The platform is an enterprise-grade Instagram public intelligence, engagement extraction, and demographic auditing suite. It allows researchers and operators to analyze audience behavior, discover high-intent commenters, determine geographic demographics (specifically Mumbai residence/business activity), and sync data dual-stream to Google Sheets and Microsoft Excel without generating loose CSV files.

### Verified Live Scale & Database Metrics
- **Mumbai Location Database (`api-anything/.checkpoints/mumbai_check_checkpoint.db`)**:
  - **5,857 unique accounts** audited, classified, and persistently cached.
  - Distribution breakdown:
    - `Likely Mumbai`: **90 accounts** (High/Medium Confidence)
    - `Serves Mumbai`: **20 accounts** (High/Medium Confidence)
    - `Possibly Mumbai`: **249 accounts** (Low Confidence)
    - `Likely outside Mumbai`: **407 accounts** (High/Medium Confidence)
    - `Not enough evidence`: **5,091 accounts** (Low Confidence)
- **Comment Checkpoint Databases (`api-anything/.checkpoints/{username}_checkpoint.db`)**:
  - Contains full post shortcodes, timestamps, parent comments, and nested replies across target competitor profiles (`licious_foods`, `meatigo_official`, `prasuma_official`, `cambaytiger`, `fishtokri`, `bastianmumbai`).
- **Google Sheets API v4 Integration**:
  - Target sheet `1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y` configured and verified using Service Account credentials (`service_account.json`).
- **Local Excel Mirror**:
  - `api-anything/exports/instagram_research.xlsx` maintained with matching worksheets and formatting via `openpyxl`.
- **Automated Unit Testing**:
  - `api-anything/test/test_mumbai_commenters.py`: **8/8 unit tests passing** (0.013s).

---

## 2. Implemented Features Directory

### Feature 1: Public Timeline & Post Extraction
- **File**: `instagram_comments.py`
- **Capabilities**:
  - Traverses public timeline posts for any target Instagram handle.
  - Scalable post volume: configurable via CLI or GUI slider from **1 to 500 posts**.
  - Collects post metadata: post shortcode, canonical URL (`https://www.instagram.com/p/{shortcode}/`), timestamp, caption text, and comment count.
  - Public web endpoint discovery with pagination cursor handling.

---

### Feature 2: High-Fidelity Comment & Reply Extraction
- **File**: `instagram_comments.py`
- **Capabilities**:
  - **Top-Level Comments**: Ingests primary comments with author `@username`, `display_name`, comment text, like count, comment ID, and creation date.
  - **Nested Child Replies**: Checks `child_comment_count > 0` and queries GraphQL reply endpoints (`graphql/query/?query_hash=...`) to retrieve all threaded replies under each comment.
  - Full author attribution distinguishing between parent commenters (`record_type='comment'`) and reply respondents (`record_type='reply'`, with `parent_comment_id` link).
  - Exponential backoff and jittered delays (1.5s to 3.5s) to defend against Instagram rate-limiting.

---

### Feature 3: Resumable Fault-Tolerant SQLite Checkpointing
- **Files**: `instagram_comments.py`, `mumbai_checker.py`
- **Capabilities**:
  - Transactional checkpointing: writes records atomically into `.checkpoints/{username}_checkpoint.db` using SQLite WAL mode (`BEGIN IMMEDIATE`).
  - Zero data loss: network interrupts, power outages, or manual cancellations safely preserve all extracted rows.
  - Resumption: restarting an extraction job checks existing shortcodes and comment IDs, skipping already-harvested posts.

---

### Feature 4: Mumbai Location Intelligence & Demographic Classification Engine
- **File**: `mumbai_checker.py`
- **Capabilities**:
  - **Multi-Brand Commenter Ingestion**: Harvests commenter and reply-author usernames across any selection of past checkpoint databases.
  - **Case-Insensitive Normalization**: Strips `@` prefixes and applies canonical lowercase keys (`normalized_username`).
  - **Public Profile Evidence Gathering**: Scrapes public bio, display name, external website URLs, and recent post location tags.
  - **5-Tier Geographic Heuristic Classifier**:
    1. `Likely Mumbai` (High/Medium Confidence): Direct bio references (e.g., "Mumbai", "Bombay"), iconic landmarks (Marine Drive, Bandra, Powai, Juhu, Colaba, BKC, CST), or Mumbai pincodes (`400001`–`400104`).
    2. `Serves Mumbai` (High/Medium Confidence): Commercial accounts, cloud kitchens, and service providers explicitly operating in Mumbai.
    3. `Possibly Mumbai` (Low Confidence): Contextual mentions without definitive residence proof.
    4. `Likely outside Mumbai` (High/Medium Confidence): Explicit negative markers for other cities/regions (Delhi, Bangalore, Bengaluru, Pune, Hyderabad, Chennai, Kolkata, Gurgaon, Noida).
    5. `Not enough evidence` (Low Confidence): Blank or private profiles with no public location evidence.
  - **Error Shielding**: If an account previously received a valid classification, a subsequent transient network error or rate limit will **never overwrite** the valid record with an error state.
  - **Persistent Cache**: Recheck flags prevent redundant API calls by skipping accounts already evaluated in `mumbai_checks`.

---

### Feature 5: Dual-Destination Real-Time Reporting (Google Sheets & Excel)
- **File**: `sheets_sync.py`
- **Capabilities**:
  - **Google Sheets API v4**:
    - Authenticates via `service_account.json` or interactive OAuth (`credentials.json` / `token.json`).
    - Syncs brand-specific worksheets into spreadsheet `1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y`.
    - Formats worksheets with dark-gray bold header banners, frozen top rows, auto-sized columns, and clickable `=HYPERLINK()` formulas for post URLs and profiles.
  - **Local Excel Mirror**:
    - Generates matching `.xlsx` sheets in `exports/instagram_research.xlsx` using `openpyxl`.
    - Retains clickable links and professional formatting offline.
  - **Strict Anti-CSV Rule**: Eliminates messy CSV file management in favor of structured databases, live Google Sheets, and native Excel workbooks.

---

### Feature 6: High-Performance Graphite Web GUI
- **Files**: `gui_app.py`, `gui_service.py`
- **Capabilities**:
  - **Dynamic Isolated View Router**:
    - Uses `st.radio` navigation to mount only the active view. Idle views consume **0% CPU and zero I/O**, resolving frontend performance bottlenecks.
  - **WebSocket Progress Fragments (`@st.fragment`)**:
    - Background job gauges, counter pills, and live log consoles update every 2 seconds via WebSocket without rerunning the main page.
    - User inputs, search fields, and pagination selectors stay fully responsive during active runs.
  - **Server-Side SQL Pagination**:
    - Queries large tables (`mumbai_checks` with 5,857+ rows) using `LIMIT` and `OFFSET` directly in SQLite.
    - Prevents browser lag by rendering only 25, 50, or 100 rows at a time in virtualized canvas grids.
  - **Four Primary Operational Views**:
    1. **💬 Instagram Comments Extractor**: Target handle input, 1–500 post slider, child reply toggle, sync switches, and live progress console.
    2. **📍 Mumbai Location Checker**: Multi-source selector (checkpoints, Excel file, or custom list), target pre-check summary, and filterable audit grid.
    3. **🗂️ Checkpoint History & Database Inspector**: Visual cards of past scrapes, total records, disk space, and raw record inspection.
    4. **⚙️ System Settings & Health Check**: Credential validator, Google Sheets write-access verifier, and path diagnostics.

---

### Feature 7: Excel Username Ingestion & Fast Pre-Check Summaries
- **Files**: `gui_service.py`, `gui_app.py`
- **Capabilities**:
  - Ingests external target accounts directly from uploaded `.xlsx` workbooks or local file paths using `openpyxl`.
  - Sheet selector dropdown dynamically lists available sheets within the uploaded workbook.
  - Target Pre-Check Summary instantly calculates `Total Commenters`, `Already Audited`, and `Pending Audits` using pre-aggregated metadata without running expensive table scans during page loads.

---

### Feature 8: Reliability, Security & Testing Suite
- **Files**: `gui_service.py`, `api-anything/test/test_mumbai_commenters.py`
- **Capabilities**:
  - **In-Memory Redacting Log Handler**: Circular buffer capturing real-time worker logs while automatically masking API keys, OAuth tokens, and sensitive strings.
  - **Thread-Safe Job Managers**: `ExtractionJobManager` and `MumbaiJobManager` manage background worker threads with start/stop triggers and status polling.
  - **Automated Regression Suite**: Tests checkpoint loading, case-insensitive deduplication, error shielding, classification rules, and SQLite persistence.

---

### Feature 9: MeatNDoor Lead Outreach Manager
- **Files**: `lead_outreach.py`, `gui_service.py`, `gui_app.py`, `sheets_sync.py`, `test/test_lead_outreach.py`
- **Capabilities**:
  - **Incremental & Idempotent Lead Repository**:
    - Ingests commenters and nested reply authors across existing brand checkpoints (`.checkpoints/*_checkpoint.db`).
    - Discards source brand profiles (e.g., `licious_foods` is never ingested as a lead).
    - Merges new source accounts and posts without overwriting operator edits, custom notes, or review statuses.
    - Persists records in SQLite at `.checkpoints/lead_outreach.db` (`outreach_leads`, `do_not_contact`, `outreach_audit_log`).
  - **Location-Based Campaign Eligibility Engine**:
    - `Likely Mumbai` -> `Eligible (Mumbai Target)`
    - `Serves Mumbai` -> `Review (Business/Serves Mumbai)`
    - `Possibly Mumbai` -> `Review (Low Confidence)`
    - `Likely outside Mumbai` -> `Suppressed (Outside Mumbai)` (auto-suppressed from Mumbai campaigns)
    - `Not enough evidence` -> `Pending Location Verification`
    - `Do Not Contact` -> `Suppressed (Do Not Contact)` (auto-suppressed across all campaigns)
    - Error Shielding: Transient lookup errors are never converted into confirmed outside-Mumbai classifications.
  - **Editable Message Drafting Engine**:
    - Standard Mumbai-specific template for eligible Mumbai prospects.
    - Standard Generic template for food lovers.
    - Operator inline draft editor with change persistence and audit logging.
    - Anti-Circumvention Rule: Generic messages cannot be used to bypass Mumbai-only exclusion rules.
  - **Interactive Streamlit Review Queue (`🍗 MeatNDoor Lead Outreach`)**:
    - Fifth isolated view in Streamlit GUI with zero idle CPU overhead.
    - KPI Metrics Bar: Total Leads, Mumbai Eligible, Outside Suppressed, DNC Suppressed, Pending Review, Approved, Contacted.
    - Full search, eligibility filters, review status filters, and server-side SQL pagination (`LIMIT / OFFSET`).
    - Actionable cards with 1-click status transitions: `✅ Approve Draft` (approves without sending), `💾 Save Draft`, `❌ Reject`, `📩 Mark Contacted`, `💬 Mark Replied`, `🌟 Mark Interested`, `🚫 Mark Not Interested`, `⛔ Add to DNC / Un-DNC`.
    - Direct compliant chat link (`https://ig.me/m/{username}`) and profile link.
    - Real-time operator audit log tracking all decisions and transitions.
  - **Official Meta Authorized Messaging Integration**:
    - `MetaMessagingClient` strictly adheres to Meta Messenger API for Instagram policies.
    - Cold commenters lacking Meta IGSIDs or outside the 24-hour messaging window are routed to `Manual / Opt-In Required`.
    - Automated bulk follower requests, bot DMs, and private API scraping are strictly prohibited.
    - Duplicate-send protection blocks re-contacting previously reached leads.
  - **Dual Reporting Synchronization**:
    - Synchronizes the review queue to `Outreach Manager` worksheet in Google Sheets (`1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y`) and local Excel (`exports/instagram_research.xlsx`).
    - Retains all existing brand tabs and Mumbai Check tabs without data loss. Zero CSV exports.

- **10. n8n Campaign Automation & Lead Intelligence Pipeline** (`n8n_client.py`, `campaign_service.py`, `campaign_api.py`, `gui_app.py`, `gui_service.py`, `n8n_workflows/meatndoor_campaign.json`):
  - **10-Stage Resumable Pipeline**:
    - Stage 1: Campaign Configuration & Parameter Validation
    - Stage 2: Multi-Checkpoint Commenter & Reply Author Discovery
    - Stage 3: Normalization & Cross-Brand Deduplication
    - Stage 4: Mumbai Location Qualification (integrating 5,857 cached evaluations)
    - Stage 5: Dual-Axis Lead Scoring (Location Score 0-50 + Customer Intent Score 0-50)
    - Stage 6: Personalized Message Drafting (Mumbai vs Generic templates)
    - Stage 7: Do-Not-Contact (DNC) Registry & Outside-Mumbai Exclusion
    - Stage 8: SQLite Database Persistence (`campaigns`, `campaign_leads`, `campaign_executions`)
    - Stage 9: Dual Reporting Synchronization (Google Sheets `Campaign Leads` tab + Excel mirror)
    - Stage 10: Execution Metrics Reporting & Telemetry Finalization
  - **Dual-Axis Intent & Location Lead Scoring (0–100 Scale)**:
    - **Location Score (0–50 pts)**: `Likely Mumbai` (50), `Serves Mumbai` (35), `Possibly Mumbai` (25), `Not enough evidence` (10), `Likely outside Mumbai` (0 - suppressed).
    - **Customer Buying Intent Score (0–50 pts)**: Specific meat product keywords (`mutton`, `chicken`, `prawns`, `fresh`, etc. +15), delivery/order terms (`delivery`, `order`, `cod`, `doorstep` +15), price/availability terms (`price`, `rate`, `available` +10), direct question presence (`?` +10), repeat engagement frequency (+5 to +10). Capped at 50 pts.
  - **Interactive Streamlit Campaign Manager (`🎯 Campaign Manager`)**:
    - First view in Streamlit GUI navigation with real-time WebSocket progress polling via `@st.fragment`.
    - Campaign Selector & Creator Form with customizable name, objective, target location, source account filters, max leads slider (10 to 1,000), minimum score filter (0 to 100), and optional n8n webhook URL.
    - Six Primary Controls: `▶️ Start Campaign`, `⏸️ Pause`, `⏯️ Resume`, `⏹️ Cancel`, `🔄 Refresh Status`, and `📊 Sync Sheets`. Duplicate execution guards prevent accidental parallel starts.
    - Stage-Level Telemetry Panel: Live progress bar, leads evaluated, qualified count, high-intent count, drafts generated, suppressed count, elapsed timer, n8n connection indicator, and error banners.
    - Actionable Review Queue: Search bar, approval status filter, contact status filter, sorting (Total Score, Intent Score, Username, Recency), SQL pagination (`LIMIT / OFFSET`), inline draft editing with save, 1-click approvals (approves draft without sending), status updates, compliant `https://ig.me/m/{username}` chat link, and persistent DNC.
  - **n8n Orchestrator Integration & Native Python Fallback**:
    - `N8nClient`: Secure client supporting URL validation, payload construction, webhook authentication (`X-Webhook-Secret`, `X-N8N-API-KEY`), execution status polling, and health checks.
    - Native Fallback: If n8n is offline or unconfigured, the background engine runs the full pipeline directly while providing complete telemetry and audit logging.
  - **Lightweight Microservice Boundary (`campaign_api.py`)**:
    - Zero-dependency authenticated HTTP server using Python standard library `http.server.ThreadingHTTPServer`.
    - Endpoints: `GET /health`, `POST /campaign/start`, `GET /campaign/<id>/status`, `POST /campaign/<id>/pause`, `POST /campaign/<id>/resume`, `POST /campaign/<id>/cancel`, and `POST /campaign/<id>/sync`.
  - **Importable n8n Workflow Blueprint**:
    - Pre-configured 13-node n8n workflow at `n8n_workflows/meatndoor_campaign.json` ready for 1-click import into n8n Cloud or self-hosted instances.

---

## 3. Incremental Feature Log

| Feature Name | Release Date | Component(s) | Summary |
| :--- | :--- | :--- | :--- |
| **Public Post & Comment Extractor** | Oct 12, 2026 | `instagram_comments.py` | Timeline discovery, comment/reply harvesting, SQLite checkpointing, post limit up to 500. |
| **Dual Google Sheets & Excel Sync** | Oct 12, 2026 | `sheets_sync.py` | Dual sync to Google Sheets and `exports/instagram_research.xlsx` with formulas and frozen headers. |
| **Mumbai Demographics Engine** | Oct 13, 2026 | `mumbai_checker.py` | Profile evidence harvesting, 5-tier classification, negative city rules, and caching of 5,857 accounts. |
| **Streamlit Graphite GUI Suite** | Oct 13, 2026 | `gui_app.py`, `gui_service.py` | 4-tab web application with WebSocket fragments and server-side SQL pagination. |
| **Excel Ingestion & Pre-Check Optimization** | Oct 14, 2026 | `gui_service.py`, `gui_app.py` | `openpyxl` workbook username parser, dynamic sheet selector, and zero-I/O pre-check summaries. |
| **MeatNDoor Lead Outreach Manager** | Oct 14, 2026 | `lead_outreach.py`, `gui_service.py`, `gui_app.py`, `sheets_sync.py` | Incremental ingestion across checkpoints, location suppression, DNC registry, editable drafts, 5th GUI view, Meta compliance links, and dual sync. |
| **n8n Campaign Automation & Lead Scoring** | Oct 15, 2026 | `n8n_client.py`, `campaign_service.py`, `campaign_api.py`, `gui_app.py`, `gui_service.py`, `n8n_workflows/meatndoor_campaign.json` | 10-stage pipeline, dual-axis intent scoring (0-100), interactive Campaign Manager GUI view, n8n webhook orchestrator & workflow blueprint, pause/resume/cancel controls, and campaign Sheets sync. |
| **Resilient GraphQL & Live Session Token Sync** | Oct 30, 2026 | `instagram_comments.py`, `campaign_service.py`, `gui_app.py` | Live token auto-refresh (`fb_dtsg`, `lsd`, `av`, `__rev`, `__spin_r`, `__spin_t`, `__hsi`, `jazoest`, `x-csrftoken`), `_NoRedirectHandler` infinite loop defense, `for (;;);` prefix stripping, direct GraphQL querying with fallback for post discovery, comments, and replies, plus empty search query null safety. |


