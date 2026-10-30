#!/usr/bin/env python3
"""
Streamlit Web GUI for Instagram Public-Post Comment & Reply Exporter.

Features:
- Professional Graphite-Grey UI with restrained styling.
- Instagram username input with auto-stripping of '@' and validation.
- Post limit slider (1 to 500, default 50).
- Feature toggles: comments, replies, Google Sheets sync, Excel workbook, resume.
- Real-time progress monitoring: posts, comments, replies, current shortcode, elapsed time.
- Graceful cancellation support without data or checkpoint corruption.
- In-memory redacted live log stream (cookies, tokens, private keys removed).
- Results dashboard with direct links to Google Spreadsheet master and local Excel mirror.
- History browser for past account checkpoints with one-click re-syncing.
"""

import math
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import streamlit as st

from gui_service import (
    CampaignJobManager,
    ExtractionJobManager,
    MumbaiJobManager,
    LeadOutreachManager,
    extract_usernames_from_excel,
    get_available_checkpoint_accounts,
    get_campaign_manager,
    get_checkpoint_history,
    get_commenters_summary,
    get_job_manager,
    get_mumbai_manager,
    get_outreach_manager,
    resync_existing_account,
    test_google_sheets_connection,
)
from mumbai_checker import MumbaiCheckDB, MumbaiCheckSync

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
    ALL_REVIEW_STATUSES,
    MSG_STATUS_MANUAL,
    MSG_STATUS_READY,
    MSG_STATUS_SENT,
    MUMBAI_OUTREACH_TEMPLATE,
    GENERIC_OUTREACH_TEMPLATE,
    LeadOutreachSync,
)
from sheets_sync import DEFAULT_EXCEL_PATH, DEFAULT_SPREADSHEET_ID


@st.cache_data(ttl=60, show_spinner=False)
def get_cached_available_checkpoint_accounts() -> List[Dict[str, Any]]:
    return get_available_checkpoint_accounts()


@st.cache_data(ttl=120, show_spinner=False)
def get_cached_commenters_summary_counts(sources_tuple: Tuple[str, ...], recheck: bool = False) -> Dict[str, int]:
    available = get_cached_available_checkpoint_accounts()
    sel = set(s.lower() for s in sources_tuple) if sources_tuple else set()
    filtered = [a for a in available if not sel or a["source_username"].lower() in sel]
    discovered = sum(a.get("total_records", 0) for a in filtered)
    unique_approx = sum(a.get("unique_commenters", 0) for a in filtered)

    db = MumbaiCheckDB()
    completed_count = db.get_checks_count()
    already_checked = min(unique_approx, completed_count) if not recheck else 0
    remaining = max(0, unique_approx - already_checked)
    return {
        "discovered_count": discovered,
        "unique_count": unique_approx,
        "already_checked_count": already_checked,
        "remaining_count": remaining,
    }


@st.cache_data(ttl=60, show_spinner=False)
def get_cached_checkpoint_history() -> List[Dict[str, Any]]:
    return get_checkpoint_history()

# Page Configuration
st.set_page_config(
    page_title="Instagram Comment Exporter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Graphite-Grey Modern Theme Styling
GRAPHITE_CSS = """
<style>
    /* Global Graphite Palette */
    :root {
        --bg-main: #141618;
        --bg-card: #1c1f24;
        --bg-card-hover: #23272e;
        --border-color: #2f343d;
        --text-primary: #e6edf3;
        --text-secondary: #9aa2ab;
        --accent-blue: #3b82f6;
        --accent-green: #22c55e;
        --accent-amber: #eab308;
        --accent-red: #ef4444;
    }

    /* Metric cards styling */
    div[data-testid="stMetric"] {
        background-color: #1e2227;
        border: 1px solid #2d333b;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25);
    }
    div[data-testid="stMetricLabel"] {
        color: #8b949e !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    div[data-testid="stMetricValue"] {
        color: #f0f6fc !important;
        font-weight: 700 !important;
        font-size: 1.65rem !important;
    }

    /* Status badge containers */
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .badge-running { background-color: rgba(59, 130, 246, 0.2); color: #60a5fa; border: 1px solid #3b82f6; }
    .badge-success { background-color: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid #22c55e; }
    .badge-warning { background-color: rgba(234, 179, 8, 0.2); color: #facc15; border: 1px solid #eab308; }
    .badge-error { background-color: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }
    .badge-idle { background-color: rgba(156, 163, 175, 0.2); color: #9ca3af; border: 1px solid #6b7280; }

    /* Mumbai Location Categories Badges */
    .badge-likely-mumbai { background-color: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid #22c55e; }
    .badge-serves-mumbai { background-color: rgba(6, 182, 212, 0.2); color: #22d3ee; border: 1px solid #06b6d4; }
    .badge-possibly-mumbai { background-color: rgba(234, 179, 8, 0.2); color: #facc15; border: 1px solid #eab308; }
    .badge-outside-mumbai { background-color: rgba(156, 163, 175, 0.2); color: #d1d5db; border: 1px solid #9ca3af; }
    .badge-not-enough { background-color: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }

    /* MeatNDoor Outreach Badges */
    .badge-eligible-mumbai { background-color: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid #22c55e; }
    .badge-business-review { background-color: rgba(6, 182, 212, 0.2); color: #22d3ee; border: 1px solid #06b6d4; }
    .badge-low-conf-review { background-color: rgba(234, 179, 8, 0.2); color: #facc15; border: 1px solid #eab308; }
    .badge-suppressed-outside { background-color: rgba(156, 163, 175, 0.2); color: #d1d5db; border: 1px solid #9ca3af; }
    .badge-suppressed-dnc { background-color: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }
    .badge-pending-location { background-color: rgba(244, 114, 182, 0.2); color: #f472b6; border: 1px solid #ec4899; }
    .badge-review-pending { background-color: rgba(156, 163, 175, 0.2); color: #9ca3af; border: 1px solid #6b7280; }
    .badge-review-approved { background-color: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid #22c55e; }
    .badge-review-contacted { background-color: rgba(59, 130, 246, 0.2); color: #60a5fa; border: 1px solid #3b82f6; }
    .badge-review-replied { background-color: rgba(168, 85, 247, 0.2); color: #c084fc; border: 1px solid #a855f7; }

    /* Campaign Lead Score Badges */
    .badge-score-high { background-color: rgba(34, 197, 94, 0.25); color: #4ade80; border: 1px solid #22c55e; font-weight: bold; }
    .badge-score-med { background-color: rgba(234, 179, 8, 0.25); color: #facc15; border: 1px solid #eab308; font-weight: bold; }
    .badge-score-low { background-color: rgba(156, 163, 175, 0.25); color: #d1d5db; border: 1px solid #9ca3af; }



    /* Custom Terminal Log Container */
    .terminal-box {
        background-color: #0d1117;
        border: 1px solid #21262d;
        border-radius: 6px;
        padding: 12px;
        font-family: 'JetBrains Mono', 'Fira Code', 'Courier New', monospace;
        font-size: 0.82rem;
        color: #c9d1d9;
        max-height: 280px;
        overflow-y: auto;
        white-space: pre-wrap;
        line-height: 1.45;
    }

    /* Button formatting */
    button[kind="primary"] {
        background-color: #2563eb !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.4rem !important;
        border-radius: 6px !important;
    }
    button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
    }

    /* Clean divider */
    hr {
        border-color: #2d333b !important;
        margin: 1.2rem 0 !important;
    }

    /* Top Segmented Navigation Tabs */
    div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        background-color: #16181d;
        padding: 6px;
        border-radius: 8px;
        border: 1px solid #2d333b;
        margin-bottom: 0.6rem;
    }
    div[data-testid="stRadio"] label {
        padding: 8px 18px !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
        background-color: transparent !important;
        cursor: pointer !important;
        border: 1px solid transparent !important;
        transition: all 0.15s ease-in-out;
    }
    div[data-testid="stRadio"] label:hover {
        background-color: #21262d !important;
    }
    div[data-testid="stRadio"] label:has(input:checked) {
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border-color: #3b82f6 !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.4);
    }
</style>
"""
st.markdown(GRAPHITE_CSS, unsafe_allow_html=True)

# Initialize Job Managers
manager = get_job_manager()
mumbai_manager = get_mumbai_manager()
outreach_manager = get_outreach_manager()
campaign_manager = get_campaign_manager()



def _render_extractor_telemetry_ui(snap: Dict[str, Any], running: bool) -> None:
    st.markdown("##### 📈 Live Progress & Real-Time Telemetry")
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        posts_disp = f"{snap.get('posts_completed', 0)} / {snap.get('total_posts', 0)}" if snap.get("total_posts") else "0 / 0"
        st.metric("Posts Processed", posts_disp)
    with m_col2:
        st.metric("Comments Extracted", f"{snap.get('comments_collected', 0):,}")
    with m_col3:
        st.metric("Child Replies Extracted", f"{snap.get('replies_collected', 0):,}")
    with m_col4:
        secs = snap.get("elapsed_seconds", 0)
        mins, s = divmod(secs, 60)
        st.metric("Elapsed Time", f"{mins:02d}:{s:02d}")

    tot = snap.get("total_posts") or 1
    done = snap.get("posts_completed") or 0
    pct = min(1.0, max(0.0, done / tot))
    st.progress(pct, text=snap.get("status_text", ""))

    if snap.get("current_post_shortcode"):
        shortcode = snap["current_post_shortcode"]
        st.caption(f"Currently active post: **[{shortcode}](https://www.instagram.com/p/{shortcode}/)**")

    with st.expander("📜 Live Execution Logs (Secrets Redacted)", expanded=running):
        log_records = snap.get("logs", [])
        if log_records:
            log_text = "\n".join(f"{r['message']}" for r in log_records[-80:])
            st.markdown(f'<div class="terminal-box">{log_text}</div>', unsafe_allow_html=True)
        else:
            st.caption("No log messages recorded yet.")


@st.fragment(run_every="2s")
def _render_extractor_progress_active() -> None:
    curr_snap = manager.get_snapshot()
    if not curr_snap.get("is_running"):
        st.rerun()
    _render_extractor_telemetry_ui(curr_snap, running=True)


def _render_mumbai_telemetry_ui(snap: Dict[str, Any]) -> None:
    st.markdown("---")
    st.markdown("### 📊 Batch Progress & Telemetry")

    tot_proc = snap.get("total_to_process", 0)
    proc_acc = snap.get("processed_count", 0)
    pct = (proc_acc / tot_proc) if tot_proc > 0 else (1.0 if snap.get("status") == "completed" else 0.0)

    st.progress(pct, text=f"Progress: {proc_acc} of {tot_proc} accounts checked ({int(pct * 100)}%)")
    st.caption(f"**Status**: {snap.get('status_text', '')}")

    m_m1, m_m2, m_m3, m_m4, m_m5, m_m6 = st.columns(6)
    m_m1.metric("Discovered", f"{snap.get('discovered_count', 0):,}")
    m_m2.metric("Unique Accounts", f"{snap.get('deduplicated_count', 0):,}")
    m_m3.metric("Skipped (Cached)", f"{snap.get('skipped_count', 0):,}")
    m_m4.metric("Remaining", f"{max(0, tot_proc - proc_acc):,}")
    m_m5.metric("Newly Checked", f"{proc_acc:,}")
    m_m6.metric("Failures / Errors", f"{snap.get('failed_count', 0):,}")

    c_c1, c_c2, c_c3, c_c4, c_c5 = st.columns(5)
    c_c1.metric("Likely Mumbai", snap.get("likely_mumbai_count", 0))
    c_c2.metric("Serves Mumbai", snap.get("serves_mumbai_count", 0))
    c_c3.metric("Possibly Mumbai", snap.get("possibly_mumbai_count", 0))
    c_c4.metric("Outside Mumbai", snap.get("outside_mumbai_count", 0))
    c_c5.metric("Not Enough Evid.", snap.get("not_enough_evidence_count", 0))


@st.fragment(run_every="2s")
def _render_mumbai_progress_active() -> None:
    curr_snap = mumbai_manager.get_snapshot()
    if not curr_snap.get("is_running"):
        st.rerun()
    _render_mumbai_telemetry_ui(curr_snap)


def _render_campaign_telemetry_ui(snap: Dict[str, Any]) -> None:
    st.markdown("---")
    st.markdown("### 📊 Real-Time Campaign Telemetry & Execution Pipeline")

    tot_proc = snap.get("total_to_process", 0)
    proc_acc = snap.get("leads_processed", 0)
    pct = (proc_acc / tot_proc) if tot_proc > 0 else (1.0 if snap.get("status") == "completed" else 0.0)

    st.progress(pct, text=f"Pipeline: {proc_acc} of {tot_proc} leads evaluated ({int(pct * 100)}%) — Stage: {snap.get('stage', 'Idle')}")

    t1, t2, t3, t4, t5, t6 = st.columns(6)
    t1.metric("Leads Evaluated", f"{proc_acc:,}")
    t2.metric("Qualified (Mumbai)", f"{snap.get('leads_qualified', 0):,}")
    t3.metric("High-Intent Leads", f"{snap.get('leads_high_intent', 0):,}")
    t4.metric("Messages Drafted", f"{snap.get('messages_drafted', 0):,}")
    t5.metric("Suppressed / Excluded", f"{snap.get('leads_suppressed', 0):,}")
    secs = snap.get("elapsed_seconds", 0)
    mins, s = divmod(secs, 60)
    t6.metric("Elapsed Time", f"{mins:02d}:{s:02d}")

    n8n_trig = snap.get("n8n_triggered", False)
    n8n_id = snap.get("n8n_execution_id", "")
    err_msg = snap.get("error_message", "")

    c_n1, c_n2 = st.columns([1.5, 2.5])
    with c_n1:
        if n8n_trig:
            st.success(f"🌐 n8n Webhook: Active (Exec: `{n8n_id or 'queued'}`)", icon="🔗")
        else:
            st.info("⚡ Execution Engine: Native Python Core Pipeline (n8n Ready)", icon="⚙️")
    with c_n2:
        if err_msg:
            st.warning(f"⚠️ Telemetry Alert: {err_msg}")
        else:
            st.caption("All pipeline services operational: SQLite checkpoints, dual-axis scoring & dual-sync active.")


@st.fragment(run_every="2s")
def _render_campaign_progress_active() -> None:
    curr_snap = campaign_manager.get_snapshot()
    if not curr_snap.get("is_running"):
        st.rerun()
    _render_campaign_telemetry_ui(curr_snap)


# Sidebar: System Status & Navigation
with st.sidebar:
    st.markdown("### 📸 Instagram Extractor")
    st.caption("Dual Sync: Google Sheets + Local Excel")
    st.markdown("---")

    # System Diagnostics summary
    st.markdown("##### 🔌 System Integrations")
    service_acc_file = Path("service_account.json")
    if service_acc_file.exists():
        st.success("Google Sheets: Service Account active", icon="✅")
    elif Path("token.json").exists():
        st.success("Google Sheets: OAuth Token active", icon="✅")
    else:
        st.warning("Google Sheets: Credentials missing", icon="⚠️")

    excel_file = Path(DEFAULT_EXCEL_PATH)
    if excel_file.exists():
        size_kb = int(excel_file.stat().st_size / 1024)
        st.info(f"Excel Mirror: {size_kb} KB", icon="📁")
    else:
        st.caption("Excel Mirror: Not created yet")

    st.markdown("---")
    st.markdown("##### 🔗 Quick Links")
    st.markdown(
        f"- [Master Google Spreadsheet](https://docs.google.com/spreadsheets/d/{DEFAULT_SPREADSHEET_ID}/edit)\n"
        f"- Local Workbook: `{DEFAULT_EXCEL_PATH}`"
    )

    st.markdown("---")
    st.caption("Antigravity Extraction Suite • v2.1")

# View Navigation (Isolated Execution - Only active page executes)
nav_options = [
    "🎯 Campaign Manager",
    "📍 Mumbai Location Checker",
    "🚀 Extractor Dashboard",
    "🍗 MeatNDoor Lead Outreach",
    "📚 History & Checkpoints",
    "🔑 Google Sheets Health",
]


if "active_nav_view" not in st.session_state:
    st.session_state["active_nav_view"] = "🎯 Campaign Manager"


selected_view = st.radio(
    "Navigation View",
    nav_options,
    index=nav_options.index(st.session_state["active_nav_view"]) if st.session_state["active_nav_view"] in nav_options else 0,
    horizontal=True,
    label_visibility="collapsed",
    key="main_nav_radio",
)
st.session_state["active_nav_view"] = selected_view

# =========================================================================
# VIEW 0: CAMPAIGN MANAGER (n8n AUTOMATION)
# =========================================================================
if selected_view == "🎯 Campaign Manager":
    c_snap = campaign_manager.get_snapshot()
    c_is_running = c_snap.get("is_running", False)

    col_camp_h1, col_camp_h2 = st.columns([3, 1])
    with col_camp_h1:
        st.markdown("## 🎯 MeatNDoor Campaign Manager (n8n Automated)")
        st.caption("Coordinate end-to-end commenter research, Mumbai qualification, dual-axis intent scoring (0-100), personalized draft generation, and dual Sheets/Excel sync.")
    with col_camp_h2:
        c_status = c_snap.get("status", "idle")
        if c_is_running:
            st.markdown('<div class="status-badge badge-running">● RUNNING</div>', unsafe_allow_html=True)
        elif c_status == "completed":
            st.markdown('<div class="status-badge badge-success">✓ COMPLETED</div>', unsafe_allow_html=True)
        elif c_status == "paused":
            st.markdown('<div class="status-badge badge-warning">⏸ PAUSED</div>', unsafe_allow_html=True)
        elif c_status == "cancelled":
            st.markdown('<div class="status-badge badge-warning">■ CANCELLED</div>', unsafe_allow_html=True)
        elif c_status == "failed":
            st.markdown('<div class="status-badge badge-error">✕ FAILED</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-badge badge-idle">○ IDLE</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Campaign Selector & Configuration
    existing_campaigns = campaign_manager.list_campaigns()
    camp_options = ["[+ Create New Campaign]"] + [f"{c['campaign_id']} | {c['name']} ({c['status'].upper()})" for c in existing_campaigns]

    if "selected_campaign_selector" not in st.session_state:
        st.session_state["selected_campaign_selector"] = camp_options[1] if len(camp_options) > 1 else camp_options[0]

    sel_idx = 0
    if st.session_state["selected_campaign_selector"] in camp_options:
        sel_idx = camp_options.index(st.session_state["selected_campaign_selector"])

    selected_camp_item = st.selectbox(
        "Select Active Campaign or Create New",
        options=camp_options,
        index=sel_idx,
        disabled=c_is_running,
        key="camp_select_box",
    )
    st.session_state["selected_campaign_selector"] = selected_camp_item

    is_new = selected_camp_item == "[+ Create New Campaign]"
    active_camp_data: Dict[str, Any] = {}
    if not is_new:
        sel_cid = selected_camp_item.split(" | ")[0].strip()
        active_camp_data = campaign_manager.get_campaign(sel_cid) or {}

    with st.container():
        st.markdown("##### ⚙️ Campaign Configuration")
        c_col1, c_col2 = st.columns(2)
        with c_col1:
            cfg_name = st.text_input(
                "Campaign Name",
                value=active_camp_data.get("name", "MeatNDoor Mumbai Commenter Outreach"),
                disabled=c_is_running,
                help="A clear descriptor for the audience or batch.",
            )
            cfg_obj = st.selectbox(
                "Campaign Objective",
                options=[
                    "Cold Commenter Conversion & Delivery Verification",
                    "Customer Research & Demand Assessment",
                    "New Meat Product Launch Awareness",
                    "Brand Re-engagement & Feedback Collection",
                ],
                index=0 if not active_camp_data else [
                    "Cold Commenter Conversion & Delivery Verification",
                    "Customer Research & Demand Assessment",
                    "New Meat Product Launch Awareness",
                    "Brand Re-engagement & Feedback Collection",
                ].index(active_camp_data.get("objective", "Cold Commenter Conversion & Delivery Verification")) if active_camp_data.get("objective") in [
                    "Cold Commenter Conversion & Delivery Verification",
                    "Customer Research & Demand Assessment",
                    "New Meat Product Launch Awareness",
                    "Brand Re-engagement & Feedback Collection",
                ] else 0,
                disabled=c_is_running,
            )
            cfg_loc = st.selectbox(
                "Target Location",
                options=["Mumbai", "Navi Mumbai", "Mumbai Metropolitan Region", "All Maharashtra"],
                index=0 if not active_camp_data else ["Mumbai", "Navi Mumbai", "Mumbai Metropolitan Region", "All Maharashtra"].index(active_camp_data.get("target_location", "Mumbai")) if active_camp_data.get("target_location") in ["Mumbai", "Navi Mumbai", "Mumbai Metropolitan Region", "All Maharashtra"] else 0,
                disabled=c_is_running,
            )

        with c_col2:
            cfg_max_leads = st.slider(
                "Maximum Leads to Process",
                min_value=10,
                max_value=1000,
                value=int(active_camp_data.get("max_leads", 100)),
                step=10,
                disabled=c_is_running,
                help="Caps the batch size to maintain healthy API and workflow pacing.",
            )
            cfg_min_score = st.slider(
                "Minimum Lead Qualification Score (0-100)",
                min_value=0,
                max_value=100,
                value=int(active_camp_data.get("min_lead_score", 30)),
                step=5,
                disabled=c_is_running,
                help="Dual-axis qualification score: Location Score (0-50) + Intent Score (0-50).",
            )
            cfg_n8n_url = st.text_input(
                "n8n Webhook URL (Optional)",
                value=os.environ.get("N8N_CAMPAIGN_WEBHOOK_URL", ""),
                disabled=c_is_running,
                placeholder="http://localhost:5678/webhook/meatndoor-campaign",
                help="When configured, clicking 'Start Campaign' invokes the n8n orchestrator workflow.",
            )

        avail_sources = get_cached_available_checkpoint_accounts()
        avail_src_names = [a["source_username"] for a in avail_sources] if avail_sources else []
        default_selected_src = active_camp_data.get("source_accounts", avail_src_names) if not is_new else avail_src_names
        valid_default_src = [s for s in default_selected_src if s in avail_src_names]

        c_col3, c_col4 = st.columns([2, 2])
        with c_col3:
            cfg_sources = st.multiselect(
                "Source Profiles / Checkpoint Datasets",
                options=avail_src_names,
                default=valid_default_src,
                disabled=c_is_running,
                help="Select brand checkpoints to ingest commenter leads from.",
            )
        with c_col4:
            cfg_notes = st.text_input(
                "Campaign Notes (Optional)",
                value=active_camp_data.get("notes", ""),
                disabled=c_is_running,
                placeholder="e.g. Focus on mutton and fresh meat queries from competitor posts",
            )

    st.markdown("##### 🕹️ Primary Action Controls")
    btn_c1, btn_c2, btn_c3, btn_c4, btn_c5, btn_c6 = st.columns(6)

    cur_cid = active_camp_data.get("campaign_id", "") if not is_new else f"camp_{int(time.time())}"
    cur_status = active_camp_data.get("status", "idle")

    with btn_c1:
        start_btn = st.button(
            "▶️ Start Campaign",
            type="primary",
            use_container_width=True,
            disabled=c_is_running,
            help="Initiates full pipeline: lead extraction, qualification, intent scoring, draft generation, and sync.",
        )
    with btn_c2:
        pause_btn = st.button(
            "⏸️ Pause",
            type="secondary",
            use_container_width=True,
            disabled=(not c_is_running),
            help="Signals current batch to safely pause at next checkpoint boundary.",
        )
    with btn_c3:
        resume_btn = st.button(
            "⏯️ Resume",
            type="secondary",
            use_container_width=True,
            disabled=(c_is_running or cur_status != "paused"),
            help="Resumes execution for paused campaign.",
        )
    with btn_c4:
        cancel_btn = st.button(
            "⏹️ Cancel",
            type="secondary",
            use_container_width=True,
            disabled=(not c_is_running and cur_status not in ("running", "paused")),
            help="Safely cancels execution and marks campaign cancelled.",
        )
    with btn_c5:
        refresh_btn = st.button(
            "🔄 Refresh Status",
            type="secondary",
            use_container_width=True,
            help="Reruns the page to fetch latest telemetry and review records.",
        )
    with btn_c6:
        sync_btn = st.button(
            "📊 Sync Sheets",
            type="secondary",
            use_container_width=True,
            disabled=(c_is_running or is_new or not cur_cid),
            help="Syncs campaign leads to Google Sheets ('Campaign Leads' tab) and local Excel.",
        )

    if start_btn:
        if not cfg_name.strip():
            st.error("Please provide a valid campaign name.")
        else:
            if cfg_n8n_url.strip():
                os.environ["N8N_CAMPAIGN_WEBHOOK_URL"] = cfg_n8n_url.strip()

            target_cid = cur_cid if not is_new else f"camp_{re.sub(r'[^a-zA-Z0-9_]', '_', cfg_name.lower())[:25]}_{int(time.time())}"
            ok, msg = campaign_manager.start_campaign_job(
                campaign_id=target_cid,
                name=cfg_name.strip(),
                objective=cfg_obj,
                target_location=cfg_loc,
                source_accounts=cfg_sources,
                max_leads=int(cfg_max_leads),
                min_lead_score=int(cfg_min_score),
                notes=cfg_notes.strip(),
            )
            if ok:
                st.success(f"🚀 {msg}")
                st.session_state["selected_campaign_selector"] = f"{target_cid} | {cfg_name.strip()} (RUNNING)"
                st.rerun()
            else:
                st.error(f"Cannot start campaign: {msg}")

    elif pause_btn:
        if cur_cid:
            campaign_manager.pause_campaign_job(cur_cid)
            st.info(f"⏸️ Pause signal sent to campaign '{cur_cid}'.")
            st.rerun()

    elif resume_btn:
        if cur_cid:
            ok, msg = campaign_manager.resume_campaign_job(cur_cid)
            if ok:
                st.success(f"⏯️ Resumed campaign '{cur_cid}'.")
                st.rerun()
            else:
                st.error(f"Failed to resume: {msg}")

    elif cancel_btn:
        if cur_cid:
            campaign_manager.cancel_campaign_job(cur_cid)
            st.warning(f"⏹️ Cancel signal sent to campaign '{cur_cid}'.")
            st.rerun()

    elif refresh_btn:
        st.rerun()

    elif sync_btn:
        if cur_cid:
            with st.spinner("Syncing campaign leads to Google Sheets & Excel..."):
                sync_out = campaign_manager.sync_campaign_to_sheets(cur_cid)
                sh_res = sync_out.get("sheets_res", {})
                sh_cnt = sh_res.get("rows_synced", 0)
                ex_cnt = sync_out.get("excel_rows", 0)
                st.success(f"✓ Synced {sh_cnt} rows to Google Sheets tab 'Campaign Leads' and {ex_cnt} rows to Excel workbook.")

    if c_is_running:
        _render_campaign_progress_active()
    elif c_snap.get("campaign_id") or active_camp_data:
        _render_campaign_telemetry_ui(c_snap if c_snap.get("campaign_id") == cur_cid else {
            "status": active_camp_data.get("status", "idle"),
            "stage": active_camp_data.get("current_stage", "Ready"),
            "total_to_process": active_camp_data.get("total_leads_processed", 0),
            "leads_processed": active_camp_data.get("total_leads_processed", 0),
            "leads_qualified": active_camp_data.get("qualified_leads_count", 0),
            "leads_high_intent": active_camp_data.get("high_intent_count", 0),
            "messages_drafted": active_camp_data.get("messages_drafted", 0),
            "leads_suppressed": max(0, active_camp_data.get("total_leads_processed", 0) - active_camp_data.get("qualified_leads_count", 0)),
            "elapsed_seconds": 0,
            "n8n_triggered": False,
            "n8n_execution_id": "",
            "error_message": active_camp_data.get("sheets_sync_error", ""),
        })

    st.markdown("---")
    st.markdown("### 📋 Campaign Review Queue & Lead Qualification")

    if is_new or not cur_cid:
        st.info("Select an existing campaign from the dropdown above or click 'Start Campaign' to create and run a new campaign.")
    else:
        q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns([2, 1.2, 1.2, 1.2, 1.2])
        with q_col1:
            q_search = st.text_input("Search Username or Signals", placeholder="Search username, keyword...", key=f"q_search_{cur_cid}")
        with q_col2:
            q_app_status = st.selectbox(
                "Approval Status",
                options=["All", "pending", "approved", "rejected"],
                index=0,
                key=f"q_app_{cur_cid}",
            )
        with q_col3:
            q_contact_status = st.selectbox(
                "Contact Status",
                options=["All", "not_contacted", "contacted", "replied", "interested", "not_interested", "dnc"],
                index=0,
                key=f"q_contact_{cur_cid}",
            )
        with q_col4:
            q_sort = st.selectbox(
                "Sort By",
                options=["Score: High to Low", "Score: Low to High", "Username", "Recently Updated"],
                index=0,
                key=f"q_sort_{cur_cid}",
            )
        with q_col5:
            q_page_size = st.selectbox("Page Size", options=[10, 25, 50], index=1, key=f"q_ps_{cur_cid}")

        sort_col_map = {
            "Score: High to Low": "lead_score DESC",
            "Score: Low to High": "lead_score ASC",
            "Username": "username ASC",
            "Recently Updated": "last_action_at DESC",
        }
        sort_by = sort_col_map.get(q_sort, "lead_score DESC")

        page_key = f"camp_page_{cur_cid}"
        if page_key not in st.session_state:
            st.session_state[page_key] = 1

        leads_slice, total_matching = campaign_manager.get_campaign_leads_paginated(
            campaign_id=cur_cid,
            approval_status=None if q_app_status == "All" else q_app_status,
            contact_status=None if q_contact_status == "All" else q_contact_status,
            search=q_search.strip() if q_search else "",
            sort_by=sort_by,
            limit=q_page_size,
            offset=(st.session_state[page_key] - 1) * q_page_size,
        )

        total_pages = max(1, (total_matching + q_page_size - 1) // q_page_size)
        if st.session_state[page_key] > total_pages:
            st.session_state[page_key] = total_pages
        curr_page = st.session_state[page_key]

        p_c1, p_c2, p_c3, p_c4 = st.columns([1.2, 1.6, 2.8, 1.2])
        with p_c1:
            if st.button("⬅️ Previous", disabled=(curr_page <= 1), use_container_width=True, key=f"p_prev_{cur_cid}"):
                st.session_state[page_key] = max(1, curr_page - 1)
                st.rerun()
        with p_c2:
            st.markdown(f"<div style='text-align: center; padding-top: 6px; font-weight: 600;'>Page {curr_page} of {total_pages}</div>", unsafe_allow_html=True)
        with p_c3:
            start_num = min((curr_page - 1) * q_page_size + 1, total_matching) if total_matching > 0 else 0
            end_num = min(curr_page * q_page_size, total_matching)
            st.caption(f"Showing **{start_num}–{end_num}** of **{total_matching:,}** matching leads in campaign `{cur_cid}`")
        with p_c4:
            if st.button("Next ➡️", disabled=(curr_page >= total_pages), use_container_width=True, key=f"p_next_{cur_cid}"):
                st.session_state[page_key] = min(total_pages, curr_page + 1)
                st.rerun()

        if not leads_slice:
            st.info("No leads match the selected filter criteria for this campaign.")
        else:
            for lead in leads_slice:
                lead_u = lead["username"]
                lead_score = int(lead.get("lead_score", 0))
                intent_score = int(lead.get("intent_score", 0))
                loc_cls = lead.get("location_category", "Unknown")
                app_st = lead.get("approval_status", "pending")
                cnt_st = lead.get("contact_status", "not_contacted")
                supp_reason = lead.get("suppression_reason", "")

                score_badge_class = "badge-score-high" if lead_score >= 70 else ("badge-score-med" if lead_score >= 40 else "badge-score-low")
                app_badge_class = "badge-review-approved" if app_st == "approved" else ("badge-error" if app_st == "rejected" else "badge-review-pending")
                cnt_badge_class = "badge-review-contacted" if cnt_st == "contacted" else ("badge-review-replied" if cnt_st in ("replied", "interested") else "badge-idle")

                with st.expander(f"👤 @{lead_u} — Score: {lead_score}/100 [{loc_cls}] — Approval: {app_st.upper()}", expanded=(app_st == "pending")):
                    hdr_c1, hdr_c2, hdr_c3, hdr_c4 = st.columns([2, 1.5, 1.5, 1])
                    with hdr_c1:
                        st.markdown(f"**Username**: [`@{lead_u}`](https://instagram.com/{lead_u}) &nbsp;•&nbsp; **Sources**: `{lead.get('source_accounts', '—')}`", unsafe_allow_html=True)
                        st.caption(f"Location: **{loc_cls}** &nbsp;|&nbsp; Confidence: **{lead.get('location_confidence', 'Low')}**")
                    with hdr_c2:
                        st.markdown(f'<span class="status-badge {score_badge_class}">Score: {lead_score}/100</span> (Intent: {intent_score}/50)', unsafe_allow_html=True)
                    with hdr_c3:
                        st.markdown(f'<span class="status-badge {app_badge_class}">Approval: {app_st.upper()}</span> &nbsp; <span class="status-badge {cnt_badge_class}">{cnt_st.upper()}</span>', unsafe_allow_html=True)
                    with hdr_c4:
                        st.link_button("💬 Open IG Chat", f"https://ig.me/m/{lead_u}", use_container_width=True)

                    raw_sig = lead.get("intent_evidence", [])
                    signals = json.loads(raw_sig) if isinstance(raw_sig, str) and raw_sig.startswith("[") else (raw_sig if isinstance(raw_sig, list) else [])
                    if signals:
                        sig_str = " &nbsp;•&nbsp; ".join([f"`{s}`" for s in signals])
                        st.markdown(f"**Intent Signals**: {sig_str}", unsafe_allow_html=True)

                    if supp_reason:
                        st.markdown(f'<div class="status-badge badge-suppressed-outside">Suppression Reason: {supp_reason}</div>', unsafe_allow_html=True)

                    draft_text = st.text_area(
                        "Personalized Message Draft",
                        value=lead.get("suggested_message", ""),
                        key=f"draft_{cur_cid}_{lead_u}",
                        height=100,
                    )

                    act_c1, act_c2, act_c3, act_c4, act_c5, act_c6 = st.columns([1.5, 1.5, 1.5, 2, 1.5, 1.5])
                    with act_c1:
                        if st.button("✅ Approve", key=f"btn_app_{cur_cid}_{lead_u}", use_container_width=True, help="Approve draft for manual outreach. Does NOT send."):
                            campaign_manager.approve_lead_draft(cur_cid, lead_u)
                            st.success(f"Approved @{lead_u}")
                            st.rerun()
                    with act_c2:
                        if st.button("💾 Save Draft", key=f"btn_sav_{cur_cid}_{lead_u}", use_container_width=True):
                            campaign_manager.edit_lead_draft(cur_cid, lead_u, draft_text)
                            st.success(f"Saved draft for @{lead_u}")
                            st.rerun()
                    with act_c3:
                        if st.button("❌ Reject", key=f"btn_rej_{cur_cid}_{lead_u}", use_container_width=True):
                            campaign_manager.reject_lead_draft(cur_cid, lead_u)
                            st.warning(f"Rejected @{lead_u}")
                            st.rerun()
                    with act_c4:
                        new_cnt_status = st.selectbox(
                            "Status",
                            options=["not_contacted", "contacted", "replied", "interested", "not_interested"],
                            index=["not_contacted", "contacted", "replied", "interested", "not_interested"].index(cnt_st) if cnt_st in ["not_contacted", "contacted", "replied", "interested", "not_interested"] else 0,
                            key=f"sel_cnt_{cur_cid}_{lead_u}",
                            label_visibility="collapsed",
                        )
                        if new_cnt_status != cnt_st:
                            campaign_manager.update_lead_contact_status(cur_cid, lead_u, new_cnt_status)
                            st.rerun()
                    with act_c5:
                        if st.button("🚫 DNC", key=f"btn_dnc_{cur_cid}_{lead_u}", use_container_width=True, help="Add user to persistent Do-Not-Contact list"):
                            campaign_manager.add_lead_to_dnc(cur_cid, lead_u)
                            st.error(f"Added @{lead_u} to DNC list")
                            st.rerun()
                    with act_c6:
                        if st.button("🗑️ Remove", key=f"btn_rm_{cur_cid}_{lead_u}", use_container_width=True):
                            campaign_manager.remove_lead_from_campaign(cur_cid, lead_u)
                            st.info(f"Removed @{lead_u} from campaign")
                            st.rerun()

# =========================================================================
# VIEW 1: EXTRACTOR DASHBOARD
# =========================================================================
elif selected_view == "🚀 Extractor Dashboard":
    snapshot = manager.get_snapshot()
    is_running = snapshot["is_running"]

    # Header section
    col_hdr1, col_hdr2 = st.columns([3, 1])
    with col_hdr1:
        st.markdown("## Instagram Comment & Reply Extractor")
        st.caption("Extracts posts, top-level comments, and indented child replies directly to Google Sheets & Excel.")
    with col_hdr2:
        status = snapshot["status"]
        if is_running:
            st.markdown('<div class="status-badge badge-running">● RUNNING</div>', unsafe_allow_html=True)
        elif status == "completed":
            st.markdown('<div class="status-badge badge-success">✓ COMPLETED</div>', unsafe_allow_html=True)
        elif status == "cancelled":
            st.markdown('<div class="status-badge badge-warning">■ CANCELLED</div>', unsafe_allow_html=True)
        elif status == "error":
            st.markdown('<div class="status-badge badge-error">✕ ERROR</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-badge badge-idle">○ IDLE</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Configuration Form
    with st.container():
        st.markdown("##### 🎯 Target & Limits")
        col_u, col_l = st.columns([2, 1])
        with col_u:
            target_username = st.text_input(
                "Instagram Username",
                value="licious_foods",
                placeholder="e.g. licious_foods, prasuma_official, meatigo_official",
                help="Accepts username with or without '@'. Captures public posts automatically.",
                disabled=is_running,
            )
        with col_l:
            post_limit = st.slider(
                "Post Limit",
                min_value=1,
                max_value=500,
                value=50,
                step=1,
                help="Maximum unique posts to discover and extract from the profile timeline.",
                disabled=is_running,
            )

        st.markdown("##### ⚙️ Extraction Features")
        col_opt1, col_opt2, col_opt3 = st.columns(3)
        with col_opt1:
            opt_comments = st.checkbox("Extract Top-Level Comments", value=True, disabled=is_running)
            opt_replies = st.checkbox("Extract Child Replies", value=True, disabled=is_running)
        with col_opt2:
            opt_sheets = st.checkbox("Sync to Google Spreadsheet", value=True, disabled=is_running)
            opt_excel = st.checkbox("Save to Local Excel Workbook", value=True, disabled=is_running)
        with col_opt3:
            opt_resume = st.checkbox("Resume Interrupted Runs", value=True, disabled=is_running)

        # Advanced Settings
        with st.expander("🛠️ Advanced Extraction Parameters", expanded=False):
            col_adv1, col_adv2 = st.columns(2)
            with col_adv1:
                adv_delay = st.slider(
                    "Request Delay (seconds)",
                    min_value=0.5,
                    max_value=5.0,
                    value=2.0,
                    step=0.5,
                    help="Conservative rate-limiting pause between GraphQL pagination calls.",
                    disabled=is_running,
                )
                adv_max_comments = st.number_input(
                    "Max Comments per Post (0 = unlimited)",
                    min_value=0,
                    max_value=5000,
                    value=0,
                    disabled=is_running,
                )
            with col_adv2:
                adv_spreadsheet_id = st.text_input(
                    "Google Spreadsheet ID",
                    value=DEFAULT_SPREADSHEET_ID,
                    disabled=is_running,
                )
                adv_excel_path = st.text_input(
                    "Excel Workbook Destination",
                    value=DEFAULT_EXCEL_PATH,
                    disabled=is_running,
                )
                adv_max_replies = st.number_input(
                    "Max Replies per Comment (0 = unlimited)",
                    min_value=0,
                    max_value=500,
                    value=0,
                    disabled=is_running,
                )

        # Action Buttons
        col_act1, col_act2, col_act3 = st.columns([1.5, 1.5, 4])
        with col_act1:
            start_clicked = st.button(
                "🚀 Start Extraction",
                type="primary",
                use_container_width=True,
                disabled=is_running,
            )
        with col_act2:
            stop_clicked = st.button(
                "⏹ Stop / Cancel",
                type="secondary",
                use_container_width=True,
                disabled=not is_running,
            )

        if start_clicked:
            clean_u = target_username.strip().lstrip("@")
            if not clean_u:
                st.error("Please enter a valid Instagram username.")
            elif not re.match(r"^[A-Za-z0-9._]{1,30}$", clean_u):
                st.error(f"'{clean_u}' contains invalid characters for an Instagram username.")
            else:
                ok, msg = manager.start_job(
                    username=clean_u,
                    limit_posts=post_limit,
                    extract_comments=opt_comments,
                    extract_replies=opt_replies,
                    sync_excel=opt_excel,
                    sync_sheets=opt_sheets,
                    resume=opt_resume,
                    delay=adv_delay,
                    max_comments_per_post=(adv_max_comments if adv_max_comments > 0 else None),
                    max_replies_per_comment=(adv_max_replies if adv_max_replies > 0 else None),
                    spreadsheet_id=adv_spreadsheet_id,
                    excel_path=adv_excel_path,
                )
                if ok:
                    st.toast(f"Extraction started for @{clean_u}!", icon="🚀")
                    st.rerun()
                else:
                    st.error(msg)

        if stop_clicked:
            ok, msg = manager.cancel_job()
            if ok:
                st.warning(msg)
                time.sleep(1)
                st.rerun()
            else:
                st.error(msg)

    st.markdown("---")

    # =========================================================================
    # LIVE PROGRESS & METRICS
    # =========================================================================
    if is_running:
        _render_extractor_progress_active()
    elif snapshot.get("posts_completed", 0) > 0 or snapshot.get("total_posts", 0) > 0:
        _render_extractor_telemetry_ui(snapshot, running=False)

    # =========================================================================
    # RESULTS CARD (Shown on completed or cancelled)
    # =========================================================================
    if not is_running and snapshot["status"] in ("completed", "cancelled", "error"):
        st.markdown("---")
        st.markdown("### 🏁 Run Summary & Output Destinations")

        res_col1, res_col2 = st.columns(2)
        with res_col1:
            st.markdown("##### 🌐 Online Master (Google Sheets)")
            if snapshot.get("sheets_synced"):
                tab_name = snapshot.get("username", "")
                ss_url = f"https://docs.google.com/spreadsheets/d/{snapshot.get('spreadsheet_id', DEFAULT_SPREADSHEET_ID)}/edit"
                st.success(
                    f"**Synchronized Successfully!**\n\n"
                    f"- Worksheet Tab: **`{tab_name}`**\n"
                    f"- Master Spreadsheet ID: `{snapshot.get('spreadsheet_id', DEFAULT_SPREADSHEET_ID)}`",
                    icon="✅",
                )
                st.link_button("🔗 Open Master Google Spreadsheet", ss_url, type="primary")
            elif snapshot.get("sheets_pending_auth"):
                st.warning(
                    f"**Pending Authorization**\n\n"
                    f"{snapshot.get('sheets_message', 'Credentials missing.')}",
                    icon="⚠️",
                )
            else:
                st.info("Google Sheets synchronization was either disabled or skipped.")

        with res_col2:
            st.markdown("##### 💻 Local Mirror (Excel Workbook)")
            if snapshot.get("excel_updated"):
                excel_p = snapshot.get("excel_path", DEFAULT_EXCEL_PATH)
                st.success(
                    f"**Updated Successfully!**\n\n"
                    f"- Local Path: `{excel_p}`\n"
                    f"- Tab Name: **`{snapshot.get('username', '')}`**\n"
                    f"- Hierarchical comments & replies saved with clickable links.",
                    icon="✅",
                )
            else:
                st.info("Local Excel export was not generated or was disabled.")

        # Re-sync action button
        if snapshot.get("username"):
            u_name = snapshot["username"]
            if st.button(f"🔄 Re-sync @{u_name} from Checkpoint to Google Sheets & Excel"):
                with st.spinner(f"Re-syncing @{u_name}..."):
                    try:
                        res = resync_existing_account(
                            u_name,
                            spreadsheet_id=snapshot.get("spreadsheet_id", DEFAULT_SPREADSHEET_ID),
                            excel_path=snapshot.get("excel_path", DEFAULT_EXCEL_PATH),
                        )
                        if res.get("sheets_success"):
                            st.success(f"Successfully re-synced @{u_name} to Google Sheets and Excel!")
                        elif res.get("sheets_pending_auth"):
                            st.warning(f"Excel updated, but Google Sheets needs authorization:\n{res.get('sheets_message')}")
                        else:
                            st.info(f"Excel records: {res.get('excel_records')}. Message: {res.get('sheets_message')}")
                    except Exception as e:
                        st.error(f"Re-sync failed: {e}")


# =========================================================================
# VIEW 2: MUMBAI LOCATION CHECKER
# =========================================================================
elif selected_view == "📍 Mumbai Location Checker":
    m_snapshot = mumbai_manager.get_snapshot()
    m_is_running = m_snapshot["is_running"]

    # Header
    col_m_hdr1, col_m_hdr2 = st.columns([3, 1])
    with col_m_hdr1:
        st.markdown("## 📍 Mumbai Location Checker")
        st.caption("Batch audit Instagram commenters and reply authors from scraped brand posts for auditable Mumbai location evidence.")
    with col_m_hdr2:
        m_status = m_snapshot["status"]
        if m_is_running:
            st.markdown('<div class="status-badge badge-running">● RUNNING</div>', unsafe_allow_html=True)
        elif m_status == "completed":
            st.markdown('<div class="status-badge badge-success">✓ COMPLETED</div>', unsafe_allow_html=True)
        elif m_status == "cancelled":
            st.markdown('<div class="status-badge badge-warning">■ STOPPED</div>', unsafe_allow_html=True)
        elif m_status == "error":
            st.markdown('<div class="status-badge badge-error">✕ ERROR</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-badge badge-idle">○ IDLE</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Input Form
    with st.container():
        st.markdown("##### 📥 Input Target Commenters")
        input_mode = st.radio(
            "Target Accounts Input Source",
            [
                "👥 Scraped Source Profiles (Commenters & Repliers)",
                "📝 Paste Usernames (Manual)",
                "📂 Import from Excel Workbook",
            ],
            horizontal=True,
            disabled=m_is_running,
        )

        selected_source_accounts: Optional[List[str]] = None
        extracted_usernames_list: List[str] = []

        if input_mode == "👥 Scraped Source Profiles (Commenters & Repliers)":
            available_sources = get_cached_available_checkpoint_accounts()
            if not available_sources:
                st.warning("No scraped profile checkpoints found yet. Scrape an account in Tab 1 first or choose 'Paste Usernames'.")
            else:
                source_names = [a["source_username"] for a in available_sources]
                col_sel1, col_sel2 = st.columns([1.5, 3.5])
                with col_sel1:
                    select_all_sources = st.checkbox(
                        "Audit All Available Sources",
                        value=True,
                        disabled=m_is_running,
                        help="Includes commenters across all scraped brand checkpoints.",
                    )
                with col_sel2:
                    if select_all_sources:
                        selected_source_accounts = source_names
                        st.caption(f"Selecting all {len(source_names)} profiles: {', '.join([f'@{s}' for s in source_names])}")
                    else:
                        selected_source_accounts = st.multiselect(
                            "Select Scraped Source Account(s)",
                            options=source_names,
                            default=[source_names[0]] if source_names else [],
                            format_func=lambda s: f"@{s} ({next((a['unique_commenters'] for a in available_sources if a['source_username'] == s), 0)} unique commenters)",
                            disabled=m_is_running,
                        )

        elif input_mode == "📝 Paste Usernames (Manual)":
            raw_input = st.text_area(
                "Instagram Usernames",
                value="thebombaycanteen\nlicious_foods\nfishtokri\nprasuma_official\nmeatigo_official",
                height=110,
                placeholder="Paste usernames separated by newlines, commas, or spaces (e.g. user1, @user2)",
                help="Paste accounts to audit. Handles '@' prefixes automatically.",
                disabled=m_is_running,
            )
            for part in re.split(r"[\n,;\s]+", raw_input):
                c = part.strip().lstrip("@")
                if c and re.match(r"^[A-Za-z0-9._]{1,30}$", c):
                    if c.lower() not in [u.lower() for u in extracted_usernames_list]:
                        extracted_usernames_list.append(c)

        else:
            excel_target = Path(DEFAULT_EXCEL_PATH)
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                if excel_target.exists():
                    wb_temp = openpyxl.load_workbook(str(excel_target), read_only=True)
                    available_sheets = [s for s in wb_temp.sheetnames if s != "Mumbai Check"]
                    wb_temp.close()
                    selected_sheet = st.selectbox(
                        "Import from Local Excel Tab",
                        available_sheets,
                        disabled=m_is_running,
                    )
                    if st.button("📋 Load Tab Usernames", disabled=m_is_running):
                        extracted_usernames_list = extract_usernames_from_excel(str(excel_target), sheet_name=selected_sheet)
                        st.session_state["mumbai_imported_usernames"] = extracted_usernames_list
                        st.success(f"Extracted {len(extracted_usernames_list)} accounts from '{selected_sheet}'!")
            with col_ex2:
                uploaded_file = st.file_uploader(
                    "Or Upload Excel Workbook (.xlsx)",
                    type=["xlsx"],
                    disabled=m_is_running,
                )
                if uploaded_file:
                    try:
                        extracted_usernames_list = extract_usernames_from_excel(uploaded_file)
                        st.session_state["mumbai_imported_usernames"] = extracted_usernames_list
                        st.success(f"Extracted {len(extracted_usernames_list)} accounts from uploaded file!")
                    except Exception as e:
                        st.error(f"Error reading uploaded file: {e}")

            if "mumbai_imported_usernames" in st.session_state and not extracted_usernames_list:
                extracted_usernames_list = st.session_state["mumbai_imported_usernames"]

        # Pre-check Summary Telemetry
        col_pre_h1, col_pre_h2 = st.columns([3.5, 1.5])
        with col_pre_h1:
            st.markdown("##### 🔎 Target Pre-Check Summary")
        with col_pre_h2:
            if st.button("🔄 Refresh Discovery Stats", help="Re-scan checkpoints for newly extracted comments", disabled=m_is_running):
                get_cached_available_checkpoint_accounts.clear()
                get_cached_commenters_summary_counts.clear()
                st.rerun()

        if input_mode == "👥 Scraped Source Profiles (Commenters & Repliers)":
            active_srcs = tuple(selected_source_accounts or [])
            pre_summary = get_cached_commenters_summary_counts(active_srcs, recheck=False)
            pre_col1, pre_col2, pre_col3, pre_col4 = st.columns(4)
            pre_col1.metric("💬 Total Comments & Replies", f"{pre_summary['discovered_count']:,}")
            pre_col2.metric("👥 Unique Commenters & Repliers", f"{pre_summary['unique_count']:,}")
            pre_col3.metric("⏭️ Already Checked in DB", f"{pre_summary['already_checked_count']:,}")
            pre_col4.metric("🎯 Remaining to Check", f"{pre_summary['remaining_count']:,}")
        else:
            pre_col1, pre_col2, pre_col3, pre_col4 = st.columns(4)
            pre_col1.metric("💬 Total Discovered", f"{len(extracted_usernames_list):,}")
            pre_col2.metric("👥 Unique Usernames", f"{len(extracted_usernames_list):,}")
            db_inst = MumbaiCheckDB()
            completed_set = db_inst.get_completed_usernames()
            already_cnt = sum(1 for u in extracted_usernames_list if u.strip().lstrip("@").lower() in completed_set)
            pre_col3.metric("⏭️ Already Checked in DB", f"{already_cnt:,}")
            pre_col4.metric("🎯 Remaining to Check", f"{max(0, len(extracted_usernames_list) - already_cnt):,}")

        st.markdown("##### ⚙️ Execution & Rate Controls")
        col_m_opt1, col_m_opt2, col_m_opt3 = st.columns(3)
        with col_m_opt1:
            m_opt_recheck = st.checkbox(
                "Force Re-check (--recheck)",
                value=False,
                disabled=m_is_running,
                help="Re-audit previously completed accounts rather than skipping them.",
            )
        with col_m_opt2:
            m_opt_limit = st.number_input(
                "Batch Limit (0 = check all remaining)",
                min_value=0,
                max_value=10000,
                value=0,
                step=25,
                disabled=m_is_running,
                help="Maximum number of accounts to check in this batch. Set 0 for unlimited.",
            )
        with col_m_opt3:
            m_adv_delay = st.slider(
                "Request Delay (seconds)",
                min_value=1.0,
                max_value=5.0,
                value=2.0,
                step=0.5,
                help="Pause between profile GraphQL calls to maintain healthy rate limits.",
                disabled=m_is_running,
            )

        col_dest1, col_dest2, col_dest3 = st.columns(3)
        with col_dest1:
            m_opt_sheets = st.checkbox("Sync to Google Spreadsheet ('Mumbai Check' Tab)", value=True, disabled=m_is_running)
        with col_dest2:
            m_opt_excel = st.checkbox("Update Local Excel Workbook ('Mumbai Check' Tab)", value=True, disabled=m_is_running)
        with col_dest3:
            m_adv_sheet_id = st.text_input(
                "Master Spreadsheet ID",
                value=DEFAULT_SPREADSHEET_ID,
                disabled=m_is_running,
            )

        # Action Buttons
        col_m_act1, col_m_act2, col_m_act3 = st.columns([2, 1.5, 3.5])
        with col_m_act1:
            start_mumbai_clicked = st.button(
                "🔍 Check Mumbai Location",
                type="primary",
                use_container_width=True,
                disabled=m_is_running,
            )
        with col_m_act2:
            stop_mumbai_clicked = st.button(
                "⏹ Stop / Cancel Batch",
                type="secondary",
                use_container_width=True,
                disabled=not m_is_running,
            )

        if start_mumbai_clicked:
            if input_mode.startswith("👥") and not selected_source_accounts:
                st.error("Please select at least one scraped source account to audit commenters from.")
            elif not input_mode.startswith("👥") and not extracted_usernames_list:
                st.error("Please supply at least one Instagram username to check.")
            else:
                ok, msg = mumbai_manager.start_batch_check(
                    source_accounts=selected_source_accounts if input_mode.startswith("👥") else None,
                    custom_usernames=extracted_usernames_list if not input_mode.startswith("👥") else None,
                    recheck=m_opt_recheck,
                    limit=int(m_opt_limit),
                    delay=m_adv_delay,
                    sync_excel=m_opt_excel,
                    sync_sheets=m_opt_sheets,
                    spreadsheet_id=m_adv_sheet_id,
                )
                if ok:
                    get_cached_commenters_summary_counts.clear()
                    st.toast("Started Mumbai location check batch on commenters!", icon="📍")
                    st.rerun()
                else:
                    st.error(msg)

        if stop_mumbai_clicked:
            mumbai_manager.cancel_job()
            st.toast("Cancellation requested. Stopping after active profile...", icon="⏹")
            st.rerun()

    # Progress & Metrics Section (Isolated Fragment)
    if m_is_running:
        _render_mumbai_progress_active()
    elif m_snapshot.get("total_to_process", 0) > 0 or m_snapshot.get("processed_count", 0) > 0:
        _render_mumbai_telemetry_ui(m_snapshot)

    # Results Table Section
    st.markdown("---")
    st.markdown("### 📋 Mumbai Audit Database & Master Results")

    db_instance = MumbaiCheckDB()
    total_in_db = db_instance.get_checks_count()

    if total_in_db == 0:
        st.info("No accounts checked yet. Select scraped sources above and click 'Check Mumbai Location' to run your first batch.")
    else:
        # Quick Classification Stats Badges
        cat_stats = db_instance.get_classification_stats()
        stat_cols = st.columns(5)
        stat_cols[0].caption(f"✓ **Likely Mumbai**: {cat_stats.get('Likely Mumbai', 0):,}")
        stat_cols[1].caption(f"⚡ **Serves Mumbai**: {cat_stats.get('Serves Mumbai', 0):,}")
        stat_cols[2].caption(f"? **Possibly Mumbai**: {cat_stats.get('Possibly Mumbai', 0):,}")
        stat_cols[3].caption(f"✕ **Outside Mumbai**: {cat_stats.get('Likely outside Mumbai', 0):,}")
        stat_cols[4].caption(f"○ **Not Enough Evid.**: {cat_stats.get('Not enough evidence', 0):,}")

        col_fil1, col_fil2, col_fil3, col_fil4 = st.columns([1.8, 2.0, 1.0, 1.2])
        with col_fil1:
            cat_filter = st.selectbox(
                "Filter Category",
                ["All Categories", "Likely Mumbai", "Serves Mumbai", "Possibly Mumbai", "Likely outside Mumbai", "Not enough evidence"],
            )
        with col_fil2:
            search_query = st.text_input("Search Results", placeholder="Filter by username, brand, location, keyword...")
        with col_fil3:
            page_size = st.selectbox("Page Size", [25, 50, 100], index=0)
        with col_fil4:
            view_mode = st.selectbox("View Mode", ["Table (Grid)", "Cards"], index=0)

        # Count filtered records directly in SQLite
        filtered_count = db_instance.get_checks_count(classification=cat_filter, search=search_query)
        total_pages = max(1, (filtered_count + page_size - 1) // page_size)

        # Session state for pagination
        if "mumbai_page" not in st.session_state:
            st.session_state["mumbai_page"] = 1

        if st.session_state["mumbai_page"] > total_pages:
            st.session_state["mumbai_page"] = total_pages
        if st.session_state["mumbai_page"] < 1:
            st.session_state["mumbai_page"] = 1

        curr_page = st.session_state["mumbai_page"]

        # Pagination Navigation Header
        p_c1, p_c2, p_c3, p_c4 = st.columns([1.2, 1.6, 2.8, 1.2])
        with p_c1:
            if st.button("⬅️ Previous", disabled=(curr_page <= 1), use_container_width=True):
                st.session_state["mumbai_page"] = max(1, curr_page - 1)
                st.rerun()
        with p_c2:
            st.markdown(f"<div style='text-align: center; padding-top: 6px; font-weight: 600;'>Page {curr_page} of {total_pages}</div>", unsafe_allow_html=True)
        with p_c3:
            start_num = min((curr_page - 1) * page_size + 1, filtered_count) if filtered_count > 0 else 0
            end_num = min(curr_page * page_size, filtered_count)
            st.caption(f"Showing **{start_num}–{end_num}** of **{filtered_count:,}** matching accounts (Database Total: {total_in_db:,})")
        with p_c4:
            if st.button("Next ➡️", disabled=(curr_page >= total_pages), use_container_width=True):
                st.session_state["mumbai_page"] = min(total_pages, curr_page + 1)
                st.rerun()

        # Fetch ONLY the current page slice from SQLite
        offset = (curr_page - 1) * page_size
        page_records = db_instance.get_checks_page(
            limit=page_size,
            offset=offset,
            classification=cat_filter,
            search=search_query,
        )

        if not page_records:
            st.info("No records matching the selected filter or search query.")
        elif view_mode == "Table (Grid)":
            df_rows = []
            for r in page_records:
                u_name = r.get("checked_username") or r.get("username") or ""
                df_rows.append({
                    "Username": f"@{u_name}",
                    "Display Name": r.get("display_name") or "—",
                    "Source Account(s)": r.get("source_accounts") or "—",
                    "Stated Location": r.get("stated_location") or "—",
                    "Classification": r.get("classification") or "—",
                    "Confidence": r.get("confidence") or "Low",
                    "Evidence Summary": r.get("evidence_summary") or "",
                    "Checked At": r.get("checked_at") or "",
                })
            df = pd.DataFrame(df_rows)
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            # Custom styled cards (only renders current page rows, e.g. 25!)
            table_cols = st.columns([1.5, 1.4, 1.4, 1.4, 3.1, 1.5, 1.0])
            table_cols[0].markdown("**Username**")
            table_cols[1].markdown("**Display Name**")
            table_cols[2].markdown("**Source Account(s)**")
            table_cols[3].markdown("**Stated Location**")
            table_cols[4].markdown("**Evidence Summary**")
            table_cols[5].markdown("**Classification**")
            table_cols[6].markdown("**Confidence**")
            st.markdown("---")

            for item in page_records:
                r1, r2, r3, r4, r5, r6, r7 = st.columns([1.5, 1.4, 1.4, 1.4, 3.1, 1.5, 1.0])
                u_name = item.get("checked_username") or item.get("username") or ""
                p_url = item.get("profile_url", f"https://www.instagram.com/{u_name}/")
                r1.markdown(f"[@{u_name}]({p_url})")
                r2.markdown(item.get("display_name") or "—")
                r3.markdown(f"`{item.get('source_accounts') or '—'}`")
                r4.markdown(f"`{item.get('stated_location') or '—'}`")
                r5.caption(item.get("evidence_summary") or "")

                cls_name = item.get("classification", "")
                if cls_name == "Likely Mumbai":
                    r6.markdown('<div class="status-badge badge-likely-mumbai">✓ LIKELY MUMBAI</div>', unsafe_allow_html=True)
                elif cls_name == "Serves Mumbai":
                    r6.markdown('<div class="status-badge badge-serves-mumbai">⚡ SERVES MUMBAI</div>', unsafe_allow_html=True)
                elif cls_name == "Possibly Mumbai":
                    r6.markdown('<div class="status-badge badge-possibly-mumbai">? POSSIBLY MUMBAI</div>', unsafe_allow_html=True)
                elif cls_name == "Likely outside Mumbai":
                    r6.markdown('<div class="status-badge badge-outside-mumbai">✕ OUTSIDE MUMBAI</div>', unsafe_allow_html=True)
                else:
                    r6.markdown('<div class="status-badge badge-not-enough">○ NOT ENOUGH EVID.</div>', unsafe_allow_html=True)

                r7.markdown(f"**{item.get('confidence', 'Low')}**")

        st.markdown("---")
        # Direct Links and Force Re-sync Action
        col_out1, col_out2 = st.columns(2)
        with col_out1:
            st.markdown("##### 🌐 Online Master Google Spreadsheet")
            sheet_url = f"https://docs.google.com/spreadsheets/d/{DEFAULT_SPREADSHEET_ID}/edit"
            st.markdown(f"- Worksheet Tab: **`Mumbai Check`**\n- Master Sheet ID: `{DEFAULT_SPREADSHEET_ID}`")
            st.link_button("🔗 Open 'Mumbai Check' in Google Sheets", sheet_url, type="primary")

        with col_out2:
            st.markdown("##### 💻 Local Excel Workbook")
            st.markdown(f"- File Path: `{DEFAULT_EXCEL_PATH}`\n- Worksheet Tab: **`Mumbai Check`**")
            if st.button("🔄 Force Re-sync 'Mumbai Check' to Google Sheets & Excel"):
                with st.spinner("Synchronizing all checked accounts to Google Sheets & Excel..."):
                    try:
                        all_db_records = db_instance.get_all_checks()
                        ex_cnt = MumbaiCheckSync.sync_to_excel(all_db_records)
                        sh_res = MumbaiCheckSync.sync_to_google_sheets(all_db_records)
                        st.success(f"Successfully synced {ex_cnt} accounts to 'Mumbai Check' in both Google Sheets and Excel!")
                    except Exception as e:
                        st.error(f"Sync failed: {e}")


# =========================================================================
# VIEW: MEATNDOOR LEAD OUTREACH MANAGER
# =========================================================================
elif selected_view == "🍗 MeatNDoor Lead Outreach":
    col_out_h1, col_out_h2 = st.columns([3, 1])
    with col_out_h1:
        st.markdown("## 🍗 MeatNDoor Lead Outreach Manager")
        st.caption("Campaign eligibility, hyper-personalized message drafting, review queue & DNC enforcement.")
    with col_out_h2:
        out_snap = outreach_manager.get_snapshot()
        if out_snap["is_running"]:
            st.markdown('<div class="status-badge badge-running">● INGESTING LEADS</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-badge badge-idle">○ OUTREACH READY</div>', unsafe_allow_html=True)

    # Ingestion status banner if running
    out_snap = outreach_manager.get_snapshot()
    if out_snap["is_running"]:
        st.info(f"⏳ {out_snap['status_text']}")
        st.caption(f"Elapsed: {out_snap['elapsed_seconds']}s")

    # Metrics Summary Row
    summary = outreach_manager.get_lead_counts_summary()
    st.markdown("##### 📊 Outreach Campaign & Queue Overview")
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    with k1:
        st.metric("Total Leads", f"{summary['total_leads']:,}")
    with k2:
        st.metric("Mumbai Eligible", f"{summary['eligible_mumbai']:,}")
    with k3:
        st.metric("Outside Suppressed", f"{summary['suppressed_outside']:,}")
    with k4:
        st.metric("DNC Suppressed", f"{summary['suppressed_dnc']:,}")
    with k5:
        st.metric("Pending Review", f"{summary['pending_review']:,}")
    with k6:
        st.metric("Approved", f"{summary['approved']:,}")
    with k7:
        st.metric("Contacted", f"{summary['contacted']:,}")

    st.markdown("---")

    # Ingestion & Sync Controls
    ctrl_col1, ctrl_col2 = st.columns(2)
    with ctrl_col1:
        with st.expander("📥 Ingest & Refresh Leads from Checkpoints", expanded=False):
            st.caption(
                "Reads commenters and reply authors from existing brand checkpoints, joins with cached Mumbai "
                "evaluations, applies location suppression and DNC rules, and assigns message drafts."
            )
            avail_accs = get_cached_available_checkpoint_accounts()
            avail_names = [a["source_username"] for a in avail_accs] if avail_accs else []
            selected_sources = st.multiselect(
                "Source Brands (leave empty to scan all checkpoints)",
                options=avail_names,
                default=[],
                key="outreach_sources_multiselect",
            )
            if st.button("🚀 Run Lead Ingestion", disabled=out_snap["is_running"], key="btn_run_ingest"):
                ok, msg = outreach_manager.start_ingestion_job(source_accounts=selected_sources if selected_sources else None)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.warning(msg)

    with ctrl_col2:
        with st.expander("🔄 Dual-Destination Reporting Sync", expanded=False):
            st.caption(
                "Synchronizes the complete outreach review queue to the master Google Spreadsheet "
                f"(`{DEFAULT_SPREADSHEET_ID}`) and local Excel mirror (`{DEFAULT_EXCEL_PATH}`)."
            )
            if st.button("📤 Sync Queue to Google Sheets & Excel", key="btn_sync_outreach_both"):
                with st.spinner("Syncing Outreach Manager worksheet..."):
                    try:
                        ex_count = outreach_manager.sync_to_excel()
                        st.success(f"✓ Excel synced: {ex_count} leads saved to '{DEFAULT_EXCEL_PATH}' -> 'Outreach Manager'")
                    except Exception as e:
                        st.error(f"Excel sync error: {e}")

                    try:
                        sh_res = outreach_manager.sync_to_google_sheets()
                        st.success(f"✓ Google Sheets synced: [{sh_res['sheet_title']}]({sh_res['spreadsheet_url']}) ({sh_res['rows_synced']} rows)")
                    except Exception as e:
                        st.warning(f"Google Sheets sync note: {e}")

    # Search & Filter Bar
    st.markdown("##### 🔍 Lead Queue Filters & Search")
    f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns([2.5, 2.0, 1.8, 1.5, 1.2])
    with f_col1:
        search_kw = st.text_input("Search Username / Location", placeholder="e.g. Bandra, rahul, licious", key="outreach_search_kw")
    with f_col2:
        elig_opts = [
            "All Eligibilities",
            ELIGIBILITY_MUMBAI_TARGET,
            ELIGIBILITY_BUSINESS_REVIEW,
            ELIGIBILITY_LOW_CONF_REVIEW,
            ELIGIBILITY_PENDING_LOCATION,
            ELIGIBILITY_SUPPRESSED_OUTSIDE,
            ELIGIBILITY_SUPPRESSED_DNC,
        ]
        sel_elig = st.selectbox("Campaign Eligibility", elig_opts, index=0, key="outreach_sel_elig")
    with f_col3:
        status_opts = ["All Statuses"] + ALL_REVIEW_STATUSES
        sel_status = st.selectbox("Review Status", status_opts, index=0, key="outreach_sel_status")
    with f_col4:
        sort_opts = {
            "Newest / Last Action": "last_action_at DESC",
            "Most Comments": "total_comments DESC",
            "Username (A-Z)": "username ASC",
            "Oldest Action": "last_action_at ASC",
        }
        sel_sort_label = st.selectbox("Sort By", list(sort_opts.keys()), index=0, key="outreach_sel_sort")
        sel_sort = sort_opts[sel_sort_label]
    with f_col5:
        page_size = st.selectbox("Page Size", [10, 25, 50, 100], index=1, key="outreach_page_size")

    # Pagination state
    if "outreach_page_num" not in st.session_state:
        st.session_state["outreach_page_num"] = 1

    # Query leads
    filter_elig = None if sel_elig == "All Eligibilities" else sel_elig
    filter_status = None if sel_status == "All Statuses" else sel_status
    offset = (st.session_state["outreach_page_num"] - 1) * page_size

    leads, total_filtered = outreach_manager.get_leads_paginated(
        limit=page_size,
        offset=offset,
        search=search_kw,
        eligibility=filter_elig,
        review_status=filter_status,
        sort_by=sel_sort,
    )

    total_pages = max(1, math.ceil(total_filtered / page_size))
    if st.session_state["outreach_page_num"] > total_pages:
        st.session_state["outreach_page_num"] = total_pages

    # Page navigation header
    p_nav1, p_nav2, p_nav3 = st.columns([2, 4, 2])
    with p_nav1:
        if st.button("⬅️ Previous Page", disabled=(st.session_state["outreach_page_num"] <= 1), key="outreach_btn_prev"):
            st.session_state["outreach_page_num"] -= 1
            st.rerun()
    with p_nav2:
        start_row = offset + 1 if total_filtered > 0 else 0
        end_row = min(offset + page_size, total_filtered)
        st.markdown(
            f"<div style='text-align: center; color: #9aa2ab; padding-top: 6px;'>"
            f"Showing <b>{start_row} - {end_row}</b> of <b>{total_filtered:,}</b> leads (Page {st.session_state['outreach_page_num']} of {total_pages})</div>",
            unsafe_allow_html=True,
        )
    with p_nav3:
        if st.button("Next Page ➡️", disabled=(st.session_state["outreach_page_num"] >= total_pages), key="outreach_btn_next"):
            st.session_state["outreach_page_num"] += 1
            st.rerun()

    # Lead Cards
    if not leads:
        st.info("No leads match the selected criteria. Try adjusting the filters or running Lead Ingestion.")
    else:
        for lead in leads:
            u = lead["username"]
            norm_u = lead["normalized_username"]
            p_url = lead["profile_url"]
            dm_url = lead["direct_message_url"]
            elig = lead["campaign_eligibility"]
            rev_stat = lead["review_status"]
            cls_name = lead["mumbai_classification"]
            conf = lead["mumbai_confidence"]
            loc = lead["stated_location"] or "—"
            evid = lead["evidence_summary"] or "—"
            tot_comments = lead["total_comments"]
            comment_types = lead["comment_types"] or "comment"
            brands = lead["source_brands"]
            if brands.startswith("["):
                try:
                    brands = ", ".join(json.loads(brands))
                except Exception:
                    pass

            # Eligibility badge styling
            badge_class = "badge-pending-location"
            if elig == ELIGIBILITY_MUMBAI_TARGET:
                badge_class = "badge-eligible-mumbai"
            elif elig == ELIGIBILITY_BUSINESS_REVIEW:
                badge_class = "badge-business-review"
            elif elig == ELIGIBILITY_LOW_CONF_REVIEW:
                badge_class = "badge-low-conf-review"
            elif elig == ELIGIBILITY_SUPPRESSED_OUTSIDE:
                badge_class = "badge-suppressed-outside"
            elif elig == ELIGIBILITY_SUPPRESSED_DNC:
                badge_class = "badge-suppressed-dnc"

            status_badge_class = "badge-review-pending"
            if rev_stat == STATUS_APPROVED:
                status_badge_class = "badge-review-approved"
            elif rev_stat == STATUS_CONTACTED:
                status_badge_class = "badge-review-contacted"
            elif rev_stat in (STATUS_REPLIED, STATUS_INTERESTED):
                status_badge_class = "badge-review-replied"
            elif rev_stat == STATUS_DNC:
                status_badge_class = "badge-suppressed-dnc"

            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #1a1d23; border: 1px solid #2d333b; border-radius: 8px; padding: 14px; margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <div>
                                <span style="font-size: 1.15rem; font-weight: 700; color: #58a6ff;">
                                    <a href="{p_url}" target="_blank" style="color: #58a6ff; text-decoration: none;">@{u}</a>
                                </span>
                                &nbsp;&nbsp;
                                <a href="{dm_url}" target="_blank" style="font-size: 0.85rem; color: #22d3ee; text-decoration: none; border: 1px solid #0891b2; padding: 2px 8px; border-radius: 4px;">
                                    💬 Open ig.me Chat
                                </a>
                            </div>
                            <div>
                                <span class="status-badge {badge_class}" style="margin-right: 6px;">{elig.upper()}</span>
                                <span class="status-badge {status_badge_class}">{rev_stat.upper()}</span>
                            </div>
                        </div>
                        <div style="display: grid; grid-template-columns: 1fr 1fr 1.5fr; gap: 8px; font-size: 0.85rem; color: #9aa2ab; margin-bottom: 8px;">
                            <div><b>Source Brand(s):</b> <code style="color: #c9d1d9;">{brands}</code></div>
                            <div><b>Engagement:</b> {tot_comments} record(s) ({comment_types})</div>
                            <div><b>Location Audit:</b> <span style="color: #e6edf3;">{cls_name}</span> ({conf} confidence)</div>
                        </div>
                        <div style="font-size: 0.82rem; color: #8b949e; margin-bottom: 6px;">
                            <b>Stated Location:</b> {loc} &nbsp;|&nbsp; <b>Evidence:</b> {evid}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # Editable draft text area
                draft_key = f"draft_input_{norm_u}"
                edited_draft = st.text_area(
                    f"Outreach Draft for @{u}",
                    value=lead["drafted_message"],
                    key=draft_key,
                    height=75,
                    label_visibility="collapsed",
                )

                # Action buttons
                b_col1, b_col2, b_col3, b_col4, b_col5, b_col6, b_col7, b_col8 = st.columns(8)
                with b_col1:
                    if st.button("✅ Approve", key=f"app_{norm_u}", help="Approve draft for outreach (does NOT send message)"):
                        try:
                            if edited_draft != lead["drafted_message"]:
                                outreach_manager.edit_lead_draft(norm_u, edited_draft)
                            outreach_manager.update_lead_status(norm_u, STATUS_APPROVED, notes="Approved by operator")
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))
                with b_col2:
                    if st.button("💾 Save Draft", key=f"sav_{norm_u}", help="Save custom draft edits without approving"):
                        outreach_manager.edit_lead_draft(norm_u, edited_draft)
                        st.toast(f"Saved draft edits for @{u}")
                        st.rerun()
                with b_col3:
                    if st.button("❌ Reject", key=f"rej_{norm_u}", help="Reject draft"):
                        outreach_manager.update_lead_status(norm_u, STATUS_REJECTED, notes="Rejected by operator")
                        st.rerun()
                with b_col4:
                    if st.button("📩 Contacted", key=f"cnt_{norm_u}", help="Mark lead as contacted"):
                        outreach_manager.update_lead_status(norm_u, STATUS_CONTACTED, notes="Marked contacted by operator")
                        st.rerun()
                with b_col5:
                    if st.button("💬 Replied", key=f"rep_{norm_u}", help="Mark lead as replied"):
                        outreach_manager.update_lead_status(norm_u, STATUS_REPLIED, notes="Lead replied")
                        st.rerun()
                with b_col6:
                    if st.button("🌟 Interested", key=f"int_{norm_u}", help="Mark lead as interested"):
                        outreach_manager.update_lead_status(norm_u, STATUS_INTERESTED, notes="Lead showed interest")
                        st.rerun()
                with b_col7:
                    if st.button("🚫 Not Int.", key=f"not_{norm_u}", help="Mark lead as not interested"):
                        outreach_manager.update_lead_status(norm_u, STATUS_NOT_INTERESTED, notes="Lead not interested")
                        st.rerun()
                with b_col8:
                    if lead["is_dnc"]:
                        if st.button("🔓 Un-DNC", key=f"undnc_{norm_u}", help="Remove from Do Not Contact list"):
                            outreach_manager.remove_from_dnc(norm_u)
                            st.rerun()
                    else:
                        if st.button("⛔ DNC", key=f"dnc_{norm_u}", help="Add to Do Not Contact list"):
                            outreach_manager.add_to_dnc(norm_u, reason="Operator marked Do Not Contact")
                            st.rerun()

                st.markdown("<div style='margin-bottom: 18px;'></div>", unsafe_allow_html=True)

    # Operator Audit Log
    st.markdown("---")
    with st.expander("📜 Operator Decision & Compliance Audit Log (Recent 50 Actions)", expanded=False):
        audit_rows = outreach_manager.get_audit_log(limit=50)
        if not audit_rows:
            st.caption("No operator actions recorded yet.")
        else:
            df_audit = pd.DataFrame(audit_rows)
            st.dataframe(df_audit[["id", "normalized_username", "action", "previous_status", "new_status", "notes", "operator", "created_at"]], use_container_width=True)


# =========================================================================
# VIEW 3: HISTORY & CHECKPOINTS
# =========================================================================

elif selected_view == "📚 History & Checkpoints":
    st.markdown("### 📚 Historical Checkpoint Accounts")
    col_hist_h1, col_hist_h2 = st.columns([4, 1])
    with col_hist_h1:
        st.caption("Inspect and manage previously extracted profiles stored in SQLite checkpoints.")
    with col_hist_h2:
        if st.button("🔄 Refresh History"):
            get_cached_checkpoint_history.clear()
            st.rerun()

    history_list = get_cached_checkpoint_history()
    if not history_list:
        st.info("No checkpoint databases found in `.checkpoints/`.")
    else:
        st.markdown(f"Found **{len(history_list)}** previously processed accounts:")

        df_hist = pd.DataFrame([
            {
                "Account": f"@{item['username']}",
                "Posts Processed": f"{item['posts_completed']} / {item['posts_total']}",
                "Comments Extracted": f"{item['comments_count']:,}",
                "Replies Extracted": f"{item['replies_count']:,}",
                "Last Updated": item['last_modified'],
            }
            for item in history_list
        ])
        st.dataframe(df_hist, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("##### 🔄 Re-synchronize Account to Google Sheets & Excel")
        col_s1, col_s2 = st.columns([3, 1.2])
        with col_s1:
            sel_user = st.selectbox(
                "Select Profile Checkpoint",
                [item['username'] for item in history_list],
                format_func=lambda u: f"@{u}",
            )
        with col_s2:
            st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("🔄 Sync Selected Account", use_container_width=True):
                with st.spinner(f"Syncing @{sel_user}..."):
                    try:
                        sync_out = resync_existing_account(sel_user)
                        if sync_out.get("sheets_success"):
                            st.success(f"Synced @{sel_user} to Google Sheets & Excel!")
                        else:
                            st.warning(f"Sync result: {sync_out.get('sheets_message')}")
                    except Exception as err:
                        st.error(f"Sync error: {err}")


# =========================================================================
# VIEW 4: GOOGLE SHEETS & AUTH HEALTH
# =========================================================================
elif selected_view == "🔑 Google Sheets Health":
    st.markdown("### 🔑 Google Sheets API Health & Diagnostics")
    st.caption("Verify that your local environment is authorized to update the master Google Spreadsheet.")

    test_col1, test_col2 = st.columns([2, 1])
    with test_col1:
        st.markdown(f"**Target Spreadsheet ID**: `{DEFAULT_SPREADSHEET_ID}`")
        st.markdown(
            f"[Open Spreadsheet in Browser](https://docs.google.com/spreadsheets/d/{DEFAULT_SPREADSHEET_ID}/edit)"
        )
    with test_col2:
        test_btn = st.button("🔍 Test Connection Now", use_container_width=True)

    if test_btn:
        with st.spinner("Connecting to Google Sheets API..."):
            ok, title_or_err, tabs = test_google_sheets_connection()
            if ok:
                st.success(f"**Connection Successful!**\n\n- Spreadsheet Title: **{title_or_err}**", icon="✅")
                st.markdown("##### Existing Worksheet Tabs in Master Sheet:")
                st.write(tabs)
            else:
                st.error(f"**Connection Test Failed:**\n\n{title_or_err}", icon="❌")

    st.markdown("---")
    st.markdown("##### 📋 Service Account Configuration Guide")
    st.markdown(
        """
1. **Service Account Key File**:
   Ensure `service_account.json` is located in `/home/nars/projects/api-anything-api/api-anything/service_account.json`.
2. **Spreadsheet Permissions**:
   The Google Spreadsheet (`1sLCfz9Eho4OWjsDTILIVKNRtwVYbCoibFK6z6un7c_Y`) must be shared with the Service Account email address with **Editor** permissions.
3. **Automatic Dual Sync**:
   Every time an extraction is executed, the tool automatically updates the tab for that username and keeps all other tabs intact.
        """
    )
