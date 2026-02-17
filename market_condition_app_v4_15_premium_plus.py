# market_condition_app_v3.py
# TKM Appraisals Market Adjustment Tool - Step-by-step workflow
#
# STEP-BY-STEP WORKFLOW:
# 1. Subject Property - Enter address and effective date
# 2. Upload Data - Upload CSV with market data
# 3. Data Diagnostics - Review and exclude outliers
# 4. Select Comparables - Choose comps for analysis
# 5. Adjustment Report - View chart and download results
#
# REQUIRED CSV HEADERS:
#   Address, Zip, Pending Date, Sold Price

import io
import os
import zipfile
import json
from datetime import date, datetime
from typing import Tuple, List

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import streamlit as st
import plotly.express as px

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle

# ----------------------------
# Report History Management
# ----------------------------
HISTORY_DIR = os.path.join(os.path.expanduser("~"), ".marketadjuster")
HISTORY_FILE = os.path.join(HISTORY_DIR, "history.json")

def _ensure_history_dir():
    os.makedirs(HISTORY_DIR, exist_ok=True)

def load_history() -> list:
    """Load report history from JSON file."""
    _ensure_history_dir()
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []
    return []

def save_history(history: list):
    """Save report history to JSON file."""
    _ensure_history_dir()
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2, default=str)

def save_report_to_history(session_state: dict):
    """Save current report state to history."""
    history = load_history()

    # Serialize the data we need to restore the report
    report_data = {
        "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
        "saved_at": datetime.now().isoformat(),
        "subject_address": session_state.get("subject_address", ""),
        "eff_date": str(session_state.get("eff_date", date.today())),
        "settings": session_state.get("settings", {}),
        "excluded_rowids": list(session_state.get("excluded_rowids", set())),
        "selected_comps": session_state.get("selected_comps", []),
    }

    # Save uploaded data as CSV string
    if session_state.get("uploaded_data") is not None:
        try:
            df = session_state["uploaded_data"]
            report_data["uploaded_data_csv"] = df.to_csv(index=False)
        except Exception:
            report_data["uploaded_data_csv"] = None
    else:
        report_data["uploaded_data_csv"] = None

    # Check if we already have a report for this address+date, update it
    existing_idx = None
    for i, r in enumerate(history):
        if (r.get("subject_address") == report_data["subject_address"]
            and r.get("eff_date") == report_data["eff_date"]):
            existing_idx = i
            break

    if existing_idx is not None:
        history[existing_idx] = report_data
    else:
        history.insert(0, report_data)  # newest first

    # Keep max 50 reports
    history = history[:50]
    save_history(history)

def load_report_from_history(report_data: dict, session_state):
    """Restore a saved report into session state."""
    session_state["subject_address"] = report_data.get("subject_address", "")
    session_state["eff_date"] = date.fromisoformat(report_data.get("eff_date", str(date.today())))
    session_state["settings"] = report_data.get("settings", {})
    session_state["excluded_rowids"] = set(report_data.get("excluded_rowids", []))
    session_state["selected_comps"] = report_data.get("selected_comps", [])

    # Restore uploaded data
    csv_str = report_data.get("uploaded_data_csv")
    if csv_str:
        try:
            df_loaded = pd.read_csv(io.StringIO(csv_str))
            # Re-convert date columns that were stored as strings
            if "ContractDate" in df_loaded.columns:
                df_loaded["ContractDate"] = pd.to_datetime(df_loaded["ContractDate"]).dt.date
            if "SoldPrice" in df_loaded.columns:
                df_loaded["SoldPrice"] = pd.to_numeric(df_loaded["SoldPrice"], errors="coerce")
            session_state["uploaded_data"] = df_loaded
        except Exception:
            session_state["uploaded_data"] = None
    else:
        session_state["uploaded_data"] = None

    # Jump to step 5 (report) if we have data, otherwise step 1
    if session_state["uploaded_data"] is not None and session_state["selected_comps"]:
        session_state["step"] = 5
    elif session_state["uploaded_data"] is not None:
        session_state["step"] = 3
    else:
        session_state["step"] = 1

def delete_report_from_history(report_id: str):
    """Delete a report from history by ID."""
    history = load_history()
    history = [r for r in history if r.get("id") != report_id]
    save_history(history)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

from html import escape as _escape

def vq_stat_card(label: str, value: str, delta: str | None = None) -> None:
    """Render a compact, readable stat card (more reliable than st.metric across browsers/themes)."""
    label_h = _escape(str(label))
    value_s = str(value)
    value_h = _escape(value_s)
    delta_html = f'<div class="vq-stat-delta">{_escape(str(delta))}</div>' if delta else ""
    value_class = "vq-stat-value vq-stat-value-sm" if len(value_s) >= 18 else "vq-stat-value"
    st.markdown(
        f'''
        <div class="vq-stat">
          <div class="vq-stat-label">{label_h}</div>
          <div class="{value_class}">{value_h}</div>
          {delta_html}
        </div>
        ''',
        unsafe_allow_html=True,
    )


# Modern color palette
THEME_COLORS = {
    'primary': '#60A5FA',      # Accent Blue
    'primary_light': '#93C5FD',
    'success': '#34C759',
    'danger': '#FF3B30',
    'warning': '#FF9500',
    'secondary': '#6B7280',
    'background': '#F9FAFB',
    'surface': '#FFFFFF',
    'text': '#111827',
    'text_secondary': '#6B7280',
    'border': '#E5E7EB',
    'upward': '#34C759',
    'downward': '#FF3B30',
    'stable': '#6B7280',
}

APP_NAME = "MarketAdjuster"
REQUIRED_COLS = ["Address", "Zip", "Pending Date", "Sold Price"]

# ----------------------------
# Custom CSS for modern UI
# ----------------------------

def inject_modern_css():
    """Inject a clean, modern 'Liquid Glass' UI theme (Streamlit CSS override)."""
    st.markdown(r"""
    <style>
:root{
  --bg:#f5f5f7;
  --surface: rgba(255,255,255,.92);
  --surface2: rgba(255,255,255,.98);
  --border: rgba(17,24,39,.14);
  --border2: rgba(17,24,39,.20);
  --ink:#111827;
  --muted:#4b5563;
  --shadow: 0 10px 30px rgba(0,0,0,.08);
  --shadow2: 0 18px 50px rgba(0,0,0,.10);
  --r-xl:18px;
  --r-lg:14px;
  --r-md:12px;
  --accent:#60A5FA;
  --accentSoft: rgba(37,99,235,.14);
}

/* Page */
.stApp{
  background: var(--bg) !important;
  color: var(--ink) !important;
}
html, body, [class*="css"]{
  -webkit-font-smoothing: antialiased;
  font-smoothing: antialiased;
}
header, #MainMenu, footer { visibility:hidden; height:0 !important; }
section.main > div { padding-top: 1.2rem !important; }
.block-container { max-width: 1120px; padding-left: 1.2rem; padding-right: 1.2rem; }

/* Wide mode override for the report step */
.report-wide-mode .block-container { max-width: 1520px !important; }

/* Modern frosted-glass hero card */
.vq-hero{
  background: linear-gradient(135deg, rgba(96,165,250,.08) 0%, rgba(255,255,255,.92) 100%);
  border: 1px solid rgba(96,165,250,.18);
  border-radius: var(--r-xl);
  padding: 20px 24px;
  margin-bottom: 16px;
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  box-shadow: 0 4px 24px rgba(96,165,250,.06);
}
.vq-hero-title{
  font-size: 15px; font-weight: 700; color: var(--ink);
  margin-bottom: 4px;
}
.vq-hero-sub{
  font-size: 13px; color: var(--muted); line-height: 1.45;
}

/* Modern report table via HTML */
.vq-table-wrap{
  border-radius: var(--r-xl);
  overflow: hidden;
  border: 1px solid var(--border);
  box-shadow: var(--shadow);
  background: var(--surface);
}
.vq-table{
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
  color: var(--ink);
}
.vq-table thead th{
  background: linear-gradient(180deg, #F0F4FF 0%, #E8EEF7 100%);
  color: var(--ink);
  font-weight: 700;
  font-size: 11px;
  letter-spacing: .4px;
  text-transform: uppercase;
  padding: 12px 14px;
  text-align: left;
  border-bottom: 2px solid rgba(96,165,250,.20);
  white-space: nowrap;
}
.vq-table tbody td{
  padding: 11px 14px;
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}
.vq-table tbody tr:last-child td{ border-bottom: none; }
.vq-table tbody tr:nth-child(even){ background: rgba(249,250,251,.6); }
.vq-table tbody tr:hover{ background: rgba(96,165,250,.04); }

/* Colored badges for direction */
.vq-badge{
  display: inline-block;
  padding: 3px 10px;
  border-radius: 999px;
  font-size: 11px;
  font-weight: 650;
  letter-spacing: .2px;
}
.vq-badge-up{ background: rgba(52,199,89,.12); color: #1B7A34; }
.vq-badge-down{ background: rgba(255,59,48,.10); color: #C62828; }
.vq-badge-stable{ background: rgba(107,114,128,.10); color: #4B5563; }

/* Section divider with label */
.vq-section-label{
  display: flex; align-items: center; gap: 10px;
  margin: 28px 0 14px 0;
  font-size: 12px; font-weight: 700; letter-spacing: .6px;
  text-transform: uppercase; color: var(--muted);
}
.vq-section-label::after{
  content:''; flex:1; height:1px; background: var(--border);
}

/* Force readable text everywhere */
.stApp, .stApp *{
  color: var(--ink);
}
.stCaption, .stMarkdown small, .stMarkdown em, .stMarkdown .st-emotion-cache-1wmy9hl{
  color: var(--muted) !important;
}

/* Headings */
h1,h2,h3,h4,h5,h6 { color: var(--ink) !important; }

/* Cards (expanders, forms, metrics, dataframes) */
div[data-testid="stExpander"],
div[data-testid="stForm"],
div[data-testid="stMetric"],
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"]{
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-xl) !important;
  box-shadow: var(--shadow) !important;
}

/* Expander header + body */
div[data-testid="stExpander"] > details {
  padding: 10px 14px 14px 14px !important;
  border-radius: var(--r-xl) !important;
}
div[data-testid="stExpander"] summary{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
  padding: 10px 12px !important;
  margin-bottom: 10px !important;
  color: var(--ink) !important;
  font-weight: 650 !important;
}
div[data-testid="stExpander"] summary svg{
  color: var(--muted) !important;
}

/* Labels (fix faint/low opacity) */
label, label *{
  opacity: 1 !important;
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
}

/* Inputs: text/number/date/textarea */
input, textarea{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  caret-color: var(--ink) !important;
}
input::placeholder, textarea::placeholder{
  color: rgba(17,24,39,.45) !important;
  -webkit-text-fill-color: rgba(17,24,39,.45) !important;
}

/* Streamlit input wrappers */
div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] > div,
div[data-testid="stNumberInput"] > div,
div[data-testid="stTextArea"] > div{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
  box-shadow: none !important;
}
/* Date input: flatten all nested wrapper backgrounds */
div[data-testid="stDateInput"] > div{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}
div[data-testid="stDateInput"] > div > div{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}
div[data-testid="stDateInput"] div[data-baseweb="input"]{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
  box-shadow: none !important;
}
div[data-testid="stDateInput"] div[data-baseweb="input"] > div{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}
div[data-testid="stDateInput"] input{
  background: transparent !important;
}
div[data-baseweb="input"] > div:focus-within,
div[data-testid="stTextInput"] > div:focus-within,
div[data-testid="stNumberInput"] > div:focus-within,
div[data-testid="stDateInput"] div[data-baseweb="input"]:focus-within,
div[data-testid="stTextArea"] > div:focus-within{
  border-color: rgba(37,99,235,.40) !important;
  box-shadow: 0 0 0 4px rgba(37,99,235,.18) !important;
}

/* Selectbox / multiselect */
div[data-baseweb="select"] > div{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] input{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
}
div[data-baseweb="select"] input::placeholder{
  color: rgba(17,24,39,.45) !important;
  -webkit-text-fill-color: rgba(17,24,39,.45) !important;
}

/* Dropdown menu */
ul[role="listbox"]{
  background: var(--surface2) !important;
  border: 1px solid var(--border2) !important;
  border-radius: var(--r-lg) !important;
  box-shadow: var(--shadow2) !important;
}
ul[role="listbox"] *{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
}
li[role="option"][aria-selected="true"]{
  background: var(--accentSoft) !important;
}

/* Checkboxes / radios */
div[data-testid="stCheckbox"] label,
div[data-testid="stRadio"] label{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
}
div[data-testid="stCheckbox"] svg,
div[data-testid="stRadio"] svg{
  color: var(--accent) !important;
}

/* Sliders */
div[data-testid="stSlider"] *{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
}

/* Alerts */
div[data-testid="stAlert"] *{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
}

/* Metrics (tiles) */
div[data-testid="stMetric"]{
  padding: 12px 14px !important;
}
div[data-testid="stMetricLabel"]{
  color: var(--muted) !important;
  -webkit-text-fill-color: var(--muted) !important;
  font-weight: 600 !important;
}
div[data-testid="stMetricValue"]{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
  font-size: 2.1rem !important;
  line-height: 1.05 !important;
  white-space: normal !important;
  overflow: visible !important;
  word-break: break-word !important;
}
div[data-testid="stMetricDelta"]{
  color: var(--muted) !important;
  -webkit-text-fill-color: var(--muted) !important;
  opacity: 1 !important;
}

/* Dataframe/editor text */
div[data-testid="stDataFrame"] *, div[data-testid="stDataEditor"] *{
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
}

/* Buttons */
button[kind="primary"]{
  background: var(--accent) !important;
  border: 1px solid rgba(96,165,250,.55) !important;
  border-radius: var(--r-lg) !important;
  color: #0b1220 !important;
  font-weight: 650 !important;
}
button[kind="primary"]:hover{ filter: brightness(0.98); }
button[kind="primary"] *{ color:#0b1220 !important; }
button[kind="secondary"]{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
  color: var(--ink) !important;
}

/* Remove any forced low-opacity text from older Streamlit themes */
.st-emotion-cache-1y4p8pa, .st-emotion-cache-1c7y2kd { opacity: 1 !important; }

/* Stat cards (replaces st.metric for cross-browser readability) */
.vq-stat{
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--r-lg);
  padding: 12px 14px;
  box-shadow: var(--shadow);
  min-height: 78px;
}
.vq-stat-label{
  color: var(--muted) !important;
  font-size: 12px;
  font-weight: 650;
  letter-spacing: .2px;
  margin-bottom: 6px;
  opacity: 1 !important;
  -webkit-text-fill-color: var(--muted) !important;
}
.vq-stat-value{
  color: var(--ink) !important;
  font-size: 20px;
  font-weight: 750;
  line-height: 1.15;
  overflow-wrap: anywhere;
  opacity: 1 !important;
  -webkit-text-fill-color: var(--ink) !important;
}

.vq-stat-value-sm{
  font-size: 16px;
  line-height: 1.2;
}
.vq-stat-delta{
  margin-top: 6px;
  color: var(--accent) !important;
  font-size: 12px;
  font-weight: 650;
  opacity: 1 !important;
  -webkit-text-fill-color: var(--accent) !important;
}

/* Step indicator bar — style the navigation buttons as a sleek breadcrumb trail */
.vq-step-bar{
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 0 12px 0;
}

/* Override Streamlit button styling for step nav buttons */
div[data-testid="stHorizontalBlock"]:has(button[key^="vq_step_nav"]),
div[data-testid="stHorizontalBlock"]:has([data-testid="stButton"]) {
  gap: 6px !important;
}

/* Make step buttons look like text links, not chunky buttons */
button[kind="secondary"]{
  background: transparent !important;
  border: 1px solid transparent !important;
  border-radius: var(--r-lg) !important;
  color: var(--muted) !important;
  font-weight: 500 !important;
  font-size: 13px !important;
  padding: 8px 4px !important;
  transition: all .15s ease;
  box-shadow: none !important;
}
button[kind="secondary"] *{ color: var(--muted) !important; }
button[kind="secondary"]:hover{
  background: rgba(96,165,250,.06) !important;
  border-color: rgba(96,165,250,.15) !important;
  color: var(--ink) !important;
}
button[kind="secondary"]:hover *{ color: var(--ink) !important; }
button[kind="secondary"]:disabled{
  opacity: 0.35 !important;
  cursor: not-allowed !important;
}

/* Primary step button = current step */
button[kind="primary"]{
  background: rgba(96,165,250,.10) !important;
  border: 1px solid rgba(96,165,250,.25) !important;
  border-radius: var(--r-lg) !important;
  color: #1D4ED8 !important;
  font-weight: 700 !important;
  font-size: 13px !important;
  padding: 8px 4px !important;
  box-shadow: none !important;
}
button[kind="primary"] *{ color: #1D4ED8 !important; }
button[kind="primary"]:hover{
  background: rgba(96,165,250,.16) !important;
  filter: none !important;
}

/* Sticky inspector panel (Report step) */
.vq-sticky{
  position: sticky;
  top: 14px;
  align-self: flex-start;
}

/* Action buttons (Next/Back/Download) — restore solid appearance.
   Step nav buttons are in the top row; everything else should look like a real button. */
div[data-testid="stExpander"] button[kind="primary"],
div[data-testid="stExpander"] button[kind="secondary"],
div[data-testid="stForm"] button[kind="primary"],
div[data-testid="stForm"] button[kind="secondary"],
div[data-testid="stDownloadButton"] button,
.vq-action-row button[kind="primary"],
.vq-action-row button[kind="secondary"]{
  font-size: 14px !important;
  padding: 10px 18px !important;
}
.vq-action-row button[kind="primary"]{
  background: var(--accent) !important;
  border: 1px solid rgba(96,165,250,.55) !important;
  color: #0b1220 !important;
  font-weight: 650 !important;
  box-shadow: 0 2px 8px rgba(96,165,250,.18) !important;
}
.vq-action-row button[kind="primary"] *{ color: #0b1220 !important; }
.vq-action-row button[kind="primary"]:hover{
  filter: brightness(0.96) !important;
  background: var(--accent) !important;
}
.vq-action-row button[kind="secondary"]{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  color: var(--ink) !important;
  font-weight: 600 !important;
}
.vq-action-row button[kind="secondary"] *{ color: var(--ink) !important; }
.vq-action-row button[kind="secondary"]:hover{
  background: rgba(96,165,250,.06) !important;
  border-color: rgba(96,165,250,.20) !important;
}

/* Download buttons always look solid */
div[data-testid="stDownloadButton"] button{
  background: var(--surface2) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important;
  color: var(--ink) !important;
  font-weight: 600 !important;
  box-shadow: var(--shadow) !important;
}
div[data-testid="stDownloadButton"] button *{ color: var(--ink) !important; }
div[data-testid="stDownloadButton"] button:hover{
  background: rgba(96,165,250,.06) !important;
  border-color: rgba(96,165,250,.20) !important;
}

</style>
    """, unsafe_allow_html=True)


def show_progress_indicator(current_step: int):
    """Modern minimal step indicator with clickable text steps.

    Uses individual st.button widgets (one per step) rendered inside a
    horizontal column layout, styled via CSS to look like a sleek breadcrumb
    / step trail rather than standard Streamlit buttons.
    """
    steps = [
        (1, "Subject"),
        (2, "Upload"),
        (3, "Diagnostics"),
        (4, "Comps"),
        (5, "Report"),
    ]

    max_step = 1
    if st.session_state.get("subject_address"):
        max_step = 2
    if st.session_state.get("uploaded_data") is not None:
        max_step = max(max_step, 3)
    if st.session_state.get("diagnostics_df") is not None:
        max_step = max(max_step, 4)
    if st.session_state.get("selected_comps"):
        max_step = max(max_step, 5)

    # Render clickable step buttons in columns
    cols = st.columns(len(steps))
    for col, (step_num, step_label) in zip(cols, steps):
        with col:
            is_current = step_num == current_step
            is_completed = step_num < current_step
            is_accessible = step_num <= max_step

            if is_current:
                status_class = "current"
                icon = f"**{step_num}**"
            elif is_completed:
                status_class = "completed"
                icon = "✓"
            else:
                status_class = "upcoming"
                icon = str(step_num)

            # Use unique key per step
            btn_key = f"vq_step_nav_{step_num}"
            if st.button(
                f"{icon}  {step_label}",
                key=btn_key,
                use_container_width=True,
                disabled=not is_accessible,
                type="primary" if is_current else "secondary",
            ):
                if is_accessible and step_num != current_step:
                    st.session_state["step"] = step_num
                    st.rerun()


# ----------------------------
# Utility Functions (from v2)
# ----------------------------



def render_step_header(title: str, subtitle: str | None = None, icon: str | None = None):
    """Render a glass-style header card for each step."""
    icon_html = f'<span style="margin-right:10px;">{icon}</span>' if icon else ''
    subtitle_html = f'<div style="margin-top:6px; color:rgba(11,18,32,.72); font-size:14px;">{subtitle}</div>' if subtitle else ''
    st.markdown(
        f'<div class="glass-card">'
        f'<div style="display:flex; align-items:center;">'
        f'{icon_html}'
        f'<div style="font-size:18px; font-weight:750; letter-spacing:.2px;">{title}</div>'
        f'</div>'
        f'{subtitle_html}'
        f'</div>',
        unsafe_allow_html=True
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    return df


def canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map common MLS/export column variants to the canonical required names."""
    df = df.copy()
    # Build a case-insensitive lookup for current columns
    cols_lower = {str(c).strip().lower(): c for c in df.columns}

    def find(*names: str):
        for n in names:
            key = n.strip().lower()
            if key in cols_lower:
                return cols_lower[key]
        return None

    colmap = {}

    # Address
    addr = find("address", "property address", "street address", "site address")
    if addr and addr != "Address":
        colmap[addr] = "Address"

    # Zip
    zipc = find("zip", "zip code", "zipcode", "postal code")
    if zipc and zipc != "Zip":
        colmap[zipc] = "Zip"

    # Contract/Pending date
    pend = find("pending date", "contract date", "pendingdate", "pending", "contractdate", "contract_date", "pending_date")
    if pend and pend != "Pending Date":
        colmap[pend] = "Pending Date"

    # Sold price
    price = find("sold price", "sale price", "soldprice", "saleprice", "price", "sold_price", "sale_price")
    if price and price != "Sold Price":
        colmap[price] = "Sold Price"

    if colmap:
        df = df.rename(columns=colmap)
    return df

def parse_dates_robust(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"": np.nan, "nan": np.nan, "NaN": np.nan, "None": np.nan})
    parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if parsed.notna().mean() < 0.50:
        parsed2 = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
        if parsed2.notna().mean() > parsed.notna().mean():
            parsed = parsed2
    return parsed

def parse_money_robust(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = (
        s.str.replace("$", "", regex=False)
         .str.replace(",", "", regex=False)
         .str.replace(" ", "", regex=False)
    )
    s = s.replace({"": np.nan, "nan": np.nan, "NaN": np.nan, "None": np.nan})
    return pd.to_numeric(s, errors="coerce")

def month_start(d) -> date:
    if isinstance(d, str):
        d = pd.to_datetime(d).date()
    elif hasattr(d, 'date'):
        d = d.date()
    return date(d.year, d.month, 1)

def money(x):
    if pd.isna(x):
        return ""
    return f"${x:,.0f}"

def pct(x, decimals=2):
    if pd.isna(x):
        return ""
    return f"{x:.{decimals}f}%"

def pct_change(i_eff, i_contract):
    if pd.isna(i_eff) or pd.isna(i_contract) or i_contract == 0:
        return np.nan
    return (i_eff / i_contract) - 1.0

def days_between(d1: date, d2: date) -> int:
    return abs((pd.Timestamp(d1) - pd.Timestamp(d2)).days)

def categorize_adjustment(adj_pct: float, threshold: float = 0.5) -> str:
    if pd.isna(adj_pct):
        return "N/A"
    elif adj_pct > threshold:
        return "Increasing"
    elif adj_pct < -threshold:
        return "Declining"
    else:
        return "Stable"

def adjustment_direction(adj_pct: float, threshold: float = 0.1) -> str:
    if pd.isna(adj_pct) or abs(adj_pct) < threshold:
        return "NO ADJUSTMENT"
    elif adj_pct > 0:
        return "UPWARD"
    else:
        return "DOWNWARD"

# ----------------------------
# Index construction
# ----------------------------
def build_monthly_index_price(df: pd.DataFrame, min_sales_per_month: int = 5) -> pd.DataFrame:
    d = df.copy()
    d = d.dropna(subset=["ContractDate", "SoldPrice"])
    d["Month"] = d["ContractDate"].apply(month_start)
    d = d.dropna(subset=["Month"])

    m = (
        d.groupby("Month", as_index=False)
         .agg(SalesCount=("SoldPrice", "size"),
              MedianPrice=("SoldPrice", "median"))
         .sort_values("Month")
    )
    m["ThinMonth"] = m["SalesCount"] < min_sales_per_month

    m["Index_Raw"] = np.nan
    if len(m) > 0:
        m.loc[m.index[0], "Index_Raw"] = 1.00
        for i in range(1, len(m)):
            prev = m.iloc[i - 1]["MedianPrice"]
            cur = m.iloc[i]["MedianPrice"]
            if pd.isna(prev) or prev == 0 or pd.isna(cur):
                m.loc[m.index[i], "Index_Raw"] = m.iloc[i - 1]["Index_Raw"]
            else:
                m.loc[m.index[i], "Index_Raw"] = m.iloc[i - 1]["Index_Raw"] * (cur / prev)

    return m

def add_smoothed_and_regression(index_df: pd.DataFrame, smooth_window: int = 6) -> pd.DataFrame:
    idx = index_df.copy()
    w = max(2, int(smooth_window))

    idx["Index_Smoothed"] = (
        idx["Index_Raw"]
        .rolling(window=w, center=True, min_periods=1)
        .mean()
    )

    # Rolling standard deviation for confidence band
    idx["Index_Std"] = (
        idx["Index_Raw"]
        .rolling(window=w, center=True, min_periods=2)
        .std()
        .fillna(0)
    )

    x = np.arange(len(idx), dtype=float)
    y = idx["Index_Raw"].astype(float).values
    if len(idx) >= 2 and np.isfinite(y).sum() >= 2:
        coef = np.polyfit(x[np.isfinite(y)], y[np.isfinite(y)], 1)
        idx["Index_Regression"] = coef[0] * x + coef[1]
    else:
        idx["Index_Regression"] = idx["Index_Smoothed"]

    return idx

def lookup_index(index_df: pd.DataFrame, target_date: date, index_col: str):
    if index_df.empty or target_date is None:
        return np.nan, None, "no_index"
    m = month_start(target_date)
    row = index_df[index_df["Month"] == m]
    if not row.empty:
        return float(row.iloc[0][index_col]), m, "exact"
    prior = index_df[index_df["Month"] <= m].sort_values("Month")
    if prior.empty:
        return np.nan, None, "no_prior"
    used_month = prior.iloc[-1]["Month"]
    return float(prior.iloc[-1][index_col]), used_month, "prior"

# ----------------------------
# Cook's Distance
# ----------------------------
def cooks_distance_time_regression(df: pd.DataFrame) -> pd.DataFrame:
    """Flag outliers based on price deviation from the time-based trend.

    Uses a simple linear regression of log(price) on time, but flags are based on
    PRICE residuals (how far a sale's price deviates from the fitted trend), not on
    temporal position.  Recent sales that are in-line with the trend are never flagged,
    even if they sit at the edge of the date range.
    """
    out = df.copy()
    out = out.dropna(subset=["ContractDate", "SoldPrice"]).copy()
    out["ContractDate"] = pd.to_datetime(out["ContractDate"]).dt.date

    if len(out) < 6:
        out["Leverage"] = np.nan
        out["CooksD"] = np.nan
        out["HighLeverage"] = False
        out["HighCooksD"] = False
        return out

    t0 = pd.to_datetime(out["ContractDate"]).min()
    x1 = (pd.to_datetime(out["ContractDate"]) - t0).dt.days.astype(float).values
    y = np.log(np.maximum(1.0, out["SoldPrice"].astype(float).values))

    n = len(y)
    p = 2

    X = np.column_stack([np.ones(n), x1])
    XtX = X.T @ X
    try:
        XtX_inv = np.linalg.inv(XtX)
    except np.linalg.LinAlgError:
        XtX_inv = np.linalg.pinv(XtX)

    beta = XtX_inv @ (X.T @ y)
    yhat = X @ beta
    resid = y - yhat

    H = np.sum((X @ XtX_inv) * X, axis=1)
    mse = np.sum(resid ** 2) / max(1, (n - p))

    eps = 1e-12
    cooks = (resid**2 / (p * (mse + eps))) * (H / (np.maximum(eps, (1 - H))**2))

    # Flag based on price deviation (studentized residual), NOT leverage
    # A sale is a price outlier if its residual is > 2 standard deviations from trend
    resid_std = np.std(resid) if np.std(resid) > 0 else 1.0
    studentized = np.abs(resid) / resid_std

    cook_thresh = 4 / n

    out["Leverage"] = H
    out["CooksD"] = cooks
    # HighLeverage now flags PRICE outliers (large residuals), not temporal edge cases
    out["HighLeverage"] = studentized > 2.0
    out["HighCooksD"] = out["CooksD"] > cook_thresh
    return out

@st.cache_data(show_spinner=False)
def compute_iqr_flags_cached(sold_prices: np.ndarray, k: float) -> np.ndarray:
    s = pd.Series(sold_prices.astype(float))
    q1 = s.quantile(0.25)
    q3 = s.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - k * iqr
    hi = q3 + k * iqr
    return ((s < lo) | (s > hi)).to_numpy()

@st.cache_data(show_spinner=False)
def build_index_cached(
    contract_dates: np.ndarray,
    sold_prices: np.ndarray,
    min_sales: int,
    smooth_window: int
) -> pd.DataFrame:
    df_temp = pd.DataFrame({
        'ContractDate': pd.to_datetime(contract_dates).date,
        'SoldPrice': sold_prices
    })
    idx = build_monthly_index_price(df_temp, min_sales)
    idx = add_smoothed_and_regression(idx, smooth_window)
    return idx

# ----------------------------
# Chart function (from v2)
# ----------------------------
def plot_fannie_style_chart(
    index_df: pd.DataFrame,
    comps_df: pd.DataFrame,
    eff_date: date,
    eff_index: float,
    index_col: str = "Index_Smoothed",
    show_raw: bool = False,
    show_thin: bool = False,
    tick_mode: str = "Monthly",
    lookback_months: int = 12
) -> plt.Figure:
    """Create a sleek, modern chart with methodology visualization"""
    fig, ax = plt.subplots(figsize=(14, 6.2), facecolor='none')
    ax.set_facecolor((1, 1, 1, 0))
    
    index_df = index_df.copy()
    index_df['MonthDT'] = pd.to_datetime(index_df['Month'])
    index_df['IndexPct'] = (index_df[index_col] - 1.0) * 100
    
    # --- Confidence band (shows smoothing uncertainty) ---
    if 'Index_Std' in index_df.columns and index_col == 'Index_Smoothed':
        std_pct = index_df['Index_Std'] * 100
        upper = index_df['IndexPct'] + std_pct
        lower = index_df['IndexPct'] - std_pct
        ax.fill_between(index_df['MonthDT'], lower, upper,
                        color=THEME_COLORS['primary'], alpha=0.08, zorder=1,
                        label='Confidence Band (±1σ)')
    
    # --- Raw index (always shown faintly to demonstrate smoothing) ---
    if 'Index_Raw' in index_df.columns:
        index_df['RawPct'] = (index_df['Index_Raw'] - 1.0) * 100
        ax.plot(index_df['MonthDT'], index_df['RawPct'], 
                color=THEME_COLORS['secondary'], linewidth=1.2, 
                alpha=0.25, linestyle='-', zorder=2,
                label='Raw Chain-Linked Index')
    
    # --- Primary index line ---
    index_label = {
        'Index_Smoothed': 'Smoothed Index',
        'Index_Raw': 'Raw Chain-Linked Index',
        'Index_Regression': 'Regression Trendline',
    }.get(index_col, 'Index')
    
    ax.plot(index_df['MonthDT'], index_df['IndexPct'], 
            color=THEME_COLORS['primary'], linewidth=3.5, 
            zorder=3, solid_capstyle='round', alpha=0.9,
            label=index_label)
    
    # --- Additional overlays ---
    if show_raw and 'Index_Raw' in index_df.columns and index_col != 'Index_Raw':
        # Already shown faintly above; make it a bit more visible if toggled on
        ax.plot(index_df['MonthDT'], index_df['RawPct'], 
                color=THEME_COLORS['secondary'], linewidth=1.8, 
                alpha=0.45, linestyle='--', zorder=2)
    
    if show_thin and 'ThinMonth' in index_df.columns:
        thin = index_df[index_df['ThinMonth']]
        if not thin.empty:
            thin_pct = (thin[index_col] - 1.0) * 100
            ax.scatter(thin['MonthDT'], thin_pct, 
                      color=THEME_COLORS['warning'], s=100, 
                      marker='x', zorder=4, alpha=0.6, linewidths=2)
    
    eff_dt = pd.to_datetime(eff_date)
    eff_idx_pct = (eff_index - 1.0) * 100
    
    if not comps_df.empty:
        comps_sorted = comps_df.sort_values('ContractDate').reset_index(drop=True)
        
        # Collect all label positions to check for overlaps later
        label_records = []
        
        for i, (_, row) in enumerate(comps_sorted.iterrows()):
            contract_dt = pd.to_datetime(row['ContractDate'])
            contract_idx_pct = (row['Index_Contract'] - 1.0) * 100
            adj_pct = row['MktAdjPct']
            address = row['CompAddress']
            
            if abs(adj_pct) < 0.1:
                color = THEME_COLORS['stable']
            elif adj_pct > 0:
                color = THEME_COLORS['upward']
            else:
                color = THEME_COLORS['downward']
            
            # Dotted connector line to effective date
            ax.plot([contract_dt, eff_dt], 
                   [contract_idx_pct, eff_idx_pct],
                   color=color, linewidth=1.2, linestyle=':', 
                   alpha=0.35, zorder=2)
            
            # Comp dot
            ax.scatter(contract_dt, contract_idx_pct, 
                      color=color, s=180, marker='o', 
                      zorder=5, edgecolors='white', linewidths=2.5,
                      alpha=0.95)
            
            # Shorten address — keep enough to be identifiable
            addr_parts = address.split()
            if len(addr_parts) > 5:
                address_short = ' '.join(addr_parts[:5]) + '…'
            elif len(address) > 35:
                address_short = address[:32] + '…'
            else:
                address_short = address
            
            label_text = f"{address_short}  {adj_pct:+.1f}%"
            
            label_records.append({
                'dt': contract_dt,
                'y': contract_idx_pct,
                'text': label_text,
                'color': color,
                'idx': i,
            })
        
        # ---- Label placement: deterministic slot-based approach ----
        # Sort labels by x-position (date), then assign vertical slots
        # so that labels near each other in time don't overlap.

        import matplotlib.dates as _mdates
        import matplotlib.transforms as _mtrans

        for rec in label_records:
            rec['dt_num'] = _mdates.date2num(rec['dt'])

        label_records_sorted = sorted(label_records, key=lambda r: r['dt_num'])

        # Convert each data point to axes fraction coords (0-1, 0-1)
        for rec in label_records_sorted:
            pt_display = ax.transData.transform((rec['dt_num'], rec['y']))
            pt_axes = ax.transAxes.inverted().transform(pt_display)
            rec['ax_x'] = pt_axes[0]
            rec['ax_y'] = pt_axes[1]

        # Label dimensions in axes fraction
        fig_w_px = fig.get_size_inches()[0] * fig.dpi
        fig_h_px = fig.get_size_inches()[1] * fig.dpi
        ax_bbox = ax.get_position()
        ax_w_px = ax_bbox.width * fig_w_px
        ax_h_px = ax_bbox.height * fig_h_px

        LABEL_W = 185 / max(1, ax_w_px)
        LABEL_H = 22 / max(1, ax_h_px)
        GAP = 6 / max(1, ax_h_px)

        placed_boxes = []  # (x_left, y_bottom, x_right, y_top)

        def _overlaps_any(x_left, y_bottom, x_right, y_top):
            for (px1, py1, px2, py2) in placed_boxes:
                if x_left < px2 and x_right > px1 and y_bottom < py2 and y_top > py1:
                    return True
            return False

        for rec in label_records_sorted:
            ax_x = rec['ax_x']
            ax_y = rec['ax_y']

            # Center label horizontally over the data point
            lbl_x_center = ax_x - LABEL_W / 2

            # Clamp to stay within axes
            lbl_x_center = max(0.01, min(lbl_x_center, 1.0 - LABEL_W - 0.01))

            # Try offsets: above first, then below, increasing distance
            # Start at step 2 so the label clears the data point marker
            best = None
            for step in range(2, 16):
                # Above
                y_above = ax_y + (LABEL_H + GAP) * step
                if 0.02 < y_above < (0.96 - LABEL_H):
                    if not _overlaps_any(lbl_x_center, y_above, lbl_x_center + LABEL_W, y_above + LABEL_H):
                        best = (lbl_x_center, y_above)
                        break
                # Below
                y_below = ax_y - (LABEL_H + GAP) * step - LABEL_H
                if 0.02 < y_below < (0.96 - LABEL_H):
                    if not _overlaps_any(lbl_x_center, y_below, lbl_x_center + LABEL_W, y_below + LABEL_H):
                        best = (lbl_x_center, y_below)
                        break

            if not best:
                best = (lbl_x_center, ax_y + LABEL_H + GAP)

            lbl_x, lbl_y = best
            placed_boxes.append((lbl_x, lbl_y, lbl_x + LABEL_W, lbl_y + LABEL_H))

            ax.annotate(
                rec['text'],
                xy=(rec['dt'], rec['y']),
                xytext=(lbl_x + LABEL_W / 2, lbl_y + LABEL_H / 2),
                textcoords='axes fraction',
                ha='center', va='center', fontsize=8.5, fontweight='600',
                color=rec['color'], zorder=7,
                bbox=dict(boxstyle='round,pad=0.35', facecolor='white',
                          edgecolor=rec['color'], alpha=0.92, linewidth=0.6),
                arrowprops=dict(arrowstyle='->', color=rec['color'],
                                linewidth=0.9, alpha=0.5,
                                connectionstyle='arc3,rad=0.1'),
            )
    
    ax.scatter([eff_dt], [eff_idx_pct], 
              color=THEME_COLORS['primary'], s=300, marker='D', 
              zorder=6, edgecolors='white', linewidths=3, alpha=0.95)
    
    ax.annotate('Effective\nDate', xy=(eff_dt, eff_idx_pct),
               xytext=(18, 18), textcoords='offset points',
               ha='left', va='bottom', fontsize=12, fontweight='700',
               color=THEME_COLORS['primary'], zorder=7,
               bbox=dict(boxstyle='round,pad=0.5', facecolor='white',
                       edgecolor='none', alpha=0.90))
    
    ax.axvline(eff_dt, color=THEME_COLORS['text_secondary'], 
              linestyle='--', linewidth=1.5, alpha=0.2, zorder=1)
    
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, _: f'{y:+.1f}%'))
    # X-axis tick density
    if str(tick_mode).lower().startswith('quarter'):
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha='center')

    ax.set_title('Market Condition Trend', fontsize=16, fontweight='700', color=THEME_COLORS['text'], pad=16)
    ax.set_xlabel('Contract Month', fontsize=12, fontweight='500', color=THEME_COLORS['text_secondary'], labelpad=10)
    ax.set_ylabel('Index Change (%)', fontsize=12, color=THEME_COLORS['text_secondary'], labelpad=10)
    ax.grid(axis='y', alpha=0.15, linewidth=1)
    ax.set_axisbelow(True)

    # Smart x-axis range: use lookback period, extending for older comps
    default_start = pd.to_datetime(eff_date) - pd.DateOffset(months=lookback_months + 1)
    x_end = pd.to_datetime(eff_date) + pd.DateOffset(months=2)

    # Extend start if any comp or index data goes further back
    earliest_needed = default_start
    if not comps_df.empty:
        earliest_comp = pd.to_datetime(comps_df['ContractDate']).min()
        if earliest_comp < earliest_needed:
            earliest_needed = earliest_comp - pd.DateOffset(months=1)
    if not index_df.empty:
        earliest_index = index_df['MonthDT'].min()
        if earliest_index < earliest_needed:
            # Only extend to index data if a comp needs it
            if not comps_df.empty and pd.to_datetime(comps_df['ContractDate']).min() < default_start:
                earliest_needed = min(earliest_needed, earliest_index)

    ax.set_xlim(earliest_needed, x_end)

    
    for spine in ax.spines.values():
        spine.set_visible(False)
    
    ax.tick_params(colors=THEME_COLORS['text_secondary'], labelsize=11,
                   length=0, pad=10)
    
    # Compact legend in upper-left to show methodology layers
    legend = ax.legend(loc='upper left', frameon=True, fontsize=8,
                       framealpha=0.85, edgecolor='none',
                       facecolor='white', borderpad=0.8,
                       labelspacing=0.4, handlelength=1.8)
    legend.get_frame().set_linewidth(0)
    for text in legend.get_texts():
        text.set_color(THEME_COLORS['text_secondary'])
    
    y_min, y_max = ax.get_ylim()
    y_range = y_max - y_min
    ax.set_ylim(y_min - y_range * 0.10, y_max + y_range * 0.12)
    
    ax.set_facecolor('white')
    fig.patch.set_facecolor('white')
    
    fig.tight_layout(pad=1.5)
    return fig

# ----------------------------
# Table image renderer
# ----------------------------
def render_table_image(
    comp_data: pd.DataFrame,
    subject_address: str,
    eff_date: date,
    overall_trend: str,
    overall_change_pct: float,
) -> bytes:
    """Render the comparable adjustments table as a professional PNG image."""

    # Build display rows
    rows = []
    for i, (_, r) in enumerate(comp_data.iterrows()):
        adj_pct = r["MktAdjPct"]
        applied = r["AppliedAdj"]
        if not applied or abs(adj_pct) < 0.1:
            adj_str = "Stable"
            dollar_str = "—"
        else:
            adj_str = f"{adj_pct:+.2f}%"
            dollar_str = f"${r['MktAdj$']:+,.0f}"

        contract_dt = r["ContractDate"]
        if hasattr(contract_dt, 'strftime'):
            date_str = contract_dt.strftime("%b %d, %Y")
        else:
            date_str = str(contract_dt)

        rows.append([
            str(i + 1),
            str(r["CompAddress"])[:35],
            date_str,
            f"${r['SalePrice']:,.0f}",
            adj_str,
            dollar_str,
        ])

    col_labels = ["#", "Address", "Contract Date", "Sale Price", "Adjustment", "Adj $"]
    n_rows = len(rows)
    n_cols = len(col_labels)

    # Figure sizing
    fig_width = 10
    row_height = 0.38
    header_height = 0.45
    title_height = 0.9
    fig_height = title_height + header_height + (n_rows * row_height) + 0.5

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, fig_height)
    ax.axis('off')
    fig.patch.set_facecolor('white')

    # Column positions (left edges)
    col_x = [0.02, 0.06, 0.38, 0.56, 0.72, 0.86]
    col_align = ['center', 'left', 'left', 'right', 'right', 'right']

    # Title block
    y_top = fig_height - 0.25
    ax.text(0.02, y_top, subject_address, fontsize=13, fontweight='bold',
            color='#1a1a2e', va='top', fontfamily='sans-serif')
    ax.text(0.02, y_top - 0.32, f"Effective Date: {eff_date.strftime('%B %d, %Y')}  ·  Trend: {overall_trend}  ·  Overall Change: {overall_change_pct:+.2f}%",
            fontsize=9, color='#6B7280', va='top', fontfamily='sans-serif')

    # Header row
    y_header = y_top - 0.85
    ax.fill_between([0, 1], y_header - 0.05, y_header + 0.32, color='#F3F4F6', zorder=0)
    for j, label in enumerate(col_labels):
        ax.text(col_x[j], y_header + 0.12, label,
                fontsize=8.5, fontweight='bold', color='#374151',
                va='center', ha=col_align[j], fontfamily='sans-serif')

    # Data rows
    for i, row in enumerate(rows):
        y_row = y_header - 0.15 - (i * row_height)
        # Alternating row background
        if i % 2 == 1:
            ax.fill_between([0, 1], y_row - 0.14, y_row + 0.22, color='#F9FAFB', zorder=0)
        # Gridline
        ax.plot([0.02, 0.98], [y_row - 0.14, y_row - 0.14], color='#E5E7EB', linewidth=0.5, zorder=1)

        for j, cell in enumerate(row):
            # Color the adjustment column
            color = '#111827'
            fw = 'normal'
            if j == 4:  # Adjustment column
                if cell.startswith('+'):
                    color = '#22c55e'
                    fw = 'bold'
                elif cell.startswith('-'):
                    color = '#ef4444'
                    fw = 'bold'
                elif cell == 'Stable':
                    color = '#6B7280'
            if j == 5 and cell != '—':  # Adj $ column
                fw = 'bold'
            if j == 0:  # Comp number
                color = THEME_COLORS.get('primary', '#3B82F6')
                fw = 'bold'

            ax.text(col_x[j], y_row + 0.04, cell,
                    fontsize=9, fontweight=fw, color=color,
                    va='center', ha=col_align[j], fontfamily='sans-serif')

    # Footer
    y_footer = y_header - 0.15 - (n_rows * row_height) - 0.15
    ax.text(0.98, y_footer, "Generated by MarketAdjuster", fontsize=7,
            color='#9CA3AF', va='top', ha='right', fontfamily='sans-serif', style='italic')

    fig.subplots_adjust(left=0.01, right=0.99, top=0.98, bottom=0.02)

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=200, bbox_inches='tight', facecolor='white', pad_inches=0.15)
    plt.close(fig)
    return buf.getvalue()

# ----------------------------
# Narrative builder
# ----------------------------
def build_narrative(
    subject_address: str,
    date_start: date,
    date_end: date,
    eff_date: date,
    eff_index: float,
    index_option: str,
    math_col: str,
    comp_rows: pd.DataFrame,
    index_df: pd.DataFrame,
    excluded_count: int,
    overall_trend: str,
    overall_change_pct: float = None,
    trend_lookback: str = "1 Year",
    raw_sales_df: pd.DataFrame = None,
) -> str:
    # Use pre-computed change if provided; otherwise calculate from full range
    if overall_change_pct is None:
        if not index_df.empty and math_col in index_df.columns:
            first_idx = float(index_df.iloc[0][math_col])
            last_idx = float(index_df.iloc[-1][math_col])
            overall_change_pct = ((last_idx / first_idx) - 1.0) * 100
        else:
            overall_change_pct = 0.0
    
    narrative = f"""MARKET CONDITION ADJUSTMENT ANALYSIS
{'=' * 60}

SUBJECT PROPERTY: {subject_address}
EFFECTIVE DATE: {eff_date.strftime('%B %d, %Y')}

ANALYSIS PERIOD: {date_start.strftime('%B %Y')} through {date_end.strftime('%B %Y')}

OVERALL MARKET TREND
{'-' * 60}
The overall property value trend for the {trend_lookback.lower()} lookback period is {overall_trend.upper()}.
Market change during this period was {overall_change_pct:+.2f}%.
Full data range: {date_start.strftime('%B %Y')} through {date_end.strftime('%B %Y')}.

"""
    # Quarterly median price breakdown
    if raw_sales_df is not None and not raw_sales_df.empty:
        eff_dt = pd.to_datetime(eff_date)
        _sd = raw_sales_df.copy()
        _sd["ContractDate"] = pd.to_datetime(_sd["ContractDate"])
        _sd["SoldPrice"] = _sd["SoldPrice"].astype(float)

        quarters = [
            ("0-3 Months", eff_dt - pd.DateOffset(months=3), eff_dt),
            ("3-6 Months", eff_dt - pd.DateOffset(months=6), eff_dt - pd.DateOffset(months=3)),
            ("6-9 Months", eff_dt - pd.DateOffset(months=9), eff_dt - pd.DateOffset(months=6)),
            ("9-12 Months", eff_dt - pd.DateOffset(months=12), eff_dt - pd.DateOffset(months=9)),
        ]

        q_lines = []
        for label, dt_start, dt_end in quarters:
            mask = (_sd["ContractDate"] >= dt_start) & (_sd["ContractDate"] < dt_end)
            subset = _sd[mask]
            if len(subset) > 0:
                med = subset["SoldPrice"].median()
                q_lines.append(f"  {label:14s}  Median: ${med:>12,.0f}   ({len(subset)} sales)")
            else:
                q_lines.append(f"  {label:14s}  No sales data")

        narrative += f"""QUARTERLY MEDIAN SALE PRICE (trailing from effective date)
{'-' * 60}
""" + "\n".join(q_lines) + "\n\n"
    
    narrative += f"""INDEX METHODOLOGY
{'-' * 60}
Index Type: {index_option}
Index at Effective Date: {eff_index:.4f} ({((eff_index - 1.0) * 100):+.2f}%)
Data Quality: {len(index_df)} months of data; {excluded_count} sale(s) excluded as outliers

INDIVIDUAL COMPARABLE ADJUSTMENTS
{'-' * 60}
Per Fannie Mae guidance, adjustments to comparable sales are based on market 
changes between the contract date of each comparable and the effective date 
of the appraisal. Therefore, different comparables may receive positive, 
negative, or no adjustments depending on their specific contract timing.

"""
    
    for idx, row in comp_rows.iterrows():
        comp_num = idx + 1
        addr = row['CompAddress']
        contract_date = row['ContractDate']
        sale_price = row['SalePrice']
        contract_idx = row['Index_Contract']
        adj_pct = row['MktAdjPct']
        adj_dollar = row['MktAdj$']
        applied = row['AppliedAdj']
        
        direction = adjustment_direction(adj_pct)
        category = categorize_adjustment(adj_pct)
        
        if not applied:
            status_text = "NO adjustment (within minimum time threshold)"
        else:
            status_text = f"{direction} adjustment of {abs(adj_pct):.2f}%"
        
        contract_idx_pct = ((contract_idx - 1.0) * 100)
        
        narrative += f"""
Comparable {comp_num}: {addr}
  Contract Date: {contract_date.strftime('%B %Y')}
  Sale Price: ${sale_price:,.0f}
  Market Index at Contract: {contract_idx:.4f} ({contract_idx_pct:+.2f}%)
  Market Index at Effective: {eff_index:.4f} ({((eff_index - 1.0) * 100):+.2f}%)
  Category: {category}
  Adjustment: {status_text}
  Dollar Adjustment: {adj_dollar:+,.0f}
"""
    
    narrative += f"""
{'=' * 60}

CONCLUSION
{'-' * 60}
This analysis demonstrates that while the overall market trend is {overall_trend},
individual comparable sales reflect varying market conditions based on their
specific contract timing. This follows Fannie Mae guidance that comparable
sales in the same appraisal report may have positive, negative, or no 
adjustments applied.

For more information, refer to Fannie Mae Selling Guide section B4-1.3-09, 
Adjustments to Comparable Sales.

Analysis prepared: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
Tool: {APP_NAME}
"""
    
    return narrative

# ----------------------------
# MAIN APP WITH WORKFLOW
# ----------------------------
def build_pdf_addendum(
    subject_address: str,
    eff_date: date,
    settings: dict,
    narrative: str,
    chart_png: bytes,
    comp_table: pd.DataFrame,
) -> bytes:
    """Create an appraisal-ready PDF addendum pack (chart + table + narrative)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40,
        title="Market Condition Adjustment Addendum",
    )
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    body = styles["BodyText"]
    body.leading = 13

    elements = []
    elements.append(Paragraph("Market Condition Adjustment Addendum", title_style))
    elements.append(Spacer(1, 8))

    meta_lines = [
        f"<b>Subject:</b> { _escape(subject_address) }",
        f"<b>Effective Date:</b> { eff_date.strftime('%B %d, %Y') }",
        f"<b>Index Type:</b> { _escape(str(settings.get('index_option',''))) }",
        f"<b>Smoothing Window:</b> { int(settings.get('smooth_window', 6)) } months",
        f"<b>Min Sales/Month:</b> { int(settings.get('min_sales_per_month', 3)) }",
        f"<b>Minimum Days for Adjustment:</b> { int(settings.get('no_adj_days', 0)) }",
    ]
    elements.append(Paragraph("<br/>".join(meta_lines), body))
    elements.append(Spacer(1, 12))

    # Chart
    img = RLImage(io.BytesIO(chart_png), width=520, height=300)
    elements.append(img)
    elements.append(Spacer(1, 12))

    if not comp_table.empty:
        elements.append(Paragraph("<b>Comparable Market Condition Adjustments</b>", body))
        elements.append(Spacer(1, 6))

        table_df = comp_table.copy()
        for c in table_df.columns:
            table_df[c] = table_df[c].astype(str)

        data = [table_df.columns.tolist()] + table_df.values.tolist()
        tbl = Table(data, hAlign="LEFT", colWidths=[40, 210, 75, 70, 55, 70])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8EEF7")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0B1220")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 9),
            ("FONTSIZE", (0,1), (-1,-1), 8),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))
        elements.append(tbl)
        elements.append(Spacer(1, 12))

    if narrative:
        elements.append(Paragraph("<b>Narrative</b>", body))
        elements.append(Spacer(1, 6))
        safe_narr = _escape(narrative).replace("\n", "<br/>")
        elements.append(Paragraph(safe_narr, body))

    doc.build(elements)
    return buf.getvalue()


def main():
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    inject_modern_css()
    
    # Initialize session state
    if "step" not in st.session_state:
        st.session_state["step"] = 1
    if "subject_address" not in st.session_state:
        st.session_state["subject_address"] = ""
    if "eff_date" not in st.session_state:
        st.session_state["eff_date"] = date.today()
    if "uploaded_data" not in st.session_state:
        st.session_state["uploaded_data"] = None
    if "excluded_rowids" not in st.session_state:
        st.session_state["excluded_rowids"] = set()
    if "selected_comps" not in st.session_state:
        st.session_state["selected_comps"] = []
    if "settings" not in st.session_state:
        st.session_state["settings"] = {
            "no_adj_days": 90,
            "trend_lookback": "1 Year",
            "trend_override": "Auto-detect",
            "index_option": "Smoothed (Recommended)",
            "smooth_window": 6,
            "min_sales_per_month": 5,
            "use_iqr": True,
            "iqr_multiplier": 1.0,
            "use_cooks": False,
            "cooks_threshold": float("nan"),
        }
    
    # Header with logo
    _logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MarketAdjuster_macOS_512.png")
    hdr_cols = st.columns([0.10, 0.64, 0.26])
    with hdr_cols[0]:
        if os.path.exists(_logo_path):
            st.image(_logo_path, width=120)
    with hdr_cols[1]:
        st.markdown(f"""
        <div style='padding: 0.2rem 0;'>
            <h1 style='color: var(--ink); margin: 0; font-size: 1.6rem; font-weight: 800; letter-spacing: -0.3px;'>{APP_NAME}</h1>
            <p style='color: var(--muted); font-size: 13px; margin: 0;'>
                Professional market condition adjustments aligned with Fannie Mae guidance
            </p>
        </div>
        """, unsafe_allow_html=True)
    with hdr_cols[2]:
        st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
        if st.button("↺ Start Over", use_container_width=True, key="start_over"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div style="border-bottom:1px solid var(--border); margin: 4px 0 18px 0;"></div>', unsafe_allow_html=True)
    
    # Progress indicator
    show_progress_indicator(st.session_state["step"])
    
    # ----------------------------
    # STEP 1: Subject Property
    # ----------------------------
    if st.session_state["step"] == 1:
        st.subheader("Enter Subject Property Information")
        
        subject_address = st.text_input(
            "Subject Property Address",
            value=st.session_state["subject_address"],
            placeholder="e.g., 123 Main Street, City, State",
            help="Enter the address of the property being appraised"
        )
        
        eff_date = st.date_input(
            "Effective Date",
            value=st.session_state["eff_date"],
            help="Date of the appraisal",
            format="MM/DD/YYYY"
        )
        
        st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns([3, 1, 1])
        with col3:
            if st.button("Next →", type="primary", use_container_width=True, key="step1_next"):
                if not subject_address:
                    st.error("Please enter a subject property address")
                else:
                    st.session_state["subject_address"] = subject_address
                    st.session_state["eff_date"] = eff_date
                    st.session_state["step"] = 2
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        # ----------------------------
        # Report History
        # ----------------------------
        history = load_history()
        if history:
            st.markdown('<div style="margin-top: 40px; border-top: 1px solid var(--border); padding-top: 20px;"></div>', unsafe_allow_html=True)
            st.markdown("#### 📋 Recent Reports")
            st.caption("Click to reopen a previous report")

            for i, report in enumerate(history):
                addr = report.get("subject_address", "Unknown")
                eff = report.get("eff_date", "")
                saved = report.get("saved_at", "")
                report_id = report.get("id", str(i))
                n_comps = len(report.get("selected_comps", []))

                # Format the saved date nicely
                try:
                    saved_dt = datetime.fromisoformat(saved)
                    saved_str = saved_dt.strftime("%b %d, %Y at %I:%M %p")
                except (ValueError, TypeError):
                    saved_str = saved

                # Format effective date
                try:
                    eff_dt = date.fromisoformat(eff)
                    eff_str = eff_dt.strftime("%B %d, %Y")
                except (ValueError, TypeError):
                    eff_str = eff

                col_info, col_open, col_del = st.columns([5, 1, 0.5])
                with col_info:
                    st.markdown(f"""
                    <div style="background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:12px 16px; margin-bottom:4px;">
                        <div style="font-size:14px; font-weight:700; color:var(--ink);">{addr}</div>
                        <div style="font-size:11px; color:var(--muted); margin-top:2px;">
                            Effective: {eff_str} · {n_comps} comp{'s' if n_comps != 1 else ''} · Saved {saved_str}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                with col_open:
                    if st.button("Open", key=f"hist_open_{report_id}", use_container_width=True, type="primary"):
                        load_report_from_history(report, st.session_state)
                        st.rerun()
                with col_del:
                    if st.button("✕", key=f"hist_del_{report_id}", use_container_width=True):
                        delete_report_from_history(report_id)
                        st.rerun()
    # ----------------------------
    # ----------------------------
    # STEP 2: Upload Data
    # ----------------------------
    elif st.session_state["step"] == 2:
        st.subheader("Upload Market Data")
        
        st.markdown("""
        Upload a CSV file containing comparable sales data with the following columns:
        - **Address** - Property address
        - **Zip** - Zip code
        - **Pending Date** - Contract/pending date
        - **Sold Price** - Sale price
        """)
        
        uploaded_file = st.file_uploader(
            "Choose CSV file",
            type=["csv"],
            help="Upload your market data CSV file"
        )
        
        if uploaded_file is not None:
            try:
                df_raw = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                df = canonicalize_columns(normalize_columns(df_raw))
                
                missing = [col for col in REQUIRED_COLS if col not in df.columns]
                if missing:
                    st.error(f"Missing required columns: {', '.join(missing)}")
                    st.info(f"Available columns: {', '.join(df.columns.tolist())}")
                else:
                    df["ContractDate"] = parse_dates_robust(df["Pending Date"])
                    df["SoldPrice"] = parse_money_robust(df["Sold Price"])
                    df_clean = df.dropna(subset=["ContractDate", "SoldPrice"]).copy()
                    df_clean["ContractDate"] = pd.to_datetime(df_clean["ContractDate"]).dt.date

                    # Create a stable RowID once at upload time (do not rebuild later)
                    df_clean = df_clean.reset_index(drop=True)
                    df_clean["RowID"] = df_clean.index.astype(int)
                    
                    if df_clean.empty:
                        st.error("No valid data after parsing dates and prices")
                    else:
                        st.success(f"✓ Successfully loaded {len(df_clean)} sales records")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            vq_stat_card("Total Sales", f"{len(df_clean):,}")
                        with col2:
                            avg_price = df_clean["SoldPrice"].mean()
                            vq_stat_card("Average Price", f"${avg_price:,.0f}")
                        with col3:
                            date_start = df_clean["ContractDate"].min()
                            date_end = df_clean["ContractDate"].max()
                            months = (date_end.year - date_start.year) * 12 + (date_end.month - date_start.month)
                            vq_stat_card("Time Period", f"{months} months")
                        with col4:
                            vq_stat_card("Date Range", f"{date_start.strftime('%b %Y')} – {date_end.strftime('%b %Y')}")
                        
                        st.session_state["uploaded_data"] = df_clean
                        
                        st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
                        col1, col2, col3 = st.columns([2, 1, 1])
                        with col2:
                            if st.button("← Back", use_container_width=True, key="step2_back"):
                                st.session_state["step"] = 1
                                st.rerun()
                        with col3:
                            if st.button("Next →", type="primary", use_container_width=True, key="step2_next"):
                                st.session_state["step"] = 3
                                st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error reading CSV: {e}")
        else:
            st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([2, 1, 1])
            with col2:
                if st.button("← Back", use_container_width=True, key="step2_back_nofile"):
                    st.session_state["step"] = 1
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    
    # ----------------------------
    # STEP 3: Data Diagnostics
    # ----------------------------
    elif st.session_state["step"] == 3:
        df_f = st.session_state["uploaded_data"].copy()
        settings = st.session_state["settings"]
        
        st.subheader("Data Diagnostics & Outlier Removal")
        
        # Settings in sidebar-like expander
        with st.expander("⚙️ Diagnostic Settings", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                use_iqr = st.checkbox("IQR Outlier Detection", value=settings["use_iqr"])
                if use_iqr:
                    iqr_multiplier = st.slider("IQR Multiplier", 1.0, 3.0, float(settings["iqr_multiplier"]), 0.1)
                else:
                    iqr_multiplier = 1.0
            with col2:
                use_cooks = st.checkbox("Cook's Distance", value=settings["use_cooks"])
            with col3:
                show_scatter = st.checkbox("Show Interactive Scatter Plot", value=True)
            
            settings["use_iqr"] = use_iqr
            settings["iqr_multiplier"] = iqr_multiplier
            settings["use_cooks"] = use_cooks
        
        # Run diagnostics
        if use_iqr:
            iqr_flags = compute_iqr_flags_cached(df_f["SoldPrice"].to_numpy(), iqr_multiplier)
            df_f["IQR_Outlier"] = iqr_flags
        else:
            df_f["IQR_Outlier"] = False
        
        if use_cooks:
            df_f = cooks_distance_time_regression(df_f)
        else:
            df_f["HighLeverage"] = False
            df_f["HighCooksD"] = False
        
        df_f["Flagged"] = df_f["IQR_Outlier"] | df_f["HighLeverage"] | df_f["HighCooksD"]
        
        # Human-readable flag reasons (audit trail)
        def _flag_reason(r):
            reasons = []
            if bool(r.get("IQR_Outlier", False)):
                reasons.append("IQR")
            if bool(r.get("HighLeverage", False)):
                reasons.append("Price Dev")
            if bool(r.get("HighCooksD", False)):
                reasons.append("Cook's D")
            return " + ".join(reasons)
        df_f["FlagReason"] = df_f.apply(_flag_reason, axis=1)
        flagged_count = df_f["Flagged"].sum()
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"🏴 {flagged_count} records flagged for review")
        with col2:
            st.success(f"✅ {len(df_f) - flagged_count} records passed diagnostics")
        
        # Main content area with exclusion sidebar
        col_main, col_sidebar = st.columns([3, 1])
        
        with col_main:
            # Interactive scatter plot with click-to-exclude
            if show_scatter:
                st.markdown("#### Interactive Data Visualization")
                st.caption("Click on points in the chart to exclude them, or use the table below")
                
                plot_df = df_f.copy()
                plot_df["ContractDateDT"] = pd.to_datetime(plot_df["ContractDate"])
                plot_df["Y"] = plot_df["SoldPrice"].astype(float)
                # Formatted columns for clean hover
                plot_df["PriceFormatted"] = plot_df["SoldPrice"].apply(lambda x: f"${x:,.0f}")
                plot_df["DateFormatted"] = plot_df["ContractDateDT"].dt.strftime("%B %d, %Y")
                
                plot_df["Status"] = np.where(
                    plot_df["RowID"].isin(st.session_state["excluded_rowids"]),
                    "Excluded",
                    np.where(plot_df["Flagged"], "Flagged", "Included")
                )
                
                fig_scatter = px.scatter(
                    plot_df,
                    x="ContractDateDT",
                    y="Y",
                    color="Status",
                    color_discrete_map={
                        "Included": THEME_COLORS['success'],
                        "Flagged": THEME_COLORS['warning'],
                        "Excluded": '#B0B8C4'
                    },
                    symbol="Status",
                    symbol_map={
                        "Included": "circle",
                        "Flagged": "diamond",
                        "Excluded": "x"
                    },
                    opacity=0.85,
                    custom_data=["RowID", "Address", "PriceFormatted", "DateFormatted"],
                    title="Sale Price Distribution Over Time (Click points to exclude)",
                )
                # Custom hover template — clean and simple
                for trace in fig_scatter.data:
                    trace.hovertemplate = (
                        "<b>%{customdata[1]}</b><br>"
                        "Sale Price: %{customdata[2]}<br>"
                        "Contract Date: %{customdata[3]}"
                        "<extra></extra>"
                    )
                # Dim excluded points further
                for trace in fig_scatter.data:
                    if trace.name == "Excluded":
                        trace.marker.opacity = 0.4
                        trace.marker.size = 7

                # Add trend-following IQR outlier bounds
                if use_iqr:
                    import plotly.graph_objects as go
                    _prices_s = df_f[["ContractDate", "SoldPrice"]].copy()
                    _prices_s["ContractDate"] = pd.to_datetime(_prices_s["ContractDate"])
                    _prices_s = _prices_s.sort_values("ContractDate")
                    _prices_s["SoldPrice"] = _prices_s["SoldPrice"].astype(float)

                    # Rolling window: ~3 months of data or min 8 points
                    _win = max(8, len(_prices_s) // 5)
                    _roll_q1 = _prices_s["SoldPrice"].rolling(_win, center=True, min_periods=3).quantile(0.25)
                    _roll_q3 = _prices_s["SoldPrice"].rolling(_win, center=True, min_periods=3).quantile(0.75)
                    _roll_iqr = _roll_q3 - _roll_q1
                    _roll_hi = _roll_q3 + iqr_multiplier * _roll_iqr
                    _roll_lo = _roll_q1 - iqr_multiplier * _roll_iqr

                    # Fill NaN edges
                    _roll_hi = _roll_hi.bfill().ffill()
                    _roll_lo = _roll_lo.bfill().ffill()

                    # Second smoothing pass to eliminate jumpiness
                    _smooth_win = max(5, _win // 2)
                    _roll_hi = _roll_hi.rolling(_smooth_win, center=True, min_periods=1).mean()
                    _roll_lo = _roll_lo.rolling(_smooth_win, center=True, min_periods=1).mean()

                    fig_scatter.add_trace(go.Scatter(
                        x=_prices_s["ContractDate"], y=_roll_hi,
                        mode='lines', name='Upper Bound',
                        line=dict(color='rgba(255,59,48,0.3)', width=1.5, shape='spline'),
                        hoverinfo='skip', showlegend=True
                    ))
                    fig_scatter.add_trace(go.Scatter(
                        x=_prices_s["ContractDate"], y=_roll_lo,
                        mode='lines', name='Lower Bound',
                        line=dict(color='rgba(255,59,48,0.3)', width=1.5, shape='spline'),
                        hoverinfo='skip', showlegend=True
                    ))
                fig_scatter.update_layout(
                    xaxis_title="Contract Date",
                    yaxis_title="Sale Price",
                    height=500,
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(color=THEME_COLORS["text"]),
                    title_font=dict(color=THEME_COLORS["text"]),
                    xaxis=dict(
                        tickfont=dict(color=THEME_COLORS["text"]),
                        title_font=dict(color=THEME_COLORS["text"]),
                    ),
                    yaxis=dict(
                        tickfont=dict(color=THEME_COLORS["text"]),
                        title_font=dict(color=THEME_COLORS["text"]),
                    ),
                    legend=dict(font=dict(color=THEME_COLORS["text"])),
                    hovermode="closest",
                )
                fig_scatter.update_yaxes(tickformat="$,.0f")
                
                # Display the plot and capture click events
                selected_points = st.plotly_chart(fig_scatter, use_container_width=True, 
                                                  on_select="rerun", selection_mode="points")
                
                # Handle point selection — use customdata[0] which holds RowID
                if selected_points and selected_points.selection and selected_points.selection.points:
                    for point in selected_points.selection.points:
                        # customdata is a list; RowID is at index 0
                        custom = point.get('customdata')
                        if custom and len(custom) > 0:
                            row_id = int(custom[0])
                        else:
                            # Fallback: skip this point
                            continue
                        # Toggle exclusion
                        if row_id in st.session_state["excluded_rowids"]:
                            st.session_state["excluded_rowids"].remove(row_id)
                        else:
                            st.session_state["excluded_rowids"].add(row_id)
                    st.rerun()
            
            # Data editor
            st.markdown("#### Review and Edit Exclusions")
            
            df_diag = df_f.copy()
            df_diag["SalePrice"] = df_diag["SoldPrice"]
            df_diag["Exclude"] = df_diag["RowID"].isin(st.session_state["excluded_rowids"])
            
            filter_opt = st.radio(
                "Show:",
                ["All Records", "Flagged Only", "Excluded Only"],
                horizontal=True
            )
            
            if filter_opt == "Flagged Only":
                df_diag = df_diag[df_diag["Flagged"]].copy()
            elif filter_opt == "Excluded Only":
                df_diag = df_diag[df_diag["Exclude"]].copy()
            
            display_cols = [
                "RowID", "Exclude", "Address", "ContractDate", "SalePrice",
                "FlagReason",
            ]
            # Only include diagnostic flag columns that exist
            for extra_col in ["IQR_Outlier", "HighLeverage", "HighCooksD"]:
                if extra_col in df_diag.columns:
                    display_cols.append(extra_col)
            
            edited = st.data_editor(
                df_diag[[c for c in display_cols if c in df_diag.columns]],
                column_config={
                    "RowID": st.column_config.NumberColumn("ID", width=60),
                    "Exclude": st.column_config.CheckboxColumn("Exclude", width=75),
                    "Address": st.column_config.TextColumn("Address", width=220),
                    "ContractDate": st.column_config.DateColumn("Contract Date", width=110),
                    "SalePrice": st.column_config.NumberColumn("Sale Price", format="$%d", width=100),
                    "FlagReason": st.column_config.TextColumn("Flag", width=90),
                    "IQR_Outlier": st.column_config.CheckboxColumn("IQR", width=55),
                    "HighLeverage": st.column_config.CheckboxColumn("Lev.", width=55),
                    "HighCooksD": st.column_config.CheckboxColumn("Cook", width=55),
                },
                hide_index=True,
                use_container_width=True,
                height=450,
                disabled=["RowID", "Address", "ContractDate", "SalePrice", "FlagReason", "IQR_Outlier", "HighLeverage", "HighCooksD"]
            )
            
            # Track exclusions from table
            pending_excluded = set()
            for rid, ex in zip(edited["RowID"].tolist(), edited["Exclude"].tolist()):
                if ex:
                    pending_excluded.add(rid)
        with col_sidebar:
            # Exclusion list sidebar
            st.markdown("#### Excluded Properties")

            # "Remove All Outliers" button — excludes all currently flagged
            flagged_ids = set(df_f[df_f["Flagged"]]["RowID"].tolist())
            unflagged_remaining = flagged_ids - st.session_state["excluded_rowids"]
            if unflagged_remaining:
                if st.button(f"🚫 Remove All Outliers ({len(unflagged_remaining)})", use_container_width=True, type="primary", key="remove_all_outliers"):
                    st.session_state["excluded_rowids"] = st.session_state["excluded_rowids"] | flagged_ids
                    st.rerun()
            
            if st.session_state["excluded_rowids"]:
                # Create a list of excluded properties
                excluded_df = df_f[df_f["RowID"].isin(st.session_state["excluded_rowids"])].copy()
                
                st.caption(f"{len(excluded_df)} properties excluded")
                
                # Show each excluded property with inline X button
                for idx, row in excluded_df.iterrows():
                    col_text, col_btn = st.columns([4, 1])
                    
                    with col_text:
                        st.markdown(f"""
                        <div style='background: #FEF2F2; padding: 8px 12px; border-radius: 6px; margin-bottom: 6px; border-left: 3px solid #FF3B30;'>
                            <div style='font-size: 11px; font-weight: 600; color: #111827;'>{row['Address'][:30]}{'...' if len(row['Address']) > 30 else ''}</div>
                            <div style='font-size: 10px; color: #6B7280;'>${row['SoldPrice']:,.0f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col_btn:
                        if st.button("✕", key=f"remove_{row['RowID']}", help="Restore this property",
                                   use_container_width=True):
                            st.session_state["excluded_rowids"].remove(row["RowID"])
                            st.rerun()
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Clear All", use_container_width=True, help="Restore all excluded properties"):
                    st.session_state["excluded_rowids"] = set()
                    st.rerun()
            else:
                st.info("No properties excluded yet")
                st.caption("Click points on the chart or check boxes in the table to exclude properties")
        
        # Navigation buttons
        st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
        with col3:
            if st.button("← Back", use_container_width=True, key="step3_back"):
                st.session_state["step"] = 2
                st.rerun()
        with col4:
            if st.button("Apply & Next →", type="primary", use_container_width=True, key="step3_next"):
                st.session_state["excluded_rowids"] = pending_excluded
                
                # Persist diagnostics output for the report pack
                diag_export = df_f.copy()
                diag_export["Excluded"] = diag_export["RowID"].isin(st.session_state["excluded_rowids"])
                st.session_state["diagnostics_df"] = diag_export
                cooks_threshold_val = settings.get("cooks_threshold", None)
                if cooks_threshold_val is None or (isinstance(cooks_threshold_val, float) and np.isnan(cooks_threshold_val)):
                    if bool(settings.get("use_cooks", False)) and "CooksD" in diag_export.columns:
                        n_cooks = int(diag_export["CooksD"].notna().sum())
                        cooks_threshold_val = 4 / max(1, n_cooks)
                    else:
                        cooks_threshold_val = float("nan")

                st.session_state["diagnostics_settings"] = {
                    "iqr_multiplier": float(settings.get("iqr_multiplier", 1.0)),
                    "use_cooks": bool(settings.get("use_cooks", False)),
                    "cooks_threshold": float(cooks_threshold_val),
                    "min_sales_per_month": int(settings.get("min_sales_per_month", 5)),
                    "smooth_window": int(settings.get("smooth_window", 6)),
                }
                
                st.session_state["step"] = 4
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.caption(f"Currently excluding {len(st.session_state['excluded_rowids'])} record(s)")
    
    # ----------------------------
    # STEP 4: Select Comparables
    # ----------------------------
    elif st.session_state["step"] == 4:
        df_model = st.session_state["uploaded_data"].copy()
        df_model = df_model[~df_model["RowID"].isin(st.session_state["excluded_rowids"])].copy()
        
        if df_model.empty:
            st.error("All records excluded. Go back and uncheck some exclusions.")
            if st.button("← Back"):
                st.session_state["step"] = 3
                st.rerun()
            st.stop()
        
        st.subheader("Select Comparable Sales")

        # Analysis settings are configured on the Report step (so you can adjust the chart while viewing it).
        settings = st.session_state["settings"]
        index_option = settings.get("index_option", "Smoothed (Recommended)")
        smooth_window = int(settings.get("smooth_window", 6))
        min_sales_per_month = int(settings.get("min_sales_per_month", 3))

        # Build index
        index_df = build_index_cached(
            df_model["ContractDate"].to_numpy(),
            df_model["SoldPrice"].to_numpy(),
            int(min_sales_per_month),
            int(smooth_window)
        )

        # Defensive warning for short time series
        if len(index_df) < 12 and not index_df.empty:
            st.warning(
                f"Limited time series: {len(index_df)} month(s) of data. "
                "Short datasets can produce unstable month-to-month indices. "
                "Consider expanding search criteria or widening the date range if possible."
            )

        if index_df.empty:
            st.error("Unable to build market index from filtered data")
            st.stop()
        
        # Select index column
        if "Raw" in index_option:
            math_col = "Index_Raw"
        elif "Regression" in index_option:
            math_col = "Index_Regression"
        else:
            math_col = "Index_Smoothed"
        
        eff_date = st.session_state["eff_date"]
        eff_index, eff_month_used, eff_mode = lookup_index(index_df, eff_date, math_col)
        
        # Overall trend
        if len(index_df) >= 2:
            first_idx = index_df.iloc[0][math_col]
            last_idx = index_df.iloc[-1][math_col]
            overall_change_pct = ((last_idx / first_idx) - 1.0) * 100
            overall_trend = categorize_adjustment(overall_change_pct, threshold=2.0)
        else:
            overall_change_pct = 0.0
            overall_trend = "Stable"
        
        col1, col2, col3 = st.columns(3)
        with col1:
            vq_stat_card("Overall Market Change", f"{overall_change_pct:+.2f}%", delta=overall_trend)
        with col2:
            vq_stat_card("Index at Effective Date", f"{eff_index:.4f}", delta=f"{((eff_index - 1.0) * 100):+.2f}%")
        with col3:
            vq_stat_card("Data Quality", f"{len(df_model)} sales", delta=f"{len(index_df)} months")
        
        st.markdown("#### Choose Comparable Sales")

        df_pick = df_model.copy()
        df_pick["Label"] = (
            df_pick["Address"].astype(str) +
            " | " + df_pick["ContractDate"].astype(str) +
            " | " + df_pick["SoldPrice"].apply(lambda x: f"${x:,.0f}")
        )

        sort_mode = st.selectbox(
            "Sort comparables by:",
            ["Contract Date (Newest First)", "Contract Date (Oldest First)", "Closest to Effective Date", "Sale Price (Low to High)", "Sale Price (High to Low)"],
            index=0
        )
        if sort_mode == "Contract Date (Newest First)":
            df_pick = df_pick.sort_values(["ContractDate", "SoldPrice"], ascending=[False, True])
        elif sort_mode == "Contract Date (Oldest First)":
            df_pick = df_pick.sort_values(["ContractDate", "SoldPrice"], ascending=[True, True])
        elif sort_mode == "Closest to Effective Date":
            eff_date = st.session_state["eff_date"]
            df_pick["_abs_days"] = df_pick["ContractDate"].apply(lambda d: abs(days_between(d, eff_date)))
            df_pick = df_pick.sort_values(["_abs_days", "ContractDate"], ascending=[True, False]).drop(columns=["_abs_days"])
        elif sort_mode == "Sale Price (Low to High)":
            df_pick = df_pick.sort_values(["SoldPrice", "ContractDate"], ascending=[True, False])
        else:
            df_pick = df_pick.sort_values(["SoldPrice", "ContractDate"], ascending=[False, False])

        label_map = dict(zip(df_pick["RowID"].tolist(), df_pick["Label"].tolist()))

        selected = st.multiselect(
            "Select comparables for adjustment analysis:",
            options=df_pick["RowID"].tolist(),
            default=st.session_state["selected_comps"],
            format_func=lambda rid: label_map.get(rid, str(rid)),
            help="Choose the sales you want to analyze",
            key="comp_selector"
        )
        
        if selected:
            st.session_state["selected_comps"] = selected
            
            st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([2, 1, 1])
            with col2:
                if st.button("← Back", use_container_width=True, key="step4_back"):
                    st.session_state["step"] = 3
                    st.rerun()
            with col3:
                if st.button("Generate Report →", type="primary", use_container_width=True, key="step4_next"):
                    st.session_state["step"] = 5
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("👆 Select at least one comparable sale to continue")
            st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([2, 1, 1])
            with col2:
                if st.button("← Back", use_container_width=True, key="step4_back_nosel"):
                    st.session_state["step"] = 3
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    
    # ----------------------------
    # STEP 5: Adjustment Report
    # ----------------------------
    elif st.session_state["step"] == 5:
        settings = st.session_state["settings"]

        # Auto-save this report to history
        try:
            save_report_to_history(st.session_state)
        except Exception:
            pass  # Don't break the app if history save fails

        # Widen the page for the report step
        st.markdown('<style>.block-container { max-width: 1520px !important; }</style>', unsafe_allow_html=True)

        render_step_header(
            "Market Condition Adjustment Report",
            "Tune analysis settings while viewing the results. Export a report pack when ready.",
            icon="📄"
        )

        # ----------------------------
        # Read live widget values from session state keys.
        # On step 5, the analysis widgets (rep_smooth, rep_min_sales, etc.)
        # write to session_state with their key. On subsequent reruns their
        # values are available BEFORE the widgets render. We sync them into
        # the settings dict first so the index build uses fresh values.
        # ----------------------------
        if "rep_smooth" in st.session_state:
            settings["smooth_window"] = int(st.session_state["rep_smooth"])
        if "rep_min_sales" in st.session_state:
            settings["min_sales_per_month"] = int(st.session_state["rep_min_sales"])
        if "rep_index_option" in st.session_state:
            settings["index_option"] = st.session_state["rep_index_option"]
        if "rep_no_adj_days" in st.session_state:
            settings["no_adj_days"] = int(st.session_state["rep_no_adj_days"])
        if "rep_trend_lookback" in st.session_state:
            settings["trend_lookback"] = st.session_state["rep_trend_lookback"]
        if "rep_trend_override" in st.session_state:
            settings["trend_override"] = st.session_state["rep_trend_override"]

        # ----------------------------
        # Build model + index
        # ----------------------------
        df_model = st.session_state["uploaded_data"].copy()
        df_model = df_model[~df_model["RowID"].isin(st.session_state["excluded_rowids"])].copy()

        index_df = build_index_cached(
            df_model["ContractDate"].to_numpy(),
            df_model["SoldPrice"].to_numpy(),
            int(settings["min_sales_per_month"]),
            int(settings["smooth_window"])
        )

        if "Raw" in settings["index_option"]:
            math_col = "Index_Raw"
        elif "Regression" in settings["index_option"]:
            math_col = "Index_Regression"
        else:
            math_col = "Index_Smoothed"

        eff_date = st.session_state["eff_date"]
        eff_index, eff_month_used, eff_mode = lookup_index(index_df, eff_date, math_col)

        # Overall trend — calculated over the lookback window, not full dataset
        if len(index_df) >= 2:
            # Determine lookback window
            lookback_str = settings.get("trend_lookback", "1 Year")
            lookback_months = {"6 Months": 6, "1 Year": 12, "1.5 Years": 18, "2 Years": 24}.get(lookback_str, 12)
            lookback_start = pd.to_datetime(eff_date) - pd.DateOffset(months=lookback_months)

            # Filter index to lookback window
            index_df_lb = index_df[pd.to_datetime(index_df['Month']) >= lookback_start]
            if len(index_df_lb) >= 2:
                first_idx = float(index_df_lb.iloc[0][math_col])
                last_idx = float(index_df_lb.iloc[-1][math_col])
            else:
                first_idx = float(index_df.iloc[0][math_col])
                last_idx = float(index_df.iloc[-1][math_col])

            overall_change_pct = ((last_idx / first_idx) - 1.0) * 100.0
            auto_trend = categorize_adjustment(overall_change_pct, threshold=2.0)

            # Allow manual override
            trend_override = settings.get("trend_override", "Auto-detect")
            if trend_override != "Auto-detect":
                overall_trend = trend_override
            else:
                overall_trend = auto_trend
        else:
            overall_change_pct = 0.0
            overall_trend = "Stable"

        # ----------------------------
        # Select comps + compute adjustments
        # ----------------------------
        df_pick = df_model.copy()
        comps = df_pick[df_pick["RowID"].isin(st.session_state["selected_comps"])].copy()

        if comps.empty:
            st.warning("No comparables are selected. Go back to Step 4 and select at least one comparable.")
        else:
            comps["Index_Contract"] = comps["ContractDate"].apply(lambda x: lookup_index(index_df, x, math_col)[0])
            comps["Index_Effective"] = eff_index
            comps["DaysFromEffective"] = comps["ContractDate"].apply(lambda d: days_between(d, eff_date))
            comps["AppliedAdj"] = comps["DaysFromEffective"] >= int(settings["no_adj_days"])

            comps["MktAdjPct"] = comps.apply(
                lambda r: pct_change(r["Index_Effective"], r["Index_Contract"]) * 100.0,
                axis=1
            )
            comps["MktAdj$"] = comps["SoldPrice"] * (comps["MktAdjPct"] / 100.0)

            comps.loc[~comps["AppliedAdj"], "MktAdjPct"] = 0.0
            comps.loc[~comps["AppliedAdj"], "MktAdj$"] = 0.0

            comps["Category"] = comps["MktAdjPct"].apply(categorize_adjustment)
            comps["Direction"] = comps["MktAdjPct"].apply(adjustment_direction)

            out = comps[[
                "Address", "ContractDate", "SoldPrice",
                "Index_Contract", "Index_Effective",
                "MktAdjPct", "MktAdj$", "Category", "Direction",
                "AppliedAdj", "DaysFromEffective"
            ]].copy().rename(columns={"Address": "CompAddress", "SoldPrice": "SalePrice"})

            out = out.sort_values("ContractDate").reset_index(drop=True)

            # Narrative inputs
            date_start = df_model["ContractDate"].min()
            date_end = df_model["ContractDate"].max()

            narrative = build_narrative(
                subject_address=st.session_state["subject_address"],
                date_start=date_start,
                date_end=date_end,
                eff_date=eff_date,
                eff_index=eff_index,
                index_option=settings["index_option"],
                math_col=math_col,
                comp_rows=out,
                index_df=index_df,
                excluded_count=len(st.session_state["excluded_rowids"]),
                overall_trend=overall_trend,
                overall_change_pct=overall_change_pct,
                trend_lookback=settings.get("trend_lookback", "1 Year"),
                raw_sales_df=df_model,
            )


            # ----------------------------
            # FULL-WIDTH HERO: Subject Info
            # ----------------------------
            st.markdown(f'''
            <div class="vq-hero">
                <div class="vq-hero-title">\U0001f4cd {_escape(st.session_state["subject_address"])}</div>
                <div class="vq-hero-sub">Effective Date: {eff_date.strftime("%B %d, %Y")}  \u00b7  Trend: {overall_trend}  \u00b7  Overall Change: {overall_change_pct:+.2f}%</div>
            </div>
            ''', unsafe_allow_html=True)

            # ----------------------------
            # TWO-COLUMN: Inspector + Chart/Results
            # ----------------------------
            left_col, right_col = st.columns([0.24, 0.76], gap="large")

            # LEFT: Inspector (sticky sidebar)
            with left_col:
                st.markdown('<div class="vq-sticky">', unsafe_allow_html=True)

                with st.expander("\u2699 Analysis", expanded=True):
                    index_option = st.selectbox(
                        "Index Type",
                        ["Smoothed (Recommended)", "Raw Chain-Linked", "Regression Trendline"],
                        index=["Smoothed (Recommended)", "Raw Chain-Linked", "Regression Trendline"].index(
                            st.session_state["settings"].get("index_option", "Smoothed (Recommended)")
                        ),
                        key="rep_index_option",
                    )
                    st.session_state["settings"]["index_option"] = index_option

                    smooth_window = st.slider(
                        "Smoothing Window",
                        2, 12,
                        int(st.session_state["settings"].get("smooth_window", 6)),
                        help="Rolling average window (months)",
                        key="rep_smooth",
                    )
                    st.session_state["settings"]["smooth_window"] = int(smooth_window)

                    min_sales_per_month = st.number_input(
                        "Min Sales/Month",
                        1, 50,
                        int(st.session_state["settings"].get("min_sales_per_month", 3)),
                        key="rep_min_sales",
                    )
                    st.session_state["settings"]["min_sales_per_month"] = int(min_sales_per_month)

                    no_adj_days = st.number_input(
                        "No-Adj Window (days)",
                        0, 365,
                        int(st.session_state["settings"].get("no_adj_days", 0)),
                        key="rep_no_adj_days",
                    )
                    st.session_state["settings"]["no_adj_days"] = int(no_adj_days)

                with st.expander("\U0001f4c8 Trend Assessment", expanded=True):
                    trend_lookback = st.selectbox(
                        "Lookback Period",
                        ["6 Months", "1 Year", "1.5 Years", "2 Years"],
                        index=["6 Months", "1 Year", "1.5 Years", "2 Years"].index(
                            st.session_state["settings"].get("trend_lookback", "1 Year")
                        ),
                        help="Period for calculating the market trend direction",
                        key="rep_trend_lookback",
                    )
                    st.session_state["settings"]["trend_lookback"] = trend_lookback

                    trend_override = st.selectbox(
                        "Market Trend (appraiser opinion)",
                        ["Auto-detect", "Declining", "Stable", "Increasing"],
                        index=["Auto-detect", "Declining", "Stable", "Increasing"].index(
                            st.session_state["settings"].get("trend_override", "Auto-detect")
                        ),
                        help="Override the calculated trend with your professional judgment",
                        key="rep_trend_override",
                    )
                    st.session_state["settings"]["trend_override"] = trend_override

                with st.expander("\U0001f4ca Chart", expanded=False):
                    show_raw_overlay = st.checkbox("Emphasize Raw Overlay", value=False, key="rep_show_raw")
                    show_thin_months = st.checkbox("Mark Thin Months", value=True, key="rep_show_thin")
                    tick_mode = st.selectbox("X-axis", ["Monthly", "Quarterly"], key="rep_tick_mode")

                # Build chart
                # Compute lookback months for chart x-axis
                _lb_str = settings.get("trend_lookback", "1 Year")
                _lb_months = {"6 Months": 6, "1 Year": 12, "1.5 Years": 18, "2 Years": 24}.get(_lb_str, 12)

                fig = plot_fannie_style_chart(
                    index_df=index_df, comps_df=out,
                    eff_date=eff_date, eff_index=eff_index,
                    index_col=math_col,
                    show_raw=show_raw_overlay,
                    show_thin=show_thin_months,
                    tick_mode=tick_mode,
                    lookback_months=_lb_months
                )

                png_buf = io.BytesIO()
                fig.savefig(png_buf, format="png", bbox_inches="tight", dpi=200, facecolor='white')
                png_bytes = png_buf.getvalue()

                csv_out = out.copy()
                csv_out["ContractDate"] = csv_out["ContractDate"].astype(str)
                csv_data = csv_out.to_csv(index=False).encode("utf-8")
                txt_data = narrative.encode("utf-8")

                # Clean address for filenames
                import re
                _addr = st.session_state.get("subject_address", "Report")
                _addr_clean = re.sub(r'[^\w\s-]', '', _addr).strip().replace('  ', ' ')
                _fn_prefix = f"{_addr_clean} MarketAdjuster"

                with st.expander("\U0001f4e5 Exports", expanded=False):
                    st.caption("**Individual Exports**")
                    st.download_button("Chart (PNG)", data=png_bytes, file_name=f"{_fn_prefix} Chart.png", mime="image/png", use_container_width=True)

                    # Render adjustment table as image
                    table_img_bytes = render_table_image(
                        comp_data=out,
                        subject_address=st.session_state["subject_address"],
                        eff_date=eff_date,
                        overall_trend=overall_trend,
                        overall_change_pct=overall_change_pct,
                    )
                    st.download_button("Adjustments Table (PNG)", data=table_img_bytes, file_name=f"{_fn_prefix} Adjustments.png", mime="image/png", use_container_width=True)
                    st.download_button("Narrative (TXT)", data=txt_data, file_name=f"{_fn_prefix} Narrative.txt", mime="text/plain", use_container_width=True)
                    st.download_button("Raw Data (CSV)", data=csv_data, file_name=f"{_fn_prefix} Data.csv", mime="text/csv", use_container_width=True)

                    st.markdown("---")
                    st.caption("**PDF Addendum**")
                    pdf_include_table = st.checkbox("Include table", value=True, key="pdf_include_table")
                    pdf_include_narr = st.checkbox("Include narrative", value=True, key="pdf_include_narr")
                    if st.button("Generate PDF", type="primary", use_container_width=True):
                        pdf_table = out.copy()
                        pdf_table.insert(0, "Comp #", range(1, len(pdf_table) + 1))
                        pdf_table_disp = pd.DataFrame({
                            "Comp #": pdf_table["Comp #"],
                            "Address": pdf_table["CompAddress"],
                            "Contract": pdf_table["ContractDate"].astype(str),
                            "Sale": pdf_table["SalePrice"].apply(lambda x: f"${x:,.0f}"),
                            "Adj %": pdf_table["MktAdjPct"].apply(lambda x: f"{x:+.2f}%"),
                            "Adj $": pdf_table["MktAdj$"].apply(lambda x: f"${x:+,.0f}"),
                        })
                        pdf_bytes = build_pdf_addendum(
                            subject_address=st.session_state["subject_address"],
                            eff_date=eff_date, settings=settings,
                            narrative=(narrative if pdf_include_narr else ""),
                            chart_png=png_bytes,
                            comp_table=(pdf_table_disp if pdf_include_table else pdf_table_disp.head(0)),
                        )
                        st.download_button("Download PDF", data=pdf_bytes, file_name=f"{_fn_prefix} Report.pdf", mime="application/pdf", use_container_width=True)

                # ZIP pack
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr(f"{_fn_prefix} Chart.png", png_bytes)
                    zf.writestr(f"{_fn_prefix} Adjustments.png", table_img_bytes)
                    zf.writestr(f"{_fn_prefix} Data.csv", csv_data)
                    zf.writestr(f"{_fn_prefix} Narrative.txt", txt_data)
                    if "diagnostics_df" in st.session_state and isinstance(st.session_state["diagnostics_df"], pd.DataFrame):
                        diag_df = st.session_state["diagnostics_df"].copy()
                        pref_cols = ["RowID", "Excluded", "Flagged", "FlagReason", "Address", "ContractDate", "SoldPrice", "IQR_Outlier", "HighLeverage", "HighCooksD"]
                        diag_cols_available = [c for c in pref_cols if c in diag_df.columns]
                        zf.writestr(f"{_fn_prefix} Diagnostics.csv", diag_df[diag_cols_available].to_csv(index=False).encode("utf-8"))
                    zf.writestr(f"{_fn_prefix} Settings.json", json.dumps(st.session_state.get("diagnostics_settings", {}), indent=2, default=str))

                st.download_button("\U0001f4e6 Complete Pack (ZIP)", data=zip_buf.getvalue(), file_name=f"{_fn_prefix} Pack.zip", mime="application/zip", use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)

            # ----------------------------
            # RIGHT: Chart Hero + Methodology Cards + Table
            # ----------------------------
            with right_col:
                # CHART as hero element
                st.pyplot(fig, use_container_width=True)

                # Methodology stat strip
                n_months = len(index_df)
                n_sales = int(index_df['SalesCount'].sum()) if 'SalesCount' in index_df.columns else 0
                n_thin = int(index_df['ThinMonth'].sum()) if 'ThinMonth' in index_df.columns else 0
                avg_sales = f"{n_sales / max(1, n_months):.0f}"
                thin_note = f' <span style="color:var(--muted);font-size:11px;">({n_thin} thin)</span>' if n_thin > 0 else ''

                st.markdown(f'''
                <div style="display:flex; gap:10px; margin: 6px 0 20px 0; flex-wrap:wrap;">
                    <div style="flex:1; min-width:110px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">Method</div>
                        <div style="font-size:14px; font-weight:700; color:var(--ink);">Chain-Linked Index</div>
                    </div>
                    <div style="flex:1; min-width:100px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">Smoothing</div>
                        <div style="font-size:14px; font-weight:700; color:var(--ink);">{int(settings["smooth_window"])}-mo rolling avg</div>
                    </div>
                    <div style="flex:1; min-width:100px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">Data</div>
                        <div style="font-size:14px; font-weight:700; color:var(--ink);">{n_sales} sales \u00b7 {n_months} mo</div>
                    </div>
                    <div style="flex:1; min-width:100px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">Avg/Month</div>
                        <div style="font-size:14px; font-weight:700; color:var(--ink);">{avg_sales}{thin_note}</div>
                    </div>
                </div>
                ''', unsafe_allow_html=True)

                # Quarterly Median Sale Price tiles
                _eff_dt_q = pd.to_datetime(eff_date)
                _sd_q = df_model.copy()
                _sd_q["ContractDate"] = pd.to_datetime(_sd_q["ContractDate"])
                _sd_q["SoldPrice"] = _sd_q["SoldPrice"].astype(float)

                _quarters = [
                    ("0–3 Mo", _eff_dt_q - pd.DateOffset(months=3), _eff_dt_q),
                    ("3–6 Mo", _eff_dt_q - pd.DateOffset(months=6), _eff_dt_q - pd.DateOffset(months=3)),
                    ("6–9 Mo", _eff_dt_q - pd.DateOffset(months=9), _eff_dt_q - pd.DateOffset(months=6)),
                    ("9–12 Mo", _eff_dt_q - pd.DateOffset(months=12), _eff_dt_q - pd.DateOffset(months=9)),
                ]

                _q_tiles = ""
                _prev_med = None
                _q_medians = {}  # store for 12-month calc
                for _qlabel, _qstart, _qend in _quarters:
                    _qmask = (_sd_q["ContractDate"] >= _qstart) & (_sd_q["ContractDate"] < _qend)
                    _qsubset = _sd_q[_qmask]
                    if len(_qsubset) > 0:
                        _qmed = _qsubset["SoldPrice"].median()
                        _qcount = len(_qsubset)
                        _q_medians[_qlabel] = _qmed
                        # Arrow: change from prior (older) quarter to this (newer) quarter
                        if _prev_med is not None:
                            _qdiff = ((_qmed / _prev_med) - 1.0) * 100
                            if _qdiff > 1:
                                _qarrow = f'<span style="color:#22c55e;font-size:11px;"> \u2191{_qdiff:+.1f}%</span>'
                            elif _qdiff < -1:
                                _qarrow = f'<span style="color:#ef4444;font-size:11px;"> \u2193{_qdiff:+.1f}%</span>'
                            else:
                                _qarrow = f'<span style="color:var(--muted);font-size:11px;"> \u2192{_qdiff:+.1f}%</span>'
                        else:
                            _qarrow = ""
                        _prev_med = _qmed
                        _q_tiles += f'''<div style="flex:1; min-width:105px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                            <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">{_qlabel} ({_qcount})</div>
                            <div style="font-size:14px; font-weight:700; color:var(--ink);">${_qmed:,.0f}{_qarrow}</div>
                        </div>'''
                    else:
                        _prev_med = None
                        _q_tiles += f'''<div style="flex:1; min-width:105px; background:var(--surface2); border:1px solid var(--border); border-radius:var(--r-lg); padding:10px 14px;">
                            <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">{_qlabel}</div>
                            <div style="font-size:14px; font-weight:600; color:var(--muted);">No data</div>
                        </div>'''

                # 12-month overall change tile (9-12 Mo → 0-3 Mo)
                if "0\u20133 Mo" in _q_medians and "9\u201312 Mo" in _q_medians:
                    _yr_recent = _q_medians["0\u20133 Mo"]
                    _yr_oldest = _q_medians["9\u201312 Mo"]
                    _yr_chg = ((_yr_recent / _yr_oldest) - 1.0) * 100
                    if _yr_chg > 1:
                        _yr_color = "#22c55e"
                        _yr_arrow = "\u2191"
                    elif _yr_chg < -1:
                        _yr_color = "#ef4444"
                        _yr_arrow = "\u2193"
                    else:
                        _yr_color = "var(--muted)"
                        _yr_arrow = "\u2192"
                    _q_tiles += f'''<div style="flex:1; min-width:105px; background:var(--surface2); border:2px solid {_yr_color}40; border-radius:var(--r-lg); padding:10px 14px;">
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:2px;">12-Mo Change</div>
                        <div style="font-size:14px; font-weight:700; color:{_yr_color};">{_yr_arrow} {_yr_chg:+.1f}%</div>
                    </div>'''

                st.markdown(f'''
                <div style="font-size:10px; text-transform:uppercase; letter-spacing:0.8px; color:var(--muted); font-weight:700; margin-bottom:2px;">Quarterly Median Sale Price</div>
                <div style="font-size:10px; color:var(--muted); margin-bottom:8px;">Arrows show change from prior quarter \u00b7 Trailing quarters from effective date</div>
                <div style="display:flex; gap:10px; margin: 0 0 20px 0; flex-wrap:wrap;">
                    {_q_tiles}
                </div>
                ''', unsafe_allow_html=True)

                # Comparable Adjustments table
                st.markdown('<div class="vq-section-label">Comparable Adjustments</div>', unsafe_allow_html=True)

                display_out = out.copy()
                display_out.insert(0, "Comp #", range(1, len(display_out) + 1))

                def _badge_html(adj_pct, applied):
                    if not applied:
                        return '<span class="vq-badge vq-badge-stable">No Adj</span>'
                    if abs(adj_pct) < 0.1:
                        return '<span class="vq-badge vq-badge-stable">Stable</span>'
                    elif adj_pct > 0:
                        return f'<span class="vq-badge vq-badge-up">+{adj_pct:.1f}%</span>'
                    else:
                        return f'<span class="vq-badge vq-badge-down">{adj_pct:.1f}%</span>'

                rows_html = ""
                for _, r in display_out.iterrows():
                    badge = _badge_html(r["MktAdjPct"], r["AppliedAdj"])
                    adj_dollar = f"${r['MktAdj$']:+,.0f}" if r["AppliedAdj"] and abs(r["MktAdjPct"]) >= 0.1 else "\u2014"
                    rows_html += f"""<tr>
                        <td style="font-weight:650;color:var(--accent);">{int(r['Comp #'])}</td>
                        <td>{_escape(str(r['CompAddress']))}</td>
                        <td>{r['ContractDate'].strftime('%b %d, %Y') if hasattr(r['ContractDate'], 'strftime') else str(r['ContractDate'])}</td>
                        <td style="font-variant-numeric:tabular-nums;">${r['SalePrice']:,.0f}</td>
                        <td>{badge}</td>
                        <td style="font-variant-numeric:tabular-nums;font-weight:600;">{adj_dollar}</td>
                    </tr>"""

                st.markdown(f'''
                <div class="vq-table-wrap">
                <table class="vq-table">
                <thead><tr>
                    <th>#</th><th>Address</th><th>Contract Date</th>
                    <th>Sale Price</th><th>Adjustment</th><th>Adj $</th>
                </tr></thead>
                <tbody>{rows_html}</tbody>
                </table>
                </div>
                ''', unsafe_allow_html=True)

                st.markdown('<div class="vq-section-label">Narrative</div>', unsafe_allow_html=True)
                with st.expander("View Full Narrative", expanded=False):
                    st.text_area("Narrative", value=narrative, height=360, label_visibility="collapsed")

                st.markdown("---")
                st.markdown('<div class="vq-action-row">', unsafe_allow_html=True)
                if st.button("\u2190 Back to Comparables", use_container_width=True, key="step5_back"):
                    st.session_state["step"] = 4
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
