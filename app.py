# ============================================================
# CS WORKLOAD & CAPACITY DASHBOARD
# BUILD: V46_FIX_SHIPMENT_VOLUME_SOURCE_COLUMNS
# BUILD: SECTION2_SAME_ROW_V6
# Python + Streamlit + Pandas + Plotly
# Data source: (100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx
# ============================================================

from __future__ import annotations

import re
import hashlib
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="CS CAPACITY & PRODUCTIVITY",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "CS CAPACITY & PRODUCTIVITY"
APP_SUBTITLE = ""
DEFAULT_FILE = "(100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx"
CAPACITY_HOURS_PER_FTE = 8 * 0.95 * 22  # 167.2 hours/FTE/month
STANDARD_OFFICES = ["HAN", "HAD", "HLC", "HCM"]
SERVICE_ORDER = ["AE", "AI", "OE", "OI", "CC", "TR", "WH"]
SERVICE_LABELS = {
    "AE": "Air Export",
    "AI": "Air Import",
    "OE": "Ocean Export",
    "OI": "Ocean Import",
    "CC": "Customs Clearance",
    "TR": "Trucking",
    "WH": "Warehouse",
}

COLORS = {
    "navy": "#003B70",
    "blue": "#005BAC",
    "light_blue": "#EAF3F8",
    "red": "#E60012",
    "green": "#169B62",
    "amber": "#F59E0B",
    "gray": "#9AA3AD",
    "gray_dark": "#5F6B7A",
    "grid": "#E8EBEF",
    "bg": "#F7F8FA",
    "white": "#FFFFFF",
    "text": "#1F2937",
    "muted": "#64748B",
    "border": "#D9E2EC",
}

# Business-meaning color map.
# UI only: Same Business Meaning = Same Color Everywhere.
BUSINESS_COLORS = {
    "actual": COLORS["blue"],
    "approved": COLORS["navy"],
    "required": COLORS["amber"],
    "positive": COLORS["green"],
    "negative": COLORS["red"],
    "critical": COLORS["red"],
    "supporting": COLORS["gray"],
}


# ============================================================
# SHARED EXECUTIVE UI CONSTANTS
# UI only — no business logic / calculation changes
# ============================================================

UI = {
    "font_family": "Inter, 'Segoe UI', Arial, sans-serif",
    "title_size": 30,
    "section_title_size": 19,
    "chart_title_size": 17,
    "kpi_value_size": 32,
    "kpi_label_size": 13,
    "body_size": 12,
    "axis_size": 11,
    "note_size": 11,
    "radius": 12,
    "card_padding": 16,
    "section_gap": 18,
    "chart_height": 380,
    "chart_height_tall": 520,
}

CORPORATE_PALETTE = [
    COLORS["blue"],
    COLORS["navy"],
    COLORS["amber"],
    COLORS["green"],
    "#6B8EAD",
    "#A7B9C9",
    COLORS["red"],
]

SHEET_NAMES = {
    "hc": "HC",
    "resolution": "CS Resolutions Rate",
    "workload": "BU allocation",
    "yvf": "YVF",
    "shipment": "Shipment volume",
    "customer_ns": "Customer Volume-N&S",
    "customer_had": "Customer Volume - HAD",
    "customer_han": "Customer Volume - HAN",
    "customer_hlc": "Customer Volume - HLC",
    "customer_hcm": "Customer Volume - HCM",
    "fte": "CS FTE",
    "core": "C",
    "ancillary": "A",
    "supporting": "S",
    "exception": "E",
    "notes": "Ghi chú",
}

# ============================================================
# STYLE
# ============================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {COLORS['bg']};
        color: {COLORS['text']};
    }}

    /* Reduce top whitespace and move dashboard content upward */
    .block-container {{
        padding-top: 1.4rem !important;
    }}

    .main-header {{
        margin-top: 0 !important;
    }}
    section[data-testid="stSidebar"] {{
        background: {COLORS['navy']};
    }}
    /* Sidebar: high-contrast labels, captions and controls */
    section[data-testid="stSidebar"] {{
        background: {COLORS['navy']};
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{
        color: #FFFFFF !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
        background: #FFFFFF !important;
        border-color: #D9E2EC !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"] div[data-baseweb="select"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
        color: {COLORS['navy']} !important;
        fill: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background: #FFFFFF !important;
        border-color: #D9E2EC !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section * {{
        color: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
        background: #FFFFFF !important;
        color: {COLORS['navy']} !important;
        border: 1px solid #B8C7D6 !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button * {{
        color: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    .main-header {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-left: 7px solid {COLORS['blue']};
        border-radius: 14px;
        padding: 18px 22px;
        margin-bottom: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }}
    .main-title {{
        color: {COLORS['navy']};
        font-size: 30px;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.02em;
    }}
    .subtitle {{
        color: {COLORS['muted']};
        font-size: 14px;
        margin-top: 4px;
    }}
    .section-title {{
        color: {COLORS['navy']};
        font-size: 18px;
        font-weight: 800;
        margin: 22px 0 8px 0;
        border-left: 5px solid {COLORS['amber']};
        padding-left: 10px;
    }}
    .kpi-card {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-radius: 14px;
        padding: 16px 14px;
        min-height: 110px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
    }}

    .hc-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 16px 16px 14px 16px;
        min-height: 190px;
        height: 190px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        text-align: center;
    }}

    .hc-kpi-card .kpi-label {{
        text-align: center;
        width: 100%;
    }}

    .hc-kpi-total {{
        color: #003B70;
        font-size: 32px;
        line-height: 1.05;
        font-weight: 850;
        margin-top: 12px;
        margin-bottom: 8px;
        text-align: center;
        width: 100%;
    }}

    .hc-detail-row {{
        display: grid;
        grid-template-columns: 1fr 1px 1fr;
        align-items: stretch;
        gap: 12px;
        margin-top: auto;
        padding-top: 12px;
        border-top: 1px solid #E5E7EB;
    }}

    .hc-detail-divider {{
        background: #E5E7EB;
        width: 1px;
    }}

    .hc-detail-item {{
        text-align: center;
    }}

    .hc-detail-label {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.04em;
    }}

    .hc-detail-value {{
        color: #003B70;
        font-size: 20px;
        line-height: 1.2;
        font-weight: 800;
        margin-top: 3px;
    }}

    .hc-variance-card {{
        justify-content: flex-start;
    }}

    .hc-variance-formula {{
        color: #64748B;
        font-size: 12px;
        font-weight: 600;
        margin-top: 12px;
        margin-bottom: 10px;
        text-align: center;
        width: 100%;
    }}

    .hc-variance-status {{
        display: block;
        width: 100%;
        box-sizing: border-box;
        text-align: center;
        margin-top: 0 !important;
    }}


    /* Section 2 - Shipment KPI cards */
    .shipment-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 18px 18px;
        min-height: 170px;
        height: 170px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
    }}

    .shipment-kpi-label {{
        color: #64748B;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 14px;
    }}

    .shipment-kpi-value {{
        color: #003B70;
        font-size: 36px;
        line-height: 1.05;
        font-weight: 850;
    }}

    .shipment-kpi-note {{
        color: #64748B;
        font-size: 11px;
        margin-top: 10px;
    }}


    /* Section 3 - Workload by PIC */
    .pic-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 14px;
        min-height: 142px;
        height: 142px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: center;
        text-align: center;
    }}

    .pic-kpi-label {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.035em;
        line-height: 1.25;
        min-height: 30px;
    }}

    .pic-kpi-value {{
        color: #003B70;
        font-size: 27px;
        line-height: 1.05;
        font-weight: 850;
        margin-top: 8px;
    }}

    .pic-kpi-note {{
        color: #64748B;
        font-size: 10.5px;
        margin-top: 7px;
        line-height: 1.25;
    }}

    .pic-status-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 16px;
        min-height: 110px;
        height: 110px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
    }}

    .pic-status-left {{
        min-width: 180px;
    }}

    .pic-status-title {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.035em;
    }}

    .pic-status-value {{
        color: #003B70;
        font-size: 30px;
        font-weight: 850;
        margin-top: 4px;
    }}

    .pic-progress-track {{
        flex: 1;
        height: 13px;
        background: #EAF3F8;
        border-radius: 999px;
        overflow: hidden;
    }}

    .pic-progress-fill {{
        height: 100%;
        border-radius: 999px;
    }}

    .workload-status-panel {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 16px;
        min-height: 110px;
        height: 110px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: center;
        text-align: center;
    }}

    .workload-status-text {{
        font-size: 20px;
        font-weight: 850;
        margin-top: 8px;
    }}
    .kpi-label {{
        color: {COLORS['muted']};
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 6px;
    }}
    .kpi-value {{
        color: {COLORS['navy']};
        font-size: 26px;
        line-height: 1.05;
        font-weight: 850;
        margin-bottom: 4px;
    }}
    .kpi-note {{
        color: {COLORS['muted']};
        font-size: 11px;
    }}
    .status-badge {{
        display:inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-top: 8px;
    }}
    .chart-box {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-radius: 14px;
        padding: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
    }}
    .warning-box {{
        background: #FFF7ED;
        color: #92400E;
        border: 1px solid #FED7AA;
        border-radius: 12px;
        padding: 10px 12px;
        margin-bottom: 10px;
        font-size: 13px;
    }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 8px; }}
    .stTabs [data-baseweb="tab"] {{
        background: {COLORS['white']};
        border-radius: 10px 10px 0 0;
        border: 1px solid {COLORS['border']};
        padding: 8px 16px;
    }}
    
    /* Section 3 compact executive layout */
    .compact-workload-kpi {{
        min-height: 138px !important;
        height: 138px !important;
        padding: 16px 18px 14px 18px !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }}
    .compact-workload-kpi .kpi-label {{
        margin-bottom: 8px !important;
        line-height: 1.15 !important;
    }}
    .compact-workload-kpi .kpi-value {{
        margin: 2px 0 6px 0 !important;
        line-height: 1.05 !important;
    }}
    .compact-workload-kpi .kpi-note {{
        margin-top: 4px !important;
        line-height: 1.15 !important;
    }}
    .workload-util-card, .workload-status-card {{
        min-height: 94px !important;
        height: 94px !important;
        padding: 13px 16px !important;
    }}
    .workload-util-card {{
        display: grid !important;
        grid-template-columns: 175px minmax(0, 1fr) !important;
        align-items: center !important;
        column-gap: 18px !important;
    }}
    .workload-status-card {{
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }}
    @media (max-width: 1100px) {{
        .workload-util-card {{
            grid-template-columns: 155px minmax(0, 1fr) !important;
            column-gap: 12px !important;
        }}
    }}

</style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# EXECUTIVE / CORPORATE UI OVERRIDES
# Shared styling layer only — business logic remains unchanged
# ============================================================

st.markdown(
    f"""
    <style>
    :root {{
        --font-main: {UI['font_family']};
        --navy: {COLORS['navy']};
        --blue: {COLORS['blue']};
        --orange: {COLORS['amber']};
        --green: {COLORS['green']};
        --red: {COLORS['red']};
        --text: {COLORS['text']};
        --muted: #667085;
        --border: #D8E1EA;
        --surface: #FFFFFF;
        --background: #F6F8FA;
        --radius: {UI['radius']}px;
    }}

    html, body, [class*="css"], .stApp,
    button, input, textarea, select {{
        font-family: var(--font-main) !important;
    }}

    .stApp {{
        background: var(--background);
        color: var(--text);
    }}

    .block-container {{
        max-width: 1680px;
        padding-top: 1.35rem !important;
        padding-left: 1.45rem !important;
        padding-right: 1.45rem !important;
        padding-bottom: 2rem !important;
    }}

    /* Header */
    .main-header {{
        border-radius: var(--radius);
        padding: 16px 20px;
        margin: 0 0 14px 0 !important;
        border: 1px solid var(--border);
        border-left: 5px solid var(--blue);
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.05);
    }}

    .main-title {{
        font-size: {UI['title_size']}px !important;
        line-height: 1.15 !important;
        font-weight: 700 !important;
        letter-spacing: -0.015em !important;
        color: var(--navy) !important;
    }}

    .subtitle {{
        font-size: {UI['body_size']}px !important;
        line-height: 1.45 !important;
        color: var(--muted) !important;
        margin-top: 3px !important;
    }}

    /* Section titles */
    .section-title {{
        font-size: {UI['section_title_size']}px !important;
        line-height: 1.25 !important;
        font-weight: 700 !important;
        color: var(--navy) !important;
        margin: 20px 0 10px 0 !important;
        padding: 0 0 0 10px !important;
        border-left: 4px solid var(--orange) !important;
    }}

    /* Shared card language */
    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card,
    .pic-status-card,
    .workload-status-panel,
    .chart-box {{
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.045) !important;
        box-sizing: border-box !important;
    }}

    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card {{
        padding: {UI['card_padding']}px !important;
    }}

    /* KPI labels */
    .kpi-label,
    .shipment-kpi-label,
    .pic-kpi-label {{
        color: #5F6B7A !important;
        font-size: {UI['kpi_label_size']}px !important;
        line-height: 1.25 !important;
        font-weight: 600 !important;
        letter-spacing: 0.025em !important;
        text-transform: uppercase !important;
        margin-bottom: 7px !important;
    }}

    /* KPI values */
    .kpi-value,
    .hc-kpi-total,
    .shipment-kpi-value,
    .pic-kpi-value,
    .pic-status-value {{
        color: var(--navy) !important;
        font-size: {UI['kpi_value_size']}px !important;
        line-height: 1.05 !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
    }}

    /* Notes / formulas / sources */
    .kpi-note,
    .shipment-kpi-note,
    .pic-kpi-note,
    .hc-variance-formula,
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] p {{
        color: var(--muted) !important;
        font-size: {UI['note_size']}px !important;
        line-height: 1.4 !important;
        font-weight: 400 !important;
    }}

    /* HC cards — equal structure */
    .hc-kpi-card {{
        height: 184px !important;
        min-height: 184px !important;
    }}

    .hc-detail-row {{
        margin-top: auto !important;
        padding-top: 11px !important;
        gap: 10px !important;
        border-top: 1px solid #E7ECF1 !important;
    }}

    .hc-detail-label {{
        color: var(--muted) !important;
        font-size: 11px !important;
        font-weight: 600 !important;
    }}

    .hc-detail-value {{
        color: var(--navy) !important;
        font-size: 19px !important;
        font-weight: 700 !important;
    }}

    /* Office Capacity Snapshot — semantic color hierarchy */
    .hc-total-approved {{
        color: var(--navy) !important;
    }}
    .hc-total-actual {{
        color: var(--blue) !important;
    }}
    .hc-total-required {{
        color: var(--orange) !important;
    }}

    /* Shipment KPI cards */
    .shipment-kpi-card {{
        height: 154px !important;
        min-height: 154px !important;
    }}

    /* PIC KPI cards */
    .pic-kpi-card {{
        height: 140px !important;
        min-height: 140px !important;
    }}

    .pic-status-card,
    .workload-status-panel {{
        height: 104px !important;
        min-height: 104px !important;
        padding: 14px 16px !important;
    }}

    .pic-status-title {{
        font-size: 12px !important;
        font-weight: 600 !important;
        color: var(--muted) !important;
        letter-spacing: 0.025em !important;
    }}

    /* Chart wrappers */
    .chart-box {{
        padding: 12px 14px 10px 14px !important;
        overflow: visible !important;
        min-width: 0 !important;
    }}

    .chart-box [data-testid="stPlotlyChart"] {{
        margin: 0 !important;
    }}

    .chart-box .js-plotly-plot,
    .chart-box .plot-container,
    .chart-box .svg-container {{
        width: 100% !important;
    }}

    /* Status */
    .status-badge {{
        font-size: 11px !important;
        font-weight: 700 !important;
        padding: 4px 9px !important;
    }}

    /* Streamlit dataframe */
    [data-testid="stDataFrame"] {{
        border: 1px solid var(--border);
        border-radius: var(--radius);
        overflow: hidden;
    }}

    /* Vertical spacing between Streamlit blocks */
    div[data-testid="stVerticalBlock"] > div {{
        gap: 0.35rem;
    }}

    /* Sidebar remains high contrast */
    section[data-testid="stSidebar"] {{
        background: var(--navy) !important;
    }}

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #FFFFFF !important;
        opacity: 1 !important;
    }}

    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background: #FFFFFF !important;
        border: 1px solid #C9D5E1 !important;
        border-radius: 8px !important;
    }}

    /* Laptop responsive behavior */
    @media (max-width: 1200px) {{
        .block-container {{
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }}

        .main-title {{
            font-size: 28px !important;
        }}

        .kpi-value,
        .hc-kpi-total,
        .shipment-kpi-value,
        .pic-kpi-value,
        .pic-status-value {{
            font-size: 28px !important;
        }}

        .kpi-label,
        .shipment-kpi-label,
        .pic-kpi-label {{
            font-size: 12px !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FINAL UI/UX QUALITY OVERRIDES
# UI only — no business logic / calculation changes
# ============================================================

st.markdown(
    f"""
    <style>
    /* Management-first density and consistent vertical rhythm */
    .block-container {{
        max-width: 1680px !important;
        padding-top: 1.10rem !important;
        padding-bottom: 1.75rem !important;
    }}

    .section-title {{
        margin-top: 22px !important;
        margin-bottom: 12px !important;
    }}

    /* Standard Streamlit charts as real cards.
       Avoid raw HTML wrappers around Streamlit widgets. */
    [data-testid="stPlotlyChart"] {{
        background: #FFFFFF;
        border: 1px solid {COLORS['border']};
        border-radius: {UI['radius']}px;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.045);
        padding: 8px 10px 4px 10px;
        box-sizing: border-box;
    }}

    [data-testid="stDataFrame"] {{
        background: #FFFFFF;
        border: 1px solid {COLORS['border']} !important;
        border-radius: {UI['radius']}px !important;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.035);
    }}

    /* Equal KPI-card language */
    .kpi-card {{
        min-height: 124px !important;
        height: 124px !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }}

    .kpi-label {{
        min-height: 18px;
    }}

    /* Make notes subordinate to management metrics */
    .kpi-note {{
        margin-top: 5px !important;
        line-height: 1.35 !important;
    }}

    /* Compact tabs and avoid visual competition */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px !important;
        border-bottom: 1px solid {COLORS['grid']} !important;
    }}

    .stTabs [data-baseweb="tab"] {{
        border: 0 !important;
        border-radius: 0 !important;
        background: transparent !important;
        padding: 7px 10px !important;
        font-size: 12px !important;
    }}

    /* Responsive laptop refinements */
    @media (max-width: 1366px) {{
        .block-container {{
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
        }}
        .main-title {{
            font-size: 28px !important;
        }}
        .section-title {{
            font-size: 18px !important;
        }}
        .kpi-value,
        .hc-kpi-total,
        .shipment-kpi-value,
        .pic-kpi-value,
        .pic-status-value {{
            font-size: 28px !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================


def safe_float(value) -> float:
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value.startswith("="):
            return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def clean_col(col) -> str:
    col = str(col).replace("\n", " ").replace("\r", " ")
    col = re.sub(r"\s+", " ", col).strip()
    return col


def normalize_office(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def parse_month(value) -> pd.Timestamp | pd.NaT:
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    text = str(value).strip()
    for fmt in ["%b-%y", "%b-%Y", "%Y-%m", "%m/%Y", "%Y/%m"]:
        try:
            dt = pd.to_datetime(text, format=fmt)
            return pd.Timestamp(year=dt.year, month=dt.month, day=1)
        except Exception:
            pass
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return pd.NaT
        return pd.Timestamp(year=dt.year, month=dt.month, day=1)
    except Exception:
        return pd.NaT


def format_month(ts) -> str:
    if pd.isna(ts):
        return ""
    return pd.Timestamp(ts).strftime("%b-%y")


def numeric_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def read_sheet(path: str | Path, sheet: str, header: int = 1) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header, engine="openpyxl")
        df.columns = [clean_col(c) for c in df.columns]
        df = df.dropna(how="all")
        return df
    except Exception:
        return pd.DataFrame()


def ensure_cols(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    return df


def first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    clean_map = {clean_col(c).lower(): c for c in df.columns}
    for c in candidates:
        key = clean_col(c).lower()
        if key in clean_map:
            return clean_map[key]
    return None


def weighted_period_avg(df: pd.DataFrame, value_col: str, group_col: str = "MonthDate") -> float:
    """Average of valid monthly totals only; blank future months are excluded."""
    if df.empty or value_col not in df.columns or group_col not in df.columns:
        return 0.0

    valid = df[[group_col, value_col]].copy()
    valid[value_col] = pd.to_numeric(valid[value_col], errors="coerce")
    valid = valid.dropna(subset=[group_col, value_col])

    if valid.empty:
        return 0.0

    monthly = (
        valid.groupby(group_col, dropna=True)[value_col]
        .sum(min_count=1)
        .reset_index()
        .dropna(subset=[value_col])
    )
    if monthly.empty:
        return 0.0
    return float(monthly[value_col].mean())


def safe_div(num: float, den: float) -> float:
    return float(num) / float(den) if den not in [0, 0.0] and not pd.isna(den) else 0.0


def fmt_num(value: float, decimals: int = 1, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value:,.{decimals}f}{suffix}"


def fmt_int(value: float) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value:,.0f}"


def fmt_pct(value: float) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value * 100:,.1f}%"


def status_from_util(util: float) -> Tuple[str, str, str]:
    if util <= 0:
        return "NO DATA", COLORS["muted"], COLORS["light_blue"]
    if util < 0.90:
        return "LESS LOAD", COLORS["blue"], "#DBEAFE"
    if util <= 0.95:
        return "BALANCED", COLORS["green"], "#DCFCE7"
    if util <= 1.00:
        return "HIGH LOAD", COLORS["amber"], "#FEF3C7"
    return "OVERLOAD", COLORS["red"], "#FEE2E2"


def kpi_card(label: str, value: str, note: str = "", status: Optional[Tuple[str, str, str]] = None):
    badge = ""
    if status:
        txt, color, bg = status
        badge = f'<span class="status-badge" style="color:{color};background:{bg};">{txt}</span>'
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
            {badge}
        </div>
        """,
        unsafe_allow_html=True,
    )



def hc_detail_card(
    label: str,
    total_value: float,
    mng_value: Optional[float] = None,
    pic_value: Optional[float] = None,
    note_left: str = "MNG",
    note_right: str = "PIC",
    status_text: Optional[str] = None,
    status_color: Optional[str] = None,
    status_bg: Optional[str] = None,
):
    """Executive HC card with equal height and two aligned detail blocks at the bottom."""
    details_html = ""
    if mng_value is not None or pic_value is not None:
        left_decimals = 2 if "REQUIRED" in label.upper() else 0
        right_decimals = 2 if "REQUIRED" in label.upper() else 0
        left_val = fmt_num(mng_value or 0, left_decimals)
        right_val = fmt_num(pic_value or 0, right_decimals)
        details_html = f"""
        <div class="hc-detail-row">
            <div class="hc-detail-item">
                <div class="hc-detail-label">{note_left}</div>
                <div class="hc-detail-value">{left_val}</div>
            </div>
            <div class="hc-detail-divider"></div>
            <div class="hc-detail-item">
                <div class="hc-detail-label">{note_right}</div>
                <div class="hc-detail-value">{right_val}</div>
            </div>
        </div>
        """

    status_html = ""
    if status_text:
        status_html = (
            f'<span class="status-badge" '
            f'style="color:{status_color};background:{status_bg};margin-top:10px;">'
            f'{status_text}</span>'
        )

    # UI-only semantic color hierarchy for Office Capacity Snapshot:
    # Approved = Navy, Actual = Corporate Blue, Required = Orange.
    label_upper = label.upper()
    if "REQUIRED" in label_upper:
        total_color_class = "hc-total-required"
    elif "ACTUAL" in label_upper:
        total_color_class = "hc-total-actual"
    else:
        total_color_class = "hc-total-approved"

    st.markdown(
        f"""
        <div class="hc-kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="hc-kpi-total {total_color_class}">{fmt_num(total_value, 2 if "REQUIRED" in label.upper() else 0)}</div>
            {status_html}
            {details_html}
        </div>
        """,
        unsafe_allow_html=True,
    )



def hc_variance_card(
    label: str,
    value: float,
    formula_text: str,
    status_text: str,
    status_color: str,
    status_bg: str,
):
    """Centered variance card to visually balance the HC cards."""
    st.markdown(
        f"""
        <div class="hc-kpi-card hc-variance-card">
            <div class="kpi-label">{label}</div>
            <div class="hc-kpi-total" style="color:{status_color} !important;">{fmt_num(value, 0)}</div>
            <div class="hc-variance-formula">{formula_text}</div>
            <span class="status-badge hc-variance-status"
                  style="color:{status_color};background:{status_bg};">
                {status_text}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def shipment_kpi_card(label: str, value: str, note: str = ""):
    """Equal-size centered KPI card for Shipment Volume section."""
    st.markdown(
        f"""
        <div class="shipment-kpi-card">
            <div class="shipment-kpi-label">{label}</div>
            <div class="shipment-kpi-value">{value}</div>
            <div class="shipment-kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def pic_kpi_card(label: str, value: str, note: str = ""):
    """Compact numeric KPI card for Workload by PIC."""
    st.markdown(
        f"""
        <div class="pic-kpi-card">
            <div class="pic-kpi-label">{label}</div>
            <div class="pic-kpi-value">{value}</div>
            <div class="pic-kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filtered_monthly_metric(
    df: pd.DataFrame,
    value_col: str,
    agg: str = "sum",
) -> float:
    """
    KPI helper:
    - Single month => exact selected-month aggregate.
    - All months => average of monthly aggregates.
    """
    if df is None or df.empty or value_col not in df.columns:
        return float("nan")

    d = df[["MonthDate", value_col]].copy()
    d[value_col] = pd.to_numeric(d[value_col], errors="coerce")
    d = d.dropna(subset=["MonthDate", value_col])
    if d.empty:
        return float("nan")

    if agg == "sum":
        monthly = d.groupby("MonthDate")[value_col].sum(min_count=1)
    elif agg == "mean":
        monthly = d.groupby("MonthDate")[value_col].mean()
    else:
        raise ValueError("agg must be 'sum' or 'mean'")

    monthly = monthly.dropna()
    if monthly.empty:
        return float("nan")
    return float(monthly.mean())


def hc_capacity_utilization(df: pd.DataFrame) -> float:
    """
    Capacity Utilization KPI source of truth:
    sheet HC -> column "Capacity Utilization (%)".

    Single Office / Month:
        use the source value directly.

    All Offices:
        calculate each month's overall utilization as the Actual-HC-weighted
        average of the office source percentages, then average across selected months.

    Blank future months are excluded.
    """
    if df is None or df.empty or "HC Utilization" not in df.columns:
        return float("nan")

    cols = ["MonthDate", "HC Utilization"]
    if "Total Actual HC" in df.columns:
        cols.append("Total Actual HC")

    d = df[cols].copy()
    d["HC Utilization"] = pd.to_numeric(d["HC Utilization"], errors="coerce")
    if "Total Actual HC" in d.columns:
        d["Total Actual HC"] = pd.to_numeric(d["Total Actual HC"], errors="coerce")

    d = d.dropna(subset=["MonthDate", "HC Utilization"])
    if d.empty:
        return float("nan")

    monthly_values = []

    for _, g in d.groupby("MonthDate"):
        # If only one row/office for the month, this is the exact source value.
        if len(g) == 1:
            monthly_values.append(float(g["HC Utilization"].iloc[0]))
            continue

        # All Offices: weight by Actual HC so larger offices contribute appropriately.
        if "Total Actual HC" in g.columns:
            valid = g["Total Actual HC"].notna() & (g["Total Actual HC"] > 0)
            if valid.any():
                weighted = (
                    g.loc[valid, "HC Utilization"]
                    * g.loc[valid, "Total Actual HC"]
                ).sum() / g.loc[valid, "Total Actual HC"].sum()
                monthly_values.append(float(weighted))
                continue

        # Fallback only if HC weights are unavailable.
        monthly_values.append(float(g["HC Utilization"].mean()))

    monthly_values = [v for v in monthly_values if not pd.isna(v)]
    return float(np.mean(monthly_values)) if monthly_values else float("nan")



def pic_utilization_card(util: float):
    if pd.isna(util):
        value = "N/A"
        pct = 0
        color = COLORS["muted"]
    else:
        value = fmt_pct(util)
        pct = max(0, min(util * 100, 125))
        _, color, _ = status_from_util(util)

    width_pct = min(pct / 125 * 100, 100)
    st.markdown(
        f"""
        <div class="pic-status-card">
            <div class="pic-status-left">
                <div class="pic-status-title">CAPACITY UTILIZATION</div>
                <div class="pic-status-value">{value}</div>
            </div>
            <div class="pic-progress-track">
                <div class="pic-progress-fill"
                     style="width:{width_pct:.1f}%;background:{color};"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def overall_workload_status_card(util: float):
    if pd.isna(util):
        status_text, color, bg = "NO DATA", COLORS["muted"], COLORS["light_blue"]
    else:
        status_text, color, bg = status_from_util(util)

    st.markdown(
        f"""
        <div class="workload-status-panel">
            <div class="pic-status-title">OVERALL WORKLOAD STATUS</div>
            <div class="workload-status-text"
                 style="color:{color};background:{bg};
                        border-radius:999px;padding:8px 14px;">
                {status_text}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(text: str):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def plotly_layout(
    fig: go.Figure,
    height: int = UI["chart_height"],
    *,
    show_legend: bool = True,
    legend_position: str = "top",
    margin_left: int = 52,
    margin_right: int = 36,
    margin_top: int = 62,
    margin_bottom: int = 44,
) -> go.Figure:
    """Shared Executive/Corporate Plotly layout — UI only."""
    legend_cfg = dict(
        font=dict(size=UI["axis_size"]),
        bgcolor="rgba(0,0,0,0)",
        borderwidth=0,
    )

    if legend_position == "top":
        legend_cfg.update(
            orientation="h",
            yanchor="bottom",
            y=1.015,
            xanchor="right",
            x=1,
        )
    elif legend_position == "bottom":
        legend_cfg.update(
            orientation="h",
            yanchor="top",
            y=-0.16,
            xanchor="left",
            x=0,
        )
    else:
        legend_cfg.update(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
        )

    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            color=COLORS["text"],
            family=UI["font_family"],
            size=UI["axis_size"],
        ),
        title=dict(
            font=dict(
                size=UI["chart_title_size"],
                color=COLORS["navy"],
                family=UI["font_family"],
            ),
            x=0.0,
            xanchor="left",
            y=0.985,
            yanchor="top",
            pad=dict(t=0, b=8),
        ),
        margin=dict(
            l=margin_left,
            r=margin_right,
            t=margin_top,
            b=margin_bottom,
        ),
        legend=legend_cfg,
        showlegend=show_legend,
        hoverlabel=dict(
            font=dict(family=UI["font_family"], size=UI["axis_size"]),
            bgcolor="#FFFFFF",
            bordercolor=COLORS["border"],
        ),
        hovermode="closest",
    )

    fig.update_xaxes(
        gridcolor=COLORS["grid"],
        gridwidth=0.7,
        zeroline=False,
        showline=False,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
        automargin=True,
        ticks="",
    )
    fig.update_yaxes(
        gridcolor=COLORS["grid"],
        gridwidth=0.7,
        zeroline=False,
        showline=False,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
        automargin=True,
        ticks="",
    )
    return fig

# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data(show_spinner=False)
def load_data(path: str, cache_token: str = "") -> Dict[str, pd.DataFrame]:
    """
    FAST workbook loader:
    - Open Excel workbook only ONCE with pd.ExcelFile.
    - Parse required sheets from the same workbook handle.
    - cache_token invalidates Streamlit cache when file content changes.
    """
    data: Dict[str, pd.DataFrame] = {}

    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        available_sheets = set(xls.sheet_names)
    except Exception:
        return {key: pd.DataFrame() for key in SHEET_NAMES}

    for key, sheet in SHEET_NAMES.items():
        if sheet not in available_sheets:
            data[key] = pd.DataFrame()
            continue

        try:
            df = pd.read_excel(
                xls,
                sheet_name=sheet,
                header=1,
            )
            df.columns = [clean_col(c) for c in df.columns]
            df = df.dropna(how="all")
            data[key] = df
        except Exception:
            data[key] = pd.DataFrame()

    return data



@st.cache_data(show_spinner=False)
def prepare_hc(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    if not office_col or not month_col:
        return pd.DataFrame(columns=["Office", "MonthDate"])
    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    mapping = {
        "Approved HC – MNG": "Approved HC MNG",
        "Approved HC – PIC": "Approved HC PIC",
        "Total Approved HC": "Total Approved HC",
        "Actual HC – MNG": "Actual HC MNG",
        "Actual HC – PIC": "Actual HC PIC",
        "Total Actual HC": "Total Actual HC",
        "Total Actual  HC": "Total Actual HC",
        "Required HC – MNG": "Required HC MNG",
        "Required HC – PIC": "Required HC PIC",
        "Total Required HC": "Total Required HC",
        "Total Available Standard Time (95%x8x22xPIC)": "HC Available Hours",
        "Total actual Working Time (=C+A+S+E)": "HC Actual Working Hours",
        "Actual workload/PIC (hour)": "HC Actual Workload per PIC",
        "Capacity Utilization (%)": "HC Utilization",
        "HC Utilization (%)": "HC Utilization",
        "Overal  Workload Status": "HC Status",
        "Overall Workload Status": "HC Status",
        "HC Status": "HC Status",
    }
    for old, new in mapping.items():
        col = first_existing(df, [old])
        if col:
            df[new] = df[col]
    # Fallback calculations
    if "Total Approved HC" not in df.columns:
        df["Total Approved HC"] = numeric_series(df.get("Approved HC MNG", 0)) + numeric_series(df.get("Approved HC PIC", 0))
    if "Total Actual HC" not in df.columns:
        df["Total Actual HC"] = numeric_series(df.get("Actual HC MNG", 0)) + numeric_series(df.get("Actual HC PIC", 0))
    hc_numeric_cols = [
        "Approved HC MNG", "Approved HC PIC", "Total Approved HC",
        "Actual HC MNG", "Actual HC PIC", "Total Actual HC",
        "Required HC MNG", "Required HC PIC", "Total Required HC",
        "HC Available Hours", "HC Actual Working Hours",
        "HC Actual Workload per PIC", "HC Utilization"
    ]
    for col in hc_numeric_cols:
        if col in df.columns:
            # IMPORTANT: keep blank Excel cells as NaN.
            # Do not convert future blank months to 0 because that distorts averages/trends.
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["MonthDate"])

    # Keep only HC rows that actually contain HC data.
    hc_key_cols = [
        c for c in [
            "Total Approved HC", "Total Actual HC", "Total Required HC",
            "Approved HC MNG", "Approved HC PIC",
            "Actual HC MNG", "Actual HC PIC",
            "Required HC MNG", "Required HC PIC"
        ] if c in df.columns
    ]
    if hc_key_cols:
        df = df.dropna(subset=hc_key_cols, how="all")

    return df


@st.cache_data(show_spinner=False)
def prepare_workload(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate", "Segment"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    segment_col = first_existing(df, ["Segment"])
    if not office_col or not month_col or not segment_col:
        return pd.DataFrame(columns=["Office", "MonthDate", "Segment"])
    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    df["Segment"] = df[segment_col].astype(str).str.strip().str.upper()
    # Use BU allocation workload columns as source of truth.
    component_map = {
        "Core Workload (min)": "Core Workload (min)",
        "Ancillary Workload (min)": "Ancillary Workload (min)",
        "Supporting Workload (min)": "Supporting Workload (min)",
        "Exception Workload (min)": "Exception Workload (min)",
        "Total Workload (min)": "Total Workload (min)",
        "% of Network": "Workload Share",
        "OFFICE HC ALLOCATION RATIO TO Bus": "Office HC Allocation Ratio",
    }
    for old, new in component_map.items():
        col = first_existing(df, [old])
        if col:
            df[new] = numeric_series(df[col])
        else:
            df[new] = 0.0
    if df["Total Workload (min)"].sum() == 0:
        df["Total Workload (min)"] = (
            df["Core Workload (min)"]
            + df["Ancillary Workload (min)"]
            + df["Supporting Workload (min)"]
            + df["Exception Workload (min)"]
        )
    df["Workload Hours"] = df["Total Workload (min)"] / 60
    df["Core Hours"] = df["Core Workload (min)"] / 60
    df["Ancillary Hours"] = df["Ancillary Workload (min)"] / 60
    df["Supporting Hours"] = df["Supporting Workload (min)"] / 60
    df["Exception Hours"] = df["Exception Workload (min)"] / 60
    df["Service Label"] = df["Segment"].map(SERVICE_LABELS).fillna(df["Segment"])
    return df.dropna(subset=["MonthDate"])


@st.cache_data(show_spinner=False)
def prepare_fte(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "CS PIC", "MonthDate", "Actual FTE"])
    df = df.copy()
    office_col = first_existing(df, ["OFFICE", "Office"])
    pic_col = first_existing(df, ["CS PIC", "PIC"])
    if not office_col or not pic_col:
        return pd.DataFrame(columns=["Office", "CS PIC", "MonthDate", "Actual FTE"])
    month_cols = [c for c in df.columns if parse_month(c) is not pd.NaT and not pd.isna(parse_month(c))]
    if not month_cols:
        return pd.DataFrame(columns=["Office", "CS PIC", "MonthDate", "Actual FTE"])
    long = df.melt(id_vars=[office_col, pic_col], value_vars=month_cols, var_name="Month", value_name="Actual FTE")
    long["Office"] = long[office_col].map(normalize_office)
    long["CS PIC"] = long[pic_col].astype(str).str.strip()
    long["MonthDate"] = long["Month"].map(parse_month)

    # IMPORTANT:
    # Keep blank future-month FTE cells as NaN.
    # Converting blanks to 0 would dilute the average FTE when Month = All
    # and can incorrectly hide overloaded PICs.
    long["Actual FTE"] = pd.to_numeric(long["Actual FTE"], errors="coerce")

    long = long[
        (long["Office"] != "")
        & (~long["MonthDate"].isna())
        & (long["Actual FTE"].notna())
    ]
    return long[["Office", "CS PIC", "MonthDate", "Actual FTE"]]


@st.cache_data(show_spinner=False)
def prepare_shipment(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if df.empty:
        return (
            pd.DataFrame(columns=["Office", "MonthDate", "Total Shipment", "Active Customers"]),
            pd.DataFrame(columns=["Office", "MonthDate", "Mode", "Volume"]),
        )

    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    active_col = first_existing(df, ["Active Customers"])
    total_col = first_existing(df, ["TOTAL", "Total"])

    if not office_col or not month_col:
        return pd.DataFrame(), pd.DataFrame()

    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)

    # Keep blanks as NaN first so future empty months are not treated as real zero-data months.
    if total_col:
        df["Total Shipment"] = pd.to_numeric(df[total_col], errors="coerce")
    else:
        df["Total Shipment"] = np.nan

    if active_col:
        df["Active Customers"] = pd.to_numeric(df[active_col], errors="coerce")
    else:
        df["Active Customers"] = np.nan

    # Exclude rows/months where both key shipment metrics are blank.
    df = df.dropna(subset=["MonthDate"])
    df = df.dropna(subset=["Total Shipment", "Active Customers"], how="all")

    # Valid rows can safely use zero fallback afterward.
    df["Total Shipment"] = df["Total Shipment"].fillna(0)
    df["Active Customers"] = df["Active Customers"].fillna(0)

    excluded = {
        office_col, month_col, active_col, total_col,
        "Office", "MonthDate", "Active Customers", "Total Shipment"
    }
    mode_cols = [
        c for c in df.columns
        if c not in excluded and not str(c).startswith("Unnamed")
    ]

    if mode_cols:
        mode_long = df.melt(
            id_vars=["Office", "MonthDate"],
            value_vars=mode_cols,
            var_name="Mode",
            value_name="Volume",
        )
        mode_long["Volume"] = pd.to_numeric(mode_long["Volume"], errors="coerce").fillna(0)
        mode_long = mode_long[mode_long["Volume"] > 0]
    else:
        mode_long = pd.DataFrame(columns=["Office", "MonthDate", "Mode", "Volume"])

    return df, mode_long


@st.cache_data(show_spinner=False)
def prepare_customer(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    # Prefer office-specific customer sheets to avoid double count with Customer Volume-N&S.
    office_sheets = ["customer_had", "customer_han", "customer_hlc", "customer_hcm"]
    frames = []
    for key in office_sheets:
        df = data.get(key, pd.DataFrame())
        if df.empty:
            continue
        frames.append(customer_wide_to_long(df))
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if combined.empty:
        combined = customer_wide_to_long(data.get("customer_ns", pd.DataFrame()))
    return combined


@st.cache_data(show_spinner=False)
def customer_wide_to_long(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    cust_col = first_existing(df, ["Customer", "CUSTOMER"])
    if not office_col or not cust_col:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    month_cols = [c for c in df.columns if parse_month(c) is not pd.NaT and not pd.isna(parse_month(c))]
    if not month_cols:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    long = df.melt(id_vars=[office_col, cust_col], value_vars=month_cols, var_name="Month", value_name="Volume")
    long["Office"] = long[office_col].map(normalize_office)
    long["Customer"] = long[cust_col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    long["MonthDate"] = long["Month"].map(parse_month)
    long["Volume"] = numeric_series(long["Volume"])
    long = long[(long["Volume"] > 0) & (long["Customer"] != "") & (~long["MonthDate"].isna())]
    return long[["Office", "Customer", "MonthDate", "Volume"]]


@st.cache_data(show_spinner=False)
def prepare_case_detail(
    df: pd.DataFrame,
    activity_type: str,
) -> pd.DataFrame:
    """
    Normalize detail sheets C / A / S / E into long format.

    C/A/S expected pattern:
        Office | Scope/Code | Apr-26 ... Mar-27 | Total

    E expected pattern:
        Office | Code | BU | Criteria | Exception Detail | Apr-26 ... Mar-27 | Total

    The parser is intentionally tolerant to header wording so the dashboard
    remains usable when the source template adds descriptive columns.
    """
    base_cols = [
        "Activity Type", "Office", "Code", "BU", "Criteria",
        "Detail", "MonthDate", "Volume"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()
    d.columns = [clean_col(c) for c in d.columns]

    office_col = first_existing(d, ["Office", "OFFICE"])
    if not office_col:
        return pd.DataFrame(columns=base_cols)

    # Identify month columns strictly by parseable month headers.
    month_cols = []
    for c in d.columns:
        parsed = parse_month(c)
        if not pd.isna(parsed):
            month_cols.append(c)

    if not month_cols:
        return pd.DataFrame(columns=base_cols)

    # Descriptive columns before month columns.
    code_col = first_existing(d, ["Scope", "CODE", "Code", "Service Code"])
    bu_col = first_existing(d, ["BU", "Segment", "Service"])
    criteria_col = first_existing(d, ["Criteria"])
    detail_col = first_existing(
        d,
        [
            "Scope details",       # Sheet A
            "Job details",         # Sheet S
            "EXCEPTION DETAIL",    # Sheet E
            "Exception Detail",
            "Detail",
            "Description",
            "Activity",
        ],
    )

    id_cols = [office_col]
    for c in [code_col, bu_col, criteria_col, detail_col]:
        if c and c not in id_cols:
            id_cols.append(c)

    long = d.melt(
        id_vars=id_cols,
        value_vars=month_cols,
        var_name="Month",
        value_name="Volume",
    )

    long["Office"] = long[office_col].map(normalize_office)
    long["MonthDate"] = long["Month"].map(parse_month)
    long["Volume"] = pd.to_numeric(long["Volume"], errors="coerce")

    long["Code"] = (
        long[code_col].astype(str).str.strip()
        if code_col else ""
    )
    long["BU"] = (
        long[bu_col].astype(str).str.strip().str.upper()
        if bu_col else ""
    )
    long["Criteria"] = (
        long[criteria_col].astype(str).str.strip()
        if criteria_col else ""
    )
    long["Detail"] = (
        long[detail_col].astype(str).str.strip()
        if detail_col else ""
    )
    long["Activity Type"] = activity_type

    # Keep only rows that contain real activity data.
    # Blank cells and 0-volume months are excluded so detail tabs show only months
    # that actually have C/A/S/E data in the corresponding source sheet.
    long = long[
        (long["Office"] != "")
        & (~long["MonthDate"].isna())
        & (long["Volume"].notna())
        & (long["Volume"] > 0)
    ].copy()

    return long[base_cols].reset_index(drop=True)



@st.cache_data(show_spinner=False)
def prepare_code_note_map(df: pd.DataFrame) -> Dict[str, str]:
    """
    Sheet 'Ghi chú':
        Col A = Scope of Job code
        Col B = description.

    Used mainly for Core Service codes such as AE-CTAB:
        suffix CTAB -> "Customs + Trucking + Air"

    A/S/E sheets already contain their own descriptive columns
    (Scope details / Job details / EXCEPTION DETAIL), which take priority.
    """
    if df is None or df.empty or df.shape[1] < 2:
        return {}

    d = df.copy()
    code_col = d.columns[0]
    desc_col = d.columns[1]

    d[code_col] = d[code_col].astype(str).str.strip().str.upper()
    d[desc_col] = d[desc_col].astype(str).str.strip()

    # Remove title/header/blank rows.
    d = d[
        d[code_col].ne("")
        & d[desc_col].ne("")
        & d[code_col].ne("SCOPE OF JOB")
        & d[code_col].ne("NAN")
        & d[desc_col].ne("nan")
    ].copy()

    return dict(zip(d[code_col], d[desc_col]))


def add_code_description(
    df: pd.DataFrame,
    activity_type: str,
    note_map: Dict[str, str],
) -> pd.DataFrame:
    """
    Create the dashboard field 'Code Description'.

    Priority:
    1) A/S/E: source description already available in the corresponding sheet.
    2) C: lookup suffix of Scope code against sheet 'Ghi chú'
       e.g. AE-ABBB -> ABBB -> Air Freight Only.
    3) Fallback: lookup whole code in 'Ghi chú'.
    """
    if df is None or df.empty:
        return df

    d = df.copy()
    d["Code"] = d["Code"].fillna("").astype(str).str.strip()
    d["Detail"] = d["Detail"].fillna("").astype(str).str.strip()

    def _decode(row):
        code = str(row.get("Code", "")).strip().upper()
        source_detail = str(row.get("Detail", "")).strip()

        # Ancillary / Supporting / Exception: use source wording first.
        if activity_type != "Core Service" and source_detail:
            return source_detail

        # Core: AE-CTAB -> CTAB
        suffix = code.split("-", 1)[1] if "-" in code else code
        if suffix in note_map:
            return note_map[suffix]

        if code in note_map:
            return note_map[code]

        # Defensive fallback if a source detail exists.
        return source_detail

    d["Code Description"] = d.apply(_decode, axis=1)
    return d


def workload_breakdown_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    Section 5 summary:
    Segment | Core Service (min) | Ancillary Service (min) |
    Supporting Activity (min) | Exception Handling (min) |
    Total Workload (min) | Ratio
    """
    cols = [
        "Segment",
        "Core Service (min)",
        "Ancillary Service (min)",
        "Supporting Activity (min)",
        "Exception Handling (min)",
        "Total Workload (min)",
        "Ratio",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    d = df.copy()
    for c in [
        "Core Workload (min)",
        "Ancillary Workload (min)",
        "Supporting Workload (min)",
        "Exception Workload (min)",
        "Total Workload (min)",
    ]:
        if c not in d.columns:
            d[c] = 0.0
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)

    agg = (
        d.groupby("Segment", as_index=False)
        .agg({
            "Core Workload (min)": "sum",
            "Ancillary Workload (min)": "sum",
            "Supporting Workload (min)": "sum",
            "Exception Workload (min)": "sum",
            "Total Workload (min)": "sum",
        })
    )

    # Ensure standard business order.
    agg = (
        pd.DataFrame({"Segment": SERVICE_ORDER})
        .merge(agg, on="Segment", how="left")
        .fillna(0)
    )

    # Recalculate Total from C+A+S+E if source Total is blank/zero.
    component_total = (
        agg["Core Workload (min)"]
        + agg["Ancillary Workload (min)"]
        + agg["Supporting Workload (min)"]
        + agg["Exception Workload (min)"]
    )
    agg["Total Workload (min)"] = np.where(
        agg["Total Workload (min)"] > 0,
        agg["Total Workload (min)"],
        component_total,
    )

    grand_total = float(agg["Total Workload (min)"].sum())
    agg["Ratio"] = np.where(
        grand_total > 0,
        agg["Total Workload (min)"] / grand_total,
        0.0,
    )

    agg = agg.rename(columns={
        "Core Workload (min)": "Core Service (min)",
        "Ancillary Workload (min)": "Ancillary Service (min)",
        "Supporting Workload (min)": "Supporting Activity (min)",
        "Exception Workload (min)": "Exception Handling (min)",
    })

    return agg[cols]


def chart_case_allocation(df: pd.DataFrame):
    """
    C/A/S/E allocation by Segment.
    Stacked horizontal bars show both total workload and its activity composition.
    """
    summary = workload_breakdown_table(df)

    if summary.empty or float(summary["Total Workload (min)"].sum()) <= 0:
        st.info("No C/A/S/E workload data available for selected filters.")
        return

    plot_df = summary.copy()
    plot_df = plot_df[plot_df["Total Workload (min)"] > 0].copy()

    # Highest-workload service at the top.
    plot_df = plot_df.sort_values("Total Workload (min)", ascending=True)

    components = [
        ("Core Service (min)", "Core Service", COLORS["blue"]),
        ("Ancillary Service (min)", "Ancillary Service", COLORS["green"]),
        ("Supporting Activity (min)", "Supporting Activity", COLORS["amber"]),
        ("Exception Handling (min)", "Exception Handling", COLORS["red"]),
    ]

    fig = go.Figure()
    for col, label, color in components:
        fig.add_trace(
            go.Bar(
                y=plot_df["Segment"],
                x=plot_df[col],
                name=label,
                orientation="h",
                marker_color=color,
                customdata=np.column_stack([
                    plot_df["Total Workload (min)"],
                    plot_df["Ratio"],
                ]),
                hovertemplate=(
                    f"<b>{label}</b><br>"
                    "Segment: %{y}<br>"
                    "Workload: %{x:,.0f} min<br>"
                    "Segment Total: %{customdata[0]:,.0f} min<br>"
                    "Share of Total: %{customdata[1]:.1%}"
                    "<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        barmode="stack",
        title="C / A / S / E Workload Allocation by Segment",
        xaxis_title="Workload (min)",
        yaxis_title="",
    )
    fig = plotly_layout(
        fig,
        390,
        show_legend=True,
        legend_position="top",
        margin_left=50,
        margin_right=35,
        margin_top=72,
        margin_bottom=48,
    )
    fig.update_xaxes(rangemode="tozero")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_workload_breakdown_table(df: pd.DataFrame):
    summary = workload_breakdown_table(df)

    if summary.empty:
        st.info("No workload breakdown data available for selected filters.")
        return

    display = summary.copy()
    grand_total = float(display["Total Workload (min)"].sum())

    total_row = pd.DataFrame([{
        "Segment": "TOTAL",
        "Core Service (min)": float(display["Core Service (min)"].sum()),
        "Ancillary Service (min)": float(display["Ancillary Service (min)"].sum()),
        "Supporting Activity (min)": float(display["Supporting Activity (min)"].sum()),
        "Exception Handling (min)": float(display["Exception Handling (min)"].sum()),
        "Total Workload (min)": grand_total,
        "Ratio": 1.0 if grand_total > 0 else 0.0,
    }])
    display = pd.concat([display, total_row], ignore_index=True)

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=390,
        column_config={
            "Segment": st.column_config.TextColumn("Segment", width="small"),
            "Core Service (min)": st.column_config.NumberColumn(
                "Core Service (min)", format="%,.0f", width="medium"
            ),
            "Ancillary Service (min)": st.column_config.NumberColumn(
                "Ancillary Service (min)", format="%,.0f", width="medium"
            ),
            "Supporting Activity (min)": st.column_config.NumberColumn(
                "Supporting Activity (min)", format="%,.0f", width="medium"
            ),
            "Exception Handling (min)": st.column_config.NumberColumn(
                "Exception Handling (min)", format="%,.0f", width="medium"
            ),
            "Total Workload (min)": st.column_config.NumberColumn(
                "Total Workload (min)", format="%,.0f", width="medium"
            ),
            "Ratio": st.column_config.NumberColumn(
                "Ratio", format="percent", width="small"
            ),
        },
    )


def render_activity_detail_table(
    df: pd.DataFrame,
    activity_type: str,
):
    """Detail table for one C/A/S/E source sheet."""
    if df is None or df.empty:
        st.info(f"No {activity_type} detail data available for selected filters.")
        return

    d = df.copy()

    # Defensive filter: detail table only shows months/rows with actual data.
    d["Volume"] = pd.to_numeric(d["Volume"], errors="coerce")
    d = d[d["Volume"].fillna(0) > 0].copy()

    if d.empty:
        st.info(f"No {activity_type} detail data available for selected filters.")
        return

    d["Month"] = d["MonthDate"].dt.strftime("%b-%y")

    # Consistent business order requested for all C / A / S / E tabs:
    # Office → Month → Code → Code Description → Volume
    preferred = ["Office", "Month", "Code", "Code Description", "Volume"]
    cols = [c for c in preferred if c in d.columns]

    # Drop descriptive columns that are completely blank.
    cols = [
        c for c in cols
        if c in ["Office", "Month", "Volume"]
        or d[c].astype(str).str.strip().replace("nan", "").ne("").any()
    ]

    # Capture months before removing MonthDate from the visible table.
    months_with_data = (
        d["MonthDate"].dropna().drop_duplicates().sort_values()
        .dt.strftime("%b-%y").tolist()
    )

    sort_source = d.copy()
    sort_cols = [c for c in ["Office", "Code", "MonthDate"] if c in sort_source.columns]
    if sort_cols:
        sort_source = sort_source.sort_values(sort_cols)

    d = sort_source[cols].copy()
    if "Volume" in d.columns:
        d["Volume"] = pd.to_numeric(d["Volume"], errors="coerce").fillna(0)

    if months_with_data:
        st.caption("Months with data: " + ", ".join(months_with_data))

    # Compact detail table: keep the four operational fields narrow and balanced.
    compact_config = {
        "Office": st.column_config.TextColumn("Office", width="small"),
        "Month": st.column_config.TextColumn("Month", width="small"),
        "Code": st.column_config.TextColumn("Code", width="medium"),
        "Code Description": st.column_config.TextColumn(
            "Code Description", width="large"
        ),
        "Volume": st.column_config.NumberColumn(
            "Volume", format="%,.0f", width="small"
        ),
    }

    st.dataframe(
        d,
        use_container_width=False,
        hide_index=True,
        height=min(420, max(160, 38 + len(d) * 34)),
        column_config={c: compact_config[c] for c in d.columns if c in compact_config},
    )


def prepare_resolution(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 'CS Resolutions Rate'
    Source structure:
        OFFICE | Month | Total abnormality/month |
        No of abnormality resolved by CS | CS Resolution rate

    Dashboard rules:
    - Only months with actual source data are retained.
    - Resolution Rate is recalculated from Resolved / Total Abnormality
      so the dashboard does not depend on Excel formula cache.
    """
    cols = [
        "Office", "MonthDate",
        "Total Abnormality", "Resolved", "Resolution Rate",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    d = df.copy()
    office_col = first_existing(d, ["OFFICE", "Office"])
    month_col = first_existing(d, ["Month"])
    total_col = first_existing(
        d, ["Total abnormality/month", "Total abnormality"]
    )
    resolved_col = first_existing(
        d, ["No of abnormality resolved by CS", "Resolved"]
    )

    if not office_col or not month_col or not total_col or not resolved_col:
        return pd.DataFrame(columns=cols)

    d["Office"] = d[office_col].map(normalize_office)
    d["MonthDate"] = d[month_col].map(parse_month)

    # Preserve blanks first; do not convert future empty months to zero.
    d["Total Abnormality"] = pd.to_numeric(d[total_col], errors="coerce")
    d["Resolved"] = pd.to_numeric(d[resolved_col], errors="coerce")

    # Keep only rows/months where the source contains actual activity data.
    d = d[
        (d["Office"] != "")
        & (~d["MonthDate"].isna())
        & (
            d["Total Abnormality"].notna()
            | d["Resolved"].notna()
        )
    ].copy()

    if d.empty:
        return pd.DataFrame(columns=cols)

    d["Total Abnormality"] = d["Total Abnormality"].fillna(0)
    d["Resolved"] = d["Resolved"].fillna(0)

    d["Resolution Rate"] = np.where(
        d["Total Abnormality"] > 0,
        d["Resolved"] / d["Total Abnormality"],
        np.nan,
    )

    return d[cols].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def prepare_yvf(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 'YVF' — source structure:
        OFFICE | [Month, if available] |
        Total YVF booking/month | Total IFF shipment/month | YVF booking ratio

    Rules:
    - Preserve the source structure; do not invent months.
    - If a Month column exists, parse it and allow Month/Year filtering.
    - Rows where both YVF Booking and IFF Shipment are blank are excluded,
      so future empty periods never appear on the dashboard.
    - Recalculate YVF Booking Ratio = YVF Booking / IFF Shipment to avoid
      stale Excel formula-cache values.
    """
    base_cols = [
        "Office", "MonthDate",
        "YVF Booking", "IFF Shipment", "YVF Booking Ratio",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()

    office_col = first_existing(d, ["OFFICE", "Office"])
    month_col = first_existing(d, ["Month"])
    yvf_col = first_existing(
        d, ["Total YVF booking/month", "Total YVF booking"]
    )
    iff_col = first_existing(
        d, ["Total IFF shipment/month", "Total IFF shipment"]
    )

    if not office_col or not yvf_col or not iff_col:
        return pd.DataFrame(columns=base_cols)

    d["Office"] = d[office_col].map(normalize_office)
    d["MonthDate"] = (
        d[month_col].map(parse_month)
        if month_col else pd.NaT
    )

    # Preserve blanks first so empty/future periods are not converted to zeros.
    d["YVF Booking"] = pd.to_numeric(d[yvf_col], errors="coerce")
    d["IFF Shipment"] = pd.to_numeric(d[iff_col], errors="coerce")

    d = d[
        (d["Office"] != "")
        & (
            d["YVF Booking"].notna()
            | d["IFF Shipment"].notna()
        )
    ].copy()

    if d.empty:
        return pd.DataFrame(columns=base_cols)

    d["YVF Booking"] = d["YVF Booking"].fillna(0)
    d["IFF Shipment"] = d["IFF Shipment"].fillna(0)

    d["YVF Booking Ratio"] = np.where(
        d["IFF Shipment"] > 0,
        d["YVF Booking"] / d["IFF Shipment"],
        np.nan,
    )

    return d[base_cols].reset_index(drop=True)



def all_periods(*dfs: pd.DataFrame) -> List[pd.Timestamp]:
    periods = []
    for df in dfs:
        if df is not None and not df.empty and "MonthDate" in df.columns:
            periods.extend(df["MonthDate"].dropna().unique().tolist())
    return sorted(pd.to_datetime(pd.Series(periods)).dropna().drop_duplicates().tolist())


def apply_filters(df: pd.DataFrame, year: str, month: str, office: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if "MonthDate" in out.columns:
        if year != "All":
            out = out[out["MonthDate"].dt.year.astype(str) == year]
        if month != "All":
            target = parse_month(month)
            out = out[out["MonthDate"] == target]
    if office != "All Offices" and "Office" in out.columns:
        out = out[out["Office"] == office]
    return out


def filter_office_only(df: pd.DataFrame, office: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if office != "All Offices" and "Office" in out.columns:
        out = out[out["Office"] == office]
    return out

# ============================================================
# KPI CALCULATION
# ============================================================



def calculate_active_customers(shipment_df: pd.DataFrame) -> float:
    """
    Active Customers source: Shipment volume.
    For a single selected month, this returns the sum across selected offices.
    For multi-month / All selection, it returns the average monthly active-customer total
    because the source sheet contains monthly counts rather than customer-level IDs.
    """
    if shipment_df is None or shipment_df.empty or "Active Customers" not in shipment_df.columns:
        return 0.0

    d = shipment_df[["MonthDate", "Active Customers"]].copy()
    d["Active Customers"] = pd.to_numeric(d["Active Customers"], errors="coerce")
    d = d.dropna(subset=["MonthDate", "Active Customers"])

    if d.empty:
        return 0.0

    monthly = (
        d.groupby("MonthDate", as_index=False)["Active Customers"]
        .sum(min_count=1)
        .dropna(subset=["Active Customers"])
    )
    if monthly.empty:
        return 0.0

    if len(monthly) == 1:
        return float(monthly["Active Customers"].iloc[0])

    return float(monthly["Active Customers"].mean())


def calculate_kpis(hc, workload, fte, shipment) -> Dict[str, float]:
    total_workload_hours = float(workload["Workload Hours"].sum()) if not workload.empty and "Workload Hours" in workload.columns else 0.0
    required_fte_total_period = total_workload_hours / CAPACITY_HOURS_PER_FTE if total_workload_hours else 0.0

    # Capacity is summed over selected months. Actual FTE card is average monthly FTE.
    monthly_fte = pd.DataFrame()
    if not fte.empty:
        monthly_fte = fte.groupby("MonthDate", dropna=True)["Actual FTE"].sum().reset_index()
    actual_fte_avg = float(monthly_fte["Actual FTE"].mean()) if not monthly_fte.empty else 0.0
    capacity_hours = float((monthly_fte["Actual FTE"] * CAPACITY_HOURS_PER_FTE).sum()) if not monthly_fte.empty else 0.0
    required_fte_avg = required_fte_total_period / max(len(monthly_fte), 1) if not monthly_fte.empty else required_fte_total_period

    # HC is average monthly total when multiple months are selected.
    approved_hc = weighted_period_avg(hc, "Total Approved HC") if not hc.empty and "Total Approved HC" in hc.columns else 0.0
    actual_hc = weighted_period_avg(hc, "Total Actual HC") if not hc.empty and "Total Actual HC" in hc.columns else 0.0
    total_shipment = float(shipment["Total Shipment"].sum()) if not shipment.empty and "Total Shipment" in shipment.columns else 0.0
    util = safe_div(total_workload_hours, capacity_hours)
    fte_gap = actual_fte_avg - required_fte_avg
    return {
        "Approved HC": approved_hc,
        "Actual HC": actual_hc,
        "Actual FTE": actual_fte_avg,
        "Required FTE": required_fte_avg,
        "Capacity Hours": capacity_hours,
        "Workload Hours": total_workload_hours,
        "Utilization": util,
        "FTE Gap": fte_gap,
        "Total Shipment": total_shipment,
    }


def build_reconciliation(hc, workload, fte, shipment) -> pd.DataFrame:
    kpis = calculate_kpis(hc, workload, fte, shipment)
    rows = []
    # Reference values from HC if available.
    ref_actual_hc = weighted_period_avg(hc, "Total Actual HC") if not hc.empty and "Total Actual HC" in hc.columns else np.nan
    ref_approved_hc = weighted_period_avg(hc, "Total Approved HC") if not hc.empty and "Total Approved HC" in hc.columns else np.nan
    ref_required_hc = weighted_period_avg(hc, "Total Required HC") if not hc.empty and "Total Required HC" in hc.columns else np.nan
    ref_ship = shipment["Total Shipment"].sum() if not shipment.empty and "Total Shipment" in shipment.columns else np.nan

    def add(kpi, calc, ref, status="PASS", note=""):
        diff = calc - ref if pd.notna(ref) else np.nan
        rows.append({"KPI": kpi, "Calculated Value": calc, "Reference Value": ref, "Difference": diff, "Status": status, "Note": note})

    add("Total Approved HC", kpis["Approved HC"], ref_approved_hc, "PASS" if pd.notna(ref_approved_hc) else "WARNING", "Source: HC")
    add("Total Actual HC", kpis["Actual HC"], ref_actual_hc, "PASS" if pd.notna(ref_actual_hc) else "WARNING", "Source: HC")
    add("Total Shipment", kpis["Total Shipment"], ref_ship, "PASS" if pd.notna(ref_ship) else "WARNING", "Source: Shipment volume")
    add("Required FTE", kpis["Required FTE"], ref_required_hc, "WARNING", "Dashboard calculates from Workload Hours / 167.2; HC value is reference only.")
    add("Utilization", kpis["Utilization"], np.nan, "WARNING", "Dashboard formula: Workload Hours / Capacity Hours.")
    return pd.DataFrame(rows)

# ============================================================
# CHARTS
# ============================================================



def chart_office_capacity_trend(df: pd.DataFrame):
    """
    Executive HC trend from sheet HC.

    Presentation only:
    - Actual HC = Corporate Blue
    - Required HC = Orange
    - Approved HC = Navy dashed reference line
    - Light orange fill = shortage gap between Actual and Required HC
    - Required HC labels show both required value and HC gap
    """
    if df.empty:
        st.info("No HC trend data available for selected filters.")
        return

    required_cols = [
        "MonthDate",
        "Total Approved HC",
        "Total Actual HC",
        "Total Required HC",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.info("HC trend cannot be displayed because required HC columns are missing.")
        return

    trend_source = df[required_cols].copy()

    for col in [
        "Total Approved HC",
        "Total Actual HC",
        "Total Required HC",
    ]:
        trend_source[col] = pd.to_numeric(
            trend_source[col], errors="coerce"
        )

    # Exclude months where all three HC values are blank.
    trend_source = trend_source.dropna(
        subset=[
            "Total Approved HC",
            "Total Actual HC",
            "Total Required HC",
        ],
        how="all",
    )

    if trend_source.empty:
        st.info("No HC trend data available for selected filters.")
        return

    trend = (
        trend_source.groupby("MonthDate", as_index=False)[
            [
                "Total Approved HC",
                "Total Actual HC",
                "Total Required HC",
            ]
        ]
        .sum(min_count=1)
        .sort_values("MonthDate")
    )

    trend["Month"] = trend["MonthDate"].dt.strftime("%b-%y")
    trend["HC Gap"] = (
        trend["Total Required HC"] - trend["Total Actual HC"]
    )

    fig = go.Figure()

    # Actual HC — primary current-capacity line.
    # Draw first so Required HC can shade only the gap above/below Actual HC.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Actual HC"],
            mode="lines+markers",
            name="Actual HC",
            line=dict(
                color=BUSINESS_COLORS["actual"],
                width=3,
            ),
            marker=dict(
                size=8,
                symbol="circle",
                color=BUSINESS_COLORS["actual"],
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Actual HC: %{y:,.1f}"
                "<extra></extra>"
            ),
        )
    )

    # Required HC — demand line.
    # fill='tonexty' shades ONLY the gap between Actual and Required HC.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Required HC"],
            mode="lines+markers+text",
            name="Required HC",
            line=dict(
                color=BUSINESS_COLORS["required"],
                width=3,
            ),
            marker=dict(
                size=8,
                symbol="circle",
                color=BUSINESS_COLORS["required"],
            ),
            fill="tonexty",
            fillcolor="rgba(245, 158, 11, 0.08)",
            text=[
                f"{req:,.2f}<br>Gap {gap:+,.2f}"
                for req, gap in zip(
                    trend["Total Required HC"],
                    trend["HC Gap"],
                )
            ],
            textposition="top center",
            textfont=dict(
                size=10,
                color=BUSINESS_COLORS["required"],
                family=UI["font_family"],
            ),
            customdata=trend[["HC Gap"]].to_numpy(),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Required HC: %{y:,.2f}<br>"
                "HC Gap: %{customdata[0]:+,.2f}"
                "<extra></extra>"
            ),
        )
    )

    # Approved HC — reference baseline.
    # Draw last so it remains visible even when Approved HC = Actual HC.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Approved HC"],
            mode="lines+markers",
            name="Approved HC",
            line=dict(
                color=BUSINESS_COLORS["approved"],
                width=2,
                dash="dash",
            ),
            marker=dict(
                size=7,
                symbol="circle-open",
                color=BUSINESS_COLORS["approved"],
                line=dict(
                    color=BUSINESS_COLORS["approved"],
                    width=1.5,
                ),
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Approved HC: %{y:,.1f}"
                "<extra></extra>"
            ),
        )
    )

    fig.update_layout(
        title="Actual vs Required HC Trend",
        yaxis_title="HC",
        hovermode="x unified",
    )

    fig = plotly_layout(
        fig,
        UI["chart_height"],
        show_legend=True,
        legend_position="top",
        margin_left=56,
        margin_right=42,
        margin_top=78,
        margin_bottom=46,
    )

    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=trend["Month"].tolist(),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )



def chart_workload_by_pic(fte_df: pd.DataFrame, selected_office: str):
    """
    PIC Workload & Capacity Utilization.

    Business display rule:
    - PIC Workload (hrs) = CS FTE Factor × Available Standard Time / PIC.
    - Available Standard Time / PIC = 167.2 hrs/month.
    - Utilization = PIC Workload / 167.2 = CS FTE Factor.

    Display logic:
    - Specific Office: show all PICs with data in that office.
    - All Offices: show Top 10 PICs by Utilization across all offices.
    - Colors:
        >100%   = Red (Overload)
        90–100% = Orange (Attention)
        <90%    = Blue (Available Capacity)
    """
    if fte_df is None or fte_df.empty:
        st.info("No CS FTE data available for selected filters.")
        return

    d = fte_df.copy()
    d["Actual FTE"] = pd.to_numeric(d["Actual FTE"], errors="coerce")
    d = d.dropna(subset=["Office", "CS PIC", "Actual FTE"])
    d = d[(d["Office"] != "") & (d["CS PIC"] != "")]

    if d.empty:
        st.info("No CS FTE data available for selected filters.")
        return

    pic_data = (
        d.groupby(["Office", "CS PIC"], as_index=False)["Actual FTE"]
        .mean()
    )
    pic_data["Standard Hours"] = CAPACITY_HOURS_PER_FTE
    pic_data["Actual Workload Hours"] = pic_data["Actual FTE"] * CAPACITY_HOURS_PER_FTE
    pic_data["Utilization"] = pic_data["Actual FTE"]

    def _status(util):
        if util > 1.0:
            return "Overload", COLORS["red"]
        if util >= 0.90:
            return "Attention", COLORS["amber"]
        return "Available", COLORS["blue"]

    mapped = pic_data["Utilization"].apply(_status)
    pic_data["Status"] = mapped.map(lambda x: x[0])
    pic_data["Bar Color"] = mapped.map(lambda x: x[1])

    total_pic = int(len(pic_data))
    overloaded_pic = int((pic_data["Utilization"] > 1.0).sum())

    if selected_office == "All Offices":
        display = (
            pic_data.sort_values(
                ["Utilization", "Actual Workload Hours"],
                ascending=[False, False],
            )
            .head(10)
            .copy()
        )
        display["PIC Label"] = display.apply(
            lambda r: f"{r['Office']} | {r['CS PIC']}",
            axis=1,
        )
        subtitle = "Top 10 PICs by Capacity Utilization – All Offices"
    else:
        display = (
            pic_data[pic_data["Office"] == selected_office]
            .sort_values(
                ["Utilization", "Actual Workload Hours"],
                ascending=[False, False],
            )
            .copy()
        )
        display["PIC Label"] = display["CS PIC"].astype(str)
        subtitle = f"All PICs – {selected_office}"

    if display.empty:
        st.info("No PIC workload data available for selected filters.")
        return

    display = display.sort_values(
        ["Utilization", "Actual Workload Hours"],
        ascending=[True, True],
    )

    display["Label"] = display.apply(
        lambda r: f"{r['Actual Workload Hours']:,.1f} hrs | {r['Utilization']*100:.0f}%",
        axis=1,
    )

    chart_height = max(360, min(650, 43 * len(display) + 150))

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=display["Actual Workload Hours"],
            y=display["PIC Label"],
            orientation="h",
            name="PIC Workload",
            marker_color=display["Bar Color"],
            text=display["Label"],
            textposition="outside",
            cliponaxis=False,
            customdata=np.column_stack([
                display["Office"],
                display["Actual FTE"],
                display["Utilization"],
                display["Status"],
            ]),
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Office: %{customdata[0]}<br>"
                "PIC Workload: %{x:,.1f} hrs<br>"
                "CS FTE Factor: %{customdata[1]:.2f}<br>"
                "Utilization: %{customdata[2]:.1%}<br>"
                "Status: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    fig.add_vline(
        x=CAPACITY_HOURS_PER_FTE,
        line_width=2.5,
        line_dash="dash",
        line_color=COLORS["navy"],
        annotation_text="Standard 167.2 h",
        annotation_position="top",
        annotation_font_color=COLORS["navy"],
    )

    max_actual = float(display["Actual Workload Hours"].max())
    x_max = max(max_actual * 1.15, CAPACITY_HOURS_PER_FTE * 1.25)

    fig.update_layout(
        title=dict(
            text=(
                "PIC Workload & Capacity Utilization"
                f"<br><span style='font-size:11px;color:#667085;font-weight:400'>{subtitle}</span>"
            ),
            x=0.0,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(size=UI["chart_title_size"], color=COLORS["navy"]),
        ),
        height=chart_height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", color=COLORS["text"]),
        margin=dict(l=125, r=80, t=76, b=48),
        bargap=0.24,
        showlegend=False,
    )

    # Keep only one compact management summary in the upper-right.
    existing_annotations = list(fig.layout.annotations) if fig.layout.annotations else []
    existing_annotations.append(
        dict(
            text=f"Overloaded PICs: <b>{overloaded_pic}</b> / Total PICs: <b>{total_pic}</b>",
            x=1,
            y=1.075,
            xref="paper",
            yref="paper",
            showarrow=False,
            xanchor="right",
            yanchor="bottom",
            font=dict(size=UI["note_size"], color=COLORS["muted"]),
        )
    )
    fig.update_layout(annotations=existing_annotations)

    fig.update_xaxes(
        title_text="Workload Hours",
        range=[0, x_max],
        gridcolor=COLORS["grid"],
        zeroline=False,
        automargin=True,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
    )
    fig.update_yaxes(
        title_text="",
        gridcolor="rgba(0,0,0,0)",
        automargin=True,
        tickfont=dict(size=UI["axis_size"]),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )

    st.markdown(
        f"""
        <div style="
            display:flex;
            justify-content:flex-end;
            align-items:center;
            gap:18px;
            margin-top:4px;
            margin-bottom:2px;
            color:#667085;
            font-size:11px;
            line-height:1.2;
            white-space:nowrap;">
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['red']};margin-right:5px;border-radius:2px;"></span>Overload &gt;100%</span>
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['amber']};margin-right:5px;border-radius:2px;"></span>Attention 90–100%</span>
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['blue']};margin-right:5px;border-radius:2px;"></span>Available &lt;90%</span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def chart_workload_by_service(df: pd.DataFrame):
    if df.empty:
        st.info("No workload data available for selected filters.")
        return
    agg = df.groupby(["Segment", "Service Label"], as_index=False)["Workload Hours"].sum()
    total = agg["Workload Hours"].sum()
    agg["% of Total"] = agg["Workload Hours"].apply(lambda x: safe_div(x, total))
    agg["Label"] = agg.apply(lambda r: f"{r['Workload Hours']:,.1f} hrs | {r['% of Total']*100:.1f}%", axis=1)
    agg["SortOrder"] = agg["Segment"].apply(lambda x: SERVICE_ORDER.index(x) if x in SERVICE_ORDER else 999)
    agg = agg.sort_values(["Workload Hours"], ascending=True)
    fig = px.bar(
        agg,
        x="Workload Hours",
        y="Service Label",
        orientation="h",
        text="Label",
        color_discrete_sequence=[COLORS["blue"]],
        title="Workload Breakdown by Service Type",
    )
    fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>%{x:,.1f} hours<extra></extra>")
    fig = plotly_layout(fig, UI["chart_height"], show_legend=False, margin_left=110, margin_right=70, margin_top=64, margin_bottom=44)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_workload_composition(df: pd.DataFrame):
    if df.empty:
        st.info("No workload composition data available.")
        return
    comp = pd.DataFrame({
        "Activity": ["Core", "Ancillary", "Supporting", "Exception"],
        "Hours": [df["Core Hours"].sum(), df["Ancillary Hours"].sum(), df["Supporting Hours"].sum(), df["Exception Hours"].sum()],
    })
    total = comp["Hours"].sum()
    comp["Share"] = comp["Hours"].apply(lambda x: safe_div(x, total))
    fig = go.Figure()
    color_map = [COLORS["blue"], COLORS["green"], COLORS["amber"], COLORS["red"]]
    for _, row in comp.iterrows():
        color = color_map[comp.index[comp["Activity"] == row["Activity"]][0]]
        fig.add_trace(go.Bar(
            y=["Total Workload"],
            x=[row["Share"]],
            name=row["Activity"],
            orientation="h",
            marker_color=color,
            text=[f"{row['Activity']}<br>{row['Share']*100:.1f}%" if row["Share"] > 0.06 else ""],
            hovertemplate=f"{row['Activity']}: {row['Hours']:,.1f} hrs ({row['Share']*100:.1f}%)<extra></extra>",
        ))
    fig.update_layout(barmode="stack", xaxis_tickformat=".0%", title="Workload Composition – C/A/S/E")
    fig = plotly_layout(fig, 320, show_legend=True, legend_position="top", margin_left=52, margin_right=40, margin_top=66, margin_bottom=44)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_workload_trend(df: pd.DataFrame):
    if df.empty:
        st.info("No trend data available.")
        return
    agg = df.groupby(["MonthDate", "Office"], as_index=False)["Workload Hours"].sum().sort_values("MonthDate")
    agg["Month"] = agg["MonthDate"].dt.strftime("%b-%y")
    fig = px.line(
        agg,
        x="Month",
        y="Workload Hours",
        color="Office",
        markers=True,
        color_discrete_sequence=[COLORS["blue"], COLORS["green"], COLORS["amber"], COLORS["red"]],
        title="Monthly Workload Trend",
    )
    fig.update_traces(hovertemplate="%{fullData.name}<br>%{x}: %{y:,.1f} hrs<extra></extra>")
    fig.update_xaxes(categoryorder="array", categoryarray=agg["Month"].drop_duplicates().tolist())
    fig = plotly_layout(fig, UI["chart_height"], show_legend=True, legend_position="top", margin_left=56, margin_right=40, margin_top=66, margin_bottom=48)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_capacity_trend(workload: pd.DataFrame, fte: pd.DataFrame):
    if workload.empty and fte.empty:
        st.info("No capacity data available.")
        return
    wl = workload.groupby("MonthDate", as_index=False)["Workload Hours"].sum() if not workload.empty else pd.DataFrame(columns=["MonthDate", "Workload Hours"])
    ft = fte.groupby("MonthDate", as_index=False)["Actual FTE"].sum() if not fte.empty else pd.DataFrame(columns=["MonthDate", "Actual FTE"])
    cap = pd.merge(wl, ft, on="MonthDate", how="outer").fillna(0)
    cap["Capacity Hours"] = cap["Actual FTE"] * CAPACITY_HOURS_PER_FTE
    cap["Month"] = cap["MonthDate"].dt.strftime("%b-%y")
    cap = cap.sort_values("MonthDate")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cap["Month"], y=cap["Capacity Hours"], name="Capacity Hours", marker_color=COLORS["light_blue"]))
    fig.add_trace(go.Scatter(x=cap["Month"], y=cap["Workload Hours"], name="Workload Hours", mode="lines+markers", line=dict(color=COLORS["red"], width=3)))
    fig.update_layout(title="Workload vs Capacity Trend", yaxis_title="Hours")
    fig.update_xaxes(categoryorder="array", categoryarray=cap["Month"].tolist())
    fig = plotly_layout(fig, UI["chart_height"], show_legend=True, legend_position="top", margin_left=58, margin_right=40, margin_top=66, margin_bottom=48)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def build_segment_workload(
    df: pd.DataFrame,
    mode_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Section 4 source table — Workload by Segment.

    Business fields:
    - Allocation Time (h): total workload hours from BU allocation.
    - Workload Share (%): Segment workload / total workload.
    - Required FTE:
        Prefer source field "Office HC Allocation Ratio" when available.
        For multiple months, calculate monthly Segment FTE and average across
        valid months because Required FTE is a monthly capacity requirement,
        not a cumulative-period quantity.
        Fallback = monthly Workload Hours / 167.2 hours/FTE.
    - Shipment Volume: Shipment volume sheet, mapped to AE/AI/OE/OI/CC/TR/WH.
    """
    base_cols = [
        "Segment", "Allocation Time (h)", "Workload Share",
        "Required FTE", "Shipment Volume"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()
    if "Segment" not in d.columns or "Workload Hours" not in d.columns:
        return pd.DataFrame(columns=base_cols)

    d["Segment"] = d["Segment"].astype(str).str.strip().str.upper()
    d["Workload Hours"] = pd.to_numeric(d["Workload Hours"], errors="coerce").fillna(0)
    d = d[d["Segment"].isin(SERVICE_ORDER)].copy()

    if d.empty:
        return pd.DataFrame(columns=base_cols)

    seg = (
        d.groupby("Segment", as_index=False)["Workload Hours"]
        .sum()
        .rename(columns={"Workload Hours": "Allocation Time (h)"})
    )

    seg = (
        pd.DataFrame({"Segment": SERVICE_ORDER})
        .merge(seg, on="Segment", how="left")
        .fillna({"Allocation Time (h)": 0.0})
    )

    total_hours = float(seg["Allocation Time (h)"].sum())
    seg["Workload Share"] = np.where(
        total_hours > 0,
        seg["Allocation Time (h)"] / total_hours,
        0.0,
    )

    fte_by_segment = pd.DataFrame({"Segment": SERVICE_ORDER, "Required FTE": 0.0})

    if "MonthDate" in d.columns and d["MonthDate"].notna().any():
        monthly = d.copy()

        use_source_fte = (
            "Office HC Allocation Ratio" in monthly.columns
            and pd.to_numeric(
                monthly["Office HC Allocation Ratio"], errors="coerce"
            ).fillna(0).abs().sum() > 0
        )

        if use_source_fte:
            monthly["Required FTE Source"] = pd.to_numeric(
                monthly["Office HC Allocation Ratio"], errors="coerce"
            )
            monthly_fte = (
                monthly.dropna(subset=["MonthDate"])
                .groupby(["MonthDate", "Segment"], as_index=False)["Required FTE Source"]
                .sum(min_count=1)
                .rename(columns={"Required FTE Source": "Required FTE"})
            )
        else:
            monthly_hours = (
                monthly.dropna(subset=["MonthDate"])
                .groupby(["MonthDate", "Segment"], as_index=False)["Workload Hours"]
                .sum()
            )
            monthly_hours["Required FTE"] = (
                monthly_hours["Workload Hours"] / CAPACITY_HOURS_PER_FTE
            )
            monthly_fte = monthly_hours[["MonthDate", "Segment", "Required FTE"]]

        if not monthly_fte.empty:
            fte_by_segment = (
                monthly_fte.groupby("Segment", as_index=False)["Required FTE"]
                .mean()
            )
            fte_by_segment = (
                pd.DataFrame({"Segment": SERVICE_ORDER})
                .merge(fte_by_segment, on="Segment", how="left")
                .fillna({"Required FTE": 0.0})
            )
    else:
        fte_by_segment = seg[["Segment", "Allocation Time (h)"]].copy()
        fte_by_segment["Required FTE"] = (
            fte_by_segment["Allocation Time (h)"] / CAPACITY_HOURS_PER_FTE
        )
        fte_by_segment = fte_by_segment[["Segment", "Required FTE"]]

    seg = seg.merge(fte_by_segment, on="Segment", how="left")
    seg["Required FTE"] = pd.to_numeric(
        seg["Required FTE"], errors="coerce"
    ).fillna(0)

    seg["Shipment Volume"] = 0.0
    if (
        mode_df is not None
        and not mode_df.empty
        and {"Mode", "Volume"}.issubset(mode_df.columns)
    ):
        vol = mode_df.copy()
        vol["Mode"] = vol["Mode"].astype(str).str.strip().str.upper()
        vol["Volume"] = pd.to_numeric(vol["Volume"], errors="coerce").fillna(0)

        volume_segment_map = {
            "AE": "AE", "AI": "AI", "OE": "OE", "OI": "OI",
            "OEFCL": "OE", "OELCL": "OE", "OIFCL": "OI", "OILCL": "OI",
            "CC": "CC", "CE": "CC", "CI": "CC",
            "TR": "TR", "DM": "TR", "DE": "TR", "DI": "TR",
            "WH": "WH", "HE": "WH", "HI": "WH",
        }
        vol["Segment"] = vol["Mode"].map(volume_segment_map)

        volume_by_segment = (
            vol.dropna(subset=["Segment"])
            .groupby("Segment", as_index=False)["Volume"]
            .sum()
            .rename(columns={"Volume": "Shipment Volume Source"})
        )
        seg = seg.merge(volume_by_segment, on="Segment", how="left")
        seg["Shipment Volume"] = pd.to_numeric(
            seg["Shipment Volume Source"], errors="coerce"
        ).fillna(0)
        seg = seg.drop(columns=["Shipment Volume Source"])

    seg["Segment"] = pd.Categorical(
        seg["Segment"], categories=SERVICE_ORDER, ordered=True
    )
    seg = seg.sort_values("Segment").reset_index(drop=True)
    seg["Segment"] = seg["Segment"].astype(str)

    return seg[
        ["Segment", "Allocation Time (h)", "Workload Share", "Required FTE", "Shipment Volume"]
    ]


def chart_service_matrix(
    df: pd.DataFrame,
    mode_df: Optional[pd.DataFrame] = None,
):
    """
    Workload by Segment — flower-style packed bubble chart.

    One bubble = one Segment.
    Bubble size = Workload Share (%).
    Bubbles are positioned close together for quick visual comparison.
    Axes are intentionally hidden for an executive view.
    """
    seg = build_segment_workload(df, mode_df)

    if seg.empty or float(seg["Allocation Time (h)"].sum()) <= 0:
        st.info("No segment workload data available for selected filters.")
        return

    # Rank only for visual placement; table keeps SERVICE_ORDER.
    plot_df = seg[seg["Allocation Time (h)"] > 0].copy()
    plot_df = plot_df.sort_values("Workload Share", ascending=False).reset_index(drop=True)

    # Flower-like positions:
    # largest bubble in center, remaining bubbles distributed tightly around it.
    flower_positions = [
        (0.00, 0.00),     # center
        (-1.55, 0.30),    # left
        (1.55, 0.30),     # right
        (-0.85, 1.35),    # upper-left
        (0.85, 1.35),     # upper-right
        (-0.75, -1.30),   # lower-left
        (0.75, -1.30),    # lower-right
        (0.00, 2.20),
        (0.00, -2.20),
        (2.25, -0.75),
    ]
    plot_df["x"] = [flower_positions[i][0] for i in range(len(plot_df))]
    plot_df["y"] = [flower_positions[i][1] for i in range(len(plot_df))]

    max_share = float(plot_df["Workload Share"].max())
    if max_share > 0:
        # Tighter size range so bubbles can sit closer without excessive overlap.
        plot_df["Bubble Size"] = 54 + (plot_df["Workload Share"] / max_share) * 76
    else:
        plot_df["Bubble Size"] = 70

    segment_color_map = {
        svc: CORPORATE_PALETTE[i % len(CORPORATE_PALETTE)]
        for i, svc in enumerate(SERVICE_ORDER)
    }

    fig = go.Figure()

    for _, r in plot_df.iterrows():
        svc = r["Segment"]
        fig.add_trace(
            go.Scatter(
                x=[r["x"]],
                y=[r["y"]],
                mode="markers+text",
                name=svc,
                text=[f"<b>{svc}</b><br>{r['Workload Share']:.1%}"],
                textposition="middle center",
                textfont=dict(
                    family="Arial",
                    size=11,
                    color="#FFFFFF" if r["Workload Share"] >= 0.06 else COLORS["navy"],
                ),
                marker=dict(
                    size=[r["Bubble Size"]],
                    color=segment_color_map.get(svc, COLORS["blue"]),
                    opacity=0.94,
                    line=dict(color="#FFFFFF", width=2.5),
                ),
                customdata=[[
                    r["Shipment Volume"],
                    r["Allocation Time (h)"],
                    r["Required FTE"],
                    r["Workload Share"],
                ]],
                hovertemplate=(
                    f"<b>{svc}</b><br>"
                    "Shipment Volume: %{customdata[0]:,.0f}<br>"
                    "Allocation Time: %{customdata[1]:,.1f} hrs<br>"
                    "Required FTE: %{customdata[2]:,.2f}<br>"
                    "Workload Share: %{customdata[3]:.1%}"
                    "<extra></extra>"
                ),
                showlegend=False,
            )
        )

    fig.update_layout(title="Workload by Segment")
    fig = plotly_layout(
        fig,
        390,
        show_legend=False,
        margin_left=20,
        margin_right=20,
        margin_top=58,
        margin_bottom=18,
    )

    fig.update_xaxes(
        visible=False,
        showgrid=False,
        zeroline=False,
        showticklabels=False,
        title_text="",
        range=[-2.8, 2.8],
        fixedrange=True,
    )
    fig.update_yaxes(
        visible=False,
        showgrid=False,
        zeroline=False,
        showticklabels=False,
        title_text="",
        range=[-2.55, 2.65],
        scaleanchor="x",
        scaleratio=1,
        fixedrange=True,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )


def segment_workload_table(df: pd.DataFrame, mode_df: pd.DataFrame):
    """
    Executive summary table for Section 4.

    Order:
    Segment | Shipment Volume | Allocation Time (h) | Required FTE | Workload Share (%)
    """
    seg = build_segment_workload(df, mode_df)

    if seg.empty:
        st.info("No segment workload data available for selected filters.")
        return

    display = seg.copy().rename(columns={
        "Workload Share": "Workload Share (%)",
    })

    # Display percentage as percentage-point values.
    display["Workload Share (%)"] = (
        pd.to_numeric(display["Workload Share (%)"], errors="coerce").fillna(0) * 100
    )

    total_hours = float(display["Allocation Time (h)"].sum())
    total_volume = float(display["Shipment Volume"].sum())
    total_required_fte = float(display["Required FTE"].sum())

    total_row = pd.DataFrame([{
        "Segment": "TOTAL",
        "Shipment Volume": total_volume,
        "Allocation Time (h)": total_hours,
        "Required FTE": total_required_fte,
        "Workload Share (%)": 100.0 if total_hours > 0 else 0.0,
    }])

    display = pd.concat([display, total_row], ignore_index=True)

    # Requested business order.
    display = display[
        ["Segment", "Shipment Volume", "Allocation Time (h)", "Required FTE", "Workload Share (%)"]
    ]

    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:{UI['chart_title_size']}px;
            font-weight:700;
            margin:2px 0 10px 2px;">
            Segment Workload Summary
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=390,
        column_config={
            "Segment": st.column_config.TextColumn(
                "Segment", width="small"
            ),
            "Shipment Volume": st.column_config.NumberColumn(
                "Shipment Volume", width="medium", format="%,.0f"
            ),
            "Allocation Time (h)": st.column_config.NumberColumn(
                "Allocation Time (hrs)", width="medium", format="%,.1f"
            ),
            "Required FTE": st.column_config.NumberColumn(
                "Required FTE", width="small", format="%.2f"
            ),
            "Workload Share (%)": st.column_config.NumberColumn(
                "Workload Share (%)", width="medium", format="%.1f%%"
            ),
        },
    )


def chart_shipment_modes(df: pd.DataFrame):
    """
    Executive transportation-mode view.
    Horizontal ranking bar is used instead of a donut because the source
    contains many modes and small shares that would otherwise overlap.
    """
    if df is None or df.empty:
        st.info("No shipment volume data available for selected filters.")
        return

    d = df.copy()
    # Source field from prepare_shipment_data(): "Volume"
    d["Volume"] = pd.to_numeric(
        d["Volume"], errors="coerce"
    ).fillna(0)

    d = d[d["Volume"] > 0].copy()
    if d.empty:
        st.info("No shipment volume data available for selected filters.")
        return

    agg = (
        d.groupby("Mode", as_index=False)["Volume"]
        .sum()
        .sort_values("Volume", ascending=True)
    )

    total = float(agg["Volume"].sum())
    agg["Share"] = np.where(
        total > 0,
        agg["Volume"] / total,
        0,
    )

    fig = go.Figure(
        go.Bar(
            x=agg["Volume"],
            y=agg["Mode"],
            orientation="h",
            marker_color=BUSINESS_COLORS["actual"],
            customdata=agg[["Share"]].to_numpy(),
            text=[
                f"{v:,.0f}  |  {s:.1%}"
                for v, s in zip(
                    agg["Volume"],
                    agg["Share"],
                )
            ],
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Shipment Volume: %{x:,.0f}<br>"
                "Share: %{customdata[0]:.1%}"
                "<extra></extra>"
            ),
        )
    )

    fig.update_layout(
        title="Shipment Volume by Transportation Mode",
        xaxis_title="Shipment Volume",
        yaxis_title="",
    )
    fig = plotly_layout(
        fig,
        420,
        show_legend=False,
        margin_left=70,
        margin_right=90,
        margin_top=62,
        margin_bottom=48,
    )
    fig.update_yaxes(
        categoryorder="array",
        categoryarray=agg["Mode"].tolist(),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )



def chart_top_customers(df: pd.DataFrame):
    """
    Executive customer concentration view.
    Top 10 is shown for faster management reading; source/filter logic unchanged.
    """
    if df is None or df.empty:
        st.info("No customer shipment data available for selected filters.")
        return

    d = df.copy()
    # Source field from customer-volume preparation: "Volume"
    d["Volume"] = pd.to_numeric(
        d["Volume"], errors="coerce"
    ).fillna(0)

    top = (
        d.groupby("Customer", as_index=False)["Volume"]
        .sum()
        .sort_values("Volume", ascending=False)
        .head(10)
        .sort_values("Volume", ascending=True)
    )

    if top.empty:
        st.info("No customer shipment data available for selected filters.")
        return

    fig = go.Figure(
        go.Bar(
            x=top["Volume"],
            y=top["Customer"],
            orientation="h",
            marker_color=BUSINESS_COLORS["actual"],
            text=top["Volume"],
            texttemplate="%{text:,.0f}",
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Shipment Volume: %{x:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    fig.update_layout(
        title="Top 10 Customers by Shipment Volume",
        xaxis_title="Shipment Volume",
        yaxis_title="",
    )
    fig = plotly_layout(
        fig,
        420,
        show_legend=False,
        margin_left=115,
        margin_right=70,
        margin_top=62,
        margin_bottom=48,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )



def chart_resolution(df: pd.DataFrame):
    """
    CS Solution performance:
    - Bars = Total Abnormalities vs Resolved by CS
    - Line = CS Resolution Rate
    - Ordered by month
    """
    if df is None or df.empty:
        st.info("No CS Resolution data available for selected filters.")
        return

    agg = (
        df.groupby("MonthDate", as_index=False)
        .agg(
            **{
                "Total Abnormality": ("Total Abnormality", "sum"),
                "Resolved": ("Resolved", "sum"),
            }
        )
        .sort_values("MonthDate")
    )
    agg["Resolution Rate"] = np.where(
        agg["Total Abnormality"] > 0,
        agg["Resolved"] / agg["Total Abnormality"],
        np.nan,
    )
    agg["Month"] = agg["MonthDate"].dt.strftime("%b-%y")

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=agg["Month"],
            y=agg["Total Abnormality"],
            name="Total Abnormalities",
            marker_color=BUSINESS_COLORS["supporting"],
            text=agg["Total Abnormality"],
            texttemplate="%{text:,.0f}",
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Total Abnormalities: %{y:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    fig.add_trace(
        go.Bar(
            x=agg["Month"],
            y=agg["Resolved"],
            name="Resolved by CS",
            marker_color=BUSINESS_COLORS["actual"],
            text=agg["Resolved"],
            texttemplate="%{text:,.0f}",
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Resolved by CS: %{y:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    fig.add_trace(
        go.Scatter(
            x=agg["Month"],
            y=agg["Resolution Rate"],
            name="CS Resolution Rate",
            mode="lines+markers+text",
            line=dict(color=COLORS["green"], width=3),
            marker=dict(size=7),
            text=agg["Resolution Rate"],
            texttemplate="%{text:.1%}",
            textposition="top center",
            yaxis="y2",
            hovertemplate=(
                "<b>%{x}</b><br>"
                "CS Resolution Rate: %{y:.1%}"
                "<extra></extra>"
            ),
        )
    )

    fig.update_layout(
        title="CS Resolution Performance by Month",
        barmode="group",
        yaxis=dict(
            title="Cases",
            rangemode="tozero",
        ),
        yaxis2=dict(
            title="Resolution Rate",
            overlaying="y",
            side="right",
            tickformat=".0%",
            range=[0, 1.08],
            showgrid=False,
        ),
    )

    fig = plotly_layout(
        fig,
        390,
        show_legend=True,
        legend_position="top",
        margin_left=58,
        margin_right=68,
        margin_top=72,
        margin_bottom=44,
    )
    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=agg["Month"].tolist(),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )


def render_cs_solution_table(df: pd.DataFrame):
    """
    Detail table sourced directly from sheet 'CS Resolutions Rate'.
    Shows only rows/months with source data after dashboard filters.
    """
    if df is None or df.empty:
        st.info("No CS Resolution data available for selected filters.")
        return

    d = df.copy()
    d = d.sort_values(["Office", "MonthDate"]).copy()
    d["Month"] = d["MonthDate"].dt.strftime("%b-%y")

    display = d[
        ["Office", "Month", "Total Abnormality", "Resolved", "Resolution Rate"]
    ].copy()

    total_abn = float(display["Total Abnormality"].sum())
    total_resolved = float(display["Resolved"].sum())
    overall_rate = safe_div(total_resolved, total_abn)

    total_row = pd.DataFrame([{
        "Office": "TOTAL",
        "Month": "",
        "Total Abnormality": total_abn,
        "Resolved": total_resolved,
        "Resolution Rate": overall_rate,
    }])
    display = pd.concat([display, total_row], ignore_index=True)

    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:{UI['chart_title_size']}px;
            font-weight:700;
            margin:2px 0 10px 2px;">
            CS Resolution Detail
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=360,
        column_config={
            "Office": st.column_config.TextColumn(
                "Office", width="small"
            ),
            "Month": st.column_config.TextColumn(
                "Month", width="small"
            ),
            "Total Abnormality": st.column_config.NumberColumn(
                "Total Abnormalities", width="medium", format="%,.0f"
            ),
            "Resolved": st.column_config.NumberColumn(
                "Resolved by CS", width="medium", format="%,.0f"
            ),
            "Resolution Rate": st.column_config.NumberColumn(
                "CS Resolution Rate", width="medium", format="percent"
            ),
        },
    )


def chart_yvf(df: pd.DataFrame):
    """
    YVF Promoter Effectiveness — donut chart.
    Shows Total YVF Bookings as the adopted portion of Total IFF Shipments.

    IMPORTANT:
    YVF Bookings are a subset of IFF Shipments, therefore the two raw totals
    must not be used as two independent pie slices.
    Pie composition:
        YVF Bookings
        Remaining IFF Shipments = Total IFF Shipments - YVF Bookings
    """
    if df is None or df.empty:
        st.info("No YVF data available for selected filters.")
        return

    d = df.copy()
    d["YVF Booking"] = pd.to_numeric(d["YVF Booking"], errors="coerce").fillna(0)
    d["IFF Shipment"] = pd.to_numeric(d["IFF Shipment"], errors="coerce").fillna(0)

    d = d[(d["YVF Booking"] != 0) | (d["IFF Shipment"] != 0)].copy()
    if d.empty:
        st.info("No YVF data available for selected filters.")
        return

    total_yvf = float(d["YVF Booking"].sum())
    total_iff = float(d["IFF Shipment"].sum())
    remaining_iff = max(total_iff - total_yvf, 0.0)
    ratio = safe_div(total_yvf, total_iff)

    fig = go.Figure(
        data=[
            go.Pie(
                labels=["YVF Bookings", "Non-YVF IFF Shipments"],
                values=[total_yvf, remaining_iff],
                hole=0.58,
                sort=False,
                direction="clockwise",
                marker=dict(
                    colors=[BUSINESS_COLORS["actual"], COLORS["grid"]],
                    line=dict(color="white", width=2),
                ),
                textinfo="label+percent",
                texttemplate="<b>%{label}</b><br>%{value:,.0f} · %{percent:.1%}",
                textposition="outside",
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "Volume: %{value:,.0f}<br>"
                    "Share: %{percent:.1%}"
                    "<extra></extra>"
                ),
            )
        ]
    )

    fig.update_layout(
        title="YVF Booking Share of Total IFF Shipments",
        annotations=[
            dict(
                text=(
                    f"<b>{ratio:.1%}</b>"
                    f"<br><span style='font-size:12px'>YVF Adoption</span>"
                    f"<br><span style='font-size:11px'>{total_yvf:,.0f} / {total_iff:,.0f}</span>"
                ),
                x=0.5,
                y=0.5,
                font=dict(
                    size=22,
                    color=COLORS["navy"],
                    family=UI["font_family"],
                ),
                showarrow=False,
                align="center",
            )
        ],
    )
    fig = plotly_layout(
        fig,
        390,
        show_legend=True,
        legend_position="top",
        margin_left=44,
        margin_right=44,
        margin_top=78,
        margin_bottom=30,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )



def render_yvf_table(df: pd.DataFrame):
    """
    Table sourced from sheet 'YVF'.
    Shows only rows with actual source data.
    """
    if df is None or df.empty:
        st.info("No YVF data available for selected filters.")
        return

    d = df.copy()
    d = d[
        (pd.to_numeric(d["YVF Booking"], errors="coerce").fillna(0) != 0)
        | (pd.to_numeric(d["IFF Shipment"], errors="coerce").fillna(0) != 0)
    ].copy()

    if d.empty:
        st.info("No YVF data available for selected filters.")
        return

    has_month = (
        "MonthDate" in d.columns
        and d["MonthDate"].notna().any()
    )

    if has_month:
        d = d.sort_values(["MonthDate", "Office"]).copy()
        d["Month"] = d["MonthDate"].dt.strftime("%b-%y")
        display = d[
            ["Office", "Month", "YVF Booking", "IFF Shipment", "YVF Booking Ratio"]
        ].copy()
    else:
        display = d[
            ["Office", "YVF Booking", "IFF Shipment", "YVF Booking Ratio"]
        ].copy()
        display = display.sort_values(["Office"])

    total_yvf = float(display["YVF Booking"].sum())
    total_iff = float(display["IFF Shipment"].sum())
    total_ratio = safe_div(total_yvf, total_iff)

    total_row = {
        "Office": "TOTAL",
        "YVF Booking": total_yvf,
        "IFF Shipment": total_iff,
        "YVF Booking Ratio": total_ratio,
    }
    if has_month:
        total_row["Month"] = ""

    display = pd.concat(
        [display, pd.DataFrame([total_row])],
        ignore_index=True,
    )

    column_cfg = {
        "Office": st.column_config.TextColumn("Office", width="small"),
        "YVF Booking": st.column_config.NumberColumn(
            "Total YVF Bookings", width="medium", format="%,.0f"
        ),
        "IFF Shipment": st.column_config.NumberColumn(
            "Total IFF Shipments", width="medium", format="%,.0f"
        ),
        "YVF Booking Ratio": st.column_config.NumberColumn(
            "YVF Booking Ratio", width="medium", format="percent"
        ),
    }
    if has_month:
        column_cfg["Month"] = st.column_config.TextColumn(
            "Month", width="small"
        )

    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:{UI['chart_title_size']}px;
            font-weight:700;
            margin:2px 0 10px 2px;">
            YVF Performance Detail
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=390,
        column_config=column_cfg,
    )



# ============================================================
# MAIN APP
# ============================================================


def main():
    # Load default workbook first. Upload remains below Month / Office filters.
    file_path = Path(DEFAULT_FILE)
    cache_token = ""

    # If an uploaded workbook was already saved in this session, reuse the saved file.
    uploaded_path_cached = st.session_state.get("dashboard_uploaded_path")
    uploaded_sig_cached = st.session_state.get("dashboard_uploaded_sig")

    if uploaded_path_cached and Path(uploaded_path_cached).exists():
        file_path = Path(uploaded_path_cached)
        cache_token = uploaded_sig_cached or ""
    elif file_path.exists():
        stat = file_path.stat()
        cache_token = f"{stat.st_mtime_ns}_{stat.st_size}"

    if not file_path.exists():
        st.error(
            f"Không tìm thấy file dữ liệu: {file_path}. "
            "Vui lòng đặt file Excel cùng thư mục app.py hoặc upload file ở Sidebar."
        )
        st.stop()

    with st.spinner("Loading and validating Excel data..."):
        raw = load_data(str(file_path), cache_token)
        hc = prepare_hc(raw["hc"])
        workload = prepare_workload(raw["workload"])
        fte = prepare_fte(raw["fte"])
        shipment, shipment_mode = prepare_shipment(raw["shipment"])
        customer = prepare_customer(raw)
        # Section 2 customer ranking/detail must use Customer Volume-N&S only.
        customer_ns = customer_wide_to_long(raw["customer_ns"])
        resolution = prepare_resolution(raw["resolution"])
        yvf = prepare_yvf(raw["yvf"])

        # Section 5 detail sources (C / A / S / E)
        core_detail = prepare_case_detail(raw["core"], "Core Service")
        ancillary_detail = prepare_case_detail(raw["ancillary"], "Ancillary Service")
        supporting_detail = prepare_case_detail(raw["supporting"], "Supporting Activity")
        exception_detail = prepare_case_detail(raw["exception"], "Exception Handling")

        # Code description lookup from sheet "Ghi chú".
        code_note_map = prepare_code_note_map(raw["notes"])

        # Add one consistent "Code Description" field to all C/A/S/E sources.
        core_detail = add_code_description(core_detail, "Core Service", code_note_map)
        ancillary_detail = add_code_description(ancillary_detail, "Ancillary Service", code_note_map)
        supporting_detail = add_code_description(supporting_detail, "Supporting Activity", code_note_map)
        exception_detail = add_code_description(exception_detail, "Exception Handling", code_note_map)

    periods = all_periods(hc, workload, fte, shipment, customer, customer_ns, resolution)
    month_options = ["All"] + [format_month(p) for p in periods]

    offices_from_data = sorted(set(
        list(hc.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(workload.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(fte.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(shipment.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(customer.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(customer_ns.get("Office", pd.Series(dtype=str)).dropna().unique())
    ))
    office_options = ["All Offices"] + sorted(set(STANDARD_OFFICES + [o for o in offices_from_data if o]))

    # Sidebar order: Month -> Office -> Upload file. No Year and no Reset button.
    with st.sidebar:
        st.markdown("## FILTERS")
        st.caption("Month / Office")
        month = st.selectbox("MONTH", month_options, key="month_filter")
        office = st.selectbox("OFFICE", office_options, key="office_filter")
        st.markdown("---")
        uploaded = st.file_uploader(
            "UPLOAD EXCEL FILE",
            type=["xlsx", "xlsm", "xls"],
            help="Nếu không upload, Dashboard sẽ đọc file mặc định trong cùng thư mục app.py.",
            key="excel_uploader",
        )
        if uploaded is not None:
            new_bytes = uploaded.getvalue()
            new_sig = hashlib.md5(new_bytes).hexdigest()

            if st.session_state.get("dashboard_uploaded_sig") != new_sig:
                tmp_path = Path("_uploaded_dashboard_data.xlsx")
                tmp_path.write_bytes(new_bytes)

                st.session_state["dashboard_uploaded_sig"] = new_sig
                st.session_state["dashboard_uploaded_path"] = str(tmp_path)
                st.rerun()
        st.caption(f"Source file: {Path(file_path).name}")
        st.caption("Capacity standard: 167.2 hrs/FTE/month")

    # Year is intentionally not exposed as a filter.
    year = "All"

    # Apply filters
    f_hc = apply_filters(hc, year, month, office)
    f_workload = apply_filters(workload, year, month, office)
    f_fte = apply_filters(fte, year, month, office)
    f_shipment = apply_filters(shipment, year, month, office)
    f_mode = apply_filters(shipment_mode, year, month, office)
    f_customer = apply_filters(customer, year, month, office)
    f_customer_ns = apply_filters(customer_ns, year, month, office)
    f_resolution = apply_filters(resolution, year, month, office)
    # YVF follows Month/Year filters only when the source sheet contains Month.
    # Current office-only source remains valid without inventing a time dimension.
    if (
        yvf is not None
        and not yvf.empty
        and "MonthDate" in yvf.columns
        and yvf["MonthDate"].notna().any()
    ):
        f_yvf = apply_filters(yvf, year, month, office)
    else:
        f_yvf = filter_office_only(yvf, office)

    f_core_detail = apply_filters(core_detail, year, month, office)
    f_ancillary_detail = apply_filters(ancillary_detail, year, month, office)
    f_supporting_detail = apply_filters(supporting_detail, year, month, office)
    f_exception_detail = apply_filters(exception_detail, year, month, office)

    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title">{APP_TITLE}</div>
            <div class="subtitle">{APP_SUBTITLE}</div>
            <div class="subtitle"><b>Selected Month:</b> {month} &nbsp; | &nbsp; <b>Selected Office:</b> {office}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if workload.empty or fte.empty:
        st.markdown(
            """
            <div class="warning-box">
            WARNING: Một số dữ liệu workload hoặc CS FTE có thể chưa đầy đủ. Dashboard vẫn chạy dynamic và sẽ tự cập nhật khi bổ sung dữ liệu vào file nguồn.
            </div>
            """,
            unsafe_allow_html=True,
        )

    section_title("1. Office Capacity Snapshot")

    # Section 1 uses the HC sheet as the single source of truth.
    approved_hc = weighted_period_avg(f_hc, "Total Approved HC") if not f_hc.empty else 0.0
    approved_mng = weighted_period_avg(f_hc, "Approved HC MNG") if not f_hc.empty else 0.0
    approved_pic = weighted_period_avg(f_hc, "Approved HC PIC") if not f_hc.empty else 0.0

    actual_hc = weighted_period_avg(f_hc, "Total Actual HC") if not f_hc.empty else 0.0
    actual_mng = weighted_period_avg(f_hc, "Actual HC MNG") if not f_hc.empty else 0.0
    actual_pic = weighted_period_avg(f_hc, "Actual HC PIC") if not f_hc.empty else 0.0

    required_hc = weighted_period_avg(f_hc, "Total Required HC") if not f_hc.empty else 0.0
    required_mng = weighted_period_avg(f_hc, "Required HC MNG") if not f_hc.empty else 0.0
    required_pic = weighted_period_avg(f_hc, "Required HC PIC") if not f_hc.empty else 0.0

    hc_variance = approved_hc - actual_hc
    if hc_variance > 0.05:
        variance_status = ("VACANCY GAP", COLORS["amber"], "#FEF3C7")
    elif hc_variance < -0.05:
        variance_status = ("ABOVE APPROVED", COLORS["red"], "#FEE2E2")
    else:
        variance_status = ("ON PLAN", COLORS["green"], "#DCFCE7")

    hc1, hc2, hc3, hc4 = st.columns(4, gap="medium")

    with hc1:
        hc_detail_card(
            "APPROVED HC",
            approved_hc,
            approved_mng,
            approved_pic,
        )

    with hc2:
        hc_detail_card(
            "ACTUAL HC",
            actual_hc,
            actual_mng,
            actual_pic,
        )

    with hc3:
        hc_detail_card(
            "REQUIRED HC",
            required_hc,
            required_mng,
            required_pic,
        )

    with hc4:
        hc_variance_card(
            "HC VARIANCE",
            hc_variance,
            "Approved HC − Actual HC",
            variance_status[0],
            variance_status[1],
            variance_status[2],
        )

    # KPI cards follow Month + Office filters.
    # The line chart keeps all available months so management can see the HC trend.
    hc_trend_data = filter_office_only(hc, office)
    st.markdown('<div class="chart-box" style="margin-top:12px;">', unsafe_allow_html=True)
    chart_office_capacity_trend(hc_trend_data)

    section_title("2. Shipment Volume")

    # Source for both KPIs: Shipment volume sheet
    shipment_total = (
        float(f_shipment["Total Shipment"].sum())
        if not f_shipment.empty and "Total Shipment" in f_shipment.columns
        else 0.0
    )
    active_customers = calculate_active_customers(f_shipment)

    # Two equal KPI cards similar to Section 1.
    sk1, sk2, sk3, sk4 = st.columns(4, gap="medium")
    with sk1:
        shipment_kpi_card(
            "TOTAL SHIPMENT VOLUME",
            fmt_int(shipment_total),
            "",
        )
    with sk2:
        shipment_kpi_card(
            "ACTIVE CUSTOMERS",
            fmt_int(active_customers),
            "",
        )
    # Keep 2 empty columns so the two KPI cards retain the same visual width as Section 1.
    with sk3:
        st.empty()
    with sk4:
        st.empty()

    # Executive layout:
    # Row 1: Donut 40% | Top 20 Customers 60%
    # Row 2: Customer Detail full width, scrollable
    chart_left, chart_right = st.columns([0.42, 0.58], gap="medium")

    with chart_left:
        st.markdown('<div class="chart-box" style="margin-top:12px;">', unsafe_allow_html=True)
        chart_shipment_modes(f_mode)

    with chart_right:
        st.markdown('<div class="chart-box" style="margin-top:12px;">', unsafe_allow_html=True)
        chart_top_customers(f_customer_ns)

    st.markdown('<div class="chart-box" style="margin-top:12px;">', unsafe_allow_html=True)
    customer_detail_table(f_customer_ns)

    section_title("3. Workload by PIC")

    # KPI source: HC
    # Total Available Standard Time:
    # prefer HC source column; if blank, calculate Total Actual HC × 167.2.
    hc_for_pic = f_hc.copy()

    if not hc_for_pic.empty:
        hc_for_pic["Calculated Available Hours"] = pd.to_numeric(
            hc_for_pic.get("HC Available Hours"), errors="coerce"
        )
        fallback_capacity = (
            pd.to_numeric(hc_for_pic.get("Total Actual HC"), errors="coerce")
            * CAPACITY_HOURS_PER_FTE
        )
        hc_for_pic["Calculated Available Hours"] = (
            hc_for_pic["Calculated Available Hours"].fillna(fallback_capacity)
        )

    total_available = filtered_monthly_metric(
        hc_for_pic, "Calculated Available Hours", "sum"
    ) if not hc_for_pic.empty else float("nan")

    standard_per_pic = CAPACITY_HOURS_PER_FTE

    total_actual_working = filtered_monthly_metric(
        hc_for_pic, "HC Actual Working Hours", "sum"
    ) if not hc_for_pic.empty else float("nan")

    actual_workload_per_pic = filtered_monthly_metric(
        hc_for_pic, "HC Actual Workload per PIC", "mean"
    ) if not hc_for_pic.empty else float("nan")

    # IMPORTANT: Do not calculate Capacity Utilization from CS FTE.
    # Source of truth is sheet HC -> "Capacity Utilization (%)".
    capacity_util = hc_capacity_utilization(hc_for_pic)

    # Row 1: 4 compact numeric cards to avoid an overly dense 6-card row.
    p1, p2, p3, p4 = st.columns(4, gap="medium")

    with p1:
        pic_kpi_card(
            "TOTAL AVAILABLE STANDARD TIME",
            fmt_num(total_available, 1, " hrs") if not pd.isna(total_available) else "N/A",
            "95% × 8 × 22 × Actual HC",
        )

    with p2:
        pic_kpi_card(
            "AVAILABLE STANDARD TIME / PIC",
            fmt_num(standard_per_pic, 1, " hrs"),
            "95% × 8 × 22",
        )

    with p3:
        pic_kpi_card(
            "TOTAL ACTUAL WORKING TIME",
            fmt_num(total_actual_working, 1, " hrs")
            if not pd.isna(total_actual_working) else "N/A",
            "HC source: C + A + S + E",
        )

    with p4:
        pic_kpi_card(
            "ACTUAL WORKLOAD / PIC",
            fmt_num(actual_workload_per_pic, 1, " hrs")
            if not pd.isna(actual_workload_per_pic) else "N/A",
            "HC source",
        )

    # Row 2: Utilization + Status as wider management indicators.
    ps1, ps2 = st.columns([1.35, 0.65], gap="medium")
    with ps1:
        pic_utilization_card(capacity_util)
    with ps2:
        overall_workload_status_card(capacity_util)

    # Chart source: CS FTE
    # PIC Workload = coefficient in sheet "CS FTE" × Available Standard Time / PIC.
    # Available Standard Time / PIC = 95% × 8 × 22 = 167.2 hours.
    # Therefore: PIC Workload = CS FTE coefficient × 167.2 hours.
    # When All Offices is selected, only overloaded PICs/offices are displayed.
    st.markdown('<div class="chart-box" style="margin-top:8px;">', unsafe_allow_html=True)
    chart_workload_by_pic(f_fte, office)

    st.markdown(
        """
        <div style="
            margin-top:6px;
            color:#667085;
            font-size:11px;
            line-height:1.45;
            font-family:Inter, 'Segoe UI', Arial, sans-serif;">
            <b>PIC Workload (hrs)</b> = CS FTE Factor × Available Standard Time / PIC<br>
            <b>Available Standard Time / PIC</b> = 8 hrs/day × 22 days/month × 95% efficiency = 167.2 hrs/month
        </div>
        """,
        unsafe_allow_html=True,
    )

    section_title("4. Workload by Segment")

    segment_summary = build_segment_workload(f_workload, f_mode)
    segment_total_hours = (
        float(segment_summary["Allocation Time (h)"].sum())
        if not segment_summary.empty
        else 0.0
    )

    # Compact Section 4 header: description and KPI aligned on one balanced row.
    seg_intro, seg_kpi = st.columns([0.76, 0.24], gap="medium")
    with seg_intro:
        st.markdown(
            """
            <div style="
                min-height:92px;
                display:flex;
                align-items:center;
                color:#667085;
                font-size:12px;
                line-height:1.55;
                padding:0 8px 0 2px;">
                Compare workload concentration across Segments using Shipment Volume,
                Allocation Time, Required FTE and Workload Share.
            </div>
            """,
            unsafe_allow_html=True,
        )
    with seg_kpi:
        kpi_card(
            "TOTAL WORKLOAD HOURS",
            fmt_num(segment_total_hours, 1, " hrs"),
            "Source: BU allocation",
        )

    # Balanced visual row.
    # IMPORTANT: do not wrap Streamlit widgets with raw HTML <div> tags.
    # Streamlit renders widgets outside the HTML element, which creates empty white boxes.
    seg_chart, seg_table = st.columns([0.44, 0.56], gap="medium")

    with seg_chart:
        chart_service_matrix(f_workload, f_mode)

    with seg_table:
        segment_workload_table(f_workload, f_mode)

    st.markdown(
        """
        <div style="
            margin-top:7px;
            padding:10px 14px;
            background:#FFFFFF;
            border:1px solid #D8E1EA;
            border-radius:12px;
            color:#667085;
            font-size:11px;
            line-height:1.45;">
            <b style="color:#003B70;">Note:</b>
            Bubbles are clustered for easy comparison; bubble size = Workload Share (%).
            Required FTE uses the HC/FTE allocation source when available and falls back to Workload ÷ 167.2 hrs/FTE.
            Shipment Volume is sourced from the Shipment volume sheet and mapped to the corresponding Service.
        </div>
        """,
        unsafe_allow_html=True,
    )

    section_title("5. Workload Breakdown by Service Type and Activity")

    st.markdown(
        """
        <div style="
            color:#667085;
            font-size:12px;
            line-height:1.5;
            margin:0 0 10px 2px;">
            Workload is broken down into Core Service (C), Ancillary Service (A),
            Supporting Activity (S) and Exception Handling (E).
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Summary table + C/A/S/E allocation chart
    wb_chart, wb_table = st.columns([0.43, 0.57], gap="medium")
    with wb_chart:
        chart_case_allocation(f_workload)
    with wb_table:
        render_workload_breakdown_table(f_workload)

    # Four source-detail tables
    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:{UI['chart_title_size']}px;
            font-weight:700;
            margin:14px 0 8px 2px;">
            C / A / S / E Details
        </div>
        """,
        unsafe_allow_html=True,
    )

    casetab_c, casetab_a, casetab_s, casetab_e = st.tabs([
        "C · Core Service",
        "A · Ancillary Service",
        "S · Supporting Activity",
        "E · Exception Handling",
    ])
    with casetab_c:
        render_activity_detail_table(f_core_detail, "Core Service")
    with casetab_a:
        render_activity_detail_table(f_ancillary_detail, "Ancillary Service")
    with casetab_s:
        render_activity_detail_table(f_supporting_detail, "Supporting Activity")
    with casetab_e:
        render_activity_detail_table(f_exception_detail, "Exception Handling")

    section_title("6. CS Solution")

    # Executive KPIs sourced from sheet "CS Resolutions Rate".
    if not f_resolution.empty:
        total_abn = float(f_resolution["Total Abnormality"].sum())
        resolved = float(f_resolution["Resolved"].sum())
        rate = safe_div(resolved, total_abn)

        cs1, cs2, cs3 = st.columns(3, gap="medium")
        with cs1:
            kpi_card(
                "TOTAL ABNORMALITIES",
                fmt_int(total_abn),
                "Source: CS Resolutions Rate",
            )
        with cs2:
            kpi_card(
                "RESOLVED BY CS",
                fmt_int(resolved),
                "Cases resolved by CS",
            )
        with cs3:
            kpi_card(
                "CS RESOLUTION RATE",
                fmt_pct(rate),
                f"{fmt_int(resolved)} / {fmt_int(total_abn)} cases",
            )

    cs_chart, cs_table = st.columns([0.55, 0.45], gap="medium")
    with cs_chart:
        chart_resolution(f_resolution)
    with cs_table:
        render_cs_solution_table(f_resolution)

    section_title("7. YVF Promoter Effectiveness")

    if not f_yvf.empty:
        yvf_booking = float(f_yvf["YVF Booking"].sum())
        iff = float(f_yvf["IFF Shipment"].sum())
        yvf_rate = safe_div(yvf_booking, iff)

        y1, y2, y3 = st.columns(3, gap="medium")
        with y1:
            kpi_card(
                "TOTAL YVF BOOKINGS",
                fmt_int(yvf_booking),
                "Source: YVF",
            )
        with y2:
            kpi_card(
                "TOTAL IFF SHIPMENTS",
                fmt_int(iff),
                "Source: YVF",
            )
        with y3:
            kpi_card(
                "YVF BOOKING RATIO",
                fmt_pct(yvf_rate),
                f"{fmt_int(yvf_booking)} / {fmt_int(iff)}",
            )

    yvf_chart_col, yvf_table_col = st.columns([0.52, 0.48], gap="medium")
    with yvf_chart_col:
        chart_yvf(f_yvf)
    with yvf_table_col:
        render_yvf_table(f_yvf)


if __name__ == "__main__":
    main()
