# ============================================================
# CS WORKLOAD & CAPACITY DASHBOARD
# Python + Streamlit + Pandas + Plotly
# Data source: (100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx
# ============================================================

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
    "bg": "#F5F7FA",
    "white": "#FFFFFF",
    "text": "#1F2937",
    "muted": "#64748B",
    "border": "#D9E2EC",
}

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

    st.markdown(
        f"""
        <div class="hc-kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="hc-kpi-total">{fmt_num(total_value, 2 if "REQUIRED" in label.upper() else 0)}</div>
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
            <div class="hc-kpi-total">{fmt_num(value, 0)}</div>
            <div class="hc-variance-formula">{formula_text}</div>
            <span class="status-badge hc-variance-status"
                  style="color:{status_color};background:{status_bg};">
                {status_text}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(text: str):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def plotly_layout(fig: go.Figure, height: int = 340) -> go.Figure:
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=COLORS["text"], family="Arial"),
        title=dict(font=dict(size=15, color=COLORS["navy"]), x=0.0),
        margin=dict(l=10, r=10, t=45, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(gridcolor="#EDF2F7", zeroline=False)
    fig.update_yaxes(gridcolor="#EDF2F7", zeroline=False)
    return fig

# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data(show_spinner=False)
def load_data(path: str) -> Dict[str, pd.DataFrame]:
    data = {}
    for key, sheet in SHEET_NAMES.items():
        data[key] = read_sheet(path, sheet, header=1)
    return data


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
        "HC Utilization (%)": "HC Utilization",
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
        "HC Available Hours", "HC Actual Working Hours", "HC Utilization"
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
    long["Actual FTE"] = numeric_series(long["Actual FTE"])
    long = long[(long["Office"] != "") & (~long["MonthDate"].isna())]
    return long[["Office", "CS PIC", "MonthDate", "Actual FTE"]]


def prepare_shipment(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate", "Total Shipment"]), pd.DataFrame(columns=["Office", "MonthDate", "Mode", "Volume"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    if not office_col or not month_col:
        return pd.DataFrame(), pd.DataFrame()
    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    total_col = first_existing(df, ["TOTAL", "Total"])
    if total_col:
        df["Total Shipment"] = numeric_series(df[total_col])
    else:
        df["Total Shipment"] = 0.0
    df["Active Customers"] = numeric_series(df.get("Active Customers", 0))
    mode_cols = [c for c in df.columns if c not in [office_col, month_col, "Office", "MonthDate", "Active Customers", total_col, "Total Shipment"]]
    mode_cols = [c for c in mode_cols if not str(c).startswith("Unnamed")]
    mode_long = df.melt(id_vars=["Office", "MonthDate"], value_vars=mode_cols, var_name="Mode", value_name="Volume")
    mode_long["Volume"] = numeric_series(mode_long["Volume"])
    mode_long = mode_long[mode_long["Volume"] > 0]
    return df.dropna(subset=["MonthDate"]), mode_long.dropna(subset=["MonthDate"])


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


def prepare_resolution(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate", "Total Abnormality", "Resolved", "Resolution Rate"])
    df = df.copy()
    office_col = first_existing(df, ["OFFICE", "Office"])
    month_col = first_existing(df, ["Month"])
    total_col = first_existing(df, ["Total abnormality/month", "Total abnormality"])
    resolved_col = first_existing(df, ["No of abnormality resolved by CS", "Resolved"])
    rate_col = first_existing(df, ["CS Resolution rate", "Resolution rate"])
    if not office_col or not month_col:
        return pd.DataFrame(columns=["Office", "MonthDate"])
    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    df["Total Abnormality"] = numeric_series(df[total_col]) if total_col else 0
    df["Resolved"] = numeric_series(df[resolved_col]) if resolved_col else 0
    if rate_col:
        df["Resolution Rate"] = numeric_series(df[rate_col])
    else:
        df["Resolution Rate"] = df.apply(lambda r: safe_div(r["Resolved"], r["Total Abnormality"]), axis=1)
    return df.dropna(subset=["MonthDate"])


def prepare_yvf(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "YVF Booking", "IFF Shipment", "YVF Booking Ratio"])
    df = df.copy()
    office_col = first_existing(df, ["OFFICE", "Office"])
    yvf_col = first_existing(df, ["Total YVF booking/month", "Total YVF booking"])
    iff_col = first_existing(df, ["Total IFF shipment/month", "Total IFF shipment"])
    ratio_col = first_existing(df, ["YVF booking ratio"])
    if not office_col:
        return pd.DataFrame(columns=["Office"])
    df["Office"] = df[office_col].map(normalize_office)
    df["YVF Booking"] = numeric_series(df[yvf_col]) if yvf_col else 0
    df["IFF Shipment"] = numeric_series(df[iff_col]) if iff_col else 0
    if ratio_col:
        df["YVF Booking Ratio"] = numeric_series(df[ratio_col])
    else:
        df["YVF Booking Ratio"] = df.apply(lambda r: safe_div(r["YVF Booking"], r["IFF Shipment"]), axis=1)
    return df


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
    """3-line HC trend from sheet HC with shaded gap between Approved HC and Actual HC."""
    if df.empty:
        st.info("No HC trend data available for selected filters.")
        return

    required_cols = ["MonthDate", "Total Approved HC", "Total Actual HC", "Total Required HC"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.info("HC trend cannot be displayed because required HC columns are missing.")
        return

    trend_source = df[
        ["MonthDate", "Total Approved HC", "Total Actual HC", "Total Required HC"]
    ].copy()

    for col in ["Total Approved HC", "Total Actual HC", "Total Required HC"]:
        trend_source[col] = pd.to_numeric(trend_source[col], errors="coerce")

    # Exclude months where all three HC values are blank.
    trend_source = trend_source.dropna(
        subset=["Total Approved HC", "Total Actual HC", "Total Required HC"],
        how="all",
    )

    if trend_source.empty:
        st.info("No HC trend data available for selected filters.")
        return

    trend = (
        trend_source.groupby("MonthDate", as_index=False)[
            ["Total Approved HC", "Total Actual HC", "Total Required HC"]
        ]
        .sum(min_count=1)
        .sort_values("MonthDate")
    )
    trend["Month"] = trend["MonthDate"].dt.strftime("%b-%y")

    fig = go.Figure()

    # Approved HC line
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Approved HC"],
            mode="lines+markers",
            name="Approved HC",
            line=dict(color=COLORS["navy"], width=3),
            marker=dict(size=7),
            hovertemplate="%{x}<br>Approved HC: %{y:,.1f}<extra></extra>",
        )
    )

    # Actual HC line — baseline for the shaded Actual vs Required gap.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Actual HC"],
            mode="lines+markers",
            name="Actual HC",
            line=dict(color=COLORS["blue"], width=3),
            marker=dict(size=7),
            hovertemplate="%{x}<br>Actual HC: %{y:,.1f}<extra></extra>",
        )
    )

    # Required HC line + shaded gap to Actual HC.
    # The fill is intentionally between Actual HC and Required HC.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Required HC"],
            mode="lines+markers",
            name="Required HC",
            line=dict(color=COLORS["red"], width=3, dash="dot"),
            marker=dict(size=7),
            fill="tonexty",
            fillcolor="rgba(230, 0, 18, 0.12)",
            hovertemplate="%{x}<br>Required HC: %{y:,.2f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="HC Trend & Gap",
        yaxis_title="HC",
        hovermode="x unified",
    )
    fig = plotly_layout(fig, 360)
    fig.update_xaxes(type="category")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


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
    fig = plotly_layout(fig, 380)
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
    fig = plotly_layout(fig, 280)
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
    fig = plotly_layout(fig, 340)
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
    fig = plotly_layout(fig, 340)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_service_matrix(df: pd.DataFrame):
    if df.empty:
        st.info("No office × service matrix data available.")
        return
    pivot = df.pivot_table(index="Office", columns="Segment", values="Workload Hours", aggfunc="sum", fill_value=0)
    for seg in SERVICE_ORDER:
        if seg not in pivot.columns:
            pivot[seg] = 0
    pivot = pivot[SERVICE_ORDER]
    fig = px.imshow(
        pivot,
        text_auto=".1f",
        aspect="auto",
        color_continuous_scale=[[0, COLORS["light_blue"]], [1, COLORS["blue"]]],
        title="Office × Service Workload Matrix (Hours)",
    )
    fig.update_layout(coloraxis_colorbar=dict(title="Hours"))
    fig = plotly_layout(fig, 330)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_shipment_modes(mode_df: pd.DataFrame):
    if mode_df.empty:
        st.info("No shipment mode data available.")
        return
    agg = mode_df.groupby("Mode", as_index=False)["Volume"].sum().sort_values("Volume", ascending=True)
    fig = px.bar(agg, x="Volume", y="Mode", orientation="h", text="Volume", color_discrete_sequence=[COLORS["blue"]], title="Shipment Volume by Transportation Mode")
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, hovertemplate="%{y}: %{x:,.0f}<extra></extra>")
    fig = plotly_layout(fig, 380)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_top_customers(df: pd.DataFrame):
    if df.empty:
        st.info("No customer volume data available.")
        return
    top = df.groupby("Customer", as_index=False)["Volume"].sum().sort_values("Volume", ascending=False).head(20)
    top = top.sort_values("Volume", ascending=True)
    fig = px.bar(top, x="Volume", y="Customer", orientation="h", text="Volume", color_discrete_sequence=[COLORS["blue"]], title="Top 20 Customers by Shipment Volume")
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, hovertemplate="%{y}: %{x:,.0f}<extra></extra>")
    fig = plotly_layout(fig, 520)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_resolution(df: pd.DataFrame):
    if df.empty:
        st.info("No CS resolution data available.")
        return
    agg = df.groupby("MonthDate", as_index=False).agg({"Total Abnormality": "sum", "Resolved": "sum"})
    agg["Resolution Rate"] = agg.apply(lambda r: safe_div(r["Resolved"], r["Total Abnormality"]), axis=1)
    agg["Month"] = agg["MonthDate"].dt.strftime("%b-%y")
    agg = agg.sort_values("MonthDate")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=agg["Month"], y=agg["Total Abnormality"], name="Total Abnormality", marker_color=COLORS["light_blue"], yaxis="y"))
    fig.add_trace(go.Scatter(x=agg["Month"], y=agg["Resolution Rate"], name="Resolution Rate", mode="lines+markers", line=dict(color=COLORS["green"], width=3), yaxis="y2"))
    fig.update_layout(
        title="Control Tower Effectiveness – CS Resolution Rate",
        yaxis=dict(title="Abnormalities"),
        yaxis2=dict(title="Resolution Rate", overlaying="y", side="right", tickformat=".0%", range=[0, max(1, agg["Resolution Rate"].max() * 1.1)]),
    )
    fig = plotly_layout(fig, 340)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_yvf(df: pd.DataFrame):
    if df.empty:
        st.info("No YVF data available.")
        return
    d = df.copy()
    d = d[d["Office"].isin(STANDARD_OFFICES)]
    fig = px.bar(d, x="Office", y="YVF Booking Ratio", text="YVF Booking Ratio", color_discrete_sequence=[COLORS["blue"]], title="YVF Promoter Effectiveness – Booking Ratio")
    fig.update_traces(texttemplate="%{text:.1%}", textposition="outside", cliponaxis=False, hovertemplate="%{x}: %{y:.1%}<extra></extra>")
    fig.update_yaxes(tickformat=".0%")
    fig = plotly_layout(fig, 340)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# ============================================================
# MAIN APP
# ============================================================


def main():
    # Load default workbook first. The upload control is intentionally placed
    # below the Month / Office filters in the sidebar.
    file_path = Path(DEFAULT_FILE)

    # If a file was uploaded in this session, use it.
    uploaded_cached = st.session_state.get("dashboard_uploaded_file")
    if uploaded_cached:
        tmp_path = Path("_uploaded_dashboard_data.xlsx")
        tmp_path.write_bytes(uploaded_cached)
        file_path = tmp_path

    if not Path(file_path).exists():
        st.error(f"Không tìm thấy file dữ liệu: {file_path}. Vui lòng đặt file Excel cùng thư mục app.py hoặc upload file ở Sidebar.")
        st.stop()

    with st.spinner("Loading and validating Excel data..."):
        raw = load_data(str(file_path))
        hc = prepare_hc(raw["hc"])
        workload = prepare_workload(raw["workload"])
        fte = prepare_fte(raw["fte"])
        shipment, shipment_mode = prepare_shipment(raw["shipment"])
        customer = prepare_customer(raw)
        resolution = prepare_resolution(raw["resolution"])
        yvf = prepare_yvf(raw["yvf"])

    periods = all_periods(hc, workload, fte, shipment, customer, resolution)
    month_options = ["All"] + [format_month(p) for p in periods]

    offices_from_data = sorted(set(
        list(hc.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(workload.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(fte.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(shipment.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(customer.get("Office", pd.Series(dtype=str)).dropna().unique())
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
            if st.session_state.get("dashboard_uploaded_file") != new_bytes:
                st.session_state["dashboard_uploaded_file"] = new_bytes
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
    f_resolution = apply_filters(resolution, year, month, office)
    f_yvf = filter_office_only(yvf, office)

    kpis = calculate_kpis(f_hc, f_workload, f_fte, f_shipment)
    status = status_from_util(kpis["Utilization"])

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
    st.markdown('<div class="chart-box" style="margin-top:14px;">', unsafe_allow_html=True)
    chart_office_capacity_trend(hc_trend_data)
    st.markdown('</div>', unsafe_allow_html=True)

    section_title("2. Workload & Capacity Trend")
    t1, t2 = st.columns(2)
    with t1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        chart_workload_trend(f_workload)
        st.markdown('</div>', unsafe_allow_html=True)
    with t2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        chart_capacity_trend(f_workload, f_fte)
        st.markdown('</div>', unsafe_allow_html=True)

    section_title("3. Workload by Service Type & Composition")
    s1, s2 = st.columns([1.1, 0.9])
    with s1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        chart_workload_by_service(f_workload)
        st.markdown('</div>', unsafe_allow_html=True)
    with s2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        chart_workload_composition(f_workload)
        st.markdown('</div>', unsafe_allow_html=True)

    section_title("4. Office × Service Workload Matrix")
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    chart_service_matrix(f_workload)
    st.markdown('</div>', unsafe_allow_html=True)

    section_title("5. Shipment & Customer Analysis")
    h1, h2 = st.columns([0.9, 1.1])
    with h1:
        kpi_card("Total Shipment", fmt_int(kpis["Total Shipment"]), "Source: Shipment volume")
        st.markdown('<div class="chart-box" style="margin-top:10px;">', unsafe_allow_html=True)
        chart_shipment_modes(f_mode)
        st.markdown('</div>', unsafe_allow_html=True)
    with h2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        chart_top_customers(f_customer)
        st.markdown('</div>', unsafe_allow_html=True)

    section_title("6. Effectiveness")
    e1, e2 = st.columns(2)
    with e1:
        if not f_resolution.empty:
            total_abn = f_resolution["Total Abnormality"].sum()
            resolved = f_resolution["Resolved"].sum()
            rate = safe_div(resolved, total_abn)
            kpi_card("CS Resolution Rate", fmt_pct(rate), f"Resolved {fmt_int(resolved)} / {fmt_int(total_abn)} cases")
        st.markdown('<div class="chart-box" style="margin-top:10px;">', unsafe_allow_html=True)
        chart_resolution(f_resolution)
        st.markdown('</div>', unsafe_allow_html=True)
    with e2:
        if not f_yvf.empty:
            yvf_booking = f_yvf["YVF Booking"].sum()
            iff = f_yvf["IFF Shipment"].sum()
            yvf_rate = safe_div(yvf_booking, iff)
            kpi_card("YVF Booking Ratio", fmt_pct(yvf_rate), f"YVF {fmt_int(yvf_booking)} / IFF {fmt_int(iff)}")
        st.markdown('<div class="chart-box" style="margin-top:10px;">', unsafe_allow_html=True)
        chart_yvf(f_yvf)
        st.markdown('</div>', unsafe_allow_html=True)

    section_title("7. Detail & Reconciliation")
    tab1, tab2, tab3, tab4 = st.tabs(["Reconciliation", "Workload Detail", "Customer Detail", "Data Audit"])
    with tab1:
        recon = build_reconciliation(f_hc, f_workload, f_fte, f_shipment)
        st.dataframe(recon, use_container_width=True, hide_index=True)
    with tab2:
        detail_cols = ["Office", "MonthDate", "Segment", "Service Label", "Core Hours", "Ancillary Hours", "Supporting Hours", "Exception Hours", "Workload Hours"]
        detail = f_workload[[c for c in detail_cols if c in f_workload.columns]].copy()
        if not detail.empty:
            detail["Month"] = detail["MonthDate"].dt.strftime("%b-%y")
            detail = detail.drop(columns=["MonthDate"])
        st.dataframe(detail, use_container_width=True, hide_index=True)
    with tab3:
        cust = f_customer.copy()
        if not cust.empty:
            cust["Month"] = cust["MonthDate"].dt.strftime("%b-%y")
            cust = cust.drop(columns=["MonthDate"])
        st.dataframe(cust, use_container_width=True, hide_index=True)
    with tab4:
        audit_rows = []
        for key, sheet in SHEET_NAMES.items():
            df = raw.get(key, pd.DataFrame())
            audit_rows.append({
                "Sheet": sheet,
                "Rows": len(df),
                "Columns": len(df.columns) if not df.empty else 0,
                "Status": "PASS" if not df.empty else "WARNING",
                "Note": "Loaded" if not df.empty else "Sheet missing or empty",
            })
        audit = pd.DataFrame(audit_rows)
        st.dataframe(audit, use_container_width=True, hide_index=True)
        st.caption("Excel formula warning: Nếu file chứa công thức nhưng chưa lưu cached results, hãy mở Excel → Calculate → Save trước khi chạy Dashboard.")


if __name__ == "__main__":
    main()
