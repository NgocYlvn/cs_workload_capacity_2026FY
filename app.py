# -*- coding: utf-8 -*-
"""
CS Capacity & Productivity Dashboard
=====================================
Executive dashboard cho Customer Service - theo dõi Workload, Capacity, FTE, Utilization.

Nguồn dữ liệu chính thức (Source of Truth) - KHÔNG được thay đổi nếu chưa có xác nhận:
    - HC              -> Approved/Actual HC theo Office x Month (dùng để reconciliation)
    - BU allocation   -> Workload (phút) theo Office x Month x Segment (Core/Ancillary/Supporting/Exception)
    - Shipment volume -> Số lượng shipment theo Office x Month x Mode
    - CS FTE          -> Actual FTE theo Office x CS PIC x Month

Công thức chuẩn (đã chốt với Ms Ngọc):
    Capacity Hours/PIC/Month = 8h x 95% x 22 ngày = 167.2 giờ
    Actual FTE (Office, Month)   = SUM(CS FTE theo từng CS PIC thuộc Office đó, tháng đó)
    Capacity Hours (Office,Month)= Actual FTE x 167.2
    Total Workload Hours         = SUM('BU allocation'.Total Workload (min)) / 60
    Required FTE                 = Total Workload Hours / 167.2
    Utilization %                = Total Workload Hours / Capacity Hours
    Gap FTE                      = Actual FTE - Required FTE  (>0: dư người, <0: thiếu người)

Nguyên tắc:
    - KHÔNG hard-code Office/Month/CS PIC - toàn bộ danh sách được đọc động từ dữ liệu,
      nên khi Ms Ngọc bổ sung HAN/HLC/HCM vào các sheet nguồn, Dashboard tự nhận diện,
      không cần sửa code.
    - KHÔNG tự tạo số liệu giả cho các Office/Month còn thiếu dữ liệu - chỉ hiển thị
      "Không đủ dữ liệu" (WARNING) thay vì suy đoán.
    - Cột J (Required HC - PIC) và cột O (Utilization) của sheet `HC` CHỈ dùng để
      đối chiếu (reconciliation), KHÔNG dùng làm KPI chính trên Dashboard.

Chạy:
    streamlit run app.py
"""

import io
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONSTANTS
# ============================================================
HOURS_PER_FTE_PER_MONTH = 8 * 0.95 * 22  # = 167.2 gio/FTE/thang

REQUIRED_SHEETS = ["HC", "BU allocation", "Shipment volume", "CS FTE"]

# Các sheet bổ sung theo yêu cầu mới của quản lý - KHÔNG bắt buộc, nếu thiếu thì
# phần Dashboard tương ứng sẽ ẩn/báo "chưa có dữ liệu" thay vì lỗi crash toàn app.
OPTIONAL_SHEETS = ["C", "A", "S", "E", "YVF", "CS Resolutions Rate"]

SEGMENT_LABELS = {
    "Core Workload (min)": "Core",
    "Ancillary Workload (min)": "Ancillary",
    "Supporting Workload (min)": "Supporting",
    "Exception Workload (min)": "Exception",
}

MONTH_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

# ---- Yusen-inspired corporate color palette ----
COLOR_NAVY = "#003B70"
COLOR_BLUE = "#005BAC"
COLOR_LIGHT_BLUE = "#EAF3F8"
COLOR_RED = "#E60012"
COLOR_GREEN = "#169B62"
COLOR_AMBER = "#F59E0B"
COLOR_BG = "#F5F7FA"
COLOR_TEXT = "#1F2937"
COLOR_TEXT_SECONDARY = "#64748B"
COLOR_BORDER = "#D9E2EC"

STATUS_COLOR = {
    "Overload": COLOR_RED,
    "High Load": COLOR_AMBER,
    "Balanced": COLOR_GREEN,
    "Less Load": COLOR_BLUE,
    "No Data": COLOR_TEXT_SECONDARY,
}

SERVICE_TYPE_LABELS = {
    "AE": "Air Export", "AI": "Air Import",
    "OE": "Ocean Export", "OI": "Ocean Import",
    "CC": "Customs Clearance", "TR": "Trucking", "WH": "Warehouse",
}

CHART_FONT = dict(family="Segoe UI, Arial, sans-serif", color=COLOR_TEXT, size=12)


def style_fig(fig, title=None, height=350):
    """Áp style corporate chung cho mọi biểu đồ Plotly."""
    fig.update_layout(
        height=height,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=CHART_FONT,
        title=dict(text=title, x=0, xanchor="left", font=dict(size=14, color=COLOR_NAVY)) if title else None,
        margin=dict(l=10, r=10, t=45 if title else 10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0),
    )
    fig.update_xaxes(showgrid=False, linecolor=COLOR_BORDER)
    fig.update_yaxes(showgrid=True, gridcolor=COLOR_BORDER, zeroline=False)
    return fig



# ============================================================
# PHẦN A - HELPER: chuẩn hóa cột Month (text "Jul-26" HOẶC datetime -> "Jul-26" + key sort)
# ============================================================
def normalize_month(value):
    """Chuẩn hóa 1 giá trị Month về text dạng 'Mon-YY'. Trả về None nếu không parse được."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%b-%y")
    s = str(value).strip()
    if s == "" or s.lower() == "none":
        return None
    # Thử parse các định dạng text phổ biến: "Jul-26", "Jul-2026", "2026-07", ...
    for fmt in ("%b-%y", "%b-%Y", "%B-%y", "%B-%Y", "%Y-%m", "%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%b-%y")
        except ValueError:
            continue
    # Nếu không match format nào, trả lại nguyên bản (đã strip) để không mất dữ liệu,
    # nhưng sẽ được đánh dấu là "không parse được" ở bước audit.
    return s


def month_sort_key(month_label: str):
    """Sinh key sắp xếp thời gian cho month dạng 'Jul-26' -> (year, month_index)."""
    try:
        dt = datetime.strptime(month_label, "%b-%y")
        return (dt.year, dt.month)
    except (ValueError, TypeError):
        return (9999, 99)  # đẩy các giá trị không parse được xuống cuối


# ============================================================
# PHẦN B - LOAD DATA (cache theo nội dung file upload)
# ============================================================
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict:
    """Đọc toàn bộ sheet cần thiết từ file Excel. Trả về dict {sheet_name: DataFrame}."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    available = set(xls.sheet_names)
    missing = [s for s in REQUIRED_SHEETS if s not in available]
    if missing:
        raise ValueError(
            "Thiếu sheet bắt buộc trong file Excel: " + ", ".join(missing)
            + f". Các sheet hiện có: {', '.join(xls.sheet_names)}"
        )

    data = {}
    for sheet in REQUIRED_SHEETS:
        df = pd.read_excel(xls, sheet_name=sheet, header=1)  # header thực tế nằm ở dòng 2
        df = df.dropna(how="all")
        data[sheet] = df

    # Sheet bổ sung: đọc nếu có, bỏ qua nếu không có (không chặn app)
    for sheet in OPTIONAL_SHEETS:
        if sheet in available:
            df = pd.read_excel(xls, sheet_name=sheet, header=1)
            data[sheet] = df.dropna(how="all")

    # Customer Volume theo từng Office (tên sheet dạng "Customer Volume - HAD"),
    # bỏ qua sheet tổng hợp "Customer Volume-N&S" để tránh double-count.
    cust_sheets = [
        s for s in xls.sheet_names
        if s.lower().startswith("customer volume") and "n&s" not in s.lower().replace(" ", "")
    ]
    cust_frames = []
    for s in cust_sheets:
        df = pd.read_excel(xls, sheet_name=s, header=1).dropna(how="all")
        cust_frames.append(df)
    if cust_frames:
        data["_customer_volume_raw"] = cust_frames

    return data


def clean_hc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Office" not in df.columns or "Month" not in df.columns:
        raise ValueError("Sheet HC thiếu cột 'Office' hoặc 'Month'.")
    df["Office"] = df["Office"].astype(str).str.strip()
    df["Month"] = df["Month"].apply(normalize_month)
    df = df.dropna(subset=["Office", "Month"])
    df = df[df["Office"] != ""]
    return df


def clean_bu_allocation(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    needed = ["Office", "Month", "Segment"]
    for col in needed:
        if col not in df.columns:
            raise ValueError(f"Sheet 'BU allocation' thiếu cột '{col}'.")
    df["Office"] = df["Office"].astype(str).str.strip()
    df["Month"] = df["Month"].apply(normalize_month)
    df = df.dropna(subset=["Office", "Month", "Segment"])

    seg_cols = [c for c in SEGMENT_LABELS if c in df.columns]
    for c in seg_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "Total Workload (min)" in df.columns:
        df["Total Workload (min)"] = pd.to_numeric(df["Total Workload (min)"], errors="coerce")
        # Nếu Total Workload (min) bị thiếu/lỗi, tự tính lại = tổng 4 segment (không suy đoán số liệu mới,
        # chỉ cộng lại đúng những gì đã có trong sheet nguồn)
        recompute_mask = df["Total Workload (min)"].isna()
        if seg_cols:
            df.loc[recompute_mask, "Total Workload (min)"] = df.loc[recompute_mask, seg_cols].sum(axis=1)
    else:
        df["Total Workload (min)"] = df[seg_cols].sum(axis=1) if seg_cols else 0

    return df


def clean_shipment(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Office" not in df.columns or "Month" not in df.columns:
        raise ValueError("Sheet 'Shipment volume' thiếu cột 'Office' hoặc 'Month'.")
    df["Office"] = df["Office"].astype(str).str.strip()
    df["Month"] = df["Month"].apply(normalize_month)
    df = df.dropna(subset=["Office", "Month"])
    if "TOTAL" in df.columns:
        df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce")
    return df


def clean_cs_fte(df: pd.DataFrame) -> pd.DataFrame:
    """CS FTE sheet có cấu trúc wide: Office, CS PIC, rồi 1 cột / tháng.
    Chuyển sang dạng long: Office, CS PIC, Month, Actual FTE."""
    df = df.copy()
    if "OFFICE" in df.columns:
        df = df.rename(columns={"OFFICE": "Office"})
    if "Office" not in df.columns or "CS PIC" not in df.columns:
        raise ValueError("Sheet 'CS FTE' thiếu cột 'OFFICE'/'Office' hoặc 'CS PIC'.")

    month_cols = [c for c in df.columns if c not in ("Office", "CS PIC")]
    long_df = df.melt(
        id_vars=["Office", "CS PIC"],
        value_vars=month_cols,
        var_name="Month",
        value_name="Actual FTE",
    )
    long_df["Office"] = long_df["Office"].astype(str).str.strip()
    long_df["CS PIC"] = long_df["CS PIC"].astype(str).str.strip()
    long_df["Month"] = long_df["Month"].apply(normalize_month)
    long_df = long_df.dropna(subset=["Office", "Month", "Actual FTE"])
    return long_df


@st.cache_data(show_spinner=False)
def build_clean_data(file_bytes: bytes) -> dict:
    raw = load_workbook(file_bytes)
    result = {
        "HC": clean_hc(raw["HC"]),
        "BU allocation": clean_bu_allocation(raw["BU allocation"]),
        "Shipment volume": clean_shipment(raw["Shipment volume"]),
        "CS FTE": clean_cs_fte(raw["CS FTE"]),
    }
    if "CS Resolutions Rate" in raw:
        result["CS Resolutions Rate"] = clean_cs_resolution(raw["CS Resolutions Rate"])
    if "YVF" in raw:
        result["YVF"] = clean_yvf(raw["YVF"])
    for code in ("C", "A", "S", "E"):
        if code in raw:
            result[code] = clean_detail_sheet(raw[code], code)
    if "_customer_volume_raw" in raw:
        result["Customer Volume"] = clean_customer_volume(raw["_customer_volume_raw"])
    return result


def clean_cs_resolution(df: pd.DataFrame) -> pd.DataFrame:
    """CS Resolutions Rate: Office, Month, Total abnormality/month, No resolved, CS Resolution rate."""
    df = df.copy()
    if "OFFICE" in df.columns:
        df = df.rename(columns={"OFFICE": "Office"})
    if "Office" not in df.columns or "Month" not in df.columns:
        return pd.DataFrame(columns=["Office", "Month", "CS Resolution rate"])
    df["Office"] = df["Office"].astype(str).str.strip()
    df["Month"] = df["Month"].apply(normalize_month)
    rate_col = next((c for c in df.columns if "Resolution rate" in str(c)), None)
    if rate_col:
        df[rate_col] = pd.to_numeric(df[rate_col], errors="coerce")
        df = df.rename(columns={rate_col: "CS Resolution rate"})
    df = df.dropna(subset=["Office", "Month"])
    return df


def clean_yvf(df: pd.DataFrame) -> pd.DataFrame:
    """YVF: Office, Total YVF booking/month, Total IFF shipment/month, YVF booking ratio.
    Sheet này KHÔNG có cột Month - là snapshot theo Office."""
    df = df.copy()
    if "OFFICE" in df.columns:
        df = df.rename(columns={"OFFICE": "Office"})
    if "Office" not in df.columns:
        return pd.DataFrame(columns=["Office", "YVF booking ratio"])
    df["Office"] = df["Office"].astype(str).str.strip()
    booking_col = next((c for c in df.columns if "YVF booking" in str(c) and "ratio" not in str(c).lower()), None)
    iff_col = next((c for c in df.columns if "IFF shipment" in str(c)), None)
    ratio_col = next((c for c in df.columns if "booking ratio" in str(c).lower()), None)
    if booking_col:
        df[booking_col] = pd.to_numeric(df[booking_col], errors="coerce")
    if iff_col:
        df[iff_col] = pd.to_numeric(df[iff_col], errors="coerce")
    if ratio_col:
        df[ratio_col] = pd.to_numeric(df[ratio_col], errors="coerce")
        # Nếu ratio bị trống nhưng có đủ booking & IFF thì tính lại = booking/IFF (không suy đoán số liệu mới,
        # chỉ tính lại đúng công thức đã có tên cột)
        if booking_col and iff_col:
            need_calc = df[ratio_col].isna() & df[iff_col].notna() & (df[iff_col] != 0)
            df.loc[need_calc, ratio_col] = df.loc[need_calc, booking_col] / df.loc[need_calc, iff_col]
        df = df.rename(columns={ratio_col: "YVF booking ratio"})
    df = df.dropna(subset=["Office"])
    df = df[df["Office"] != ""]
    return df


def clean_detail_sheet(df: pd.DataFrame, code: str) -> pd.DataFrame:
    """Chuẩn hóa sheet chi tiết C/A/S/E (volume theo Office x Scope/Job/Exception x Month)
    về dạng long: Office, Label, Month, Volume."""
    df = df.copy()
    if "Office" not in df.columns:
        return pd.DataFrame(columns=["Office", "Label", "Month", "Volume"])
    df["Office"] = df["Office"].astype(str).str.strip()

    label_col = next(
        (c for c in ("EXCEPTION DETAIL", "Job details", "Scope details", "Scope") if c in df.columns),
        None,
    )
    if label_col is None:
        return pd.DataFrame(columns=["Office", "Label", "Month", "Volume"])

    month_cols = [c for c in df.columns if isinstance(c, str) and normalize_month(c) not in (None,) and c not in ("Office",)]
    # Chỉ giữ cột thực sự parse được thành Month hợp lệ (loại Scope/Code/Criteria...)
    month_cols = [c for c in month_cols if month_sort_key(normalize_month(c)) != (9999, 99)]
    if not month_cols:
        return pd.DataFrame(columns=["Office", "Label", "Month", "Volume"])

    long_df = df.melt(
        id_vars=["Office", label_col],
        value_vars=month_cols,
        var_name="Month",
        value_name="Volume",
    )
    long_df = long_df.rename(columns={label_col: "Label"})
    long_df["Month"] = long_df["Month"].apply(normalize_month)
    long_df["Volume"] = pd.to_numeric(long_df["Volume"], errors="coerce")
    long_df = long_df.dropna(subset=["Office", "Month", "Volume"])
    long_df["Segment Code"] = code
    return long_df


def clean_customer_volume(frames: list) -> pd.DataFrame:
    """Gộp các sheet 'Customer Volume - <Office>' về dạng long: Office, Customer, Month, Volume."""
    all_long = []
    for df in frames:
        df = df.copy()
        if "Office" not in df.columns or "Customer" not in df.columns:
            continue
        df["Office"] = df["Office"].astype(str).str.strip()
        df["Customer"] = df["Customer"].astype(str).str.strip()
        month_cols = [c for c in df.columns if isinstance(c, str) and month_sort_key(normalize_month(c)) != (9999, 99)]
        if not month_cols:
            continue
        long_df = df.melt(
            id_vars=["Office", "Customer"],
            value_vars=month_cols,
            var_name="Month",
            value_name="Volume",
        )
        long_df["Month"] = long_df["Month"].apply(normalize_month)
        long_df["Volume"] = pd.to_numeric(long_df["Volume"], errors="coerce")
        long_df = long_df.dropna(subset=["Office", "Customer", "Month", "Volume"])
        all_long.append(long_df)
    if not all_long:
        return pd.DataFrame(columns=["Office", "Customer", "Month", "Volume"])
    return pd.concat(all_long, ignore_index=True)


# ============================================================
# PHẦN C - TÍNH KPI (theo công thức chuẩn đã chốt)
# ============================================================
def calc_workload_hours(bu_df: pd.DataFrame) -> pd.DataFrame:
    """Trả về DataFrame Office x Month: Total Workload Hours + breakdown theo Segment."""
    if bu_df.empty:
        return pd.DataFrame(columns=["Office", "Month", "Workload Hours"])
    g = bu_df.groupby(["Office", "Month"], as_index=False)["Total Workload (min)"].sum()
    g["Workload Hours"] = g["Total Workload (min)"] / 60
    return g[["Office", "Month", "Workload Hours"]]


def calc_actual_fte(cs_fte_df: pd.DataFrame) -> pd.DataFrame:
    """Actual FTE (Office, Month) = tổng Actual FTE của tất cả CS PIC thuộc Office đó."""
    if cs_fte_df.empty:
        return pd.DataFrame(columns=["Office", "Month", "Actual FTE"])
    return cs_fte_df.groupby(["Office", "Month"], as_index=False)["Actual FTE"].sum()


def calc_kpi_table(bu_df: pd.DataFrame, cs_fte_df: pd.DataFrame) -> pd.DataFrame:
    """Bảng KPI chính: Office x Month -> Workload Hours, Actual FTE, Capacity Hours,
    Required FTE, Utilization %, Gap FTE."""
    workload = calc_workload_hours(bu_df)
    fte = calc_actual_fte(cs_fte_df)

    kpi = pd.merge(workload, fte, on=["Office", "Month"], how="outer")
    kpi["Workload Hours"] = kpi["Workload Hours"].fillna(0)
    # Actual FTE có thể NaN nếu Office/Month đó chưa có data CS FTE -> giữ NaN, KHÔNG điền giả số 0
    # để tránh Utilization bị chia cho 0 một cách sai lệch (0 FTE thật khác với "chưa có data").

    kpi["Capacity Hours"] = kpi["Actual FTE"] * HOURS_PER_FTE_PER_MONTH
    kpi["Required FTE"] = kpi["Workload Hours"] / HOURS_PER_FTE_PER_MONTH

    def safe_utilization(row):
        cap = row["Capacity Hours"]
        if pd.isna(cap) or cap == 0:
            return None
        return row["Workload Hours"] / cap

    kpi["Utilization %"] = kpi.apply(safe_utilization, axis=1)
    kpi["Gap FTE"] = kpi["Actual FTE"] - kpi["Required FTE"]

    def status(u):
        if u is None or pd.isna(u):
            return "No Data"
        if u > 1.0:
            return "Overload"
        if u > 0.95:
            return "High Load"
        if u >= 0.90:
            return "Balanced"
        return "Less Load"

    kpi["Status"] = kpi["Utilization %"].apply(status)
    kpi["_sort"] = kpi["Month"].apply(month_sort_key)
    kpi = kpi.sort_values(["Office", "_sort"]).drop(columns="_sort")
    return kpi


def calc_segment_breakdown(bu_df: pd.DataFrame) -> pd.DataFrame:
    """Workload (giờ) theo Office x Month x Segment (Core/Ancillary/Supporting/Exception)."""
    seg_cols = [c for c in SEGMENT_LABELS if c in bu_df.columns]
    if bu_df.empty or not seg_cols:
        return pd.DataFrame(columns=["Office", "Month", "Segment", "Hours"])
    g = bu_df.groupby(["Office", "Month"], as_index=False)[seg_cols].sum()
    long_df = g.melt(id_vars=["Office", "Month"], value_vars=seg_cols,
                      var_name="Segment", value_name="Minutes")
    long_df["Segment"] = long_df["Segment"].map(SEGMENT_LABELS)
    long_df["Hours"] = long_df["Minutes"] / 60
    return long_df[["Office", "Month", "Segment", "Hours"]]


def calc_office_capacity_snapshot(hc_df: pd.DataFrame) -> pd.DataFrame:
    """Item 1: bảng Approved/Actual/Required HC + Variance theo MNG/PIC/Total, cho mỗi Office x Month."""
    if hc_df.empty:
        return pd.DataFrame(columns=["Office", "Month", "Level", "Approved", "Actual", "Required", "Variance"])

    col_map = {
        "MNG": {"Approved": "Approved HC – MNG", "Actual": "Actual HC – MNG", "Required": "Required HC – MNG"},
        "PIC": {"Approved": "Approved HC – PIC", "Actual": "Actual HC – PIC", "Required": "Required HC – PIC"},
        "Total": {"Approved": "Total Approved HC", "Actual": "Total Actual  HC", "Required": "Total Required HC"},
    }
    rows = []
    for _, r in hc_df.iterrows():
        for level, cols in col_map.items():
            approved = pd.to_numeric(r.get(cols["Approved"]), errors="coerce")
            actual = pd.to_numeric(r.get(cols["Actual"]), errors="coerce")
            required = pd.to_numeric(r.get(cols["Required"]), errors="coerce")
            variance = actual - required if pd.notna(actual) and pd.notna(required) else None
            rows.append({
                "Office": r["Office"], "Month": r["Month"], "Level": level,
                "Approved": approved, "Actual": actual, "Required": required, "Variance": variance,
            })
    return pd.DataFrame(rows)


def calc_segment_matrix(bu_df: pd.DataFrame, kpi_df: pd.DataFrame) -> pd.DataFrame:
    """Item 4: Office x Segment (AE/AI/OE/OI/CC/TR/WH) Workload Matrix -> Volume | Tỷ lệ | FTE %.
    'Segment' ở đây là cột Segment gốc trong BU allocation (loại hình dịch vụ), KHÁC với
    Core/Ancillary/Supporting/Exception (đó là 4 nhóm activity, xem calc_segment_breakdown)."""
    if bu_df.empty or "Segment" not in bu_df.columns:
        return pd.DataFrame(columns=["Office", "Month", "Segment", "Volume", "Workload Hours", "% of Office Workload", "FTE %"])

    vol_cols = [c for c in ("Core Volume", "Ancillary Volume", "Supporting Volume", "Exception Volume") if c in bu_df.columns]
    g = bu_df.groupby(["Office", "Month", "Segment"], as_index=False).agg(
        Volume=("Total Workload (min)", "size") if not vol_cols else (vol_cols[0], "sum"),
        Workload_min=("Total Workload (min)", "sum"),
    )
    if vol_cols:
        vol_sum = bu_df.groupby(["Office", "Month", "Segment"], as_index=False)[vol_cols].sum()
        g["Volume"] = vol_sum[vol_cols].sum(axis=1)

    g["Workload Hours"] = g["Workload_min"] / 60
    office_total = g.groupby(["Office", "Month"])["Workload Hours"].transform("sum")
    g["% of Office Workload"] = g["Workload Hours"] / office_total.replace(0, pd.NA)

    cap = kpi_df[["Office", "Month", "Capacity Hours"]].drop_duplicates()
    g = pd.merge(g, cap, on=["Office", "Month"], how="left")
    g["FTE %"] = g["Workload Hours"] / g["Capacity Hours"].replace(0, pd.NA)
    return g[["Office", "Month", "Segment", "Volume", "Workload Hours", "% of Office Workload", "FTE %"]]



def calc_reconciliation(hc_df: pd.DataFrame, kpi_df: pd.DataFrame) -> pd.DataFrame:
    """Đối chiếu KPI tính động (từ BU allocation / CS FTE) vs giá trị tham chiếu trong sheet HC.

    Chỉ 2 KPI dưới đây được kỳ vọng khớp nhau (cùng đơn vị/khái niệm) nên mới chấm PASS/WARNING/ERROR:
        - Required FTE  (tính động)      vs  cột J 'Required HC – PIC'  (hardcode trong HC)
        - Utilization %  (tính động)     vs  cột O 'HC Utilization (%)' (trong HC)

    2 dòng dưới đây chỉ hiển thị THAM KHẢO (status = INFO), KHÔNG chấm PASS/ERROR vì khác bản chất,
    không được kỳ vọng bằng nhau:
        - Actual FTE (CS FTE, đơn vị: workload ratio)  vs  Actual HC (HC, đơn vị: headcount)
        - Utilization % (Workload Hours / Capacity Hours)  vs  cột O của HC (= Required HC / Actual HC —
          công thức khác hẳn, không dựa trên Workload Hours)
    """
    if hc_df.empty or kpi_df.empty:
        return pd.DataFrame(columns=["Office", "Month", "KPI", "Calculated Value", "Reference Value", "Difference", "Status"])

    ref_cols = {}
    for col in hc_df.columns:
        if "Required HC" in str(col) and "PIC" in str(col):
            ref_cols["Required FTE"] = col
        if "Total Actual" in str(col) and "HC" in str(col):
            ref_cols["Actual HC (headcount)"] = col
        if "Utilization" in str(col):
            ref_cols["Utilization %"] = col

    # Chỉ Required FTE được chấm PASS/WARNING/ERROR (cùng khái niệm với cột J của HC, kỳ vọng khớp).
    # Utilization % và Actual FTE dùng công thức/đơn vị khác cột tham chiếu trong HC theo thiết kế,
    # nên chỉ hiển thị INFO kèm giải thích, không chấm là lỗi.
    scored_kpis = {"Required FTE"}

    rows = []
    merged = pd.merge(kpi_df, hc_df, on=["Office", "Month"], how="inner", suffixes=("", "_hc"))
    for _, r in merged.iterrows():
        comparisons = [
            ("Required FTE", "Required FTE"),
            ("Utilization %", "Utilization %"),
            ("Actual FTE (CS FTE) vs Actual HC (headcount)", "Actual FTE"),
        ]
        for kpi_name, calc_col in comparisons:
            ref_key = "Required FTE" if kpi_name == "Required FTE" else (
                "Utilization %" if kpi_name == "Utilization %" else "Actual HC (headcount)"
            )
            ref_col = ref_cols.get(ref_key)
            calc_val = r.get(calc_col)
            is_scored = kpi_name in scored_kpis

            if ref_col is None or ref_col not in merged.columns:
                status = "WARNING" if is_scored else "INFO"
                ref_val = None
                diff = None
            else:
                ref_val = r.get(ref_col)
                if pd.isna(calc_val) or pd.isna(ref_val):
                    status = "WARNING" if is_scored else "INFO"
                    diff = None
                else:
                    diff = calc_val - ref_val
                    if is_scored:
                        tol = 0.01 if kpi_name == "Utilization %" else 0.05
                        status = "PASS" if abs(diff) <= tol else "ERROR"
                    else:
                        status = "INFO"  # khác biệt là bình thường, không phải lỗi
            rows.append({
                "Office": r["Office"], "Month": r["Month"], "KPI": kpi_name,
                "Calculated Value": calc_val, "Reference Value": ref_val,
                "Difference": diff, "Status": status,
            })
    return pd.DataFrame(rows)


# ============================================================
# PHẦN D - UI HELPERS
# ============================================================
def inject_css():
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: {COLOR_BG}; }}
        div[data-testid="stMetric"] {{
            background-color: white;
            border: 1px solid {COLOR_BORDER};
            border-radius: 8px;
            padding: 12px 14px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }}
        div[data-testid="stMetricLabel"] {{ color: {COLOR_TEXT_SECONDARY}; }}
        div[data-testid="stMetricValue"] {{ color: {COLOR_NAVY}; }}
        h2, h3 {{ color: {COLOR_NAVY}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def status_badge(status: str) -> str:
    color = STATUS_COLOR.get(status, COLOR_TEXT_SECONDARY)
    return (
        f'<span style="background-color:{color}1A;color:{color};'
        f'padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;">{status}</span>'
    )


def extract_year(month_label: str):
    try:
        return "20" + month_label.split("-")[1]
    except (IndexError, AttributeError):
        return None


# ============================================================
# PHẦN E - MAIN
# ============================================================
def main():
    st.set_page_config(page_title="CS Workload & Capacity Dashboard", layout="wide", initial_sidebar_state="expanded")
    inject_css()

    with st.sidebar:
        st.header("📁 Dữ liệu nguồn")
        uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
        st.caption(
            "Khi bổ sung dữ liệu HAN/HLC/HCM vào các sheet nguồn, chỉ cần upload lại file — "
            "Dashboard tự động nhận diện Office/Month mới, không cần sửa code."
        )

    if uploaded is None:
        st.title("CS WORKLOAD & CAPACITY DASHBOARD")
        st.caption("Operations Performance | Capacity | Workload | Productivity")
        st.info("⬅️ Vui lòng upload file Excel nguồn để bắt đầu.")
        return

    file_bytes = uploaded.getvalue()

    try:
        data = build_clean_data(file_bytes)
    except ValueError as e:
        st.error(f"❌ Lỗi cấu trúc file: {e}")
        return
    except Exception as e:  # noqa: BLE001
        st.error(f"❌ Không đọc được file Excel: {e}")
        return

    hc_df = data["HC"]
    bu_df = data["BU allocation"]
    ship_df = data["Shipment volume"]
    fte_df = data["CS FTE"]

    if bu_df.empty:
        st.warning("⚠️ Sheet 'BU allocation' không có dữ liệu hợp lệ - không thể tính Workload/Utilization.")
        return

    kpi_df = calc_kpi_table(bu_df, fte_df)
    seg_df = calc_segment_breakdown(bu_df)
    reconciliation_df = calc_reconciliation(hc_df, kpi_df)
    snapshot_df = calc_office_capacity_snapshot(hc_df)
    matrix_df = calc_segment_matrix(bu_df, kpi_df)
    cs_res_df = data.get("CS Resolutions Rate", pd.DataFrame())
    yvf_df = data.get("YVF", pd.DataFrame())
    cust_df = data.get("Customer Volume", pd.DataFrame())

    kpi_df["Year"] = kpi_df["Month"].apply(extract_year)

    # ---- SIDEBAR: Year / Month / Office / Reset ----
    st.sidebar.header("🔎 FILTERS")

    if st.sidebar.button("↺ RESET FILTERS", use_container_width=True):
        for key in ("f_year", "f_month", "f_office"):
            st.session_state.pop(key, None)
        st.rerun()

    all_years = sorted(kpi_df["Year"].dropna().unique().tolist())
    sel_year = st.sidebar.selectbox("YEAR", ["All"] + all_years, key="f_year")

    year_filtered = kpi_df if sel_year == "All" else kpi_df[kpi_df["Year"] == sel_year]

    all_offices = sorted(year_filtered["Office"].dropna().unique().tolist())
    sel_offices = st.sidebar.multiselect("OFFICE", all_offices, default=all_offices, key="f_office")

    months_pool = year_filtered[year_filtered["Office"].isin(sel_offices)]["Month"].dropna().unique().tolist()
    months_pool = sorted(months_pool, key=month_sort_key)
    sel_months = st.sidebar.multiselect("MONTH", months_pool, default=months_pool, key="f_month")

    pic_pool = fte_df[fte_df["Office"].isin(sel_offices)]["CS PIC"].dropna().unique().tolist()
    sel_pics = st.sidebar.multiselect("CS PIC (tùy chọn)", sorted(pic_pool), default=[])

    if not sel_offices or not sel_months:
        st.warning("Vui lòng chọn ít nhất 1 Office và 1 Month để xem Dashboard.")
        return

    f_kpi = kpi_df[kpi_df["Office"].isin(sel_offices) & kpi_df["Month"].isin(sel_months)]
    f_seg = seg_df[seg_df["Office"].isin(sel_offices) & seg_df["Month"].isin(sel_months)]
    f_ship = ship_df[ship_df["Office"].isin(sel_offices) & ship_df["Month"].isin(sel_months)]
    f_fte_detail = fte_df[fte_df["Office"].isin(sel_offices) & fte_df["Month"].isin(sel_months)]
    f_snapshot = snapshot_df[snapshot_df["Office"].isin(sel_offices) & snapshot_df["Month"].isin(sel_months)] if not snapshot_df.empty else snapshot_df
    f_matrix = matrix_df[matrix_df["Office"].isin(sel_offices) & matrix_df["Month"].isin(sel_months)] if not matrix_df.empty else matrix_df
    f_cs_res = cs_res_df[cs_res_df["Office"].isin(sel_offices) & cs_res_df["Month"].isin(sel_months)] if not cs_res_df.empty else cs_res_df
    f_yvf = yvf_df[yvf_df["Office"].isin(sel_offices)] if not yvf_df.empty else yvf_df
    f_cust = cust_df[cust_df["Office"].isin(sel_offices) & cust_df["Month"].isin(sel_months)] if not cust_df.empty else cust_df
    if sel_pics:
        f_fte_detail = f_fte_detail[f_fte_detail["CS PIC"].isin(sel_pics)]

    # ---- HEADER ----
    hcol1, hcol2 = st.columns([3, 1])
    with hcol1:
        st.title("CS WORKLOAD & CAPACITY DASHBOARD")
        st.caption("Operations Performance | Capacity | Workload | Productivity")
    with hcol2:
        period_label = f"{sel_months[0]} → {sel_months[-1]}" if len(sel_months) > 1 else (sel_months[0] if sel_months else "N/A")
        st.markdown(
            f"<div style='text-align:right;color:{COLOR_TEXT_SECONDARY};font-size:13px;'>"
            f"<b>Period:</b> {period_label}<br/><b>Office:</b> {', '.join(sel_offices)}</div>",
            unsafe_allow_html=True,
        )

    if f_kpi.empty:
        st.warning("Không có dữ liệu cho lựa chọn hiện tại.")
        return

    # ==============================================================
    # 1. OFFICE CAPACITY SNAPSHOT (Executive Summary)
    # ==============================================================
    st.subheader("1️⃣ Office Capacity Snapshot")

    total_shipment = f_ship["TOTAL"].sum() if "TOTAL" in f_ship.columns else None
    total_workload_hours = f_kpi["Workload Hours"].sum()
    total_actual_fte = f_kpi["Actual FTE"].sum(skipna=True)
    total_required_fte = f_kpi["Required FTE"].sum(skipna=True)
    total_capacity_hours = f_kpi["Capacity Hours"].sum(skipna=True)
    overall_utilization = (total_workload_hours / total_capacity_hours) if total_capacity_hours else None
    gap_fte = (total_actual_fte - total_required_fte) if pd.notna(total_actual_fte) and pd.notna(total_required_fte) else None
    approved_hc = None
    actual_hc = None
    if not f_snapshot.empty:
        latest_month = max(sel_months, key=month_sort_key)
        snap_latest = f_snapshot[(f_snapshot["Level"] == "Total") & (f_snapshot["Month"] == latest_month)]
        approved_hc = snap_latest["Approved"].sum() if not snap_latest.empty else None
        actual_hc = snap_latest["Actual"].sum() if not snap_latest.empty else None

    r1 = st.columns(4)
    r1[0].metric("Approved HC", f"{approved_hc:,.0f}" if pd.notna(approved_hc) else "N/A")
    r1[1].metric("Actual HC", f"{actual_hc:,.0f}" if pd.notna(actual_hc) else "N/A")
    r1[2].metric("Actual FTE", f"{total_actual_fte:,.1f}" if pd.notna(total_actual_fte) else "N/A")
    r1[3].metric("Required FTE", f"{total_required_fte:,.1f}" if pd.notna(total_required_fte) else "N/A")

    r2 = st.columns(4)
    r2[0].metric("Capacity Hours", f"{total_capacity_hours:,.1f} h" if pd.notna(total_capacity_hours) else "N/A")
    r2[1].metric("Workload Hours", f"{total_workload_hours:,.1f} h")
    r2[2].metric(
        "Utilization %",
        f"{overall_utilization:.1%}" if overall_utilization is not None else "N/A",
    )
    r2[3].metric("FTE Gap", f"{gap_fte:+.1f}" if gap_fte is not None else "N/A",
                 help=">0: dư năng lực | ≈0: cân bằng | <0: thiếu nguồn lực")

    overall_status = "No Data"
    if overall_utilization is not None:
        if overall_utilization > 1.0:
            overall_status = "Overload"
        elif overall_utilization > 0.95:
            overall_status = "High Load"
        elif overall_utilization >= 0.90:
            overall_status = "Balanced"
        else:
            overall_status = "Less Load"
    st.markdown(f"Overall Workload Status: {status_badge(overall_status)}", unsafe_allow_html=True)

    missing_offices = [o for o in sel_offices if o not in f_fte_detail["Office"].unique()
                        and o not in f_kpi.dropna(subset=["Actual FTE"])["Office"].unique()]
    if missing_offices:
        st.info(f"ℹ️ Chưa có dữ liệu Workload/FTE đầy đủ cho: **{', '.join(missing_offices)}** — hiển thị N/A, không phải 0.")
    st.caption("Actual FTE (sheet `CS FTE`, workload ratio) KHÁC Actual HC (headcount, sheet `HC`) — không quy đổi qua lại. "
               "Approved/Actual HC hiển thị theo tháng gần nhất trong lựa chọn (snapshot), các KPI Workload/Capacity/FTE là tổng theo kỳ đã chọn.")

    st.divider()

    # ==============================================================
    # 2. WORKLOAD & CAPACITY TREND
    # ==============================================================
    st.subheader("2️⃣ Workload & Capacity Trend")
    trend = f_kpi.copy()
    trend["_sort"] = trend["Month"].apply(month_sort_key)
    trend = trend.sort_values("_sort")

    tcol1, tcol2 = st.columns(2)
    with tcol1:
        fig_t1 = go.Figure()
        for office, sub in trend.groupby("Office"):
            fig_t1.add_trace(go.Scatter(x=sub["Month"], y=sub["Workload Hours"], mode="lines+markers", name=office, line=dict(color=COLOR_BLUE)))
        style_fig(fig_t1, "Monthly Workload Trend (Hours)")
        st.plotly_chart(fig_t1, use_container_width=True)
    with tcol2:
        fig_t2 = go.Figure()
        for office, sub in trend.groupby("Office"):
            fig_t2.add_trace(go.Scatter(x=sub["Month"], y=sub["Utilization %"], mode="lines+markers", name=office))
        fig_t2.add_hline(y=1.0, line_dash="dash", line_color=COLOR_RED, annotation_text="100%")
        fig_t2.update_yaxes(tickformat=".0%")
        style_fig(fig_t2, "Utilization % Trend")
        st.plotly_chart(fig_t2, use_container_width=True)

    st.divider()

    # ==============================================================
    # 3. WORKLOAD BY SERVICE TYPE  &  4. WORKLOAD COMPOSITION (C/A/S/E)
    # ==============================================================
    scol1, scol2 = st.columns(2)
    with scol1:
        st.subheader("3️⃣ Workload by Service Type")
        if not f_matrix.empty:
            svc_total = f_matrix.groupby("Segment", as_index=False)["Workload Hours"].sum()
            svc_total["Label"] = svc_total["Segment"].map(SERVICE_TYPE_LABELS).fillna(svc_total["Segment"])
            svc_total["Pct"] = svc_total["Workload Hours"] / svc_total["Workload Hours"].sum()
            svc_total = svc_total.sort_values("Workload Hours", ascending=True)
            fig_svc = px.bar(
                svc_total, x="Workload Hours", y="Label", orientation="h",
                text=svc_total["Pct"].apply(lambda v: f"{v:.0%}"),
                color_discrete_sequence=[COLOR_BLUE],
            )
            style_fig(fig_svc, "Workload Breakdown by Service Type")
            st.plotly_chart(fig_svc, use_container_width=True)
        else:
            st.info("Không có dữ liệu.")

    with scol2:
        st.subheader("4️⃣ Workload Composition (C/A/S/E)")
        if not f_seg.empty:
            comp = f_seg.groupby(["Office", "Segment"], as_index=False)["Hours"].sum()
            fig_comp = px.bar(
                comp, x="Office", y="Hours", color="Segment", barmode="stack",
                color_discrete_sequence=[COLOR_NAVY, COLOR_BLUE, COLOR_AMBER, COLOR_RED],
            )
            style_fig(fig_comp, "Total Workload = Core + Ancillary + Supporting + Exception")
            st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.info("Không có dữ liệu.")

    st.divider()

    # ==============================================================
    # 5. OFFICE × SERVICE WORKLOAD MATRIX
    # ==============================================================
    st.subheader("5️⃣ Office × Service Workload Matrix")
    if f_matrix.empty:
        st.info("Không có dữ liệu để hiển thị Matrix.")
    else:
        heat = f_matrix.groupby(["Office", "Segment"], as_index=False)["Workload Hours"].sum()
        heat_pivot = heat.pivot(index="Office", columns="Segment", values="Workload Hours").fillna(0)
        fig_heat = go.Figure(data=go.Heatmap(
            z=heat_pivot.values, x=heat_pivot.columns, y=heat_pivot.index,
            colorscale=[[0, COLOR_LIGHT_BLUE], [1, COLOR_NAVY]],
            text=heat_pivot.values.round(0), texttemplate="%{text}",
        ))
        style_fig(fig_heat, "Workload Hours: Office × Service Type", height=320)
        st.plotly_chart(fig_heat, use_container_width=True)
        with st.expander("Chi tiết bảng Matrix (Volume / % / FTE%)", expanded=False):
            detail = f_matrix.groupby(["Office", "Segment"], as_index=False).agg(
                Volume=("Volume", "sum"), **{"Workload Hours": ("Workload Hours", "sum")},
            )
            tot = detail.groupby("Office")["Workload Hours"].transform("sum")
            detail["% of Office Workload"] = detail["Workload Hours"] / tot.replace(0, pd.NA)
            st.dataframe(
                detail.style.format({"Volume": "{:,.0f}", "Workload Hours": "{:,.1f}", "% of Office Workload": "{:.1%}"}),
                use_container_width=True, hide_index=True,
            )

    st.divider()

    # ==============================================================
    # 6. SHIPMENT VOLUME BY TRANSPORTATION MODE
    # ==============================================================
    st.subheader("6️⃣ Shipment Volume by Transportation Mode")
    kcol, ccol = st.columns([1, 2])
    with kcol:
        st.metric("Total Shipment", f"{total_shipment:,.0f}" if pd.notna(total_shipment) else "N/A")
        if "Active Customers" in f_ship.columns:
            latest_month_ship = max(sel_months, key=month_sort_key)
            active_latest = f_ship[f_ship["Month"] == latest_month_ship]["Active Customers"].sum()
            st.metric("Active Customers", f"{active_latest:,.0f}", help=f"Snapshot tháng {latest_month_ship}")
    with ccol:
        mode_cols = [c for c in f_ship.columns if c not in ("Office", "Month", "Active Customers", "TOTAL")]
        if mode_cols:
            mode_sum = f_ship[mode_cols].sum().sort_values(ascending=True)
            fig_mode = px.bar(
                x=mode_sum.values, y=mode_sum.index, orientation="h",
                labels={"x": "Volume", "y": "Mode"}, color_discrete_sequence=[COLOR_BLUE],
            )
            style_fig(fig_mode, "Volume theo Transportation Mode")
            st.plotly_chart(fig_mode, use_container_width=True)

    st.divider()

    # ==============================================================
    # 7. TOP 20 CUSTOMERS
    # ==============================================================
    st.subheader("7️⃣ Top 20 Customers by Shipment Volume")
    if not f_cust.empty:
        top20 = f_cust.groupby("Customer", as_index=False)["Volume"].sum().sort_values("Volume", ascending=False).head(20)
        top20 = top20.sort_values("Volume", ascending=True)
        fig_top20 = px.bar(top20, x="Volume", y="Customer", orientation="h", color_discrete_sequence=[COLOR_NAVY])
        style_fig(fig_top20, "Top 20 Customers", height=450)
        st.plotly_chart(fig_top20, use_container_width=True)
    else:
        st.info("Không có dữ liệu Customer Volume cho lựa chọn hiện tại.")

    st.divider()

    # ==============================================================
    # 8. CONTROL TOWER EFFECTIVENESS (CS Resolution Rate)
    # ==============================================================
    st.subheader("8️⃣ Control Tower Effectiveness — CS Resolution Rate")
    if f_cs_res.empty or "CS Resolution rate" not in f_cs_res.columns:
        st.info("Không có dữ liệu CS Resolution Rate cho lựa chọn hiện tại.")
    else:
        res_valid = f_cs_res.dropna(subset=["CS Resolution rate"])
        avg_rate = res_valid["CS Resolution rate"].mean() if not res_valid.empty else None
        rcol1, rcol2 = st.columns([1, 2])
        with rcol1:
            st.metric("Avg CS Resolution Rate", f"{avg_rate:.1%}" if avg_rate is not None else "N/A")
        with rcol2:
            res_trend = res_valid.copy()
            res_trend["_sort"] = res_trend["Month"].apply(month_sort_key)
            res_trend = res_trend.sort_values("_sort")
            fig_res = go.Figure()
            for office, sub in res_trend.groupby("Office"):
                fig_res.add_trace(go.Scatter(x=sub["Month"], y=sub["CS Resolution rate"], mode="lines+markers", name=office))
            fig_res.update_yaxes(tickformat=".0%")
            style_fig(fig_res, "CS Resolution Rate theo tháng")
            st.plotly_chart(fig_res, use_container_width=True)

    st.divider()

    # ==============================================================
    # 9. YVF PROMOTER EFFECTIVENESS
    # ==============================================================
    st.subheader("9️⃣ YVF Promoter Effectiveness")
    if f_yvf.empty or "YVF booking ratio" not in f_yvf.columns:
        st.info("Không có dữ liệu YVF cho lựa chọn hiện tại.")
    else:
        yvf_plot = f_yvf.dropna(subset=["YVF booking ratio"])
        if yvf_plot.empty:
            st.info("Chưa có Office nào có đủ dữ liệu YVF booking ratio.")
        else:
            fig_yvf = px.bar(yvf_plot, x="Office", y="YVF booking ratio", color_discrete_sequence=[COLOR_GREEN])
            fig_yvf.update_yaxes(tickformat=".0%")
            style_fig(fig_yvf, "YVF Booking Ratio theo Office")
            st.plotly_chart(fig_yvf, use_container_width=True)
        st.caption("Sheet `YVF` không có cột Month — đây là snapshot tổng, không lọc theo Month.")

    st.divider()

    # ==============================================================
    # EXCEPTIONS
    # ==============================================================
    st.subheader("⚠️ Exceptions — Office/Month cần chú ý")
    exceptions = f_kpi[f_kpi["Status"].isin(["Overload", "No Data"])]
    if exceptions.empty:
        st.success("Không phát hiện Office/Month nào ở trạng thái Overload hoặc thiếu dữ liệu.")
    else:
        st.dataframe(
            exceptions[["Office", "Month", "Workload Hours", "Actual FTE", "Required FTE", "Utilization %", "Status"]],
            use_container_width=True, hide_index=True,
        )

    st.divider()

    # ==============================================================
    # 10. DETAIL / RECONCILIATION
    # ==============================================================
    st.subheader("🔟 Detail / Reconciliation")
    with st.expander("🔍 Reconciliation - Đối chiếu với sheet HC", expanded=False):
        st.caption(
            "**Required FTE** được chấm PASS/WARNING/ERROR vì cùng khái niệm với cột J của `HC`, kỳ vọng khớp. "
            "**Utilization %** và **Actual FTE** chỉ mang tính THAM KHẢO (INFO): khác công thức/đơn vị với `HC`, "
            "KHÔNG được kỳ vọng bằng nhau."
        )
        if reconciliation_df.empty:
            st.info("Không đủ dữ liệu để đối chiếu.")
        else:
            f_recon = reconciliation_df[
                reconciliation_df["Office"].isin(sel_offices) & reconciliation_df["Month"].isin(sel_months)
            ]
            st.dataframe(f_recon, use_container_width=True, hide_index=True)
            n_error = (f_recon["Status"] == "ERROR").sum()
            n_warning = (f_recon["Status"] == "WARNING").sum()
            if n_error:
                st.error(f"{n_error} dòng lệch vượt ngưỡng (ERROR) trên Required FTE — cần kiểm tra lại nguồn dữ liệu.")
            if n_warning:
                st.warning(f"{n_warning} dòng chưa đủ giá trị tham chiếu để đối chiếu (WARNING).")

    with st.expander("👥 Headcount tham chiếu (từ sheet `HC`)", expanded=False):
        f_hc = hc_df[hc_df["Office"].isin(sel_offices) & hc_df["Month"].isin(sel_months)]
        hc_ref_cols = [c for c in f_hc.columns if str(c).strip() in ("Total Approved HC", "Total Actual  HC", "Total Actual HC")]
        if f_hc.empty or not hc_ref_cols:
            st.info("Không có dữ liệu headcount cho lựa chọn hiện tại.")
        else:
            st.dataframe(f_hc[["Office", "Month"] + hc_ref_cols], use_container_width=True, hide_index=True)

    with st.expander("📄 Detail Table — KPI theo Office x Month", expanded=False):
        st.dataframe(f_kpi, use_container_width=True, hide_index=True)

    with st.expander("👤 Actual FTE theo CS PIC", expanded=False):
        if not f_fte_detail.empty:
            fig6 = px.bar(f_fte_detail, x="CS PIC", y="Actual FTE", color="Office", color_discrete_sequence=[COLOR_BLUE, COLOR_NAVY, COLOR_AMBER, COLOR_GREEN])
            style_fig(fig6, "Actual FTE theo CS PIC")
            st.plotly_chart(fig6, use_container_width=True)
        else:
            st.info("Không có dữ liệu CS FTE cho lựa chọn hiện tại.")


if __name__ == "__main__":
    main()
