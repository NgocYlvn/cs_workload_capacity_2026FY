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

st.set_page_config(
    page_title="CS Capacity & Productivity Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)


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
    long_df["Month"] = long_df["Month"].apply(normalize_month)
    long_df["Actual FTE"] = pd.to_numeric(long_df["Actual FTE"], errors="coerce")
    long_df = long_df.dropna(subset=["Office", "Month", "Actual FTE"])
    return long_df


@st.cache_data(show_spinner=False)
def build_clean_data(file_bytes: bytes) -> dict:
    raw = load_workbook(file_bytes)
    return {
        "HC": clean_hc(raw["HC"]),
        "BU allocation": clean_bu_allocation(raw["BU allocation"]),
        "Shipment volume": clean_shipment(raw["Shipment volume"]),
        "CS FTE": clean_cs_fte(raw["CS FTE"]),
    }


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
# PHẦN D - UI
# ============================================================
def main():
    st.title("📊 CS Capacity & Productivity Dashboard")
    st.caption(
        "Nguồn: HC · BU allocation · Shipment volume · CS FTE  |  "
        "Utilization % = Workload Hours / Capacity Hours  (Capacity Hours = Actual FTE × 167.2h)"
    )

    with st.sidebar:
        st.header("📁 Dữ liệu nguồn")
        uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
        st.caption(
            "Khi bổ sung dữ liệu HAN/HLC/HCM vào các sheet nguồn, chỉ cần upload lại file — "
            "Dashboard tự động nhận diện Office/Month mới, không cần sửa code."
        )

    if uploaded is None:
        st.info("⬅️ Vui lòng upload file Excel nguồn (template CS Capacity & Productivity) để bắt đầu.")
        return

    file_bytes = uploaded.getvalue()

    # ---- Load + clean, với xử lý lỗi rõ ràng ----
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

    # ---- FILTERS ----
    st.sidebar.header("🔎 Bộ lọc")
    all_offices = sorted(kpi_df["Office"].dropna().unique().tolist())
    sel_offices = st.sidebar.multiselect("Office", all_offices, default=all_offices)

    months_for_offices = kpi_df[kpi_df["Office"].isin(sel_offices)]["Month"].dropna().unique().tolist()
    months_for_offices = sorted(months_for_offices, key=month_sort_key)
    sel_months = st.sidebar.multiselect("Month", months_for_offices, default=months_for_offices)

    pic_pool = fte_df[fte_df["Office"].isin(sel_offices)]["CS PIC"].dropna().unique().tolist()
    sel_pics = st.sidebar.multiselect("CS PIC (tùy chọn)", sorted(pic_pool), default=[])

    if not sel_offices or not sel_months:
        st.warning("Vui lòng chọn ít nhất 1 Office và 1 Month để xem Dashboard.")
        return

    f_kpi = kpi_df[kpi_df["Office"].isin(sel_offices) & kpi_df["Month"].isin(sel_months)]
    f_seg = seg_df[seg_df["Office"].isin(sel_offices) & seg_df["Month"].isin(sel_months)]
    f_ship = ship_df[ship_df["Office"].isin(sel_offices) & ship_df["Month"].isin(sel_months)]
    f_fte_detail = fte_df[fte_df["Office"].isin(sel_offices) & fte_df["Month"].isin(sel_months)]
    if sel_pics:
        f_fte_detail = f_fte_detail[f_fte_detail["CS PIC"].isin(sel_pics)]

    if f_kpi.empty:
        st.warning("Không có dữ liệu cho lựa chọn hiện tại.")
        return

    # ---- EXECUTIVE KPI ----
    st.subheader("Executive KPI")
    total_shipment = f_ship["TOTAL"].sum() if "TOTAL" in f_ship.columns else None
    total_workload_hours = f_kpi["Workload Hours"].sum()
    total_actual_fte = f_kpi["Actual FTE"].sum(skipna=True)
    total_required_fte = f_kpi["Required FTE"].sum(skipna=True)
    total_capacity_hours = f_kpi["Capacity Hours"].sum(skipna=True)
    overall_utilization = (
        total_workload_hours / total_capacity_hours if total_capacity_hours else None
    )
    gap_fte = total_actual_fte - total_required_fte if total_actual_fte else None

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Shipment", f"{total_shipment:,.0f}" if pd.notna(total_shipment) else "N/A")
    c2.metric("Total Workload (h)", f"{total_workload_hours:,.0f}")
    c3.metric("Actual FTE", f"{total_actual_fte:,.1f}" if pd.notna(total_actual_fte) else "N/A")
    c4.metric("Required FTE", f"{total_required_fte:,.1f}" if pd.notna(total_required_fte) else "N/A")
    c5.metric(
        "Utilization %",
        f"{overall_utilization:.0%}" if overall_utilization is not None else "N/A",
        delta=f"Gap {gap_fte:+.1f} FTE" if gap_fte is not None else None,
        delta_color="inverse",
    )

    missing_offices = [o for o in sel_offices if o not in f_fte_detail["Office"].unique()
                        and o not in f_kpi.dropna(subset=["Actual FTE"])["Office"].unique()]
    if missing_offices:
        st.info(
            f"ℹ️ Chưa có dữ liệu CS FTE / BU allocation đầy đủ cho: **{', '.join(missing_offices)}** "
            "- KPI của các Office này đang hiển thị 'No Data' hoặc N/A, không phải bằng 0."
        )

    st.caption(
        "ℹ️ **Actual FTE** ở trên là tổng workload-ratio theo CS PIC (sheet `CS FTE`) — "
        "KHÔNG phải số lượng nhân sự. Số lượng nhân sự thực tế (headcount) xem ở bảng "
        "\"Headcount tham chiếu\" bên dưới."
    )

    f_hc = hc_df[hc_df["Office"].isin(sel_offices) & hc_df["Month"].isin(sel_months)]
    hc_ref_cols = [c for c in f_hc.columns if str(c).strip() in (
        "Total Approved HC", "Total Actual  HC", "Total Actual HC",
    )]
    with st.expander("👥 Headcount tham chiếu (từ sheet `HC` — đơn vị: số người)", expanded=False):
        if f_hc.empty or not hc_ref_cols:
            st.info("Không có dữ liệu headcount cho lựa chọn hiện tại.")
        else:
            st.dataframe(
                f_hc[["Office", "Month"] + hc_ref_cols],
                use_container_width=True, hide_index=True,
            )

    st.divider()

    # ---- WORKLOAD & CAPACITY ----
    st.subheader("Workload & Capacity theo Office")
    colA, colB = st.columns(2)
    with colA:
        fig = px.bar(
            f_kpi, x="Office", y=["Workload Hours", "Capacity Hours"],
            barmode="group", title="Workload Hours vs Capacity Hours",
        )
        st.plotly_chart(fig, use_container_width=True)
    with colB:
        gap_by_office = f_kpi.groupby("Office", as_index=False)[["Actual FTE", "Required FTE"]].sum()
        gap_by_office["Gap FTE"] = gap_by_office["Actual FTE"] - gap_by_office["Required FTE"]
        fig2 = px.bar(
            gap_by_office, x="Office", y="Gap FTE",
            color=gap_by_office["Gap FTE"] > 0,
            color_discrete_map={True: "#2e7d32", False: "#c62828"},
            title="FTE Gap / Surplus theo Office (Actual − Required)",
        )
        fig2.update_layout(showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # ---- SERVICE PERFORMANCE (segment breakdown) ----
    st.subheader("Workload theo Service Segment (Core / Ancillary / Supporting / Exception)")
    if not f_seg.empty:
        fig3 = px.bar(
            f_seg, x="Office", y="Hours", color="Segment",
            barmode="stack", title="Tỷ trọng Workload theo Segment",
        )
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Không có dữ liệu breakdown theo segment cho lựa chọn hiện tại.")

    st.divider()

    # ---- TREND ----
    st.subheader("Xu hướng Workload & Utilization theo tháng")
    trend = f_kpi.copy()
    trend["_sort"] = trend["Month"].apply(month_sort_key)
    trend = trend.sort_values("_sort")
    fig4 = go.Figure()
    for office, sub in trend.groupby("Office"):
        fig4.add_trace(go.Scatter(x=sub["Month"], y=sub["Workload Hours"], mode="lines+markers", name=f"{office} - Workload (h)"))
    fig4.update_layout(title="Workload Hours theo tháng", xaxis_title="Month", yaxis_title="Hours")
    st.plotly_chart(fig4, use_container_width=True)

    fig5 = go.Figure()
    for office, sub in trend.groupby("Office"):
        fig5.add_trace(go.Scatter(x=sub["Month"], y=sub["Utilization %"], mode="lines+markers", name=office))
    fig5.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="100% (Overload threshold)")
    fig5.update_layout(title="Utilization % theo tháng", xaxis_title="Month", yaxis_tickformat=".0%")
    st.plotly_chart(fig5, use_container_width=True)

    st.divider()

    # ---- OFFICE / CS PIC BREAKDOWN ----
    st.subheader("Chi tiết theo CS PIC")
    if not f_fte_detail.empty:
        fig6 = px.bar(
            f_fte_detail, x="CS PIC", y="Actual FTE", color="Office",
            facet_col="Month" if len(sel_months) <= 4 else None,
            title="Actual FTE theo CS PIC",
        )
        st.plotly_chart(fig6, use_container_width=True)
    else:
        st.info("Không có dữ liệu CS FTE cho lựa chọn hiện tại.")

    st.divider()

    # ---- EXCEPTIONS ----
    st.subheader("⚠️ Exceptions - Office/Month cần chú ý")
    exceptions = f_kpi[f_kpi["Status"].isin(["Overload", "No Data"])]
    if exceptions.empty:
        st.success("Không phát hiện Office/Month nào ở trạng thái Overload hoặc thiếu dữ liệu.")
    else:
        st.dataframe(
            exceptions[["Office", "Month", "Workload Hours", "Actual FTE", "Required FTE", "Utilization %", "Status"]],
            use_container_width=True, hide_index=True,
        )

    st.divider()

    # ---- RECONCILIATION ----
    with st.expander("🔍 Reconciliation - Đối chiếu với sheet HC (cột J / O)", expanded=False):
        st.caption(
            "**Required FTE** được chấm PASS/WARNING/ERROR vì cùng khái niệm với cột J của `HC`, kỳ vọng khớp. "
            "**Utilization %** và **Actual FTE** chỉ mang tính THAM KHẢO (INFO): Utilization % dùng công thức "
            "Workload Hours/Capacity Hours (khác cột O của `HC` = Required/Actual HC); Actual FTE (CS FTE, "
            "workload ratio) khác đơn vị với Actual HC (headcount) — cả hai KHÔNG được kỳ vọng bằng nhau."
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
                st.error(f"{n_error} dòng lệch vượt ngưỡng (ERROR) trên Required FTE/Utilization — cần kiểm tra lại nguồn dữ liệu.")
            if n_warning:
                st.warning(f"{n_warning} dòng chưa đủ giá trị tham chiếu để đối chiếu (WARNING).")

    # ---- DETAIL TABLE ----
    with st.expander("📄 Detail Table - KPI theo Office x Month", expanded=False):
        st.dataframe(f_kpi, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
