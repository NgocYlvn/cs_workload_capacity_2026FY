# CS Workload & Capacity Dashboard – 7 Section Version

## Data source
`(100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx`

## Dashboard order
1. Office Capacity Snapshot
2. Workload / Capacity by CS PIC
3. Shipment Volume & Active Customers
4. Office × Segment Workload Matrix
5. Workload Breakdown by Service Type & Activity
6. Control Tower Effectiveness
7. YVF Promoter Effectiveness

`Data Quality & Reconciliation` is placed in an expander after Section 07 and is not treated as Section 08.

## Filters
Sidebar filters: Month → Office. Không có Year và không có Reset Filters. Nút Upload Excel file đặt phía dưới 2 bộ lọc. Filters là dynamic từ dữ liệu nguồn. YVF hiện chưa có Month field, vì vậy Section 07 chỉ áp dụng Office cho đến khi dữ liệu YVF theo tháng được bổ sung.

## Main business logic
- 1 FTE capacity standard = `8 × 95% × 22 = 167.2 hours/month`.
- Workload Hours = `BU allocation → Total Workload (min) / 60`.
- Required FTE = `Workload Hours / 167.2`.
- Actual FTE is sourced from `CS FTE` as decimal FTE/workload ratio; it is not headcount.
- Capacity Hours = `Actual FTE × 167.2` per the agreed dashboard logic.
- Utilization = `Workload Hours / Capacity Hours`.
- BU allocation workload fields are treated as Source of Truth because the original Standard Time table is not in this workbook.
- Missing HAN/HLC/HCM workload/FTE data is treated as a data-availability warning. Logic remains dynamic and will pick up new source rows automatically.

## PIC-level note
The workbook does not provide BU allocation workload split directly by CS PIC. Section 02 therefore derives PIC workload-equivalent hours from `CS FTE × 167.2` and clearly labels this in the dashboard instead of inventing a BU allocation PIC split.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

Keep `app.py` and `(100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx` in the same folder, or upload the Excel file from the Streamlit sidebar.
