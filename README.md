# CS Workload & Capacity Dashboard

Dashboard được thiết kế theo yêu cầu Executive/Corporate, sử dụng Python + Streamlit + Pandas + Plotly.

## File cần có
Đặt file Excel nguồn cùng thư mục với `app.py` và đổi đúng tên:

`(100826)TEMPLATE_DATA FOR DASHBOARD_V1.xlsx`

Hoặc có thể upload trực tiếp trên sidebar khi chạy app.

## Cách chạy

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Logic chính

- Actual FTE: lấy từ sheet `CS FTE`, SUM theo Office/Month.
- Required FTE: `Workload Hours / 167.2`.
- Capacity Hours: `Actual FTE × 167.2`.
- Workload Hours: `Total Workload (min) / 60` từ `BU allocation`.
- Utilization: `Workload Hours / Capacity Hours`.
- C/A/S/E Workload: lấy trực tiếp từ `BU allocation` làm Source of Truth.

## Lưu ý

- Code không hard-code riêng HAD.
- Khi bổ sung dữ liệu HAN/HLC/HCM vào file nguồn, dashboard tự động nhận nếu cấu trúc cột giữ nguyên.
- Nếu Excel có công thức, nên mở file bằng Excel → Calculate → Save trước khi chạy dashboard.
