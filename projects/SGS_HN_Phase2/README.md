# SGS Hải Phòng – Cải tạo khu H&N Phase 2 (Nhà xưởng 3) – Thiết kế HVAC & BOQ

Thư mục làm việc cho thiết kế HVAC và BOQ thiết bị + vật tư lắp đặt. Đọc theo thứ tự:

1. `01_design_basis/DESIGN_BASIS.md` – cơ sở thiết kế (bắt buộc đọc trước).
2. `STATUS.md` – tiến độ từng gói việc, việc còn lại, cách tiếp tục.
3. `02_calc` – tải lạnh, lưu lượng, chọn thiết bị.
4. `03_geometry` – tọa độ phòng/tủ hood trích từ PDF kiến trúc.
5. `04_distribution` – tuyến & kích thước ống gió, ống gas, nước ngưng; bóc khối lượng.
6. `05_electrical` – DB.HVAC, tủ DDC, cáp, máng, điểm điều khiển.
7. `06_boq` – BOQ thiết bị và vật tư lắp đặt.
8. `07_qa` – báo cáo kiểm tra chéo.

## Công cụ (`tools/`)
- Đọc DWG: `npm i @mlightcad/libredwg-web@0.8.8` (thư mục bất kỳ), sửa đường dẫn import trong `dwgtext.mjs`/`dwg2json.mjs`, chạy `node dwgtext.mjs file.dwg > file.txt`.
- `render.py`: vẽ vùng DWG (JSON từ `dwg2json.mjs`) ra PNG (cần matplotlib).
- `build.py`, `scen.py`: bảng tính tải v3 (phiên bản trước khi chốt cơ sở thiết kế).
- `rfi_v5.py`: tạo file RFI.
- LibreOffice (`soffice`) có sẵn để tính lại công thức Excel.
