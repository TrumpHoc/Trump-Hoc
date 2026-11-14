# STATUS – tiến độ & bàn giao giữa các phiên

Quy tắc: mỗi gói việc (WP) khi xong phải (1) lưu kết quả vào thư mục của mình, (2) cập nhật dòng tương ứng dưới đây, (3) commit + push lên nhánh `ccr-4eb5ff8a-s8s8p3`.

| WP | Nội dung | Phụ thuộc | Trạng thái | Kết quả |
|---|---|---|---|---|
| WP0 | Cơ sở thiết kế, nhập dữ liệu | – | XONG | `01_design_basis/DESIGN_BASIS.md`, `00_inputs/` |
| WP1 | Tính tải, cân bằng gió, chọn thiết bị (AHU, CDU, quạt, scrubber, máy hút ẩm, bơm nước ngưng) | WP0 | CHƯA | `02_calc/` |
| WP2 | Hình học: tọa độ phòng, tủ hood, vách, vị trí đặt thiết bị & điểm xả | WP0 | CHƯA | `03_geometry/` |
| WP3 | Tuyến & kích thước ống gió/ống gas/nước ngưng + bóc khối lượng | WP1, WP2 | CHƯA | `04_distribution/` |
| WP4 | Điện & điều khiển (DB.HVAC, DDC, cáp, máng, I/O) | WP1, WP2 | CHƯA | `05_electrical/` |
| WP5 | Tổng hợp BOQ thiết bị + vật tư lắp đặt | WP3, WP4 | CHƯA | `06_boq/` |
| WP6 | QA kiểm tra chéo toàn bộ, liệt kê lỗi → sửa → QA lại đến khi sạch | WP5 | CHƯA | `07_qa/` |

## Nhật ký
- 14/11/2026: WP0 xong. Chạy WP1 + WP2 song song.
