# HANDOFF PTGT – Buồng lạnh -20°C + Chiller 3x80HP

## CHECKPOINT 01 – 2026-11-13
- Dự án / Quota: Buồng lạnh -20°C (7x3,4x2,6m, 5 CDU 20HP + 5 dàn 25HP) + Chiller 3x80HP / TD-2026.11.13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v01.xlsx` (sheet DT, NOI_SUY, GIA_DINH)
- Script + tham số: `_scripts/ptgt_build.py`, `_scripts/params.json`; cache giá: `_cache/BANG_GIA_MEPF_v2.0_2026-11-12.csv`
- Đầu vào: `00. Tài Liệu/BANG_GIA_VAT_TU_MEPF_v2.0_2026-11-12.xlsx`, mẫu `00. Tài Liệu/MAU_BG_LMT1_But_Son_2026.08.13.xlsx`
- Đã xong: bóc khối lượng sơ bộ theo giả định, gán giá, nội suy (4 dòng), NC theo hệ số × K (SGC), recalc 0 lỗi, mọi Lô F=1
- Đang làm: chờ user trả lời câu hỏi mở
- Quyết định đã áp (CHƯA được user chốt): thiết bị chính do CĐT cấp (chỉ tính NC); TĐ cấp panel/cửa; uPVC Tiền Phong C2;
  nước lạnh 7/12°C → bảo ôn 25mm; P8 = 1,00; chi phí khác 8%; vật tư tiêu hao 2,5% VT; cỡ ống gas Ø41,28/Ø19,05; gas R404A 2 bình/hệ
- Đơn giá: quy đổi 13 dòng (ô cam); cần hỏi giá 13 dòng (ô đỏ: panel, cửa, phụ kiện panel, van cân bằng áp, dầu POE, phụ kiện đồng + que hàn,
  dây sưởi xả, 2 vỏ tủ điện, bộ điều khiển kho, PLC liên động, 2 lô phụ kiện điều khiển); NC đề xuất 11 dòng (ô vàng cột R)
- Số liệu (vốn, chưa VAT): VT 680,1 tr · NC + chi phí 414,0 tr · Tổng 1.094,1 tr · sau VAT 1.181,6 tr (CHƯA gồm các dòng cần hỏi giá)
  - I Buồng lạnh 458,0 tr · II Chiller 531,3 tr · III Vật tư phụ + chi phí khác 104,8 tr
- Kiểm tra §8: recalc 0 lỗi; Lô F=1 đạt; dòng con Lô không có (tủ điện tách dòng riêng); còn 13 ô Q trống (đã liệt kê)
- Rủi ro / câu hỏi mở: công suất 100HP cho buồng 62 m³ (cấp đông?); 5 dàn 25HP có lắp vừa buồng 7 m không; model CDU/chiller/bơm/tháp;
  nhiệt độ nước lạnh; mặt bằng/khoảng cách; P8; tổng khoán NC; giá LS MCCB đang là giá niêm yết (CK = 0)
- BƯỚC TIẾP THEO: user trả lời câu hỏi → sửa `params.json` (và giá panel/tủ nếu có) → chạy `python3 _scripts/ptgt_build.py v02`

## CHECKPOINT 02 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v02.xlsx` (v01 backup ở `90. Backup/`)
- Thay đổi: áp báo giá NCC panel (params `panel_NCC`): 77,88 m² tường+trần × 1.160.000; NC 80.000/m²; phụ kiện vỏ (gồm đèn, van thông áp) 5 tr; vận chuyển 1 tr.
  Bỏ dòng van cân bằng áp riêng (đã nằm trong lô phụ kiện). Sàn tách dòng "cần hỏi giá"; cửa vẫn "cần hỏi giá".
- Số liệu (vốn, chưa VAT): VT 777,9 tr · NC+CP 409,3 tr · Tổng 1.187,2 tr · sau VAT 1.282,1 tr; I 541,8 · II 531,3 · III 114,1
- Cần hỏi giá còn 11 dòng: cách nhiệt sàn, cửa, dầu POE, phụ kiện đồng + que hàn, dây sưởi xả, 2 vỏ tủ, bộ ĐK kho, PLC liên động, 2 lô phụ kiện ĐK
- BƯỚC TIẾP THEO: xác nhận sàn + cửa với NCC panel; các câu hỏi mở ở CHECKPOINT 01 → v03
