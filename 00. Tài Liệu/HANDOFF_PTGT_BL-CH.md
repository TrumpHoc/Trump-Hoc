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

## CHECKPOINT 03 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v03.xlsx` (v02 backup ở `90. Backup/`)
- Quyết định user: chỉ dùng ĐƠN GIÁ NCC panel (1.160.000 VT + 80.000 NC /m²), khối lượng theo bóc tách 6 mặt = 101,68 m² (gồm sàn) → bỏ dòng sàn riêng.
  Giữ lô phụ kiện vỏ 5 tr và vận chuyển 1 tr. Cửa vẫn chờ giá.
- Số liệu (vốn, chưa VAT): VT 806,2 tr · NC+CP 413,2 tr · Tổng 1.219,4 tr · sau VAT 1.316,9 tr; I 571,3 · II 531,3 · III 116,8
- Cần hỏi giá còn 10 dòng (cửa, dầu POE, phụ kiện đồng + que hàn, dây sưởi xả, 2 vỏ tủ, bộ ĐK kho, PLC, 2 lô phụ kiện ĐK)
- BƯỚC TIẾP THEO: giá cửa + các câu hỏi mở CHECKPOINT 01 → v04

## CHECKPOINT 04 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v04.xlsx` (v03 backup ở `90. Backup/`)
- Quyết định user: Phần I buồng lạnh dùng đơn giá ghi chép tay 13/11/2026 (ảnh lưu `00. Tài Liệu/GHI_CHEP_GIA_BUONG_LANH_2026-11-13_*.jpg`)
  × 1,10 = giá đầu vào (params `gia_ghi_chep`). TĐ cấp thiết bị buồng lạnh (máy nén Pana 20HP + bình ngưng tự lắp cụm, dàn lạnh, tháp, bơm).
  Đơn giá theo hệ ×5: cụm máy, dàn, TES5, gas, ống PVC, tủ ĐK, dây điện, NC lắp đặt. Ống đồng cặp Ø42/Ø16 + bảo ôn + phụ kiện 800k/m × 105 m.
  Panel 2 mặt inox 1.160.000 × 101,68 m² (6 mặt). Tháp/bơm SL theo phạm vi (1 tháp 20RT, 2 bơm 1,5kW) – chưa xác nhận công suất.
  Vật tư tiêu hao 2,5% chỉ tính trên VT không thuộc ghi chép.
- Số liệu (vốn, chưa VAT): VT 1.375,6 tr · NC+CP 397,7 tr · Tổng 1.773,3 tr · sau VAT 1.915,1 tr; I 1.094,6 · II 531,3 · III 147,4
- Sai lệch trong ghi chép cần hỏi: gas 5,7tr nhưng tổng 178,1tr khớp 4,5tr (lệch 1,2tr/hệ); inox 1 mặt 940k×77,88=73,2tr ≠ 74,76tr ghi (=960k)
- Cần hỏi giá còn 4 dòng: cửa, vỏ tủ chiller, PLC liên động chiller, phụ kiện ĐK tủ chiller
- BƯỚC TIẾP THEO: xác nhận tháp/bơm buồng lạnh (1 tháp 20RT+2 bơm 1,5kW hay 5 bộ), gas/hệ, chi phí khác 8% trên phần thiết bị → v05

## CHECKPOINT 05 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v05.xlsx` (v04 backup ở `90. Backup/`)
- Quyết định user: buồng lạnh 5 tháp 20RT + 5 bơm 1,5kW (1 bộ/cụm, không bơm dự phòng)
- Số liệu (vốn, chưa VAT): VT 1.460,8 tr · NC+CP 404,7 tr · Tổng 1.865,5 tr · sau VAT 2.014,8 tr; I 1.179,8 · II 531,3 · III 154,4
- Còn mở: gas 5,7 hay 4,5 tr/hệ; chi phí khác 8% có tính trên thiết bị không; 4 dòng chờ giá (cửa, vỏ tủ chiller, PLC, phụ kiện ĐK); P8
- BƯỚC TIẾP THEO: trả lời các điểm mở → v06

## CHECKPOINT 06 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v06.xlsx` (v05 backup ở `90. Backup/`)
- Quyết định user 13/11: gas 5,7tr/hệ; P8 = 1,00; bơm chiller (lạnh + GN) 13,5tr/cái; tháp chiller 42tr/cái (dùng nguyên giá, không ×1,1);
  ống chiller toàn bộ uPVC D90, 3 vòng độc lập, tank có sẵn 6 bích chờ DN80; van dùng van nhựa PVC (không van bướm); bảo ôn 15mm;
  tủ điện chiller trọn bộ giá nhập 45tr; thang cáp có nắp dày 1,2mm.
- Excel: tỷ lệ chi phí khác ô P9 (8%), tỷ lệ VT tiêu hao ô P10 (2,5%) – dòng III.A/III.B là công thức; cột X = 1 dòng tính tiêu hao.
- Nội suy mới: cút D90, van bi PVC D90 (ngoại suy, tin cậy thấp), bảo ôn Ø89 15mm, bu lông DN80. Y lọc giữ gang DN80 (Wonil).
- Số liệu Excel (vốn, chưa VAT): Tổng 1.953,7 tr · sau VAT 2.110,0 tr; I 1.179,8 · II 616,1 · III 157,8
- Cần hỏi giá: cửa buồng lạnh; van một chiều PVC D90 (6 cái); xác nhận giá van bi PVC D90
- Rủi ro: D90 với lưu lượng giả định 45 m³/h vòng GN → v ≈ 2,28 m/s (> 2,1 hút); cần lưu lượng thực của chiller
- BƯỚC TIẾP THEO: giá cửa + van 1 chiều PVC; lưu lượng chiller → v07

## CHECKPOINT 07 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v07.xlsx` (v06 backup ở `90. Backup/`)
- Quyết định user: lưu lượng lạnh 59 m³/h, GN 70 m³/h/chiller; ống uPVC D110; chỉ van bi nhựa tay gạt PVC D110 + Y lọc DN100;
  bỏ khớp mềm, van 1 chiều; chiller giá nhập 400tr/cái (TĐ cấp). Van cấp bù/xả đáy/đồng hồ đổi sang van bi PVC.
- Số liệu Excel (vốn, chưa VAT): Tổng 3.242,7 tr · sau VAT 3.502,2 tr; I 1.179,8 · II 1.810,0 · III 253,0 (chi phí khác 240,2)
- Rủi ro: v D110 lạnh 2,02 / GN 2,40 m/s (> 2,1 hút); chi phí khác 8% đang tính trên cả thiết bị (chiller 1,2 tỷ) – chờ user quyết
- Cần hỏi giá: cửa buồng lạnh; xác nhận van bi PVC D110 (ngoại suy 173.600)
- BƯỚC TIẾP THEO: user quyết nền tính 8%, xác nhận giữ đồng hồ/nhiệt kế/flow switch → v08

## CHECKPOINT 08 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v08.xlsx` (v07 backup ở `90. Backup/`)
- Quyết định user: (1) bỏ tính chi phí khác trên VT thiết bị chiller (chiller, bơm, tháp) – cột Y = 1, công thức III.B trừ ra;
  (2) giữ toàn bộ D110; (3) bỏ đồng hồ, nhiệt kế, công tắc dòng chảy, van cấp bù/xả đáy; chỉ giữ van D110 + Y lọc DN100;
  (4) van bi PVC D110 650k/cái: 1/bơm, 1/chiller, 1/tháp = 12 cái; (5) cửa 11tr/cái (dùng nguyên giá, không ×1,1).
- Số liệu Excel (vốn, chưa VAT): Tổng 3.105,3 tr · sau VAT 3.353,7 tr; I 1.201,8 · II 1.765,8 · III 137,7 (chi phí khác 125,8)
- Cần hỏi giá: 0. NC đề xuất (ô vàng) còn: đấu nối tank, súc rửa chạy thử, NC lắp chiller/bơm/tháp/tủ trong cột R
- BƯỚC TIẾP THEO: user duyệt NC đề xuất + điền thông tin khách hàng → bản chào giá

## CHECKPOINT 09 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v09.xlsx` (v08 backup ở `90. Backup/`)
- Quyết định user – NC chiller: lắp chiller 2tr/máy; bơm 300k/cái; tháp 1tr/bộ; tủ 1,5tr; bỏ dòng đấu nối tank và súc rửa/chạy thử.
- Không còn dòng NC đề xuất, không còn dòng cần hỏi giá.
- Số liệu Excel (vốn, chưa VAT): Tổng 3.061,7 tr · sau VAT 3.306,7 tr; I 1.201,8 · II 1.725,5 · III 134,5 (chi phí khác 122,6)
- Còn mở: thông tin khách hàng; P8 = 1,00 (bán = vốn); ẩn cột P–Y trước khi gửi
- BƯỚC TIẾP THEO: điền KH + chốt P8 → xuất bản chào giá (ẩn cột vốn)

## CHECKPOINT 10 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v10.xlsx` (v09 backup ở `90. Backup/`)
- Quyết định user: bích chờ tank DN100 (bỏ côn 110/90, bích D90; +6 bích D110, bu lông DN100 42 bộ); chi phí khác 8% → 4% (P9), bỏ phí văn phòng kho bãi;
  bỏ tủ điện + cáp + thang chiller → dòng "Tủ điện, dây điện cấp nguồn – CĐT cấp"; giá đỡ ống chiller Lô 8tr VT + 3tr NC;
  van PVC D110 gộp 1 dòng 12 cái; bảo ôn bỏ chữ Superlon, xuất xứ China; Y lọc DN100 China 700k/cái.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.766,6 tr · sau VAT 2.988,0 tr; I 1.201,8 · II 1.522,2 · III 42,7 (chi phí khác 34,1)
- Còn mở: thông tin khách hàng; P8; vật tư phụ điện 6tr+1,35tr (III.A) có giữ không khi điện chiller CĐT cấp
- BƯỚC TIẾP THEO: điền KH + chốt P8 → bản chào giá ẩn cột vốn

## CHECKPOINT 11 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v11.xlsx` (v10 backup ở `90. Backup/`)
- Quyết định user: mỗi phần lớn chỉ 3 mục: 1 Thiết bị / 2 Đường ống / 3 Đường điện; vật tư tiêu hao chuyển vào II.2 (cuối mục);
  bỏ vật tư phụ điện; Phần III chỉ còn "CHI PHÍ KHÁC"; tủ điện buồng lạnh 1 tủ/cụm (5 tủ × 6,6tr).
- Công thức: tiêu hao = P10 × SUMPRODUCT(P,Q,X) các dòng phía trên; CPK = P9 × (U_I + U_II − SUMPRODUCT(S,Y)) × w/140.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.759,1 tr · sau VAT 2.979,8 tr; I 1.201,8 · II 1.523,4 · III 33,9
- Còn mở: thông tin khách hàng; P8
- BƯỚC TIẾP THEO: điền KH + chốt P8 → bản chào giá ẩn cột vốn

## CHECKPOINT 12 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v12.xlsx` (v11 backup ở `90. Backup/`)
- Quyết định user: gộp máy nén + bình ngưng + POE + phụ kiện lên máy → "Cụm máy nén bình ngưng 20HP" (90,2tr VT, NC lắp ráp 4,4tr gộp vào);
  NC lắp đặt 11tr/hệ (55tr) chia đều 1/3 cho I.1, I.2, I.3, trong mục chia đều theo dòng (I.1: cụm, dàn, tháp, bơm; I.2: ống đồng, TES5, gas, ống PVC; I.3: tủ, dây);
  bảo ôn ống chiller tròn bội 4 (56 m); NC bảo ôn phụ kiện 700k; bỏ freeze; xoá cột Ghi chú/Nguồn/Trạng thái.
- BỐ CỤC CỘT MỚI: A–K bán; N nhãn; **O8 hệ số giá bán, O9 tỷ lệ CPK, O10 tỷ lệ tiêu hao**; O SL vốn, P ĐG VT vốn, Q ĐG NC vốn, R/S/T thành tiền vốn; U cờ tiêu hao; V cờ loại khỏi nền CPK.
  Nguồn/trạng thái/ghi chú từng dòng chuyển sang sheet NGUON.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.760,1 tr · sau VAT 2.980,9 tr; I 1.201,8 · II 1.524,3 · III 34,0
- BƯỚC TIẾP THEO: thông tin KH + hệ số giá bán (O8) → bản chào giá ẩn cột N–V

## CHECKPOINT 13 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v13.xlsx` (v12 backup ở `90. Backup/`)
- Quyết định user: bỏ dòng diễn giải vận tốc chiller; bỏ mọi chữ "ghi chép" (params đổi tên `gia_user_buong_lanh`, nguồn "Giá user cấp × 1,10");
  van D110 Taijan; ống đồng Hailiang; gas China; chiller "(đầu nén Hanbell - Đài Loan)", cột xuất xứ = Hstar; bơm chiller Teco; tháp chiller TTP;
  vận chuyển panel 1,1tr chuyển sang III (cộng vào dòng Phí vận chuyển); cửa VT 7,5tr + NC 1tr; NC lắp đặt 55tr phân bổ theo tỷ lệ VT (không chia đều);
  VT tiêu hao chiller cộng thêm 1tr (params `tieu_hao_cong_them`).
- Số liệu Excel (vốn, chưa VAT): Tổng 2.755,9 tr · sau VAT 2.976,4 tr; I 1.195,7 · II 1.525,3 · III 34,9
- BƯỚC TIẾP THEO: thông tin KH + hệ số giá bán (O8) → bản chào giá ẩn cột N–V

## CHECKPOINT 14 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v14.xlsx` (v13 backup ở `90. Backup/`)
- Quyết định user: tháp buồng lạnh TTP; bơm buồng lạnh "Bơm giải nhiệt", cột D = Ewara; mọi hãng ghi cột D (Panasonic, TTP, Ewara, Hailiang, Danfoss, Tiền Phong, Hstar, Teco, Taijan);
  NC cố định Phần I: cụm 1,5tr, dàn 2tr, tủ 1tr, tháp 1tr, bơm 300k, tiết lưu 500k (params `nc_co_dinh_buong_lanh`);
  phần còn lại của tổng NC 77tr (= (4+10)tr × 1,1 × 5) = 45,5tr phân bổ theo tỷ lệ VT cho ống đồng, gas, ống PVC, dây điện.
  Chiều cao dòng tự tính theo độ dài chữ cột B/C/D. Sale: Mr. Bùi Hoàng Nghĩa – 0964826488 – project.tiendatco@gmail.com; KH: Mr. Quang Anh.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.755,9 tr · sau VAT 2.976,4 tr; I 1.195,7 · II 1.525,3 · III 34,9
- Còn mở: tên công ty/địa chỉ KH; hệ số giá bán O8 (1,00); ẩn cột N–V khi gửi

## CHECKPOINT 15 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v15.xlsx` (v14 backup ở `90. Backup/`)
- Quyết định user: ống đồng ghi "+ bảo ôn và phụ kiện"; cụm máy model "Máy nén Panasonic 20HP + …", xuất xứ Việt Nam; "Dàn lạnh inox", model DJ 14.3/85-201, China;
  Phần I tách 4 mục: I.1 Vỏ buồng lạnh / I.2 Thiết bị / I.3 Đường ống / I.4 Đường điện; dây điện buồng lạnh 7,4tr/hệ (giá dùng nguyên, params `day_dien_buong_lanh_he`);
  tháp chiller "100RT"; thêm "Băng quấn trắng" 25.000 đ/kg, 1 kg/3 m ống nước lạnh → 19 kg.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.760,7 tr · sau VAT 2.981,5 tr; I 1.199,7 · II 1.525,9 · III 35,1
- Còn mở: công ty/địa chỉ KH; hệ số giá bán O8; ẩn cột N–V khi gửi

## CHECKPOINT 16 – 2026-11-13
- File hiện hành: `4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_v16.xlsx` (v15 backup ở `90. Backup/`)
- Quyết định user: băng quấn 40 kg × 25.000, NC 14.000/kg, bỏ ghi chú; chiller 450tr/cái; chi phí khác giảm 30% (O9: 4% → 2,8%, giữ +1,1tr vận chuyển panel);
  in đậm dòng "Rất mong…" và "ĐẠI DIỆN…/XÁC NHẬN KHÁCH HÀNG" (dòng 79–80); mọi số điện thoại (Tel, Hotline, ghi chú 11) đổi sang 0964826488; chỉnh độ rộng cột/chiều cao dòng.
- Số liệu Excel (vốn, chưa VAT): Tổng 2.901,5 tr · sau VAT 3.133,7 tr; I 1.199,7 · II 1.676,9 · III 24,9
- Còn mở: công ty/địa chỉ KH; hệ số giá bán O8; ẩn cột N–V khi gửi
