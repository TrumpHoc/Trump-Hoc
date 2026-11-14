# WP2 – HÌNH HỌC KHU H&N PHASE 2 (NHÀ XƯỞNG 3, SGS HẢI PHÒNG)

> Gói việc WP2 – v1.0 – 14/11/2026. Phụ thuộc: `01_design_basis/DESIGN_BASIS.md` (WP0).
> Kết quả: `geometry.json` (mô hình tọa độ đầy đủ), `extract_geometry.py` (script tái lập), `plan_overlay_p7.png` (phòng + điểm hút + thiết bị đề xuất trên nền DW-A-03-06), `plan_rooms_p6.png` (đa giác phòng + tường tiếp xúc trên nền DW-A-03-05).
> Chạy lại: `python3 -I 03_geometry/extract_geometry.py` (cần pdfplumber, pypdfium2, matplotlib).

## 1. Hệ tọa độ & phương pháp

- **Hệ tọa độ**: đơn vị mm, gốc tại giao trục **E/3**; X dương sang phải về trục **A** (tường bao ngoài nhà), Y dương lên trên về trục **5**. Cao độ: FFL = 0; trần panel +3.000; vùng thiết bị trên trần +3.300…+6.000 (giả định); ống gió Phase 1 BOD +4.200…+6.200.
- **Nguồn**: PDF vector xuất từ CAD `00_inputs/261109_Cai_tao_khu_HN_xuong_3.pdf`: trang 5 (DW-A-03-04, vách & cửa + kích thước thông thủy), trang 6 (DW-A-03-05, tên phòng + diện tích), trang 7 (DW-A-03-06, nội thất), trang 2 (DW-A-03-01, mặt bằng toàn xưởng – chỉ dùng cho bối cảnh ngoài tường A).
- **Các bước** (đều nằm trong `extract_geometry.py`):
  1. Hiệu chuẩn pt → mm bằng đường lưới (đường nét đứt dài) + chữ ký hiệu trục trên trang 6, hồi quy tuyến tính theo khoảng lưới 7.000 / 7.700.
  2. Đo **tâm vách** từ đường vector của vách panel (trang 6); dựng **mặt trong vách** theo kích thước thông thủy ghi trên DW-A-03-04 (panel 50 mm), rồi đối chiếu ngược với tâm vách đo được.
  3. Kiểm tra mỗi phòng: diện tích so với bảng (±5 %), và nhãn diện tích trên trang 6 phải nằm trong đa giác.
  4. Điểm hút: tâm ký hiệu thiết bị trên trang 7 (vector); mỗi điểm có "neo" là nhãn chữ trên bản vẽ ("FUME HOOD - 1600", "HOOD ACID", "TỦ HÓA CHẤT", "BÀN ĐỂ MÁY"…) – script tự kiểm tra ký tự trong khung neo. Kết quả: **32/32 neo khớp**, 35/35 điểm nằm trong đúng đa giác phòng.
  5. Tính tường tiếp xúc, đề xuất vị trí thiết bị, khoảng cách Manhattan.

## 2. Hiệu chuẩn

| Hạng mục | Kết quả |
|---|---|
| Tâm trục X (pt, trang 6) | E 122,94 · D 308,38 · C 493,84 · B 679,29 · A 864,75 |
| Tâm trục Y (pt từ mép trên) | 5: 144,86 · 4: 348,85 · 3: 552,85 |
| Tỷ lệ | **sx = 37,745 mm/pt; sy = 37,746 mm/pt** (đẳng hướng) ≈ 1:107 trên giấy A3 |
| Công thức | X = 37,74541·x_pt − 4.640,2 ; Y = −37,74602·y_top_pt + 20.867,8 |
| Kiểm tra khoảng lưới | E-D 6.999,5 · D-C 7.000,3 · C-B 6.999,9 · B-A 7.000,3 ; 3-4 7.700,2 · 4-5 7.699,8 mm |
| Sai số dư | X ≤ 0,3 mm; Y ≤ 0,1 mm |
| Trang 5 và 7 | Lưới trùng khít trang 6 (cùng tọa độ pt) → dùng chung phép biến đổi |
| Pixel render | pypdfium2 scale 2 → 2 px/pt (E ≈ 246 px, A ≈ 1.730 px – khớp ước lượng ban đầu 245/1.727) |
| Trang 2 (toàn xưởng) | 130,4 mm/pt; E x = 210,6 pt, trục 3 y = 622,2 pt (chỉ dùng định vị FPS, cửa, bậc ngoài tường A) |

Đối chiếu tâm vách đo bằng vector với mô hình dựng từ kích thước DW-A-03-04: 15/19 vách lệch ≤ 2 mm; vách Dividing/Coding −28 mm, Instrument/Reporting −12 mm, vách giữa 2 phòng Standard −51 mm và vách phía A của ST2 −49 mm (≈ 1 chiều dày panel – nét vẽ là mặt vách, không phải tâm). Không ảnh hưởng kết quả.

## 3. Phòng – đa giác, diện tích, tâm

13 phòng đều là hình chữ nhật (đa giác 4 đỉnh, mặt trong vách; tọa độ đỉnh trong `geometry.json` → `rooms[].polygon_mm`). "DT mô hình" = theo mặt trong vách đã dựng; "DT chỉ từ vector" = kiểm tra độc lập chỉ dùng tâm vách đo trên PDF ± 25 mm.

| Mã | Phòng | Kích thước thông thủy (mm) | DT bảng phòng (m²) | DT mô hình (m²) | Sai lệch | DT chỉ từ vector (m²) | Sai lệch | Tâm (X;Y) mm | BBox mm (X0;Y0 → X1;Y1) |
|---|---|---|---|---|---|---|---|---|---|
| DIG | Digestion room | 3650 × 4000 | 14,60 | 14,60 | +0,0 % | 14,61 | +0,1 % | 7720; 16675 | 5895; 14675 → 9545; 18675 |
| KJE | Kjeldahl room | 4740 × 4000 | 18,88 | 18,96 | +0,4 % | 18,96 | +0,4 % | 11965; 16675 | 9595; 14675 → 14335; 18675 |
| ST1 | Sample treatment 1 | 8440 × 7800 | 65,70 | 65,83 | +0,2 % | 65,86 | +0,2 % | 10115; 10725 | 5895; 6825 → 14335; 14625 |
| BAL | Balance room | 5735 × 3215 | 18,35 | 18,44 | +0,5 % | 18,44 | +0,5 % | 11468; 5168 | 8600; 3560 → 14335; 6775 |
| PRE | Sample preparation | 7455 × 7950 | 58,99 | 59,27 | +0,5 % | 59,25 | +0,4 % | 4822; 2800 | 1095; -1175 → 8550; 6775 |
| COD | Sample coding | 5305 × 4750 | 25,20 | 25,20 | -0,0 % | 25,34 | +0,5 % | 5898; -3600 | 3245; -5975 → 8550; -1225 |
| DIV | Sample dividing | 2100 × 4750 | 9,98 | 9,97 | -0,1 % | 9,84 | -1,4 % | 2145; -3600 | 1095; -5975 → 3195; -1225 |
| INS | Instrument room | 3785 × 9485 | 35,82 | 35,90 | +0,2 % | 35,79 | -0,1 % | 10492; -1232 | 8600; -5975 → 12385; 3510 |
| REP | Reporting room | 1900 × 9485 | 18,02 | 18,02 | +0,0 % | 18,14 | +0,7 % | 13385; -1232 | 12435; -5975 → 14335; 3510 |
| OFF | Office | 4700 × 8775 | 40,96 | 41,24 | +0,7 % | 41,23 | +0,7 % | 18585; 5262 | 16235; 875 → 20935; 9650 |
| STD20 | Standard (19,72 m²) | 2900 × 6800 | 19,72 | 19,72 | +0,0 % | 19,37 | -1,8 % | 17685; -2575 | 16235; -5975 → 19135; 825 |
| STD12 | Standard (11,90 m²) | 1750 × 6800 | 11,90 | 11,90 | +0,0 % | 12,24 | +2,8 % | 20060; -2575 | 19185; -5975 → 20935; 825 |
| ST2 | Sample treatment 2 | 6720 × 15625 | 103,64 | 105,00 | +1,3 % | 104,26 | +0,6 % | 24345; 1838 | 20985; -5975 → 27705; 9650 |

Tổng: mô hình 444,05 m² / bảng 441,76 m² (+0,5 %). **Tất cả phòng đạt ±5 %** (lớn nhất ST2 +1,3 % theo kích thước 6.720 × 15.625 ghi trên DW-A-03-04; chênh 1,36 m² nhiều khả năng do cột/hộp cột dọc tường A trừ vào diện tích; Office +0,7 % tương tự). Nhãn diện tích trang 6 của cả 13 phòng nằm trong đa giác tương ứng. Thể tích (trần 3,0 m) có trong JSON (`volume_m3_at_ceiling_3m`) cho WP1 tính ACH.

## 4. Vách chung và tường tiếp xúc không gian không điều hòa

Phân loại phía ngoài mỗi đoạn vách (điểm thử cách mặt vách 400 mm):
- **Tường ngoài A**: chỉ ST2 (vách panel ST2 song song, cách tường bao trục A ~0,3 m).
- **Xưởng KĐH**: không gian xưởng không điều hòa (phía trục E, phía trên Digestion/Kjeldahl, phía dưới hàng phòng dưới cùng).
- **Lối đi giữa 2 khối**: dải X = 14.385…16.185 (thông thủy ~1,8 m) giữa khối trái (E–C) và khối phải (B–A) – coi là không điều hòa.
- **Khu ngoài phạm vi**: khu trái trên (trên Sample prep, trái Digestion/ST1) và khu phải trên (trên Office/ST2) – không ghi tên trên DW-A-03-05; theo DW-A-03-01 tương ứng khoảng DIR (nhà ăn) và CST (kho hóa chất). Giả định **không điều hòa** (an toàn cho tải) – cần xác nhận.

Chiều dài tiếp xúc (m, theo mặt trong; diện tích = × 3,0 m chiều cao trần – có trong JSON):

| Mã | Vách chung với | Tường ngoài A (m) | Xưởng KĐH (m) | Lối đi (m) | Khu ngoài PV (m) | Tổng tiếp xúc (m) |
|---|---|---|---|---|---|---|
| DIG | KJE, ST1 | 0,00 | 3,65 | 0,00 | 4,00 | 7,65 |
| KJE | DIG, ST1 | 0,00 | 4,74 | 4,00 | 0,00 | 8,74 |
| ST1 | BAL, DIG, KJE, PRE | 0,00 | 0,00 | 7,80 | 7,80 | 15,60 |
| BAL | INS, PRE, REP, ST1 | 0,00 | 0,00 | 3,21 | 0,00 | 3,21 |
| PRE | BAL, COD, DIV, INS, ST1 | 0,00 | 7,95 | 0,00 | 4,80 | 12,75 |
| COD | DIV, INS, PRE | 0,00 | 5,30 | 0,00 | 0,00 | 5,30 |
| DIV | COD, PRE | 0,00 | 6,85 | 0,00 | 0,00 | 6,85 |
| INS | BAL, COD, PRE, REP | 0,00 | 3,79 | 0,00 | 0,00 | 3,79 |
| REP | BAL, INS | 0,00 | 1,90 | 9,48 | 0,00 | 11,38 |
| OFF | ST2, STD12, STD20 | 0,00 | 0,00 | 8,78 | 4,70 | 13,47 |
| STD20 | OFF, STD12 | 0,00 | 2,90 | 6,80 | 0,00 | 9,70 |
| STD12 | OFF, ST2, STD20 | 0,00 | 1,75 | 0,00 | 0,00 | 1,75 |
| ST2 | OFF, STD12 | 15,62 | 6,72 | 0,00 | 6,72 | 29,07 |

Ghi chú: trần panel PU 50 mm tiếp xúc khoang trên trần (không điều hòa) cho toàn bộ diện tích sàn; sàn trên nền đất. Danh sách từng đoạn vách (tọa độ đầu/cuối, loại) trong `rooms[].exposure.exposed_segments`.

## 5. Điểm hút – tọa độ và đối chiếu số lượng

"Hướng mặt" = hướng mặt thao tác (sash) nhìn vào phòng (N = +Y, S = −Y, E = +X, W = −X). Q theo DESIGN_BASIS §3. "Kiểm nhãn" = ký tự nhãn trên PDF khớp neo.

| ID | Phòng | Loại | Acid | X (mm) | Y (mm) | Hướng mặt | Q (m³/h) | Kiểm nhãn | Ghi chú |
|---|---|---|---|---|---|---|---|---|---|
| DIG-FH1 | DIG | Tủ hood 1600 | X | 6668 | 18301 | S | 2270 | OK | Hood acid đặt trên bàn 1600×800 sát vách trên |
| DIG-FH2 | DIG | Tủ hood 1600 | X | 8469 | 18301 | S | 2270 | OK | Hood acid 1600 đặt trên bàn 2000×800 sát vách trên |
| DIG-CH1 | DIG | Chụp treo |  | 8518 | 15104 | N | 1100 | OK | Chụp treo 2000×800 trên bàn lò nung |
| KJE-CH1 | KJE | Chụp treo |  | 11711 | 15104 | N | 1100 | OK | Chụp treo 2000×800 trên bàn (máy Kjeldahl) |
| ST1-FH1 | ST1 | Tủ hood 1600 | X | 6268 | 13855 | E | 2270 | OK | Hood acid, vách trái |
| ST1-FH2 | ST1 | Tủ hood 1600 |  | 6268 | 12250 | E | 2270 | OK | Hood hữu cơ (lọc than) |
| ST1-FH3 | ST1 | Tủ hood 1600 |  | 6268 | 10650 | E | 2270 | OK | Hood hữu cơ (lọc than) |
| ST1-CC1 | ST1 | Tủ hóa chất |  | 13980 | 14051 | W | 1300 | OK | Tủ hóa chất 2000×650 (mua mới), vách phải |
| ST1-CC2 | ST1 | Tủ hóa chất |  | 8910 | 7177 | N | 1300 | OK | Tủ hóa chất 2000×650 (mua mới), vách dưới – CẦN XÁC NHẬN có hút (cơ sở TK chỉ 1 tủ) |
| ST2-FH1 | ST2 | Tủ hood 1600 |  | 22906 | 9280 | S | 2270 | OK | Hàng 3 hood sát vách trên |
| ST2-FH2 | ST2 | Tủ hood 1600 |  | 24507 | 9280 | S | 2270 | OK |  |
| ST2-FH3 | ST2 | Tủ hood 1600 | X | 26111 | 9280 | S | 2270 | OK | GIẢ ĐỊNH là hood acid (bản vẽ không phân biệt; chọn hood gần trục A/scrubber nhất) |
| ST2-FH4 | ST2 | Tủ hood 1600 |  | 21381 | 7385 | E | 2270 | OK | Vách trái ST2 |
| ST2-FH5 | ST2 | Tủ hood 1600 |  | 21381 | 334 | E | 2270 | OK | Vách trái ST2 |
| ST2-FH6 | ST2 | Tủ hood 1600 |  | 21381 | -1270 | E | 2270 | OK | Vách trái ST2 |
| ST2-SK1 | ST2 | Chụp bồn rửa |  | 23473 | 6675 | E | 300 | OK | Bồn rửa đôi 1400×800 (đầu bàn đảo) – chụp hút |
| ST2-SK2 | ST2 | Chụp bồn rửa |  | 23473 | 1685 | E | 300 | OK | Bồn rửa đôi 1400×800 (đầu bàn đảo) – chụp hút |
| ST2-CC1 | ST2 | Tủ hóa chất |  | 27274 | 1391 | W | 1300 | OK | Tủ hóa chất 800×800, vách phải (trục A) |
| ST2-CC2 | ST2 | Tủ hóa chất |  | 27274 | -568 | W | 1300 | OK | Tủ hóa chất 800×800, vách phải (trục A) |
| STD20-FH1 | STD20 | Tủ hood 1600 |  | 18686 | 89 | W | 2270 | OK | Sát vách ngăn 2 phòng Standard |
| STD20-FH2 | STD20 | Tủ hood 1600 |  | 18686 | -1482 | W | 2270 | OK |  |
| STD20-FH3 | STD20 | Tủ hood 1600 |  | 18686 | -5113 | W | 2270 | OK | Góc dưới phải |
| STD20-FH4 | STD20 | Tủ hood 1600 |  | 16614 | -5113 | E | 2270 | OK | Góc dưới trái |
| STD20-SK1 | STD20 | Chụp bồn rửa |  | 16610 | 51 | E | 300 | – | Bồn rửa đôi (ký hiệu, không nhãn) góc trên trái – chụp hút |
| STD12-CC1 | STD12 | Tủ hóa chất |  | 20366 | -5351 | W | 1300 | OK | "TỦ KHÔ" = tủ khô chứa hóa chất 1000×600 có hút (DM nội thất 9.6) |
| PRE-AH1 | PRE | Arm hood |  | 3169 | 2316 | - | 800 | – | GIẢ ĐỊNH: arm hood treo trên đảo bàn trái (bản vẽ không thể hiện) |
| PRE-AH2 | PRE | Arm hood |  | 5819 | 2316 | - | 800 | – | GIẢ ĐỊNH: arm hood treo trên đảo bàn phải (bản vẽ không thể hiện) |
| INS-V1 | INS | Điểm hút TB |  | 9729 | 1051 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V2 | INS | Điểm hút TB |  | 9729 | -949 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V3 | INS | Điểm hút TB |  | 9729 | -2954 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V4 | INS | Điểm hút TB |  | 9729 | -4958 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V5 | INS | Điểm hút TB |  | 11228 | 1051 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V6 | INS | Điểm hút TB |  | 11228 | -949 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V7 | INS | Điểm hút TB |  | 11228 | -2954 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |
| INS-V8 | INS | Điểm hút TB |  | 11228 | -4958 | - | 150 | OK | Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm) |

### 5.1 Đối chiếu với DESIGN_BASIS §3

| Phòng | Hạng mục | Bản vẽ DW-A-03-06 | DESIGN_BASIS | Kết luận |
|---|---|---|---|---|
| Digestion | Hood 1600 (acid) | 2 (2) – nhãn "HOOD ACID" | 2 (2) | Khớp |
| Digestion | Chụp treo | 1 (bàn lò nung) | 1 | Khớp; DM nội thất mục 6.5 ghi **2 chụp** → cần xác nhận |
| Kjeldahl | Chụp treo | 1 (bàn 2000×800 vách dưới) | 1 | Khớp; DM nội thất mục 7.4 ghi **2 chụp** (có lọc than); máy chưng cất đạm tự động góc trên trái chưa có chụp → cần xác nhận |
| ST1 | Hood 1600 (acid) | 3 (1) – 1 "HOOD ACID" + 2 "HOOD HỮU CƠ" | 3 (1) | Khớp |
| ST1 | Tủ hóa chất | **2** "CHEMICAL (MUA MỚI) 2000×650" (vách phải + vách dưới) | 1 | **Lệch +1** – mô hình giữ cả 2 (ST1-CC2 cần xác nhận), +1.300 m³/h nếu cả 2 có hút |
| ST2 | Hood 1600 (acid) | 6 – 3 sát vách trên + 3 sát vách trái | 6 (1) | Số lượng khớp; bản vẽ **không phân biệt acid** → GIẢ ĐỊNH ST2-FH3 (gần trục A/scrubber nhất) |
| ST2 | Tủ hóa chất | **2** "TỦ HÓA CHẤT 800×800" | 3 | **Lệch −1** (tủ thứ 3 ở Standard 11,90) |
| ST2 | Chụp bồn rửa | **2** "BỒN RỬA (MUA MỚI) 1400×800" | 3 | **Lệch −1** (bồn thứ 3 ở Standard 19,72) |
| Standard 19,72 | Hood 1600 | **4** (2 giữa sát vách phải + 2 góc dưới) | 3 | **Lệch +1** |
| Standard 19,72 | Chụp bồn rửa | **1** (ký hiệu bồn đôi góc trên trái) | 0 | **Lệch +1** |
| Standard 11,90 | Hood 1600 | **0** (chỉ tủ lạnh, bàn cân, tủ khô) | 1 | **Lệch −1** |
| Standard 11,90 | Tủ hóa chất có hút | **1** "TỦ KHÔ" (DM 9.6: tủ khô 1000×600 có hút) | 0 | **Lệch +1** |
| Sample prep | Arm hood | không thể hiện | 2 × 800 | GIẢ ĐỊNH 1 arm hood/đảo bàn (PRE-AH1, PRE-AH2) |
| Instrument | Điểm hút thiết bị | 8 "BÀN ĐỂ MÁY 2000×1000" | 8 × 150 | Khớp (1 điểm/bàn) |

**Tổng** (bản vẽ / cơ sở TK): hood 1600 15/15 (acid 4/4); chụp treo 2/2; tủ hóa chất 5/4; chụp bồn rửa 3/3; arm hood 2/2 (giả định); điểm hút TB 8/8.
Nhóm **ST2 + 2 phòng Standard** (DM nội thất mục 9 gộp 3 phòng: 9 hood hữu cơ + 1 hood acid, 3 bồn rửa đôi có chụp, 3 tủ khô có hút): **tổng khớp 10 / 3 / 3**, chỉ **phân bổ theo phòng khác** DESIGN_BASIS → WP1 cần cân bằng gió theo bản vẽ: STD20 = 4 hood + 1 bồn (9.380 m³/h); STD12 = 1 tủ khô (1.300 m³/h); ST2 = 6 hood + 2 tủ + 2 bồn (16.820 m³/h).

Lưu lượng hút theo bản vẽ (Q cơ sở TK): acid 4 × 2.270 = 9.080 m³/h → scrubber; hood hữu cơ 11 × 2.270 = 24.970; chụp treo 2.200; tủ hóa chất 5 × 1.300 = 6.500 (5.200 nếu ST1-CC2 không hút); chụp bồn 900; arm hood 1.600; điểm hút TB 1.200 m³/h.

## 6. Cửa, lối đi, khu ngoài phạm vi

Tâm cửa lấy từ cung mở cửa (vector trang 6) và nhãn cửa trang 5. FACTORY_W/S/N = xưởng phía tây/nam/bắc; OUTSIDE_A = ngoài nhà qua tường A.

| ID | Giữa | Vách | Tâm (X;Y) | Rộng | Loại / ghi chú |
|---|---|---|---|---|---|
| D-01 | DIG ↔ KJE | x | 9571; 16678 | 900 | cánh đơn – Digestion chỉ đi qua Kjeldahl |
| D-02 | KJE ↔ ST1 | y | 10148; 14650 | 1000 | cánh đơn D4.5 (tận dụng) |
| D-03 | KJE ↔ CORRIDOR | x | 14360; 16950 | 1400 | cửa đôi 1,4 m (giữ hiện trạng) |
| D-04 | ST1 ↔ CORRIDOR | x | 14360; 8019 | 1400 | cửa đôi lệch cánh (hiện trạng) |
| D-05 | BAL ↔ CORRIDOR | x | 14360; 5241 | 1400 | cửa đôi lệch cánh (hiện trạng) |
| D-06 | REP ↔ CORRIDOR | x | 14360; 2580 | 1400 | cửa đôi lệch cánh (hiện trạng) |
| D-07 | INS ↔ REP | x | 12410; 2760 | 1400 | D_1.4x2.2 (làm mới) – Instrument vào qua Reporting |
| D-08 | BAL ↔ PRE | x | 8575; 5128 | 1000 | cửa lùa D_1x2.2 (làm mới) |
| D-09 | PRE ↔ FACTORY_W | x | 1070; 5241 | 1400 | cửa đôi lệch cánh (hiện trạng) |
| D-10 | PRE ↔ COD | y | 7808; -1200 | 1000 | D_1x2.2 (tận dụng) |
| D-11 | COD ↔ FACTORY_S | y | 7627; -6000 | 1400 | D_1.4x2.2 (làm mới) |
| D-12 | DIV ↔ FACTORY_W | x | 1070; -3320 | 1000 | D_1x2.2 (tận dụng), đổi chiều mở |
| PB-01 | COD ↔ DIV | x | 3220; -2384 | 800 | cửa inox 800×800 kiểu passbox, đáy cách sàn 0,9–1,0 m |
| D-13 | OFF ↔ CORRIDOR | x | 16210; 5177 | 1600 | cửa đôi 1,6 m |
| D-14 | OFF ↔ ST2 | x | 20960; 2448 | 1000 | cửa lùa 1 m D_1x2.2 (làm mới) |
| D-15 | STD20 ↔ CORRIDOR | x | 16210; -3101 | 1400 | cửa đôi lệch cánh |
| D-16 | STD20 ↔ STD12 | x | 19160; -3101 | 1400 | D_1.4x2.2 (làm mới) |
| D-17 | STD12 ↔ ST2 | x | 20960; -3101 | 1400 | D_1.4x2.2 (làm mới) |
| D-18 | ST2 ↔ FACTORY_S | y | 22661; -6000 | 1600 | cửa đôi 1,6 m mở 2 chiều D_1.6x2.2 (làm mới) |
| D-N1 | DINING_nonscope ↔ FACTORY_W | x | 1070; 12526 | 1400 | hiện trạng – ngoài phạm vi |
| D-N2 | CST_nonscope ↔ CORRIDOR | x | 16210; 10801 | 1700 | cửa đôi hiện trạng – ngoài phạm vi |
| D-N3 | CST_nonscope ↔ FACTORY_N | y | 26538; 18700 | 1700 | cửa đôi hiện trạng – ngoài phạm vi |
| D-N4 | CST_nonscope ↔ OUTSIDE_A | x | 28000; 13726 | 900 | cửa ra ngoài trên tường trục A – giữ lối ra – không đặt thiết bị trước cửa |
| D-N5 | ST2 (khe tường A) ↔ OUTSIDE_A | x | 28000; -850 | 900 | EXISTING D3 (theo DW-A-03-01) – cần khảo sát: có còn sau cải tạo? |

- **Lối đi giữa 2 khối** (X 14.385…16.185, suốt Y −6.000…18.700): khối trái mở cửa ra lối đi qua vách X = 14.360 (Kjeldahl, ST1, Balance, Reporting); khối phải qua vách X = 16.210 (Office, Standard 19,72, cửa đôi khu CST). Đây là hành lang kỹ thuật tự nhiên để **vượt ống chính từ khối trái sang trục A** (treo cao +4,0…+5,5 trên lối đi).
- Chuỗi đi lại: Digestion → Kjeldahl → ST1 → lối đi; Instrument → Reporting → lối đi; Coding ↔ Dividing chỉ qua passbox; Standard 19,72 → Standard 11,90 → ST2; Office ↔ ST2 cửa lùa; Sample prep, Dividing, Coding, ST2 có cửa ra xưởng.
- **Khu ngoài phạm vi**: (1) trái trên X 1.095…5.845, Y 6.825…18.675 (~56 m², theo DW-A-03-01 ~ DIR nhà ăn); (2) phải trên X 16.235…27.705, Y 9.700…18.675 (~103 m², theo DW-A-03-01 ~ CST kho hóa chất, có cửa ra tường A tại Y ≈ 13.300…14.200). Không thiết kế HVAC cho 2 khu này; tuy nhiên tuyến ống phía bắc (scrubber, DP-2.01) phải đi trên trần khu CST → cần CĐT cho phép và xác nhận kết cấu treo.

## 7. Đề xuất vị trí thiết bị

### 7.1 Ràng buộc
- Tường bao ngoài duy nhất của khu là **trục A** (X = 28.000). Các phía khác là xưởng/lối đi/khu hiện hữu → louver gió tươi, điểm xả, dàn nóng, scrubber đều ở **ngoài tường A**.
- Ngoài tường A (theo DW-A-03-01): **FPS – phòng FUMI & PCO** (kho hóa chất xông hơi, X 28.390…31.010, Y 640…8.650); cửa CST Y ≈ 13.300…14.200; cửa "EXISTING D3" Y ≈ −300…−1.400 (cần khảo sát còn hay không sau cải tạo); bậc/ram dốc Y −7.600…−5.900; cửa sổ hiện hữu Y ≈ 10.100…13.000 và −2.500…−4.500.
- Vùng trên trần +3,3…+6,0 m nhưng ống Phase 1 ở +4,2…+6,2 → ống chính Phase 2 nên đi **+3,3…+4,1** nơi giao cắt Phase 1; AHU (cao ~1,8–2,0 m) phải đặt nơi không có ống Phase 1 phía trên (cần khảo sát).

### 7.2 Phân vùng ngoài tường A (từ nam lên bắc)

| Y (mm) | Sử dụng | Lý do |
|---|---|---|
| −15.000…−9.000 | **Bãi CDU** (CDU-2.01/2.02/(2.03)) | Dải trống dọc tường A, không chặn cửa/bậc, cách louver ≥ 3,6 m, CDU thổi gió nóng lên trên |
| −7.600…−5.900 | Bậc/ram dốc hiện hữu | Giữ nguyên |
| −5.400…−2.000 | **Louver gió tươi** LV-2.01/2.02/2.03 tại +4,0…+5,5 | Cách mọi điểm xả ≥ 12 m; sát AHU-2.02 |
| −1.400…−300 | Cửa EXISTING D3 | Giữ lối |
| 640…8.650 | FPS (kho fumigation) | Không đặt louver/điểm xả sát FPS |
| 9.600…12.000 | **Điểm xả** DP-2.03, DP-2.02 (ống đứng vượt mái) | Gần ST2 (nhiều hood nhất) |
| 13.300…14.200 | Cửa CST ra ngoài | Giữ lối |
| 15.000…17.500 | **Scrubber SCR-2.01** + ống xả DP-2.00 | Gần trọng tâm 4 hood acid (3/4 ở khối trái phía bắc), cách FPS > 6 m, cách louver ≥ 19 m |
| 18.200 | **Điểm xả** DP-2.01 | Tuyến ngắn nhất từ khối trái phía bắc |

### 7.3 Thiết bị đề xuất (tâm, kích thước chiếm chỗ sơ bộ – WP1 chốt kích thước)

| ID | Loại | Tâm (X;Y) mm | Kích thước chiếm chỗ (mm) | Cao độ | Vị trí | Phục vụ |
|---|---|---|---|---|---|---|
| AHU-2.01 | AHU | 11700; 5150 | 4800×2000 | +3,3…+5,3 (sàn kết cấu trên trần) | Trên trần Balance room, sát lối đi giữa 2 khối (trục C) | DIG, KJE, ST1, BAL, PRE, COD, DIV, REP, INS* |
| AHU-2.02 | AHU | 24600; -3300 | 4800×2000 | +3,3…+5,3 | Trên trần ST2 (khu "đặt máy tương lai"), cách tường A ~3 m | OFF, STD20, STD12, ST2 |
| AHU-2.03 | AHU (phương án) | 10500; -3800 | 3000×1600 | +3,3…+5,0 | Trên trần Instrument room (nếu WP1 chọn tách AHU) | INS |
| DH-2.01 | Máy hút ẩm | 11900; -1200 | 1200×800 | +3,3… (dạng nối ống) | Trên trần Instrument, cạnh AHU-2.03 | INS |
| FAF-2.01 | Quạt cấp bù FAFE (G4+F8) | 13000; 11000 | 1500×900 | +3,3…+4,3 | Trên trần ST1 sát lối đi | DIG, KJE, ST1 |
| FAF-2.02 | Quạt cấp bù FAFE (G4+F8) | 24600; 400 | 1500×900 | +3,3…+4,3 | Trên trần ST2 gần AHU-2.02 và louver LV-2.02 | ST2, STD20, STD12 |
| EAF-2.01 | Quạt hút EAF (SS304) | 13000; 16000 | 1200×900 | +3,3…+4,3 | Trên trần Kjeldahl sát lối đi | DIG-CH1, KJE-CH1, ST1-CC1, ST1-CC2 |
| EAF-2.02 | Quạt hút EAF (SS304) | 26500; 5000 | 1200×900 | +3,3…+4,3 | Trên trần ST2 phía tường A | ST2-SK1, ST2-SK2, ST2-CC1, ST2-CC2, STD20-SK1, STD12-CC1 |
| EAF-2.03 | Quạt hút EAF (GI/SS) | 8575; 1500 | 1200×900 | +3,3…+4,3 | Trên vách Sample prep/Instrument | PRE-AH1, PRE-AH2, INS-V1, INS-V2, INS-V3, INS-V4, INS-V5, INS-V6, INS-V7, INS-V8 |
| SCR-2.01 | Scrubber ướt (PP) – ngoài nhà | 30000; 16250 | 2000×2500 | mặt đất ngoài nhà, ống xả ≥ đỉnh mái | Ngoài tường trục A, phía bắc cửa CST (D-N4), cách FPS > 6 m | DIG-FH1, DIG-FH2, ST1-FH1, ST2-FH3 |
| CDU-YARD | Bãi dàn nóng CDU-2.01/2.02/(2.03) | 29500; -12000 | 2000×6000 | mặt đất (móng: xây dựng) | Ngoài tường A, phía nam bậc/ram dốc hiện hữu (Y -5,9…-7,6 m) | AHU-2.01, AHU-2.02, AHU-2.03 |
| DB.HVAC | Tủ điện tổng HVAC (gợi ý cho WP4) | 16400; 2200 | 800×400 | sàn | Vách khối phải phía lối đi (giữa D-13 và D-15) |  |

Lý do chọn:
- **AHU-2.01** trên trần Balance (X 9,3…14,1 m; Y 4,15…6,15 m): gần trọng tâm tải khối trái, sát lối đi nên OA/ống gas vượt lối đi ngắn nhất; ống cấp chia 2 nhánh bắc (ST1, Kjeldahl, Digestion) và nam/tây (Sample prep, Coding, Dividing, Reporting, Instrument). Nhược điểm: OA và ống gas tới tường A dài (~24 m và ~35 m Manhattan) – không tránh được vì khối trái không có tường ngoài.
- **AHU-2.02** trên trần ST2 phía đông-nam (khu "đặt máy tương lai"): OA từ LV-2.02 ~4 m, ống gas tới bãi CDU ~14 m, nước ngưng ~9 m; cấp gió cho Office/Standard ≤ 15 m.
- **AHU-2.03** (phương án) + **DH-2.01** trên trần Instrument: tải thiết bị ~31 kW tập trung, ống phân phối 2–3 m; nếu WP1 không tách, Instrument nhập AHU-2.01 (7,6 m).
- **FAF-2.01 / FAF-2.02**: đặt cạnh nhóm phòng hóa tương ứng; FAF-2.02 dùng chung louver với AHU-2.02. FAF-2.01 lấy gió tươi qua LV-2.01 (~28 m) – nên gộp chung ống OA trục chính với AHU-2.01 trên lối đi.
- **EAF-2.01/2.02/2.03**: đặt ở trọng tâm nhóm điểm hút (nhánh ≤ 16 m), xả ra cụm điểm xả phía bắc tường A.
- **Scrubber**: hood acid ở Digestion (Y 18,3 m), ST1 (Y 13,9 m), ST2 (Y 9,3 m) → ngoài tường A tại Y ≈ 16 m: tổng ống PP ≈ 86 m Manhattan, không chặn cửa CST, xa FPS và louver.
- **Thoát nước ngưng**: điểm giả định CD-2.01 ngoài tường A (Y ≈ −8.500, cạnh bãi CDU). Ống uPVC trên trần, dốc 1 %: AHU-2.01 dài ~30 m → tụt 0,30 m (cần đáy khay AHU ≥ +3,6 m nếu xuyên tường ở +3,3) → **tự chảy khả thi** nếu bệ AHU đủ cao; dự phòng bơm nước ngưng cho AHU-2.01, AHU-2.03, DH-2.01 (tuyến 22–30 m).
- **DB.HVAC** (gợi ý cho WP4): vách khối phải phía lối đi, X ≈ 16.400, Y ≈ 2.200.

### 7.4 Louver, điểm xả, thoát nước ngưng

| ID | Loại | (X;Y) mm | Cao độ | Ghi chú |
|---|---|---|---|---|
| LV-2.01 | louver_OA | 28000; -2400 | +4,0…+5,5 | Gió tươi AHU-2.01 + FAF-2.01 (+AHU-2.03 nếu gộp) |
| LV-2.02 | louver_OA | 28000; -3700 | +4,0…+5,5 | Gió tươi AHU-2.02 + FAF-2.02 |
| LV-2.03 | louver_OA | 28000; -5000 | +4,0…+5,5 | Gió tươi AHU-2.03 (phương án) |
| DP-2.00 | discharge | 30700; 17300 | ống xả scrubber, đỉnh ≥ mái +3 m | Khí acid sau scrubber |
| DP-2.01 | discharge | 28000; 18200 | xuyên tường A tại +4,5 → ống đứng vượt mái | Hood hữu cơ ST1 (2) + EAF-2.01 |
| DP-2.02 | discharge | 28000; 12000 | xuyên tường A tại +4,5 → ống đứng vượt mái | Hood hữu cơ ST2 (5) + Standard (4) + EAF-2.02 |
| DP-2.03 | discharge | 28000; 9600 | xuyên tường A tại +4,5 → ống đứng vượt mái | EAF-2.03 (arm hood + Instrument) |
| CD-2.01 | condensate_drain | 28100; -8500 | ngoài nhà, cao độ đất | Điểm thoát nước ngưng giả định (rãnh/ga thoát nước mưa ngoài tường A) |

Khoảng cách louver ↔ điểm xả: nhỏ nhất **12,0 m** (LV-2.01 ↔ DP-2.03); tất cả ≥ 10 m (bảng `intake_exhaust_separation`). Louver ↔ mép FPS ≈ 3 m theo phương ngang (louver đặt cao +4,0…+5,5) – cần kiểm tra FPS có miệng thông gió/xả không.

## 8. Chiều dài tuyến sơ bộ (cho WP3 bóc tách)

L Manhattan = |ΔX| + |ΔY| (đi trên trần/dọc lối đi). **L ước tính = 1,10 × L Manhattan + phần đứng** (1,5 m ống gió vào phòng; 2,0–2,5 m ống hút từ cổ tủ/chụp; 0,5 m OA; 3,5 m nước ngưng; 4,5 m ống gas gồm đoạn xuống bãi CDU). Ma trận đầy đủ thiết bị × (tâm phòng, mọi điểm hút, louver/điểm xả) trong `geometry.json` → `distance_matrix_manhattan_mm`.

| Hệ | Từ | Đến | L Manhattan (m) | L ước tính (m) |
|---|---|---|---|---|
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | DIG | 15,5 | 18,6 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | KJE | 11,8 | 14,5 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | ST1 | 7,2 | 9,4 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | BAL | 0,2 | 1,8 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | PRE | 9,2 | 11,7 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | COD | 14,6 | 17,5 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | DIV | 18,3 | 21,6 |
| SA/RA AHU-2,01 (tới tâm phòng) | AHU-2,01 | REP | 8,1 | 10,4 |
| SA/RA AHU-2,01 (nếu không tách AHU-2,03) | AHU-2,01 | INS | 7,6 | 9,8 |
| SA/RA AHU-2,02 (tới tâm phòng) | AHU-2,02 | OFF | 14,6 | 17,5 |
| SA/RA AHU-2,02 (tới tâm phòng) | AHU-2,02 | STD20 | 7,6 | 9,9 |
| SA/RA AHU-2,02 (tới tâm phòng) | AHU-2,02 | STD12 | 5,3 | 7,3 |
| SA/RA AHU-2,02 (tới tâm phòng) | AHU-2,02 | ST2 | 5,4 | 7,4 |
| SA/RA AHU-2,03 | AHU-2,03 | INS | 2,6 | 4,3 |
| Máy hút ẩm | DH-2,01 | INS | 1,4 | 3,1 |
| OA | LV-2,01 | AHU-2,01 | 23,9 | 26,7 |
| OA | LV-2,01 | FAF-2,01 | 28,4 | 31,7 |
| OA | LV-2,02 | AHU-2,02 | 3,8 | 4,7 |
| OA | LV-2,02 | FAF-2,02 | 7,5 | 8,8 |
| OA | LV-2,03 | AHU-2,03 | 18,7 | 21,1 |
| Ống gas AHU→CDU | AHU-2,01 | CDU-YARD | 35,0 | 42,9 |
| Nước ngưng | AHU-2,01 | CD-2,01 | 30,1 | 36,6 |
| Ống gas AHU→CDU | AHU-2,02 | CDU-YARD | 13,6 | 19,5 |
| Nước ngưng | AHU-2,02 | CD-2,01 | 8,7 | 13,1 |
| Ống gas AHU→CDU | AHU-2,03 | CDU-YARD | 27,2 | 34,4 |
| Nước ngưng | AHU-2,03 | CD-2,01 | 22,3 | 28,0 |
| Nước ngưng | DH-2,01 | CD-2,01 | 23,5 | 29,4 |
| Gió bù FAFE | FAF-2,01 | DIG-FH1 | 13,6 | 16,5 |
| Hút acid → scrubber (PP) | DIG-FH1 | SCR-2,01 | 25,4 | 30,4 |
| Gió bù FAFE | FAF-2,01 | DIG-FH2 | 11,8 | 14,5 |
| Hút acid → scrubber (PP) | DIG-FH2 | SCR-2,01 | 23,6 | 28,4 |
| Gió bù FAFE | FAF-2,01 | DIG-CH1 | 8,6 | 10,9 |
| Gió bù FAFE | FAF-2,01 | KJE-CH1 | 5,4 | 7,4 |
| Gió bù FAFE | FAF-2,01 | ST1-FH1 | 9,6 | 12,0 |
| Hút acid → scrubber (PP) | ST1-FH1 | SCR-2,01 | 26,1 | 31,2 |
| Gió bù FAFE | FAF-2,01 | ST1-FH2 | 8,0 | 10,3 |
| Hood hữu cơ → xả (SS304) | ST1-FH2 | DP-2,01 | 27,7 | 33,0 |
| Gió bù FAFE | FAF-2,01 | ST1-FH3 | 7,1 | 9,3 |
| Hood hữu cơ → xả (SS304) | ST1-FH3 | DP-2,01 | 29,3 | 34,7 |
| Gió bù FAFE | FAF-2,01 | ST1-CC1 | 4,0 | 5,9 |
| Gió bù FAFE | FAF-2,01 | ST1-CC2 | 7,9 | 10,2 |
| Gió bù FAFE | FAF-2,02 | ST2-FH1 | 10,6 | 13,1 |
| Hood hữu cơ → xả (SS304) | ST2-FH1 | DP-2,02 | 7,8 | 11,1 |
| Gió bù FAFE | FAF-2,02 | ST2-FH2 | 9,0 | 11,4 |
| Hood hữu cơ → xả (SS304) | ST2-FH2 | DP-2,02 | 6,2 | 9,3 |
| Gió bù FAFE | FAF-2,02 | ST2-FH3 | 10,4 | 12,9 |
| Hút acid → scrubber (PP) | ST2-FH3 | SCR-2,01 | 10,9 | 14,4 |
| Gió bù FAFE | FAF-2,02 | ST2-FH4 | 10,2 | 12,7 |
| Hood hữu cơ → xả (SS304) | ST2-FH4 | DP-2,02 | 11,2 | 14,9 |
| Gió bù FAFE | FAF-2,02 | ST2-FH5 | 3,3 | 5,1 |
| Hood hữu cơ → xả (SS304) | ST2-FH5 | DP-2,02 | 18,3 | 22,6 |
| Gió bù FAFE | FAF-2,02 | ST2-FH6 | 4,9 | 6,9 |
| Hood hữu cơ → xả (SS304) | ST2-FH6 | DP-2,02 | 19,9 | 24,4 |
| Gió bù FAFE | FAF-2,02 | ST2-SK1 | 7,4 | 9,6 |
| Gió bù FAFE | FAF-2,02 | ST2-SK2 | 2,4 | 4,2 |
| Gió bù FAFE | FAF-2,02 | ST2-CC1 | 3,7 | 5,5 |
| Gió bù FAFE | FAF-2,02 | ST2-CC2 | 3,6 | 5,5 |
| Gió bù FAFE | FAF-2,02 | STD20-FH1 | 6,2 | 8,3 |
| Hood hữu cơ → xả (SS304) | STD20-FH1 | DP-2,02 | 21,2 | 25,8 |
| Gió bù FAFE | FAF-2,02 | STD20-FH2 | 7,8 | 10,1 |
| Hood hữu cơ → xả (SS304) | STD20-FH2 | DP-2,02 | 22,8 | 27,6 |
| Gió bù FAFE | FAF-2,02 | STD20-FH3 | 11,4 | 14,1 |
| Hood hữu cơ → xả (SS304) | STD20-FH3 | DP-2,02 | 26,4 | 31,6 |
| Gió bù FAFE | FAF-2,02 | STD20-FH4 | 13,5 | 16,3 |
| Hood hữu cơ → xả (SS304) | STD20-FH4 | DP-2,02 | 28,5 | 33,8 |
| Gió bù FAFE | FAF-2,02 | STD20-SK1 | 8,3 | 10,7 |
| Gió bù FAFE | FAF-2,02 | STD12-CC1 | 10,0 | 12,5 |
| Hút EAF nhánh | DIG-CH1 | EAF-2,01 | 5,4 | 7,9 |
| Hút EAF nhánh | KJE-CH1 | EAF-2,01 | 2,2 | 4,4 |
| Hút EAF nhánh | ST1-CC1 | EAF-2,01 | 2,9 | 5,2 |
| Hút EAF nhánh | ST1-CC2 | EAF-2,01 | 12,9 | 16,2 |
| Hút EAF nhánh | ST2-SK1 | EAF-2,02 | 4,7 | 7,2 |
| Hút EAF nhánh | ST2-SK2 | EAF-2,02 | 6,3 | 9,0 |
| Hút EAF nhánh | ST2-CC1 | EAF-2,02 | 4,4 | 6,8 |
| Hút EAF nhánh | ST2-CC2 | EAF-2,02 | 6,3 | 9,0 |
| Hút EAF nhánh | STD20-SK1 | EAF-2,02 | 14,8 | 18,3 |
| Hút EAF nhánh | STD12-CC1 | EAF-2,02 | 16,5 | 20,1 |
| Hút EAF nhánh | PRE-AH1 | EAF-2,03 | 6,2 | 8,8 |
| Hút EAF nhánh | PRE-AH2 | EAF-2,03 | 3,6 | 5,9 |
| Hút EAF nhánh | INS-V1 | EAF-2,03 | 1,6 | 3,8 |
| Hút EAF nhánh | INS-V2 | EAF-2,03 | 3,6 | 6,0 |
| Hút EAF nhánh | INS-V3 | EAF-2,03 | 5,6 | 8,2 |
| Hút EAF nhánh | INS-V4 | EAF-2,03 | 7,6 | 10,4 |
| Hút EAF nhánh | INS-V5 | EAF-2,03 | 3,1 | 5,4 |
| Hút EAF nhánh | INS-V6 | EAF-2,03 | 5,1 | 7,6 |
| Hút EAF nhánh | INS-V7 | EAF-2,03 | 7,1 | 9,8 |
| Hút EAF nhánh | INS-V8 | EAF-2,03 | 9,1 | 12,0 |
| Hút EAF → xả | EAF-2,01 | DP-2,01 | 17,2 | 20,4 |
| Hút EAF → xả | EAF-2,02 | DP-2,02 | 8,5 | 10,8 |
| Hút EAF → xả | EAF-2,03 | DP-2,03 | 27,5 | 31,8 |

## 9. Giả định
1. Mặt trong vách theo kích thước thông thủy DW-A-03-04; panel 50 mm; trần +3,0 m.
2. Khu ngoài phạm vi (trái trên, phải trên), lối đi giữa 2 khối và xưởng xung quanh: **không điều hòa**.
3. ST2-FH3 là hood acid của ST2 (bản vẽ không phân biệt).
4. Arm hood Sample prep: 1 cái/đảo bàn, treo trần tại tâm đảo.
5. Điểm hút Instrument: 1 điểm/bàn để máy (8 bàn).
6. Cả 2 tủ "CHEMICAL" ở ST1 có hút (bảo thủ); "TỦ KHÔ" ở Standard 11,90 là tủ khô có hút của DM 9.6.
7. Vị trí ngoài nhà (FPS, cửa D3, bậc, cửa sổ) lấy từ DW-A-03-01 (bản vẽ toàn xưởng, có thể là hiện trạng trước cải tạo).
8. Bãi CDU Y −15…−9 m ngoài tường A còn trống; móng/bệ thuộc gói xây dựng.
9. Có sàn kết cấu/khung đỡ trên trần cho AHU/quạt tại vị trí đề xuất; chiều cao mái đủ cho ống đứng xả vượt mái.
10. Điểm thoát nước ngưng ngoài tường A (rãnh/ga thoát nước mưa) tại Y ≈ −8,5 m.

## 10. Vấn đề còn mở (cần CĐT/GGV xác nhận)
1. **Phân bổ hood/tủ/bồn giữa ST2 – Standard 19,72 – Standard 11,90** khác DESIGN_BASIS (tổng khớp). Đề nghị cập nhật DESIGN_BASIS §3 theo bản vẽ: STD20 4 hood + 1 bồn; STD12 0 hood + 1 tủ khô; ST2 6 hood + 2 tủ + 2 bồn.
2. **ST1 có 2 tủ hóa chất** trên bản vẽ (cơ sở TK 1) – tủ nào có hút?
3. Hood nào trong 6 hood ST2 là hood acid? (ảnh hưởng tuyến PP tới scrubber).
4. **Số chụp treo**: DM nội thất ghi 2 cho Digestion và 2 cho Kjeldahl; bản vẽ 1 + 1. Máy chưng cất đạm tự động có cần chụp?
5. Vị trí thực tế arm hood ở Sample prep.
6. Hiện trạng ngoài tường A: FPS (có xả khí fumigant?), cửa D3, mặt bằng trống cho bãi CDU và scrubber, ranh giới khu đất, đường nội bộ.
7. Chiều cao mái/kèo; vị trí, cao độ ống gió Phase 1 đi qua khu Phase 2 → kiểm tra va chạm với AHU (cao ~2 m) và ống chính.
8. Cho phép đi ống trên trần khu CST (ngoài phạm vi) cho tuyến scrubber/DP-2.01; nếu không, phương án B: xả qua mái ngay trên EAF-2.01 / hood ST1 (cần thông tin mái).
9. Điểm đấu thoát nước ngưng thực tế (trục A hay trục E).
10. Diện tích ST2 theo kích thước 105,0 m² vs bảng 103,64 m² (cột/hộp kỹ thuật) – dùng 103,64 cho tải, 105,0 cho bố trí.
