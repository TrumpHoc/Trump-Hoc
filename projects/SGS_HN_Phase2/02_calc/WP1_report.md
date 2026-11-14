# WP1 – TÍNH TẢI, CÂN BẰNG GIÓ VÀ CHỌN THIẾT BỊ HVAC – PHASE 2 (v4, đã cập nhật theo quyết định chủ trì)

SGS Hải Phòng – cải tạo khu H&N Phase 2 (Nhà xưởng 3). Tổng thầu GGV. Ngày 14/11/2026.
Cơ sở bắt buộc: `01_design_basis/DESIGN_BASIS.md` (gọi tắt **DB**). Hình học: `03_geometry/geometry.json` + `WP2_report.md` (gọi tắt **WP2**). Phiên bản trước: v3 (`tools/build.py`, `HVAC_SGS_HN_Phase2_TinhTai_v3.xlsx`).

## 0. Quyết định của chủ trì (14/11/2026) – đã áp dụng

| # | Nội dung | Quyết định |
|---|---|---|
| 1 | Bản nháp WP1: gió bù FAFE = 90 % gió thải, 10 % còn lại lấy bằng gió chuyển từ phòng sạch (miệng gió chuyển trên vách), AHU bù bằng gió tươi đã xử lý (+80 kW) | **KHÔNG DUYỆT.** Làm theo IFC/Phase 1 (chỉ đạo CĐT: mọi thứ theo IFC như Phase 1, FAFE ≈ gió thải; danh mục thiết bị CĐT không yêu cầu chênh áp): **FAFE = 100 % gió thải từng phòng, không gió chuyển, bỏ miệng gió chuyển**; gió tươi AHU = nhu cầu người (kiểm tra Ev ASHRAE 62.1) + nhu cầu có lý do khác. Thay cho “90 %” ở DB §2. |
| 2 | Instrument room: bù 100 % gió thải (1.200 m³/h) bằng gió tươi đã xử lý qua AHU-2.03, không dùng FAFE | **DUYỆT** (thay cho DB §4.2 riêng phòng này). |
| 3 | 3 AHU (tách AHU-2.03 cho Instrument); chấp nhận ACH cao do tải hiện; T gió cấp 16,36 °C; áp dụng số liệu WP2 | **DUYỆT.** |
| 4 | Rủi ro gió bù tràn (spill) | Giữ là vấn đề mở; khuyến nghị miệng FAFE đặt tại mặt tủ. |

## 1. Sản phẩm (thư mục `02_calc/`)

| Tệp | Nội dung |
|---|---|
| `build_calc.py` | `python3 build_calc.py`: tạo `HVAC_Phase2_Calc_v4.xlsx` (mọi phép tính là công thức Excel), tính lại bằng LibreOffice, quét lỗi (#VALUE!, #REF!, #DIV/0!, #NAME?, #N/A), đọc KiemTra, xuất 2 tệp JSON. Độ nhạy: `OUT=/đường/dẫn.xlsx OVR='{"SPL":0.05}' python3 build_calc.py` (không ghi đè JSON). |
| `HVAC_Phase2_Calc_v4.xlsx` | Sheets: **CoSo** (đầu vào + nguồn, trạng thái không khí Magnus, bảng động cơ IEC), **ThietBi** (68 dòng nhiệt thiết bị), **TaiPhong** (tải phòng, cân bằng gió; bảng gió chuyển giữ lại nhưng “không áp dụng”, mọi giá trị 0), **AHU** (3 AHU + so sánh PA 2 AHU, CDU, sưởi, máy hút ẩm, bơm nước ngưng), **QuatHut_Bu** (FAF, EAF, quạt tủ hood, scrubber), **LichThietBi**, **KiemTra** (133 hạng mục). |
| `equipment_schedule.json` | 20 tag thiết bị: loại, SL, vị trí, lưu lượng (kèm phân bổ theo phòng/điểm cho FAF/EAF), ESP, lạnh, sưởi, điện, điện áp/pha, FLA, cổ ống đề xuất, ghi chú. |
| `airflow_by_room.json` | Từng phòng: cấp, hồi, rò ra (dư áp), tươi theo người, thải theo loại, FAFE, gió chuyển (= 0), AHU, FAF/EAF, điểm nguồn quạt tủ hood, scrubber, kiểm tra ẩm, tỷ lệ tràn chịu được; tổng. |

Kết quả sau khi LibreOffice tính lại: **0 lỗi công thức; KiemTra 133/133 ĐẠT**.

## 2. Phương pháp

1. **Trạng thái không khí**: Psat theo Magnus `0,61094·exp(17,625T/(T+243,04))` kPa; `W = 0,622·φ·Ps/(P − φ·Ps)`; `h = 1,006T + W(2501 + 1,86T)`.
2. **Tải phòng (hè)**: vách giáp không gian không điều hòa (xưởng, lối đi, khu ngoài phạm vi) ở 35 °C; vách giáp tường ngoài A (chỉ ST2, 15,6 m) ở 37,8 °C – chiều dài từ WP2; U vách 0,45; trần U 0,42, plenum 45 °C; cao 3,0 m; đèn 12 W/m²; người 75/55 W; thiết bị = Σ(W × (SL2027 + bổ sung) × hệ số) từ `equipment_cdt_classified.json` + 2 dòng giả định cho 2 phòng Standard; hệ số an toàn 10 % cho RSH và RLH.
3. **Gió cấp phòng** = MAX(tải hiện; tải ẩn; 6 ACH; thải không có FAFE), làm tròn lên 10 m³/h; ΔT = 25 − T gió cấp.
4. **T gió cấp** = 15 °C (sau coil) + ΔT quạt = ΔP_tổng/(ρ·cp·η) = 985/(1,2·1006·0,6) = 1,36 K → **16,36 °C** (đã duyệt).
5. **Cân bằng từng phòng**: Cấp + FAFE + Chuyển vào = Hồi + Thải + Chuyển ra + Rò ra. FAFE = 100 % gió thải (trừ Instrument = 0), gió chuyển = 0. SPL = 0 (gió bù không xử lý không tính vào tải phòng/AHU – như IFC/Phase 1).
6. **Gió tươi AHU** = MAX(Vot; Σ thải không có FAFE), Vot = Σ(25 m³/h/người)/Ev, Ev = 1 + Xs − Zp,max (ASHRAE 62.1 App. A). Gió tươi dư (= gió tươi − Σ thải không có FAFE) thoát qua khe cửa; trong bảng tính được **chia cho các phòng không phải phòng hóa theo tỷ lệ gió cấp** → phòng sạch hơi dương áp, phòng hóa trung tính (phòng hóa không nhận dư áp để tránh đẩy hơi hóa chất ra ngoài).
7. **Tải coil** = S·ρ·(h_hòa trộn − h_sau coil); hiện = S·ρ·cp·ΔT; nước ngưng = S·ρ·ΔW. Công suất chọn = tải coil × 1,10.
8. **Kiểm tra ẩm từng phòng**: `S·ρ·(W_phòng − W_sau coil)·hfg` ≥ RLH. Chỉ số “tỷ lệ tràn FAFE chịu được” = phần gió bù chưa xử lý được lọt vào vùng phòng mà RH vẫn ≤ 70 % (Instrument ≤ 55 %).
9. **Sưởi** = MAX(nhu cầu đông ×1,1; gia nhiệt lại tối thiểu 2 K × gió cấp), bậc 3 kW. Nhu cầu đông = gió tươi 9,8 → 18 °C + bao che (xưởng 12 °C, tường A 9,8 °C, plenum 10 °C) + 10 % FAFE tràn; không kể nhiệt thiết bị.
10. **CDU**: suy giảm theo chiều dài ống gas (WP2) −0,2 %/m sau 7,5 m; số module = ROUNDUP(chọn / hệ số / 50 kW).
11. **Quạt**: trục = Q·ΔP/(3600·η), η = 0,6; động cơ = cỡ IEC nhỏ nhất ≥ 1,15 × trục. FLA 3 pha = P/(√3·380·0,85·0,85); 1 pha = P/(220·0,85·0,85); sưởi = P/(√3·380); CDU = P/(√3·380·0,85).

## 3. Giả định và nguồn

| # | Giả định | Giá trị | Nguồn |
|---|---|---|---|
| 1 | Ngoài trời hè / đông | 37,8 °C – 62,8 % / 9,8 °C – 71 % | DB §2 (Phase 1) |
| 2 | Phòng | 25 °C; RH thiết kế 60 % (tối đa 70 %); Instrument 55 % | DB §2 |
| 3 | Sau coil / gió cấp | 15 °C – 97 % / 16,36 °C | DB §2; chủ trì duyệt |
| 4 | ≥ 6 ACH, lọc G4+F8; gió tươi 25 m³/h/người | – | DB §2 |
| 5 | Lưu lượng hút: hood 2.270, chụp 1.100, tủ HC 1.300, bồn rửa 300, arm 800, điểm thiết bị 150 m³/h | – | DB §3 |
| 6 | Điểm hút khối phải theo bản vẽ: ST2 = 6 hood (1 acid – ST2-FH3) + 2 tủ HC + 2 bồn rửa; Standard 19,72 = 4 hood + 1 bồn rửa; Standard 11,90 = 1 tủ HC | tổng 10/3/3 không đổi | WP2 (DW-A-03-06); chủ trì |
| 7 | Khối trái: Digestion 2 hood acid + 1 chụp; Kjeldahl 1 chụp; ST1 3 hood (1 acid) + 1 tủ HC; Sample prep 2 arm hood; Instrument 8 điểm | – | DB §3 |
| 8 | **FAFE = 100 % gió thải từng phòng** (Instrument 0 %), không gió chuyển; SPL = 0 | – | Chủ trì 14/11 (IFC/Phase 1) |
| 9 | Diện tích theo bảng DB (WP2 đo lệch ≤ 1,3 %) | 441,76 m² | DB §3; WP2 |
| 10 | Vách tiếp xúc không điều hòa / tường A | theo phòng | WP2 `rooms[].exposure` |
| 11 | U 0,45/0,42; xưởng 35 °C; plenum 45 °C; cao 3,0 m; đèn 12 W/m²; người 75/55 W; SF 10 % | – | v3 |
| 12 | Thiết bị SL2027 + bổ sung, hệ số theo JSON; Standard 11,90: 3 × 485 W × 0,5; Standard 19,72: 1 kW × 0,5 | 88,4 kW | JSON CĐT; chỉ đạo WP1 |
| 13 | Tổn thất trong AHU: G4 100 + F8 185 + coil 120 + sưởi 20 + vỏ 60 = 485 Pa | – | NCC điển hình |
| 14 | ESP AHU = MAX(500; ước tính × 1,25); tuyến cấp/hồi/tươi 25/22/26,7; 20/18/4,7; 6/5/21,1 m | 500 Pa | WP2; ma sát 1 Pa/m (DB §6); L_td = 2L; đầu cuối 130 Pa |
| 15 | FAF: ESP 400 Pa (IFC) + hộp lọc G4+F8 285 Pa = 685 Pa; EAF: 400 Pa | – | DB §1 (IFC); NCC |
| 16 | η quạt 0,6; dự phòng động cơ 1,15 | – | Chỉ đạo WP1 |
| 17 | CDU: COP 3,0 @37,8 °C; ≤ 50 kW danh định/module; suy giảm −0,2 %/m sau 7,5 m; ống gas 42,9 / 19,5 / 34,4 m | – | Chỉ đạo WP1; catalogue điển hình; WP2 |
| 18 | Scrubber: +10 %; v tháp 1,8 m/s; L/G 2,5 L/m³; bơm 20 m, η 0,55; ΔP quạt 600 + 400 Pa | – | Chỉ đạo WP1; thực hành chung |
| 19 | Máy hút ẩm: nồm 22 °C/95 %; thấm 1 ACH; 10 h có người; dự phòng 1,3; năng suất @25/55 = 0,45 × định mức @30/80; 2,5 L/kWh | – | Giả định WP1 |
| 20 | Nước ngưng: độ rơi khả dụng 0,15 m, dốc 1 % → 15 m tự chảy; bơm khi > 12 m; tuyến tới CD-2.01: 36,6 / 13,1 / 29,4 m | – | Giả định WP1; WP2 |
| 21 | Sưởi: gia nhiệt lại tối thiểu 2 K; 10 % FAFE tràn khi tính đông | – | Giả định WP1 |

## 4. Quyết định thiết kế (đã duyệt)

### 4.1 Gió bù = 100 % gió thải, không gió chuyển
Mỗi phòng có thải được bù 100 % bằng FAFE không xử lý, lọc G4+F8, cấp ngay trước tủ/chụp (IFC/Phase 1). Không có miệng gió chuyển trên vách. Các phòng hóa **trung tính** (Cấp AHU = Hồi); gió tươi người dư được bố trí thoát ở các phòng sạch (Balance, Sample prep, Coding, Dividing, Reporting, Office) nên các phòng này hơi dương áp so với lối đi/xưởng – không có yêu cầu đo chênh áp (CĐT). Gió tươi qua AHU chỉ còn theo người (ASHRAE 62.1 Ev): **AHU-2.01 1.230, AHU-2.02 650 m³/h**; AHU-2.03 1.200 m³/h (bù thải Instrument). Σ 3.080 m³/h (bản nháp 90 %: 5.600).

### 4.2 Instrument room – bù thải bằng gió tươi đã xử lý (AHU-2.03)
8 điểm hút thiết bị không phải tủ/chụp, gió bù không xử lý sẽ hòa vào phòng: 1.080–1.200 m³/h ở 26,3 g/kg mang ≈ 14 kW ẩn, gấp ~2,7 lần khả năng khử ẩm của gió cấp ở 55 % (≈ 5,1 kW). AHU-2.03 lấy 1.200 m³/h gió tươi = lượng thải; chi phí ≈ 20,4 kW lạnh.

### 4.3 3 AHU (PA B)

| Tiêu chí | PA A: 2 AHU (Instrument trong AHU-2.01) | PA B: 3 AHU (chọn) |
|---|---|---|
| Gió cấp AHU-2.01 | 34.610 m³/h (Instrument 36 %) | 22.180 + 12.430 (AHU-2.03) |
| Tải coil AHU-2.01 | 195,2 kW (Instrument 36 %) | 125,8 + 69,4 kW |
| Tổng gió / tổng tải | 39.860 m³/h / 230,9 kW (cộng tuyến tính; gió tươi khối trái vẫn ≈ Vot các phòng khác + 1.200 bù thải Instrument nên không tiết kiệm đáng kể) | 39.860 m³/h / 230,9 kW |
| Điều khiển T/RH | 1 vùng cho Instrument (≤ 55 %, 24/7) và 8 phòng (≤ 70 %, giờ hành chính) – không giữ được đồng thời | Instrument điều khiển riêng, 24/7 |
| Dự phòng | Hỏng AHU-2.01 dừng thiết bị phân tích | Tách rủi ro |
| Bố trí | 1 AHU ~35.000 m³/h trên trần Balance | AHU-2.03 trên trần Instrument (WP2), ống cấp ≈ 4 m |

### 4.4 Bội số trao đổi cao
Kjeldahl 63,4; ST1 60,2; Instrument 115,7 ACH – do **tải hiện** (0,47–0,87 kW/m² thiết bị); ΔT cấp đã tối đa (15 °C sau coil → 16,36 °C cấp → ΔT 8,64 K). Bố trí miệng gió (WP3): Instrument ~12 miệng khuếch tán xoáy ~1.000 m³/h; ST1/Kjeldahl miệng đục lỗ vận tốc thấp, xa mặt tủ (gió ngang tại mặt tủ ≤ 0,2 m/s, EN 14175).

### 4.5 Kiểm tra ẩm và gió bù tràn (vấn đề mở)
Mọi phòng đạt kiểm tra ẩm với SPL = 0. Tỷ lệ tràn FAFE chịu được (giữ RH ≤ 70 %): **Standard 19,72: 1,3 %; ST2: 4,7 %; Digestion: 6,2 %; Standard 11,90: 8,9 %**; ST1 41,5 %; Sample prep 47,8 %; Kjeldahl 93 %. Độ nhạy: SPL 5 % → tải coil 295 kW và 6 hạng mục KiemTra không đạt; SPL 10 % → 380 kW, 7 hạng mục không đạt. **Khuyến nghị**: miệng FAFE đặt tại mặt tủ (slot/đục lỗ, ≤ 0,3 m/s, hướng vào vùng hút; ưu tiên tủ có cửa cấp gió phụ kiểu auxiliary-air nếu gói nội thất cho phép); van gió động cơ tại mỗi miệng bù mở theo tín hiệu chạy quạt tủ/chụp; quạt FAF biến tần giữ áp ống.

## 5. Kết quả theo phòng

| Phòng | AHU | RSH (W) | RLH (W) | Cấp | ACH | Hồi | Rò ra | Thải | FAFE | Tràn FAFE chịu được | FAF | EAF |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Digestion room | AHU-2.01 | 3.625 | 121 | 1.260 | 28,8 | 1.260 | 0 | 5.640 | 5.640 | 6,2 % | FAF-2.01 | EAF-2.01 |
| Kjeldahl room | AHU-2.01 | 10.390 | 121 | 3.590 | 63,4 | 3.590 | 0 | 1.100 | 1.100 | 93,1 % | FAF-2.01 | EAF-2.01 |
| Sample treatment 1 | AHU-2.01 | 34.370 | 605 | 11.870 | 60,2 | 11.870 | 0 | 8.110 | 8.110 | 41,5 % | FAF-2.01 | EAF-2.01 |
| Balance room | AHU-2.01 | 1.389 | 242 | 480 | 8,7 | 372 | 108 | 0 | 0 | – | – | – |
| Sample preparation | AHU-2.01 | 8.234 | 605 | 2.850 | 16,1 | 2.208 | 642 | 1.600 | 1.600 | 47,8 % | FAF-2.01 | EAF-2.03 |
| Sample coding | AHU-2.01 | 1.332 | 121 | 460 | 6,1 | 356 | 104 | 0 | 0 | – | – | – |
| Sample dividing | AHU-2.01 | 771 | 60 | 270 | 9,0 | 209 | 61 | 0 | 0 | – | – | – |
| Reporting room | AHU-2.01 | 4.038 | 605 | 1.400 | 25,9 | 1.085 | 315 | 0 | 0 | – | – | – |
| Instrument room | AHU-2.03 | 36.000 | 605 | 12.430 | 115,7 | 11.230 | 0 | 1.200 | 0 | – | – | EAF-2.03 |
| Sample treatment 2 | AHU-2.02 | 8.520 | 605 | 2.950 | 9,5 | 2.950 | 0 | 16.820 | 16.820 | 4,7 % | FAF-2.02 | EAF-2.02 |
| Standard (19,72 m²) | AHU-2.02 | 1.302 | 121 | 450 | 7,6 | 450 | 0 | 9.380 | 9.380 | 1,3 % | FAF-2.02 | EAF-2.02 |
| Standard (11,90 m²) | AHU-2.02 | 1.258 | 121 | 440 | 12,3 | 440 | 0 | 1.300 | 1.300 | 8,9 % | FAF-2.02 | EAF-2.02 |
| Office | AHU-2.02 | 4.065 | 605 | 1.410 | 11,5 | 760 | 650 | 0 | 0 | – | – | – |
| **Tổng** | | **115.293** | **4.538** | **39.860** | 30,1 | **36.780** | **1.880** | **45.150** | **43.950** | | | |

(m³/h trừ khi ghi khác.) Nhiệt thiết bị Σ 88,4 kW (Instrument 31,2; ST1 28,9; Kjeldahl 8,8; Sample prep 5,4; ST2 4,4). T phòng dự kiến 24,8–25,0 °C mọi phòng.
Gió thải: hood 34.050 (acid 9.080 → scrubber; hữu cơ 24.970 xả thẳng SS304), chụp 2.200, tủ HC 5.200, bồn rửa 900, arm hood 1.600, điểm thiết bị 1.200 → **Σ 45.150 m³/h**; qua EAF của HVAC 11.100 m³/h. FAFE Σ 43.950 m³/h (= thải − 1.200 Instrument).

## 6. AHU và dàn nóng

| Hạng mục | AHU-2.01 | AHU-2.02 | AHU-2.03 | Tổng |
|---|---|---|---|---|
| Vị trí (WP2) | trên trần Balance | trên trần ST2 phía ĐN | trên trần Instrument | |
| Phòng phục vụ | Digestion, Kjeldahl, ST1, Balance, Sample prep, Coding, Dividing, Reporting | ST2, Standard 19,72, Standard 11,90, Office | Instrument | |
| Gió cấp (m³/h) | 22.180 | 5.250 | 12.430 | 39.860 |
| Gió tươi (m³/h) / tỷ lệ | 1.230 / 5,5 % | 650 / 12,4 % | 1.200 / 9,7 % | 3.080 |
| Vou / Ev / Vot (m³/h) | 1.025 / 0,84 / 1.223 | 600 / 0,94 / 640 | 250 / 1,0 / 250 | |
| Gió hồi về AHU (m³/h) | 20.950 | 4.600 | 11.230 | 36.780 |
| Vào coil (T / RH) | 25,71 °C / 61,3 % | 26,59 °C / 62,7 % | 26,24 °C / 58,0 % | |
| Ra coil | 15 °C / 97 % | 15 °C / 97 % | 15 °C / 97 % | |
| Tải coil tổng / hiện / ẩn (kW) | 125,8 / 79,7 / 46,1 | 35,8 / 20,4 / 15,4 | 69,4 / 46,8 / 22,6 | **230,9** |
| SHR | 0,63 | 0,57 | 0,67 | |
| **Công suất lạnh chọn (+10 %)** | **138,3 kW** (39,3 TR) | **39,3 kW** (11,2 TR) | **76,3 kW** (21,7 TR) | **254,0 kW** |
| Công suất hiện chọn (kW) | 87,6 | 22,4 | 51,5 | 161,6 |
| Nước ngưng (kg/h) | 63,0 | 21,1 | 30,6 | 114,7 |
| ESP thiết kế / ước tính tuyến (Pa) | 500 / 233 | 500 / 206 | 500 / 184 | |
| Áp tổng quạt (Pa) / ΔT quạt | 985 / 1,36 K | 985 / 1,36 K | 985 / 1,36 K | |
| Trục / động cơ quạt (kW) | 10,1 / 15 | 2,4 / 3 | 5,7 / 7,5 | 25,5 |
| Sưởi điện chọn (kW) | 15 | 12 | 9 | 36 |
| Nhu cầu đông / gia nhiệt 2 K (kW) | 10,2 / 14,9 | 11,5 / 3,5 | 3,8 / 8,3 | |
| Mùa đông: T sau sưởi ≥ T cấp yêu cầu | 19,6 ≥ 18,9 °C | 23,8 ≥ 23,4 °C | 19,4 ≥ 18,0 °C | |
| Lọc / mặt coil @2,5 m/s | G4+F8 / 2,46 m² | G4+F8 / 0,58 m² | G4+F8 / 1,38 m² | |
| Điện AHU (quạt + sưởi) kW / FLA (A) | 30 / 54,3 | 15 / 24,5 | 16,5 / 29,4 | 61,5 |
| Cổ ống cấp / hồi / tươi (mm) | 1300×600 / 1350×550 / 350×200 | 450×400 / 400×400 / 250×150 | 800×550 / 650×600 / 350×200 | |
| Ống gas tới bãi CDU / hệ số suy giảm | 42,9 m / 0,929 | 19,5 m / 0,976 | 34,4 m / 0,946 | |
| **CDU** module × danh định @37,8 °C | **CDU-2.01: 3 × 49,6 kW** | **CDU-2.02: 1 × 40,3 kW** | **CDU-2.03: 2 × 40,3 kW** | 6 module |
| CDU thực sau suy giảm / module | 46,1 kW | 39,3 kW | 38,2 kW | |
| CDU điện/module (COP 3) kW / FLA | 16,5 / 29,6 | 13,4 / 24,0 | 13,4 / 24,0 | 90,0 kW |
| Nước ngưng | bơm CP-2.02 | bơm CP-2.03 | bơm CP-2.01 (chung DH-2.01) | |

Ghi chú: (1) CDU-2.01 cần 49,6 kW danh định/module, sát giới hạn 50 kW – nếu model hãng không đạt thì dùng 4 module (≈ 37 kW). (2) CDU-2.02 chỉ 1 module (không dự phòng) – có thể tách 2 × 20 kW nếu CĐT cần dự phòng. (3) Mỗi module CDU một mạch coil riêng, inverter, R410A/R32.

## 7. Quạt cấp bù FAF, quạt hút EAF, quạt tủ hood, scrubber

Tag theo đề xuất WP2: 2 FAF (theo khối) + 3 EAF. Lý do gộp: phòng hóa mỗi khối liền nhau, chung louver và tuyến; quạt tủ hood bật/tắt riêng nên dù quạt theo phòng hay theo nhóm đều cần van gió động cơ tại mỗi miệng bù → gộp nhóm + biến tần giữ áp ống là đơn giản và ít thiết bị nhất.

**FAF** (gió tươi không xử lý, lọc G4+F8, 685 Pa = ESP 400 + lọc 285, biến tần, van gió động cơ tại mỗi miệng bù):

| Tag | Vị trí (WP2) | Phục vụ (m³/h) | Σ m³/h | Động cơ | Điện / FLA | Cổ ống |
|---|---|---|---|---|---|---|
| FAF-2.01 | trên trần ST1 sát lối đi; tươi từ LV-2.01 (≈ 32 m) | Digestion 5.640, Kjeldahl 1.100, ST1 8.110, Sample prep 1.600 | **16.450** | 7,5 kW | 380V/3P / 15,8 A | 1150×500 |
| FAF-2.02 | trên trần ST2 gần AHU-2.02; tươi từ LV-2.02 | ST2 16.820, Standard 19,72 9.380, Standard 11,90 1.300 | **27.500** | 11 kW | 380V/3P / 23,1 A | 1950×500 |
| **Σ** | | | **43.950** | 18,5 kW | | |

**EAF** (ESP 400 Pa, trên trần, xả qua tường A + weather cover + lưới côn trùng, biến tần):

| Tag | Vị trí → điểm xả | Phục vụ | m³/h | Vật liệu | Động cơ | Cổ ống |
|---|---|---|---|---|---|---|
| EAF-2.01 | trên trần Kjeldahl → DP-2.01 | chụp Digestion 1.100 + Kjeldahl 1.100, tủ HC ST1 1.300 | 3.500 | SS304 | 0,75 kW 380V/3P | Ø400 |
| EAF-2.02 | trên trần ST2 phía tường A → DP-2.02 | 2 tủ HC + 2 bồn rửa ST2 3.200, bồn rửa Standard 19,72 300, tủ HC Standard 11,90 1.300 | 4.800 | SS304 | 1,1 kW 380V/3P | Ø450 |
| EAF-2.03 | trên vách Sample prep/Instrument → DP-2.03 (≈ 32 m) | 2 arm hood 1.600 + 8 điểm Instrument 1.200 | 2.800 | SS304 (nhánh arm hood GI) | 0,75 kW 380V/3P | Ø350 |
| **Σ** | | | **11.100** | | 2,6 kW | |

**Quạt tủ hood (gói nội thất – HVAC chỉ cấp nguồn + điều khiển)**: 15 điểm × 0,6 kW, 380V/3P, ≈ 1,3 A: FHF-2.01 Digestion 2 (acid), FHF-2.02 ST1 3 (1 acid), FHF-2.03 ST2 6 (1 acid: ST2-FH3), FHF-2.04 Standard 19,72 4. Σ 9,0 kW. Hood hữu cơ (11) xả thẳng ống SS304 qua DP-2.01/DP-2.02; 4 hood acid về scrubber bằng ống PP.

**SCR-2.01** – scrubber ướt tháp đệm PP, ngoài tường A phía bắc cửa CST: hood acid 9.080 → **10.000 m³/h**; Ø tháp **1.450 mm**; tuần hoàn 25 m³/h; bơm 3 kW; quạt FRP/PP 1.000 Pa **5,5 kW** (biến tần giữ áp âm ống gom); định lượng NaOH + pH 0,18 kW; tổng **8,7 kW, 380V/3P, ≈ 18 A**; cổ vào Ø650 PP. Quạt tủ acid nối tiếp quạt scrubber. Nước cấp + đấu XLNT do gói khác.

## 8. Máy hút ẩm và nước ngưng

**DH-2.01** (Instrument, trên trần cạnh AHU-2.03): yêu cầu 30,1 L/ngày @25 °C/55 % (người 7,9 + thấm cửa khi nồm 15,2 kg/ngày, ×1,3) ≈ 66,8 L/ngày @30 °C/80 % → chọn **70 L/ngày**, ≈ 1,17 kW, 220V/1P, âm trần nối ống, bơm ngưng tích hợp, chạy khi RH > 55 %.

**Bơm nước ngưng** (độ rơi khả dụng 0,15 m, dốc 1 % → 15 m; bơm khi > 12 m; điểm thoát CD-2.01 ngoài tường A):

| Tag | Phục vụ | Tuyến (WP2) | Lưu lượng bơm | Điện |
|---|---|---|---|---|
| CP-2.01 | AHU-2.03 + DH-2.01 | 29,4 m | ≥ 96 L/h, 6 m | 0,1 kW 220V/1P |
| CP-2.02 | AHU-2.01 | 36,6 m | ≥ 189 L/h, 6 m | 0,1 kW 220V/1P |
| CP-2.03 | AHU-2.02 | 13,1 m | ≥ 63 L/h, 6 m | 0,1 kW 220V/1P |

Bơm có bình chứa, phao mức cao cắt cấp lạnh AHU. Nếu bệ AHU (phần xây dựng) cho độ rơi ≥ 0,45 m thì CP-2.02/CP-2.03 có thể bỏ.

## 9. Lịch thiết bị – tổng hợp điện (đầu vào WP4/WP5)

| Nhóm | Tag | SL | Tổng điện (kW) |
|---|---|---|---|
| AHU (quạt + sưởi) | AHU-2.01 / 2.02 / 2.03 | 3 | 61,5 |
| Dàn nóng | CDU-2.01 (3), CDU-2.02 (1), CDU-2.03 (2) | 6 | 90,0 |
| Quạt cấp bù | FAF-2.01, FAF-2.02 | 2 | 18,5 |
| Quạt hút | EAF-2.01, 2.02, 2.03 | 3 | 2,6 |
| Điểm nguồn quạt tủ hood | FHF-2.01…2.04 | 15 điểm | 9,0 |
| Scrubber | SCR-2.01 | 1 | 8,7 |
| Máy hút ẩm | DH-2.01 | 1 | 1,2 |
| Bơm nước ngưng | CP-2.01, 2.02, 2.03 | 3 | 0,3 |
| **Tổng điện đặt HVAC** | | | **191,7 kW** |

Chi tiết FLA/điện áp/cổ ống trong LichThietBi và `equipment_schedule.json`. Sưởi và lạnh cùng AHU chỉ chạy đồng thời ở chế độ khử ẩm – WP4 quyết định hệ số đồng thời.

## 10. Độ nhạy (chạy lại toàn bảng bằng `OVR`)

| Kịch bản | Gió tươi AHU (m³/h) | Tải coil (kW) | KiemTra |
|---|---|---|---|
| **Cơ sở (đã duyệt)**: FAFE 100 %, SPL 0, ESP 500 | 3.080 | **230,9** | đạt |
| SPL 5 % (gió bù tràn vào phòng) | 3.140 | 295,4 | 6 lỗi |
| SPL 10 % | 3.160 | 379,8 | 7 lỗi |
| ESP tối thiểu 800 Pa (như Phase 1) | 3.050 | 239,4 | đạt; quạt 18,5 / 4 / 11 kW |
| Chỉ thiết bị năm 2027 | 3.060 | 174,8 | đạt |
| (Tham khảo – không duyệt) FAFE 90 % + gió chuyển | 5.600 | 273,1 | đạt |

So với v3 (202,6 kW): +28,3 kW, do bù thải Instrument bằng gió tươi xử lý (+20,4 kW), T cấp 16,36 °C (gió cấp +4 %), tường ngoài A của ST2, Ev ASHRAE 62.1 và tải nhỏ khác.

## 11. Danh sách tự kiểm tra (KiemTra – 133/133 ĐẠT)

1. Cân bằng khối lượng từng phòng (13): |Cấp + FAFE + Chuyển vào − Hồi − Thải − Chuyển ra − Rò ra| < 0,5 m³/h.
2. Hồi ≥ 0 (13); ACH ≥ 6 (13); ACH ≤ 40 hoặc do tải hiện (13); khử ẩm đủ RLH (13); T phòng dự kiến 23–27 °C (13).
3. Σ gió chuyển vào = Σ ra (= 0); gió chuyển chỉ trong cùng AHU (không có).
4. Σ diện tích = 441,76 m²; Σ người = 75; 15 hood / 4 acid; Σ hood acid = 9.080; Σ nhiệt thiết bị phòng = Σ ThietBi.
5. Σ thải = Σ EAF + Σ hood; Σ FAF = Σ FAFE phòng; scrubber ≥ 1,1 × hood acid.
6. Mỗi AHU (×3): Cấp = Hồi + Tươi; Tươi ≥ Vot; dư áp ≥ 0; Σ hồi phòng = hồi AHU và Σ rò ra phòng = dư áp; coil khử ẩm, RH vào ≤ 100 %; tải coil ≥ Σ tải phòng; Σ CDU thực ≥ công suất chọn; module danh định ≤ 50 kW; sưởi đông đủ; động cơ ≥ 1,15 × trục; ESP thiết kế ≥ ước tính; T cấp ở TaiPhong = T sau coil + ΔT quạt.
7. Tổng PA B = PA A (gió); Σ gió cấp phòng = Σ AHU; Instrument RH ≤ 55 %; máy hút ẩm ≥ yêu cầu; có bơm nơi cần; động cơ FAF/EAF/scrubber ≥ 1,15 × trục; lịch thiết bị khớp AHU.
8. Quét lỗi sau khi LibreOffice tính lại: 0 lỗi; không ô văn bản nào bắt đầu bằng “=”.

## 12. Vấn đề mở còn lại

1. **Gió bù tràn (spill)**: Standard 19,72 chịu 1,3 %, ST2 4,7 %, Digestion 6,2 %, Standard 11,90 8,9 % – miệng FAFE phải đặt tại mặt tủ (§4.5); Standard 19,72 (4 hood trong 19,7 m², ≈ 159 ACH thải) nên xem lại bố trí với CĐT/nội thất.
2. **ST1 tủ hóa chất thứ 2** (ST1-CC2 có trên bản vẽ, danh mục nội thất 5.12 chỉ 1 tủ có hút): nếu có hút → EAF-2.01 4.800 m³/h, FAF-2.01 17.750 m³/h.
3. **Nhiệt tỏa thiết bị Instrument / ST1 / Kjeldahl** quyết định ACH 60–116 và cỡ AHU-2.03: cần “site preparation” của CĐT và hệ số đồng thời bếp/bể ST1.
4. **Số chụp treo**: DB 1 chụp/phòng (Digestion, Kjeldahl); danh mục nội thất 6.5/7.4 ghi 2 chụp/phòng (+1.100 m³/h thải và FAFE mỗi chụp).
5. **Số quạt tủ hút trong danh mục CĐT** (27) ≠ 15 tủ hood theo DB – HVAC tính 15 điểm nguồn.
6. **Bố trí miệng gió cấp ở phòng ACH cao** gần tủ hood – WP3.
7. **Cao độ bệ AHU** quyết định có bỏ được CP-2.02/CP-2.03 không.
8. **CDU-2.01** sát giới hạn module (49,6/50 kW); **CDU-2.02** 1 module không dự phòng; suy giảm ống gas −0,2 %/m là giả định – kiểm tra catalogue hãng (chiều dài, chênh cao, hồi dầu).
9. **Mùa đông**: FAFE không xử lý cấp 9,8 °C tại mặt tủ – sưởi AHU đã tính 10 % tràn; nếu người dùng phàn nàn cần sưởi FAFE (ngoài DB).
10. **Máy hút ẩm**: điều kiện nồm và tỷ lệ năng suất 0,45 là giả định; chọn model theo đặc tính NCC tại 25 °C/55 %.
11. FLA là ước tính (PF 0,85, η 0,85) – thay bằng số liệu NCC khi chọn model.
12. DB §2 (“FAFE = 90 %”) và §4.2 (FAFE cho Instrument) cần được cập nhật theo quyết định chủ trì 14/11 (ngoài phạm vi WP1 sửa).
