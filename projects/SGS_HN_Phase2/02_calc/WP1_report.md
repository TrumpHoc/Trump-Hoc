# WP1 – TÍNH TẢI, CÂN BẰNG GIÓ VÀ CHỌN THIẾT BỊ HVAC – PHASE 2 (v4)

SGS Hải Phòng – cải tạo khu H&N Phase 2 (Nhà xưởng 3). Tổng thầu GGV. Ngày 14/11/2026.
Cơ sở bắt buộc: `01_design_basis/DESIGN_BASIS.md` (gọi tắt **DB**). Hình học: `03_geometry/geometry.json` + `WP2_report.md` (gọi tắt **WP2**). Phiên bản trước: v3 (`tools/build.py`, `HVAC_SGS_HN_Phase2_TinhTai_v3.xlsx`).

## 1. Sản phẩm (thư mục `02_calc/`)

| Tệp | Nội dung |
|---|---|
| `build_calc.py` | Chạy `python3 build_calc.py`: tạo `HVAC_Phase2_Calc_v4.xlsx` (mọi phép tính là công thức Excel), tính lại bằng LibreOffice (`soffice --headless --calc --convert-to xlsx`), quét lỗi (#VALUE!, #REF!, #DIV/0!, #NAME?, #N/A), đọc kết quả sheet KiemTra, xuất 2 tệp JSON. Chạy độ nhạy: `OUT=/đường/dẫn.xlsx OVR='{"fm":1.0}' python3 build_calc.py` (không ghi đè JSON). |
| `HVAC_Phase2_Calc_v4.xlsx` | Sheets: **CoSo** (đầu vào + nguồn, trạng thái không khí Magnus, bảng động cơ IEC), **ThietBi** (68 dòng nhiệt thiết bị), **TaiPhong** (tải phòng, cân bằng gió, bảng gió chuyển), **AHU** (3 AHU + so sánh phương án 2 AHU, CDU, sưởi, máy hút ẩm, bơm nước ngưng), **QuatHut_Bu** (FAF, EAF, quạt tủ hood, scrubber), **LichThietBi** (lịch thiết bị), **KiemTra** (133 hạng mục tự kiểm tra). |
| `equipment_schedule.json` | 20 dòng thiết bị (tag): tag, loại, SL, vị trí, lưu lượng (kèm phân bổ theo phòng/điểm cho FAF/EAF), ESP, lạnh, sưởi, điện, điện áp/pha, FLA, cổ ống đề xuất, ghi chú. |
| `airflow_by_room.json` | Từng phòng: cấp, hồi, tươi theo người, thải theo loại, FAFE, gió chuyển vào/ra (kèm nguồn/đích), AHU, FAF/EAF, điểm nguồn quạt tủ hood, scrubber, kiểm tra ẩm; danh sách gió chuyển; tổng. |

Kết quả sau khi LibreOffice tính lại: **0 lỗi công thức; KiemTra 133/133 ĐẠT**.

## 2. Phương pháp

1. **Trạng thái không khí**: Psat theo Magnus `0,61094·exp(17,625T/(T+243,04))` kPa; `W = 0,622·φ·Ps/(P − φ·Ps)`; `h = 1,006T + W(2501 + 1,86T)` (như v3).
2. **Tải phòng (hè)**: vách giáp không gian không điều hòa (xưởng, lối đi giữa 2 khối, khu ngoài phạm vi) ở 35 °C, vách giáp tường ngoài trục A (chỉ ST2, 15,6 m) ở 37,8 °C – chiều dài lấy từ WP2 `rooms[].exposure`; U vách 0,45; trần U 0,42 với plenum 45 °C; cao 3,0 m; đèn 12 W/m²; người 75 W hiện / 55 W ẩn; thiết bị = Σ(W × SL tính × hệ số tỏa nhiệt) từ `equipment_cdt_classified.json`, SL = SL2027 + bổ sung, cộng 2 dòng giả định cho 2 phòng Standard. Hệ số an toàn 10 % cho RSH và RLH.
3. **Gió cấp phòng** = MAX(tải hiện; tải ẩn; 6 ACH; Cấp − Hồi cần cho cân bằng), làm tròn lên 10 m³/h. Tải hiện dùng ΔT = 25 − T gió cấp của AHU phục vụ.
4. **T gió cấp** = 15 °C (sau coil) + ΔT nhiệt quạt (AHU hút qua coil, động cơ trong dòng khí): ΔT = ΔP_tổng/(ρ·cp·η) = 985/(1,2·1006·0,6) = **1,36 K → 16,36 °C**. DB ghi 16 °C (ứng với ΔT quạt 1 K); v4 dùng giá trị tính để nhất quán với áp quạt và η = 0,6 (gió cấp tăng ≈ 4 %).
5. **Cân bằng khối lượng từng phòng**: Cấp AHU + Bù FAFE + Chuyển vào = Hồi + Thải + Chuyển ra. FAFE = 90 % gió thải (DB §2), không xử lý, cấp ngay trước tủ/chụp, SPL = 0 (không tính vào tải phòng/AHU – như IFC/Phase 1).
6. **Gió tươi AHU** = MAX(Vot; Σ(Cấp − Hồi) các phòng). Vot = Σ(25 m³/h/người)/Ev, Ev = 1 + Xs − Zp,max (ASHRAE 62.1 App. A) để mọi phòng trong hệ nhiều phòng thực nhận đủ 25 m³/h/người.
7. **Tải coil** = S·ρ·(h_hòa trộn − h_sau coil); hiện = S·ρ·cp·(T_hòa trộn − 15); nước ngưng = S·ρ·(W_hòa trộn − W_sau coil). Công suất chọn = tải coil × 1,10.
8. **Kiểm tra ẩm từng phòng**: khả năng khử ẩm của gió cấp `S·ρ·(W_phòng − W_sau coil)·hfg` ≥ RLH (kể cả ẩn mang theo gió chuyển). Thêm chỉ số “tỷ lệ tràn FAFE chịu được” = phần gió bù chưa xử lý có thể lọt vào vùng phòng mà RH vẫn ≤ RH tối đa.
9. **Sưởi** = MAX(nhu cầu đông ×1,1; gia nhiệt lại tối thiểu 2 K × gió cấp), làm tròn bậc 3 kW. Nhu cầu đông = sưởi gió tươi 9,8 → 18 °C + tổn thất bao che (xưởng 12 °C, tường A 9,8 °C, plenum 10 °C) + 10 % FAFE tràn vào phòng; không kể nhiệt thiết bị.
10. **Dàn nóng**: công suất thực cần = công suất chọn AHU; suy giảm theo chiều dài ống gas (WP2) −0,2 %/m sau 7,5 m; số module = ROUNDUP(chọn/hệ số/50 kW).
11. **Quạt**: trục = Q·ΔP/(3600·η), η = 0,6; động cơ = cỡ IEC nhỏ nhất ≥ 1,15 × trục. FLA 3 pha = P/(√3·380·0,85·0,85); 1 pha = P/(220·0,85·0,85); sưởi = P/(√3·380); CDU = P/(√3·380·0,85).

## 3. Giả định và nguồn

| # | Giả định | Giá trị | Nguồn |
|---|---|---|---|
| 1 | Ngoài trời hè / đông | 37,8 °C – 62,8 % / 9,8 °C – 71 % | DB §2 (Phase 1 F3-SD-AC-004) |
| 2 | Phòng | 25 °C; RH thiết kế 60 % (tối đa 70 %); Instrument 55 % | DB §2 |
| 3 | Sau coil / gió cấp | 15 °C – 97 % / 15 + ΔT quạt = 16,36 °C | DB §2 + tính (mục 2.4) |
| 4 | Tối thiểu 6 ACH, lọc G4+F8; gió tươi 25 m³/h/người | – | DB §2 |
| 5 | Lưu lượng hút: hood 2.270, chụp 1.100, tủ HC 1.300, bồn rửa 300, arm 800, điểm thiết bị 150 m³/h | – | DB §3 |
| 6 | Phân bổ điểm hút khối phải theo bản vẽ: ST2 = 6 hood (1 acid – ST2-FH3) + 2 tủ HC + 2 bồn rửa; Standard 19,72 = 4 hood + 1 bồn rửa; Standard 11,90 = 1 tủ HC; tổng khối phải không đổi (10/3/3) | – | WP2 (DW-A-03-06), chỉ đạo điều phối |
| 7 | Khối trái: Digestion 2 hood acid + 1 chụp; Kjeldahl 1 chụp; ST1 3 hood (1 acid) + 1 tủ HC; Sample prep 2 arm hood; Instrument 8 điểm | – | DB §3 |
| 8 | FAFE = 90 % gió thải, không xử lý; SPL = 0 | – | DB §2, §4.2; v3 |
| 9 | Diện tích phòng theo bảng DB (WP2 đo lệch ≤ 1,3 %) | 441,76 m² | DB §3; WP2 |
| 10 | Vách tiếp xúc không điều hòa / tường A theo WP2 | xem TaiPhong | WP2 `rooms[].exposure` |
| 11 | U 0,45/0,42; xưởng 35 °C; plenum 45 °C; cao 3,0 m; đèn 12 W/m²; người 75/55 W; SF 10 % | – | v3 (`tools/build.py`) |
| 12 | Thiết bị: SL2027 + bổ sung, hệ số theo JSON; Standard 11,90: 3 tủ lạnh 485 W × 0,5; Standard 19,72: 1 kW × 0,5 | 88,4 kW | `equipment_cdt_classified.json`; chỉ đạo WP1 |
| 13 | Instrument: thải 1.200 m³/h bù 100 % bằng gió tươi đã xử lý qua AHU-2.03 (không FAFE) | fmi = 0 | **Quyết định WP1** – §4.2 |
| 14 | Gió chuyển Office → ST2 | 1.000 m³/h | Quyết định WP1 – §4.1 |
| 15 | Tổn thất trong AHU: G4 100 + F8 185 + coil ướt 120 + sưởi 20 + vỏ 60 = 485 Pa | – | NCC điển hình |
| 16 | ESP AHU = MAX(500; ước tính tuyến × 1,25); tuyến cấp/hồi/tươi theo WP2: 25/22/26,7; 20/18/4,7; 6/5/21,1 m | 500 Pa | WP2 `runs_for_takeoff`; ma sát 1 Pa/m (DB §6); L_td = 2L; đầu cuối 130 Pa |
| 17 | FAF: ESP 400 Pa (IFC) + hộp lọc G4+F8 285 Pa = 685 Pa; EAF: 400 Pa | – | DB §1 (IFC); NCC |
| 18 | η quạt tổng 0,6; dự phòng động cơ 1,15 | – | Chỉ đạo WP1; thực hành chung |
| 19 | CDU: COP 3,0 @37,8 °C; ≤ 50 kW danh định/module; suy giảm ống gas −0,2 %/m sau 7,5 m; ống gas 42,9 / 19,5 / 34,4 m | – | Chỉ đạo WP1; catalogue điển hình; WP2 |
| 20 | Scrubber: +10 % lưu lượng; v tháp 1,8 m/s; L/G 2,5 L/m³; bơm 20 m, η 0,55; ΔP quạt 600 + 400 Pa | – | Chỉ đạo WP1; thực hành chung |
| 21 | Máy hút ẩm: nồm 22 °C/95 %; thấm 1 ACH; 10 h có người; dự phòng 1,3; năng suất @25/55 = 0,45 × định mức @30/80; 2,5 L/kWh | – | Giả định WP1 |
| 22 | Nước ngưng: độ rơi khả dụng 0,15 m, dốc 1 % → L_max = 15 m; bơm khi L > 12 m; tuyến tới CD-2.01 ngoài tường A: 36,6 / 13,1 / 29,4 m | – | Giả định WP1; WP2 |
| 23 | Sưởi: gia nhiệt lại tối thiểu 2 K; 10 % FAFE tràn khi tính đông | – | Giả định WP1 |

## 4. Các quyết định thiết kế chính

### 4.1 Phần thiếu hụt 10 % của gió bù – lấy từ đâu
Mỗi phòng có FAFE nhận vào 90 % gió thải; 10 % còn lại (Σ **4.395 m³/h**: 4.235 m³/h ở 6 phòng hóa + 160 m³/h Sample preparation) phải vào phòng dưới dạng gió chuyển. Khối trái và khối phải chỉ giáp xưởng / lối đi không điều hòa (WP2) – không có hành lang điều hòa. Nếu để gió này lọt từ xưởng vào, riêng khối phải nhận ≈ 2.750 m³/h không xử lý (≈ 33 kW ẩn) trong khi gió cấp ST2 chỉ khử được ≈ 3 kW → RH vượt 70 %. Vì vậy:

- Gió chuyển lấy từ **phòng sạch cùng AHU**, qua khe cửa / miệng gió chuyển theo đúng cửa và vách chung của WP2; phòng nguồn được cấp dư (Cấp > Hồi):

| Từ | Đến | m³/h | Đường đi (WP2) |
|---|---|---|---|
| Sample preparation | Sample treatment 1 | 1.485 | miệng gió chuyển có cánh chặn sáng trên vách chung Sample prep / ST1 (2,65 m); gồm phần đi tiếp sang Kjeldahl, Digestion |
| Sample treatment 1 | Kjeldahl room | 674 | khe cửa D-02; gồm phần đi tiếp sang Digestion |
| Kjeldahl room | Digestion room | 564 | khe cửa D-01 (Digestion chỉ đi qua Kjeldahl) |
| Office | Sample treatment 2 | 1.000 | khe cửa D-14 + miệng gió chuyển vách chung (8,8 m) |
| Sample treatment 2 | Standard (11,90 m²) | 1.068 | khe cửa D-17; gồm phần đi tiếp sang Standard 19,72 |
| Standard (11,90 m²) | Standard (19,72 m²) | 938 | khe cửa D-16 |

- Phân cấp áp: Sample prep (+) → ST1 (−) → Kjeldahl (−−) → Digestion (−−−); Office (+) → ST2 (−) → Standard 11,90 (−−) → Standard 19,72 (−−−). Phần còn lại của ST2 (1.750 m³/h) và Sample prep (160 m³/h) do chính AHU cấp dư vào phòng đó. Mọi phòng hóa có Hồi = Cấp; với xưởng/lối đi các phòng gần trung tính (cửa ra lối đi D-03/D-04/D-15/D-18 không có dòng ròng).
- Ở cấp AHU: Σ(Cấp − Hồi) = Σ(Thải − FAFE) = **gió tươi phải đi qua AHU** (gió chuyển triệt tiêu trong khối). Gió tươi AHU-2.01 = 1.650, AHU-2.02 = 2.750, AHU-2.03 = 1.200 m³/h (Σ 5.600, so với 1.875 theo người).
- **Hệ quả**: phần gió tươi vượt nhu cầu người tốn **80,2 kW** tải coil (AHU-2.02: 46,3 kW). Đây là giá của FAFE = 90 %. Độ nhạy (§10): FAFE 95 % → 243,1 kW; 100 % → 230,9 kW.

### 4.2 Instrument room – gió thải bù bằng gió tươi đã xử lý
8 điểm hút thiết bị không phải tủ/chụp nên gió bù không xử lý không bị hút ngay mà hòa vào phòng: 1.080 m³/h ở 26,3 g/kg mang ≈ 14 kW ẩn, gấp ~2,7 lần khả năng khử ẩm của 12.430 m³/h gió cấp ở 55 % (5,1 kW) → không giữ được RH ≤ 55 %. WP1 quyết định **không dùng FAFE cho Instrument**; AHU-2.03 lấy 1.200 m³/h gió tươi bằng lượng thải (phòng trung tính). Chi phí ≈ 20,4 kW lạnh. **Sai khác so với DB §4.2 – cần chủ trì duyệt.**

### 4.3 Phân khối AHU – chọn 3 AHU (PA B)

| Tiêu chí | PA A: 2 AHU (Instrument trong AHU-2.01) | PA B: 3 AHU (chọn) |
|---|---|---|
| Gió cấp AHU-2.01 | 34.610 m³/h (Instrument 36 %) | 22.180 + 12.430 (AHU-2.03) |
| Tải coil AHU-2.01 | 202,2 kW (Instrument 34 %) | 132,8 + 69,4 kW |
| Tổng gió / tổng tải | 39.860 m³/h / 273,1 kW | bằng PA A (cộng tuyến tính) |
| Điều khiển T/RH | 1 vùng cho Instrument (≤ 55 %, 24/7) và 8 phòng (60–70 %, giờ hành chính) – không giữ được đồng thời | Instrument điều khiển riêng, chạy 24/7 độc lập |
| Dự phòng | Hỏng AHU-2.01 dừng cả thiết bị phân tích | Tách rủi ro |
| Kích thước / bố trí | 1 AHU ~35.000 m³/h, ~222 kW – lớn cho sàn trên trần Balance | 3 máy vừa phải; AHU-2.03 trên trần Instrument (vị trí WP2), ống cấp ≈ 4 m |

Kết luận: **PA B** – AHU-2.01 (trên trần Balance, khối trái trừ Instrument), AHU-2.02 (trên trần ST2 phía ĐN), AHU-2.03 (trên trần Instrument, cạnh DH-2.01). Không tăng công suất lạnh; thêm 1 AHU + 1 tủ DDC + 1 louver LV-2.03.

### 4.4 Bội số trao đổi cao (> 40 ACH)
Kjeldahl 63,4; ST1 60,2; Instrument 115,7 ACH – cả ba do **tải hiện** (thiết bị 0,47–0,87 kW/m²). ΔT cấp đã tối đa: T sau coil cố định 15 °C (DB) → gió cấp 16,36 °C → ΔT 8,64 K; không thể giảm ACH bằng ΔT. Muốn giảm phải giảm nhiệt tỏa (xác nhận công suất tỏa thực – §12) hoặc thêm dàn lạnh cục bộ chỉ khử nhiệt hiện (ngoài DB). 100+ ACH là bình thường cho phòng thiết bị kiểu phòng máy chủ. Yêu cầu bố trí miệng gió (WP3): Instrument ~12 miệng khuếch tán xoáy ~1.000 m³/h; ST1/Kjeldahl dùng miệng đục lỗ vận tốc thấp, xa mặt tủ hood (gió ngang tại mặt tủ ≤ 0,2 m/s, EN 14175). KiemTra coi ACH > 40 là ĐẠT khi gió cấp do tải hiện quyết định.

### 4.5 Kiểm tra ẩm và gió bù tràn
Mọi phòng đạt kiểm tra ẩm ở RH thiết kế (gió chuyển lấy từ phòng cùng điểm ẩm nên không mang ẩn). Tuy nhiên giả thiết SPL = 0 rất nhạy: tỷ lệ tràn FAFE chịu được (giữ ≤ 70 %) chỉ **1,4 % ở Standard 19,72 (4 hood trong 19,7 m²); 5,2 % ở ST2; 6,9 % ở Digestion; 9,8 % ở Standard 11,90** (ST1 46 %, Sample prep 53 %, Kjeldahl > 100 %). Độ nhạy SPL = 10 % cho tải 403,6 kW và 7 hạng mục KiemTra không đạt. → Miệng cấp FAFE phải đặt sát/tích hợp mặt tủ (slot/đục lỗ, ≤ 0,3 m/s, hướng vào vùng hút); van gió động cơ tại mỗi miệng bù mở theo tín hiệu chạy quạt tủ; AHU-2.02 có sưởi 18 kW (gia nhiệt lại tới 10 K) cho chế độ khử ẩm.

## 5. Kết quả theo phòng

| Phòng | AHU | RSH (W) | RLH (W) | Cấp (m³/h) | ACH | Hồi | Thải | FAFE | Chuyển vào | Chuyển ra | Tràn FAFE chịu được | FAF | EAF |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Digestion room | AHU-2.01 | 3.625 | 121 | 1.260 | 28,8 | 1.260 | 5.640 | 5.076 | 564 | 0 | 6,9 % | FAF-2.01 | EAF-2.01 |
| Kjeldahl room | AHU-2.01 | 10.390 | 121 | 3.590 | 63,4 | 3.590 | 1.100 | 990 | 674 | 564 | > 100 % | FAF-2.01 | EAF-2.01 |
| Sample treatment 1 | AHU-2.01 | 34.370 | 605 | 11.870 | 60,2 | 11.870 | 8.110 | 7.299 | 1.485 | 674 | 46,1 % | FAF-2.01 | EAF-2.01 |
| Balance room | AHU-2.01 | 1.389 | 242 | 480 | 8,7 | 480 | 0 | 0 | 0 | 0 | – | – | – |
| Sample preparation | AHU-2.01 | 8.234 | 605 | 2.850 | 16,1 | 1.205 | 1.600 | 1.440 | 0 | 1.485 | 53,1 % | FAF-2.01 | EAF-2.03 |
| Sample coding | AHU-2.01 | 1.332 | 121 | 460 | 6,1 | 460 | 0 | 0 | 0 | 0 | – | – | – |
| Sample dividing | AHU-2.01 | 771 | 60 | 270 | 9,0 | 270 | 0 | 0 | 0 | 0 | – | – | – |
| Reporting room | AHU-2.01 | 4.038 | 605 | 1.400 | 25,9 | 1.400 | 0 | 0 | 0 | 0 | – | – | – |
| Instrument room | AHU-2.03 | 36.000 | 605 | 12.430 | 115,7 | 11.230 | 1.200 | 0 | 0 | 0 | – | – | EAF-2.03 |
| Sample treatment 2 | AHU-2.02 | 8.520 | 605 | 2.950 | 9,5 | 1.200 | 16.820 | 15.138 | 1.000 | 1.068 | 5,2 % | FAF-2.02 | EAF-2.02 |
| Standard (19,72 m²) | AHU-2.02 | 1.302 | 121 | 450 | 7,6 | 450 | 9.380 | 8.442 | 938 | 0 | 1,4 % | FAF-2.02 | EAF-2.02 |
| Standard (11,90 m²) | AHU-2.02 | 1.258 | 121 | 440 | 12,3 | 440 | 1.300 | 1.170 | 1.068 | 938 | 9,8 % | FAF-2.02 | EAF-2.02 |
| Office | AHU-2.02 | 4.065 | 605 | 1.410 | 11,5 | 410 | 0 | 0 | 0 | 1.000 | – | – | – |
| **Tổng** | | **115.293** | **4.538** | **39.860** | 30,1 | **34.265** | **45.150** | **39.555** | 5.729 | 5.729 | | | |

Nhiệt thiết bị vào phòng Σ 88,4 kW (Instrument 31,2; ST1 28,9; Kjeldahl 8,8; Sample prep 5,4; ST2 4,4). T phòng dự kiến khi cấp theo thiết kế: 24,8–25,0 °C ở mọi phòng.

Gió thải theo loại: hood 34.050 (acid 9.080 → scrubber; hữu cơ 24.970 xả thẳng SS304), chụp treo 2.200, tủ hóa chất 5.200, bồn rửa 900, arm hood 1.600, điểm thiết bị 1.200 → **Σ 45.150 m³/h**; trong đó qua quạt EAF của HVAC 11.100 m³/h.

## 6. AHU và dàn nóng

| Hạng mục | AHU-2.01 | AHU-2.02 | AHU-2.03 | Tổng |
|---|---|---|---|---|
| Vị trí (WP2) | trên trần Balance | trên trần ST2 phía ĐN | trên trần Instrument | |
| Phòng phục vụ | Digestion, Kjeldahl, ST1, Balance, Sample prep, Coding, Dividing, Reporting | ST2, Standard 19,72, Standard 11,90, Office | Instrument | |
| Gió cấp (m³/h) | 22.180 | 5.250 | 12.430 | 39.860 |
| Gió tươi (m³/h) / tỷ lệ | 1.650 / 7,4 % | 2.750 / 52,4 % | 1.200 / 9,7 % | 5.600 |
| Gió hồi về AHU (m³/h) | 20.530 | 2.500 | 11.230 | 34.260 |
| Vot theo người / Ev | 1.223 / 0,84 | 640 / 0,94 | 250 / 1,0 | |
| Vào coil (T / RH) | 25,95 °C / 61,7 % | 31,70 °C / 65,8 % | 26,24 °C / 58,0 % | |
| Ra coil | 15 °C / 97 % | 15 °C / 97 % | 15 °C / 97 % | |
| Tải coil tổng / hiện / ẩn (kW) | 132,8 / 81,5 / 51,3 | 70,9 / 29,4 / 41,5 | 69,4 / 46,8 / 22,6 | 273,1 |
| SHR | 0,61 | 0,41 | 0,67 | |
| **Công suất lạnh chọn (+10 %)** | **146,1 kW** (41,5 TR) | **78,0 kW** (22,2 TR) | **76,3 kW** (21,7 TR) | **300,4 kW** |
| Công suất hiện chọn (kW) | 89,6 | 32,4 | 51,5 | |
| Nước ngưng (kg/h) | 70,3 | 57,6 | 30,6 | 158,4 |
| ESP thiết kế / ước tính tuyến (Pa) | 500 / 233 | 500 / 206 | 500 / 184 | |
| Áp tổng quạt (Pa) | 985 | 985 | 985 | |
| Trục quạt / động cơ (kW) | 10,1 / 15 | 2,4 / 3 | 5,7 / 7,5 | |
| Sưởi điện chọn (kW) | 15 | 18 | 9 | 42 |
| Nhu cầu đông / gia nhiệt 2 K (kW) | 11,0 / 14,9 | 17,0 / 3,5 | 3,8 / 8,3 | |
| Mùa đông: T sau sưởi ≥ T cấp yêu cầu | 19,4 ≥ 18,8 °C | 23,9 ≥ 22,9 °C | 19,4 ≥ 18,0 °C | |
| Lọc / mặt coil @2,5 m/s | G4+F8 / 2,46 m² | G4+F8 / 0,58 m² | G4+F8 / 1,38 m² | |
| Điện AHU (quạt + sưởi) kW / FLA (A) | 30 / 54,3 | 21 / 33,7 | 16,5 / 29,4 | 67,5 |
| Ống gas tới bãi CDU / hệ số suy giảm | 42,9 m / 0,93 | 19,5 m / 0,98 | 34,4 m / 0,95 | |
| **CDU**: module × danh định @37,8 °C | **CDU-2.01: 4 × 39,3 kW** | **CDU-2.02: 2 × 40,0 kW** | **CDU-2.03: 2 × 40,3 kW** | 8 module |
| CDU thực sau suy giảm / module | 36,5 kW | 39,0 kW | 38,2 kW | |
| CDU điện/module (COP 3) kW / FLA | 13,1 / 23,4 | 13,3 / 23,8 | 13,4 / 24,0 | 105,9 kW |

Ghi chú: (1) AHU-2.02 có 52 % gió tươi, SHR 0,41 – ADP ≈ 13,9 °C, hệ số bypass ≈ 0,06 → coil DX sâu 6–8 hàng, bay hơi ≈ 6–7 °C; NCC phải xác nhận đạt 15 °C/97 %. (2) Mỗi module CDU một mạch coil riêng (coil xen kẽ), inverter, R410A/R32. Do ống gas AHU-2.01 dài 42,9 m (WP2), CDU-2.01 tăng từ 3 lên 4 module. (3) Cổ ống đề xuất (8 m/s; tươi 5 m/s): AHU-2.01 cấp 1300×600, hồi 1300×550, tươi 400×250; AHU-2.02 cấp 450×400, hồi 350×250, tươi 450×350; AHU-2.03 cấp 800×550, hồi 650×600, tươi 350×200 – WP3 chốt.

## 7. Quạt cấp bù FAF, quạt hút EAF, quạt tủ hood, scrubber

Tag theo đề xuất WP2 (thay cho bản nháp 7 FAF + 5 EAF theo từng phòng). Lý do gộp: các phòng hóa mỗi khối liền nhau, chung louver và tuyến; vì quạt tủ hood bật/tắt riêng lẻ nên dù quạt theo phòng hay theo nhóm đều cần van gió động cơ tại mỗi miệng bù → gộp nhóm + biến tần giữ áp ống là đơn giản và ít thiết bị nhất.

**FAF** (gió tươi không xử lý, lọc G4+F8, áp tổng 685 Pa = ESP 400 + lọc 285; biến tần giữ áp ống; van gió động cơ ON/OFF tại mỗi tủ/chụp, mở theo tín hiệu chạy quạt tủ / EAF):

| Tag | Vị trí (WP2) | Phục vụ (m³/h) | Σ m³/h | Động cơ | Điện / FLA | Cổ ống |
|---|---|---|---|---|---|---|
| FAF-2.01 | trên trần ST1 sát lối đi; tươi từ LV-2.01 (≈ 32 m) | Digestion 5.076, Kjeldahl 990, ST1 7.299, Sample prep 1.440 | **14.805** | 5,5 kW | 380V/3P / 11,6 A | 1050×500 |
| FAF-2.02 | trên trần ST2 gần AHU-2.02; tươi từ LV-2.02 | ST2 15.138, Standard 19,72 8.442, Standard 11,90 1.170 | **24.750** | 11 kW | 380V/3P / 23,1 A | 1750×500 |
| **Σ** | | | **39.555** | 16,5 kW | | |

**EAF** (ESP 400 Pa, trên trần, xả qua tường A + weather cover + lưới côn trùng, biến tần):

| Tag | Vị trí → điểm xả (WP2) | Phục vụ | m³/h | Vật liệu | Động cơ / điện | Cổ ống |
|---|---|---|---|---|---|---|
| EAF-2.01 | trên trần Kjeldahl → DP-2.01 | chụp treo Digestion 1.100 + Kjeldahl 1.100, tủ HC ST1 1.300 | 3.500 | SS304 | 0,75 kW 380V/3P | Ø400 |
| EAF-2.02 | trên trần ST2 phía tường A → DP-2.02 | 2 tủ HC + 2 bồn rửa ST2 3.200, bồn rửa Standard 19,72 300, tủ HC Standard 11,90 1.300 | 4.800 | SS304 | 1,1 kW 380V/3P | Ø450 |
| EAF-2.03 | trên vách Sample prep / Instrument → DP-2.03 (≈ 32 m) | 2 arm hood 1.600 + 8 điểm Instrument 1.200 | 2.800 | SS304 (nhánh arm hood GI) | 0,75 kW 380V/3P | Ø350 |
| **Σ** | | | **11.100** | | 2,6 kW | |

ESP 400 Pa đủ cho tuyến dài nhất EAF-2.03 (≈ 32 m chính + 12 m nhánh → ≈ 90 Pa ma sát ×2 phụ kiện + weather cover); WP3 kiểm lại. Nhánh Instrument chạy 24/7, nhánh arm hood gián đoạn → van gió động cơ trên nhánh arm hood.

**Quạt tủ hood (thuộc gói nội thất – HVAC chỉ cấp nguồn + điều khiển)**: 15 điểm × 0,6 kW, 380V/3P, ≈ 1,3 A/điểm: FHF-2.01 Digestion 2 (acid), FHF-2.02 ST1 3 (1 acid), FHF-2.03 ST2 6 (1 acid: ST2-FH3), FHF-2.04 Standard 19,72 4. Σ 9,0 kW. Hood hữu cơ (11 tủ, lọc than tại tủ) xả thẳng ống SS304 qua DP-2.01/DP-2.02; 4 hood acid về scrubber bằng ống PP.

**SCR-2.01** – scrubber ướt tháp đệm PP, ngoài tường A phía bắc cửa CST (WP2): Σ hood acid 9.080 → thiết kế **10.000 m³/h**; tiết diện 1,54 m², **Ø 1.450 mm**; tuần hoàn 25 m³/h; bơm 3 kW (trục 2,5 kW, 20 m); quạt FRP/PP 1.000 Pa – **5,5 kW** (trục 4,6 kW), biến tần giữ áp âm ống gom; bơm định lượng NaOH + pH 0,18 kW; tổng **8,7 kW, 380V/3P, ≈ 18 A**; cổ vào Ø650 PP. Quạt tủ acid làm việc nối tiếp với quạt scrubber. Tuyến PP dài nhất ≈ 31 m (WP2) – trong mức 400 Pa ống. Nước cấp + đấu XLNT do gói khác.

## 8. Máy hút ẩm và nước ngưng

**DH-2.01** (Instrument, trên trần cạnh AHU-2.03): ẩm do người 7,9 kg/ngày + thấm cửa khi nồm 15,2 kg/ngày, ×1,3 → **30,1 L/ngày @25 °C/55 %** ≈ 66,8 L/ngày định mức @30 °C/80 % → chọn **70 L/ngày**, ≈ 1,17 kW, 220V/1P, loại âm trần nối ống, bơm ngưng tích hợp, chạy khi RH > 55 %. Nhiệt thải khi chạy ≈ 2,0 kW (không cộng ở điểm hè vì lúc đó AHU-2.03 khử ẩm).

**Nước ngưng** (AHU đặt trên sàn trên trần, điểm thoát CD-2.01 ngoài tường A – WP2): độ rơi khả dụng ≈ 0,15 m, dốc 1 % → tự chảy tối đa 15 m; dùng bơm khi tuyến > 12 m. Cả 3 tuyến đều vượt:

| Tag | Phục vụ | Tuyến (WP2) | Lưu lượng bơm | Điện |
|---|---|---|---|---|
| CP-2.01 | AHU-2.03 + DH-2.01 | 29,4 m | ≥ 96 L/h, 6 m | 0,1 kW 220V/1P |
| CP-2.02 | AHU-2.01 | 36,6 m | ≥ 211 L/h, 6 m | 0,1 kW 220V/1P |
| CP-2.03 | AHU-2.02 | 13,1 m | ≥ 173 L/h, 6 m | 0,1 kW 220V/1P |

Bơm loại có bình chứa, phao mức cao cắt cấp lạnh AHU. Ống uPVC bọc cách nhiệt. WP2 nêu tự chảy khả thi nếu đáy khay AHU-2.01 ≥ +3,6 m; nếu bệ AHU (phần xây dựng) đủ cao để có độ rơi ≥ 0,45 m thì bỏ được CP-2.02/CP-2.03 – quyết định khi chốt cao độ bệ.

## 9. Lịch thiết bị – tổng hợp điện (đầu vào WP4/WP5)

| Nhóm | Tag | SL | Tổng điện (kW) |
|---|---|---|---|
| AHU (quạt + sưởi) | AHU-2.01 / 2.02 / 2.03 | 3 | 67,5 |
| Dàn nóng | CDU-2.01 (4), CDU-2.02 (2), CDU-2.03 (2) | 8 | 105,9 |
| Quạt cấp bù | FAF-2.01, FAF-2.02 | 2 | 16,5 |
| Quạt hút | EAF-2.01, 2.02, 2.03 | 3 | 2,6 |
| Điểm nguồn quạt tủ hood | FHF-2.01…2.04 | 15 điểm | 9,0 |
| Scrubber | SCR-2.01 | 1 | 8,7 |
| Máy hút ẩm | DH-2.01 | 1 | 1,2 |
| Bơm nước ngưng | CP-2.01, 2.02, 2.03 | 3 | 0,3 |
| **Tổng điện đặt HVAC** | | | **211,7 kW** |

Chi tiết (FLA, điện áp, cổ ống, phân bổ lưu lượng) trong sheet LichThietBi và `equipment_schedule.json`. WP4 lưu ý: sưởi và lạnh của cùng AHU chỉ chạy đồng thời ở chế độ khử ẩm; hệ số đồng thời do WP4 quyết định.

## 10. Độ nhạy (chạy lại toàn bảng bằng `OVR`)

| Kịch bản | Gió tươi AHU (m³/h) | Tải coil (kW) | KiemTra |
|---|---|---|---|
| **Cơ sở v4** (FAFE 90 %, SPL 0, ESP 500) | 5.600 | **273,1** | đạt |
| FAFE 95 % | 3.810 | 243,1 | đạt |
| FAFE 100 % (phòng trung tính) | 3.080 | 230,9 | đạt; mất âm áp phòng hóa |
| SPL 10 % (gió bù tràn vào phòng) | 5.600 | 403,6 | 7 lỗi (ẩm, T phòng hóa) → phải khống chế tràn |
| ESP tối thiểu 800 Pa (như Phase 1) | 5.600 | 282,1 | đạt; quạt 18,5 / 4 / 11 kW |
| Chỉ thiết bị năm 2027 | 5.600 | 217,3 | đạt; tham khảo phân kỳ |

So với v3 (202,6 kW): +70,5 kW, chủ yếu do gió tươi xử lý tăng 1.875 → 5.600 m³/h (+80,2 kW: thiếu hụt 10 % gió bù + bù thải Instrument), trừ phần gió hồi giảm tương ứng (≈ −18 kW), cộng gió cấp tăng do T cấp 16,36 °C, thải mới (arm hood, Instrument) và tường ngoài A (≈ +8 kW). v3 dùng FAFE 100 % và chưa có arm hood/điểm hút Instrument.

## 11. Danh sách tự kiểm tra (sheet KiemTra – 133/133 ĐẠT)

1. Cân bằng khối lượng từng phòng (13): |Cấp + Bù + Chuyển vào − Hồi − Thải − Chuyển ra| < 0,5 m³/h.
2. Gió hồi ≥ 0 (13); ACH ≥ 6 (13); ACH ≤ 40 hoặc do tải hiện (13); khử ẩm đủ RLH (13); T phòng dự kiến 23–27 °C (13).
3. Σ gió chuyển vào = Σ ra; gió chuyển chỉ trong cùng AHU.
4. Σ diện tích = 441,76 m²; Σ người = 75; 15 hood / 4 acid; Σ hood acid = 9.080; Σ nhiệt thiết bị phòng = Σ sheet ThietBi.
5. Σ thải = Σ EAF + Σ hood; Σ FAF = Σ FAFE phòng; scrubber ≥ 1,1 × hood acid.
6. Mỗi AHU (×3): Cấp = Hồi + Tươi; Tươi ≥ Vot; dư áp ≥ 0; Σ hồi phòng − dư áp = hồi AHU; coil khử ẩm và RH vào ≤ 100 %; tải coil ≥ Σ tải phòng; Σ CDU thực ≥ công suất chọn; module danh định ≤ 50 kW; sưởi đông đủ; động cơ ≥ 1,15 × trục; ESP thiết kế ≥ ước tính; T cấp dùng ở TaiPhong = T sau coil + ΔT quạt.
7. Tổng PA B = PA A (gió, tải); Σ gió cấp phòng = Σ AHU; Instrument RH ≤ 55 %; máy hút ẩm ≥ yêu cầu; có bơm ở nơi cần; động cơ FAF/EAF/scrubber ≥ 1,15 × trục; lịch thiết bị khớp sheet AHU.
8. Quét lỗi công thức sau khi LibreOffice tính lại: 0 lỗi; không ô văn bản nào bắt đầu bằng “=”.

## 12. Vấn đề mở / cần xác nhận

1. **FAFE 90 % → +80 kW lạnh** (gió tươi xử lý thay cho 10 % thiếu hụt). Chủ trì chốt giữ 90 % hay 95 % (§4.1, §10).
2. **Instrument không dùng FAFE** (bù 100 % qua AHU-2.03) – sai khác DB §4.2, cần duyệt (§4.2).
3. **Standard 19,72**: 4 hood + 1 bồn rửa (9.380 m³/h thải ≈ 159 ACH thải) trong 19,7 m², chỉ chịu được 1,4 % FAFE tràn → miệng bù phải tích hợp tại mặt tủ (kiểu auxiliary-air) hoặc chuyển bớt 1 hood sang ST2; cần chủ trì/CĐT xem lại bố trí.
4. **ST1 tủ hóa chất thứ 2** (ST1-CC2, bản vẽ có, danh mục nội thất 5.12 chỉ 1 tủ có hút): nếu có hút thêm 1.300 m³/h → EAF-2.01 4.800 m³/h, FAF-2.01 +1.170, gió tươi AHU-2.01 +130 m³/h (≈ +3 kW lạnh).
5. **Nhiệt tỏa thiết bị Instrument / ST1 / Kjeldahl** quyết định ACH 60–116 và cỡ AHU-2.03: đề nghị CĐT cung cấp “site preparation” (LC-MS/MS, GC-MS/MS, HPAEC…) và hệ số đồng thời bếp/bể điều nhiệt ST1.
6. **Số chụp treo**: DB = 1 chụp/phòng (Digestion, Kjeldahl); danh mục nội thất 6.5 và 7.4 ghi 2 chụp/phòng → mỗi chụp thêm 1.100 m³/h thải + 990 FAFE + 110 tươi xử lý.
7. **Số quạt tủ hút trong danh mục CĐT** (dòng “Tủ hút/chụp hút”: 27) ≠ 15 tủ hood theo DB – HVAC tính 15 điểm nguồn.
8. **Khống chế gió bù tràn** ở ST2 (5,2 %), Digestion (6,9 %), Standard 11,90 (9,8 %): miệng FAFE sát mặt tủ; van gió động cơ mỗi miệng bù; tín hiệu chạy quạt tủ về DDC-FAN – WP3/WP4.
9. **Bố trí miệng gió cấp ở phòng ACH cao** gần tủ hood (≤ 0,2 m/s tại mặt tủ) – WP3.
10. **Cao độ bệ AHU** (xây dựng) quyết định có bỏ được bơm nước ngưng không (§8).
11. **AHU-2.02 SHR 0,41** – NCC xác nhận coil đạt 15 °C/97 % (6–8 hàng) hoặc tách tiền xử lý gió tươi.
12. **Ống gas dài** (AHU-2.01 42,9 m, AHU-2.03 34,4 m): hệ số suy giảm −0,2 %/m là giả định; kiểm tra theo catalogue hãng (giới hạn chiều dài, chênh cao, dầu hồi).
13. **Mùa đông**: FAFE không xử lý cấp 9,8 °C trước tủ – sưởi AHU đã tính 10 % tràn; nếu người dùng phàn nàn lạnh tại mặt tủ cần sưởi FAFE (ngoài DB).
14. **Miệng gió chuyển trên vách** (Sample prep/ST1, Office/ST2): kiểm tra yêu cầu ngăn cháy (QCVN 06) – nếu là vách ngăn cháy phải có van chặn lửa.
15. **Máy hút ẩm**: điều kiện nồm 22 °C/95 % và tỷ lệ năng suất 0,45 là giả định; chọn model theo đặc tính NCC tại 25 °C/55 %.
16. FLA là ước tính (PF 0,85, η 0,85) – thay bằng số liệu NCC khi chọn model.
