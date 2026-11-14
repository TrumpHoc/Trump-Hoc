# WP4 – ĐIỆN & ĐIỀU KHIỂN HVAC – SGS HẢI PHÒNG, KHU H&N PHASE 2 (NHÀ XƯỞNG 3)

> v1.0 – 14/11/2026. Phụ thuộc: `01_design_basis/DESIGN_BASIS.md` v1.1, WP1 `02_calc/equipment_schedule.json` (v4), WP2 `03_geometry/geometry.json`.
> Kết quả: `electrical_design.json`, `takeoff_electrical.json`, `WP4_Electrical.xlsx` (7 sheet: Bảng tra, Phụ tải, Sơ đồ nguyên lý, Chọn cáp & sụt áp, IO – Excel không cho ký tự “/” trong tên sheet, Máng cáp, Tổng hợp khối lượng; công thức, đã tính lại bằng LibreOffice – 0 lỗi), `WP4_single_line.png`, `WP4_containment.png`. Chạy lại: `python3 -I 05_electrical/build_electrical.py`.

## 1. Phạm vi

Theo DESIGN_BASIS §0 và §4.6: **từ tủ tổng DB.HVAC trở đi** (gồm vỏ tủ, MCCB tổng, thanh cái, các lộ ra), các tủ DDC/động lực riêng, toàn bộ cáp lộ ra, máng/ống, cáp điều khiển, cảm biến, lập trình & T&C. **Ngoài phạm vi**: cáp và thiết bị cấp nguồn tới DB.HVAC, thanh tiếp địa chính tòa nhà, hệ báo cháy (chỉ nhận tín hiệu từ module rơ le của gói PCCC), thiết bị (AHU, CDU, quạt, scrubber, DH, bơm – WP5), thân van gió động cơ (WP3), quạt tủ hood (gói nội thất).

## 2. Vị trí tủ (giả định WP4)

Không có phòng điện/“ELE”/“Tủ IT” trong khu Phase 2 theo WP2 (13 phòng đều là phòng chức năng; khu trái trên/phải trên là ngoài phạm vi). Do đó các tủ đặt **tại lối đi kỹ thuật giữa 2 khối** (X 14.385…16.185, thông thủy 1,8 m, không điều hòa, không phải phòng thí nghiệm), cao độ sàn, tiếp cận trực tiếp – đúng vị trí gợi ý DB.HVAC của WP2 (X ≈ 16.400, Y ≈ 2.200). Biến tần/contactor nằm trong tủ ở sàn (không đặt trên sàn AHU) để bảo trì an toàn; trên sàn thiết bị chỉ có cầu dao cách ly + ổ cắm bảo trì.

| Tủ | Tọa độ tâm (mm) | Lắp đặt | Kích thước | IP | Vị trí |
|---|---|---|---|---|---|
| DB.HVAC | 15935; 2200 | Đặt sàn (đế 100 mm) | 1200W × 2100H × 500D | IP42 | Vách khối phải (Office) phía lối đi, Y 1.600…2.800 (giữa cửa D-13 và D-15) – đúng vị trí gợi ý WP2 |
| DDC-FAN | 15935; 700 | Đặt sàn (đế 100 mm) | 1200W × 2100H × 500D | IP42 | Vách khối phải phía lối đi, Y 100…1.300, cạnh DB.HVAC |
| DDC-AHU-2.02 | 16035; -900 | Treo tường, đáy +0,5 m | 1000W × 1400H × 300D | IP54 | Vách khối phải phía lối đi, Y −1.400…−400 |
| DDC-AHU-2.01 | 14585; 3900 | Treo tường, đáy +0,3 m | 1000W × 1800H × 400D | IP54 | Vách khối trái (Balance/Reporting) phía lối đi, Y 3.400…4.400 (giữa cửa D-05 và D-06) |
| DDC-AHU-2.03 | 14585; -4800 | Treo tường, đáy +0,3 m | 1000W × 1600H × 400D | IP54 | Vách khối trái (Reporting) phía lối đi, Y −5.300…−4.300 (đối diện vách đặc) |


Khoảng trống trước tủ đặt sàn còn ≥ 1,3 m (lối đi 1,8 m − tủ 0,5 m), tủ treo tường hai bên lệch nhau theo Y (không đối diện) – thỏa IEC 60364-7-729 (≥ 0,7 m). Cần CĐT xác nhận lối đi được phép đặt tủ (lối thoát nạn?) – vấn đề mở #1.

## 3. Phương pháp & giả định

| # | Giả định / tham số | Giá trị | Nguồn |
|---|---|---|---|
| 1 | Điện áp | 380/220 V, 3P+N+PE, 50 Hz, TN-S | Lưới VN; Phase 1 |
| 2 | cosφ | Động cơ/CDU 0,85; tải qua biến tần 0,95 (phía lưới); sưởi điện trở 1,0 | Chỉ đạo WP4; vật lý tải điện trở |
| 3 | Hiệu suất | Động cơ IE3 4 cực (0,75 kW 82,5 % … 15 kW 92,1 %); biến tần 97 %; quạt hood 85 %; scrubber 88 % | IEC 60034-30-1; WP1 FLA |
| 4 | P CDU | Điện vào theo WP1 (COP 3,0): 16,54 / 13,43 / 13,45 kW/module; Ib = max(tính, FLA WP1) | equipment_schedule.json |
| 5 | Chọn bảo vệ | In ≥ 1,25·Ib (MCB/MCCB); động cơ trực tiếp: MPCB theo dải dòng; lộ tổng tủ DDC ≥ 1,25·ΣIb và ≥ 1,6·In lớn nhất phía sau (chọn lọc) | IEC 60364-4-43; thực hành |
| 6 | Khả năng mang tải cáp | Cu/XLPE/PVC (CXV) nhiều lõi trên khay đục lỗ – PP E, Bảng B.52.12 | IEC 60364-5-52 ≡ TCVN 9207:2012 |
| 7 | Hệ số hiệu chỉnh | Nhiệt độ 45 °C trên trần: 0,87; nhóm: 0,80 → Iz = 0,696·Iz0 | B.52.14; chỉ đạo WP4 (B.52.17: mạch < 30 % Iz bỏ qua) |
| 8 | Chiều dài | Theo mạng máng (tọa độ WP2) + đoạn xuống thiết bị (Manhattan); ngang × 1,15 + đứng (máng đứng nóc tủ 1,5 m, trong tủ 1 m, xuống thiết bị trên trần 1 m, cảm biến tường 2,6 m, CDU 3,1 + 1 m, scrubber 4,1 m) | WP2; chỉ đạo WP4 |
| 9 | Sụt áp | ΔU tích lũy DB.HVAC → tải ≤ 4 %; cáp lộ tổng mục tiêu ≤ 1,5 %; ρ = 0,0225 Ω·mm²/m, X = 0,08 mΩ/m | IEC 60364-5-52 G.52 |
| 10 | Tiết diện PE | S ≤ 16: PE = S; 16 < S ≤ 35: 16; S > 35: S/2 (lõi PE trong cáp) | IEC 60364-5-54 Bảng 54.2 |
| 11 | Máng cáp | BOP +3,6 m (như Phase 1 F3-SD-AC-013 “CABLE TRAY … BOP=FFL3600”); máng động lực và máng điều khiển riêng; độ đầy ≤ 40 % | Phase 1; thực hành VN |
| 12 | Dự phòng | 20 % số lộ mỗi tủ (làm tròn lên), 20 % I/O mỗi DDC | Chỉ đạo WP4 |
| 13 | Ik tại DB.HVAC | Chưa có số liệu → MCCB tổng Icu ≥ 36 kA, lộ ra ≥ 25 kA, tủ DDC 10 kA | Giả định – gói điện xác nhận |
| 14 | Quạt EAF | Khởi động trực tiếp (MPCB + contactor) theo chỉ đạo WP4 (lưu lượng cố định, cân chỉnh bằng van). WP1 ghi “biến tần” – xem vấn đề mở #4 | Chỉ đạo WP4 |
| 15 | Quạt AHU | Động cơ AC + biến tần trong tủ DDC (nếu chọn quạt EC thì bỏ biến tần, giữ tín hiệu 0–10 V) | WP1: “quạt plug EC/biến tần” |
| 16 | Bản vẽ điện Phase 1 | Thư mục `00_inputs/phase1_dwg` chỉ có bản vẽ AC (không có F3-SD-EE-002/011/012). Tham khảo được: máng cáp 400×200 BOP FFL3600, “MÀN HÌNH HMI” (F3-SD-AC-017), DB-27…30 | phase1_dwg/*.txt |


## 4. Phụ tải & DB.HVAC

CDU **cấp trực tiếp từ DB.HVAC** (6 lộ MCCB 4P + ELCB 300 mA theo hãng), cầu dao cách ly 4P IP65 ngoài trời tại mỗi module. Lý do: tải lớn nhất (90 kW), ở ngoài nhà, được AHU điều khiển qua DX-kit bằng tín hiệu (enable/0–10 V/trạng thái/lỗi) – không cần đóng cắt nguồn trong tủ DDC; tủ DDC nhỏ gọn, dễ chọn lọc. Scrubber: HVAC cấp nguồn tới cầu dao cách ly + tủ điều khiển trọn bộ của nhà cung cấp; DDC-FAN trao đổi lệnh chạy/trạng thái/lỗi.

| Chỉ tiêu | Giá trị |
|---|---|
| Tổng công suất đặt (định mức, không gồm ổ cắm/điều khiển) | 191,7 kW (WP1: 191,7 kW) |
| Tổng công suất điện vào (kể hiệu suất) / S đặt | 205,2 kW / 228,1 kVA (347 A) |
| Kịch bản: Hè (lạnh 100 %, sưởi 0) | 166,8 kW / 189,7 kVA / 288 A |
| Kịch bản: Đông (sưởi 100 %, lạnh chỉ CDU-2.03 50 %) | 125,7 kW / 135,1 kVA / 205 A |
| Kịch bản: Nồm/khử ẩm (lạnh 70 % + sưởi gia nhiệt lại 100 % + hút ẩm) | 175,8 kW / 194,0 kVA / 295 A |
| **Nhu cầu lớn nhất (thiết kế)** | **175,8 kW / 194,0 kVA / Ib = 295 A** (Nồm/khử ẩm (lạnh 70 % + sưởi gia nhiệt lại 100 % + hút ẩm)) |
| Hệ số nhu cầu chung | 0,86 |
| MCCB tổng | MCCB 4P 400 AF, Ir chỉnh 340 A (0,4–1,0 In), Icu ≥ 36 kA (giả định – gói điện xác nhận Ik) |
| Thanh cái | Thanh cái Cu 400 A 3P+N+PE (≈ 40×10 mm/pha, N = pha, PE 40×5), Icw ≥ 25 kA/1 s |
| Cáp cấp nguồn khuyến nghị (thông tin cho gói điện) | 1 × CXV 4×240 mm² + 1 × CV 1×120 mm² (PE) |
| Ghi chú cáp nguồn | Thông tin cho gói điện: Iz = 1 × 538 × 0.87 × 0.80 = 374 A ≥ Ir 340 A; chiều dài/tuyến, Ik, sụt áp tới DB.HVAC (khuyến nghị ≤ 1 %) do gói điện tính. |
| Lộ ra DB.HVAC | 11 lộ dùng + 3 lộ dự phòng (MCCB 4P 63 A, 25 kA (lắp sẵn)) |


Hệ số đồng thời theo kịch bản (Hè / Đông / Nồm): CDU-2.01, 2.02: 1 / 0 / 0,7; CDU-2.03 (Instrument tải thiết bị lớn, chạy cả mùa đông): 1 / 0,5 / 0,7; sưởi: 0 / 1 / 1 (sưởi gia nhiệt lại khi khử ẩm); DH-2.01: 1 / 0,5 / 1; quạt, quạt hood, scrubber, bơm, điều khiển: 1; ổ cắm bảo trì 0,2. Kịch bản Nồm (khử ẩm + gia nhiệt lại) là bất lợi nhất; sưởi và lạnh không cộng 100 %/100 %.


### 4.1 Lộ ra DB.HVAC và cáp lộ tổng tủ DDC

| Lộ | Phụ tải | Ib (A) | Bảo vệ | Cáp | L (m) | ΔU (%) |
|---|---|---|---|---|---|---|
| DB-CDU-2.01-1 | Dàn nóng CDU-2.01-1 (inverter, 380V/3P+N) | 29,6 | MCCB 4P 40 A, Icu ≥ 25 kA | CXV 4×10 + E10 | 37,3 | 0,98 |
| DB-CDU-2.01-2 | Dàn nóng CDU-2.01-2 (inverter, 380V/3P+N) | 29,6 | MCCB 4P 40 A, Icu ≥ 25 kA | CXV 4×10 + E10 | 38,4 | 1,01 |
| DB-CDU-2.01-3 | Dàn nóng CDU-2.01-3 (inverter, 380V/3P+N) | 29,6 | MCCB 4P 40 A, Icu ≥ 25 kA | CXV 4×10 + E10 | 39,6 | 1,04 |
| DB-CDU-2.02 | Dàn nóng CDU-2.02 (inverter, 380V/3P+N) | 24,0 | MCCB 4P 32 A, Icu ≥ 25 kA | CXV 4×6 + E6 | 40,7 | 1,44 |
| DB-CDU-2.03-1 | Dàn nóng CDU-2.03-1 (inverter, 380V/3P+N) | 24,0 | MCCB 4P 32 A, Icu ≥ 25 kA | CXV 4×6 + E6 | 41,9 | 1,48 |
| DB-CDU-2.03-2 | Dàn nóng CDU-2.03-2 (inverter, 380V/3P+N) | 24,0 | MCCB 4P 32 A, Icu ≥ 25 kA | CXV 4×6 + E6 | 43,0 | 1,52 |
| DB-SCR-2.01 | Tủ điều khiển trọn bộ scrubber SCR-2.01 (quạt 5,5 + bơm 3 +  | 18,3 | MCCB 4P 25 A, Icu ≥ 25 kA | CXV 4×4 + E4 | 38,6 | 1,55 |
| DB-DDC-AHU-2.01 | Cáp lộ tổng DB.HVAC → DDC-AHU-2.01 | 56,2 | MCCB 4P 80 A, Icu ≥ 25 kA | CXV 4×25 + E16 | 8,5 | 0,18 |
| DB-DDC-AHU-2.02 | Cáp lộ tổng DB.HVAC → DDC-AHU-2.02 | 30,4 | MCCB 4P 40 A, Icu ≥ 25 kA | CXV 4×10 + E10 | 10,2 | 0,29 |
| DB-DDC-AHU-2.03 | Cáp lộ tổng DB.HVAC → DDC-AHU-2.03 | 41,2 | MCCB 4P 63 A, Icu ≥ 25 kA | CXV 4×16 + E16 | 14,6 | 0,36 |
| DB-DDC-FAN | Cáp lộ tổng DB.HVAC → DDC-FAN | 60,3 | MCCB 4P 80 A, Icu ≥ 25 kA | CXV 4×25 + E16 | 8,2 | 0,19 |


### 4.2 Mạch cuối trong tủ DDC

| Mạch | Phụ tải | Ib / I cáp (A) | Bảo vệ | Dòng rò | Cáp | L (m) | ΔU tích lũy (%) | Cách ly |
|---|---|---|---|---|---|---|---|---|
| DDC-AHU-2.01-M1 | Quạt cấp AHU-2.01 15,0 kW qua biến tần (cáp màn chắn) | 26,9 / 29,1 | MCB 3P 40 A (C) 10 kA + biến tần 15,0 kW | – | CXV/S (màn chắn Cu, cáp biến tần) 3×10 + E10 | 7,2 | 0,37 | 40 A |
| DDC-AHU-2.01-H1 | Sưởi điện AHU-2.01 cấp 1/3 – 5,0 kW (SSR/thyristor điều bi | 7,6 / 7,6 | MCB 3P 10 A (C) + SSR/thyristor 3P | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 7,2 | 0,41 | 20 A |
| DDC-AHU-2.01-H2 | Sưởi điện AHU-2.01 cấp 2/3 – 5,0 kW (contactor) | 7,6 / 7,6 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 7,2 | 0,41 | 20 A |
| DDC-AHU-2.01-H3 | Sưởi điện AHU-2.01 cấp 3/3 – 5,0 kW (contactor) | 7,6 / 7,6 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 7,2 | 0,41 | 20 A |
| DDC-AHU-2.01-CP | Bơm nước ngưng CP-2.02 (220V) | 0,6 / 0,6 | RCBO 1P+N 6 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 8,8 | 0,22 | 20 A |
| DDC-AHU-2.01-SK | Ổ cắm bảo trì 1P 16A IP66 tại sàn AHU-2.01 | 4,5 / 4,5 | RCBO 1P+N 16 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 9,9 | 0,55 | – |
| DDC-AHU-2.01-CT | Nguồn điều khiển (DDC, 24 VAC/DC, quạt tủ) | 1,4 / 1,4 | MCB 1P+N 6 A (C), 10 kA | – | (trong tủ) | – | 0,18 | – |
| DDC-AHU-2.02-M1 | Quạt cấp AHU-2.02 3,0 kW qua biến tần (cáp màn chắn) | 5,6 / 6,1 | MCB 3P 10 A (C) 10 kA + biến tần 3,0 kW | – | CXV/S (màn chắn Cu, cáp biến tần) 3×2,5 + E2,5 | 28,4 | 0,90 | 20 A |
| DDC-AHU-2.02-H1 | Sưởi điện AHU-2.02 cấp 1/3 – 4,0 kW (SSR/thyristor điều bi | 6,1 / 6,1 | MCB 3P 10 A (C) + SSR/thyristor 3P | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 28,4 | 1,00 | 20 A |
| DDC-AHU-2.02-H2 | Sưởi điện AHU-2.02 cấp 2/3 – 4,0 kW (contactor) | 6,1 / 6,1 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 28,4 | 1,00 | 20 A |
| DDC-AHU-2.02-H3 | Sưởi điện AHU-2.02 cấp 3/3 – 4,0 kW (contactor) | 6,1 / 6,1 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 28,4 | 1,00 | 20 A |
| DDC-AHU-2.02-CP | Bơm nước ngưng CP-2.03 (220V) | 0,6 / 0,6 | RCBO 1P+N 6 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 30,1 | 0,42 | 20 A |
| DDC-AHU-2.02-SK | Ổ cắm bảo trì 1P 16A IP66 tại sàn AHU-2.02 | 4,5 / 4,5 | RCBO 1P+N 16 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 31,2 | 1,45 | – |
| DDC-AHU-2.02-CT | Nguồn điều khiển (DDC, 24 VAC/DC, quạt tủ) | 1,4 / 1,4 | MCB 1P+N 6 A (C), 10 kA | – | (trong tủ) | – | 0,29 | – |
| DDC-AHU-2.03-M1 | Quạt cấp AHU-2.03 7,5 kW qua biến tần (cáp màn chắn) | 13,7 / 14,8 | MCB 3P 20 A (C) 10 kA + biến tần 7,5 kW | – | CXV/S (màn chắn Cu, cáp biến tần) 3×2,5 + E2,5 | 16,0 | 1,19 | 20 A |
| DDC-AHU-2.03-H1 | Sưởi điện AHU-2.03 cấp 1/3 – 3,0 kW (SSR/thyristor điều bi | 4,6 / 4,6 | MCB 3P 10 A (C) + SSR/thyristor 3P | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 16,0 | 0,66 | 20 A |
| DDC-AHU-2.03-H2 | Sưởi điện AHU-2.03 cấp 2/3 – 3,0 kW (contactor) | 4,6 / 4,6 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 16,0 | 0,66 | 20 A |
| DDC-AHU-2.03-H3 | Sưởi điện AHU-2.03 cấp 3/3 – 3,0 kW (contactor) | 4,6 / 4,6 | MCB 3P 10 A (C) + contactor AC1 | RCD 300 mA 4P chung nhóm sưởi | CXV 3×2,5 + E2,5 | 16,0 | 0,66 | 20 A |
| DDC-AHU-2.03-CP | Bơm nước ngưng CP-2.01 (220V) | 0,6 / 0,6 | RCBO 1P+N 6 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 16,5 | 0,43 | 20 A |
| DDC-AHU-2.03-DH | Máy hút ẩm DH-2.01 (220V, bơm ngưng tích hợp) | 7,3 / 7,3 | RCBO 1P+N 10 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 12,3 | 0,98 | 20 A |
| DDC-AHU-2.03-SK | Ổ cắm bảo trì 1P 16A IP66 tại sàn AHU-2.03 | 4,5 / 4,5 | RCBO 1P+N 16 A 30 mA (C) | RCBO 30 mA | CXV 2×2,5 + E2,5 | 16,9 | 0,99 | – |
| DDC-AHU-2.03-CT | Nguồn điều khiển (DDC, 24 VAC/DC, quạt tủ) | 1,4 / 1,4 | MCB 1P+N 6 A (C), 10 kA | – | (trong tủ) | – | 0,36 | – |
| DDC-FAN-FAF-2.01 | Quạt cấp bù FAF-2.01 7,5 kW qua biến tần (cáp màn chắn) | 13,7 / 14,8 | MCB 3P 20 A (C) 10 kA + biến tần 7,5 kW | – | CXV/S (màn chắn Cu, cáp biến tần) 3×2,5 + E2,5 | 18,7 | 1,16 | 20 A |
| DDC-FAN-FAF-2.02 | Quạt cấp bù FAF-2.02 11,0 kW qua biến tần (cáp màn chắn) | 19,8 / 21,5 | MCB 3P 25 A (C) 10 kA + biến tần 11,0 kW | – | CXV/S (màn chắn Cu, cáp biến tần) 3×4 + E4 | 18,2 | 1,05 | 25 A |
| DDC-FAN-EAF-2.01 | Quạt hút EAF-2.01 0,75 kW – khởi động trực tiếp (MPCB + co | 1,6 / 1,6 | MPCB 3P 1,6–2,5 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 24,5 | 0,33 | 20 A |
| DDC-FAN-EAF-2.02 | Quạt hút EAF-2.02 1,10 kW – khởi động trực tiếp (MPCB + co | 2,3 / 2,3 | MPCB 3P 1,6–2,5 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 23,9 | 0,39 | 20 A |
| DDC-FAN-EAF-2.03 | Quạt hút EAF-2.03 0,75 kW – khởi động trực tiếp (MPCB + co | 1,6 / 1,6 | MPCB 3P 1,6–2,5 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 17,3 | 0,29 | 20 A |
| DDC-FAN-HF01 | Quạt tủ hood DIG-FH1 (FHF-2.01, ACID) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 34,9 | 0,34 | 20 A |
| DDC-FAN-HF02 | Quạt tủ hood DIG-FH2 (FHF-2.01, ACID) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 36,2 | 0,35 | 20 A |
| DDC-FAN-HF03 | Quạt tủ hood ST1-FH1 (FHF-2.02, ACID) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 30,2 | 0,32 | 20 A |
| DDC-FAN-HF04 | Quạt tủ hood ST1-FH2 (FHF-2.02) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 28,4 | 0,32 | 20 A |
| DDC-FAN-HF05 | Quạt tủ hood ST1-FH3 (FHF-2.02) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 27,4 | 0,31 | 20 A |
| DDC-FAN-HF06 | Quạt tủ hood ST2-FH1 (FHF-2.03) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 23,4 | 0,29 | 20 A |
| DDC-FAN-HF07 | Quạt tủ hood ST2-FH2 (FHF-2.03) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 25,2 | 0,30 | 20 A |
| DDC-FAN-HF08 | Quạt tủ hood ST2-FH3 (FHF-2.03, ACID) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 27,1 | 0,31 | 20 A |
| DDC-FAN-HF09 | Quạt tủ hood ST2-FH4 (FHF-2.03) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 20,2 | 0,28 | 20 A |
| DDC-FAN-HF10 | Quạt tủ hood ST2-FH5 (FHF-2.03) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 16,4 | 0,26 | 20 A |
| DDC-FAN-HF11 | Quạt tủ hood ST2-FH6 (FHF-2.03) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 18,2 | 0,27 | 20 A |
| DDC-FAN-HF12 | Quạt tủ hood STD20-FH1 (FHF-2.04) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 12,8 | 0,25 | 20 A |
| DDC-FAN-HF13 | Quạt tủ hood STD20-FH2 (FHF-2.04) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 14,6 | 0,26 | 20 A |
| DDC-FAN-HF14 | Quạt tủ hood STD20-FH3 (FHF-2.04) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 18,8 | 0,27 | 20 A |
| DDC-FAN-HF15 | Quạt tủ hood STD20-FH4 (FHF-2.04) 0,6 kW | 1,3 / 1,3 | MPCB 3P 1–1,6 A + contactor AC3 | – | CXV 3×2,5 + E2,5 | 19,6 | 0,28 | 20 A |
| DDC-FAN-CT | Nguồn điều khiển + biến áp 24 VAC cho 27 bộ chấp hành van  | 2,3 / 2,3 | MCB 1P+N 6 A (C), 10 kA | – | (trong tủ) | – | 0,19 | – |


Kiểm tra (Excel, sau khi LibreOffice tính lại): 0 mạch không đạt; ΔU tích lũy lớn nhất 1,55 % (≤ 4 %); độ đầy máng lớn nhất 37,5 % (≤ 40 %).

## 5. Tủ & thành phần chính

**DB.HVAC** – 1200W × 2100H × 500D, Đặt sàn (đế 100 mm), IP42.

- Vỏ tủ thép 1,5–2,0 mm sơn tĩnh điện, 1200W × 2100H × 500D, IP42, Form 2b, cửa có khóa × 1 bộ
- MCCB 4P 400 AF, Ir chỉnh 340 A (0,4–1,0 In), Icu ≥ 36 kA (giả định – gói điện xác nhận Ik) × 1 cái
- Thanh cái Cu 400 A 3P+N+PE (≈ 40×10 mm/pha, N = pha, PE 40×5), Icw ≥ 25 kA/1 s × 1 bộ
- Đồng hồ đa năng V/A/kW/kWh/PF/THD, Modbus RTU + 3 CT 400/5 A × 1 bộ
- Thiết bị chống sét lan truyền SPD Type 2, 4P, Imax 40 kA + MCB 4P 32 A × 1 bộ
- Đèn báo pha + cầu chì bảo vệ × 3 bộ
- MCCB 4P 25 A, Icu ≥ 25 kA × 1 cái
- MCCB 4P 32 A, Icu ≥ 25 kA × 3 cái
- MCCB 4P 40 A, Icu ≥ 25 kA × 4 cái
- MCCB 4P 63 A, Icu ≥ 25 kA × 1 cái
- MCCB 4P 80 A, Icu ≥ 25 kA × 2 cái
- Rơ le dòng rò ELR 300 mA + ZCT (cho lộ CDU, theo hãng CDU) × 6 bộ
- Lộ dự phòng 20 %: MCCB 4P 63 A, 25 kA (lắp sẵn) × 3 cái
- Cầu đấu, máng nhựa đi dây, nhãn, sơ đồ trong tủ, phụ kiện × 1 lô



**DDC-FAN** – 1200W × 2100H × 500D, Đặt sàn (đế 100 mm), IP42.

- Vỏ tủ thép sơn tĩnh điện 1200W × 2100H × 500D, IP42, quạt hút + lọc gió mặt tủ × 1 bộ
- MCCB tổng 4P 80 A 25 kA (tủ) / cầu dao tổng × 1 cái
- Thanh cái Cu 100 A 3P+N+PE × 1 bộ
- Biến tần 7,5 kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho FAF-2.01 × 1 bộ
- Biến tần 11,0 kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho FAF-2.02 × 1 bộ
- MCB 3P 20 A (C) 10 kA + biến tần 7,5 kW × 1 cái
- MCB 3P 25 A (C) 10 kA + biến tần 11,0 kW × 1 cái
- MPCB 3P 1,6–2,5 A + contactor AC3 × 3 cái
- MPCB 3P 1–1,6 A + contactor AC3 × 15 cái
- Contactor AC3 3P 9 A cuộn 24 VAC (quạt EAF + quạt tủ hood) × 18 cái
- Rơ le dòng lõi xuyên cho quạt tủ hood × 15 cái
- Biến áp 220/24 VAC 630 VA cấp bộ chấp hành van gió + MCB bảo vệ × 1 bộ
- Rơ le trung gian 24 VAC (lệnh van gió) × 24 cái
- Bộ điều khiển DDC lập trình tự do BACnet + module mở rộng: ≥ 62 DI / 54 DO / 3 AI / 3 AO (đã gồm 20 % dự phòng) × 1 bộ
- Màn hình hiển thị/cài đặt cục bộ trên cửa tủ (LCD/HMI 7") × 1 bộ
- Công tắc H-O-A, nút dừng khẩn, đèn báo chạy/lỗi cho từng thiết bị × 1 lô
- Lộ dự phòng 20 %: MPCB 3P 1,6–2,5 A + contactor (lắp sẵn) × 4 bộ
- Cầu đấu, máng nhựa đi dây, nhãn, phụ kiện × 1 lô



**DDC-AHU-2.02** – 1000W × 1400H × 300D, Treo tường, đáy +0,5 m, IP54.

- Vỏ tủ thép sơn tĩnh điện 1000W × 1400H × 300D, IP54, quạt hút + lọc gió mặt tủ × 1 bộ
- MCCB tổng 4P 40 A 25 kA (tủ) / cầu dao tổng × 1 cái
- Thanh cái Cu 63 A 3P+N+PE × 1 bộ
- Biến tần 3,0 kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho AHU-2.02-SF × 1 bộ
- MCB 3P 10 A (C) + SSR/thyristor 3P × 1 cái
- MCB 3P 10 A (C) + contactor AC1 × 2 cái
- MCB 3P 10 A (C) 10 kA + biến tần 3,0 kW × 1 cái
- RCBO 1P+N 16 A 30 mA (C) × 1 cái
- RCBO 1P+N 6 A 30 mA (C) × 1 cái
- RCD 4P 40 A 300 mA cho nhóm sưởi × 1 cái
- Bộ điều khiển thyristor/SSR 3P (0–10 V) cho cấp sưởi 1 × 1 bộ
- Contactor AC1 3P cho cấp sưởi 2, 3 × 2 cái
- Rơ le an toàn liên động sưởi (quạt chạy + airflow + thermostat) + rơ le thời gian chạy quạt sau 3 phút × 1 bộ
- Biến áp 220/24 VAC 100 VA + nguồn 24 VDC 5 A cho DDC/cảm biến × 1 bộ
- Rơ le cắt cứng: báo cháy (STO/enable biến tần + sưởi), mức nước ngưng cao (cắt DX-kit) × 1 bộ
- Bộ điều khiển DDC lập trình tự do BACnet + module mở rộng: ≥ 15 DI / 6 DO / 18 AI / 4 AO (đã gồm 20 % dự phòng) × 1 bộ
- Màn hình hiển thị/cài đặt cục bộ trên cửa tủ (LCD/HMI 7") × 1 bộ
- Công tắc H-O-A, nút dừng khẩn, đèn báo chạy/lỗi cho từng thiết bị × 1 lô
- Lộ dự phòng 20 %: MCB 3P 16 A (C) + 1 MCB 1P 10 A × 2 bộ
- Cầu đấu, máng nhựa đi dây, nhãn, phụ kiện × 1 lô



**DDC-AHU-2.01** – 1000W × 1800H × 400D, Treo tường, đáy +0,3 m, IP54.

- Vỏ tủ thép sơn tĩnh điện 1000W × 1800H × 400D, IP54, quạt hút + lọc gió mặt tủ × 1 bộ
- MCCB tổng 4P 80 A 25 kA (tủ) / cầu dao tổng × 1 cái
- Thanh cái Cu 100 A 3P+N+PE × 1 bộ
- Biến tần 15,0 kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho AHU-2.01-SF × 1 bộ
- MCB 3P 10 A (C) + SSR/thyristor 3P × 1 cái
- MCB 3P 10 A (C) + contactor AC1 × 2 cái
- MCB 3P 40 A (C) 10 kA + biến tần 15,0 kW × 1 cái
- RCBO 1P+N 16 A 30 mA (C) × 1 cái
- RCBO 1P+N 6 A 30 mA (C) × 1 cái
- RCD 4P 40 A 300 mA cho nhóm sưởi × 1 cái
- Bộ điều khiển thyristor/SSR 3P (0–10 V) cho cấp sưởi 1 × 1 bộ
- Contactor AC1 3P cho cấp sưởi 2, 3 × 2 cái
- Rơ le an toàn liên động sưởi (quạt chạy + airflow + thermostat) + rơ le thời gian chạy quạt sau 3 phút × 1 bộ
- Biến áp 220/24 VAC 100 VA + nguồn 24 VDC 5 A cho DDC/cảm biến × 1 bộ
- Rơ le cắt cứng: báo cháy (STO/enable biến tần + sưởi), mức nước ngưng cao (cắt DX-kit) × 1 bộ
- Bộ điều khiển DDC lập trình tự do BACnet + module mở rộng: ≥ 20 DI / 9 DO / 26 AI / 6 AO (đã gồm 20 % dự phòng) × 1 bộ
- Màn hình hiển thị/cài đặt cục bộ trên cửa tủ (LCD/HMI 7") × 1 bộ
- Công tắc H-O-A, nút dừng khẩn, đèn báo chạy/lỗi cho từng thiết bị × 1 lô
- Lộ dự phòng 20 %: MCB 3P 16 A (C) + 1 MCB 1P 10 A × 2 bộ
- Cầu đấu, máng nhựa đi dây, nhãn, phụ kiện × 1 lô



**DDC-AHU-2.03** – 1000W × 1600H × 400D, Treo tường, đáy +0,3 m, IP54.

- Vỏ tủ thép sơn tĩnh điện 1000W × 1600H × 400D, IP54, quạt hút + lọc gió mặt tủ × 1 bộ
- MCCB tổng 4P 63 A 25 kA (tủ) / cầu dao tổng × 1 cái
- Thanh cái Cu 63 A 3P+N+PE × 1 bộ
- Biến tần 7,5 kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho AHU-2.03-SF × 1 bộ
- MCB 3P 10 A (C) + SSR/thyristor 3P × 1 cái
- MCB 3P 10 A (C) + contactor AC1 × 2 cái
- MCB 3P 20 A (C) 10 kA + biến tần 7,5 kW × 1 cái
- RCBO 1P+N 10 A 30 mA (C) × 1 cái
- RCBO 1P+N 16 A 30 mA (C) × 1 cái
- RCBO 1P+N 6 A 30 mA (C) × 1 cái
- RCD 4P 40 A 300 mA cho nhóm sưởi × 1 cái
- Bộ điều khiển thyristor/SSR 3P (0–10 V) cho cấp sưởi 1 × 1 bộ
- Contactor AC1 3P cho cấp sưởi 2, 3 × 2 cái
- Rơ le an toàn liên động sưởi (quạt chạy + airflow + thermostat) + rơ le thời gian chạy quạt sau 3 phút × 1 bộ
- Biến áp 220/24 VAC 100 VA + nguồn 24 VDC 5 A cho DDC/cảm biến × 1 bộ
- Rơ le cắt cứng: báo cháy (STO/enable biến tần + sưởi), mức nước ngưng cao (cắt DX-kit) × 1 bộ
- Bộ điều khiển DDC lập trình tự do BACnet + module mở rộng: ≥ 20 DI / 9 DO / 11 AI / 5 AO (đã gồm 20 % dự phòng) × 1 bộ
- Màn hình hiển thị/cài đặt cục bộ trên cửa tủ (LCD/HMI 7") × 1 bộ
- Công tắc H-O-A, nút dừng khẩn, đèn báo chạy/lỗi cho từng thiết bị × 1 lô
- Lộ dự phòng 20 %: MCB 3P 16 A (C) + 1 MCB 1P 10 A × 2 bộ
- Cầu đấu, máng nhựa đi dây, nhãn, phụ kiện × 1 lô



## 6. I/O DDC & cảm biến

| Tủ | DI | DO | AI | AO | Tổng | Chọn DDC (+20 %) |
|---|---|---|---|---|---|---|
| DDC-AHU-2.01 | 16 | 7 | 21 | 5 | 49 | 20/9/26/6 |
| DDC-AHU-2.02 | 12 | 5 | 15 | 3 | 35 | 15/6/18/4 |
| DDC-AHU-2.03 | 16 | 7 | 9 | 4 | 36 | 20/9/11/5 |
| DDC-FAN | 51 | 45 | 2 | 2 | 100 | 62/54/3/3 |
| **Tổng** | 95 | 64 | 47 | 14 | 220 |  |


Chi tiết từng điểm (mô tả, thiết bị, vị trí, cáp, chiều dài) trong sheet `IO` và `electrical_design.json → io_schedule`. Liên động an toàn được **đi dây cứng** (không phụ thuộc DDC): báo cháy → STO/enable biến tần + contactor sưởi; airflow switch + thermostat an toàn + quạt chạy → mạch cuộn hút sưởi; phao mức cao bơm ngưng → DX-kit.

| Cảm biến / thiết bị hiện trường | Quy cách | SL |
|---|---|---|
| Cảm biến nhiệt độ/độ ẩm phòng gắn tường | 0–50 °C ±0,3 K; 0–100 %RH ±2 %; ra 0–10 V; 24 VAC/DC; vỏ chịu hóa chất cho phòng hóa | 14 |
| Cảm biến nhiệt độ/độ ẩm ống gió | Que 200 mm, ±0,3 K / ±2 %RH, 0–10 V (gió cấp + gió hồi mỗi AHU) | 6 |
| Cảm biến nhiệt độ/độ ẩm ngoài trời | IP65, có hộp che nắng, 0–10 V | 1 |
| Bộ chuyển đổi chênh áp quạt AHU (đo lưu lượng k-factor) | 0–1000 Pa, 0–10 V, kèm ống mềm & đầu đo | 3 |
| Cảm biến áp suất tĩnh ống gió bù FAF | 0–500 Pa, 0–10 V, đầu đo tĩnh áp | 2 |
| Công tắc chênh áp lọc | 50–500 Pa, tiếp điểm SPDT (G4 + F8 mỗi AHU/FAF) | 10 |
| Công tắc dòng gió (airflow switch) | 20–300 Pa, SPDT (3 AHU – liên động sưởi; 3 EAF – xác nhận hút) | 6 |
| Thermostat an toàn sưởi reset tay | Kèm cụm sưởi AHU – WP4 chỉ đấu nối (không tính tiền nếu AHU đã kèm) | 3 |
| Bộ chấp hành van gió on/off lò xo hồi 24 VAC | ≥ 10 Nm, thời gian chạy ≤ 90 s; van gió bù 15 hood + 7 nhóm EAF + 3 OA AHU + 2 OA FAF (thân van thuộc WP3) | 27 |
| Hộp công tắc tại tủ hood | Công tắc ON/OFF + đèn báo chạy, 24 VDC, IP54, gắn trên tủ hood | 15 |
| Rơ le dòng (current sensing relay) quạt tủ hood | Lõi xuyên, 0,5–5 A, tiếp điểm khô (lắp trong DDC-FAN) | 15 |


Van gió bù nhóm theo quạt EAF: MD-G-DIG ← EAF-2.01 (DIG-CH1); MD-G-KJE ← EAF-2.01 (KJE-CH1); MD-G-ST1 ← EAF-2.01 (ST1-CC1, ST1-CC2); MD-G-PRE ← EAF-2.03 (PRE-AH1, PRE-AH2); MD-G-ST2 ← EAF-2.02 (ST2-CC1, ST2-CC2, ST2-SK1, ST2-SK2); MD-G-STD20 ← EAF-2.02 (STD20-SK1); MD-G-STD12 ← EAF-2.02 (STD12-CC1).


## 7. Trình tự điều khiển (tóm tắt)

- **AHU – khởi động**: Lệnh chạy (lịch/BMS/H-O-A) → mở van OA → trễ 60 s → biến tần chạy, giữ lưu lượng gió cấp theo chênh áp quạt (k-factor, PID). AHU-2.03 chạy 24/7. Quạt mất xác nhận (biến tần/airflow) → báo lỗi, khóa sưởi và CDU.
- **AHU – làm lạnh**: Điều khiển theo nhiệt độ hồi (giới hạn bởi giá trị cao nhất của cảm biến phòng trong vùng): yêu cầu công suất 0–10 V tới DX-kit, phân cấp module CDU (lead/lag, luân phiên 24 h, chống chu kỳ ngắn 5 phút).
- **AHU – khử ẩm**: RH phòng > 60 % (Instrument > 55 %) → ưu tiên khử ẩm: CDU chạy để T sau coil ≈ 15 °C; nhiệt độ phòng giữ bằng sưởi gia nhiệt lại (SSR điều biến + 2 cấp contactor).
- **AHU – sưởi (đông)**: T phòng < 20 °C → sưởi theo PID (SSR cấp 1 điều biến, cấp 2–3 đóng tuần tự). Liên động cứng: chỉ cấp điện sưởi khi biến tần chạy + airflow switch đóng + thermostat an toàn bình thường; sau khi tắt sưởi quạt chạy thêm 3 phút. Không chạy sưởi và lạnh ở chế độ thường (chỉ đồng thời khi khử ẩm).
- **AHU – lọc, nước ngưng**: Công tắc ΔP lọc G4/F8 → báo bẩn. Phao mức cao bơm nước ngưng → cắt cứng DX-kit (CDU), báo động; quạt vẫn chạy.
- **Máy hút ẩm DH-2.01**: DDC-AHU-2.03 cho phép DH chạy khi RH Instrument > 55 % (trễ 5 %), hoặc khi AHU-2.03 ở non tải không giữ được RH; giám sát chạy/lỗi/bồn đầy.
- **Hút tủ hood ↔ gió bù**: Người dùng bật công tắc tại tủ → DDC mở van gió bù MD-<hood> → trễ 30 s → đóng contactor quạt tủ → rơ le dòng xác nhận → đèn tại tủ. Tắt: dừng quạt, đóng van sau 30 s.
- **Hood acid ↔ scrubber**: 4 hood acid (DIG-FH1/FH2, ST1-FH1, ST2-FH3): yêu cầu chạy → DO khởi động scrubber → chỉ đóng contactor quạt hood khi scrubber báo chạy; scrubber chạy thêm 10 phút sau hood acid cuối cùng tắt. Lỗi scrubber → báo động tại tủ + DDC (quạt hood giữ chạy, người dùng hạ cửa tủ).
- **EAF ↔ van bù nhóm**: EAF-2.01/2.02/2.03 chạy theo lịch (giờ làm việc; EAF-2.03 theo Instrument 24/7) → mở van bù nhóm MD-G-xx tương ứng; xác nhận dòng gió bằng công tắc ΔP. Instrument: thải bù bằng gió tươi AHU-2.03 (van OA mở khi EAF-2.03 chạy).
- **FAF (gió bù)**: FAF-2.01/2.02 được phép chạy khi có ≥ 1 van bù mở; biến tần PID giữ áp suất tĩnh ống gió bù (đặt khi T&C, ~250 Pa) nên lưu lượng tự theo số van mở; van OA cửa hút mở trước khi quạt chạy. Lỗi FAF → chỉ báo động (không dừng quạt hút).
- **Báo cháy**: Tín hiệu từ module rơ le báo cháy (gói PCCC): dừng 3 AHU (cắt cứng enable/STO biến tần + sưởi), dừng FAF, đóng van OA và van bù. Quạt hút EAF/hood/scrubber GIỮ chạy để không thoát hơi hóa chất ra phòng – cần tư vấn PCCC/CĐT xác nhận.
- **Chế độ mùa**: Cảm biến T/RH ngoài trời (DDC-AHU-2.02, chia sẻ qua mạng) chọn chế độ Hè/Đông/Nồm; mạng RS485 BACnet MS/TP nối 4 tủ DDC + đồng hồ DB.HVAC.


## 8. Máng cáp, ống, cách ly

Tuyến (xem `WP4_containment.png`): máng chính dọc lối đi X = 15.285 (Y −4.800…16.000) nối 5 tủ; khối trái: nhánh AHU-2.01 (Y 5.150), nhánh bắc Y 11.000 → FAF-2.01, hood ST1/Digestion (X 7.000), nhánh EAF-2.01/Kjeldahl (Y 16.000), nhánh nam Y −1.200 → AHU-2.03, DH-2.01, EAF-2.03, Sample prep/Coding/Dividing; khối phải: máng chính Y 2.200 qua Office → ST2 → dọc tường A (X 27.300), nhánh Standard 19,72 (X 18.000), nhánh hood ST2 (X 21.700 và Y 9.000). Ngoài nhà: xuyên tường A tại Y −5.500 (ống lồng) → máng HDG có nắp ốp tường +3,6 m vượt ram dốc → máng đứng xuống bãi CDU → máng HDG có nắp dọc bãi, ống GI ren tới từng cầu dao cách ly; phía bắc: xuyên tường A tại Y 9.300 → máng HDG có nắp ốp tường tới scrubber (Y 15.500) → ống GI xuống tủ scrubber.

| Loại máng | Chiều dài (m) |
|---|---|
| Máng cáp đục lỗ sơn tĩnh điện 200×100 dày 1,2 mm | 56,9 |
| Máng cáp đục lỗ sơn tĩnh điện 100×50 dày 1,2 mm | 157,9 |
| Máng cáp đục lỗ mạ kẽm nhúng nóng 200×100 có nắp (ngoài trời) | 18,9 |
| Máng cáp đục lỗ mạ kẽm nhúng nóng 100×50 có nắp (ngoài trời) | 18,9 |
| **Tổng** | 252,6 |


| Phụ kiện máng | SL |
|---|---|
| Co ngang 90° | 4 |
| Tê ngang (T) | 9 |
| Chữ thập (+) | 2 |
| Nắp bịt đầu máng | 12 |
| Co đứng 90° (máng đứng nóc tủ) | 7 |
| Nối máng (khớp nối + bu lông), mỗi 2,4 m | 106 |
| Giá treo trong nhà (ty M10 ×2 + U-channel 41×41, mỗi 1,5 m) | 144 |
| Giá đỡ ngoài trời ốp tường/bệ (mạ kẽm nhúng nóng, mỗi 1,5 m) | 26 |
| Ống lồng xuyên tường A Ø100 thép mạ kẽm + chèn chống cháy/chống thấm | 4 |


| Ống luồn / ống mềm | m |
|---|---|
| Ống nhựa PVC cứng Ø20 (cáp điều khiển/cảm biến) | 326 |
| Ống thép EMT Ø20 (đoạn xuống thiết bị trên trần) | 15 |
| Ống thép EMT Ø25 (đoạn xuống thiết bị trên trần) | 35 |
| Ống thép EMT Ø32 (đoạn xuống thiết bị trên trần) | 2 |
| Ống thép mạ kẽm (GI/IMC) ren Ø20 ngoài trời | 17 |
| Ống thép mạ kẽm (GI/IMC) ren Ø25 ngoài trời | 4 |
| Ống thép mạ kẽm (GI/IMC) ren Ø32 ngoài trời | 5 |
| Ống mềm kim loại bọc PVC kín nước Ø20 + đầu nối | 4 |
| Ống mềm kim loại bọc PVC kín nước Ø25 + đầu nối | 31 |
| Ống mềm kim loại bọc PVC kín nước Ø32 + đầu nối | 7 |
| Ống mềm kim loại bọc PVC kín nước Ø40 + đầu nối | 1 |


| Cầu dao cách ly tại chỗ | SL |
|---|---|
| Cầu dao cách ly tại chỗ 2P 20A IP65 | 4 |
| Cầu dao cách ly tại chỗ 3P 20A IP65 | 27 |
| Cầu dao cách ly tại chỗ 3P 20A IP65 + tiếp điểm phụ (báo biến tần) | 3 |
| Cầu dao cách ly tại chỗ 3P 25A IP65 + tiếp điểm phụ (báo biến tần) | 1 |
| Cầu dao cách ly tại chỗ 3P 40A IP65 + tiếp điểm phụ (báo biến tần) | 1 |
| Cầu dao cách ly tại chỗ 4P 25A IP65 ngoài trời (chống UV) | 1 |
| Cầu dao cách ly tại chỗ 4P 32A IP65 ngoài trời (chống UV) | 3 |
| Cầu dao cách ly tại chỗ 4P 40A IP65 ngoài trời (chống UV) | 3 |


| Ốc siết cáp | SL |
|---|---|
| Ốc siết cáp nhựa nylon IP68 M20 | 398 |
| Ốc siết cáp nhựa nylon IP68 M32 | 4 |
| Ốc siết cáp nhựa nylon IP68 M40 | 4 |
| Ốc siết cáp đồng mạ niken EMC M25 | 16 |
| Ốc siết cáp đồng mạ niken EMC M32 | 4 |
| Ốc siết cáp đồng mạ niken IP68 M25 | 16 |
| Ốc siết cáp đồng mạ niken IP68 M32 | 12 |


## 9. Khối lượng cáp

| Cáp động lực | m |
|---|---|
| CXV 2×2,5 + E2,5 | 126 |
| CXV 3×2,5 + E2,5 | 574 |
| CXV 4×4 + E4 | 39 |
| CXV 4×6 + E6 | 126 |
| CXV 4×10 + E10 | 125 |
| CXV 4×16 + E16 | 15 |
| CXV 4×25 + E16 | 17 |
| CXV/S (màn chắn Cu, cáp biến tần) 3×2,5 + E2,5 | 63 |
| CXV/S (màn chắn Cu, cáp biến tần) 3×4 + E4 | 18 |
| CXV/S (màn chắn Cu, cáp biến tần) 3×10 + E10 | 7 |
| **Tổng** | 1.109 |


| Cáp điều khiển / tín hiệu / mạng | m |
|---|---|
| Cáp tín hiệu chống nhiễu 3C×0,75 mm² | 95 |
| Cáp ĐK chống nhiễu 2C×1,0 mm² | 869 |
| Cáp tín hiệu chống nhiễu 4C×0,75 mm² | 449 |
| Cáp ĐK chống nhiễu 3C×1,0 mm² | 597 |
| Cáp tín hiệu chống nhiễu 2C×0,75 mm² | 95 |
| Cáp chống cháy FR 2C×1,5 mm² | 42 |
| Cáp ĐK chống nhiễu 4C×1,0 mm² | 353 |
| Cáp mạng RS485 1 đôi xoắn chống nhiễu 0,75 mm² (BACnet MS/TP/Modbus) | 45 |
| Cáp truyền tín hiệu DX-kit ↔ CDU 2C×1,25 mm² chống nhiễu | 258 |
| **Tổng** | 2.803 |


## 10. Tiếp địa

PE đi chung trong mỗi cáp (lõi PE theo §3 #10); thanh PE mỗi tủ DDC nối về DB.HVAC qua lõi PE cáp lộ tổng; DB.HVAC nối thanh tiếp địa chính tòa nhà (MET). Máng cáp động lực có dây tiếp địa CV 1×25 mm² chạy suốt + cầu nối CV 6 mm² tại mỗi mối nối; vỏ CDU/scrubber nối đẳng thế CV 16 mm².

| Hạng mục | ĐVT | SL |
|---|---|---|
| Dây tiếp địa dọc máng cáp động lực CV 1×25 mm² vàng-xanh | m | 130 |
| Dây nối đẳng thế mối nối máng CV 1×6 mm² + đầu cos | bộ | 106 |
| Dây nối đất vỏ CDU/scrubber/khung bệ CV 1×16 mm² (3 m/thiết bị) | m | 21 |
| Đấu nối thanh PE DB.HVAC → thanh tiếp địa chính tòa nhà (MET) CV 1×95 mm² (giả định 10 m) | m | 10 |
| Kẹp tiếp địa/đầu cos đồng các loại | bộ | 60 |


## 11. Phương án chọn thêm – kết nối BMS/HMI Phase 1 (không cộng vào tổng)

| Mã | Hạng mục | ĐVT | SL |
|---|---|---|---|
| OP-1 | Bộ điều khiển mạng (network controller/BACnet router) BACnet MS/TP ↔ BACnet/IP, lắp trong DDC-FAN | bộ | 1 |
| OP-2 | Màn hình HMI cảm ứng 15,6" (đồ họa 3 AHU, FAF/EAF, hood, scrubber, báo động, xu hướng) đặt tại Office hoặc phòng điều khiển Phase 1 | bộ | 1 |
| OP-3 | Phần mềm + lập trình đồ họa, tích hợp vào HMI Phase 1 ("MÀN HÌNH HMI" – F3-SD-AC-017) qua BACnet/IP hoặc Modbus TCP | lô | 1 |
| OP-4 | Gateway BACnet/Modbus cho CDU (giám sát chi tiết lỗi/công suất từ hãng CDU) | bộ | 3 |
| OP-5 | Cáp mạng Cat6 FTP (DDC-FAN → HMI Phase 1, giả định 80 m) + ống PVC Ø20 + switch công nghiệp 8 cổng | lô | 1 |


## 12. Vấn đề còn mở

1. Vị trí tủ trên lối đi giữa 2 khối: CĐT/PCCC xác nhận lối đi không phải lối thoát nạn bắt buộc chiều rộng ≥ 1,8 m; nếu không, chuyển DB.HVAC + DDC-FAN vào khu ngoài phạm vi phải trên (CST) hoặc làm hốc tủ âm vách Office.
2. Gói điện cung cấp: Ik tại DB.HVAC, vị trí tủ hạ thế tổng, chiều dài & sụt áp cáp nguồn (khuyến nghị ≤ 1 %), có cho phép 295 A dự kiến không (kiểm tra máy biến áp hiện hữu).
3. Hãng CDU: yêu cầu ELCB (loại/độ nhạy), cấu hình DX-kit (số tín hiệu, 0–10 V hay Modbus), cáp truyền tín hiệu – chốt khi chọn hãng (WP5). Nếu CDU-2.01 phải 4 module (WP1 ghi chú) → thêm 1 lộ MCCB + cáp + cách ly (đã có lộ dự phòng 63 A).
4. Quạt EAF: WP1 ghi “biến tần”, chỉ đạo WP4 dùng khởi động trực tiếp. Nếu cần biến tần để cân chỉnh lưu lượng → thay MPCB+contactor bằng biến tần 0,75–1,1 kW trong DDC-FAN (cáp đổi sang loại màn chắn, cùng tiết diện).
5. Quạt tủ hood (gói nội thất): xác nhận quạt 380 V/3P 0,6 kW, vị trí quạt (nóc tủ/trên trần) và việc dùng hộp công tắc của HVAC thay công tắc đi kèm tủ; tủ có cảm biến sash báo động (không vào DDC).
6. Báo cháy: tư vấn PCCC xác nhận chiến lược – WP4 dừng AHU/FAF, GIỮ chạy quạt hút hóa chất; module rơ le báo cháy và cáp từ tủ báo cháy tới module thuộc gói PCCC.
7. Tuyến máng ngoài nhà phía bắc tới scrubber đi ốp tường A (không đi qua trần khu CST ngoài phạm vi) – cần CĐT đồng ý; vị trí tủ điều khiển scrubber theo nhà cung cấp.
8. Cao độ máng +3,6 m trùng vùng ống gió Phase 2 (+3,3…+4,1 m) – phối hợp với WP3 tại giao cắt; tại khu có ống Phase 1 (+4,2…+6,2) giữ nguyên.
9. Ổ cắm bảo trì và chiếu sáng sàn thiết bị trên trần: WP4 chỉ đưa ổ cắm tại 3 AHU; chiếu sáng bảo trì do gói điện (nếu cần).
10. Kết nối HMI Phase 1 (phương án OP-1…OP-5): cần thông tin giao thức/hãng hệ điều khiển Phase 1 và vị trí màn hình HMI.
11. ST1-CC2 (tủ hóa chất thứ 2 ở ST1) – nếu có hút, đã nằm trong van nhóm MD-G-ST1 (không đổi I/O); nếu tách van riêng +1 DO.


## 13. Kiểm tra

- Excel: 0 ô lỗi sau tính lại LibreOffice; 0 mạch “KHÔNG ĐẠT” (In ≥ Ib, PE, Iz ≥ In, ΔU ≤ 4 %).
- Nhu cầu max Excel = 194,0 kVA, Python = 194,0 kVA (khớp).
- Σ quạt tủ hood = 15 điểm (4 acid) khớp WP1/WP2; 6 module CDU khớp WP1.
- Mọi I/O có vị trí → có chiều dài cáp; mọi mạch có cáp → có tuyến máng.
