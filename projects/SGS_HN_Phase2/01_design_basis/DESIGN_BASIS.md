# CƠ SỞ THIẾT KẾ HVAC – SGS HẢI PHÒNG, KHU H&N PHASE 2 (NHÀ XƯỞNG 3)

> Tài liệu gốc cho mọi bước tính toán, thiết kế và BOQ. Mọi số liệu dùng trong `02_calc` … `06_boq` phải truy được về đây hoặc về `00_inputs`.
> Phiên bản: v1.1 – 14/11/2026 (cập nhật phân bổ điểm hút theo bản vẽ – WP2; gió bù 100 %; 3 AHU). Tổng thầu: **GGV (Greengrow Việt)**; thầu phụ MEP coi như cùng bên GGV.

## 0. Trạng thái RFI
Toàn bộ RFI (`RFI_HVAC_SGS_HN_Phase2.xlsx`, 19 câu) **được duyệt theo giả định của GGV** (cột "Giả định tạm thời của GGV"). Bổ sung chỉ đạo của chủ trì thiết kế (14/11/2026):
- Quạt, AHU đặt **trên trần panel, trên sàn kết cấu đã có**; dàn nóng (CDU/ODU) và scrubber mới đặt **ngoài nhà**.
- Điện: phạm vi **từ tủ tổng DB.HVAC trở đi** (gồm DB.HVAC, các tủ DDC riêng, cáp/máng tới thiết bị, điều khiển). Nguồn cấp tới DB.HVAC thuộc gói điện.
- Nước ngưng đi trên trần, bổ sung **bơm nước ngưng** nếu không đủ độ dốc tự chảy.
- Máy hút ẩm/bù ẩm (nếu có): phạm vi chỉ gồm cấp + lắp máy, ống gió (nếu có) và điện; nước cấp không thuộc HVAC.
- Giá đỡ/bệ thiết bị (sàn AHU, sàn quạt, móng bê tông CDU, scrubber…) thuộc phạm vi **xây dựng** → KHÔNG đưa vào BOQ HVAC. Ty treo/giá đỡ ống gió, ống gas, ống nước ngưng, máng cáp vẫn thuộc HVAC.

## 1. Nguồn dữ liệu (`00_inputs`)
| File | Nội dung |
|---|---|
| `261109_Cai_tao_khu_HN_xuong_3.pdf` | KT Phase 2: DW-A-03-03…08 (hiện trạng, vách, nhiệt ẩm, nội thất, nền, trần). Trục ngang E–D–C–B–A cách 7000; trục dọc 3–4–5 cách 7700. Trang 6 (DW-A-03-05) có diện tích + điều kiện phòng, trang 7 (DW-A-03-06) bố trí nội thất/tủ hood. |
| `HAI_PHONG_LAB_30102027_Equipment.xlsx` | Danh mục thiết bị CĐT (công suất, SL 2027 + bổ sung), số người, điều kiện; sheet HSE (scrubber, lọc than, lưu lượng chụp 1.100 m³/h). |
| `equipment_cdt_classified.json` | 66 thiết bị đã gán phòng + nhóm + hệ số tỏa nhiệt: `[dòng, phòng, tên, W, SL2027, SL bổ sung, loại điện, nhóm, hệ số, khí tươi, khí hút]`. |
| `noi_that_10.11.26_text_dump.txt` | Danh mục nội thất (tủ hood, chụp, tủ hóa chất, bồn rửa…). |
| `09112026_Phan_hoi_ban_ve_GGV.xlsx` | Phản hồi bản vẽ. |
| `phase1_dwg/` | Bản vẽ thi công HVAC Phase 1 (DWG + text trích xuất). AHU-1.01: 12.459 m³/h, OA 1.200, 90,05 kW; AHU-1.02: 12.010 m³/h, OA 2.000, 91,21 kW; ESP 800 Pa; lọc G4/F8; autoclave EAFE 3.050 + FAFE 3.100 m³/h (in-line, G4/F8). |
| `ifc_dwg/` | IFC tư vấn (MIN lab + H&N): hood 2.270 m³/h; canopy 2000×750 2.200; tủ hóa chất 1.300; arm hood 800; AAS 540; ICP 360; quạt ly tâm ESP 400 Pa; ống hút SS304 / PP (scrubber); weather cover + insect screen. |
| `PA_may_hut_am_HD_HCD.jpg` | PA máy hút ẩm (tham khảo). |

## 2. Điều kiện thiết kế
| Thông số | Giá trị | Nguồn |
|---|---|---|
| Ngoài trời hè | 37,8 °C / 62,8 % (h ≈ 105,7 kJ/kg, W ≈ 26,3 g/kg) | Phase 1 F3-SD-AC-004 |
| Ngoài trời đông | 9,8 °C / 71 % | Phase 1 |
| Trong phòng | 25 °C ±2; RH 40–70 % (Instrument RH ≤ 55 %) | RFI I.2 – GGV |
| Mùa đông | Không tạo ẩm; **có sưởi điện trong AHU** giữ ≥ 18 °C | RFI I.2 |
| Sau coil | 15 °C / 97 % (Phase 1); gió cấp 16 °C | Phase 1 |
| Cấp sạch | Không ISO; tối thiểu **6 ACH**; lọc **G4 + F8** | RFI I.4 |
| Gió tươi cho người | 25 m³/h/người (TCVN 5687) cho tất cả các phòng | TCVN 5687 |
| Chênh áp | Không yêu cầu (file thiết bị CĐT cột Q). **Gió bù FAFE = 100 % gió thải từng phòng** như Phase 1 (FAFE 3.100 / EAFE 3.050); không dùng gió chuyển tiếp giữa các phòng | CĐT + IFC/Phase 1 – chủ trì duyệt 14/11 |
| Cao độ | Trần panel PU 50 mm cao **3,0 m**; thiết bị/ống trên trần (sàn kết cấu có sẵn); Phase 1 ống gió BOD FFL+4.2…6.2 m | DW-A-03-04/08, Phase 1 |

## 3. Phòng (DW-A-03-05) – 13 phòng, 441,76 m²
| Phòng | m² | Người | Hood 1600 (acid) | Chụp treo | Tủ HC có hút | Chụp bồn rửa | Hút khác | Nhóm |
|---|---|---|---|---|---|---|---|---|
| Digestion room | 14,60 | 2 | 2 (2) | 1 (lò nung) | – | – | – | Hóa |
| Kjeldahl room | 18,88 | 2 | – | 1 | – | – | – | Hóa |
| Sample treatment 1 | 65,70 | 10 | 3 (1) | – | 1 | – | – | Hóa |
| Sample treatment 2 | 103,64 | 10 | 6 (1 – ST2-FH3) | – | 2 | 2 | – | Hóa |
| Standard (19,72 m²) | 19,72 | 2 | 4 (0) | – | – | 1 | – | Hóa |
| Standard (11,90 m²) | 11,90 | 2 | – | – | 1 (tủ khô) | – | – | Hóa |
| Balance room | 18,35 | 4 | – | – | – | – | – | Thường |
| Sample preparation | 58,99 | 10 | – | – | – | – | 2 arm hood × 800 | Thường (có hút) |
| Sample coding | 25,20 | 2 | – | – | – | – | – | Thường |
| Sample dividing | 9,98 | 1 | – | – | – | – | – | Thường |
| Instrument room | 35,82 | 10 | – | – | – | – | 8 điểm × 150 (3 LC-MS/MS, 3 HS-GC-MS/MS, 2 GC-FID) | Thường (có hút) |
| Reporting room | 18,02 | 10 | – | – | – | – | – | Thường |
| Office | 40,96 | 10 | – | – | – | – | – | Thường |

Lưu lượng hút (RFI IV – GGV): hood 1600 **2.270 m³/h** (CAV, cảm biến sash chỉ báo động); chụp treo **1.100 m³/h**; tủ hóa chất **1.300 m³/h**; chụp bồn rửa **300 m³/h**; arm hood **800 m³/h**; điểm hút thiết bị phân tích **150 m³/h**. Tính tải theo **SL thiết bị khi mua đủ** (SL 2027 + bổ sung).

## 4. Phương án hệ thống (RFI duyệt theo GGV)
1. **Điều hòa**: AHU DX (như Phase 1) + dàn nóng CDU ngoài nhà, hồi gió cho phép (kể cả phòng hóa, như Phase 1). Lọc G4+F8, sưởi điện, ESP tham khảo Phase 1 (≤ 800 Pa, phải tính lại theo tuyến ống). **Chốt 3 AHU** (chủ trì duyệt 14/11): AHU-2.01 khối trái trừ Instrument (đặt trên Balance room), AHU-2.02 khối phải (trên ST2), AHU-2.03 riêng Instrument room (chạy 24/7). Gió cấp 16,36 °C (15 °C sau coil + nhiệt quạt). Chấp nhận ACH cao ở Kjeldahl/ST1/Instrument do tải hiện.
2. **Gió bù**: quạt FAFE in-line/hộp, lọc G4+F8, **không xử lý lạnh**, cấp miệng gió ngay trước tủ hood/chụp; lưu lượng = 100 % gió thải từng phòng. **Ngoại lệ: Instrument room** – gió hút điểm thiết bị được bù 100 % bằng gió tươi đã xử lý qua AHU-2.03 (RH ≤ 55 %). Lấy gió tươi qua louver/weather cover + insect screen ở tường ngoài.
3. **Hút khí độc**:
   - Tủ hood (cả hữu cơ và acid) có **quạt riêng lắp sẵn theo máy** (600 W, 3 pha) – thiết bị thuộc gói nội thất; HVAC làm ống từ cổ tủ đến miệng xả/scrubber, cấp điện + điều khiển quạt từ DB.HVAC.
   - Hood hữu cơ: lọc than tại tủ → xả thẳng ra ngoài (ống SS304).
   - Hood acid (Digestion 2, ST1 1, ST2 1 = 4 tủ, 9.080 m³/h) → **scrubber ướt mới, ngoài nhà** (ống PP); scrubber trọn bộ có quạt + bơm tuần hoàn; nước thải scrubber đấu về XLNT (đầu chờ do gói khác).
   - Chụp treo, tủ hóa chất, chụp bồn rửa, arm hood, điểm hút Instrument → quạt hút EAF của HVAC (đặt trên trần, xả ra ngoài, ống SS304 cho hóa chất; GI cho arm hood bụi).
4. **Máy hút ẩm**: chỉ Instrument room (RFI II.4) – cấp + lắp + điện (+ ống gió nếu dạng nối ống); thoát nước ngưng đấu vào hệ nước ngưng trên trần.
5. **Nước ngưng**: AHU, máy hút ẩm → ống uPVC bọc cách nhiệt đi trên trần, về điểm thoát gần nhất (giả định đầu chờ thoát nước ở tường ngoài/trục A hoặc E – cần xác nhận); bơm nước ngưng nếu cần.
6. **Điện & điều khiển**: DB.HVAC (tủ tổng HVAC Phase 2) → các tủ động lực/DDC: DDC-AHU (mỗi AHU 1 tủ DDC kèm khởi động/biến tần quạt, sưởi điện), DDC-FAN (quạt FAFE/EAF, liên động hút–cấp), cấp nguồn quạt tủ hood, scrubber (tủ điều khiển đi kèm scrubber, HVAC cấp nguồn), máy hút ẩm. Cáp CXV/Cu, máng cáp trên trần, ống luồn. Liên động: quạt hút ↔ quạt cấp bù theo phòng; cảm biến T/RH phòng, chênh áp lọc, báo cháy cắt AHU/quạt cấp.

## 5. Ngoài phạm vi HVAC
Giá đỡ/bệ thiết bị; nguồn tới DB.HVAC; tủ hood/chụp/tủ hóa chất (nội thất); nước cấp máy bù ẩm, nước cấp scrubber và đấu nối XLNT phía ngoài; PCCC; hoàn trả panel/trần (chỉ khoét lỗ – ghi chú phối hợp).

## 6. Tiêu chuẩn & quy ước
TCVN 5687:2024 (thông gió – điều hòa), QCVN 06:2022/BXD (an toàn cháy: van chặn lửa khi xuyên vách ngăn cháy), SMACNA (ống gió), ASHRAE Fundamentals 2021 (tải lạnh, hệ số thiết bị), ANSI/ASHRAE 110 / EN 14175 (tham khảo hood). Đơn vị: m³/h, kW, Pa, mm. Vận tốc ống gió thiết kế: ống chính ≤ 8 m/s (cấp/hồi), ống nhánh ≤ 5 m/s, ống hút hóa chất 8–10 m/s; ma sát ≈ 1 Pa/m.
