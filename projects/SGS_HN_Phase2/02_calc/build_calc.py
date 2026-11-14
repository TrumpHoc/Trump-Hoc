#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WP1 - SGS Hai Phong, H&N Phase 2 - HVAC load calc, air balance, equipment selection (v4).

Usage:  python3 build_calc.py
  1) builds HVAC_Phase2_Calc_v4.xlsx (openpyxl, all engineering numbers are Excel formulas)
  2) recalculates it with LibreOffice (soffice --headless --convert-to xlsx) so cached values exist
  3) reloads data_only, scans for Excel errors, reads KiemTra status
  4) exports equipment_schedule.json and airflow_by_room.json from the recalculated values

Binding inputs: ../01_design_basis/DESIGN_BASIS.md, ../00_inputs/equipment_cdt_classified.json,
v3 envelope assumptions from ../tools/build.py.
"""
import json, os, shutil, subprocess, sys, tempfile, math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
PRJ = os.path.dirname(HERE)
XLSX = os.environ.get('OUT') or os.path.join(HERE, 'HVAC_Phase2_Calc_v4.xlsx')
EXPORT = not os.environ.get('OUT')  # scenario runs (OUT set) do not overwrite the JSON deliverables
OVR = json.loads(os.environ.get('OVR', '{}'))  # sensitivity overrides of CoSo inputs, e.g. OVR='{"fm":1.0}'
EQJSON = os.path.join(PRJ, '00_inputs', 'equipment_cdt_classified.json')

B = Font(bold=True); BT = Font(bold=True, size=13)
INP = PatternFill('solid', fgColor='FFF2CC'); HDR = PatternFill('solid', fgColor='D9E1F2')
TOT = PatternFill('solid', fgColor='E2EFDA'); SEC = PatternFill('solid', fgColor='FCE4D6')
th = Side(style='thin', color='999999'); BD = Border(left=th, right=th, top=th, bottom=th)
WR = Alignment(wrap_text=True, vertical='center')

wb = openpyxl.Workbook()
N = {}  # name -> absolute reference

# =====================================================================================
# Sheet CoSo (inputs)
# =====================================================================================
s = wb.active; s.title = 'CoSo'
s['A1'] = 'CƠ SỞ TÍNH TOÁN – HVAC PHASE 2 (KHU H&N, NHÀ XƯỞNG 3, SGS HẢI PHÒNG) – v4'; s['A1'].font = BT
s['A2'] = 'Ô vàng = số liệu đầu vào (có nguồn). Ô trắng = công thức. DB = 01_design_basis/DESIGN_BASIS.md; v3 = tools/build.py; GĐ = giả định của kỹ sư WP1 (cần chủ trì duyệt).'
P = [
 ('SEC', 'ĐIỀU KIỆN KHÍ HẬU & TRẠNG THÁI'),
 ('P', 'Áp suất khí quyển', 101.325, 'kPa', 'Mực nước biển (Hải Phòng)'),
 ('To', 'Ngoài trời hè – nhiệt độ khô', 37.8, '°C', 'DB §2 (Phase 1 F3-SD-AC-004)'),
 ('RHo', 'Ngoài trời hè – độ ẩm', 0.628, '-', 'DB §2'),
 ('Tw', 'Ngoài trời đông – nhiệt độ khô', 9.8, '°C', 'DB §2'),
 ('RHw', 'Ngoài trời đông – độ ẩm', 0.71, '-', 'DB §2'),
 ('Tn', 'Điều kiện “nồm” (định cỡ máy hút ẩm) – nhiệt độ', 22, '°C', 'GĐ – mùa nồm Bắc Bộ (tháng 2–4)'),
 ('RHn', 'Điều kiện “nồm” – độ ẩm', 0.95, '-', 'GĐ'),
 ('Tr', 'Nhiệt độ phòng thiết kế', 25, '°C', 'DB §2 (25 ±2)'),
 ('RHr', 'Độ ẩm thiết kế phòng thường/phòng hóa', 0.60, '-', 'DB §2 dải 40–70 % → điểm thiết kế 60 % (như v3/Phase 1)'),
 ('RHmax', 'Độ ẩm tối đa cho phép phòng thường/phòng hóa', 0.70, '-', 'DB §2'),
 ('RHi', 'Độ ẩm thiết kế = tối đa Instrument room', 0.55, '-', 'DB §2 (Instrument RH ≤ 55 %)'),
 ('Twr', 'Nhiệt độ phòng tối thiểu mùa đông', 18, '°C', 'DB §2 (sưởi điện giữ ≥ 18 °C)'),
 ('Tc', 'Nhiệt độ sau coil', 15, '°C', 'DB §2 (Phase 1: 15 °C/97 %)'),
 ('RHc', 'Độ ẩm sau coil', 0.97, '-', 'DB §2'),
 ('rho', 'Khối lượng riêng không khí', 1.2, 'kg/m³', 'Tiêu chuẩn'),
 ('cp', 'Nhiệt dung riêng không khí', 1.006, 'kJ/kg.K', 'ASHRAE Fund. Ch.1'),
 ('hfg', 'Ẩn nhiệt hóa hơi (dùng cho tải ẩn)', 2501, 'kJ/kg', 'ASHRAE Fund. Ch.1'),
 ('SEC', 'KẾT CẤU BAO CHE, NỘI NHIỆT (như v3)'),
 ('Tsur', 'Nhiệt độ xưởng bao quanh (không ĐH) – hè', 35, '°C', 'v3 / GĐ'),
 ('Tpl', 'Nhiệt độ plenum trên trần – hè', 45, '°C', 'v3 / GĐ'),
 ('Tsurw', 'Nhiệt độ xưởng bao quanh – đông', 12, '°C', 'GĐ (xưởng không sưởi, OA 9,8 °C)'),
 ('Tplw', 'Nhiệt độ plenum trên trần – đông', 10, '°C', 'GĐ'),
 ('Uw', 'U vách panel 50 mm', 0.45, 'W/m².K', 'v3'),
 ('Uc', 'U trần panel PU 50 mm', 0.42, 'W/m².K', 'v3'),
 ('H', 'Chiều cao thông thủy', 3.0, 'm', 'DB §2 (trần panel 3,0 m)'),
 ('LPD', 'Mật độ chiếu sáng', 12, 'W/m²', 'v3'),
 ('qps', 'Nhiệt hiện / người', 75, 'W', 'v3 (ASHRAE Fund. Ch.18)'),
 ('qpl', 'Nhiệt ẩn / người', 55, 'W', 'v3'),
 ('SF', 'Hệ số an toàn tải phòng', 0.10, '-', 'v3 / chỉ đạo WP1'),
 ('QB', 'Tính SL khi mua đủ (1 = SL2027 + bổ sung)', 1, '-', 'DB §3'),
 ('SEC', 'GIÓ TƯƠI, GIÓ THẢI, GIÓ BÙ'),
 ('OAp', 'Gió tươi / người', 25, 'm³/h.người', 'DB §2 (TCVN 5687)'),
 ('FH', 'Hút tủ hood 1600 (CAV)', 2270, 'm³/h/tủ', 'DB §3 (IFC)'),
 ('CAN', 'Hút chụp treo', 1100, 'm³/h/chụp', 'DB §3 (CĐT sheet HSE)'),
 ('CAB', 'Hút tủ hóa chất', 1300, 'm³/h/tủ', 'DB §3 (IFC)'),
 ('SNK', 'Hút chụp bồn rửa', 300, 'm³/h/chụp', 'DB §3 (RFI IV)'),
 ('ARM', 'Hút arm hood', 800, 'm³/h/cái', 'DB §3 (IFC)'),
 ('INS', 'Hút điểm thiết bị phân tích', 150, 'm³/h/điểm', 'DB §3 (RFI IV)'),
 ('fm', 'Gió bù FAFE / gió thải (phòng có FAFE)', 1.0, '-', 'Chủ trì duyệt 14/11/2026: theo IFC/Phase 1 (FAFE ≈ gió thải); CĐT không yêu cầu chênh áp – thay DB §2 (90 %)'),
 ('fmi', 'Gió bù FAFE / gió thải – Instrument room', 0.0, '-', 'GĐ WP1: bù 100 % bằng gió tươi đã xử lý qua AHU-2.03 (RH ≤ 55 %) – xem báo cáo §4'),
  ('SPL', 'Tỷ lệ gió bù FAFE tràn vào vùng phòng – tính tải lạnh', 0.0, '-', 'v3 / IFC-Phase 1 (=0); xem cột “tỷ lệ tràn chịu được”'),
 ('SPLw', 'Tỷ lệ gió bù FAFE tràn vào phòng – tính sưởi mùa đông', 0.10, '-', 'GĐ WP1 (an toàn cho sưởi)'),
 ('ACHmin', 'Bội số trao đổi tối thiểu', 6, '1/h', 'DB §2 (RFI I.4)'),
 ('ACHmax', 'Ngưỡng bội số cần lý giải', 40, '1/h', 'Chỉ đạo WP1'),
 ('SEC', 'AHU – CHỌN MÁY, QUẠT, SƯỞI'),
 ('SEL', 'Dự phòng chọn máy (lạnh)', 0.10, '-', 'Chỉ đạo WP1'),
 ('ETA', 'Hiệu suất tổng quạt (quạt × động cơ × truyền động)', 0.60, '-', 'Chỉ đạo WP1'),
 ('MM', 'Hệ số dự phòng công suất động cơ', 1.15, '-', 'Thực hành chung'),
 ('dG4', 'Tổn thất lọc G4 (trung bình đầu/cuối)', 100, 'Pa', 'NCC điển hình (50→150)'),
 ('dF8', 'Tổn thất lọc F8 (trung bình đầu/cuối)', 185, 'Pa', 'NCC điển hình (120→250)'),
 ('dCoil', 'Tổn thất coil DX ướt', 120, 'Pa', 'NCC điển hình @2,5 m/s'),
 ('dHeat', 'Tổn thất sưởi điện', 20, 'Pa', 'NCC điển hình'),
 ('dCas', 'Tổn thất vỏ, hộp trộn, van gió, cửa vào quạt', 60, 'Pa', 'GĐ'),
 ('fr', 'Ma sát ống gió', 1.0, 'Pa/m', 'DB §6'),
 ('kfit', 'Hệ số chiều dài tương đương phụ kiện (L_td = L×(1+k))', 1.0, '-', 'GĐ'),
 ('dTerm', 'Tổn thất đầu cuối (miệng cấp 30 + hồi 20 + VCD/FD 30 + tiêu âm 50)', 130, 'Pa', 'GĐ'),
 ('ESPmin', 'ESP tối thiểu chọn cho AHU', 500, 'Pa', 'GĐ WP1 (Phase 1 800 Pa có HEPA; Phase 2 không HEPA)'),
 ('ESPm', 'Dự phòng ESP so với ước tính', 0.25, '-', 'GĐ'),
 ('vCoil', 'Vận tốc mặt coil/lọc', 2.5, 'm/s', 'Thực hành chung'),
 ('vDuct', 'Vận tốc ống chính cấp/hồi', 8.0, 'm/s', 'DB §6'),
 ('dTrh', 'Khả năng gia nhiệt lại tối thiểu (chế độ khử ẩm)', 2.0, 'K', 'GĐ WP1'),
 ('Hstep', 'Bậc công suất sưởi điện', 3, 'kW', 'NCC điển hình'),
 ('SEC', 'DÀN NÓNG CDU'),
 ('COP', 'COP dàn nóng @37,8 °C', 3.0, '-', 'Chỉ đạo WP1'),
 ('CDUmax', 'Công suất lạnh danh định tối đa / module CDU', 50, 'kW', 'GĐ (≈ 18–20 HP, inverter)'),
 ('kref', 'Suy giảm công suất theo chiều dài ống gas', 0.002, '1/m', 'GĐ – catalogue DX điển hình (≈ −2 %/10 m)'),
 ('Lref0', 'Chiều dài ống gas chuẩn của catalogue', 7.5, 'm', 'Catalogue điển hình'),
 ('SEC', 'QUẠT HÚT / CẤP BÙ / SCRUBBER'),
 ('ESPf', 'ESP quạt FAF (ống, louver, miệng gió – không gồm lọc)', 400, 'Pa', 'DB §1 (IFC: quạt 400 Pa)'),
 ('ESPe', 'ESP quạt EAF', 400, 'Pa', 'DB §1 (IFC)'),
 ('HOODW', 'Công suất quạt tích hợp tủ hood', 0.6, 'kW', 'DB §4.3 (600 W, 3 pha)'),
 ('SCRm', 'Dự phòng lưu lượng scrubber', 0.10, '-', 'Chỉ đạo WP1'),
 ('vScr', 'Vận tốc khí trong tháp đệm', 1.8, 'm/s', 'Thực hành chung (1,5–2,0)'),
 ('LG', 'Tỷ lệ lỏng/khí L/G', 2.5, 'L/m³', 'Thực hành chung (2–3)'),
 ('Hp', 'Cột áp bơm tuần hoàn', 20, 'm', 'GĐ'),
 ('etaP', 'Hiệu suất bơm', 0.55, '-', 'GĐ'),
 ('dScr', 'Tổn thất scrubber (đệm + khử sương + vào/ra)', 600, 'Pa', 'NCC điển hình'),
 ('dScrD', 'Tổn thất ống PP gom + ống xả (sau quạt tủ hood)', 400, 'Pa', 'GĐ'),
 ('DOSE', 'Bơm định lượng NaOH', 0.18, 'kW', 'NCC điển hình'),
 ('SEC', 'MÁY HÚT ẨM, NƯỚC NGƯNG, ĐIỆN'),
 ('ACHinf', 'Thấm khí (cửa) Instrument khi nồm', 1.0, '1/h', 'GĐ'),
 ('HRocc', 'Số giờ có người / ngày', 10, 'h', 'GĐ'),
 ('DHm', 'Dự phòng máy hút ẩm', 1.3, '-', 'GĐ'),
 ('DHr', 'Tỷ lệ năng suất @25 °C/55 % so với định mức @30 °C/80 %', 0.45, '-', 'Đường đặc tính NCC điển hình'),
 ('DHef', 'Hiệu suất máy hút ẩm @định mức', 2.5, 'L/kWh', 'NCC điển hình'),
 ('Hav', 'Độ cao rơi khả dụng cho ống nước ngưng trên trần', 0.15, 'm', 'GĐ (đáy máng – bẫy – đỉnh trần panel)'),
 ('slope', 'Độ dốc tối thiểu ống nước ngưng', 0.01, '-', 'Thực hành chung (1 %)'),
 ('CPk', 'Ngưỡng dùng bơm: L > k × L_max tự chảy', 0.8, '-', 'GĐ (dự phòng sai khác hiện trường 20 %)'),
 ('CPf', 'Hệ số lưu lượng bơm nước ngưng / lượng ngưng', 3, '-', 'Thực hành chung'),
 ('CPh', 'Cột áp bơm nước ngưng', 6, 'm', 'GĐ'),
 ('CPkw', 'Công suất bơm nước ngưng (loại có bình chứa)', 0.1, 'kW', 'NCC điển hình'),
 ('V3', 'Điện áp 3 pha', 380, 'V', 'Lưới VN'),
 ('V1', 'Điện áp 1 pha', 220, 'V', 'Lưới VN'),
 ('PF', 'Hệ số công suất (CDU, động cơ)', 0.85, '-', 'GĐ'),
 ('EM', 'Hiệu suất động cơ', 0.85, '-', 'GĐ'),
 ('M3', 'Ngưỡng dùng động cơ 3 pha', 0.75, 'kW', 'GĐ'),
]
s.append([]); s.append(['Ký hiệu', 'Thông số', 'Giá trị', 'Đơn vị', 'Nguồn'])
for c in s[4]: c.font = B; c.fill = HDR
r = 5
for it in P:
    if it[0] == 'SEC':
        s.cell(r, 1, it[1]).font = B
        for k in range(1, 6): s.cell(r, k).fill = SEC
        r += 1; continue
    k, t, v, u, src = it
    s.cell(r, 1, k); s.cell(r, 2, t); c = s.cell(r, 3, OVR.get(k, v)); c.fill = INP; s.cell(r, 4, u); s.cell(r, 5, src)
    N[k] = f"CoSo!$C${r}"; r += 1
r += 1
s.cell(r, 1, 'TRẠNG THÁI KHÔNG KHÍ (Magnus: Psat = 0,61094·exp(17,625T/(T+243,04)))').font = B; r += 1
for i, h in enumerate(['Điểm', 'T (°C)', 'RH', 'Psat (kPa)', 'W (kg/kg)', 'h (kJ/kg)']):
    c = s.cell(r, i + 1, h); c.font = B; c.fill = HDR
r += 1
def state(rr, label, T, RH, key):
    s.cell(rr, 1, label); s.cell(rr, 2, f'={T}'); s.cell(rr, 3, f'={RH}')
    s.cell(rr, 4, f'=0.61094*EXP(17.625*B{rr}/(B{rr}+243.04))')
    s.cell(rr, 5, f"=0.622*C{rr}*D{rr}/({N['P']}-C{rr}*D{rr})")
    s.cell(rr, 6, f'=1.006*B{rr}+E{rr}*(2501+1.86*B{rr})')
    N['W' + key] = f"CoSo!$E${rr}"; N['h' + key] = f"CoSo!$F${rr}"
    s.cell(rr, 5).number_format = '0.00000'; s.cell(rr, 6).number_format = '0.00'; s.cell(rr, 4).number_format = '0.000'
for lab, T, RH, key in [('Ngoài trời hè', 'To', 'RHo', 'o'), ('Sau coil', 'Tc', 'RHc', 'c'), ('Ngoài trời đông', 'Tw', 'RHw', 'w'),
                        ('Phòng 25 °C/60 %', 'Tr', 'RHr', 'r'), ('Phòng 25 °C/70 % (tối đa)', 'Tr', 'RHmax', 'rmax'),
                        ('Instrument 25 °C/55 %', 'Tr', 'RHi', 'i'), ('Nồm', 'Tn', 'RHn', 'n')]:
    state(r, lab, N[T], N[RH], key); r += 1
r += 1
s.cell(r, 1, 'CÔNG SUẤT ĐỘNG CƠ TIÊU CHUẨN IEC (kW) – dùng chọn động cơ').font = B; r += 1
STD = [0.18, 0.25, 0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3, 4, 5.5, 7.5, 11, 15, 18.5, 22, 30, 37]
r0s = r
for v in STD:
    s.cell(r, 3, v); r += 1
N['STD'] = f"CoSo!$C${r0s}:$C${r-1}"
for col, w in zip('ABCDE', [10, 62, 11, 11, 70]): s.column_dimensions[col].width = w
s.freeze_panes = 'A5'

def motor(x):  # smallest standard motor >= x*MM
    return f"SMALL({N['STD']},COUNTIF({N['STD']},\"<\"&({x})*{N['MM']})+1)"
def fla3(kw):  # 3-phase motor current
    return f"({kw})*1000/(SQRT(3)*{N['V3']}*{N['PF']}*{N['EM']})"
def fla1(kw):
    return f"({kw})*1000/({N['V1']}*{N['PF']}*{N['EM']})"

# =====================================================================================
# Sheet ThietBi (equipment heat)
# =====================================================================================
e = wb.create_sheet('ThietBi')
e['A1'] = 'NHIỆT TỎA THIẾT BỊ VÀO PHÒNG'; e['A1'].font = BT
e['A2'] = 'Nguồn: 00_inputs/equipment_cdt_classified.json (66 dòng CĐT, hệ số tỏa nhiệt đã gán) + 2 dòng bổ sung cho 2 phòng Standard (không có trong danh mục CĐT – chỉ đạo WP1). SL tính = SL2027 + QB × SL bổ sung.'
hdr = ['Dòng CĐT', 'Phòng', 'Thiết bị', 'Công suất (W/cái)', 'SL 2027', 'SL bổ sung', 'SL tính tải', 'Công suất tổng (W)', 'Nhóm', 'Hệ số tỏa nhiệt', 'Nhiệt vào phòng (W)', 'Loại điện']
for i, h in enumerate(hdr):
    c = e.cell(4, i + 1, h); c.font = B; c.fill = HDR; c.alignment = WR; c.border = BD
EQ = json.load(open(EQJSON, encoding='utf-8'))
EQ = [[x[0], x[1], x[2], x[3], x[4], x[5], x[7], x[8], x[6]] for x in EQ]
EQ += [['—', 'Standard (11.90 m²)', 'Tủ lạnh ×3 (theo mặt bằng DW-A-03-06) – không có trong danh mục CĐT', 485, 3, 0, 'Tủ lạnh/đông, chiller', 0.5, '1 pha'],
       ['—', 'Standard (19.72 m²)', 'Thiết bị nhỏ trên bàn 1 kW – không có trong danh mục CĐT', 1000, 1, 0, 'Khác', 0.5, '1 pha']]
for k, (rw, room, dev, w, q1, q2, cat, fct, ph) in enumerate(EQ):
    rr = 5 + k
    for col, v in zip('ABCDEFIJL', [rw, room, dev, w, q1, q2, cat, fct, ph]): e[f'{col}{rr}'] = v
    for col in 'DEFJ': e[f'{col}{rr}'].fill = INP
    e[f'G{rr}'] = f"=E{rr}+{N['QB']}*F{rr}"; e[f'H{rr}'] = f'=D{rr}*G{rr}'; e[f'K{rr}'] = f'=H{rr}*J{rr}'
    for col in 'ABCDEFGHIJKL': e[f'{col}{rr}'].border = BD
    for col, fm in (('D', '#,##0'), ('H', '#,##0'), ('K', '#,##0'), ('J', '0.00')): e[f'{col}{rr}'].number_format = fm
EQ_LAST = 5 + len(EQ) - 1
EQ_TOT = EQ_LAST + 1
e[f'C{EQ_TOT}'] = 'TỔNG'; e[f'C{EQ_TOT}'].font = B
for col in 'HK':
    e[f'{col}{EQ_TOT}'] = f'=SUM({col}5:{col}{EQ_LAST})'; e[f'{col}{EQ_TOT}'].font = B; e[f'{col}{EQ_TOT}'].number_format = '#,##0'
N['EQ_K_TOT'] = f"ThietBi!$K${EQ_TOT}"
EQ_B = f"ThietBi!$B$5:$B${EQ_LAST}"; EQ_K = f"ThietBi!$K$5:$K${EQ_LAST}"; EQ_H = f"ThietBi!$H$5:$H${EQ_LAST}"
for col, w in zip('ABCDEFGHIJKL', [8, 22, 50, 11, 8, 9, 9, 12, 30, 9, 12, 14]): e.column_dimensions[col].width = w
e.freeze_panes = 'C5'

# =====================================================================================
# Sheet TaiPhong (room loads & air balance)
# =====================================================================================
t = wb.create_sheet('TaiPhong')
t['A1'] = 'TẢI LẠNH PHÒNG & CÂN BẰNG GIÓ – PHASE 2 (v4)'; t['A1'].font = BT
t['A2'] = ('Diện tích, số người, số tủ: DB §3 (DW-A-03-05/06). Chiều dài vách tiếp giáp xưởng: v3. Cân bằng khối lượng mỗi phòng: '
           'Cấp AHU + Bù FAFE + Chuyển vào = Hồi + Thải + Chuyển ra. Gió chuyển lấy từ bảng “GIÓ CHUYỂN” bên dưới.')
ROOMS = [
 # name, group, ahu, area, ppl, wall(non-A exposed, m), wallA(external wall A, m), rh, rhmax, fh, acid, can, cab, snk, arm, ins, fmkey, FAF, EAF
 # exposed wall lengths: 03_geometry/geometry.json rooms[].exposure (WP2); exhaust split: DB §3 + WP2 (DW-A-03-06) for right block
 ('Digestion room', 'Hóa', 'AHU-2.01', 14.60, 2, 7.65, 0, 'RHr', 'RHmax', 2, 2, 1, 0, 0, 0, 0, 'fm', 'FAF-2.01', 'EAF-2.01'),
 ('Kjeldahl room', 'Hóa', 'AHU-2.01', 18.88, 2, 8.74, 0, 'RHr', 'RHmax', 0, 0, 1, 0, 0, 0, 0, 'fm', 'FAF-2.01', 'EAF-2.01'),
 ('Sample treatment 1', 'Hóa', 'AHU-2.01', 65.70, 10, 15.60, 0, 'RHr', 'RHmax', 3, 1, 0, 1, 0, 0, 0, 'fm', 'FAF-2.01', 'EAF-2.01'),
 ('Balance room', 'Thường', 'AHU-2.01', 18.35, 4, 3.215, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 0, 0, 'fm', '', ''),
 ('Sample preparation', 'Thường (có hút)', 'AHU-2.01', 58.99, 10, 12.75, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 2, 0, 'fm', 'FAF-2.01', 'EAF-2.03'),
 ('Sample coding', 'Thường', 'AHU-2.01', 25.20, 2, 5.305, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 0, 0, 'fm', '', ''),
 ('Sample dividing', 'Thường', 'AHU-2.01', 9.98, 1, 6.85, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 0, 0, 'fm', '', ''),
 ('Reporting room', 'Thường', 'AHU-2.01', 18.02, 10, 11.385, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 0, 0, 'fm', '', ''),
 ('Instrument room', 'Thường (có hút)', 'AHU-2.03', 35.82, 10, 3.785, 0, 'RHi', 'RHi', 0, 0, 0, 0, 0, 0, 8, 'fmi', '', 'EAF-2.03'),
 ('Sample treatment 2', 'Hóa', 'AHU-2.02', 103.64, 10, 13.44, 15.625, 'RHr', 'RHmax', 6, 1, 0, 2, 2, 0, 0, 'fm', 'FAF-2.02', 'EAF-2.02'),
 ('Standard (19.72 m²)', 'Hóa', 'AHU-2.02', 19.72, 2, 9.70, 0, 'RHr', 'RHmax', 4, 0, 0, 0, 1, 0, 0, 'fm', 'FAF-2.02', 'EAF-2.02'),
 ('Standard (11.90 m²)', 'Hóa', 'AHU-2.02', 11.90, 2, 1.75, 0, 'RHr', 'RHmax', 0, 0, 0, 1, 0, 0, 0, 'fm', 'FAF-2.02', 'EAF-2.02'),
 ('Office', 'Thường', 'AHU-2.02', 40.96, 10, 13.475, 0, 'RHr', 'RHmax', 0, 0, 0, 0, 0, 0, 0, 'fm', '', ''),
]
COLS = [
 ('stt', 'STT'), ('name', 'Phòng'), ('grp', 'Nhóm'), ('ahu', 'AHU phục vụ'), ('A', 'Diện tích (m²)'), ('H', 'Cao (m)'), ('V', 'Thể tích (m³)'),
 ('T', 'T thiết kế (°C)'), ('RH', 'RH thiết kế'), ('RHM', 'RH tối đa'), ('Ps', 'Psat (kPa)'), ('W', 'W phòng (kg/kg)'), ('h', 'h phòng (kJ/kg)'), ('WM', 'W @RH tối đa'),
 ('ppl', 'Số người'), ('eq', 'Nhiệt thiết bị (kW) – ThietBi'), ('wl', 'Dài vách giáp KV không ĐH (m) – WP2'), ('wlA', 'Dài vách giáp tường ngoài A (m) – WP2'),
 ('Qw', 'Q vách (W)'), ('Qc', 'Q trần (W)'), ('Ql', 'Q đèn (W)'), ('Qps', 'Q người hiện (W)'), ('Qe', 'Q thiết bị (W)'), ('RSH', 'RSH (W) ×(1+SF)'),
 ('Qpl', 'Q người ẩn (W)'), ('Qtl', 'Ẩn do gió chuyển vào (W)'), ('RLH', 'RLH (W) ×(1+SF)'),
 ('nfh', 'Tủ hood 1600'), ('nac', '…trong đó hood acid'), ('ncan', 'Chụp treo'), ('ncab', 'Tủ hóa chất'), ('nsnk', 'Chụp bồn rửa'), ('narm', 'Arm hood'), ('nins', 'Điểm hút thiết bị'),
 ('Efh', 'Hút hood (m³/h)'), ('Eac', '…hood acid → scrubber'), ('Eor', '…hood hữu cơ → xả thẳng'), ('Ecan', 'Hút chụp treo'), ('Ecab', 'Hút tủ HC'), ('Esnk', 'Hút bồn rửa'),
 ('Earm', 'Hút arm hood'), ('Eins', 'Hút thiết bị PT'), ('E', 'TỔNG GIÓ THẢI (m³/h)'), ('Eeaf', 'Thải qua quạt EAF (m³/h)'), ('ACHE', 'ACH thải'),
 ('fm', 'Tỷ lệ bù FAFE'), ('M', 'Gió bù FAFE (m³/h)'), ('Ti', 'Gió chuyển VÀO (m³/h)'), ('To', 'Gió chuyển RA (m³/h)'), ('Net', 'Thải − FAFE ± chuyển (m³/h)'),
 ('OAp', 'Gió tươi theo người (m³/h)'), ('Tsup', 'T gió cấp (°C) – AHU'), ('Ss', 'Cấp cho tải hiện (m³/h)'), ('Sl', 'Cấp cho tải ẩn (m³/h)'), ('Sa', 'Cấp 6 ACH (m³/h)'),
 ('S', 'GIÓ CẤP THIẾT KẾ (m³/h)'), ('ACH', 'ACH cấp'), ('drv', 'Yếu tố quyết định'), ('Xf', 'Gió tươi dư rò ra qua khe cửa (chỉ phòng không phải Hóa) (m³/h)'), ('R', 'GIÓ HỒI (m³/h)'), ('Tact', 'T phòng dự kiến (°C)'),
 ('LC', 'Khả năng khử ẩm gió cấp @RH TK (W)'), ('LOK', 'Kiểm tra ẩm'), ('LCM', 'Khả năng khử ẩm @RH tối đa (W)'), ('SPT', 'Tỷ lệ tràn FAFE chịu được (giữ ≤ RH tối đa)'),
 ('Zp', 'Zp = OA người / cấp (ASHRAE 62.1)'), ('Qwin', 'Tổn thất bao che mùa đông (W)'), ('Qrm', 'Tải lạnh phòng (kW)'), ('Wm2', 'W/m² (lạnh phòng)'),
 ('FAF', 'Quạt bù FAF'), ('EAF', 'Quạt hút EAF'), ('SCR', 'Đấu scrubber'), ('nhf', 'Quạt tủ hood (điểm cấp nguồn)'),
]
CI = {k: CL(i + 1) for i, (k, _) in enumerate(COLS)}
HR = 4
for i, (k, h) in enumerate(COLS):
    c = t.cell(HR, i + 1, h); c.font = B; c.fill = HDR; c.alignment = WR; c.border = BD
t.row_dimensions[HR].height = 80
R0 = HR + 1; RL = R0 + len(ROOMS) - 1; RT = RL + 1
ROW = {rm[0]: R0 + i for i, rm in enumerate(ROOMS)}
# transfer table location
TR0 = RT + 4
def _sh(rm): return f"({CI['E']}{ROW[rm]}-{CI['M']}{ROW[rm]})"   # shortfall E - FAFE of a room
# Transfer air: none. Lead decision 14/11/2026 rejected the 90 % make-up + transfer-air scheme (v4 draft);
# FAFE = 100 % of each room's exhaust per IFC/Phase 1, no transfer grilles. Mechanism kept (list may be filled later).
TRANS = []  # (from, to, formula(lambda), path)
TRL = TR0 + max(len(TRANS), 1) - 1
TR_FROM = f"$B${TR0}:$B${TRL}"; TR_TO = f"$C${TR0}:$C${TRL}"; TR_Q = f"$D${TR0}:$D${TRL}"; TR_LAT = f"$H${TR0}:$H${TRL}"
NAMES = f"$B${R0}:$B${RL}"

def rng(k):
    return f"TaiPhong!${CI[k]}${R0}:${CI[k]}${RL}"

for i, rm in enumerate(ROOMS):
    (nm, grp, ahu, A, ppl, wl, wlA, rhk, rhmk, nfh, nac, ncan, ncab, nsnk, narm, nins, fmk, faf, eaf) = rm
    r = R0 + i
    c = lambda k: f"{CI[k]}{r}"
    V = {
     'stt': i + 1, 'name': nm, 'grp': grp, 'ahu': ahu, 'A': A, 'ppl': ppl, 'wl': wl, 'wlA': wlA,
     'nfh': nfh, 'nac': nac, 'ncan': ncan, 'ncab': ncab, 'nsnk': nsnk, 'narm': narm, 'nins': nins,
     'FAF': faf, 'EAF': eaf,
     'H': f"={N['H']}", 'V': f"={c('A')}*{c('H')}", 'T': f"={N['Tr']}", 'RH': f"={N[rhk]}", 'RHM': f"={N[rhmk]}",
     'Ps': f"=0.61094*EXP(17.625*{c('T')}/({c('T')}+243.04))",
     'W': f"=0.622*{c('RH')}*{c('Ps')}/({N['P']}-{c('RH')}*{c('Ps')})",
     'h': f"=1.006*{c('T')}+{c('W')}*(2501+1.86*{c('T')})",
     'WM': f"=0.622*{c('RHM')}*{c('Ps')}/({N['P']}-{c('RHM')}*{c('Ps')})",
     'eq': f"=SUMIF({EQ_B},{c('name')},{EQ_K})/1000",
     'Qw': f"={c('wl')}*{c('H')}*{N['Uw']}*({N['Tsur']}-{c('T')})+{c('wlA')}*{c('H')}*{N['Uw']}*({N['To']}-{c('T')})",
     'Qc': f"={c('A')}*{N['Uc']}*({N['Tpl']}-{c('T')})",
     'Ql': f"={c('A')}*{N['LPD']}", 'Qps': f"={c('ppl')}*{N['qps']}", 'Qe': f"={c('eq')}*1000",
     'RSH': f"=({c('Qw')}+{c('Qc')}+{c('Ql')}+{c('Qps')}+{c('Qe')}+{N['SPL']}*{c('M')}*{N['rho']}*{N['cp']}*({N['To']}-{c('T')})/3.6)*(1+{N['SF']})",
     'Qpl': f"={c('ppl')}*{N['qpl']}",
     'Qtl': f"=SUMIF({TR_TO},{c('name')},{TR_LAT})",
     'RLH': f"=({c('Qpl')}+{c('Qtl')}+{N['SPL']}*{c('M')}*{N['rho']}*({N['Wo']}-{c('W')})*{N['hfg']}/3.6)*(1+{N['SF']})",
     'Efh': f"={c('nfh')}*{N['FH']}", 'Eac': f"={c('nac')}*{N['FH']}", 'Eor': f"={c('Efh')}-{c('Eac')}",
     'Ecan': f"={c('ncan')}*{N['CAN']}", 'Ecab': f"={c('ncab')}*{N['CAB']}", 'Esnk': f"={c('nsnk')}*{N['SNK']}",
     'Earm': f"={c('narm')}*{N['ARM']}", 'Eins': f"={c('nins')}*{N['INS']}",
     'E': f"={c('Efh')}+{c('Ecan')}+{c('Ecab')}+{c('Esnk')}+{c('Earm')}+{c('Eins')}",
     'Eeaf': f"={c('Ecan')}+{c('Ecab')}+{c('Esnk')}+{c('Earm')}+{c('Eins')}",
     'ACHE': f"={c('E')}/{c('V')}",
     'fm': f"={N[fmk]}", 'M': f"={c('fm')}*{c('E')}",
     'Ti': f"=SUMIF({TR_TO},{c('name')},{TR_Q})", 'To': f"=SUMIF({TR_FROM},{c('name')},{TR_Q})",
     'Net': f"={c('E')}-{c('M')}-{c('Ti')}+{c('To')}",
     'OAp': f"={c('ppl')}*{N['OAp']}",
     'Tsup': f"=INDEX(AHU!$C$__TSUP__:$E$__TSUP__,MATCH({c('ahu')},AHU!$C$4:$E$4,0))",
     'Ss': f"={c('RSH')}*3.6/({N['rho']}*{N['cp']}*({c('T')}-{c('Tsup')}))",
     'Sl': f"={c('RLH')}*3.6/({N['rho']}*{N['hfg']}*({c('W')}-{N['Wc']}))",
     'Sa': f"={N['ACHmin']}*{c('V')}",
     'S': f"=ROUNDUP(MAX({c('Ss')},{c('Sl')},{c('Sa')},{c('Net')})/10,0)*10",
     'ACH': f"={c('S')}/{c('V')}",
     'drv': f"=IF({c('Ss')}>=MAX({c('Sl')},{c('Sa')},{c('Net')}),\"Tải hiện\",IF({c('Sa')}>=MAX({c('Sl')},{c('Net')}),\"6 ACH\",IF({c('Net')}>={c('Sl')},\"Cân bằng gió\",\"Tải ẩn\")))",
     'Xf': f"=IF({c('grp')}=\"Hóa\",0,INDEX(AHU!$C$__XR__:$E$__XR__,MATCH({c('ahu')},AHU!$C$4:$E$4,0))*{c('S')}/SUMPRODUCT(($D${R0}:$D${RL}={c('ahu')})*($C${R0}:$C${RL}<>\"Hóa\")*${CI['S']}${R0}:${CI['S']}${RL}))",
     'R': f"={c('S')}+{c('M')}+{c('Ti')}-{c('E')}-{c('To')}-{c('Xf')}",
     'Tact': f"={c('Tsup')}+{c('RSH')}*3.6/({N['rho']}*{N['cp']}*{c('S')})",
     'LC': f"={c('S')}*{N['rho']}*({c('W')}-{N['Wc']})*{N['hfg']}/3.6",
     'LOK': f"=IF({c('LC')}>={c('RLH')},\"ĐẠT\",\"KHÔNG ĐẠT\")",
     'LCM': f"={c('S')}*{N['rho']}*({c('WM')}-{N['Wc']})*{N['hfg']}/3.6",
     'SPT': f"=IF({c('M')}>0,MAX(0,{c('LCM')}-{c('RLH')})/({c('M')}*{N['rho']}*({N['Wo']}-{c('WM')})*{N['hfg']}/3.6),\"–\")",
     'Zp': f"={c('OAp')}/{c('S')}",
     'Qwin': f"={c('wl')}*{c('H')}*{N['Uw']}*({N['Twr']}-{N['Tsurw']})+{c('wlA')}*{c('H')}*{N['Uw']}*({N['Twr']}-{N['Tw']})+{c('A')}*{N['Uc']}*({N['Twr']}-{N['Tplw']})",
     'Qrm': f"=({c('RSH')}+{c('RLH')})/1000", 'Wm2': f"={c('Qrm')}*1000/{c('A')}",
     'SCR': f"=IF({c('Eac')}>0,\"Có\",\"Không\")", 'nhf': f"={c('nfh')}",
    }
    for k, v in V.items():
        t[f"{c(k)}"] = v
    for k in ('A', 'ppl', 'wl', 'wlA', 'nfh', 'nac', 'ncan', 'ncab', 'nsnk', 'narm', 'nins', 'ahu', 'grp'):
        t[c(k)].fill = INP
SUMK = ['A', 'V', 'ppl', 'eq', 'Qw', 'Qc', 'Ql', 'Qps', 'Qe', 'RSH', 'Qpl', 'Qtl', 'RLH', 'nfh', 'nac', 'ncan', 'ncab', 'nsnk', 'narm', 'nins',
        'Efh', 'Eac', 'Eor', 'Ecan', 'Ecab', 'Esnk', 'Earm', 'Eins', 'E', 'Eeaf', 'M', 'Ti', 'To', 'Net', 'Xf', 'OAp', 'Ss', 'Sl', 'Sa', 'S', 'R', 'LC', 'LCM', 'Qwin', 'Qrm', 'nhf']
t[f"{CI['name']}{RT}"] = 'TỔNG'
for k in SUMK:
    t[f"{CI[k]}{RT}"] = f"=SUM({CI[k]}{R0}:{CI[k]}{RL})"
t[f"{CI['ACH']}{RT}"] = f"={CI['S']}{RT}/{CI['V']}{RT}"
t[f"{CI['Wm2']}{RT}"] = f"={CI['Qrm']}{RT}*1000/{CI['A']}{RT}"
for c_ in t[RT]: c_.font = B; c_.fill = TOT
FMT = {'A': '0.00', 'V': '0.0', 'RH': '0%', 'RHM': '0%', 'Ps': '0.000', 'W': '0.00000', 'WM': '0.00000', 'h': '0.00', 'eq': '0.00',
       'Tsup': '0.00', 'ACH': '0.0', 'ACHE': '0.0', 'Tact': '0.0', 'fm': '0%', 'SPT': '0.0%', 'Zp': '0.000', 'Qrm': '0.00', 'Wm2': '0'}
for r in range(R0, RT + 1):
    for k, _ in COLS:
        cc = t[f"{CI[k]}{r}"]; cc.border = BD
        cc.number_format = FMT.get(k, '#,##0')
t.freeze_panes = f"C{R0}"
t.column_dimensions['B'].width = 22
for i in range(3, len(COLS) + 1): t.column_dimensions[CL(i)].width = 11
# transfer table
t.cell(TR0 - 2, 1, 'GIÓ CHUYỂN (TRANSFER) – KHÔNG ÁP DỤNG: chủ trì duyệt FAFE = 100 % gió thải từng phòng (IFC/Phase 1), không miệng gió chuyển. Bảng giữ lại để kiểm tra (mọi giá trị = 0).').font = B
for i, h in enumerate(['STT', 'Từ phòng', 'Đến phòng', 'Lưu lượng (m³/h)', 'Đường đi', 'W phòng nguồn', 'W phòng nhận', 'Ẩn mang vào (W)']):
    c_ = t.cell(TR0 - 1, i + 1, h); c_.font = B; c_.fill = HDR; c_.border = BD
if not TRANS:
    t.cell(TR0, 1, '–'); t.cell(TR0, 2, 'Không áp dụng'); t.cell(TR0, 3, 'Không áp dụng'); t.cell(TR0, 4, 0)
    t.cell(TR0, 5, 'Không có gió chuyển (FAFE = 100 % gió thải)'); t.cell(TR0, 6, 0); t.cell(TR0, 7, 0); t.cell(TR0, 8, 0)
for j, (fr, to, fq, path) in enumerate(TRANS):
    r = TR0 + j
    t.cell(r, 1, j + 1); t.cell(r, 2, fr); t.cell(r, 3, to); t.cell(r, 4, '=' + fq()); t.cell(r, 5, path)
    t.cell(r, 6, f"=INDEX(${CI['W']}${R0}:${CI['W']}${RL},MATCH(B{r},{NAMES},0))")
    t.cell(r, 7, f"=INDEX(${CI['W']}${R0}:${CI['W']}${RL},MATCH(C{r},{NAMES},0))")
    t.cell(r, 8, f"=D{r}*{N['rho']}*(F{r}-G{r})*{N['hfg']}/3.6")
    for k in range(1, 9): t.cell(r, k).border = BD
    t.cell(r, 4).number_format = '#,##0'; t.cell(r, 6).number_format = '0.00000'; t.cell(r, 7).number_format = '0.00000'; t.cell(r, 8).number_format = '0'
t.cell(TRL + 1, 3, 'TỔNG').font = B; t.cell(TRL + 1, 4, f"=SUM(D{TR0}:D{TRL})").number_format = '#,##0'
nr = TRL + 3
for note in [
    'Ghi chú: (1) Gió cấp thiết kế = MAX(tải hiện; tải ẩn; 6 ACH; Cấp − Hồi cần thiết), làm tròn lên 10 m³/h. (2) T gió cấp = T sau coil + ΔT nhiệt quạt của AHU phục vụ (sheet AHU).',
    '(3) Gió hồi = Cấp + Bù + Chuyển vào − Thải − Chuyển ra − Rò ra (gió tươi dư của AHU chia cho các phòng không phải phòng hóa theo tỷ lệ gió cấp). Chủ trì duyệt: FAFE = 100 % gió thải, không gió chuyển. (4) Gió bù FAFE không xử lý, cấp ngay trước tủ/chụp; SPL = 0 (như IFC/Phase 1) → không tính vào tải phòng/AHU.',
    '(5) “Tỷ lệ tràn FAFE chịu được” = phần gió bù chưa xử lý được phép lọt vào vùng phòng mà gió cấp vẫn giữ RH ≤ RH tối đa (chỉ xét ẩm). Phòng có tỷ lệ thấp cần bố trí miệng bù sát mặt tủ.',
    '(6) Instrument room: thải 1.200 m³/h được bù 100 % bằng gió tươi đã xử lý qua AHU-2.03 (không dùng FAFE) để giữ RH ≤ 55 %.',
    '(7) ACH > 40 ở Kjeldahl, ST1, Instrument do mật độ nhiệt thiết bị (0,5–0,9 kW/m²); ΔT cấp đã tối đa theo T sau coil 15 °C → không giảm được bằng ΔT. Xem báo cáo WP1 §4.4.',
    '(8) Số điểm hút khối phải theo bản vẽ DW-A-03-06 (WP2): ST2 6 hood + 2 tủ HC + 2 bồn rửa; Standard 19,72 4 hood + 1 bồn rửa; Standard 11,90 1 tủ HC. Chiều dài vách tiếp xúc theo 03_geometry/geometry.json.']:
    t.cell(nr, 2, note); nr += 1

# =====================================================================================
# Sheet AHU
# =====================================================================================
a = wb.create_sheet('AHU')
a['A1'] = 'TÍNH CHỌN AHU (DX) + DÀN NÓNG CDU + THIẾT BỊ PHỤ – PHƯƠNG ÁN CHỌN: 3 AHU (B)'; a['A1'].font = BT
a['A2'] = 'Cột G = Phương án A (2 AHU: Instrument gộp vào AHU-2.01) để so sánh. Các đại lượng cộng được lấy = AHU-2.01 + AHU-2.03.'
AH = ['AHU-2.01', 'AHU-2.02', 'AHU-2.03']
for i, h in enumerate(['Hạng mục', 'Đơn vị'] + AH + ['Tổng (PA B)', 'PA A: AHU-2.01 gộp Instrument', 'Ghi chú / công thức']):
    c_ = a.cell(4, i + 1, h); c_.font = B; c_.fill = HDR; c_.alignment = WR; c_.border = BD
TPA = 'TaiPhong!'
ROOMLIST = {ah: ', '.join(rm[0] for rm in ROOMS if rm[2] == ah) for ah in AH}
LSUP = {'AHU-2.01': 25, 'AHU-2.02': 20, 'AHU-2.03': 6}   # WP2 runs_for_takeoff (to farthest room centre) + ~3 m in room
LRET = {'AHU-2.01': 22, 'AHU-2.02': 18, 'AHU-2.03': 5}
LOA = {'AHU-2.01': 26.7, 'AHU-2.02': 4.7, 'AHU-2.03': 21.1}   # WP2: louver LV-2.0x on wall A -> AHU
LREF = {'AHU-2.01': 42.9, 'AHU-2.02': 19.5, 'AHU-2.03': 34.4}  # WP2: refrigerant AHU -> CDU yard
LCOND = {'AHU-2.01': 36.6, 'AHU-2.02': 13.1, 'AHU-2.03': 29.4}  # WP2: to CD-2.01 outside wall A (AHU-2.03 line incl. DH-2.01)
AROWS = []   # (key, label, unit, func(col, tag) -> formula/value, total('sum'|None), optA('add'|None|str), note, fmt)
def sumif(k):
    return lambda col, tag: f"=SUMIF({rng('ahu')},{col}$4,{rng(k)})"
def wavg(k):
    return lambda col, tag: f"=SUMPRODUCT(({rng('ahu')}={col}$4)*{rng('R')}*{rng(k)})/{col}{{Rsum}}"
AROWS += [
 ('rooms', 'Phòng phục vụ', '', lambda col, tag: ROOMLIST[tag], None, None, '', '@'),
 ('A', 'Diện tích', 'm²', sumif('A'), 'sum', 'add', '', '0.00'),
 ('ppl', 'Số người', 'người', sumif('ppl'), 'sum', 'add', '', '0'),
 ('S', 'Gió cấp (Σ phòng)', 'm³/h', sumif('S'), 'sum', 'add', 'TaiPhong cột GIÓ CẤP', '#,##0'),
 ('ACH', 'ACH cấp trung bình', '1/h', lambda col, tag: f"={col}{{S}}/({col}{{A}}*{N['H']})", None, 'f:={c}{{S}}/({c}{{A}}*' + N['H'] + ')', '', '0.0'),
 ('Rsum', 'Σ gió hồi các phòng (sau trừ rò ra)', 'm³/h', sumif('R'), 'sum', 'add', '', '#,##0'),
 ('Net', 'Σ (Thải − FAFE) các phòng = thải không có gió bù FAFE', 'm³/h', sumif('Net'), 'sum', 'add', 'Gió chuyển (nếu có) triệt tiêu trong khối', '#,##0'),
 ('Vou', 'Gió tươi theo người Σ(25 × người)', 'm³/h', sumif('OAp'), 'sum', 'add', 'TCVN 5687', '#,##0'),
 ('Xs', 'Xs = Vou / cấp', '-', lambda col, tag: f"={col}{{Vou}}/{col}{{S}}", None, 'f:={c}{{Vou}}/{c}{{S}}', 'ASHRAE 62.1 App. A', '0.000'),
 ('Zmax', 'Zp lớn nhất trong khối', '-', lambda col, tag: f"=SUMPRODUCT(MAX(({rng('ahu')}={col}$4)*{rng('Zp')}))", None, 'f:=MAX(C{Zmax},E{Zmax})', '', '0.000'),
 ('Ev', 'Hiệu quả thông gió hệ Ev = 1 + Xs − Zmax', '-', lambda col, tag: f"=MIN(1,1+{col}{{Xs}}-{col}{{Zmax}})", None, 'f:=MIN(1,1+{c}{{Xs}}-{c}{{Zmax}})', '', '0.000'),
 ('Vot', 'Gió tươi yêu cầu cho người Vot = Vou / Ev', 'm³/h', lambda col, tag: f"={col}{{Vou}}/{col}{{Ev}}", 'sum', 'f:={c}{{Vou}}/{c}{{Ev}}', '', '#,##0'),
 ('OA', 'GIÓ TƯƠI QUA AHU = MAX(Vot; Σ thải không có FAFE)', 'm³/h', lambda col, tag: f"=ROUNDUP(MAX({col}{{Vot}},{col}{{Net}})/10,0)*10", 'sum', 'f:=ROUNDUP(MAX({c}{{Vot}},{c}{{Net}})/10,0)*10', 'Gió bù thiếu hụt 10 % phải là gió tươi đã xử lý', '#,##0'),
 ('X', 'Dư áp khối (gió tươi dư thoát qua khe cửa, chia cho các phòng KHÔNG phải phòng hóa theo tỷ lệ gió cấp → phòng sạch dương áp, phòng hóa trung tính)', 'm³/h', lambda col, tag: f"={col}{{OA}}-{col}{{Net}}", 'sum', 'f:={c}{{OA}}-{c}{{Net}}', '≥ 0', '#,##0'),
 ('Ret', 'Gió hồi về AHU = Cấp − Tươi', 'm³/h', lambda col, tag: f"={col}{{S}}-{col}{{OA}}", 'sum', 'f:={c}{{S}}-{c}{{OA}}', '', '#,##0'),
 ('OAf', 'Tỷ lệ gió tươi', '-', lambda col, tag: f"={col}{{OA}}/{col}{{S}}", None, 'f:={c}{{OA}}/{c}{{S}}', '', '0.0%'),
 ('Tret', 'T gió hồi (bình quân theo lưu lượng hồi)', '°C', wavg('T'), None, None, '', '0.00'),
 ('Wret', 'W gió hồi', 'kg/kg', wavg('W'), None, None, '', '0.00000'),
 ('hret', 'h gió hồi', 'kJ/kg', lambda col, tag: f"=1.006*{col}{{Tret}}+{col}{{Wret}}*(2501+1.86*{col}{{Tret}})", None, None, '', '0.00'),
 ('Tmix', 'T hòa trộn vào coil', '°C', lambda col, tag: f"=({col}{{OA}}*{N['To']}+{col}{{Ret}}*{col}{{Tret}})/{col}{{S}}", None, None, '', '0.00'),
 ('Wmix', 'W hòa trộn', 'kg/kg', lambda col, tag: f"=({col}{{OA}}*{N['Wo']}+{col}{{Ret}}*{col}{{Wret}})/{col}{{S}}", None, None, '', '0.00000'),
 ('hmix', 'h hòa trộn', 'kJ/kg', lambda col, tag: f"=1.006*{col}{{Tmix}}+{col}{{Wmix}}*(2501+1.86*{col}{{Tmix}})", None, None, '', '0.00'),
 ('RHmix', 'RH vào coil', '-', lambda col, tag: f"={col}{{Wmix}}*{N['P']}/((0.622+{col}{{Wmix}})*0.61094*EXP(17.625*{col}{{Tmix}}/({col}{{Tmix}}+243.04)))", None, None, '', '0.0%'),
 ('Tc', 'T ra coil', '°C', lambda col, tag: f"={N['Tc']}", None, None, 'DB §2', '0.0'),
 ('RHc', 'RH ra coil', '-', lambda col, tag: f"={N['RHc']}", None, None, 'DB §2', '0%'),
 ('Wc', 'W ra coil', 'kg/kg', lambda col, tag: f"={N['Wc']}", None, None, '', '0.00000'),
 ('hc', 'h ra coil', 'kJ/kg', lambda col, tag: f"={N['hc']}", None, None, '', '0.00'),
 ('Qt', 'TẢI LẠNH COIL – TỔNG', 'kW', lambda col, tag: f"={col}{{S}}*{N['rho']}/3600*({col}{{hmix}}-{col}{{hc}})", 'sum', 'add', 'S·ρ·(h_mix − h_c)', '0.0'),
 ('Qs', 'Tải lạnh coil – hiện', 'kW', lambda col, tag: f"={col}{{S}}*{N['rho']}*{N['cp']}/3600*({col}{{Tmix}}-{col}{{Tc}})", 'sum', 'add', '', '0.0'),
 ('Ql', 'Tải lạnh coil – ẩn', 'kW', lambda col, tag: f"={col}{{Qt}}-{col}{{Qs}}", 'sum', 'add', '', '0.0'),
 ('SHR', 'SHR coil', '-', lambda col, tag: f"={col}{{Qs}}/{col}{{Qt}}", None, 'f:={c}{{Qs}}/{c}{{Qt}}', '', '0.00'),
 ('QOA', '…trong đó tải gió tươi', 'kW', lambda col, tag: f"={col}{{OA}}*{N['rho']}/3600*({N['ho']}-{col}{{hc}})", 'sum', 'add', '', '0.0'),
 ('QOAx', '…trong đó do bù thiếu hụt (OA − Vou)', 'kW', lambda col, tag: f"=MAX(0,{col}{{OA}}-{col}{{Vou}})*{N['rho']}/3600*({N['ho']}-{col}{{hc}})", 'sum', 'add', 'Chủ yếu: bù thải Instrument qua AHU-2.03; còn lại do làm tròn/Ev', '0.0'),
 ('RSH', 'Σ RSH phòng (tham khảo)', 'kW', lambda col, tag: f"=SUMIF({rng('ahu')},{col}$4,{rng('RSH')})/1000", 'sum', 'add', '', '0.0'),
 ('RLH', 'Σ RLH phòng (tham khảo)', 'kW', lambda col, tag: f"=SUMIF({rng('ahu')},{col}$4,{rng('RLH')})/1000", 'sum', 'add', '', '0.0'),
 ('Qsel', 'CÔNG SUẤT LẠNH CHỌN (+ dự phòng)', 'kW', lambda col, tag: f"={col}{{Qt}}*(1+{N['SEL']})", 'sum', 'add', '', '0.0'),
 ('Qssel', 'Công suất lạnh hiện chọn', 'kW', lambda col, tag: f"={col}{{Qs}}*(1+{N['SEL']})", 'sum', 'add', '', '0.0'),
 ('TR', 'Công suất chọn', 'TR', lambda col, tag: f"={col}{{Qsel}}/3.517", 'sum', 'add', '', '0.0'),
 ('BTU', 'Công suất chọn', 'BTU/h', lambda col, tag: f"={col}{{Qsel}}*3412.14", 'sum', 'add', '', '#,##0'),
 ('cond', 'Lượng nước ngưng (thiết kế)', 'kg/h', lambda col, tag: f"={col}{{S}}*{N['rho']}*({col}{{Wmix}}-{col}{{Wc}})", 'sum', 'add', 'S·ρ·(W_mix − W_c)', '0.0'),
 ('condd', 'Lượng nước ngưng / ngày (24 h đỉnh)', 'L/ngày', lambda col, tag: f"={col}{{cond}}*24", 'sum', 'add', '', '#,##0'),
 ('Acoil', 'Diện tích mặt coil/lọc (@v mặt)', 'm²', lambda col, tag: f"={col}{{S}}/3600/{N['vCoil']}", 'sum', 'add', '', '0.00'),
 ('Aduct', 'Tiết diện cổ cấp/hồi (@v ống chính)', 'm²', lambda col, tag: f"={col}{{S}}/3600/{N['vDuct']}", None, 'f:={c}{{S}}/3600/' + N['vDuct'], '', '0.000'),
 ('Lsup', 'Chiều dài tuyến cấp xa nhất (ước tính)', 'm', lambda col, tag: LSUP[tag], None, None, 'WP2 runs_for_takeoff + 3 m – WP3 kiểm lại', '0'),
 ('Lret', 'Chiều dài tuyến hồi xa nhất (ước tính)', 'm', lambda col, tag: LRET[tag], None, None, 'GĐ – WP3 kiểm lại', '0'),
 ('Loa', 'Chiều dài ống gió tươi louver → AHU', 'm', lambda col, tag: LOA[tag], None, None, 'WP2 geometry.json', '0.0'),
 ('ESPe', 'ESP ước tính = (Lc + MAX(Lh; Ltươi))·f·(1+k) + đầu cuối', 'Pa', lambda col, tag: f"=({col}{{Lsup}}+MAX({col}{{Lret}},{col}{{Loa}}))*{N['fr']}*(1+{N['kfit']})+{N['dTerm']}", None, None, '', '0'),
 ('ESP', 'ESP THIẾT KẾ = MAX(ESPmin; làm tròn(ước tính × 1,25))', 'Pa', lambda col, tag: f"=MAX({N['ESPmin']},ROUNDUP({col}{{ESPe}}*(1+{N['ESPm']})/50,0)*50)", None, None, '', '0'),
 ('dPi', 'Tổn thất trong AHU (G4 + F8 + coil + sưởi + vỏ)', 'Pa', lambda col, tag: f"={N['dG4']}+{N['dF8']}+{N['dCoil']}+{N['dHeat']}+{N['dCas']}", None, None, '', '0'),
 ('dPt', 'Áp tổng quạt', 'Pa', lambda col, tag: f"={col}{{ESP}}+{col}{{dPi}}", None, None, '', '0'),
 ('dTf', 'ΔT nhiệt quạt (động cơ trong dòng khí)', 'K', lambda col, tag: f"={col}{{dPt}}/({N['rho']}*{N['cp']}*1000*{N['ETA']})", None, None, 'ΔP/(ρ·cp·η)', '0.00'),
 ('Tsup', 'T GIÓ CẤP = T sau coil + ΔT quạt', '°C', lambda col, tag: f"={col}{{Tc}}+{col}{{dTf}}", None, None, 'Hút qua coil (draw-through)', '0.00'),
 ('Pf', 'Công suất trục quạt', 'kW', lambda col, tag: f"={col}{{S}}/3600*{col}{{dPt}}/{N['ETA']}/1000", 'sum', None, '', '0.00'),
 ('Mf', 'Động cơ quạt chọn (IEC, ×1,15)', 'kW', lambda col, tag: '=' + motor(f"{col}{{Pf}}"), 'sum', None, 'Biến tần', '0.00'),
 ('Tmw', 'Mùa đông: T hòa trộn (phòng 18 °C, OA 9,8 °C)', '°C', lambda col, tag: f"=({col}{{OA}}*{N['Tw']}+{col}{{Ret}}*{N['Twr']})/{col}{{S}}", None, None, '', '0.00'),
 ('Qwoa', 'Mùa đông: sưởi gió tươi', 'kW', lambda col, tag: f"={col}{{OA}}*{N['rho']}*{N['cp']}/3600*({N['Twr']}-{N['Tw']})", 'sum', None, '', '0.00'),
 ('Qwen', 'Mùa đông: tổn thất bao che', 'kW', lambda col, tag: f"=SUMIF({rng('ahu')},{col}$4,{rng('Qwin')})/1000", 'sum', None, '', '0.00'),
 ('Qwsp', 'Mùa đông: FAFE tràn vào phòng (SPLw)', 'kW', lambda col, tag: f"={N['SPLw']}*SUMIF({rng('ahu')},{col}$4,{rng('M')})*{N['rho']}*{N['cp']}/3600*({N['Twr']}-{N['Tw']})", 'sum', None, 'Không kể nhiệt thiết bị (khởi động sáng)', '0.00'),
 ('Qw', 'Nhu cầu sưởi mùa đông ×(1+SF)', 'kW', lambda col, tag: f"=({col}{{Qwoa}}+{col}{{Qwen}}+{col}{{Qwsp}})*(1+{N['SF']})", 'sum', None, '', '0.00'),
 ('Qrh', 'Nhu cầu gia nhiệt lại tối thiểu (ΔT_rh × cấp)', 'kW', lambda col, tag: f"={col}{{S}}*{N['rho']}*{N['cp']}/3600*{N['dTrh']}", 'sum', None, '', '0.00'),
 ('Hs', 'SƯỞI ĐIỆN CHỌN', 'kW', lambda col, tag: f"=ROUNDUP(MAX({col}{{Qw}},{col}{{Qrh}})/{N['Hstep']},0)*{N['Hstep']}", 'sum', None, 'Bậc 3 kW, SCR/bậc', '0'),
 ('Two', 'Mùa đông: T sau sưởi (đầy tải)', '°C', lambda col, tag: f"={col}{{Tmw}}+{col}{{Hs}}/({col}{{S}}*{N['rho']}*{N['cp']}/3600)", None, None, '', '0.00'),
 ('Twq', 'Mùa đông: T cấp yêu cầu để phòng ≥ 18 °C', '°C', lambda col, tag: f"={N['Twr']}+({col}{{Qwen}}+{col}{{Qwsp}})*(1+{N['SF']})/({col}{{S}}*{N['rho']}*{N['cp']}/3600)", None, None, '', '0.00'),
 ('dTh', 'Khả năng gia nhiệt lại của sưởi', 'K', lambda col, tag: f"={col}{{Hs}}/({col}{{S}}*{N['rho']}*{N['cp']}/3600)", None, None, '', '0.00'),
 ('Pahu', 'Điện AHU (quạt + sưởi)', 'kW', lambda col, tag: f"={col}{{Mf}}+{col}{{Hs}}", 'sum', None, '380 V/3P/50 Hz', '0.00'),
 ('Iahu', 'Dòng định mức AHU (ước tính)', 'A', lambda col, tag: f"={fla3(col + '{Mf}')}+{col}{{Hs}}*1000/(SQRT(3)*{N['V3']})", 'sum', None, '', '0.0'),
 ('Lref', 'Chiều dài ống gas AHU → bãi CDU (một chiều)', 'm', lambda col, tag: LREF[tag], None, None, 'WP2 geometry.json', '0.0'),
 ('kref', 'Hệ số suy giảm công suất do chiều dài ống gas', '-', lambda col, tag: f"=MAX(0.85,1-{N['kref']}*MAX(0,{col}{{Lref}}-{N['Lref0']}))", None, None, 'GĐ: −0,2 %/m sau 7,5 m', '0.000'),
 ('Ncdu', 'Số module CDU', 'bộ', lambda col, tag: f"=ROUNDUP({col}{{Qsel}}/{col}{{kref}}/{N['CDUmax']},0)", 'sum', None, 'Mỗi module 1 mạch coil riêng', '0'),
 ('Qcdu', 'Công suất lạnh thực / module @37,8 °C, sau suy giảm ống (yêu cầu)', 'kW', lambda col, tag: f"={col}{{Qsel}}/{col}{{Ncdu}}", None, None, 'Inverter, R410A/R32', '0.0'),
 ('Qcdun', 'Công suất danh định / module @37,8 °C, ống chuẩn (chọn máy)', 'kW', lambda col, tag: f"={col}{{Qcdu}}/{col}{{kref}}", None, None, 'Dùng để chọn model', '0.0'),
 ('Pcdu', 'Điện / module (danh định / COP)', 'kW', lambda col, tag: f"={col}{{Qcdun}}/{N['COP']}", None, None, '', '0.00'),
 ('Icdu', 'Dòng / module (ước tính)', 'A', lambda col, tag: f"={col}{{Pcdu}}*1000/(SQRT(3)*{N['V3']}*{N['PF']})", None, None, '380 V/3P', '0.0'),
 ('PcduT', 'Tổng điện CDU', 'kW', lambda col, tag: f"={col}{{Pcdu}}*{col}{{Ncdu}}", 'sum', None, '', '0.00'),
 ('Lc', 'Chiều dài ống nước ngưng tới điểm thoát (ước tính)', 'm', lambda col, tag: LCOND[tag], None, None, 'WP2: tới CD-2.01 ngoài tường A', '0'),
 ('Lmax', 'Chiều dài tự chảy tối đa = H_khả dụng / độ dốc', 'm', lambda col, tag: f"={N['Hav']}/{N['slope']}", None, None, '', '0.0'),
 ('CPn', 'Cần bơm nước ngưng?', '', lambda col, tag: f"=IF({col}{{Lc}}>{N['CPk']}*{col}{{Lmax}},\"Cần bơm\",\"Tự chảy\")", None, None, 'Bơm nếu L > 0,8·L_max', '@'),
]
AR = {}
r = 5
for it in AROWS:
    AR[it[0]] = r; r += 1
def sub(fs, col=None):
    out = fs
    for k, rr in AR.items():
        out = out.replace('{' + k + '}', str(rr))
    if col: out = out.replace('{c}', col)
    return out
for it in AROWS:
    key, lab, unit, fn, tot, optA, note, fmt = it
    rr = AR[key]
    a.cell(rr, 1, lab); a.cell(rr, 2, unit)
    for j, tag in enumerate(AH):
        col = CL(3 + j)
        v = fn(col, tag)
        if isinstance(v, str) and v.startswith('='):
            v = sub(v.replace('{{', '{').replace('}}', '}'))
        a.cell(rr, 3 + j, v)
        if key in ('Lsup', 'Lret', 'Lc'): a.cell(rr, 3 + j).fill = INP
    if tot == 'sum': a.cell(rr, 6, f"=SUM(C{rr}:E{rr})")
    if optA == 'add': a.cell(rr, 7, f"=C{rr}+E{rr}")
    elif isinstance(optA, str) and optA.startswith('f:'):
        a.cell(rr, 7, sub(optA[2:].replace('{{', '{').replace('}}', '}'), 'G'))
    a.cell(rr, 8, note)
    for k in range(1, 9):
        a.cell(rr, k).border = BD
        if 3 <= k <= 7: a.cell(rr, k).number_format = fmt
    if key in ('OA', 'Qt', 'Qsel', 'S', 'Hs', 'ESP', 'Tsup'):
        for k in range(1, 9): a.cell(rr, k).font = B
        for k in range(1, 9): a.cell(rr, k).fill = TOT
a.cell(AR['rooms'], 7, ROOMLIST['AHU-2.01'] + ', Instrument room')
# fill placeholder for TSUP row index into TaiPhong
for rr in range(R0, RL + 1):
    cc = t[f"{CI['Tsup']}{rr}"]
    cc.value = cc.value.replace('__TSUP__', str(AR['Tsup']))
    cx = t[f"{CI['Xf']}{rr}"]
    cx.value = cx.value.replace('__XR__', str(AR['X'])).replace('__SR__', str(AR['S']))
AEND = r
# option comparison extra rows
r = AEND + 1
a.cell(r, 1, 'SO SÁNH PHÂN KHỐI').font = B; r += 1
AR['shareI'] = r
a.cell(r, 1, 'Tỷ trọng gió cấp Instrument trong AHU-2.01 nếu gộp (PA A)'); a.cell(r, 2, '-')
a.cell(r, 7, f"=E{AR['S']}/G{AR['S']}").number_format = '0.0%'; r += 1
AR['shareIq'] = r
a.cell(r, 1, 'Tỷ trọng tải lạnh Instrument trong AHU-2.01 nếu gộp (PA A)'); a.cell(r, 2, '-')
a.cell(r, 7, f"=E{AR['Qt']}/G{AR['Qt']}").number_format = '0.0%'; r += 1
for txt in [
    'Kết luận: chọn PA B (3 AHU). Lý do: (1) Instrument chiếm ~1/3 gió cấp và tải của khối trái, chạy 24/7 trong khi các phòng còn lại chạy giờ hành chính;',
    '(2) RH ≤ 55 % và điều khiển nhiệt độ riêng – một AHU một vùng không giữ được đồng thời; (3) dự phòng: hỏng AHU-2.01 không làm dừng thiết bị phân tích;',
    '(4) tổng gió và tải lạnh hai PA bằng nhau (cộng tuyến tính) nên PA B không làm tăng công suất lạnh; chi phí thêm 1 AHU + 1 bộ DDC được bù bằng ống gió ngắn hơn.']:
    a.cell(r, 1, txt); r += 1
r += 1
# ---- auxiliary: dehumidifier & condensate pumps
a.cell(r, 1, 'THIẾT BỊ PHỤ: MÁY HÚT ẨM DH-2.01 (Instrument room) – định cỡ cho mùa nồm khi AHU-2.03 chạy non tải').font = B
for k in range(1, 9): a.cell(r, k).fill = SEC
r += 1
IR = ROW['Instrument room']
DH = [
 ('dhV', 'Thể tích phòng', 'm³', f"=TaiPhong!{CI['V']}{IR}", '0.0'),
 ('dhP', 'Ẩm do người (giờ có người)', 'kg/ngày', f"=TaiPhong!{CI['ppl']}{IR}*{N['qpl']}/1000/{N['hfg']}*3600*{N['HRocc']}", '0.0'),
 ('dhI', 'Ẩm do thấm khí cửa khi nồm (24 h)', 'kg/ngày', f"={N['ACHinf']}*C{{dhV}}*{N['rho']}*({N['Wn']}-{N['Wi']})*24", '0.0'),
 ('dhQ', 'Năng suất yêu cầu @25 °C/55 % (×dự phòng)', 'L/ngày', f"=(C{{dhP}}+C{{dhI}})*{N['DHm']}", '0.0'),
 ('dhR', 'Năng suất định mức tương đương @30 °C/80 %', 'L/ngày', f"=C{{dhQ}}/{N['DHr']}", '0'),
 ('dhS', 'CHỌN (làm tròn lên 10 L/ngày)', 'L/ngày', f"=ROUNDUP(C{{dhR}}/10,0)*10", '0'),
 ('dhW', 'Công suất điện ước tính', 'kW', f"=C{{dhS}}/24/{N['DHef']}", '0.00'),
 ('dhH', 'Nhiệt thải vào phòng khi chạy (điện + ẩn ngưng tụ) – không cộng ở điểm thiết kế hè', 'kW', f"=C{{dhW}}+C{{dhQ}}/24*{N['hfg']}/3600", '0.00'),
 ('dhC', 'Nước ngưng', 'L/h', f"=C{{dhQ}}/24", '0.00'),
]
for key, lab, unit, f, fmt in DH:
    AR[key] = r; r += 1
for key, lab, unit, f, fmt in DH:
    rr = AR[key]; a.cell(rr, 1, lab); a.cell(rr, 2, unit); a.cell(rr, 3, sub(f.replace('{{', '{').replace('}}', '}'))).number_format = fmt
    for k in range(1, 4): a.cell(rr, k).border = BD
r += 1
a.cell(r, 1, 'BƠM NƯỚC NGƯNG').font = B
for k in range(1, 9): a.cell(r, k).fill = SEC
r += 1
for i, h in enumerate(['Tag', 'Phục vụ', 'Số lượng', 'Lưu lượng ngưng (L/h)', 'Lưu lượng bơm (L/h)', 'Cột áp (m)', 'Điện (kW)', 'Ghi chú']):
    c_ = a.cell(r, i + 1, h); c_.font = B; c_.fill = HDR; c_.border = BD
r += 1
AR['CP1'] = r
a.cell(r, 1, 'CP-2.01'); a.cell(r, 2, 'AHU-2.03 + DH-2.01')
a.cell(r, 3, f"=IF(E{AR['CPn']}=\"Cần bơm\",1,0)")
a.cell(r, 4, f"=E{AR['cond']}+C{AR['dhC']}"); a.cell(r, 5, f"=D{r}*{N['CPf']}"); a.cell(r, 6, f"={N['CPh']}"); a.cell(r, 7, f"={N['CPkw']}")
a.cell(r, 8, 'Bơm có bình chứa, phao báo mức cao → cắt cấp lạnh AHU; 220 V/1P'); r += 1
AR['CP2'] = r
a.cell(r, 1, 'CP-2.02'); a.cell(r, 2, 'AHU-2.01')
a.cell(r, 3, f"=IF(C{AR['CPn']}=\"Cần bơm\",1,0)")
a.cell(r, 4, f"=C{AR['cond']}"); a.cell(r, 5, f"=D{r}*{N['CPf']}"); a.cell(r, 6, f"={N['CPh']}"); a.cell(r, 7, f"={N['CPkw']}")
a.cell(r, 8, 'Như trên'); r += 1
AR['CP3'] = r
a.cell(r, 1, 'CP-2.03'); a.cell(r, 2, 'AHU-2.02')
a.cell(r, 3, f"=IF(D{AR['CPn']}=\"Cần bơm\",1,0)")
a.cell(r, 4, f"=D{AR['cond']}"); a.cell(r, 5, f"=D{r}*{N['CPf']}"); a.cell(r, 6, f"={N['CPh']}"); a.cell(r, 7, f"={N['CPkw']}")
a.cell(r, 8, 'Chỉ lắp nếu kết quả “Cần bơm”'); r += 1
for rr in (AR['CP1'], AR['CP2'], AR['CP3']):
    for k in range(1, 9): a.cell(rr, k).border = BD
    for k in (4, 5): a.cell(rr, k).number_format = '0.0'
a.column_dimensions['A'].width = 58; a.column_dimensions['B'].width = 9
for col in 'CDEF': a.column_dimensions[col].width = 15
a.column_dimensions['G'].width = 20; a.column_dimensions['H'].width = 45
a.freeze_panes = 'C5'

# =====================================================================================
# Sheet QuatHut_Bu (FAF, EAF, hood fans, scrubber)
# =====================================================================================
q = wb.create_sheet('QuatHut_Bu')
q['A1'] = 'QUẠT CẤP BÙ FAF, QUẠT HÚT EAF, QUẠT TỦ HOOD, SCRUBBER'; q['A1'].font = BT
q['A2'] = 'Áp tổng quạt FAF = ESP (IFC 400 Pa) + hộp lọc G4+F8 (trung bình). EAF: ESP 400 Pa. Công suất trục = Q·ΔP/(3600·η). Động cơ = cỡ IEC ≥ 1,15 × trục.'
def tp(k, room): return f"TaiPhong!${CI[k]}${ROW[room]}"
FAFS = [  # tag, rooms served, type, location (WP2 equipment_proposed)
 ('FAF-2.01', ['Digestion room', 'Kjeldahl room', 'Sample treatment 1', 'Sample preparation'], 'Hộp ly tâm (cabinet) + G4/F8, biến tần', 'Trên trần ST1 sát lối đi; gió tươi từ LV-2.01 (≈ 32 m)'),
 ('FAF-2.02', ['Sample treatment 2', 'Standard (19.72 m²)', 'Standard (11.90 m²)'], 'Hộp ly tâm (cabinet) + G4/F8, biến tần', 'Trên trần ST2 gần AHU-2.02; gió tươi từ LV-2.02 (≈ 9 m)'),
]
EAFS = [
 ('EAF-2.01', 'Chụp treo Digestion + Kjeldahl, tủ hóa chất ST1', [('Ecan', 'Digestion room'), ('Ecan', 'Kjeldahl room'), ('Ecab', 'Sample treatment 1')], 'SS304', 'Ly tâm, vỏ/cánh SS304, biến tần', 'Trên trần Kjeldahl sát lối đi → DP-2.01'),
 ('EAF-2.02', 'Tủ hóa chất + chụp bồn rửa ST2, bồn rửa Standard 19,72, tủ hóa chất Standard 11,90', [('Ecab', 'Sample treatment 2'), ('Esnk', 'Sample treatment 2'), ('Esnk', 'Standard (19.72 m²)'), ('Ecab', 'Standard (11.90 m²)')], 'SS304', 'Ly tâm, vỏ/cánh SS304, biến tần', 'Trên trần ST2 phía tường A → DP-2.02'),
 ('EAF-2.03', 'Arm hood Sample preparation + điểm hút Instrument', [('Earm', 'Sample preparation'), ('Eins', 'Instrument room')], 'SS304 (nhánh arm hood GI)', 'Ly tâm in-line SS304, biến tần', 'Trên vách Sample prep / Instrument → DP-2.03 (≈ 32 m)'),
]
r = 4
q.cell(r, 1, 'A. QUẠT CẤP BÙ FAF (gió tươi không xử lý lạnh, lọc G4+F8, cấp trước tủ/chụp; liên động theo quạt hút trong phòng, biến tần theo số tủ đang chạy)').font = B
for k in range(1, 15): q.cell(r, k).fill = SEC
r += 1
FH_ = ['Tag', 'Phòng', 'Gió thải phòng (m³/h)', 'Tỷ lệ bù', 'Lưu lượng (m³/h)', 'ESP (Pa)', 'Lọc G4+F8 (Pa)', 'Áp tổng (Pa)', 'Trục (kW)', 'Động cơ (kW)', 'Điện', 'FLA (A)', 'Loại', 'Cổ ống @8 m/s (m²)']
for i, h in enumerate(FH_):
    c_ = q.cell(r, i + 1, h); c_.font = B; c_.fill = HDR; c_.alignment = WR; c_.border = BD
r += 1
QR = {}
F0 = r
for tag, rms, typ, loc in FAFS:
    QR[tag] = r
    q.cell(r, 1, tag); q.cell(r, 2, ', '.join(rms))
    q.cell(r, 3, '=' + '+'.join(tp('E', x) for x in rms)); q.cell(r, 4, f"=E{r}/C{r}"); q.cell(r, 5, '=' + '+'.join(tp('M', x) for x in rms))
    q.cell(r, 6, f"={N['ESPf']}"); q.cell(r, 7, f"={N['dG4']}+{N['dF8']}"); q.cell(r, 8, f"=F{r}+G{r}")
    q.cell(r, 9, f"=E{r}/3600*H{r}/{N['ETA']}/1000"); q.cell(r, 10, '=' + motor(f"I{r}"))
    q.cell(r, 11, f"=IF(J{r}>={N['M3']},\"380V/3P\",\"220V/1P\")")
    q.cell(r, 12, f"=IF(J{r}>={N['M3']},{fla3(f'J{r}')},{fla1(f'J{r}')})")
    q.cell(r, 13, typ); q.cell(r, 14, f"=E{r}/3600/{N['vDuct']}")
    r += 1
F1 = r - 1
q.cell(r, 2, 'TỔNG FAF').font = B
for col in 'EIJ': q[f'{col}{r}'] = f"=SUM({col}{F0}:{col}{F1})"
QR['FAF_TOT'] = r; r += 2
q.cell(r, 1, 'B. QUẠT HÚT EAF (HVAC) – chụp treo, tủ hóa chất, bồn rửa, arm hood, điểm hút Instrument; đặt trên trần, xả ra ngoài qua weather cover + lưới chắn côn trùng').font = B
for k in range(1, 15): q.cell(r, k).fill = SEC
r += 1
EH_ = ['Tag', 'Phục vụ', 'Phòng', 'Vật liệu', 'Lưu lượng (m³/h)', 'ESP (Pa)', '', 'Áp tổng (Pa)', 'Trục (kW)', 'Động cơ (kW)', 'Điện', 'FLA (A)', 'Loại', 'Cổ ống tròn @9 m/s Ø (mm)']
for i, h in enumerate(EH_):
    c_ = q.cell(r, i + 1, h); c_.font = B; c_.fill = HDR; c_.alignment = WR; c_.border = BD
r += 1
E0 = r
for tag, desc, parts, mat, typ, loc in EAFS:
    QR[tag] = r
    q.cell(r, 1, tag); q.cell(r, 2, desc); q.cell(r, 3, ', '.join(sorted(set(p[1] for p in parts)))); q.cell(r, 4, mat)
    q.cell(r, 5, '=' + '+'.join(tp(k, rm) for k, rm in parts))
    q.cell(r, 6, f"={N['ESPe']}"); q.cell(r, 8, f"=F{r}")
    q.cell(r, 9, f"=E{r}/3600*H{r}/{N['ETA']}/1000"); q.cell(r, 10, '=' + motor(f"I{r}"))
    q.cell(r, 11, f"=IF(J{r}>={N['M3']},\"380V/3P\",\"220V/1P\")")
    q.cell(r, 12, f"=IF(J{r}>={N['M3']},{fla3(f'J{r}')},{fla1(f'J{r}')})")
    q.cell(r, 13, typ); q.cell(r, 14, f"=ROUNDUP(SQRT(4*E{r}/3600/9/PI())*1000/50,0)*50")
    r += 1
E1 = r - 1
q.cell(r, 2, 'TỔNG EAF').font = B
for col in 'EIJ': q[f'{col}{r}'] = f"=SUM({col}{E0}:{col}{E1})"
QR['EAF_TOT'] = r; r += 2
q.cell(r, 1, 'C. QUẠT TÍCH HỢP TỦ HOOD (thuộc gói nội thất) – HVAC chỉ cấp nguồn + điều khiển từ DB.HVAC').font = B
for k in range(1, 15): q.cell(r, k).fill = SEC
r += 1
for i, h in enumerate(['Tag', 'Phòng', 'Số tủ', '…acid → scrubber', '…hữu cơ → xả thẳng SS304', 'Lưu lượng (m³/h)', 'kW/tủ', 'Tổng kW', 'Điện', 'FLA/tủ (A)']):
    c_ = q.cell(r, i + 1, h); c_.font = B; c_.fill = HDR; c_.alignment = WR; c_.border = BD
r += 1
HF = [('FHF-2.01', 'Digestion room'), ('FHF-2.02', 'Sample treatment 1'), ('FHF-2.03', 'Sample treatment 2'), ('FHF-2.04', 'Standard (19.72 m²)')]
H0 = r
for tag, room in HF:
    QR[tag] = r
    q.cell(r, 1, tag); q.cell(r, 2, room); q.cell(r, 3, f"={tp('nfh', room)}"); q.cell(r, 4, f"={tp('nac', room)}"); q.cell(r, 5, f"=C{r}-D{r}")
    q.cell(r, 6, f"={tp('Efh', room)}"); q.cell(r, 7, f"={N['HOODW']}"); q.cell(r, 8, f"=C{r}*G{r}"); q.cell(r, 9, '380V/3P')
    q.cell(r, 10, '=' + fla3(f"G{r}"))
    r += 1
H1 = r - 1
q.cell(r, 2, 'TỔNG').font = B
for col in 'CDEFH': q[f'{col}{r}'] = f"=SUM({col}{H0}:{col}{H1})"
QR['HF_TOT'] = r; r += 2
q.cell(r, 1, 'D. SCRUBBER ƯỚT SCR-2.01 (4 hood acid; PP; tháp đệm; tuần hoàn NaOH; đặt ngoài nhà)').font = B
for k in range(1, 15): q.cell(r, k).fill = SEC
r += 1
SC = [
 ('Qa', 'Σ gió hood acid (Digestion 2, ST1 1, ST2 1)', 'm³/h', f"=SUM({rng('Eac')})", '#,##0'),
 ('Qd', 'Lưu lượng thiết kế = Qa × (1 + dự phòng), làm tròn 100', 'm³/h', f"=ROUNDUP(C{{Qa}}*(1+{N['SCRm']})/100,0)*100", '#,##0'),
 ('As', 'Tiết diện tháp @v khí', 'm²', f"=C{{Qd}}/3600/{N['vScr']}", '0.00'),
 ('Ds', 'Đường kính tháp', 'mm', f"=ROUNDUP(SQRT(4*C{{As}}/PI())*1000/50,0)*50", '0'),
 ('Ql', 'Lưu lượng tuần hoàn = Qd × L/G', 'm³/h', f"=C{{Qd}}*{N['LG']}/1000", '0.0'),
 ('Pp', 'Công suất trục bơm', 'kW', f"=1000*9.81*C{{Ql}}/3600*{N['Hp']}/{N['etaP']}/1000", '0.00'),
 ('Mp', 'Động cơ bơm tuần hoàn (PP/FRP, chịu hóa chất)', 'kW', '=' + motor('C{Pp}'), '0.00'),
 ('dP', 'Áp tổng quạt scrubber (scrubber + ống PP gom/xả)', 'Pa', f"={N['dScr']}+{N['dScrD']}", '0'),
 ('Pf', 'Công suất trục quạt', 'kW', f"=C{{Qd}}/3600*C{{dP}}/{N['ETA']}/1000", '0.00'),
 ('Mf', 'Động cơ quạt (vỏ FRP/PP, biến tần theo áp ống gom)', 'kW', '=' + motor('C{Pf}'), '0.00'),
 ('Md', 'Bơm định lượng NaOH + bộ điều khiển pH', 'kW', f"={N['DOSE']}", '0.00'),
 ('Pt', 'Tổng điện scrubber (380 V/3P)', 'kW', "=C{Mp}+C{Mf}+C{Md}", '0.00'),
 ('It', 'Dòng định mức tổng (ước tính)', 'A', '=' + fla3('C{Pt}'), '0.0'),
]
for key, *_ in SC:
    QR['s' + key] = r; r += 1
def subq(fs):
    out = fs
    for key, *_ in SC:
        out = out.replace('{' + key + '}', str(QR['s' + key]))
    return out
for key, lab, unit, f, fmt in SC:
    rr = QR['s' + key]
    q.cell(rr, 1, lab); q.cell(rr, 2, unit); q.cell(rr, 3, subq(f.replace('{{', '{').replace('}}', '}'))).number_format = fmt
    for k in range(1, 4): q.cell(rr, k).border = BD
r += 1
q.cell(r, 1, 'E. TỔNG HỢP GIÓ THẢI').font = B; r += 1
QR['sumE'] = r
q.cell(r, 1, 'Σ thải qua EAF + Σ hood (TaiPhong)'); q.cell(r, 3, f"=E{QR['EAF_TOT']}+F{QR['HF_TOT']}").number_format = '#,##0'; r += 1
QR['sumE2'] = r
q.cell(r, 1, 'Σ gió thải toàn khu (TaiPhong)'); q.cell(r, 3, f"=SUM({rng('E')})").number_format = '#,##0'; r += 1
fmtq = {3: '#,##0', 4: '0%', 5: '#,##0', 6: '0', 7: '0', 8: '0', 9: '0.00', 10: '0.00', 12: '0.0', 14: '0.000'}
for rr in list(range(F0, F1 + 2)):
    for k, fm in fmtq.items(): q.cell(rr, k).number_format = fm
    for k in range(1, 15): q.cell(rr, k).border = BD
for rr in list(range(E0, E1 + 2)):
    for k, fm in {5: '#,##0', 6: '0', 8: '0', 9: '0.00', 10: '0.00', 12: '0.0', 14: '0'}.items(): q.cell(rr, k).number_format = fm
    for k in range(1, 15): q.cell(rr, k).border = BD
for rr in list(range(H0, H1 + 2)):
    for k, fm in {6: '#,##0', 7: '0.00', 8: '0.00', 10: '0.0'}.items(): q.cell(rr, k).number_format = fm
    for k in range(1, 11): q.cell(rr, k).border = BD
q.column_dimensions['A'].width = 52; q.column_dimensions['B'].width = 28; q.column_dimensions['C'].width = 22
for col in 'DEFGHIJKL': q.column_dimensions[col].width = 12
q.column_dimensions['M'].width = 32; q.column_dimensions['N'].width = 14

# =====================================================================================
# Sheet LichThietBi (equipment schedule)
# =====================================================================================
L = wb.create_sheet('LichThietBi')
L['A1'] = 'LỊCH THIẾT BỊ HVAC – PHASE 2 (đầu vào WP4 điện, WP5 BOQ)'; L['A1'].font = BT
L['A2'] = 'Số liệu liên kết công thức tới AHU / QuatHut_Bu. “Điện (kW)” = công suất điện định mức mỗi thiết bị (động cơ IEC, sưởi, máy nén theo COP).'
LH = ['STT', 'Tag', 'Loại thiết bị', 'SL', 'Vị trí / phục vụ', 'Lưu lượng (m³/h)', 'ESP / áp (Pa)', 'Lạnh (kW)', 'Sưởi (kW)', 'Điện / bộ (kW)', 'Điện áp / pha', 'FLA / bộ (A)', 'Tổng điện (kW)', 'Ghi chú']
for i, h in enumerate(LH):
    c_ = L.cell(4, i + 1, h); c_.font = B; c_.fill = HDR; c_.alignment = WR; c_.border = BD
SCH = []
for j, tag in enumerate(AH):
    col = CL(3 + j)
    SCH.append([tag, 'AHU DX (coil DX, quạt plug EC/biến tần, lọc G4+F8, sưởi điện, máng nước ngưng inox)', '1', ROOMLIST[tag] + ' – đặt trên sàn kết cấu trên trần',
                f"=AHU!{col}{AR['S']}", f"=AHU!{col}{AR['ESP']}", f"=AHU!{col}{AR['Qsel']}", f"=AHU!{col}{AR['Hs']}", f"=AHU!{col}{AR['Pahu']}", '380V/3P',
                f"=AHU!{col}{AR['Iahu']}", f"=D{{r}}*J{{r}}", f'="OA "&TEXT(AHU!{col}{AR["OA"]},"#,##0")&" m³/h; hiện "&TEXT(AHU!{col}{AR["Qssel"]},"0.0")&" kW; quạt "&TEXT(AHU!{col}{AR["Mf"]},"0.0")&" kW; ngưng "&TEXT(AHU!{col}{AR["cond"]},"0.0")&" kg/h"'])
for j, tag in enumerate(AH):
    col = CL(3 + j); ct = 'CDU-2.0' + tag[-1]
    SCH.append([ct, 'Dàn nóng DX inverter (R410A/R32) ghép AHU – module', f"=AHU!{col}{AR['Ncdu']}", f'Bãi CDU ngoài tường A – cho {tag}', '', '', f"=AHU!{col}{AR['Qcdun']}", '',
                f"=AHU!{col}{AR['Pcdu']}", '380V/3P', f"=AHU!{col}{AR['Icdu']}", f"=D{{r}}*J{{r}}", f'="Danh định/module @37,8 °C; thực sau suy giảm ống gas "&TEXT(AHU!{col}{AR["Qcdu"]},"0.0")&" kW (L "&AHU!{col}{AR["Lref"]}&" m); COP 3,0"'])
for tag, rms, typ, loc in FAFS:
    rr = QR[tag]
    SCH.append([tag, 'Quạt cấp bù FAFE – ' + typ, '1', loc + '; phục vụ ' + ', '.join(rms), f"=QuatHut_Bu!E{rr}", f"=QuatHut_Bu!F{rr}", '', '', f"=QuatHut_Bu!J{rr}", f"=QuatHut_Bu!K{rr}",
                f"=QuatHut_Bu!L{rr}", f"=D{{r}}*J{{r}}", 'Gió tươi không xử lý; ESP ngoài hộp lọc; lọc G4+F8 thêm ' + str(100 + 185) + ' Pa; biến tần theo áp ống; van gió động cơ ON/OFF tại mỗi tủ/chụp liên động quạt tủ'])
for tag, desc, parts, mat, typ, loc in EAFS:
    rr = QR[tag]
    SCH.append([tag, 'Quạt hút EAF – ' + typ, '1', loc + '; phục vụ ' + desc, f"=QuatHut_Bu!E{rr}", f"=QuatHut_Bu!F{rr}", '', '', f"=QuatHut_Bu!J{rr}", f"=QuatHut_Bu!K{rr}",
                f"=QuatHut_Bu!L{rr}", f"=D{{r}}*J{{r}}", 'Vật liệu: ' + mat + '; xả ngoài + weather cover + lưới côn trùng'])
for tag, room in HF:
    rr = QR[tag]
    SCH.append([tag, 'Điểm cấp nguồn quạt tủ hood (quạt thuộc gói nội thất)', f"=QuatHut_Bu!C{rr}", room, f"=QuatHut_Bu!F{rr}", '', '', '', f"=QuatHut_Bu!G{rr}", '380V/3P',
                f"=QuatHut_Bu!J{rr}", f"=D{{r}}*J{{r}}", f'="Acid → SCR-2.01: "&QuatHut_Bu!D{rr}&"; hữu cơ xả thẳng: "&QuatHut_Bu!E{rr}'])
SCH.append(['SCR-2.01', 'Scrubber ướt tháp đệm PP, trọn bộ quạt FRP + bơm tuần hoàn + định lượng NaOH + tủ điều khiển', '1', 'Ngoài nhà – 4 hood acid',
            f"=QuatHut_Bu!C{QR['sQd']}", f"=QuatHut_Bu!C{QR['sdP']}", '', '', f"=QuatHut_Bu!C{QR['sPt']}", '380V/3P', f"=QuatHut_Bu!C{QR['sIt']}", f"=D{{r}}*J{{r}}",
            f'="Ø tháp "&QuatHut_Bu!C{QR["sDs"]}&" mm; tuần hoàn "&TEXT(QuatHut_Bu!C{QR["sQl"]},"0.0")&" m³/h; quạt "&QuatHut_Bu!C{QR["sMf"]}&" kW; bơm "&QuatHut_Bu!C{QR["sMp"]}&" kW"'])
SCH.append(['DH-2.01', 'Máy hút ẩm âm trần nối ống (bơm ngưng tích hợp)', '1', 'Instrument room', '', '', '', '', f"=AHU!C{AR['dhW']}", '220V/1P',
            '=' + fla1(f"AHU!C{AR['dhW']}"), f"=D{{r}}*J{{r}}", f'="Định mức "&AHU!C{AR["dhS"]}&" L/ngày @30 °C/80 % (≥ "&TEXT(AHU!C{AR["dhQ"]},"0")&" L/ngày @25 °C/55 %)"'])
for key, tag in (('CP1', 'CP-2.01'), ('CP2', 'CP-2.02'), ('CP3', 'CP-2.03')):
    rr = AR[key]
    SCH.append([tag, 'Bơm nước ngưng có bình chứa + phao báo mức cao', f"=AHU!C{rr}", f"=AHU!B{rr}", '', f'=AHU!F{rr}*9810', '', '', f"=AHU!G{rr}", '220V/1P',
                '=' + fla1(f"AHU!G{rr}"), f"=D{{r}}*J{{r}}", f'="Lưu lượng bơm ≥ "&TEXT(AHU!E{rr},"0")&" L/h; cột áp "&AHU!F{rr}&" m; SL = 0 nghĩa là không cần"'])
r = 5
LS0 = r
for i, row in enumerate(SCH):
    L.cell(r, 1, i + 1)
    for k, v in enumerate(row):
        if isinstance(v, str): v = v.replace('{r}', str(r))
        if k == 2 and isinstance(v, str) and not v.startswith('='): v = int(v)
        L.cell(r, 2 + k, v)
    for k in range(1, 15): L.cell(r, k).border = BD
    for k, fm in {4: '0', 6: '#,##0', 7: '0', 8: '0.0', 9: '0', 10: '0.00', 12: '0.0', 13: '0.00'}.items(): L.cell(r, k).number_format = fm
    r += 1
LS1 = r - 1
L.cell(r, 3, 'TỔNG CÔNG SUẤT ĐIỆN ĐẶT HVAC (cột M) / Σ LẠNH CHỌN AHU (cột H)').font = B
L.cell(r, 13, f"=SUM(M{LS0}:M{LS1})").number_format = '0.0'; L.cell(r, 8, f"=SUMPRODUCT(D{LS0}:D{LS0+2},H{LS0}:H{LS0+2})").number_format = '0.0'
LSTOT = r
L.column_dimensions['B'].width = 11; L.column_dimensions['C'].width = 48; L.column_dimensions['E'].width = 40; L.column_dimensions['N'].width = 60
for col in 'DFGHIJKLM': L.column_dimensions[col].width = 11
L.freeze_panes = 'C5'

# =====================================================================================
# Sheet KiemTra (self checks)
# =====================================================================================
k_ = wb.create_sheet('KiemTra')
k_['A1'] = 'TỰ KIỂM TRA (mọi dòng phải ĐẠT)'; k_['A1'].font = BT
for i, h in enumerate(['STT', 'Hạng mục kiểm tra', 'Giá trị / sai lệch', 'Tiêu chí', 'Kết quả']):
    c_ = k_.cell(3, i + 1, h); c_.font = B; c_.fill = HDR; c_.border = BD
CHK = []
TP_ = 'TaiPhong!'
for rm in ROOMS:
    rr = ROW[rm[0]]; cc = lambda k: f"{TP_}{CI[k]}{rr}"
    CHK.append((f"Cân bằng khối lượng – {rm[0]}: Cấp + Bù + Chuyển vào − Hồi − Thải − Chuyển ra − Rò ra",
                f"={cc('S')}+{cc('M')}+{cc('Ti')}-{cc('R')}-{cc('E')}-{cc('To')}-{cc('Xf')}", '|x| < 0,5 m³/h', 'ABS(C{r})<0.5'))
    CHK.append((f"Gió hồi ≥ 0 – {rm[0]}", f"={cc('R')}", '≥ 0', 'C{r}>=0'))
    CHK.append((f"ACH ≥ 6 – {rm[0]}", f"={cc('ACH')}", '≥ ACHmin', f"C{{r}}>={N['ACHmin']}-0.001"))
    CHK.append((f"ACH ≤ 40 hoặc do tải hiện (ΔT cấp tối đa) – {rm[0]}", f"={cc('ACH')}", '≤ 40 hoặc yếu tố = Tải hiện',
                f"OR(C{{r}}<={N['ACHmax']},{cc('drv')}=\"Tải hiện\")"))
    CHK.append((f"Khử ẩm: gió cấp hấp thụ đủ RLH @RH thiết kế – {rm[0]}", f"={cc('LC')}-{cc('RLH')}", 'dư ≥ 0 W', 'C{r}>=0'))
    CHK.append((f"T phòng dự kiến trong 25 ±2 °C – {rm[0]}", f"={cc('Tact')}", '23…27 °C', 'AND(C{r}>=23,C{r}<=27)'))
CHK += [
 ('Σ gió chuyển vào = Σ gió chuyển ra', f"=SUM({rng('Ti')})-SUM({rng('To')})", 'Bằng 0', 'ABS(C{r})<0.5'),
 ('Gió chuyển chỉ trong cùng AHU (nguồn và đích cùng khối)', '=' + '+'.join(
     f"IF(INDEX({rng('ahu')},MATCH(TaiPhong!B{TR0+j},TaiPhong!{NAMES},0))=INDEX({rng('ahu')},MATCH(TaiPhong!C{TR0+j},TaiPhong!{NAMES},0)),0,1)" for j in range(len(TRANS))) if TRANS else '=0',
  '0 vi phạm', 'C{r}=0'),
 ('Σ diện tích = 441,76 m² (DB §3)', f"=SUM({rng('A')})", '441,76', 'ABS(C{r}-441.76)<0.01'),
 ('Σ số người = 75 (DB §3)', f"=SUM({rng('ppl')})", '75', 'C{r}=75'),
 ('Σ nhiệt thiết bị phòng = Σ ThietBi', f"=SUM({rng('eq')})*1000-{N['EQ_K_TOT']}", 'Bằng 0 (mọi dòng thiết bị gán đúng phòng)', 'ABS(C{r})<0.5'),
 ('Số tủ hood = 15, trong đó acid = 4 (DB §3)', f"=SUM({rng('nfh')})*100+SUM({rng('nac')})", '1504', 'C{r}=1504'),
 ('Σ gió thải = Σ EAF + Σ hood', f"=QuatHut_Bu!C{QR['sumE']}-QuatHut_Bu!C{QR['sumE2']}", 'Bằng 0', 'ABS(C{r})<0.5'),
 ('Σ FAF = Σ gió bù FAFE các phòng', f"=QuatHut_Bu!E{QR['FAF_TOT']}-SUM({rng('M')})", 'Bằng 0', 'ABS(C{r})<0.5'),
 ('Scrubber ≥ 1,1 × Σ hood acid', f"=QuatHut_Bu!C{QR['sQd']}/QuatHut_Bu!C{QR['sQa']}", '≥ 1,10', 'C{r}>=1.1-0.0001'),
 ('Σ hood acid = 9.080 m³/h (DB §4.3)', f"=QuatHut_Bu!C{QR['sQa']}", '9080', 'ABS(C{r}-9080)<0.5'),
]
for j, tag in enumerate(AH):
    col = CL(3 + j); A_ = lambda k: f"AHU!{col}{AR[k]}"
    CHK += [
     (f"{tag}: Cấp = Hồi về AHU + Tươi", f"={A_('S')}-{A_('Ret')}-{A_('OA')}", 'Bằng 0', 'ABS(C{r})<0.5'),
     (f"{tag}: Gió tươi ≥ Vot (người, ASHRAE 62.1 Ev)", f"={A_('OA')}-{A_('Vot')}", '≥ 0', 'C{r}>=-0.01'),
     (f"{tag}: Gió tươi ≥ Σ(Cấp − Hồi) phòng (dư áp khối ≥ 0)", f"={A_('X')}", '≥ 0', 'C{r}>=-0.01'),
     (f"{tag}: Σ hồi phòng = Hồi về AHU; Σ rò ra phòng = dư áp AHU", f"=ABS({A_('Rsum')}-{A_('Ret')})+ABS(SUMIF({rng('ahu')},\"{tag}\",{rng('Xf')})-{A_('X')})", 'Bằng 0', 'ABS(C{r})<0.5'),
     (f"{tag}: Coil khử ẩm (W vào > W ra) và RH vào ≤ 100 %", f"={A_('Wmix')}-{A_('Wc')}", '> 0', f"AND(C{{r}}>0,{A_('RHmix')}<=1)"),
     (f"{tag}: Tải coil ≥ Σ tải phòng (RSH + RLH)", f"={A_('Qt')}-{A_('RSH')}-{A_('RLH')}", '≥ 0', 'C{r}>=0'),
     (f"{tag}: CDU Σ module ≥ công suất chọn", f"={A_('Ncdu')}*{A_('Qcdu')}-{A_('Qsel')}", '≥ 0', 'C{r}>=-0.01'),
     (f"{tag}: CDU module (danh định) ≤ giới hạn module", f"={A_('Qcdun')}", f"≤ CDUmax", f"C{{r}}<={N['CDUmax']}+0.01"),
     (f"{tag}: Sưởi đông – T sau sưởi ≥ T cấp yêu cầu (phòng ≥ 18 °C)", f"={A_('Two')}-{A_('Twq')}", '≥ 0', 'C{r}>=0'),
     (f"{tag}: Động cơ quạt ≥ 1,15 × trục", f"={A_('Mf')}/{A_('Pf')}", '≥ 1,15', f"C{{r}}>={N['MM']}-0.0001"),
     (f"{tag}: ESP thiết kế ≥ ESP ước tính tuyến", f"={A_('ESP')}-{A_('ESPe')}", '≥ 0', 'C{r}>=0'),
     (f"{tag}: T cấp dùng trong TaiPhong = T sau coil + ΔT quạt", f"={A_('Tsup')}-{A_('Tc')}-{A_('dTf')}", 'Bằng 0', 'ABS(C{r})<0.0001'),
    ]
CHK += [
 ('Tổng PA B = PA A (gió cấp)', f"=F{AR['S']}-G{AR['S']}-D{AR['S']}", 'Bằng 0', 'ABS(C{r})<0.5'),
 ('Tổng PA B = PA A (tải coil)', f"=AHU!F{AR['Qt']}-AHU!G{AR['Qt']}-AHU!D{AR['Qt']}", 'Bằng 0', 'ABS(C{r})<0.01'),
 ('Σ tải coil = Σ AHU (TaiPhong ↔ AHU): Σ gió cấp phòng = Σ gió cấp AHU', f"=SUM({rng('S')})-AHU!F{AR['S']}", 'Bằng 0', 'ABS(C{r})<0.5'),
 ('Instrument: RH thiết kế ≤ 55 %', f"=TaiPhong!{CI['RH']}{IR}", '≤ 55 %', 'C{r}<=0.55'),
 ('Máy hút ẩm chọn ≥ yêu cầu quy đổi', f"=AHU!C{AR['dhS']}-AHU!C{AR['dhR']}", '≥ 0', 'C{r}>=0'),
 ('Bơm nước ngưng có ở nơi cần bơm', '=' + '+'.join(f"IF(AND(AHU!{CL(3+j)}{AR['CPn']}=\"Cần bơm\",AHU!C{AR[key]}<>1),1,0)" for j, key in enumerate(['CP2', 'CP3', 'CP1'])),
  '0 thiếu', 'C{r}=0'),
 ('FAF/EAF: động cơ ≥ 1,15 × trục (số vi phạm)', '=' + '+'.join(f"IF(QuatHut_Bu!J{QR[t_]}<{N['MM']}*QuatHut_Bu!I{QR[t_]},1,0)" for t_ in [f[0] for f in FAFS] + [e_[0] for e_ in EAFS]), '0', 'C{r}=0'),
 ('Scrubber: động cơ quạt & bơm ≥ 1,15 × trục', f"=MIN(QuatHut_Bu!C{QR['sMf']}/QuatHut_Bu!C{QR['sPf']},QuatHut_Bu!C{QR['sMp']}/QuatHut_Bu!C{QR['sPp']})", '≥ 1,15', f"C{{r}}>={N['MM']}-0.0001"),
 ('Lịch thiết bị: Σ lạnh AHU = Σ chọn (sheet AHU)', f"=SUMPRODUCT(LichThietBi!D{LS0}:D{LS0+2},LichThietBi!H{LS0}:H{LS0+2})-AHU!F{AR['Qsel']}", 'Bằng 0', 'ABS(C{r})<0.01'),
]
r = 4
for i, (lab, f, crit, test) in enumerate(CHK):
    k_.cell(r, 1, i + 1); k_.cell(r, 2, lab); k_.cell(r, 3, f); k_.cell(r, 4, crit)
    k_.cell(r, 5, f'=IF({test.replace("{r}", str(r))},"ĐẠT","KHÔNG ĐẠT")')
    k_.cell(r, 3).number_format = '#,##0.000'
    for kk in range(1, 6): k_.cell(r, kk).border = BD
    r += 1
KL = r - 1
r += 1
k_.cell(r, 2, 'SỐ HẠNG MỤC KHÔNG ĐẠT').font = B
k_.cell(r, 3, f'=COUNTIF(E4:E{KL},"KHÔNG ĐẠT")'); k_.cell(r, 5, f'=IF(C{r}=0,"TẤT CẢ ĐẠT","CÓ LỖI")'); k_.cell(r, 5).font = B
KSUM = r
k_.column_dimensions['B'].width = 80; k_.column_dimensions['C'].width = 16; k_.column_dimensions['D'].width = 34; k_.column_dimensions['E'].width = 14

# fix a typo-prone ref: in CHK 'Tổng PA B = PA A (gió cấp)' uses AHU sheet cells
for rr in range(4, KL + 1):
    v = k_.cell(rr, 3).value
    if isinstance(v, str) and v.startswith(f"=F{AR['S']}-G{AR['S']}"):
        k_.cell(rr, 3).value = f"=AHU!F{AR['S']}-AHU!G{AR['S']}-AHU!D{AR['S']}"

# guard: no text cell starting with '=' that is not a formula intended (all our formulas start with '=')
wb.save(XLSX)
print('written', XLSX)

# =====================================================================================
# Recalculate with LibreOffice, verify, export JSON
# =====================================================================================
def recalc(path):
    out = tempfile.mkdtemp(prefix='lo_')
    prof = tempfile.mkdtemp(prefix='lo_prof_')
    cmd = ['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', out, path]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=300)
    res = os.path.join(out, os.path.basename(path))
    shutil.copy(res, path)
    shutil.rmtree(out, ignore_errors=True); shutil.rmtree(prof, ignore_errors=True)

if shutil.which('soffice') and os.environ.get('NORECALC') != '1':
    recalc(XLSX)
    wv = openpyxl.load_workbook(XLSX, data_only=True)
    errs = []
    for ws in wv.worksheets:
        for row in ws.iter_rows():
            for c_ in row:
                if isinstance(c_.value, str) and c_.value.strip().startswith('#') and c_.value.strip().rstrip('!?0/').upper() in ('#VALUE', '#REF', '#DIV', '#NAME', '#N/A', '#NUM', '#NULL', '#ERR'):
                    errs.append(f'{ws.title}!{c_.coordinate}={c_.value}')
                if isinstance(c_.value, str) and any(x in c_.value for x in ('#VALUE!', '#REF!', '#DIV/0!', '#NAME?', '#N/A', 'Err:')):
                    errs.append(f'{ws.title}!{c_.coordinate}={c_.value}')
    print('errors:', len(errs), errs[:20])
    kv = wv['KiemTra']
    print('KiemTra:', kv.cell(KSUM, 3).value, kv.cell(KSUM, 5).value)
    for rr in range(4, KL + 1):
        if kv.cell(rr, 5).value != 'ĐẠT':
            print('  FAIL', kv.cell(rr, 2).value, kv.cell(rr, 3).value)

    if not EXPORT: sys.exit(0)
    # ---------------- export JSON ----------------
    tv = wv['TaiPhong']; av = wv['AHU']; qv = wv['QuatHut_Bu']; lv = wv['LichThietBi']
    def g(ws, ref):
        return ws[ref].value
    def rnd(x, n=1):
        return round(x, n) if isinstance(x, (int, float)) else x
    def rect(area_m2, hmax=600):
        # rectangular W x H (mm, 50 mm steps), 150 <= H <= hmax, H <= W <= 2.5 H; smallest section >= area
        best = None
        for hh in range(150, hmax + 1, 50):
            ww = max(hh, math.ceil(area_m2 * 1e6 / hh / 50) * 50)
            if ww > 2.5 * hh: continue
            cand = (ww * hh, ww - hh, ww, hh)
            if best is None or cand < best: best = cand
        if best is None:
            hh = hmax; ww = math.ceil(area_m2 * 1e6 / hh / 50) * 50; best = (0, 0, ww, hh)
        return f'{best[2]}x{best[3]}'
    # airflow_by_room
    trans = []
    for j in range(len(TRANS)):
        rr = TR0 + j
        trans.append({'from': tv.cell(rr, 2).value, 'to': tv.cell(rr, 3).value, 'flow_m3h': rnd(tv.cell(rr, 4).value, 0), 'path': tv.cell(rr, 5).value})
    rooms = []
    eaf_of = {}
    for tag, desc, parts, mat, typ, loc in EAFS:
        for _, rm in parts:
            if tag not in eaf_of.setdefault(rm, []): eaf_of[rm].append(tag)
    hf_of = {room: tag for tag, room in HF}
    for rm in ROOMS:
        rr = ROW[rm[0]]; v = lambda k: tv[f"{CI[k]}{rr}"].value
        rooms.append({
            'room': rm[0], 'group': rm[1], 'area_m2': v('A'), 'volume_m3': rnd(v('V'), 2), 'people': v('ppl'),
            'design_T_C': v('T'), 'design_RH': v('RH'), 'max_RH': v('RHM'),
            'serving_AHU': rm[2], 'supply_temp_C': rnd(v('Tsup'), 2),
            'sensible_load_W': rnd(v('RSH'), 0), 'latent_load_W': rnd(v('RLH'), 0), 'equipment_heat_kW': rnd(v('eq'), 3),
            'supply_m3h': rnd(v('S'), 0), 'supply_ACH': rnd(v('ACH'), 1), 'supply_driver': v('drv'),
            'return_m3h': rnd(v('R'), 0), 'exfiltration_m3h': rnd(v('Xf'), 0), 'oa_people_m3h': rnd(v('OAp'), 0),
            'oa_via_AHU_note': 'Gió tươi được trộn tại AHU; phòng nhận gió cấp hỗn hợp. Lượng tươi theo AHU xem equipment_schedule.',
            'exhaust_m3h': {'fume_hood_total': rnd(v('Efh'), 0), 'fume_hood_acid_to_scrubber': rnd(v('Eac'), 0), 'fume_hood_organic_direct': rnd(v('Eor'), 0),
                            'canopy': rnd(v('Ecan'), 0), 'chem_cabinet': rnd(v('Ecab'), 0), 'sink_hood': rnd(v('Esnk'), 0), 'arm_hood': rnd(v('Earm'), 0),
                            'instrument_vent': rnd(v('Eins'), 0), 'total': rnd(v('E'), 0), 'via_HVAC_EAF': rnd(v('Eeaf'), 0)},
            'counts': {'fume_hood': v('nfh'), 'fume_hood_acid': v('nac'), 'canopy': v('ncan'), 'chem_cabinet': v('ncab'), 'sink_hood': v('nsnk'), 'arm_hood': v('narm'), 'instrument_vent': v('nins')},
            'makeup_FAFE_m3h': rnd(v('M'), 0), 'makeup_ratio': v('fm'),
            'transfer_in_m3h': rnd(v('Ti'), 0), 'transfer_out_m3h': rnd(v('To'), 0),
            'transfer_in_from': [x['from'] for x in trans if x['to'] == rm[0]], 'transfer_out_to': [x['to'] for x in trans if x['from'] == rm[0]],
            'exhaust_minus_FAFE_m3h': rnd(v('Net'), 0),
            'serving_FAF': v('FAF') or None, 'serving_EAF': eaf_of.get(rm[0], []), 'hood_fan_power_points': hf_of.get(rm[0]),
            'scrubber': v('SCR') == 'Có', 'latent_check': v('LOK'), 'tolerable_FAFE_spill_fraction': rnd(v('SPT'), 4) if isinstance(v('SPT'), (int, float)) else None,
        })
    ahus = {}
    for j, tag in enumerate(AH):
        col = CL(3 + j)
        ahus[tag] = {k: rnd(av[f"{col}{AR[k]}"].value, 5 if k == 'Wmix' else 3 if k == 'RHmix' else 2) for k in ('S', 'OA', 'Ret', 'Vot', 'Net', 'X', 'Qt', 'Qs', 'Ql', 'Qsel', 'Qssel', 'Hs', 'ESP', 'dPt', 'Tsup', 'Mf', 'Pf', 'cond', 'Tmix', 'Wmix', 'RHmix', 'Ncdu', 'Qcdu', 'Qcdun', 'kref', 'Lref', 'Loa', 'Pcdu', 'Icdu', 'Pahu', 'Iahu', 'Acoil', 'Aduct', 'CPn', 'Lc')}
    json.dump({'project': 'SGS Hai Phong – H&N Phase 2 – HVAC', 'source': 'HVAC_Phase2_Calc_v4.xlsx (TaiPhong)', 'units': 'm3/h, W, kW, °C',
               'transfers': trans, 'rooms': rooms,
               'totals': {'supply_m3h': rnd(tv[f"{CI['S']}{RT}"].value, 0), 'exhaust_m3h': rnd(tv[f"{CI['E']}{RT}"].value, 0), 'makeup_FAFE_m3h': rnd(tv[f"{CI['M']}{RT}"].value, 0),
                          'oa_via_AHU_m3h': sum(ahus[t_]['OA'] for t_ in AH)}},
              open(os.path.join(HERE, 'airflow_by_room.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    # equipment_schedule
    eq = []
    for j, tag in enumerate(AH):
        d = ahus[tag]
        eq.append({'tag': tag, 'type': 'AHU DX (draw-through, quạt plug EC/biến tần, lọc G4+F8, coil DX, sưởi điện, máng ngưng inox)', 'qty': 1,
                   'location': 'Trên sàn kết cấu trên trần panel; phục vụ ' + ROOMLIST[tag],
                   'airflow_m3h': d['S'], 'outdoor_air_m3h': d['OA'], 'return_air_m3h': d['Ret'], 'esp_Pa': d['ESP'], 'fan_total_pressure_Pa': d['dPt'],
                   'cooling_total_kW_design': d['Qt'], 'cooling_total_kW_selected': d['Qsel'], 'cooling_sensible_kW_selected': d['Qssel'],
                   'coil_entering': {'T_C': d['Tmix'], 'W_kgkg': d['Wmix'], 'RH': d['RHmix']}, 'coil_leaving': {'T_C': 15.0, 'RH': 0.97},
                   'supply_temp_C': d['Tsup'], 'heater_kW': d['Hs'], 'fan_motor_kW': d['Mf'], 'fan_shaft_kW': d['Pf'], 'electrical_kW': d['Pahu'],
                   'voltage_phase': '380V/3P/50Hz', 'FLA_A_est': d['Iahu'], 'filters': 'G4 + F8', 'condensate_kg_h': d['cond'],
                   'coil_face_area_m2': d['Acoil'], 'duct_connection_mm': {'supply': rect(d['Aduct']), 'return': rect((d['Ret'] or 0) / 3600 / 8.0), 'outdoor_air': rect(d['OA'] / 3600 / 5.0, 400)},
                   'condensate_drain': d['CPn'], 'notes': 'Công suất lạnh chọn = tải coil × 1,10. ESP ước tính theo tuyến sơ bộ, WP3 kiểm lại.'})
    for j, tag in enumerate(AH):
        d = ahus[tag]
        eq.append({'tag': 'CDU-2.0' + tag[-1], 'type': 'Dàn nóng DX inverter R410A/R32 (module ghép AHU)', 'qty': int(d['Ncdu']), 'location': 'Ngoài nhà, móng BT (xây dựng) – cho ' + tag,
                   'cooling_kW_per_unit_nominal_at_37.8C': d['Qcdun'], 'cooling_kW_per_unit_actual_after_pipe_derate': d['Qcdu'], 'refrigerant_pipe_length_m': d['Lref'], 'pipe_derate_factor': d['kref'], 'electrical_kW_per_unit': d['Pcdu'], 'voltage_phase': '380V/3P/50Hz', 'FLA_A_est_per_unit': d['Icdu'],
                   'location': 'Bãi CDU ngoài tường A (phía nam ram dốc) – cho ' + tag,
                   'notes': f"Tổng thực {rnd(d['Qcdu']*d['Ncdu'],1)} kW ≥ {d['Qsel']} kW yêu cầu; COP 3,0; mỗi module một mạch coil riêng trong {tag}."})
    for tag, rms, typ, loc in FAFS:
        rr = QR[tag]
        Q_ = qv[f'E{rr}'].value
        eq.append({'tag': tag, 'type': 'Quạt cấp bù FAFE – ' + typ, 'qty': 1, 'location': loc, 'serves_rooms': rms,
                   'airflow_by_room_m3h': {x: rnd(tv[f"{CI['M']}{ROW[x]}"].value, 0) for x in rms},
                   'airflow_m3h': rnd(Q_, 0), 'esp_Pa': qv[f'F{rr}'].value, 'filter_box_Pa': qv[f'G{rr}'].value, 'fan_total_pressure_Pa': qv[f'H{rr}'].value,
                   'motor_kW': qv[f'J{rr}'].value, 'electrical_kW': qv[f'J{rr}'].value, 'voltage_phase': qv[f'K{rr}'].value, 'FLA_A_est': rnd(qv[f'L{rr}'].value, 1),
                   'duct_connection_mm': rect(qv[f'N{rr}'].value, 500), 'cooling_kW': 0, 'heater_kW': 0,
                   'notes': 'Không xử lý lạnh; = 100 % gió thải các phòng phục vụ (chủ trì duyệt, theo IFC/Phase 1); biến tần giữ áp ống; van gió động cơ tại mỗi miệng bù, mở theo tín hiệu chạy quạt tủ/chụp/EAF tương ứng.'})
    for tag, desc, parts, mat, typ, loc in EAFS:
        rr = QR[tag]
        eq.append({'tag': tag, 'type': 'Quạt hút EAF – ' + typ, 'qty': 1, 'location': loc, 'serves': desc,
                   'airflow_by_point_m3h': [{'room': rm_, 'type': k_x, 'm3h': rnd(tv[f"{CI[k_x]}{ROW[rm_]}"].value, 0)} for k_x, rm_ in parts], 'airflow_m3h': rnd(qv[f'E{rr}'].value, 0), 'esp_Pa': qv[f'F{rr}'].value,
                   'material': mat, 'motor_kW': qv[f'J{rr}'].value, 'electrical_kW': qv[f'J{rr}'].value, 'voltage_phase': qv[f'K{rr}'].value,
                   'FLA_A_est': rnd(qv[f'L{rr}'].value, 1), 'duct_connection_mm': f"Ø{qv[f'N{rr}'].value}",
                   'notes': 'Đặt trên trần, xả ra ngoài qua weather cover + lưới côn trùng; liên động FAF của phòng.'})
    for tag, room in HF:
        rr = QR[tag]
        eq.append({'tag': tag, 'type': 'Điểm cấp nguồn quạt tủ hood (quạt tích hợp, thuộc gói nội thất)', 'qty': int(qv[f'C{rr}'].value), 'location': room,
                   'airflow_m3h_total': rnd(qv[f'F{rr}'].value, 0), 'electrical_kW_per_unit': qv[f'G{rr}'].value, 'voltage_phase': '380V/3P',
                   'FLA_A_est_per_unit': rnd(qv[f'J{rr}'].value, 2),
                   'notes': f"Acid → SCR-2.01: {qv[f'D{rr}'].value}; hữu cơ (lọc than tại tủ) xả thẳng ống SS304: {qv[f'E{rr}'].value}. HVAC cấp nguồn + điều khiển, KHÔNG cung cấp quạt."})
    sv = lambda k: qv[f"C{QR['s'+k]}"].value
    eq.append({'tag': 'SCR-2.01', 'type': 'Scrubber ướt tháp đệm PP trọn bộ (quạt FRP/PP, bơm tuần hoàn, định lượng NaOH, pH, tủ điều khiển)', 'qty': 1,
               'location': 'Ngoài nhà – xử lý 4 hood acid (Digestion 2, ST1 1, ST2 1)', 'airflow_m3h': sv('Qd'), 'acid_hood_flow_m3h': sv('Qa'),
               'fan_total_pressure_Pa': sv('dP'), 'tower_diameter_mm': sv('Ds'), 'recirculation_m3h': rnd(sv('Ql'), 1),
               'fan_motor_kW': sv('Mf'), 'pump_motor_kW': sv('Mp'), 'dosing_kW': sv('Md'), 'electrical_kW': rnd(sv('Pt'), 2), 'voltage_phase': '380V/3P', 'FLA_A_est': rnd(sv('It'), 1),
               'duct_connection_mm': 'Ø' + str(math.ceil(math.sqrt(4 * sv('Qd') / 3600 / 9 / math.pi) * 1000 / 50) * 50) + ' (PP)',
               'notes': 'Nước cấp & đấu XLNT do gói khác (đầu chờ). Quạt tủ hood acid nối tiếp với quạt scrubber; quạt scrubber biến tần giữ áp âm ống gom.'})
    eq.append({'tag': 'DH-2.01', 'type': 'Máy hút ẩm âm trần nối ống, bơm ngưng tích hợp', 'qty': 1, 'location': 'Instrument room',
               'capacity_L_day_rated_30C_80': av[f"C{AR['dhS']}"].value, 'capacity_L_day_required_25C_55': rnd(av[f"C{AR['dhQ']}"].value, 1),
               'electrical_kW': rnd(av[f"C{AR['dhW']}"].value, 2), 'voltage_phase': '220V/1P', 'FLA_A_est': rnd(av[f"C{AR['dhW']}"].value * 1000 / (220 * 0.85 * 0.85), 1),
               'notes': 'Chạy theo RH phòng > 55 % (mùa nồm/non tải). Nước ngưng đấu vào ống ngưng AHU-2.03 (qua CP-2.01). Nước cấp không thuộc HVAC.'})
    for key, tag in (('CP1', 'CP-2.01'), ('CP2', 'CP-2.02'), ('CP3', 'CP-2.03')):
        rr = AR[key]
        qty = av[f'C{rr}'].value
        if qty and qty > 0:
            eq.append({'tag': tag, 'type': 'Bơm nước ngưng có bình chứa + phao báo mức cao (cắt lạnh AHU)', 'qty': int(qty), 'location': 'Cạnh ' + av[f'B{rr}'].value + ' trên trần',
                       'flow_L_h': rnd(av[f'E{rr}'].value, 0), 'head_m': av[f'F{rr}'].value, 'electrical_kW': av[f'G{rr}'].value, 'voltage_phase': '220V/1P',
                       'FLA_A_est': rnd(av[f'G{rr}'].value * 1000 / (220 * 0.85 * 0.85), 2), 'notes': f"Lượng ngưng thiết kế {rnd(av[f'D{rr}'].value,1)} L/h; tuyến > 0,8·L_max tự chảy."})
    json.dump({'project': 'SGS Hai Phong – H&N Phase 2 – HVAC', 'source': 'HVAC_Phase2_Calc_v4.xlsx', 'revision': 'v4 (WP1)',
               'total_installed_electrical_kW': rnd(lv[f'M{LSTOT}'].value, 1), 'total_cooling_selected_kW': rnd(av[f"F{AR['Qsel']}"].value, 1),
               'equipment': eq}, open(os.path.join(HERE, 'equipment_schedule.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('json written')
