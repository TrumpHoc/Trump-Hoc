#!/usr/bin/env python3
"""Dựng PTGT Buồng lạnh -20°C + Chiller 3x80HP theo skill 1-ptgt-mep (mẫu Tiến Đạt).

Chạy:  python3 _scripts/ptgt_build.py [vXX]
Đầu vào: _scripts/params.json, _cache/BANG_GIA_MEPF_v2.0_2026-11-12.csv,
         00. Tài Liệu/MAU_BG_LMT1_But_Son_2026.08.13.xlsx (lấy header + định dạng)
Đầu ra:  4. Khối lượng - BOQ/PTGT Buong lanh + Chiller - 2026.11.13_vXX.xlsx
"""
import copy, csv, json, math, sys
from pathlib import Path
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ROOT = Path(__file__).resolve().parents[1]
P = json.loads((ROOT / "_scripts/params.json").read_text(encoding="utf-8"))
VER = sys.argv[1] if len(sys.argv) > 1 else "v01"
OUT = ROOT / "4. Khối lượng - BOQ" / f"PTGT Buong lanh + Chiller - {P.get('ngay_file', '2026.11.13')}_{VER}.xlsx"
TPL = ROOT / "00. Tài Liệu/MAU_BG_LMT1_But_Son_2026.08.13.xlsx"
SRC_BG = "BẢNG GIÁ MEPF v2.0 12/11/2026"
SRC_SGC = "Bản chốt SGC 05/10/2026 (skill §6)"

# ---------------------------------------------------------------- bảng giá
CK = {("ỐNG THÉP MẠ KẼM", "Hòa Phát"): 0.178, ("ỐNG UPVC", "Tiền Phong"): 0.086}
PRICE = {}
with open(ROOT / "_cache/BANG_GIA_MEPF_v2.0_2026-11-12.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        PRICE[r["ma"]] = r


def bg(code):
    """Đơn giá áp dụng = giá gốc × (1 − CK nếu dòng chưa trừ CK); làm tròn trăm đồng."""
    r = PRICE[code]
    g = float(r["gia"])
    if r["ck"] == "N":
        g *= 1 - CK.get((r["nhom"], r["hang"]), 0)
    g = round(g, -2) if g >= 10000 else round(g, -1)
    src = f"{SRC_BG} – {code} ({r['hang']}, {r['ngay']}, loại {r['loai']}, tin cậy {r['hang_tc']})"
    return g, src


KA = P["K_lap_dat_thiet_bi_d_per_he_so"]
KB = P["K_ong_phu_kien_d_per_he_so"]
HH = P["hao_hut"]


def ceil_to(x, step):
    return round(math.ceil(x / step - 1e-9) * step, 2)


def L_upvc(L):
    return ceil_to(L * (1 + HH), P["cay_upvc_m"])


def L_cap(L):
    return ceil_to(L, 10)


def nc(hs):
    return round(hs * KB, -2)


def nca(hs):
    return round(hs * KA, -2)


NOI_SUY = []  # (hạng mục, cách tính, giá)
ROWS = []      # ('part'|'sec'|'item'|'text', ...)


def part(t, title):
    ROWS.append(dict(kind="part", tt=t, name=title))


def sec(t, title):
    ROWS.append(dict(kind="sec", tt=t, name=title))


def text(title):
    ROWS.append(dict(kind="text", name=title))


def item(name, spec, xx, dv, sl, q, r, src, st, note=""):
    """st: 'Có nguồn' | 'Quy đổi' | 'Lô' | 'Chỉ NC' | 'Cần hỏi giá' | 'Đề xuất NC'"""
    ROWS.append(dict(kind="item", name=name, spec=spec, xx=xx, dv=dv, sl=sl, q=q, r=r,
                     src=src, st=st, note=note))


# ---------------------------------------------------------------- giá dùng chung
def p_sgc(v):
    return v, SRC_SGC


# uPVC Tiền Phong C2 (báo giá NCC loại A)
UPVC = {60: bg("CT-01062"), 110: bg("CT-01066"), 140: bg("CT-01067"), 160: bg("CT-01068")}
CUT = {60: bg("CT-00096"), 110: bg("CT-00098"), 140: bg("CT-00099"), 160: bg("CT-00100")}
BICH = {110: bg("CT-00108"), 140: bg("CT-00109"), 160: bg("CT-00110")}
# bích D60: hồi quy hàm mũ Tiền Phong (skill §4.4) P = e^2,319·D^2,029
_b60 = round(math.exp(2.319) * 63 ** 2.029, -2)
BICH[60] = (_b60, "Nội suy hàm mũ bích PVC Tiền Phong P=e^2,319·D^2,029 (skill §4.4), D=63")
NOI_SUY.append(("Mặt bích uPVC D60", "P = e^2,319 × 63^2,029 (hồi quy Tiền Phong, khớp ±3%)", _b60))
TE_THU_110_60 = bg("CT-00344")
TE_THU_160 = bg("CT-00122")
CON_140_110 = bg("CT-00252")
CON_160_140 = bg("CT-00257")

# Van Wonil (Tuấn Hưng Phát – nguồn ưu tiên skill §3)
VB = {50: bg("CT-00445"), 100: bg("CT-00448"), 125: bg("CT-00449"), 150: bg("CT-00450")}
YL = {100: bg("CT-00887"), 125: bg("CT-00888")}
VMC = {100: bg("CT-00618"), 125: bg("CT-00619")}
# khớp mềm Shinyi – sát giá chốt SGC DN150 (1.653.000 vs 1.715.000)
KM = {50: bg("CT-00013"), 100: bg("CT-00016"), 125: bg("CT-00017")}
VBI = {25: bg("CT-00824"), 50: bg("CT-00827")}
DH = p_sgc(975000)
VBI15 = p_sgc(159000)
SIPHON = p_sgc(175000)
NHIET_KE = p_sgc(850000)
FLOW_SW = p_sgc(680000)
BL = {100: p_sgc(95000), 150: p_sgc(96500)}
_bl125 = round(95000 + (96500 - 95000) * (125 - 100) / 50, -2)
BL[125] = (_bl125, "Nội suy tuyến tính bộ bu lông DN100 (95.000) – DN150 (96.500), SGC")
NOI_SUY.append(("Bộ bu lông + gioăng DN125", "95.000 + (96.500−95.000)×(125−100)/(150−100)", _bl125))
BL[50] = (95000, "Quy đổi: lấy giá DN100 SGC (không có DN50) – độ tin cậy thấp")
NOI_SUY.append(("Bộ bu lông + gioăng DN50", "Không có dữ liệu DN50 → lấy DN100 (cỡ gần nhất có giá), độ tin cậy thấp", 95000))
GD_LON = p_sgc(550000)
GD_NHO = p_sgc(200000)

# bảo ôn Superlon
BO_42_32 = bg("HV-00098")
BO_19_19 = bg("HV-00060")
BO_114_25 = bg("HV-00090")
BO_140_25 = bg("HV-00091")
# D160 25mm: ngoại suy hàm mũ từ 114 & 140
_b = math.log(BO_140_25[0] / BO_114_25[0]) / math.log(140 / 114)
_bo160 = round(BO_140_25[0] * (160 / 140) ** _b, -2)
BO_160_25 = (_bo160, "Ngoại suy hàm mũ Superlon 25mm từ Ø114 & Ø140 (HV-00090/91)")
NOI_SUY.append(("Bảo ôn Superlon 25mm Ø160", f"y = 259.563×(160/140)^{_b:.3f} (ngoại suy, ngoài dải)", _bo160))
BANG_DINH = bg("HV-00225")
KEO_BO = bg("HV-00230")

# Ống đồng Hailiang cây cứng
DONG_HUT = bg("HV-00576")   # 41.28 x 1.2
DONG_LONG = bg("HV-00562")  # 19.05 x 1.0
GAS = bg("HV-00328")        # R404A 10.9 kg

# Điện
CAP = {"3x50+25": p_sgc(752650), "3x6+4": p_sgc(110645), "3x16+10": bg("DI-00093"),
       "4x4": bg("DI-00121"), "CV25": p_sgc(110419), "CV10": bg("DI-00025"), "CV4": p_sgc(17776),
       "TH2x1.5": p_sgc(25723)}
MANG_200 = bg("DI-00714")
MANG_300 = bg("DI-00716")
RUOT_GA = (12000, "Mẫu BG Bút Sơn 13/08/2026 – ống ruột gà D25")
MCCB = {15: bg("DI-01207"), 30: bg("DI-01209"), 50: bg("DI-01211"), 150: bg("DI-01216"),
        350: bg("DI-01243"), 500: bg("DI-01245")}
CONT = {12: bg("DI-01013"), 18: bg("DI-01014"), 32: bg("DI-01016"), 50: bg("DI-01018")}
RL = {18: bg("DI-01320"), 65: bg("DI-01322")}


# ---------------------------------------------------------------- helper hạng mục
def pipe(D, L, label="Ống uPVC Tiền Phong", spec="C2 (PN8)"):
    q, s = UPVC[D]
    item(f"{label} D{D}", spec, "Việt Nam", "m", L_upvc(L), q, nc(D / 100), s, "Có nguồn")


def fit(kind, D, sl, price, hs):
    q, s = price
    item(f"{kind}", "uPVC Tiền Phong", "Việt Nam", "cái", sl, q, nc(hs * D / 100), s,
         "Quy đổi" if "Nội suy" in s or "Quy đổi" in s else "Có nguồn")


def valve(name, DN, sl, price, hs, xx="Hàn Quốc"):
    q, s = price
    item(name, f"DN{DN}", xx, "cái", sl, q, nc(hs * DN / 100), s, "Có nguồn")


def acc(name, spec, sl, price, hs, dv="cái", xx="Hàn Quốc"):
    q, s = price
    st = "Quy đổi" if ("Nội suy" in s or "Quy đổi" in s) else "Có nguồn"
    item(name, spec, xx, dv, sl, q, nc(hs), s, st)


def pump_set(n, Ds, Dd, tag):
    """Bộ phụ kiện mỗi bơm theo quy tắc user (skill §2)."""
    text(f"Cụm van {tag} – {n} bộ (hút DN{Ds}, đẩy DN{Dd})")
    valve("Van bướm tay quay – phía hút", Ds, n, VB[Ds], 1)
    valve("Y lọc mặt bích – phía hút", Ds, n, YL[Ds], 1)
    valve("Khớp nối mềm cao su – phía hút", Ds, n, KM[Ds], 0.8, "Đài Loan")
    valve("Khớp nối mềm cao su – phía đẩy", Dd, n, KM[Dd], 0.8, "Đài Loan")
    valve("Van một chiều cánh bướm kép – phía đẩy", Dd, n, VMC[Dd], 1)
    valve("Van bướm tay quay – phía đẩy", Dd, n, VB[Dd], 1)
    acc("Đồng hồ áp suất 0–6 bar mặt 100", "Wise", 2 * n, DH, 0.4)
    acc("Van bi DN15 cho đồng hồ", "DN15", 2 * n, VBI15, 0.3)
    acc("Ống siphon đồng hồ phía đẩy", "DN15", n, SIPHON, 0.3)


def cable(name, key, L_each, n, cv=None):
    sl = L_cap(n * (L_each + 3) * 1.10)
    q, s = CAP[key]
    item(name, f"CXV {key} – Goldcup" if key[0].isdigit() else key, "Việt Nam", "m", sl, q,
         round(q * P["ty_le_NC_cap"], -2), s, "Có nguồn" if "SGC" in s else "Quy đổi",
         "Dài = n×(L+3)×1,10, làm tròn 10 m")
    if cv:
        q2, s2 = CAP[cv]
        item(f"Cáp tiếp địa {cv} đi kèm", "CV – Goldcup", "Việt Nam", "m", sl, q2,
             round(q2 * P["ty_le_NC_cap"], -2), s2, "Có nguồn" if "SGC" in s2 else "Quy đổi",
             "SL = cáp động lực tương ứng")


def ask(name, spec, dv, sl, r=0, note="Cần hỏi giá NCC"):
    item(name, spec, "", dv, sl, None, r, "CHƯA CÓ GIÁ trong bảng giá MEPF v2.0", "Cần hỏi giá", note)


# ================================================================ PHẦN I – BUỒNG LẠNH
B = P["buong_lanh"]
nC, nE = B["so_CDU"], B["so_dan_lanh"]
part("I", "HỆ THỐNG BUỒNG LẠNH -20°C (7 x 3,4 x 2,6 m)")
GC = P["gia_ghi_chep"]
HS = GC["he_so_dau_vao"]
SRC_GC = GC["nguon"]


def g(v):
    """Giá đầu vào = đơn giá ghi chép × hệ số (1,10), làm tròn nghìn."""
    return round(v * HS, -3)


def gc(name, spec, xx, dv, sl, vt, ncv, note=""):
    item(name, spec, xx, dv, sl, g(vt) if vt else 0, g(ncv) if ncv else 0, SRC_GC,
         "Lô" if dv == "Lô" else "Có nguồn",
         (note + "; " if note else "") + f"Đơn giá ghi chép × {HS:.2f}")


sec("I.A", "Vỏ buồng lạnh (panel PU 100mm 2 mặt inox, cửa)")
a, b, c = B["dai"], B["rong"], B["cao"]
S_panel = round(2 * (a * b + a * c + b * c), 2)   # 6 mặt: tường + trần + sàn
gc("Panel PU 100mm 2 mặt inox – 6 mặt (tường, trần, sàn)",
   f"2×({a}×{b}+{a}×{c}+{b}×{c}) = {S_panel} m²", "", "m2", S_panel, GC["panel_2mat_m2"], GC["nc_lap_vo_m2"],
   "KT giả định phủ bì, không trừ ô cửa")
ask("Cửa bản lề panel 100mm 2 mặt inox, có điện trở sưởi khung", "700 x 1700 mm", "bộ", B["so_cua"], 0,
    "Ghi chép chưa có giá cửa – cần hỏi NCC panel")
gc("Vật tư phụ lắp vỏ (silicone, keo bọt, ốc vít, rivet, đèn, van thông áp)", "Trọn bộ", "", "Lô", 1,
   GC["vt_phu_lap_vo_lo"], 0)
gc("Vận chuyển panel đến công trình", "", "", "Lô", 1, 0, GC["van_chuyen_panel_lo"])

sec("I.B", f"Cụm máy nén ngưng tụ 20HP giải nhiệt nước (TĐ lắp cụm) – {nC} cụm")
gc("Máy nén Panasonic 20HP", "", "", "cái", nC, GC["may_nen_20HP"], 0)
gc("Bình ngưng giải nhiệt nước 20HP", "", "", "cái", nC, GC["binh_ngung_20HP"], 0)
gc("Nhớt lạnh POE", "", "", "bộ", nC, GC["nhot_POE_bo"], 0, "Theo cụm")
gc("Phụ kiện lên máy (bình chứa, phin, mắt gas, van, rơ le áp...)", "", "", "bộ", nC, GC["phu_kien_len_may"], 0)
gc("Nhân công lắp ráp cụm máy", "", "", "cụm", nC, 0, GC["nc_len_cum"])

sec("I.C", "Dàn lạnh, tiết lưu, gas")
gc("Dàn lạnh DJ 14.3/85-201 (≈25HP)", "theo ghi chép", "", "cái", nE, GC["dan_lanh"], 0,
   "Model đọc từ ghi chép tay – cần xác nhận")
gc("Van tiết lưu TES5", "Danfoss", "", "bộ", nC, GC["tiet_luu_TES5"], 0)
gc("Gas lạnh R404A/R507", "", "", "hệ", nC, GC["gas_moi_he"], 0,
   "Ghi chép: 5.700.000/hệ nhưng tổng 178,1tr khớp với 4.500.000 – cần xác nhận")

sec("I.D", "Tháp giải nhiệt, bơm giải nhiệt")
gc("Tháp giải nhiệt 20RT", "", "", "bộ", B["so_thap"], GC["thap_20RT"], 0,
   "Phạm vi 1 tháp cho 5 CDU 20HP – 20RT chỉ đủ ~1 cụm, CẦN XÁC NHẬN công suất/số lượng")
gc("Bơm giải nhiệt 1,5kW (Ewara) – 1 chạy 1 dự phòng", "", "Ý", "cái", B["so_bom"], GC["bom_1_5kW"], 0,
   "1,5kW chỉ đủ ~1 cụm; 5 cụm cần ≈39 m³/h – CẦN XÁC NHẬN")

sec("I.E", "Đường ống gas, ống nước, giá đỡ")
Lg = nC * (B["L_CDU_buong_m"] + B["L_them_moi_tuyen_m"])
L_cu = round(Lg * (1 + HH))
gc("Ống đồng Ø42x1,0 + Ø16x0,8 + bảo ôn 42x25 + phụ kiện đồng (theo cặp ống)", "800.000 đ/m (ghi chép)",
   "", "m", L_cu, GC["ong_dong_cap_m"], 0,
   f"{nC} tuyến × ({B['L_CDU_buong_m']}+{B['L_them_moi_tuyen_m']}) m × 1,05")
gc("Ống PVC + phụ kiện (nước giải nhiệt, thoát nước xả đá)", "", "Việt Nam", "hệ", nC, GC["ong_pvc_moi_he"], 0)
item("Giá đỡ chung tuyến ống gas (thép hình, ty treo, kẹp)", "@1,5 m", "Việt Nam", "cái",
     math.ceil(B["L_CDU_buong_m"] / 1.5) + 2, GD_LON[0], nc(1), GD_LON[1], "Có nguồn")

sec("I.F", "Tủ điện, dây điện")
gc("Tủ điện tổng điều khiển buồng lạnh (đơn giá theo hệ)", "1 tủ cho 5 hệ", "Việt Nam", "hệ", nC,
   GC["tu_dk_moi_he"], 0)
gc("Dây điện, cáp động lực và điều khiển", "", "Việt Nam", "hệ", nC, GC["day_dien_moi_he"], 0)

sec("I.G", "Nhân công lắp đặt hệ thống")
gc("Nhân công lắp đặt (tháp, bơm, đường ống, điện, nạp gas, chạy thử)", "", "", "hệ", nC, 0,
   GC["nc_lap_dat_moi_he"])

# ================================================================ PHẦN II – CHILLER
C = P["chiller"]
n = C["so_chiller"]
part("II", "HỆ THỐNG CHILLER 03 x 80HP")
sec("II.A", "Lắp đặt, căn chỉnh thiết bị (thiết bị do CĐT/NCC cấp; tank 20m³ CĐT cấp)")
item("Lắp đặt, căn chỉnh chiller 80HP giải nhiệt nước", "Giá NC", "", "bộ", n, None, nca(10),
     "Hệ số 10 × K_A (SGC 2.685 kW = 30)", "Đề xuất NC")
item("Lắp đặt, căn chỉnh bơm tuần hoàn nước lạnh", "Giá NC", "", "bộ", n, None, nca(2),
     "Hệ số 2 × K_A", "Đề xuất NC")
item("Lắp đặt, căn chỉnh bơm giải nhiệt", "Giá NC", "", "bộ", n, None, nca(2), "Hệ số 2 × K_A", "Đề xuất NC")
item("Lắp đặt tháp giải nhiệt", "Giá NC", "", "bộ", n, None, nca(4), "Hệ số 4 × K_A", "Đề xuất NC")
item("Đấu nối vào tank 20m³ (tank CĐT cấp)", "Giá NC", "", "Lô", 1, None, nca(1), "Hệ số 1 × K_A", "Đề xuất NC")
item("Súc rửa, thử áp, chạy thử, cân chỉnh hệ thống", "Giá NC", "", "hệ", 1, None, nca(8),
     "Hệ số 8 × K_A", "Đề xuất NC")

sec("II.B", "Ống uPVC vòng nước lạnh và giải nhiệt (Tiền Phong C2)")
text(f"Nước lạnh {n}×{C['luu_luong_nuoc_lanh_m3h']} m³/h; giải nhiệt {n}×{C['luu_luong_giai_nhiet_m3h']} m³/h (giả định 80HP ≈ {C['cong_suat_lanh_kW_moi_may']} kW)")
pipe(160, C["L_header_hut_tank_m"] + C["L_header_ve_tank_m"], spec="C2 (PN8) – header hút/về tank")
pipe(140, n * C["L_nhanh_hut_bom_m"] + n * (2 * C["L_chiller_thap_m"] + 4),
     spec="C2 (PN8) – hút bơm lạnh + vòng giải nhiệt")
pipe(110, n * (C["L_bom_chiller_m"] + C["L_chiller_header_ve_m"]), spec="C2 (PN8) – bơm → chiller → header")
fit("Cút 90° uPVC D160", 160, 4, CUT[160], 0.6)
fit("Tê thu uPVC D160 (nhánh 140/110)", 160, 2 * n, TE_THU_160, 0.9)
fit("Côn thu uPVC D160/140", 160, n, CON_160_140, 0.6)
fit("Cút 90° uPVC D140", 140, 2 * n + 8 * n, CUT[140], 0.6)
fit("Cút 90° uPVC D110", 110, 4 * n, CUT[110], 0.6)
fit("Côn thu uPVC D140/110", 140, 2 * n, CON_140_110, 0.6)
fit("Mặt bích uPVC D160 (nối tank)", 160, 2, BICH[160], 0.6)
fit("Mặt bích uPVC D140", 140, n * (1 + 2 + 1 + 2), BICH[140], 0.6)
fit("Mặt bích uPVC D110", 110, n * (1 + 2 + 1), BICH[110], 0.6)
text("Bảo ôn ống nước lạnh (giả định nước lạnh 7/12°C)")
for D, L, bo in [(160, C["L_header_hut_tank_m"] + C["L_header_ve_tank_m"], BO_160_25),
                 (140, n * C["L_nhanh_hut_bom_m"], BO_140_25),
                 (110, n * (C["L_bom_chiller_m"] + C["L_chiller_header_ve_m"]), BO_114_25)]:
    item(f"Bảo ôn cao su lưu hóa dày 25mm cho ống D{D}", "Superlon", "Malaysia", "m",
         ceil_to(L * (1 + HH), P["cay_bao_on_m"]), bo[0], round(bo[0] * P["ty_le_NC_bao_on"], -2), bo[1],
         "Quy đổi" if "Ngoại suy" in bo[1] else "Có nguồn")
item("Bảo ôn van, phụ kiện nước lạnh (cuộn 25mm + keo + băng dính)", "Superlon", "Malaysia", "Lô", 1,
     round(0.15 * sum(r["sl"] * r["q"] for r in ROWS[-3:]), -3), 0,
     "Quy đổi 15% giá trị bảo ôn ống", "Quy đổi")

sec("II.C", "Van, vòi, thiết bị đường ống")
pump_set(n, 125, 100, "bơm nước lạnh")
pump_set(n, 125, 100, "bơm giải nhiệt")
text(f"Cổ chiller – {n} bộ")
valve("Van bướm tay quay vào/ra bình bay hơi", 100, 2 * n, VB[100], 1)
valve("Khớp nối mềm vào/ra bình bay hơi", 100, 2 * n, KM[100], 0.8, "Đài Loan")
valve("Van bướm tay quay vào/ra bình ngưng", 125, 2 * n, VB[125], 1)
valve("Khớp nối mềm vào/ra bình ngưng", 125, 2 * n, KM[125], 0.8, "Đài Loan")
acc("Nhiệt kế vào/ra bình bay hơi, bình ngưng", "0–100°C", 4 * n, NHIET_KE, 0.4)
acc("Đồng hồ áp vào/ra bình bay hơi, bình ngưng", "0–6 bar", 4 * n, DH, 0.4)
acc("Van bi DN15 cho đồng hồ", "DN15", 4 * n, VBI15, 0.3)
acc("Công tắc dòng chảy ra bình bay hơi, bình ngưng", "", 2 * n, FLOW_SW, 0.4)
text(f"Tháp giải nhiệt – {n} bộ")
valve("Van bướm tay quay vào/ra tháp", 125, 2 * n, VB[125], 1)
valve("Van bi đồng cấp nước bù tháp", 25, n, VBI[25], 0.3 * 100 / 25, "Việt Nam")
valve("Van bi đồng xả đáy tháp", 50, n, VBI[50], 0.3 * 100 / 50, "Việt Nam")

sec("II.D", "Vật tư phụ (bu lông, gioăng)")
acc("Bộ bu lông + gioăng mối nối DN150 (tank)", "8 bu lông", 2, BL[150], 0.2 * 1.5)
acc("Bộ bu lông + gioăng mối nối DN125", "8 bu lông", n * (3 + 3 + 2 + 2 + 2), BL[125], 0.2 * 1.25)
acc("Bộ bu lông + gioăng mối nối DN100", "8 bu lông", n * (3 + 3 + 2), BL[100], 0.2 * 1.0)

sec("II.E", "Giá đỡ")
L_lanh = L_upvc(C["L_header_hut_tank_m"] + C["L_header_ve_tank_m"]) + L_upvc(
    n * (C["L_bom_chiller_m"] + C["L_chiller_header_ve_m"])) + L_upvc(n * C["L_nhanh_hut_bom_m"])
L_gn = L_upvc(n * (2 * C["L_chiller_thap_m"] + 4))
item("Giá đỡ ống D110–D160 (bát V40 + đai omega; ống lạnh có gối PU)", "@1,2 m", "Việt Nam", "bộ",
     math.ceil((L_lanh + L_gn) / 1.2), GD_NHO[0], nc(0.3), GD_NHO[1], "Có nguồn")
item("Giá đỡ cụm van bơm, cổ chiller, cổ tháp (thép hình)", "", "Việt Nam", "cái", 4 * n,
     GD_LON[0], nc(1), GD_LON[1], "Có nguồn")

sec("II.F", "Tủ điện tổng chiller")
ask("Vỏ tủ điện sơn tĩnh điện IP55 + thanh cái, cầu đấu, đèn báo, nút nhấn", "", "bộ", 1, nca(6),
    "Cần hỏi giá; NC = lắp đặt + đấu nối tủ")
acc("MCCB tổng 3P 500A", "LS ABS", 1, MCCB[500], 0.5, xx="Hàn Quốc")
acc("MCCB 3P 150A – chiller", "LS ABN", n, MCCB[150], 0.4, xx="Hàn Quốc")
acc("MCCB 3P 30A – bơm lạnh, bơm GN", "LS ABN", 2 * n, MCCB[30], 0.3, xx="Hàn Quốc")
acc("MCCB 3P 15A – quạt tháp", "LS ABN", n, MCCB[15], 0.3, xx="Hàn Quốc")
acc("Contactor 18A – bơm", "LS MC", 2 * n, CONT[18], 0.3, xx="Hàn Quốc")
acc("Rơ le nhiệt 18A – bơm", "LS MT", 2 * n, RL[18], 0.2, xx="Hàn Quốc")
acc("Contactor 12A – quạt tháp", "LS MC", n, CONT[12], 0.3, xx="Hàn Quốc")
acc("Rơ le nhiệt 18A – quạt tháp", "LS MT", n, RL[18], 0.2, xx="Hàn Quốc")
ask("Bộ điều khiển liên động chiller – bơm – tháp (PLC/relay)", "", "bộ", 1, 0)
ask("Phụ kiện điều khiển (relay trung gian, MCB, máng nhựa trong tủ)", "", "Lô", 1)

sec("II.G", "Dây điện, cáp, máng cáp")
cable("Cáp nguồn tủ → chiller", "3x50+25", C["L_tu_chiller_m"], n, "CV25")
cable("Cáp nguồn tủ → bơm nước lạnh", "3x6+4", C["L_tu_bom_m"], n, "CV4")
cable("Cáp nguồn tủ → bơm giải nhiệt", "3x6+4", C["L_tu_bom_m"], n, "CV4")
cable("Cáp nguồn tủ → quạt tháp", "4x4", C["L_tu_thap_m"], n, "CV4")
q, s = CAP["TH2x1.5"]
item("Cáp tín hiệu liên động, công tắc dòng chảy", "2x1,5 chống nhiễu", "Việt Nam", "m",
     C["L_cap_tin_hieu_m"], q, round(q * P["ty_le_NC_cap"], -2), s, "Có nguồn")
item("Máng cáp mạ kẽm nhúng nóng 300x100 dày 1,5 có nắp", "2DE", "Việt Nam", "m", C["L_mang_cap_m"],
     MANG_300[0], round(MANG_300[0] * P["ty_le_NC_mang"], -2), MANG_300[1], "Có nguồn")
item("Phụ kiện máng cáp (co, tê, giá treo)", "Trọn bộ", "Việt Nam", "Lô", 1,
     round(0.15 * C["L_mang_cap_m"] * MANG_300[0], -3), round(0.05 * C["L_mang_cap_m"] * MANG_300[0], -3),
     "Quy đổi 15% giá trị máng; NC = VT/3 (skill §7.5)", "Quy đổi")
item("Ống ruột gà D25 bảo vệ cáp", "", "Việt Nam", "m", 60, RUOT_GA[0], 18000, RUOT_GA[1], "Có nguồn")

# ================================================================ PHẦN III – CHI PHÍ CHUNG
vt_sum = sum((r["q"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item")
vt_th = sum((r["q"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item" and r["src"] != SRC_GC)
nc_sum = sum((r["r"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item")
part("III", "VẬT TƯ PHỤ VÀ CHI PHÍ KHÁC")
sec("III.A", "Vật tư phụ")
tieu_hao = round(P["ty_le_vat_tu_tieu_hao"] * vt_th, -5)
item("Vật tư tiêu hao (keo, dung môi uPVC, que hàn, sơn, băng tan, khí cắt hàn, đá cắt...)", "Trọn bộ",
     "VN/TQ", "Lô", 1, tieu_hao, 0, f"Quy đổi {P['ty_le_vat_tu_tieu_hao']:.1%} VT phần chiller + giá đỡ (không tính hạng mục theo ghi chép đã gồm phụ kiện)",
     "Lô")
item("Vật tư phụ điện (băng dính, dây thít, đầu cốt, ốc siết cáp)", "Trọn bộ", "VN/TQ", "Lô", 1,
     6000000, 1350000, "Mẫu BG Bút Sơn 13/08/2026 – vật tư phụ phần điện", "Lô")
base = vt_sum + nc_sum + tieu_hao + 6000000 + 1350000
cpk_total = round(P["ty_le_chi_phi_khac"] * base, -6)
sec("III.B", "Chi phí khác")
for name, w in [("Phí vận chuyển, nâng hạ", 20), ("Xe nâng hạ thi công, dàn giáo", 15),
                ("Phí vệ sinh môi trường", 5), ("Công dụng cụ thi công", 10),
                ("Phí văn phòng, kho bãi", 50), ("Phí quản lý", 40)]:
    item(name, "", "Việt Nam", "Lô", 1, 0, round(cpk_total * w / 140, -5),
         f"Quy đổi {P['ty_le_chi_phi_khac']:.0%} tổng, chia theo cơ cấu SGC ({w}/140)", "Lô")
item("Móng bê tông thiết bị", "CĐT PHỤ TRÁCH", "", "Lô", 1, 0, 0, "Skill §5", "Lô", "CĐT PHỤ TRÁCH")
item("Cáp nguồn đến tủ điện tổng", "CĐT PHỤ TRÁCH", "", "Lô", 1, 0, 0, "Phạm vi", "Lô", "CĐT PHỤ TRÁCH")

# ================================================================ GHI FILE
tpl = openpyxl.load_workbook(TPL)
ws = tpl["BG"]
ws.title = "DT"
# chụp style mẫu
def cap(r):
    return {c: copy.copy(ws.cell(r, c)._style) for c in range(1, 13)}
ST = {"part": cap(16), "sec": cap(19), "item": cap(20), "text": cap(22), "tot": cap(61), "note": cap(65)}
H_ROW = {k: ws.row_dimensions[r].height for k, r in [("part", 16), ("sec", 19), ("item", 20)]}
notes = [ws.cell(r, 1).value for r in range(64, 78)]
for mr in list(ws.merged_cells.ranges):
    if mr.min_row >= 16:
        ws.unmerge_cells(str(mr))
ws.delete_rows(16, ws.max_row)

ws["A7"] = f"BẢNG BÁO GIÁ\nDỰ ÁN : {P['du_an']}"
ws.row_dimensions[7].height = 80
_f = copy.copy(ws["A7"].font); _f.sz = 20; ws["A7"].font = _f
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws["B8"] = P["khach_hang"]
ws["B11"] = P["dia_chi_kh"]
ws["H11"] = P["quota_no"]
ws["H12"] = P["ngay"]
YEL = PatternFill("solid", fgColor="FFFF00")
ORG = PatternFill("solid", fgColor="F8CBAD")
RED = PatternFill("solid", fgColor="FFC7CE")
BLU = PatternFill("solid", fgColor="DDEBF7")
thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ws["O8"] = "HỆ SỐ GIÁ BÁN"
ws["O8"].font = Font(bold=True)
ws["P8"] = P["he_so_gia_ban_P8"]
ws["P8"].fill = YEL
ws["P8"].font = Font(bold=True, color="0000FF")
ws["P8"].comment = Comment("Hệ số giá bán – CHỜ USER CHỐT (đang để 1,00 = bán bằng vốn)", "Claude")
ws["L14"] = "GHI CHÚ"
for col, t in zip("PQRSTUVW", ["SL", "ĐG VT vốn", "ĐG NC vốn", "TT VT vốn", "TT NC vốn", "TỔNG vốn",
                               "NGUỒN ĐƠN GIÁ", "TRẠNG THÁI"]):
    c = ws[f"{col}14"]
    c.value = t
    c._style = copy.copy(ws["G14"]._style)
ws.merge_cells("L14:L15")
for col, w in {"B": 63, "C": 22, "D": 12, "G": 15, "H": 14, "I": 17, "J": 16, "K": 18, "L": 34,
               "P": 8, "Q": 14, "R": 14, "S": 16, "T": 16, "U": 17, "V": 60, "W": 14}.items():
    ws.column_dimensions[col].width = w

r0 = 16
part_rows, sec_rows = [], []
cur_part = cur_sec = None
excel_rows = []
for i, it in enumerate(ROWS):
    rr = r0 + i
    it["row"] = rr
    k = it["kind"]
    st = ST[k]
    for c in range(1, 13):
        ws.cell(rr, c)._style = copy.copy(st[c])
    for c in range(16, 24):
        ws.cell(rr, c)._style = copy.copy(st[min(c - 8, 11)])
        ws.cell(rr, c).number_format = "#,##0"
    if k in ("part", "sec", "item"):
        ws.row_dimensions[rr].height = None
    if k == "part":
        cur_part = it; it["children"] = []; part_rows.append(it)
        ws.cell(rr, 1, it["tt"]); ws.cell(rr, 2, it["name"])
    elif k == "sec":
        cur_sec = it; it["children"] = []; sec_rows.append(it); cur_part["children"].append(it)
        ws.cell(rr, 1, it["tt"]); ws.cell(rr, 2, it["name"])
    elif k == "text":
        ws.cell(rr, 2, it["name"])
        cur_sec["children"].append(it)
    else:
        cur_sec["children"].append(it)
        ws.cell(rr, 2, it["name"]); ws.cell(rr, 3, it["spec"]); ws.cell(rr, 4, it["xx"])
        ws.cell(rr, 5, it["dv"]); ws.cell(rr, 6, it["sl"])
        ws.cell(rr, 7, f"=ROUND(Q{rr}*$P$8,-3)"); ws.cell(rr, 8, f"=ROUND(R{rr}*$P$8,-3)")
        ws.cell(rr, 9, f"=F{rr}*G{rr}"); ws.cell(rr, 10, f"=F{rr}*H{rr}"); ws.cell(rr, 11, f"=I{rr}+J{rr}")
        ws.cell(rr, 12, it["note"])
        ws.cell(rr, 16, f"=F{rr}")
        if it["q"] is not None:
            ws.cell(rr, 17, it["q"])
        if it["r"] is not None:
            ws.cell(rr, 18, it["r"])
        ws.cell(rr, 19, f"=P{rr}*Q{rr}"); ws.cell(rr, 20, f"=P{rr}*R{rr}"); ws.cell(rr, 21, f"=S{rr}+T{rr}")
        ws.cell(rr, 22, it["src"]); ws.cell(rr, 23, it["st"])
        ws.cell(rr, 22).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(rr, 12).alignment = Alignment(wrap_text=True, vertical="center")
        if it["st"] == "Quy đổi":
            for c in (17, 18):
                ws.cell(rr, c).fill = ORG
            ws.cell(rr, 17).comment = Comment(it["src"], "Claude")
        elif it["st"] == "Cần hỏi giá":
            ws.cell(rr, 17).fill = RED
            ws.cell(rr, 23).fill = RED
        elif it["st"] == "Đề xuất NC":
            ws.cell(rr, 18).fill = YEL
            ws.cell(rr, 18).comment = Comment(it["src"] + " – đề xuất, chờ user duyệt", "Claude")
    for c in (6, 7, 8, 9, 10, 11, 17, 18, 19, 20, 21):
        ws.cell(rr, c).number_format = "#,##0" if c != 6 else "#,##0.0#"

last = r0 + len(ROWS) - 1
for grp in part_rows + sec_rows:
    rr = grp["row"]
    # phạm vi: từ dòng sau đến dòng cuối của nhóm
    nxt = [g["row"] for g in (part_rows if grp in part_rows else part_rows + sec_rows) if g["row"] > rr]
    if grp in sec_rows:
        nxt = [g["row"] for g in part_rows + sec_rows if g["row"] > rr]
    end = (min(nxt) - 1) if nxt else last
    for c in "IJKSTU":
        ws[f"{c}{rr}"] = f"=SUBTOTAL(9,{c}{rr + 1}:{c}{end})"
        ws[f"{c}{rr}"].number_format = "#,##0"
# part rows: phạm vi đến trước part kế tiếp
for idx, grp in enumerate(part_rows):
    rr = grp["row"]
    end = part_rows[idx + 1]["row"] - 1 if idx + 1 < len(part_rows) else last
    for c in "IJKSTU":
        ws[f"{c}{rr}"] = f"=SUBTOTAL(9,{c}{rr + 1}:{c}{end})"

t = last + 1
for rr, lab in [(t, "TỔNG GIÁ TRỊ TRƯỚC THUẾ"), (t + 1, f"THUẾ VAT {P['VAT']:.0%}"), (t + 2, "TỔNG GIÁ TRỊ SAU THUẾ")]:
    for c in range(1, 13):
        ws.cell(rr, c)._style = copy.copy(ST["tot"][c])
    ws.cell(rr, 1, lab).font = Font(name="Times New Roman", size=14, bold=True)
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8 if rr == t else 10)
for c in "IJKSTU":
    ws[f"{c}{t}"] = f"=SUBTOTAL(9,{c}{r0}:{c}{last})"
    ws[f"{c}{t}"].number_format = "#,##0"
ws[f"K{t + 1}"] = f"=K{t}*{P['VAT']}"
ws[f"K{t + 2}"] = f"=K{t}+K{t + 1}"
ws[f"U{t + 1}"] = f"=U{t}*{P['VAT']}"
ws[f"U{t + 2}"] = f"=U{t}+U{t + 1}"
ws[f"P{t}"] = "Lãi gộp"; ws[f"Q{t}"] = f"=K{t}-U{t}"
ws[f"P{t + 1}"] = "Tỷ suất"; ws[f"Q{t + 1}"] = f"=IF(K{t}=0,0,Q{t}/K{t})"
ws[f"Q{t + 1}"].number_format = "0.0%"
for c in ("K", "U", "Q"):
    for rr in (t, t + 1, t + 2):
        if ws[f"{c}{rr}"].number_format != "0.0%":
            ws[f"{c}{rr}"].number_format = "#,##0"
        ws[f"{c}{rr}"].font = Font(name="Times New Roman", size=14, bold=True)

notes_new = list(notes)
notes_new[1] = "1. Báo giá trên đã bao gồm VAT, chi phí giao hàng, chi phí lắp đặt hệ thống và phí bảo hành "
notes_new[2] = ("2. Báo giá trên chưa bao gồm: thiết bị hệ chiller (chiller, bơm, tháp giải nhiệt, tank 20m³), "
                "móng máy, dây nguồn đến tủ điện tổng và chi phí phát sinh")
nr = t + 4
for k_, v in enumerate(notes_new):
    rr = nr + k_
    for c in range(1, 13):
        ws.cell(rr, c)._style = copy.copy(ST["note"][c])
    ws.cell(rr, 1, v)
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=11 if k_ == len(notes_new) - 1 else 7)
ws.freeze_panes = "C16"
ws.print_area = f"A1:L{nr + len(notes_new) - 1}"

# ---- sheet NOI_SUY
ns = tpl.create_sheet("NOI_SUY")
ns.append(["Hạng mục", "Cách tính", "Đơn giá (đ)"])
for x in NOI_SUY:
    ns.append(list(x))
ns.append([])
ns.append(["Ghi chú", "Ô cam ở sheet DT = đơn giá quy đổi/nội suy; ô đỏ = cần hỏi giá; ô vàng cột R = NC đề xuất chờ duyệt", ""])
for col, w in zip("ABC", (40, 90, 16)):
    ns.column_dimensions[col].width = w
for row in ns.iter_rows(min_row=1, max_row=1):
    for c in row:
        c.font = Font(bold=True); c.fill = BLU

# ---- sheet GIA_DINH
gd = tpl.create_sheet("GIA_DINH")
gd.append(["Tham số", "Giá trị"])
def flat(d, pre=""):
    for k_, v in d.items():
        if isinstance(v, dict):
            flat(v, pre + k_ + ".")
        else:
            gd.append([pre + k_, v if not isinstance(v, (list,)) else str(v)])
flat(P)
gd.column_dimensions["A"].width = 40
gd.column_dimensions["B"].width = 110
for c in gd[1]:
    c.font = Font(bold=True); c.fill = BLU

OUT.parent.mkdir(parents=True, exist_ok=True)
tpl.save(OUT)

# ---- in tóm tắt
items = [r for r in ROWS if r["kind"] == "item"]
VT = sum((r["q"] or 0) * r["sl"] for r in items)
NC = sum((r["r"] or 0) * r["sl"] for r in items)
miss = [r for r in items if r["st"] == "Cần hỏi giá"]
print(f"OUT: {OUT.relative_to(ROOT)}")
print(f"Dòng hạng mục: {len(items)} | Cần hỏi giá: {len(miss)} | Quy đổi: {sum(r['st']=='Quy đổi' for r in items)}"
      f" | NC đề xuất: {sum(r['st']=='Đề xuất NC' for r in items)}")
for pr in part_rows:
    s_vt = s_nc = 0
    on = False
    for r in ROWS:
        if r is pr:
            on = True; continue
        if on and r["kind"] == "part":
            break
        if on and r["kind"] == "item":
            s_vt += (r["q"] or 0) * r["sl"]; s_nc += (r["r"] or 0) * r["sl"]
    print(f"  {pr['tt']}: VT {s_vt:,.0f} | NC {s_nc:,.0f} | Tổng {s_vt + s_nc:,.0f}")
print(f"VỐN: VT {VT:,.0f} | NC+CP {NC:,.0f} | Tổng trước thuế {VT + NC:,.0f} | sau VAT {(VT + NC) * (1 + P['VAT']):,.0f}")
print(f"Lô F=1 kiểm: {all(r['sl'] == 1 for r in items if r['dv'] == 'Lô')}")
