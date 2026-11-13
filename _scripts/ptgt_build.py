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

UPVC[90] = bg("CT-01065")
BICH[90] = bg("CT-00107")
# Cút 90° D90: không có dòng A → nội suy tuyến tính giữa D60 (A) và D110 (A)
_c90 = round(CUT[60][0] + (CUT[110][0] - CUT[60][0]) * (90 - 60) / (110 - 60), -2)
CUT[90] = (_c90, "Nội suy tuyến tính cút 90° uPVC Tiền Phong (A) D60–D110")
# Van bi nhựa PVC D90: dữ liệu NCC (A) chỉ đến D63 → ngoại suy hàm mũ từ D48, D63
_bb = math.log(85000 / 60000) / math.log(63 / 48)
_vb90 = round(85000 * (90 / 63) ** _bb, -2)
VBI_PVC90 = (_vb90, "Ngoại suy hàm mũ van bi PVC dán keo (CT-00714/715, A) – ngoài dải, độ tin cậy thấp")
_vb110 = round(85000 * (110 / 63) ** _bb, -2)
VBI_PVC110 = (_vb110, "Ngoại suy hàm mũ van bi PVC dán keo (CT-00714/715, A) – ngoài dải, độ tin cậy thấp, cần hỏi giá")
NOI_SUY.append(("Van bi nhựa PVC tay gạt D110", f"85.000×(110/63)^{_bb:.3f} (ngoại suy từ D48, D63) – cần hỏi giá xác nhận", _vb110))
CON_110_90 = bg("CT-00284")
# Van Wonil (Tuấn Hưng Phát – nguồn ưu tiên skill §3)
VB = {50: bg("CT-00445"), 100: bg("CT-00448"), 125: bg("CT-00449"), 150: bg("CT-00450")}
YL = {100: bg("CT-00887"), 125: bg("CT-00888")}
YL80 = bg("CT-00886")
VMC = {100: bg("CT-00618"), 125: bg("CT-00619")}
# khớp mềm Shinyi – sát giá chốt SGC DN150 (1.653.000 vs 1.715.000)
KM = {50: bg("CT-00013"), 80: bg("CT-00015"), 100: bg("CT-00016"), 125: bg("CT-00017")}
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
BL[80] = (95000, "Quy đổi: lấy giá DN100 SGC (không có DN80) – độ tin cậy thấp")
NOI_SUY.append(("Bộ bu lông + gioăng DN80", "Không có DN80 → lấy DN100 (cỡ gần nhất có giá)", 95000))
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
_bo15b = round(102459 + (163934 - 102459) * (15 - 13) / (19 - 13), -2)
BO_114_15 = (_bo15b, "Nội suy tuyến tính Superlon Ø114 dày 13mm (HV-00054) – 19mm (HV-00072)")
NOI_SUY.append(("Bảo ôn Superlon Ø114 dày 15mm", "102.459 + (163.934−102.459)×(15−13)/(19−13)", _bo15b))
_bo15 = round(50546 + (81967 - 50546) * (15 - 13) / (19 - 13), -2)
BO_89_15 = (_bo15, "Nội suy tuyến tính Superlon Ø89 dày 13mm (HV-00053) – 19mm (HV-00071)")
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
THANG_300 = bg("DI-00799")
NAP_300 = bg("DI-00707")
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
    item(f"{kind}", "", "Tiền Phong", "cái", sl, q, nc(hs * D / 100), s,
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
GC = P["gia_user_buong_lanh"]
HS = GC["he_so_dau_vao"]
SRC_GC = GC["nguon"]


def g(v):
    """Giá đầu vào = đơn giá user cấp × hệ số (1,10), làm tròn nghìn."""
    return round(v * HS, -3)


def gc(name, spec, xx, dv, sl, vt, ncv, note=""):
    item(name, spec, xx, dv, sl, g(vt) if vt else 0, g(ncv) if ncv else 0, SRC_GC,
         "Lô" if dv == "Lô" else "Có nguồn",
         (note + "; " if note else "") + f"Đơn giá user cấp × {HS:.2f}")


sec("I.1", "Thiết bị")
a, b, c = B["dai"], B["rong"], B["cao"]
S_panel = round(2 * (a * b + a * c + b * c), 2)   # 6 mặt: tường + trần + sàn
gc("Panel PU 100mm 2 mặt inox – 6 mặt (tường, trần, sàn)",
   f"2×({a}×{b}+{a}×{c}+{b}×{c}) = {S_panel} m²", "", "m2", S_panel, GC["panel_2mat_m2"], GC["nc_lap_vo_m2"],
   "KT giả định phủ bì, không trừ ô cửa")
item("Cửa bản lề panel 100mm 2 mặt inox, có điện trở sưởi khung", "700 x 1700 mm", "", "bộ", B["so_cua"],
     P["gia_user_chiller"]["cua"], P["gia_user_chiller"]["nc_cua"], "Giá user cấp 13/11/2026", "Có nguồn")
ROWS[-1]["user"] = True
gc("Vật tư phụ lắp vỏ (silicone, keo bọt, ốc vít, rivet, đèn, van thông áp)", "Trọn bộ", "", "Lô", 1,
   GC["vt_phu_lap_vo_lo"], 0)
gc("Cụm máy nén bình ngưng 20HP", "Máy nén 20HP + bình ngưng + nhớt POE + phụ kiện lên máy", "Panasonic",
   "cụm", nC, GC["may_nen_20HP"] + GC["binh_ngung_20HP"] + GC["nhot_POE_bo"] + GC["phu_kien_len_may"],
   GC["nc_len_cum"], "NC = nhân công lắp ráp cụm")
ROWS[-1]["nc_share"] = "I.1"
gc("Dàn lạnh DJ 14.3/85-201 (≈25HP)", "", "", "cái", nE, GC["dan_lanh"], 0)
ROWS[-1]["nc_share"] = "I.1"
gc("Tháp giải nhiệt 20RT", "", "TTP", "bộ", B["so_thap"], GC["thap_20RT"], 0,
   "1 tháp 20RT cho mỗi cụm CDU 20HP (user chốt 13/11)")
ROWS[-1]["nc_share"] = "I.1"
gc("Bơm giải nhiệt", "", "Ewara", "cái", B["so_bom"], GC["bom_1_5kW"], 0)
ROWS[-1]["nc_share"] = "I.1"

sec("I.2", "Đường ống")
Lg = nC * (B["L_CDU_buong_m"] + B["L_them_moi_tuyen_m"])
L_cu = round(Lg * (1 + HH))
gc("Ống đồng Ø42x1,0 + Ø16x0,8 + bảo ôn 42x25 + phụ kiện đồng (theo cặp ống)", "",
   "Hailiang", "m", L_cu, GC["ong_dong_cap_m"], 0,
   f"{nC} tuyến × ({B['L_CDU_buong_m']}+{B['L_them_moi_tuyen_m']}) m × 1,05")
ROWS[-1]["nc_share"] = "I.2"
gc("Van tiết lưu TES5", "", "Danfoss", "bộ", nC, GC["tiet_luu_TES5"], 0)
ROWS[-1]["nc_share"] = "I.2"
gc("Gas lạnh R404A/R507", "", "China", "hệ", nC, GC["gas_moi_he"], 0)
ROWS[-1]["nc_share"] = "I.2"
gc("Ống PVC + phụ kiện (nước giải nhiệt, thoát nước xả đá)", "", "Việt Nam", "hệ", nC, GC["ong_pvc_moi_he"], 0)
ROWS[-1]["nc_share"] = "I.2"
item("Giá đỡ chung tuyến ống gas (thép hình, ty treo, kẹp)", "@1,5 m", "Việt Nam", "cái",
     math.ceil(B["L_CDU_buong_m"] / 1.5) + 2, GD_LON[0], nc(1), GD_LON[1], "Có nguồn")

sec("I.3", "Đường điện")
gc("Tủ điện điều khiển cụm máy (1 tủ/cụm)", "", "Việt Nam", "tủ", nC, GC["tu_dk_moi_he"], 0)
ROWS[-1]["nc_share"] = "I.3"
gc("Dây điện, cáp động lực và điều khiển", "", "Việt Nam", "hệ", nC, GC["day_dien_moi_he"], 0)
ROWS[-1]["nc_share"] = "I.3"

# ---- NC Phần I: tổng NC user cấp (lắp ráp cụm 4tr + lắp đặt 10tr)/hệ × 1,10 × nC giữ nguyên;
#      đặt NC cố định theo user, phần còn lại phân bổ theo tỷ lệ VT cho ống đồng, gas, ống PVC, dây điện
_nc_pool = (g(GC["nc_len_cum"]) + g(GC["nc_lap_dat_moi_he"])) * nC
_fix = P["nc_co_dinh_buong_lanh"]
_map = {"Cụm máy nén bình ngưng 20HP": "cum", "Dàn lạnh DJ 14.3/85-201 (≈25HP)": "dan_lanh",
        "Tủ điện điều khiển cụm máy (1 tủ/cụm)": "tu", "Tháp giải nhiệt 20RT": "thap", "Bơm giải nhiệt": "bom",
        "Van tiết lưu TES5": "tiet_luu"}
_rows = [r for r in ROWS if r.get("nc_share")]
_fixed_sum = 0
for r in _rows:
    if r["name"] in _map:
        r["r"] = _fix[_map[r["name"]]]
        r["note"] = "NC user cấp"
        _fixed_sum += r["r"] * r["sl"]
_rest = [r for r in _rows if r["name"] not in _map]
_vt = sum(r["q"] * r["sl"] for r in _rest)
_remain = _nc_pool - _fixed_sum
for r in _rest:
    r["r"] = round(_remain * r["q"] / _vt)
    r["note"] = f"NC = phần còn lại {_remain:,.0f} đ phân bổ theo tỷ lệ VT ({r['q'] * r['sl'] / _vt:.2%})"
# ================================================================ PHẦN II – CHILLER
C = P["chiller"]
n = C["so_chiller"]
U = P["gia_user_chiller"]
SRC_U = U["nguon"]
D = 110
part("II", "HỆ THỐNG CHILLER 03 x 80HP")
sec("II.1", "Thiết bị (tank 20m³ do CĐT cấp, có sẵn 6 bích chờ DN100)")
item("Chiller 80HP giải nhiệt nước (đầu nén Hanbell - Đài Loan)", "", "Hstar", "bộ", n, U["chiller"], U["nc_chiller"], SRC_U + " – giá nhập", "Có nguồn",
     "VT: giá nhập, NC: user cấp")
item(f"Bơm tuần hoàn nước lạnh {C['luu_luong_nuoc_lanh_m3h']} m³/h", "", "Teco", "cái", n, U["bom"], U["nc_bom"], SRC_U,
     "Có nguồn", "VT, NC: user cấp")
item(f"Bơm giải nhiệt {C['luu_luong_giai_nhiet_m3h']} m³/h", "", "Teco", "cái", n, U["bom"], U["nc_bom"], SRC_U,
     "Có nguồn", "VT, NC: user cấp")
item("Tháp giải nhiệt", "", "TTP", "bộ", n, U["thap"], U["nc_thap"], SRC_U, "Có nguồn",
     "VT, NC: user cấp")
for r_ in ROWS[-4:]:
    r_["user"] = True
    r_["no_cpk"] = 1

sec("II.2", "Đường ống")
L_lanh_1 = C["L_tank_bom_m"] + C["L_bom_chiller_m"] + C["L_chiller_tank_m"]
L_gn_1 = 2 * C["L_chiller_thap_m"] + 4
A_ = math.pi / 4 * C["ID_D110_m"] ** 2
v_l = C["luu_luong_nuoc_lanh_m3h"] / 3600 / A_
v_g = C["luu_luong_giai_nhiet_m3h"] / 3600 / A_
item(f"Ống uPVC D{D}", "C2 (PN8)", "Tiền Phong", "m", L_upvc(n * (L_lanh_1 + L_gn_1)),
     UPVC[D][0], nc(D / 100), UPVC[D][1], "Có nguồn", f"{n}×({L_lanh_1}+{L_gn_1}) m × 1,05, tròn cây 4 m")
fit(f"Cút 90° uPVC D{D}", D, n * (6 + 8), CUT[D], 0.6)
fit(f"Mặt bích uPVC D{D}", D, 2 * n + 2 * (2 * n) + 4 * n + 2 * n, BICH[D], 0.6)
ROWS[-1]["note"] = "Bích chờ tank DN100 6 + bơm 2/cái×6 + cổ chiller 4×3 + cổ tháp 2×3"
item(f"Bảo ôn cao su lưu hóa dày 15mm cho ống D{D}", "", "China", "m",
     ceil_to(n * L_lanh_1 * (1 + HH), 4), BO_114_15[0],
     round(BO_114_15[0] * P["ty_le_NC_bao_on"], -2), BO_114_15[1], "Quy đổi")
item("Bảo ôn van, phụ kiện nước lạnh (cuộn + keo + băng dính)", "", "China", "Lô", 1,
     round(0.15 * ROWS[-1]["sl"] * ROWS[-1]["q"], -3), U["nc_bao_on_pk"], "Quy đổi 15% giá trị bảo ôn ống; NC user cấp", "Quy đổi")

VBI110 = (P["gia_user_chiller"]["van_pvc_110"], "Giá user cấp 13/11/2026")
acc(f"Van bi nhựa PVC tay gạt D{D}", "", 2 * n + n + n, VBI110, 1.1, xx="Taijan")
item("Y lọc mặt bích DN100", "DN100", "China", "cái", 2 * n, U["y_loc"], nc(1), "Giá user cấp 13/11/2026", "Có nguồn")

acc("Bộ bu lông + gioăng mối nối DN100", "8 bu lông", 2 * n + 3 * 2 * n + 4 * n + 2 * n, BL[100], 0.2 * 1.0, xx="VN/TQ")
ROWS[-1]["note"] = "Tank 6 + bơm 3/cái×6 (bích–Y, Y–bơm, bơm–bích) + cổ chiller 12 + tháp 6"

item("Giá đỡ đường ống chiller (thép hình, bát V, đai omega, gối PU ống lạnh)", "Trọn bộ", "Việt Nam", "Lô", 1,
     U["gia_do_vt"], U["gia_do_nc"], "Giá user cấp 13/11/2026", "Lô")

item("Vật tư tiêu hao (keo, dung môi uPVC, que hàn, sơn, băng tan, đá cắt...)", "Trọn bộ", "VN/TQ", "Lô", 1,
     0, 0, "Công thức: ô P10 (tỷ lệ) × tổng VT vốn các dòng phía trên có cột X = 1 "
     "(không tính thiết bị, hạng mục giá user cấp)", "Lô")
ROWS[-1]["tag"] = "th"

sec("II.3", "Đường điện")
item("Tủ điện, dây điện cấp nguồn", "CĐT cấp", "", "Lô", 1, 0, 0, "User 13/11/2026", "Lô", "CĐT cấp")

# ================================================================ PHẦN III – CHI PHÍ CHUNG
vt_sum = sum((r["q"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item")
for r_ in ROWS:
    if r_["kind"] == "item":
        r_["th"] = int(r_["src"] != SRC_GC and not r_.get("user") and r_["st"] in ("Có nguồn", "Quy đổi"))
vt_th = sum((r["q"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item" and r["th"])
nc_sum = sum((r["r"] or 0) * r["sl"] for r in ROWS if r["kind"] == "item")
part("III", "CHI PHÍ KHÁC")
tieu_hao = round(P["ty_le_vat_tu_tieu_hao"] * vt_th, -5)
next(r for r in ROWS if r.get("tag") == "th")["q"] = tieu_hao + P["tieu_hao_cong_them"]
base = vt_sum + nc_sum + tieu_hao - sum(r["q"] * r["sl"] for r in ROWS if r.get("no_cpk"))
cpk_total = round(P["ty_le_chi_phi_khac"] * base, -6)
for name, w in [("Phí vận chuyển, nâng hạ", 20), ("Xe nâng hạ thi công, dàn giáo", 15),
                ("Phí vệ sinh môi trường", 5), ("Công dụng cụ thi công", 10),
                ("Phí quản lý", 40)]:
    item(name, "", "Việt Nam", "Lô", 1, 0, round(cpk_total * w / 140, -5),
         f"Công thức: ô P9 × [tổng vốn (I + II) − VT thiết bị chiller (cột Y = 1)] × {w}/140 (cơ cấu SGC)", "Lô")
    ROWS[-1]["tag"] = ("cpk", w)
    if w == 20:
        ROWS[-1]["extra"] = g(GC["van_chuyen_panel_lo"])
        ROWS[-1]["r"] += ROWS[-1]["extra"]
        ROWS[-1]["src"] += f" + vận chuyển panel {ROWS[-1]['extra']:,.0f} đ"
item("Móng bê tông thiết bị", "CĐT PHỤ TRÁCH", "", "Lô", 1, 0, 0, "Skill §5", "Lô", "CĐT PHỤ TRÁCH")
item("Cáp nguồn đến tủ điện tổng", "CĐT PHỤ TRÁCH", "", "Lô", 1, 0, 0, "Phạm vi", "Lô", "CĐT PHỤ TRÁCH")

# ================================================================ GHI FILE
# Bố cục cột (v12 – đã xoá cột Ghi chú/Nguồn/Trạng thái):
#   A–K bán | L, M trống | N nhãn tham số | O SL vốn | P ĐG VT vốn | Q ĐG NC vốn | R TT VT | S TT NC | T TỔNG vốn
#   U cờ tiêu hao (1/0) | V cờ loại khỏi nền CPK (1)
from openpyxl.utils import get_column_letter as CL
cSL, cQ, cR, cS, cT, cU, cX, cY = 15, 16, 17, 18, 19, 20, 21, 22
L_ = {k: CL(v) for k, v in dict(SL=cSL, Q=cQ, R=cR, S=cS, T=cT, U=cU, X=cX, Y=cY).items()}
HS_CELL, CPK_CELL, TH_CELL = "O8", "O9", "O10"
tpl = openpyxl.load_workbook(TPL)
ws = tpl["BG"]
ws.title = "DT"
def cap(r):
    return {c: copy.copy(ws.cell(r, c)._style) for c in range(1, 12)}
ST = {"part": cap(16), "sec": cap(19), "item": cap(20), "text": cap(22), "tot": cap(61), "note": cap(65)}
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
ws["H8"] = P["sale"]["ten"]; ws["H9"] = P["sale"]["mobile"]; ws["H10"] = P["sale"]["email"]
ws["B11"] = P["dia_chi_kh"]
ws["H11"] = P["quota_no"]
ws["H12"] = P["ngay"]
YEL = PatternFill("solid", fgColor="FFFF00")
ORG = PatternFill("solid", fgColor="F8CBAD")
BLU = PatternFill("solid", fgColor="DDEBF7")
for cell, lab, val, fmt in [(HS_CELL, "HỆ SỐ GIÁ BÁN", P["he_so_gia_ban_P8"], "0.00"),
                            (CPK_CELL, "TỶ LỆ CHI PHÍ KHÁC", P["ty_le_chi_phi_khac"], "0.0%"),
                            (TH_CELL, "TỶ LỆ VT TIÊU HAO", P["ty_le_vat_tu_tieu_hao"], "0.0%")]:
    lc = ws.cell(ws[cell].row, cSL - 1, lab); lc.font = Font(bold=True)
    ws[cell] = val; ws[cell].number_format = fmt
    ws[cell].fill = YEL; ws[cell].font = Font(bold=True, color="0000FF")
ws[HS_CELL].comment = Comment("Hệ số giá bán: G = ROUND(ĐG VT vốn × hệ số, -3); H tương tự", "Claude")
ws[CPK_CELL].comment = Comment("Chi phí khác = ô này × [tổng vốn Phần I + II − VT các dòng cột V = 1 (chiller, bơm, tháp chiller)] × tỷ trọng/140 (SGC)", "Claude")
ws[TH_CELL].comment = Comment("Vật tư tiêu hao = ô này × Σ(SL × ĐG VT vốn × cột U) các dòng phía trên", "Claude")
for col, t in [(cSL, "SL"), (cQ, "ĐG VT vốn"), (cR, "ĐG NC vốn"), (cS, "TT VT vốn"), (cT, "TT NC vốn"),
               (cU, "TỔNG vốn"), (cX, "TÍNH TIÊU HAO (1/0)"), (cY, "LOẠI VT KHỎI NỀN CPK (1)")]:
    c = ws.cell(14, col, t)
    c._style = copy.copy(ws["G14"]._style)
COLW = {"A": 8, "B": 60, "C": 30, "D": 14, "E": 9, "F": 9, "G": 15, "H": 14, "I": 18, "J": 16, "K": 19}
for col, w in {**COLW, "N": 22,
               L_["SL"]: 8, L_["Q"]: 14, L_["R"]: 14, L_["S"]: 16, L_["T"]: 16, L_["U"]: 17,
               L_["X"]: 10, L_["Y"]: 12}.items():
    ws.column_dimensions[col].width = w

SLc, Qc, Rc, Sc, Tc, Uc, Xc, Yc = (L_[k] for k in ("SL", "Q", "R", "S", "T", "U", "X", "Y"))
r0 = 16
part_rows, sec_rows = [], []
for i, it in enumerate(ROWS):
    rr = r0 + i
    it["row"] = rr
    k = it["kind"]
    st = ST[k]
    for c in range(1, 12):
        ws.cell(rr, c)._style = copy.copy(st[c])
    for c in range(cSL, cY + 1):
        ws.cell(rr, c)._style = copy.copy(st[11])
        ws.cell(rr, c).number_format = "#,##0"
    if k in ("part", "sec"):
        (part_rows if k == "part" else sec_rows).append(it)
        ws.cell(rr, 1, it["tt"]); ws.cell(rr, 2, it["name"])
    elif k == "text":
        ws.cell(rr, 2, it["name"])
    else:
        ws.cell(rr, 2, it["name"]); ws.cell(rr, 3, it["spec"]); ws.cell(rr, 4, it["xx"])
        ws.cell(rr, 5, it["dv"]); ws.cell(rr, 6, it["sl"])
        ws.cell(rr, 7, f"=ROUND({Qc}{rr}*${HS_CELL[0]}${HS_CELL[1:]},-3)")
        ws.cell(rr, 8, f"=ROUND({Rc}{rr}*${HS_CELL[0]}${HS_CELL[1:]},-3)")
        ws.cell(rr, 9, f"=F{rr}*G{rr}"); ws.cell(rr, 10, f"=F{rr}*H{rr}"); ws.cell(rr, 11, f"=I{rr}+J{rr}")
        ws.cell(rr, cSL, f"=F{rr}")
        if it["q"] is not None:
            ws.cell(rr, cQ, it["q"])
        if it["r"] is not None:
            ws.cell(rr, cR, it["r"])
        ws.cell(rr, cS, f"={SLc}{rr}*{Qc}{rr}"); ws.cell(rr, cT, f"={SLc}{rr}*{Rc}{rr}")
        ws.cell(rr, cU, f"={Sc}{rr}+{Tc}{rr}")
        if "th" in it:
            ws.cell(rr, cX, it["th"])
        ws.cell(rr, cY, it.get("no_cpk", 0))
        if it["st"] == "Quy đổi":
            for c in (cQ, cR):
                ws.cell(rr, c).fill = ORG
            ws.cell(rr, cQ).comment = Comment(it["src"], "Claude")
    for c in (6, 7, 8, 9, 10, 11):
        ws.cell(rr, c).number_format = "#,##0" if c != 6 else "#,##0.0#"

# ---- chiều cao dòng theo độ dài chữ (B, C, D xuống dòng)
for it in ROWS:
    rr = it["row"]
    n_line = 1
    for col in ("B", "C", "D"):
        c = ws[f"{col}{rr}"]
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="left" if col == "B" else "center")
        txt = str(c.value or "")
        if txt:
            n_line = max(n_line, math.ceil(len(txt) / (COLW[col] * 0.95)))
    ws.row_dimensions[rr].height = max(24, 20 * n_line + 6)

last = r0 + len(ROWS) - 1
SUBC = ["I", "J", "K", Sc, Tc, Uc]
for grp in sec_rows:
    rr = grp["row"]
    nxt = [g["row"] for g in part_rows + sec_rows if g["row"] > rr]
    end = (min(nxt) - 1) if nxt else last
    for c in SUBC:
        ws[f"{c}{rr}"] = f"=SUBTOTAL(9,{c}{rr + 1}:{c}{end})"
        ws[f"{c}{rr}"].number_format = "#,##0"
for idx, grp in enumerate(part_rows):
    rr = grp["row"]
    end = part_rows[idx + 1]["row"] - 1 if idx + 1 < len(part_rows) else last
    for c in SUBC:
        ws[f"{c}{rr}"] = f"=SUBTOTAL(9,{c}{rr + 1}:{c}{end})"
        ws[f"{c}{rr}"].number_format = "#,##0"

_pI = next(g for g in part_rows if g["tt"] == "I")["row"]
_pII = next(g for g in part_rows if g["tt"] == "II")["row"]
_pIII = next(g for g in part_rows if g["tt"] == "III")["row"]
for it in ROWS:
    tg = it.get("tag")
    if tg == "th":
        _e = it["row"] - 1
        ws.cell(it["row"], cQ, f"=ROUND(${TH_CELL[0]}${TH_CELL[1:]}*SUMPRODUCT({SLc}{r0}:{SLc}{_e},{Qc}{r0}:{Qc}{_e},{Xc}{r0}:{Xc}{_e}),-5)+{P['tieu_hao_cong_them']}")
    elif isinstance(tg, tuple):
        _e = _pIII - 1
        ws.cell(it["row"], cR, f"=ROUND(${CPK_CELL[0]}${CPK_CELL[1:]}*({Uc}{_pI}+{Uc}{_pII}-SUMPRODUCT({Sc}{r0}:{Sc}{_e},{Yc}{r0}:{Yc}{_e}))*{tg[1]}/140,-5)" + (f"+{it['extra']}" if it.get("extra") else ""))
t = last + 1
for rr, lab in [(t, "TỔNG GIÁ TRỊ TRƯỚC THUẾ"), (t + 1, f"THUẾ VAT {P['VAT']:.0%}"), (t + 2, "TỔNG GIÁ TRỊ SAU THUẾ")]:
    for c in range(1, 12):
        ws.cell(rr, c)._style = copy.copy(ST["tot"][c])
    ws.cell(rr, 1, lab).font = Font(name="Times New Roman", size=14, bold=True)
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8 if rr == t else 10)
for c in SUBC:
    ws[f"{c}{t}"] = f"=SUBTOTAL(9,{c}{r0}:{c}{last})"
ws[f"K{t + 1}"] = f"=K{t}*{P['VAT']}"
ws[f"K{t + 2}"] = f"=K{t}+K{t + 1}"
ws[f"{Uc}{t + 1}"] = f"={Uc}{t}*{P['VAT']}"
ws[f"{Uc}{t + 2}"] = f"={Uc}{t}+{Uc}{t + 1}"
ws[f"{SLc}{t}"] = "Lãi gộp"; ws[f"{Qc}{t}"] = f"=K{t}-{Uc}{t}"
ws[f"{SLc}{t + 1}"] = "Tỷ suất"; ws[f"{Qc}{t + 1}"] = f"=IF(K{t}=0,0,{Qc}{t}/K{t})"
for c in ("I", "J", "K", Sc, Tc, Uc, Qc):
    for rr in (t, t + 1, t + 2):
        ws[f"{c}{rr}"].number_format = "#,##0"
        ws[f"{c}{rr}"].font = Font(name="Times New Roman", size=14, bold=True)
ws[f"{Qc}{t + 1}"].number_format = "0.0%"

notes_new = list(notes)
notes_new[1] = "1. Báo giá trên đã bao gồm VAT, chi phí giao hàng, chi phí lắp đặt hệ thống và phí bảo hành "
notes_new[2] = ("2. Báo giá trên chưa bao gồm: tank 20m³, "
                "móng máy, dây nguồn đến tủ điện tổng và chi phí phát sinh")
nr = t + 4
for k_, v in enumerate(notes_new):
    rr = nr + k_
    for c in range(1, 12):
        ws.cell(rr, c)._style = copy.copy(ST["note"][c])
    ws.cell(rr, 1, v)
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=11 if k_ == len(notes_new) - 1 else 7)
ws.freeze_panes = None
ws.print_area = f"A1:K{nr + len(notes_new) - 1}"

# ---- sheet NGUON (thay cho cột Ghi chú/Nguồn/Trạng thái đã xoá ở DT)
ng = tpl.create_sheet("NGUON")
ng.append(["Dòng DT", "Hạng mục", "Nguồn đơn giá", "Trạng thái", "Ghi chú"])
for it in ROWS:
    if it["kind"] == "item":
        ng.append([it["row"], it["name"], it["src"], it["st"], it["note"]])
for col, w in zip("ABCDE", (9, 60, 70, 14, 70)):
    ng.column_dimensions[col].width = w
for c in ng[1]:
    c.font = Font(bold=True); c.fill = BLU

# ---- sheet NOI_SUY
ns = tpl.create_sheet("NOI_SUY")
ns.append(["Hạng mục", "Cách tính", "Đơn giá (đ)"])
_used = {r["q"] for r in ROWS if r["kind"] == "item"}
for x in NOI_SUY:
    if x[2] in _used:
        ns.append(list(x))
ns.append([])
ns.append(["Ghi chú", "Ô cam ở sheet DT (cột P, Q) = đơn giá quy đổi/nội suy; nguồn từng dòng xem sheet NGUON", ""])
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
