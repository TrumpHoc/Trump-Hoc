#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WP4 – Điện & điều khiển HVAC – SGS Hải Phòng, khu H&N Phase 2 (Nhà xưởng 3)

Script tái lập toàn bộ kết quả WP4:
  electrical_design.json, takeoff_electrical.json, WP4_Electrical.xlsx (công thức, đã tính lại bằng LibreOffice),
  WP4_single_line.png, WP4_containment.png, WP4_report.md

Chạy:  python3 -I 05_electrical/build_electrical.py
Đầu vào (chỉ đọc): 01_design_basis/DESIGN_BASIS.md (v1.1), 02_calc/equipment_schedule.json,
                   03_geometry/geometry.json, 00_inputs/261109_Cai_tao_khu_HN_xuong_3.pdf (nền bản vẽ)
Phạm vi (DESIGN_BASIS §0, §4.6): từ DB.HVAC (gồm DB.HVAC) trở đi; nguồn tới DB.HVAC thuộc gói điện.
"""
import json, math, os, heapq, collections, subprocess, shutil, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
EQS = json.load(open(os.path.join(ROOT, '02_calc', 'equipment_schedule.json'), encoding='utf-8'))
GEO = json.load(open(os.path.join(ROOT, '03_geometry', 'geometry.json'), encoding='utf-8'))
PDF = os.path.join(ROOT, '00_inputs', '261109_Cai_tao_khu_HN_xuong_3.pdf')
EQ = {e['tag']: e for e in EQS['equipment']}
GEQ = {e['id']: e for e in GEO['equipment_proposed']}
EXH = {e['id']: e for e in GEO['exhaust_points']}
ROOMS = {r['id']: r for r in GEO['rooms']}
PTS = {p['id']: p for p in GEO['louvers_discharge_drain']}

# ==============================================================================================
# 1. HẰNG SỐ THIẾT KẾ (nguồn ghi bên cạnh)
# ==============================================================================================
U3, U1 = 380.0, 220.0                    # TCVN 7995 / lưới 380/220 V 50 Hz
PF_STD, PF_VFD, PF_RES = 0.85, 0.95, 1.0  # chỉ đạo WP4: cosφ 0,85; tải qua biến tần 0,95; sưởi điện trở 1,0
ETA_VFD = 0.97                            # hiệu suất biến tần (catalogue điển hình)
ETA_IE3 = {0.6: 0.85, 0.75: 0.825, 1.1: 0.841, 3.0: 0.877, 5.5: 0.896, 7.5: 0.904, 11.0: 0.914, 15.0: 0.921}  # IEC 60034-30-1 IE3 4P
K_TEMP = 0.87     # IEC 60364-5-52 B.52.14, XLPE, 45 °C trên trần
K_GROUP = 0.80    # chỉ đạo WP4 (≈ B.52.20 khay đục lỗ 1 lớp ~5–6 mạch; mạch tải nhẹ < 30 % Iz bỏ qua theo B.52.17)
K_IN = 1.25       # In ≥ 1,25·Ib (tải liên tục, khởi động; giả định WP4)
SLACK = 1.15      # chỉ đạo WP4: +15 % chiều dài ngang
RHO = 0.0225      # Ω·mm²/m Cu ở nhiệt độ làm việc (IEC 60364-5-52 Phụ lục G)
XL = 0.08e-3      # Ω/m cảm kháng cáp nhiều lõi
DU_MAX = 4.0      # % tổng từ DB.HVAC tới phụ tải (chỉ đạo WP4)
DU_FEEDER = 1.5   # % mục tiêu cho cáp DB.HVAC → tủ DDC
TRAY_Z = 3600     # mm BOP máng cáp (như Phase 1 F3-SD-AC-013: CABLE TRAY BOP=FFL3600)
SPARE = 0.20      # 20 % lộ dự phòng / I/O dự phòng
SIZES = [1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240]
# IEC 60364-5-52 Bảng B.52.12 (XLPE, Cu, phương pháp E – cáp nhiều lõi trên khay đục lỗ, 30 °C)
IZ3 = {1.5: 23, 2.5: 32, 4: 42, 6: 54, 10: 75, 16: 100, 25: 127, 35: 158, 50: 192, 70: 246, 95: 298, 120: 346, 150: 399, 185: 456, 240: 538}
IZ2 = {1.5: 26, 2.5: 36, 4: 49, 6: 63, 10: 86, 16: 115, 25: 149, 35: 185, 50: 225, 70: 289, 95: 352, 120: 410, 150: 473, 185: 542, 240: 641}
OD4 = {1.5: 11.5, 2.5: 12.5, 4: 14.0, 6: 15.5, 10: 18.5, 16: 21.5, 25: 25.0, 35: 27.5, 50: 31.0, 70: 35.5, 95: 40.0, 120: 44.0, 150: 49.0, 185: 54.0, 240: 60.0}  # CXV 4 lõi ~ catalogue
BRK = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 320, 400, 500, 630]
MPCB = [(0.63, 1.0), (1.0, 1.6), (1.6, 2.5), (2.5, 4.0), (4.0, 6.3), (6.0, 10.0)]
ISO = [20, 25, 32, 40, 63, 80, 100, 125]
CONDUIT = [(20, 16.9), (25, 21.4), (32, 27.8), (40, 35.4), (50, 44.8)]  # (danh định, đường kính trong mm) ống EMT/PVC
GLAND = [('M20', 13), ('M25', 18), ('M32', 25), ('M40', 32), ('M50', 38), ('M63', 44)]
TRAYS = [(100, 50), (200, 100), (300, 100), (400, 100)]
TRAY_FILL = 0.40

def pe_of(S):
    v = S if S <= 16 else (16 if S <= 35 else S / 2)
    return min(s for s in SIZES if s >= v - 1e-9)

def r1(x, n=1):
    return round(x + 0.0, n)

def vn(x, n=1):
    """số kiểu Việt Nam: dấu phẩy thập phân, dấu chấm nghìn"""
    s = ('{:,.%df}' % n).format(x)
    return s.replace(',', '§').replace('.', ',').replace('§', '.')

# ==============================================================================================
# 2. VỊ TRÍ TỦ (giả định WP4 – trên lối đi kỹ thuật giữa 2 khối, cao độ sàn, ngoài phòng thí nghiệm)
# ==============================================================================================
# Lối đi X 14.385…16.185 (thông thủy 1,8 m), tường khối phải phía lối đi X = 16.185, khối trái X = 14.385.
PANELS = collections.OrderedDict([
    ('DB.HVAC', dict(xy=(15935, 2200), mount='Đặt sàn (đế 100 mm)', dims='1200W × 2100H × 500D', ip='IP42',
                     wall='Vách khối phải (Office) phía lối đi, Y 1.600…2.800 (giữa cửa D-13 và D-15) – đúng vị trí gợi ý WP2')),
    ('DDC-FAN', dict(xy=(15935, 700), mount='Đặt sàn (đế 100 mm)', dims='1200W × 2100H × 500D', ip='IP42',
                     wall='Vách khối phải phía lối đi, Y 100…1.300, cạnh DB.HVAC')),
    ('DDC-AHU-2.02', dict(xy=(16035, -900), mount='Treo tường, đáy +0,5 m', dims='1000W × 1400H × 300D', ip='IP54',
                          wall='Vách khối phải phía lối đi, Y −1.400…−400')),
    ('DDC-AHU-2.01', dict(xy=(14585, 3900), mount='Treo tường, đáy +0,3 m', dims='1000W × 1800H × 400D', ip='IP54',
                          wall='Vách khối trái (Balance/Reporting) phía lối đi, Y 3.400…4.400 (giữa cửa D-05 và D-06)')),
    ('DDC-AHU-2.03', dict(xy=(14585, -4800), mount='Treo tường, đáy +0,3 m', dims='1000W × 1600H × 400D', ip='IP54',
                          wall='Vách khối trái (Reporting) phía lối đi, Y −5.300…−4.300 (đối diện vách đặc)')),
])

# ==============================================================================================
# 3. MẠNG MÁNG CÁP (nút & đoạn, mm; trong nhà +3,6 m, ngoài nhà ốp tường A +3,6 m rồi xuống bãi CDU)
# ==============================================================================================
N = {}
def node(n, x, y):
    N[n] = (x, y)
XC = 15285   # trục máng chính dọc lối đi
for y in [-4800, -1200, -900, 700, 2200, 3900, 5150, 11000, 16000]:
    node('C%d' % y, XC, y)
node('A1', 14100, 5150)                                    # đầu đấu điện AHU-2.01
node('N1', 13000, 11000); node('N2', 10000, 11000); node('N3', 7000, 11000)
node('N4', 7000, 13855); node('N5', 7000, 18000)
node('E1', 13000, 16000); node('E2', 8500, 16000)
node('S1', 11900, -1200); node('S2', 10500, -1200); node('S3', 8575, -1200); node('S4', 4500, -1200)
node('R1', 18000, 2200); node('R2', 21700, 2200); node('R3', 24600, 2200); node('R4', 27300, 2200)
node('W0', 18000, 89); node('W1', 18000, -1482); node('W2', 18000, -5000)
node('V1', 21700, 334); node('V2', 21700, -1270); node('V3', 21700, 7385); node('V4', 21700, 9000)
node('H1', 22906, 9000); node('H2', 24507, 9000); node('H3', 26111, 9000)
node('R5', 27300, 5000); node('R6', 27300, 9000); node('R7', 27300, 9300); node('R8', 27300, -3300); node('R9', 27300, -5500)
node('O1', 28300, -5500); node('O2', 28300, -9000)
CDU_MODULES = ['CDU-2.01-1', 'CDU-2.01-2', 'CDU-2.01-3', 'CDU-2.02', 'CDU-2.03-1', 'CDU-2.03-2']
CDU_XY = {m: (29500, -9600 - 1000 * i) for i, m in enumerate(CDU_MODULES)}  # bãi CDU X 28.500…30.500, Y −15.000…−9.000 (WP2)
for i, m in enumerate(CDU_MODULES):
    node('Y%d' % i, 28800, CDU_XY[m][1])
node('P1', 28300, 9300); node('P2', 28300, 15500)
SCR_PANEL_XY = (29200, 15500)   # tủ điều khiển trọn bộ scrubber (vendor) – mép skid phía tường A

EDGES = []   # (a, b, môi trường, cao độ thêm m)
def edge(a, b, env='in', vz=0.0, name=''):
    EDGES.append(dict(a=a, b=b, env=env, vz=vz, name=name))
cs = ['C-4800', 'C-1200', 'C-900', 'C700', 'C2200', 'C3900', 'C5150', 'C11000', 'C16000']
for a, b in zip(cs[:-1], cs[1:]):
    edge(a, b, name='Máng chính lối đi')
edge('C5150', 'A1', name='Nhánh AHU-2.01')
edge('C11000', 'N1', name='Nhánh bắc trái (ST1)'); edge('N1', 'N2', name='Nhánh bắc trái (ST1)'); edge('N2', 'N3', name='Nhánh bắc trái (ST1)')
edge('N3', 'N4', name='Nhánh hood ST1/Digestion'); edge('N4', 'N5', name='Nhánh hood ST1/Digestion')
edge('C16000', 'E1', name='Nhánh EAF-2.01/Kjeldahl'); edge('E1', 'E2', name='Nhánh EAF-2.01/Kjeldahl')
edge('C-1200', 'S1', name='Nhánh nam trái (Instrument)'); edge('S1', 'S2', name='Nhánh nam trái (Instrument)')
edge('S2', 'S3', name='Nhánh nam trái (Instrument)'); edge('S3', 'S4', name='Nhánh Sample prep/Coding/Dividing')
edge('C2200', 'R1', name='Máng chính khối phải'); edge('R1', 'R2', name='Máng chính khối phải')
edge('R2', 'R3', name='Máng chính khối phải'); edge('R3', 'R4', name='Máng chính khối phải')
edge('R1', 'W0', name='Nhánh Standard 19,72'); edge('W0', 'W1', name='Nhánh Standard 19,72'); edge('W1', 'W2', name='Nhánh Standard 19,72')
edge('R2', 'V1', name='Nhánh hood ST2 (vách trái)'); edge('V1', 'V2', name='Nhánh hood ST2 (vách trái)')
edge('R2', 'V3', name='Nhánh hood ST2 (vách trái)'); edge('V3', 'V4', name='Nhánh hood ST2 (vách trái)')
edge('V4', 'H1', name='Nhánh hood ST2 (vách trên)'); edge('H1', 'H2', name='Nhánh hood ST2 (vách trên)')
edge('H2', 'H3', name='Nhánh hood ST2 (vách trên)'); edge('H3', 'R6', name='Nhánh hood ST2 (vách trên)')
edge('R4', 'R5', name='Máng dọc tường A (trong ST2)'); edge('R5', 'R6', name='Máng dọc tường A (trong ST2)')
edge('R6', 'R7', name='Máng dọc tường A (trong ST2)')
edge('R4', 'R8', name='Máng dọc tường A (trong ST2)'); edge('R8', 'R9', name='Máng dọc tường A (trong ST2)')
edge('R9', 'O1', env='wall', name='Xuyên tường A (ống lồng) → ngoài nhà')
edge('O1', 'O2', env='out', name='Máng ngoài nhà ốp tường A (có nắp) – vượt ram dốc')
edge('O2', 'Y0', env='out', vz=3.1, name='Máng đứng xuống bãi CDU')
for i in range(len(CDU_MODULES) - 1):
    edge('Y%d' % i, 'Y%d' % (i + 1), env='out', name='Máng bãi CDU (có nắp)')
edge('R7', 'P1', env='wall', name='Xuyên tường A (ống lồng) → ngoài nhà')
edge('P1', 'P2', env='out', name='Máng ngoài nhà ốp tường A (có nắp) – tới scrubber')
# nối tủ vào máng (máng đứng từ nóc tủ +2,1 m lên +3,6 m)
PANEL_NODE = {'DB.HVAC': 'C2200', 'DDC-FAN': 'C700', 'DDC-AHU-2.02': 'C-900', 'DDC-AHU-2.01': 'C3900', 'DDC-AHU-2.03': 'C-4800'}
for p, nn in PANEL_NODE.items():
    node('P_' + p, *PANELS[p]['xy'])
    edge('P_' + p, nn, env='riser', vz=1.5, name='Máng đứng nóc tủ %s' % p)

def elen(e):
    (xa, ya), (xb, yb) = N[e['a']], N[e['b']]
    return (abs(xa - xb) + abs(ya - yb)) / 1000.0

ADJ = collections.defaultdict(list)
for i, e in enumerate(EDGES):
    ADJ[e['a']].append((e['b'], i)); ADJ[e['b']].append((e['a'], i))

def dijkstra(src, dst):
    dist = {src: 0.0}; prev = {}
    pq = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if u == dst:
            break
        if d > dist.get(u, 1e18):
            continue
        for v, ei in ADJ[u]:
            nd = d + elen(EDGES[ei]) + EDGES[ei]['vz']
            if nd < dist.get(v, 1e18):
                dist[v] = nd; prev[v] = (u, ei); heapq.heappush(pq, (nd, v))
    path = []; u = dst
    while u != src:
        u, ei = prev[u]; path.append(ei)
    return path[::-1]

INDOOR_NODES = [n for n in N if not n.startswith('P_') and not n.startswith('Y') and n not in ('O1', 'O2', 'P1', 'P2')]
def nearest(xy, pool=None):
    pool = pool or INDOOR_NODES
    return min(pool, key=lambda n: (abs(N[n][0] - xy[0]) + abs(N[n][1] - xy[1])))

# độ dài đứng ở đầu thiết bị theo loại vị trí (m)
END_VZ = {'ceil': 1.0, 'hood': 1.5, 'wall': 2.6, 'panel': 0.0, 'cdu': 1.0, 'scr': 4.1, 'duct': 1.0}
FLEX = {'ceil': 1.0, 'hood': 1.0, 'cdu': 1.0, 'scr': 1.0}
PANEL_INTERNAL = 1.0   # m đi dây trong tủ / tới cầu đấu

def route(src_panel, xy, end, pool=None, dst_node=None):
    """trả về dict chiều dài cáp: ngang trên máng, ngang ống (drop), đứng, tổng; danh sách đoạn máng"""
    a = 'P_' + src_panel
    if dst_node is None:
        dst_node = nearest(xy, pool)
    path = dijkstra(a, dst_node)
    tray_h = sum(elen(EDGES[i]) for i in path)
    vz = sum(EDGES[i]['vz'] for i in path)
    drop_h = (abs(N[dst_node][0] - xy[0]) + abs(N[dst_node][1] - xy[1])) / 1000.0
    v_end = END_VZ[end]
    L_h = tray_h + drop_h
    total = L_h * SLACK + vz + v_end + PANEL_INTERNAL
    return dict(path=path, node=dst_node, tray_h=r1(tray_h, 2), drop_h=r1(drop_h, 2), L_h=r1(L_h, 2),
                L_v=r1(vz + v_end + PANEL_INTERNAL, 2), L=r1(total, 1))

def route_pts(xy1, end1, xy2, end2, pool1=None, pool2=None):
    n1, n2 = nearest(xy1, pool1), nearest(xy2, pool2)
    path = dijkstra(n1, n2) if n1 != n2 else []
    tray_h = sum(elen(EDGES[i]) for i in path)
    vz = sum(EDGES[i]['vz'] for i in path)
    d1 = (abs(N[n1][0] - xy1[0]) + abs(N[n1][1] - xy1[1])) / 1000.0
    d2 = (abs(N[n2][0] - xy2[0]) + abs(N[n2][1] - xy2[1])) / 1000.0
    L_h = tray_h + d1 + d2
    return dict(path=path, L_h=r1(L_h, 2), L_v=r1(vz + END_VZ[end1] + END_VZ[end2], 2),
                L=r1(L_h * SLACK + vz + END_VZ[end1] + END_VZ[end2], 1), drop_h=r1(d1 + d2, 2))

OUT_POOL = ['Y%d' % i for i in range(len(CDU_MODULES))]
def pool_for(end):
    return OUT_POOL if end == 'cdu' else (['P2'] if end == 'scr' else None)

# ==============================================================================================
# 4. VỊ TRÍ ĐIỂM ĐẤU THIẾT BỊ
# ==============================================================================================
XY = {
    'AHU-2.01': (14000, 5150), 'AHU-2.02': (27000, -3300), 'AHU-2.03': (10500, -3000),
    'AHU-2.01c': tuple(GEQ['AHU-2.01']['center_mm']), 'AHU-2.02c': tuple(GEQ['AHU-2.02']['center_mm']),
    'AHU-2.03c': tuple(GEQ['AHU-2.03']['center_mm']),
    'FAF-2.01': tuple(GEQ['FAF-2.01']['center_mm']), 'FAF-2.02': (24600, 850),
    'EAF-2.01': tuple(GEQ['EAF-2.01']['center_mm']), 'EAF-2.02': tuple(GEQ['EAF-2.02']['center_mm']),
    'EAF-2.03': tuple(GEQ['EAF-2.03']['center_mm']),
    'DH-2.01': tuple(GEQ['DH-2.01']['center_mm']),
    'CP-2.01': (11000, -2900), 'CP-2.02': (13500, 4300), 'CP-2.03': (26500, -2300),
    'SCR-2.01': SCR_PANEL_XY, 'LV-2.02': tuple(PTS['LV-2.02']['xy_mm']),
}
XY.update(CDU_XY)
HOODS = [e for e in GEO['exhaust_points'] if e['type'] == 'fume_hood_1600']
HOOD_TAG = {'DIG': 'FHF-2.01', 'ST1': 'FHF-2.02', 'ST2': 'FHF-2.03', 'STD20': 'FHF-2.04'}
assert len(HOODS) == 15 and sum(1 for h in HOODS if h['acid']) == 4
assert sum(EQ[t]['qty'] for t in HOOD_TAG.values()) == 15

# ==============================================================================================
# 5. DANH SÁCH MẠCH (phụ tải cuối + cáp lộ tổng)
# ==============================================================================================
LOADS = []
def load(cid, src, desc, tag, kind, P, ph=3, eta=1.0, eta_vfd=1.0, pf=PF_STD, pf_m=PF_STD, fla=None, xy=None,
         end='ceil', neutral=False, df=(1, 1, 1), rcd='', iso=True, grp='', note='', shield=False, cable=True):
    LOADS.append(dict(id=cid, src=src, desc=desc, tag=tag, kind=kind, P=P, ph=ph, U=U3 if ph == 3 else U1, eta=eta,
                      eta_vfd=eta_vfd, pf=pf, pf_m=pf_m, fla=fla, xy=xy, end=end, neutral=neutral, df=df, rcd=rcd,
                      iso=iso, grp=grp, note=note, shield=shield, cable=cable))

# --- DB.HVAC: CDU (cấp trực tiếp, cách ly IP65 tại mỗi module) + scrubber ---
for m in CDU_MODULES:
    base = m[:8]
    e = EQ[base]
    P = e['electrical_kW_per_unit']; fla = e['FLA_A_est_per_unit']
    df = (1.0, 0.5 if base == 'CDU-2.03' else 0.0, 0.7)
    load('DB-' + m, 'DB.HVAC', 'Dàn nóng %s (inverter, 380V/3P+N)' % m, m, 'INV', P, fla=fla, xy=XY[m], end='cdu',
         neutral=True, df=df, rcd='ELCB 300 mA (loại B/tương thích inverter – theo yêu cầu hãng)', grp='CDU',
         note='Cấp trực tiếp từ DB.HVAC; AHU điều khiển CDU qua bộ DX-kit (tín hiệu), không đóng cắt nguồn')
e = EQ['SCR-2.01']
load('DB-SCR-2.01', 'DB.HVAC', 'Tủ điều khiển trọn bộ scrubber SCR-2.01 (quạt 5,5 + bơm 3 + định lượng 0,18 kW)', 'SCR-2.01',
     'PKG', e['electrical_kW'], eta=0.88, fla=e['FLA_A_est'], xy=XY['SCR-2.01'], end='scr', neutral=True, grp='SCR',
     note='HVAC cấp nguồn tới cầu dao cách ly + tủ vendor; tủ vendor có khởi động/biến tần riêng')

# --- DDC-AHU-x ---
AHU_CFG = {
    'AHU-2.01': dict(panel='DDC-AHU-2.01', stages=3, cp='CP-2.02'),
    'AHU-2.02': dict(panel='DDC-AHU-2.02', stages=3, cp='CP-2.03'),
    'AHU-2.03': dict(panel='DDC-AHU-2.03', stages=3, cp='CP-2.01'),
}
for a, c in AHU_CFG.items():
    e = EQ[a]; p = c['panel']
    fm = e['fan_motor_kW']
    load('%s-M1' % p, p, 'Quạt cấp %s %s kW qua biến tần (cáp màn chắn)' % (a, vn(fm, 1)), a + '-SF', 'VFD', fm,
         eta=ETA_IE3[fm], eta_vfd=ETA_VFD, pf=PF_VFD, xy=XY[a], end='ceil', grp='AHU fan', shield=True)
    hk = e['heater_kW'] / c['stages']
    for s in range(c['stages']):
        load('%s-H%d' % (p, s + 1), p, 'Sưởi điện %s cấp %d/%d – %s kW (%s)' % (a, s + 1, c['stages'], vn(hk, 1),
             'SSR/thyristor điều biến' if s == 0 else 'contactor'), a + '-EH%d' % (s + 1), 'HEATER', hk,
             pf=PF_RES, xy=XY[a], end='ceil', df=(0, 1, 1), rcd='RCD 300 mA 4P chung nhóm sưởi', grp='Heater')
    ce = EQ[c['cp']]
    load('%s-CP' % p, p, 'Bơm nước ngưng %s (220V)' % c['cp'], c['cp'], '1P', ce['electrical_kW'], ph=1,
         fla=ce['FLA_A_est'], xy=XY[c['cp']], end='ceil', rcd='RCBO 30 mA', grp='CP')
    if a == 'AHU-2.03':
        de = EQ['DH-2.01']
        load('%s-DH' % p, p, 'Máy hút ẩm DH-2.01 (220V, bơm ngưng tích hợp)', 'DH-2.01', '1P', de['electrical_kW'], ph=1,
             fla=de['FLA_A_est'], xy=XY['DH-2.01'], end='ceil', df=(1, 0.5, 1), rcd='RCBO 30 mA', grp='DH')
    load('%s-SK' % p, p, 'Ổ cắm bảo trì 1P 16A IP66 tại sàn %s' % a, a + '-SO', 'SOCKET', 1.0, ph=1, pf=PF_RES,
         xy=(XY[a + 'c'][0], XY[a + 'c'][1]), end='ceil', df=(0.2, 0.2, 0.2), rcd='RCBO 30 mA', iso=False, grp='Socket')
    load('%s-CT' % p, p, 'Nguồn điều khiển (DDC, 24 VAC/DC, quạt tủ)', '', 'CTRL', 0.3, ph=1, pf=PF_RES, iso=False,
         cable=False, end='panel', grp='Ctrl')

# --- DDC-FAN ---
for t in ['FAF-2.01', 'FAF-2.02']:
    fm = EQ[t]['motor_kW']
    load('DDC-FAN-%s' % t, 'DDC-FAN', 'Quạt cấp bù %s %s kW qua biến tần (cáp màn chắn)' % (t, vn(fm, 1)), t, 'VFD', fm,
         eta=ETA_IE3[fm], eta_vfd=ETA_VFD, pf=PF_VFD, xy=XY[t], end='ceil', grp='FAF', shield=True)
for t in ['EAF-2.01', 'EAF-2.02', 'EAF-2.03']:
    fm = EQ[t]['motor_kW']
    load('DDC-FAN-%s' % t, 'DDC-FAN', 'Quạt hút %s %s kW – khởi động trực tiếp (MPCB + contactor)' % (t, vn(fm, 2)), t,
         'DOL', fm, eta=ETA_IE3[fm], fla=EQ[t]['FLA_A_est'], xy=XY[t], end='ceil', grp='EAF')
for i, h in enumerate(HOODS):
    tag = HOOD_TAG[h['room']]
    load('DDC-FAN-HF%02d' % (i + 1), 'DDC-FAN', 'Quạt tủ hood %s (%s%s) 0,6 kW' % (h['id'], tag, ', ACID' if h['acid'] else ''),
         h['id'], 'DOL', 0.6, eta=ETA_IE3[0.6], fla=EQ[tag]['FLA_A_est_per_unit'], xy=tuple(h['xy_mm']), end='hood',
         grp='Hood fan', note='Quạt thuộc gói nội thất; HVAC cấp nguồn + cách ly tại tủ + điều khiển')
load('DDC-FAN-CT', 'DDC-FAN', 'Nguồn điều khiển + biến áp 24 VAC cho 27 bộ chấp hành van gió', '', 'CTRL', 0.5, ph=1,
     pf=PF_RES, iso=False, cable=False, end='panel', grp='Ctrl')
FEEDERS = ['DDC-AHU-2.01', 'DDC-AHU-2.02', 'DDC-AHU-2.03', 'DDC-FAN']

# ==============================================================================================
# 6. TÍNH DÒNG, CHỌN BẢO VỆ, CÁP, SỤT ÁP
# ==============================================================================================
def pick(lst, v):
    for x in lst:
        if x >= v - 1e-9:
            return x
    raise ValueError(v)

def du_pct(I, L, S, ph, pf, npar=1):
    sinp = math.sqrt(max(0.0, 1 - pf * pf))
    z = RHO / S * pf + XL * sinp
    if ph == 3:
        return math.sqrt(3) * I * L * z / U3 * 100 / npar
    return 2 * I * L * z / U1 * 100 / npar

for d in LOADS:
    d['Pin'] = d['P'] / (d['eta'] * d['eta_vfd'])
    if d['ph'] == 3:
        d['Ib_calc'] = d['Pin'] * 1000 / (math.sqrt(3) * U3 * d['pf'])
    else:
        d['Ib_calc'] = d['Pin'] * 1000 / (U1 * d['pf'])
    d['Ib'] = max(d['Ib_calc'], d['fla'] or 0)
    if d['kind'] == 'VFD':      # dòng phía động cơ (cáp VFD→động cơ)
        d['Icab'] = d['P'] / d['eta'] * 1000 / (math.sqrt(3) * U3 * d['pf_m'])
        d['pf_cab'] = d['pf_m']
    else:
        d['Icab'] = d['Ib']; d['pf_cab'] = d['pf']
    # thiết bị bảo vệ
    if d['kind'] == 'DOL':
        lo, hi = next(r for r in MPCB if r[0] <= d['Ib'] <= r[1])
        d['prot'] = 'MPCB 3P %s–%s A + contactor AC3' % (vn(lo, 2).rstrip('0').rstrip(','), vn(hi, 1).rstrip('0').rstrip(','))
        d['In'] = hi; d['prot_type'] = 'MPCB'
    else:
        d['In'] = max(pick(BRK, K_IN * d['Ib']), 10 if d['kind'] in ('VFD', 'INV', 'PKG', 'HEATER') else 6)
        if d['kind'] == 'SOCKET':
            d['In'] = 16
        if d['src'] == 'DB.HVAC':
            d['prot'] = 'MCCB 4P %d A, Icu ≥ 25 kA' % d['In'] if d['neutral'] else 'MCCB 3P %d A, Icu ≥ 25 kA' % d['In']
            d['prot_type'] = 'MCCB'
        else:
            poles = '3P' if d['ph'] == 3 else '1P+N'
            d['prot'] = 'MCB %s %d A (C), 10 kA' % (poles, d['In'])
            d['prot_type'] = 'MCB'
            if d['kind'] == 'VFD':
                d['prot'] = 'MCB 3P %d A (C) 10 kA + biến tần %s kW' % (d['In'], vn(d['P'], 1))
            if d['kind'] == 'HEATER':
                d['prot'] = 'MCB 3P %d A (C) + %s' % (d['In'], 'SSR/thyristor 3P' if d['id'].endswith('H1') else 'contactor AC1')
            if d['kind'] in ('1P', 'SOCKET'):
                d['prot'] = 'RCBO 1P+N %d A 30 mA (C)' % d['In']
    if d['cable']:
        d['route'] = route(d['src'], d['xy'], d['end'], pool=pool_for(d['end']))

# feeders: Ib = Σ Ib phụ tải trong tủ (cộng số học, an toàn), toàn bộ tải tủ DDC có thể chạy đồng thời (đông/nồm)
FEED = []
for p in FEEDERS:
    ch = [d for d in LOADS if d['src'] == p]
    Pin = sum(d['Pin'] for d in ch)
    Ib = sum(d['Ib'] for d in ch)
    In = pick(BRK, max(K_IN * Ib, max(d['In'] for d in ch) * 1.6))
    f = dict(id='DB-' + p, src='DB.HVAC', to=p, desc='Cáp lộ tổng DB.HVAC → %s' % p, Pin=Pin, Ib=Ib, In=In, ph=3,
             pf=0.9, neutral=True, prot='MCCB 4P %d A, Icu ≥ 25 kA' % In, prot_type='MCCB', kind='FEEDER')
    r = dijkstra('P_DB.HVAC', 'P_' + p)
    tray_h = sum(elen(EDGES[i]) for i in r)
    vz = sum(EDGES[i]['vz'] for i in r)
    f['route'] = dict(path=r, tray_h=r1(tray_h, 2), drop_h=0.0, L_h=r1(tray_h, 2), L_v=r1(vz + 2 * PANEL_INTERNAL, 2),
                      L=r1(tray_h * SLACK + vz + 2 * PANEL_INTERNAL, 1))
    FEED.append(f)

def size_cable(I_des, In, L, ph, pf, du_budget, minS=2.5):
    tab = IZ3 if ph == 3 else IZ2
    for S in SIZES:
        if S < minS:
            continue
        Iz = tab[S] * K_TEMP * K_GROUP
        if Iz + 1e-9 < In:
            continue
        du = du_pct(I_des, L, S, ph, pf)
        if du <= du_budget:
            return S, Iz, du
    raise ValueError('no size')

for f in FEED:
    S, Iz, du = size_cable(f['Ib'], f['In'], f['route']['L'], 3, f['pf'], DU_FEEDER, minS=4)
    f.update(S=S, Iz=Iz, du=du, du_cum=du, PE=pe_of(S))
    f['cable'] = 'CXV 4×%s + E%s' % (vn(S, 1).replace(',0', ''), vn(pe_of(S), 1).replace(',0', ''))
FEED_BY = {f['to']: f for f in FEED}

def sname(S):
    return vn(S, 1).replace(',0', '')

for d in LOADS:
    if not d['cable']:
        d.update(S=None, du=0.0, du_cum=FEED_BY[d['src']]['du'] if d['src'] in FEED_BY else 0.0)
        continue
    up = FEED_BY[d['src']]['du'] if d['src'] in FEED_BY else 0.0
    S, Iz, du = size_cable(d['Icab'], d['In'], d['route']['L'], d['ph'], d['pf_cab'], DU_MAX - up)
    d.update(S=S, Iz=Iz, du=du, du_cum=du + up, PE=pe_of(S))
    s = sname(S); pe = sname(pe_of(S))
    if d['ph'] == 1:
        d['cable'] = 'CXV 2×%s + E%s' % (s, pe); d['cores'] = 3
    elif d['neutral']:
        d['cable'] = 'CXV 4×%s + E%s' % (s, pe); d['cores'] = 5
    else:
        d['cable'] = 'CXV 3×%s + E%s' % (s, pe); d['cores'] = 4
    if d['shield']:
        d['cable'] = d['cable'].replace('CXV', 'CXV/S (màn chắn Cu, cáp biến tần)')
    d['OD'] = OD4[S] * (1.08 if d['cores'] == 5 else (0.92 if d['cores'] == 3 else 1.0)) + (2.0 if d['shield'] else 0.0)
    d['iso_A'] = pick(ISO, d['In']) if d['iso'] else None
for f in FEED:
    f['OD'] = OD4[f['S']] * 1.08

# ==============================================================================================
# 7. DB.HVAC – TỔNG ĐẶT, NHU CẦU 3 KỊCH BẢN, LỘ DỰ PHÒNG
# ==============================================================================================
CASES = ['Hè (lạnh 100 %, sưởi 0)', 'Đông (sưởi 100 %, lạnh chỉ CDU-2.03 50 %)', 'Nồm/khử ẩm (lạnh 70 % + sưởi gia nhiệt lại 100 % + hút ẩm)']
def case_kva(k):
    kw = sum(d['Pin'] * d['df'][k] for d in LOADS)
    kva = sum(d['Pin'] * d['df'][k] / d['pf'] for d in LOADS)
    return kw, kva
P_conn_rated = sum(d['P'] for d in LOADS if d['kind'] != 'CTRL' and d['kind'] != 'SOCKET')
P_conn_in = sum(d['Pin'] for d in LOADS)
S_conn = sum(d['Pin'] / d['pf'] for d in LOADS)
cases = []
for k in range(3):
    kw, kva = case_kva(k)
    cases.append(dict(case=CASES[k], kW=r1(kw, 1), kVA=r1(kva, 1), I_A=r1(kva * 1000 / (math.sqrt(3) * U3), 1)))
dmax = max(cases, key=lambda c: c['kVA'])
Ib_db = dmax['I_A']
Ir_set = math.ceil(1.15 * Ib_db / 10) * 10           # chỉnh Ir ≈ 1,15·Ib_max (dư cho 20 % lộ dự phòng)
In_db = pick([250, 400, 630], Ir_set)
BUSBAR = In_db
# cáp cấp nguồn khuyến nghị cho gói điện (Iz ≥ Ir; hệ số như trên – gói điện kiểm lại theo tuyến thực)
inc = None
for npar in (1, 2, 3):
    for S in SIZES:
        if S < 50:
            continue
        if IZ3[S] * K_TEMP * K_GROUP * npar >= Ir_set:
            inc = (npar, S); break
    if inc:
        break
INCOMING = dict(P_conn_rated_kW=r1(P_conn_rated, 1), P_conn_input_kW=r1(P_conn_in, 1), S_conn_kVA=r1(S_conn, 1),
                I_conn_A=r1(S_conn * 1000 / (math.sqrt(3) * U3), 1), cases=cases, max_case=dmax['case'],
                demand_kW=dmax['kW'], demand_kVA=dmax['kVA'], Ib_max_A=Ib_db,
                demand_factor_overall=r1(dmax['kW'] / P_conn_in, 2),
                MCCB='MCCB 4P %d AF, Ir chỉnh %d A (0,4–1,0 In), Icu ≥ 36 kA (giả định – gói điện xác nhận Ik)' % (In_db, Ir_set),
                In_A=In_db, Ir_A=Ir_set, busbar='Thanh cái Cu %d A 3P+N+PE (≈ 40×10 mm/pha, N = pha, PE 40×5), Icw ≥ 25 kA/1 s' % BUSBAR,
                incoming_cable='%d × CXV 4×%s mm² + %d × CV 1×%s mm² (PE)' % (inc[0], sname(inc[1]), inc[0], sname(pe_of(inc[1]))),
                incoming_cable_note='Thông tin cho gói điện: Iz = %d × %d × %.2f × %.2f = %.0f A ≥ Ir %d A; chiều dài/tuyến, Ik, sụt áp tới DB.HVAC (khuyến nghị ≤ 1 %%) do gói điện tính.' %
                                    (inc[0], IZ3[inc[1]], K_TEMP, K_GROUP, inc[0] * IZ3[inc[1]] * K_TEMP * K_GROUP, Ir_set))

def spare_n(n):
    return int(math.ceil(SPARE * n))
DB_OUT = [d for d in LOADS if d['src'] == 'DB.HVAC'] + FEED
SPARES = {'DB.HVAC': dict(n_used=len(DB_OUT), n_spare=spare_n(len(DB_OUT)),
                          spec='MCCB 4P 63 A, 25 kA (lắp sẵn)' )}
for p in FEEDERS:
    used = [d for d in LOADS if d['src'] == p and d['kind'] != 'CTRL']
    SPARES[p] = dict(n_used=len(used), n_spare=spare_n(len(used)),
                     spec='MCB 3P 16 A (C) + 1 MCB 1P 10 A' if p != 'DDC-FAN' else 'MPCB 3P 1,6–2,5 A + contactor (lắp sẵn)')

# ==============================================================================================
# 8. I/O DDC & CẢM BIẾN
# ==============================================================================================
IO = []
def io(panel, typ, desc, dev, loc, end='ceil', cab='', internal=False, grp=''):
    IO.append(dict(panel=panel, type=typ, desc=desc, device=dev, loc=loc, end=end, cable=cab, internal=internal, grp=grp))

C_DI = 'Cáp ĐK chống nhiễu 2C×1,0 mm²'
C_DO = 'Cáp ĐK chống nhiễu 3C×1,0 mm²'
C_TH = 'Cáp tín hiệu chống nhiễu 4C×0,75 mm²'
C_AI = 'Cáp tín hiệu chống nhiễu 3C×0,75 mm²'
C_AO = 'Cáp tín hiệu chống nhiễu 2C×0,75 mm²'
C_HS = 'Cáp ĐK chống nhiễu 4C×1,0 mm²'
C_FA = 'Cáp chống cháy FR 2C×1,5 mm²'
C_485 = 'Cáp mạng RS485 1 đôi xoắn chống nhiễu 0,75 mm² (BACnet MS/TP/Modbus)'
C_DX = 'Cáp truyền tín hiệu DX-kit ↔ CDU 2C×1,25 mm² chống nhiễu'
CABLE_OD = {C_DI: 7.0, C_DO: 7.5, C_TH: 8.0, C_AI: 7.5, C_AO: 7.0, C_HS: 8.5, C_FA: 9.0, C_485: 6.0, C_DX: 7.5}

ROOM_SENS = {'AHU-2.01': ['DIG', 'KJE', 'ST1', 'BAL', 'PRE', 'COD', 'DIV', 'REP'],
             'AHU-2.02': ['ST2', 'STD20', 'STD12', 'OFF'], 'AHU-2.03': ['INS', 'INS']}
CDU_OF = {'AHU-2.01': ['CDU-2.01-1', 'CDU-2.01-2', 'CDU-2.01-3'], 'AHU-2.02': ['CDU-2.02'], 'AHU-2.03': ['CDU-2.03-1', 'CDU-2.03-2']}
for a, c in AHU_CFG.items():
    p = c['panel']; A = a + 'c'
    io(p, 'DO', 'Lệnh chạy biến tần quạt %s' % a, 'Biến tần (trong tủ)', None, internal=True, grp='Quạt')
    io(p, 'DI', 'Trạng thái chạy biến tần', 'Biến tần', None, internal=True, grp='Quạt')
    io(p, 'DI', 'Lỗi biến tần', 'Biến tần', None, internal=True, grp='Quạt')
    io(p, 'AO', 'Đặt tốc độ biến tần 0–10 V', 'Biến tần', None, internal=True, grp='Quạt')
    io(p, 'AI', 'Chênh áp quạt (k-factor → lưu lượng gió cấp) 0–10 V', 'Bộ chuyển đổi chênh áp 0–1000 Pa', A, 'duct', C_AI, grp='Quạt')
    io(p, 'DI', 'Công tắc dòng gió quạt (airflow switch) – liên động sưởi', 'Công tắc chênh áp 20–300 Pa', A, 'duct', C_DI, grp='Quạt')
    io(p, 'DI', 'Chênh áp lọc G4 bẩn', 'Công tắc chênh áp 50–500 Pa', A, 'duct', C_DI, grp='Lọc')
    io(p, 'DI', 'Chênh áp lọc F8 bẩn', 'Công tắc chênh áp 50–500 Pa', A, 'duct', C_DI, grp='Lọc')
    io(p, 'AI', 'Nhiệt độ gió cấp', 'Cảm biến T/RH ống gió (gió cấp)', A, 'duct', C_TH, grp='Cảm biến AHU')
    io(p, 'AI', 'Độ ẩm gió cấp', '(chung cảm biến trên)', A, 'duct', '', grp='Cảm biến AHU')
    io(p, 'AI', 'Nhiệt độ gió hồi', 'Cảm biến T/RH ống gió (gió hồi)', A, 'duct', C_TH, grp='Cảm biến AHU')
    io(p, 'AI', 'Độ ẩm gió hồi', '(chung cảm biến trên)', A, 'duct', '', grp='Cảm biến AHU')
    io(p, 'DO', 'Đóng contactor sưởi cấp 2', 'Contactor AC1 (trong tủ)', None, internal=True, grp='Sưởi')
    io(p, 'DO', 'Đóng contactor sưởi cấp 3', 'Contactor AC1 (trong tủ)', None, internal=True, grp='Sưởi')
    io(p, 'AO', 'Điều biến sưởi cấp 1 (SSR/thyristor 0–10 V)', 'Bộ điều khiển thyristor 3P (trong tủ)', None, internal=True, grp='Sưởi')
    io(p, 'DI', 'Bảo vệ quá nhiệt sưởi (thermostat an toàn reset tay) – cắt cứng', 'Thermostat an toàn (kèm AHU)', A, 'duct', C_DI, grp='Sưởi')
    io(p, 'DI', 'Lỗi mạch sưởi (RCD/MCB nhảy)', 'Tiếp điểm phụ RCD/MCB', None, internal=True, grp='Sưởi')
    io(p, 'DO', 'Mở van gió tươi OA (on/off, lò xo hồi)', 'Bộ chấp hành van gió 24 VAC', A, 'duct', C_DO, grp='Van gió')
    for m in CDU_OF[a]:
        io(p, 'DO', 'Cho phép chạy %s (DX-kit)' % m, 'Bộ DX-kit tại AHU', A, 'duct', C_DI, grp='CDU')
        io(p, 'AO', 'Yêu cầu công suất %s 0–10 V' % m, 'Bộ DX-kit tại AHU', A, 'duct', C_AO, grp='CDU')
        io(p, 'DI', 'Trạng thái chạy %s' % m, 'Bộ DX-kit tại AHU', A, 'duct', C_DI, grp='CDU')
        io(p, 'DI', 'Lỗi %s' % m, 'Bộ DX-kit tại AHU', A, 'duct', C_DI, grp='CDU')
    io(p, 'DI', 'Báo mức nước ngưng cao %s (cắt cứng CDU)' % c['cp'], 'Phao báo mức của bơm', c['cp'], 'ceil', C_DI, grp='Nước ngưng')
    io(p, 'DI', 'Tín hiệu báo cháy – dừng AHU (cắt cứng biến tần + sưởi)', 'Module rơ le báo cháy (gói PCCC, đặt cạnh DB.HVAC)',
       'FA', 'panel', C_FA, grp='Báo cháy')
    io(p, 'DI', 'Công tắc chọn Tay/Tự động (H-O-A)', 'Công tắc mặt tủ', None, internal=True, grp='Tủ')
    for rid in ROOM_SENS[a]:
        r = ROOMS[rid]
        io(p, 'AI', 'Nhiệt độ phòng %s' % r['name'], 'Cảm biến T/RH phòng gắn tường', 'room:' + rid, 'wall', C_TH, grp='Phòng')
        io(p, 'AI', 'Độ ẩm phòng %s' % r['name'], '(chung cảm biến trên)', 'room:' + rid, 'wall', '', grp='Phòng')
    if a == 'AHU-2.02':
        io(p, 'AI', 'Nhiệt độ ngoài trời (chọn chế độ hè/đông/nồm)', 'Cảm biến T/RH ngoài trời có che nắng', 'LV-2.02', 'wall', C_TH, grp='Ngoài trời')
        io(p, 'AI', 'Độ ẩm ngoài trời', '(chung cảm biến trên)', 'LV-2.02', 'wall', '', grp='Ngoài trời')
    if a == 'AHU-2.03':
        io(p, 'DO', 'Cho phép chạy máy hút ẩm DH-2.01 (RH > 55 %)', 'Tiếp điểm điều khiển từ xa DH', 'DH-2.01', 'ceil', C_DI, grp='Hút ẩm')
        io(p, 'DI', 'Trạng thái chạy DH-2.01', 'Tiếp điểm DH', 'DH-2.01', 'ceil', C_DI, grp='Hút ẩm')
        io(p, 'DI', 'Lỗi/bồn đầy DH-2.01', 'Tiếp điểm DH', 'DH-2.01', 'ceil', C_DI, grp='Hút ẩm')

p = 'DDC-FAN'
for t in ['FAF-2.01', 'FAF-2.02']:
    io(p, 'DO', 'Lệnh chạy biến tần %s' % t, 'Biến tần (trong tủ)', None, internal=True, grp=t)
    io(p, 'DI', 'Trạng thái chạy %s' % t, 'Biến tần', None, internal=True, grp=t)
    io(p, 'DI', 'Lỗi biến tần %s' % t, 'Biến tần', None, internal=True, grp=t)
    io(p, 'AO', 'Đặt tốc độ %s' % t, 'Biến tần', None, internal=True, grp=t)
    io(p, 'AI', 'Áp suất tĩnh ống gió bù %s (PID)' % t, 'Cảm biến áp suất ống gió 0–500 Pa', t, 'ceil', C_AI, grp=t)
    io(p, 'DI', 'Chênh áp lọc G4 %s' % t, 'Công tắc chênh áp 50–500 Pa', t, 'ceil', C_DI, grp=t)
    io(p, 'DI', 'Chênh áp lọc F8 %s' % t, 'Công tắc chênh áp 50–500 Pa', t, 'ceil', C_DI, grp=t)
    io(p, 'DO', 'Mở van gió tươi cửa hút %s' % t, 'Bộ chấp hành van gió 24 VAC', t, 'ceil', C_DO, grp=t)
for t in ['EAF-2.01', 'EAF-2.02', 'EAF-2.03']:
    io(p, 'DO', 'Lệnh chạy %s (contactor)' % t, 'Contactor (trong tủ)', None, internal=True, grp=t)
    io(p, 'DI', 'Trạng thái chạy %s (tiếp điểm phụ)' % t, 'Contactor', None, internal=True, grp=t)
    io(p, 'DI', 'Quá tải %s (MPCB)' % t, 'MPCB', None, internal=True, grp=t)
    io(p, 'DI', 'Xác nhận dòng gió %s' % t, 'Công tắc chênh áp 20–300 Pa', t, 'ceil', C_DI, grp=t)
for i, h in enumerate(HOODS):
    hx = 'hood:' + h['id']
    io(p, 'DI', 'Yêu cầu chạy quạt tủ %s (công tắc tại tủ)' % h['id'], 'Hộp công tắc ON/OFF + đèn tại tủ hood', hx, 'hood', C_HS, grp='Hood')
    io(p, 'DO', 'Đóng contactor quạt tủ %s' % h['id'], 'Contactor (trong tủ)', None, internal=True, grp='Hood')
    io(p, 'DI', 'Xác nhận chạy quạt tủ %s (rơ le dòng)' % h['id'], 'Rơ le dòng (trong tủ)', None, internal=True, grp='Hood')
    io(p, 'DO', 'Mở van gió bù MD-%s' % h['id'], 'Bộ chấp hành van gió 24 VAC', hx, 'ceil', C_DO, grp='Hood')
GROUP_DAMPERS = collections.OrderedDict([
    ('MD-G-DIG', ('EAF-2.01', ['DIG-CH1'])), ('MD-G-KJE', ('EAF-2.01', ['KJE-CH1'])),
    ('MD-G-ST1', ('EAF-2.01', ['ST1-CC1', 'ST1-CC2'])), ('MD-G-PRE', ('EAF-2.03', ['PRE-AH1', 'PRE-AH2'])),
    ('MD-G-ST2', ('EAF-2.02', ['ST2-CC1', 'ST2-CC2', 'ST2-SK1', 'ST2-SK2'])),
    ('MD-G-STD20', ('EAF-2.02', ['STD20-SK1'])), ('MD-G-STD12', ('EAF-2.02', ['STD12-CC1']))])
for g, (fan, pts) in GROUP_DAMPERS.items():
    xs = [EXH[q]['xy_mm'] for q in pts]
    XY[g] = (sum(x[0] for x in xs) / len(xs), sum(x[1] for x in xs) / len(xs))
    io(p, 'DO', 'Mở van gió bù %s (theo %s chạy)' % (g, fan), 'Bộ chấp hành van gió 24 VAC', g, 'ceil', C_DO, grp='Van bù EAF')
io(p, 'DO', 'Lệnh chạy scrubber SCR-2.01 (từ xa)', 'Tủ điều khiển scrubber (vendor)', 'SCR-2.01', 'scr', C_DI, grp='SCR')
io(p, 'DI', 'Trạng thái chạy scrubber (cho phép quạt hood acid)', 'Tủ scrubber (tiếp điểm khô)', 'SCR-2.01', 'scr', C_DI, grp='SCR')
io(p, 'DI', 'Lỗi tổng scrubber (pH, mức, quạt, bơm)', 'Tủ scrubber (tiếp điểm khô)', 'SCR-2.01', 'scr', C_DI, grp='SCR')
io(p, 'DI', 'Tín hiệu báo cháy – dừng FAF, đóng van bù', 'Module rơ le báo cháy (gói PCCC)', 'FA', 'panel', C_FA, grp='Báo cháy')
io(p, 'DI', 'Công tắc chọn Tay/Tự động (H-O-A)', 'Công tắc mặt tủ', None, internal=True, grp='Tủ')

def loc_xy(loc):
    if loc is None:
        return None
    if loc.startswith('room:'):
        return tuple(ROOMS[loc[5:]]['centroid_mm'])
    if loc.startswith('hood:'):
        return tuple(EXH[loc[5:]]['xy_mm'])
    if loc == 'FA':
        return PANELS['DB.HVAC']['xy']
    return XY[loc]

ROUTED_CTRL = []   # (cable type, length, path) cho máng
for q in IO:
    if q['internal'] or not q['cable']:
        q['L'] = 0.0; continue
    xy = loc_xy(q['loc'])
    if q['loc'] == 'FA':
        r = dict(path=[], L=r1(3.0 + (abs(PANELS[q['panel']]['xy'][0] - xy[0]) + abs(PANELS[q['panel']]['xy'][1] - xy[1])) / 1000 * SLACK + 3.0, 1), drop_h=0)
        if q['panel'] != 'DB.HVAC':
            rr = dijkstra('P_DB.HVAC', 'P_' + q['panel'])
            th = sum(elen(EDGES[i]) for i in rr); vz = sum(EDGES[i]['vz'] for i in rr)
            r = dict(path=rr, L=r1(th * SLACK + vz + 2 * PANEL_INTERNAL, 1), drop_h=0)
    else:
        r = route(q['panel'], xy, q['end'], pool=pool_for(q['end']))
    q['L'] = r['L']; q['path'] = r['path']; q['drop_h'] = r.get('drop_h', 0)
    ROUTED_CTRL.append((q['cable'], r['L'], r['path'], q))

# cáp mạng / liên kết ngoài I/O
EXTRA_CTRL = []
chain = ['DB.HVAC', 'DDC-FAN', 'DDC-AHU-2.02', 'DDC-AHU-2.01', 'DDC-AHU-2.03']
for a, b in zip(chain[:-1], chain[1:]):
    rr = dijkstra('P_' + a, 'P_' + b)
    th = sum(elen(EDGES[i]) for i in rr); vz = sum(EDGES[i]['vz'] for i in rr)
    EXTRA_CTRL.append(dict(cable=C_485, desc='Mạng BACnet MS/TP/Modbus %s → %s (đồng hồ điện + DDC)' % (a, b), L=r1(th * SLACK + vz + 2 * PANEL_INTERNAL, 1), path=rr))
for a, mods in CDU_OF.items():
    for m in mods:
        r = route_pts(XY[a + 'c'], 'duct', XY[m], 'cdu', pool2=OUT_POOL)
        EXTRA_CTRL.append(dict(cable=C_DX, desc='Truyền tín hiệu DX-kit %s ↔ %s' % (a, m), L=r['L'], path=r['path'], end='cdu',
                               drop_h=(abs(N['Y0'][0] - XY[m][0])) / 1000.0))
for x in EXTRA_CTRL:
    ROUTED_CTRL.append((x['cable'], x['L'], x['path'], x))

def io_counts(panel):
    c = collections.Counter(q['type'] for q in IO if q['panel'] == panel)
    return {k: c.get(k, 0) for k in ('DI', 'DO', 'AI', 'AO')}
IO_SUM = collections.OrderedDict()
for pn in ['DDC-AHU-2.01', 'DDC-AHU-2.02', 'DDC-AHU-2.03', 'DDC-FAN']:
    c = io_counts(pn)
    IO_SUM[pn] = dict(c, total=sum(c.values()), **{k + '_spare': int(math.ceil(v * (1 + SPARE))) for k, v in c.items()})
IO_TOT = {k: sum(IO_SUM[p][k] for p in IO_SUM) for k in ('DI', 'DO', 'AI', 'AO', 'total')}

SENSORS = collections.OrderedDict()
def sens(name, spec, n):
    SENSORS[name] = dict(spec=spec, qty=SENSORS.get(name, {}).get('qty', 0) + n)
sens('Cảm biến nhiệt độ/độ ẩm phòng gắn tường', '0–50 °C ±0,3 K; 0–100 %RH ±2 %; ra 0–10 V; 24 VAC/DC; vỏ chịu hóa chất cho phòng hóa',
     sum(len(v) for v in ROOM_SENS.values()))
sens('Cảm biến nhiệt độ/độ ẩm ống gió', 'Que 200 mm, ±0,3 K / ±2 %RH, 0–10 V (gió cấp + gió hồi mỗi AHU)', 6)
sens('Cảm biến nhiệt độ/độ ẩm ngoài trời', 'IP65, có hộp che nắng, 0–10 V', 1)
sens('Bộ chuyển đổi chênh áp quạt AHU (đo lưu lượng k-factor)', '0–1000 Pa, 0–10 V, kèm ống mềm & đầu đo', 3)
sens('Cảm biến áp suất tĩnh ống gió bù FAF', '0–500 Pa, 0–10 V, đầu đo tĩnh áp', 2)
sens('Công tắc chênh áp lọc', '50–500 Pa, tiếp điểm SPDT (G4 + F8 mỗi AHU/FAF)', 10)
sens('Công tắc dòng gió (airflow switch)', '20–300 Pa, SPDT (3 AHU – liên động sưởi; 3 EAF – xác nhận hút)', 6)
sens('Thermostat an toàn sưởi reset tay', 'Kèm cụm sưởi AHU – WP4 chỉ đấu nối (không tính tiền nếu AHU đã kèm)', 3)
sens('Bộ chấp hành van gió on/off lò xo hồi 24 VAC', '≥ 10 Nm, thời gian chạy ≤ 90 s; van gió bù 15 hood + 7 nhóm EAF + 3 OA AHU + 2 OA FAF (thân van thuộc WP3)', 27)
sens('Hộp công tắc tại tủ hood', 'Công tắc ON/OFF + đèn báo chạy, 24 VDC, IP54, gắn trên tủ hood', 15)
sens('Rơ le dòng (current sensing relay) quạt tủ hood', 'Lõi xuyên, 0,5–5 A, tiếp điểm khô (lắp trong DDC-FAN)', 15)

# ==============================================================================================
# 9. MÁNG CÁP – TẢI TRÊN MỖI ĐOẠN, CHỌN KÍCH THƯỚC
# ==============================================================================================
seg_pw = collections.defaultdict(list); seg_ct = collections.defaultdict(list)
for d in LOADS:
    if d['cable']:
        for i in d['route']['path']:
            seg_pw[i].append(d['OD'])
for f in FEED:
    for i in f['route']['path']:
        seg_pw[i].append(f['OD'])
for cab, L, path, q in ROUTED_CTRL:
    for i in path:
        seg_ct[i].append(CABLE_OD[cab])

def pick_tray(ods):
    a = sum(math.pi * o * o / 4 for o in ods)
    for w, h in TRAYS:
        if a <= TRAY_FILL * w * h:
            return (w, h), a
    return TRAYS[-1], a

SEGS = []
for i, e in enumerate(EDGES):
    L = elen(e) + e['vz']
    pw = seg_pw.get(i, []); ct = seg_ct.get(i, [])
    rec = dict(i=i, a=e['a'], b=e['b'], name=e['name'], env=e['env'], xa=N[e['a']][0], ya=N[e['a']][1],
               xb=N[e['b']][0], yb=N[e['b']][1], vz=e['vz'], L=r1(L, 2), n_pw=len(pw), n_ct=len(ct))
    if pw:
        t, a = pick_tray(pw)
        t = max(t, (200, 100)) if e['name'].startswith('Máng chính') or e['env'] in ('out', 'riser') else t
        rec.update(pw_tray=t, pw_area=r1(a, 0), pw_fill=r1(a / (t[0] * t[1]) * 100, 1))
    if ct:
        t, a = pick_tray(ct)
        rec.update(ct_tray=t, ct_area=r1(a, 0), ct_fill=r1(a / (t[0] * t[1]) * 100, 1))
    SEGS.append(rec)

def tray_desc(t, env):
    w, h = t
    if env in ('out', 'wall'):
        return 'Máng cáp đục lỗ mạ kẽm nhúng nóng %d×%d có nắp (ngoài trời)' % (w, h)
    return 'Máng cáp đục lỗ sơn tĩnh điện %d×%d dày 1,2 mm' % (w, h)

TRAY_Q = collections.OrderedDict()
for s in SEGS:
    if s['env'] == 'wall':
        continue
    for k in ('pw_tray', 'ct_tray'):
        if k in s:
            dsc = tray_desc(s[k], s['env'])
            TRAY_Q[dsc] = TRAY_Q.get(dsc, 0) + s['L']
# phụ kiện: co, tê, nối theo bậc nút
deg = collections.Counter()
for s in SEGS:
    if s['env'] != 'wall':
        deg[s['a']] += 1; deg[s['b']] += 1
def turn_at(n):
    vs = []
    for s in SEGS:
        if s['env'] == 'wall':
            continue
        if s['a'] == n or s['b'] == n:
            o = s['b'] if s['a'] == n else s['a']
            vs.append('x' if N[o][1] == N[n][1] else 'y')
    return vs
FIT = collections.Counter()
for n, k in deg.items():
    if n.startswith('P_'):
        FIT['Co đứng 90° (máng đứng nóc tủ)'] += 2 if n == 'P_DB.HVAC' or n == 'P_DDC-FAN' else 1
        continue
    if k == 1:
        FIT['Nắp bịt đầu máng'] += 1
    elif k == 2:
        vs = turn_at(n)
        if vs[0] != vs[1]:
            FIT['Co ngang 90°'] += 1
    elif k == 3:
        FIT['Tê ngang (T)'] += 1
    elif k >= 4:
        FIT['Chữ thập (+)'] += 1
L_tray_total = sum(TRAY_Q.values())
n_ct_pw = 2 if any('ct_tray' in s for s in SEGS) else 1
FIT['Nối máng (khớp nối + bu lông), mỗi 2,4 m'] = int(math.ceil(L_tray_total / 2.4))
L_in = sum(v for k, v in TRAY_Q.items() if 'ngoài trời' not in k)
L_out = sum(v for k, v in TRAY_Q.items() if 'ngoài trời' in k)
FIT['Giá treo trong nhà (ty M10 ×2 + U-channel 41×41, mỗi 1,5 m)'] = int(math.ceil(L_in / 1.5))
FIT['Giá đỡ ngoài trời ốp tường/bệ (mạ kẽm nhúng nóng, mỗi 1,5 m)'] = int(math.ceil(L_out / 1.5))
FIT['Ống lồng xuyên tường A Ø100 thép mạ kẽm + chèn chống cháy/chống thấm'] = sum(1 for s in SEGS if s['env'] == 'wall') * 2

# ==============================================================================================
# 10. ỐNG LUỒN, ỐNG MỀM, ỐC SIẾT CÁP, CÁCH LY TẠI CHỖ, TIẾP ĐỊA
# ==============================================================================================
def conduit_for(od):
    for dn, idm in CONDUIT:
        if od <= idm * math.sqrt(0.53):
            return dn
    return 50
def gland_for(od):
    for g, mx in GLAND:
        if od <= mx:
            return g
    return 'M63'
CONDQ = collections.Counter(); FLEXQ = collections.Counter(); GLQ = collections.Counter(); ISOQ = collections.Counter()
for d in LOADS:
    if not d['cable']:
        continue
    r = d['route']
    dn = conduit_for(d['OD'])
    h = r['drop_h'] * SLACK + max(0.0, END_VZ[d['end']] - FLEX.get(d['end'], 0))
    if d['end'] in ('cdu', 'scr'):
        CONDQ['Ống thép mạ kẽm (GI/IMC) ren Ø%d ngoài trời' % dn] += h
    else:
        CONDQ['Ống thép EMT Ø%d (đoạn xuống thiết bị trên trần)' % dn] += h
    if d['end'] in FLEX and d['kind'] not in ('SOCKET',):
        FLEXQ['Ống mềm kim loại bọc PVC kín nước Ø%d + đầu nối' % dn] += FLEX[d['end']]
    gtyp = 'đồng mạ niken EMC' if d['shield'] else ('đồng mạ niken IP68' if d['end'] in ('cdu', 'scr') else 'nhựa nylon IP68')
    GLQ['Ốc siết cáp %s %s' % (gtyp, gland_for(d['OD']))] += 2 + (2 if d['iso'] else 0)
    if d['iso']:
        if d['ph'] == 1:
            nm = 'Cầu dao cách ly tại chỗ 2P %dA IP65' % d['iso_A']
        else:
            pol = '4P' if d['neutral'] else '3P'
            nm = 'Cầu dao cách ly tại chỗ %s %dA IP65%s%s' % (pol, d['iso_A'], ' ngoài trời (chống UV)' if d['end'] in ('cdu', 'scr') else '',
                                                             ' + tiếp điểm phụ (báo biến tần)' if d['kind'] == 'VFD' else '')
        ISOQ[nm] += 1
for f in FEED:
    GLQ['Ốc siết cáp nhựa nylon IP68 %s' % gland_for(f['OD'])] += 2
for cab, L, path, q in ROUTED_CTRL:
    od = CABLE_OD[cab]
    dh = q.get('drop_h', 0) if isinstance(q, dict) else 0
    end = q.get('end', 'ceil') if isinstance(q, dict) else 'ceil'
    if end in ('cdu', 'scr'):
        CONDQ['Ống thép mạ kẽm (GI/IMC) ren Ø20 ngoài trời'] += dh * SLACK + END_VZ[end] * 0.5
    elif end != 'panel':
        CONDQ['Ống nhựa PVC cứng Ø20 (cáp điều khiển/cảm biến)'] += dh * SLACK + (END_VZ.get(end, 1.0) if end == 'wall' else 0.5)
    GLQ['Ốc siết cáp nhựa nylon IP68 M20'] += 2
CONDQ = {k: r1(v, 0) for k, v in CONDQ.items() if v >= 0.5}
# tiếp địa
EARTH = collections.OrderedDict()
L_pw_tray = sum(s['L'] for s in SEGS if 'pw_tray' in s)
EARTH['Dây tiếp địa dọc máng cáp động lực CV 1×25 mm² vàng-xanh'] = ('m', r1(L_pw_tray * 1.05, 0))
EARTH['Dây nối đẳng thế mối nối máng CV 1×6 mm² + đầu cos'] = ('bộ', FIT['Nối máng (khớp nối + bu lông), mỗi 2,4 m'])
EARTH['Dây nối đất vỏ CDU/scrubber/khung bệ CV 1×16 mm² (3 m/thiết bị)'] = ('m', 3.0 * (len(CDU_MODULES) + 1))
EARTH['Đấu nối thanh PE DB.HVAC → thanh tiếp địa chính tòa nhà (MET) CV 1×95 mm² (giả định 10 m)'] = ('m', 10.0)
EARTH['Kẹp tiếp địa/đầu cos đồng các loại'] = ('bộ', 60)

# ==============================================================================================
# 11. THÀNH PHẦN TỦ
# ==============================================================================================
def comp(panel):
    L = [d for d in LOADS if d['src'] == panel]
    out = []
    if panel == 'DB.HVAC':
        out.append(('Vỏ tủ thép 1,5–2,0 mm sơn tĩnh điện, %s, %s, Form 2b, cửa có khóa' % (PANELS[panel]['dims'], PANELS[panel]['ip']), 'bộ', 1))
        out.append((INCOMING['MCCB'], 'cái', 1))
        out.append((INCOMING['busbar'], 'bộ', 1))
        out.append(('Đồng hồ đa năng V/A/kW/kWh/PF/THD, Modbus RTU + 3 CT %d/5 A' % In_db, 'bộ', 1))
        out.append(('Thiết bị chống sét lan truyền SPD Type 2, 4P, Imax 40 kA + MCB 4P 32 A', 'bộ', 1))
        out.append(('Đèn báo pha + cầu chì bảo vệ', 'bộ', 3))
        mc = collections.Counter(d['prot'] for d in L)
        for f in FEED:
            mc[f['prot']] += 1
        for k, v in sorted(mc.items()):
            out.append((k, 'cái', v))
        out.append(('Rơ le dòng rò ELR 300 mA + ZCT (cho lộ CDU, theo hãng CDU)', 'bộ', len(CDU_MODULES)))
        out.append(('Lộ dự phòng 20 %%: %s' % SPARES[panel]['spec'], 'cái', SPARES[panel]['n_spare']))
        out.append(('Cầu đấu, máng nhựa đi dây, nhãn, sơ đồ trong tủ, phụ kiện', 'lô', 1))
        return out
    f = FEED_BY[panel]
    out.append(('Vỏ tủ thép sơn tĩnh điện %s, %s, quạt hút + lọc gió mặt tủ' % (PANELS[panel]['dims'], PANELS[panel]['ip']), 'bộ', 1))
    out.append(('MCCB tổng 4P %d A 25 kA (tủ) / cầu dao tổng' % f['In'], 'cái', 1))
    out.append(('Thanh cái Cu %d A 3P+N+PE' % pick([63, 100, 160, 250], f['In']), 'bộ', 1))
    vf = [d for d in L if d['kind'] == 'VFD']
    for d in vf:
        out.append(('Biến tần %s kW 380 V (HVAC, cuộn kháng DC/EMC C2, Modbus/BACnet) cho %s' % (vn(d['P'], 1), d['tag']), 'bộ', 1))
    mc = collections.Counter(d['prot'] for d in L if d['kind'] not in ('CTRL',))
    for k, v in sorted(mc.items()):
        out.append((k, 'cái', v))
    ht = [d for d in L if d['kind'] == 'HEATER']
    if ht:
        out.append(('RCD 4P 40 A 300 mA cho nhóm sưởi', 'cái', 1))
        out.append(('Bộ điều khiển thyristor/SSR 3P (0–10 V) cho cấp sưởi 1', 'bộ', 1))
        out.append(('Contactor AC1 3P cho cấp sưởi 2, 3', 'cái', len(ht) - 1))
        out.append(('Rơ le an toàn liên động sưởi (quạt chạy + airflow + thermostat) + rơ le thời gian chạy quạt sau 3 phút', 'bộ', 1))
    if panel == 'DDC-FAN':
        nd = len([d for d in L if d['kind'] == 'DOL'])
        out.append(('Contactor AC3 3P 9 A cuộn 24 VAC (quạt EAF + quạt tủ hood)', 'cái', nd))
        out.append(('Rơ le dòng lõi xuyên cho quạt tủ hood', 'cái', 15))
        out.append(('Biến áp 220/24 VAC 630 VA cấp bộ chấp hành van gió + MCB bảo vệ', 'bộ', 1))
        out.append(('Rơ le trung gian 24 VAC (lệnh van gió)', 'cái', 22 + 2))
    else:
        out.append(('Biến áp 220/24 VAC 100 VA + nguồn 24 VDC 5 A cho DDC/cảm biến', 'bộ', 1))
        out.append(('Rơ le cắt cứng: báo cháy (STO/enable biến tần + sưởi), mức nước ngưng cao (cắt DX-kit)', 'bộ', 1))
    c = IO_SUM[panel]
    out.append(('Bộ điều khiển DDC lập trình tự do BACnet + module mở rộng: ≥ %d DI / %d DO / %d AI / %d AO (đã gồm 20 %% dự phòng)' %
                (c['DI_spare'], c['DO_spare'], c['AI_spare'], c['AO_spare']), 'bộ', 1))
    out.append(('Màn hình hiển thị/cài đặt cục bộ trên cửa tủ (LCD/HMI 7")', 'bộ', 1))
    out.append(('Công tắc H-O-A, nút dừng khẩn, đèn báo chạy/lỗi cho từng thiết bị', 'lô', 1))
    out.append(('Lộ dự phòng 20 %%: %s' % SPARES[panel]['spec'], 'bộ', SPARES[panel]['n_spare']))
    out.append(('Cầu đấu, máng nhựa đi dây, nhãn, phụ kiện', 'lô', 1))
    return out
PANEL_BOM = {p: comp(p) for p in PANELS}

OPTIONS = [
    ('OP-1', 'Bộ điều khiển mạng (network controller/BACnet router) BACnet MS/TP ↔ BACnet/IP, lắp trong DDC-FAN', 'bộ', 1),
    ('OP-2', 'Màn hình HMI cảm ứng 15,6" (đồ họa 3 AHU, FAF/EAF, hood, scrubber, báo động, xu hướng) đặt tại Office hoặc phòng điều khiển Phase 1', 'bộ', 1),
    ('OP-3', 'Phần mềm + lập trình đồ họa, tích hợp vào HMI Phase 1 ("MÀN HÌNH HMI" – F3-SD-AC-017) qua BACnet/IP hoặc Modbus TCP', 'lô', 1),
    ('OP-4', 'Gateway BACnet/Modbus cho CDU (giám sát chi tiết lỗi/công suất từ hãng CDU)', 'bộ', 3),
    ('OP-5', 'Cáp mạng Cat6 FTP (DDC-FAN → HMI Phase 1, giả định 80 m) + ống PVC Ø20 + switch công nghiệp 8 cổng', 'lô', 1),
]

# ==============================================================================================
# 12. TỔNG HỢP KHỐI LƯỢNG (BOQ)
# ==============================================================================================
PWQ = collections.OrderedDict()
for d in LOADS:
    if d['cable']:
        PWQ[d['cable']] = PWQ.get(d['cable'], 0) + d['route']['L']
for f in FEED:
    PWQ[f['cable']] = PWQ.get(f['cable'], 0) + f['route']['L']
CTQ = collections.OrderedDict()
for cab, L, path, q in ROUTED_CTRL:
    CTQ[cab] = CTQ.get(cab, 0) + L

BOQ = []
def boq(group, item, spec, unit, qty, note=''):
    BOQ.append(dict(no=len(BOQ) + 1, group=group, item=item, spec=spec, unit=unit, qty=qty, note=note))
G1 = 'A. Tủ điện & điều khiển'
for p in PANELS:
    sub = '; '.join('%s ×%s' % (c[0], c[2]) for c in PANEL_BOM[p][:6])
    boq(G1, 'Tủ %s (trọn bộ, lắp đặt, đấu nối, thử nghiệm)' % p, '%s; %s. Thành phần chính: xem electrical_design.json → panels.%s.components' % (PANELS[p]['dims'], PANELS[p]['mount'], p), 'tủ', 1, PANELS[p]['wall'])
G2 = 'B. Cáp động lực Cu/XLPE/PVC 0,6/1 kV'
def skey(k):
    import re
    m = re.search(r'(\d+)×([\d,]+)', k)
    return (('CXV/S' in k), int(m.group(1)), float(m.group(2).replace(',', '.')))
for k in sorted(PWQ, key=skey):
    boq(G2, 'Cáp ' + k, 'Cu/XLPE/PVC 0,6/1 kV, TCVN 5935 / IEC 60502-1; đi trên máng + ống', 'm', int(math.ceil(PWQ[k])), 'Đã gồm +15 % chiều dài ngang + đoạn đứng')
G3 = 'C. Cáp điều khiển, tín hiệu, mạng'
for k, v in CTQ.items():
    boq(G3, k, 'Cu, cách điện PVC/PE, màn chắn lá nhôm + dây thoát (trừ cáp FR: IEC 60331, 90 phút)', 'm', int(math.ceil(v)))
G4 = 'D. Máng cáp & phụ kiện'
for k, v in TRAY_Q.items():
    boq(G4, k, 'BOP +3,6 m (như Phase 1); kèm phụ kiện; độ đầy ≤ 40 %', 'm', int(math.ceil(v)))
for k, v in FIT.items():
    boq(G4, k, 'Theo kích thước máng tương ứng', 'cái' if 'Giá' not in k else 'bộ', int(v))
G5 = 'E. Ống luồn, ống mềm, ốc siết cáp'
for k, v in sorted(CONDQ.items()):
    boq(G5, k, 'Kèm khớp nối, kẹp, hộp kéo dây', 'm', int(math.ceil(v)))
for k, v in sorted(FLEXQ.items()):
    boq(G5, k, 'Đoạn cuối ~1 m vào động cơ/cụm sưởi/CDU (chống rung)', 'm', int(math.ceil(v)))
for k, v in sorted(GLQ.items()):
    boq(G5, k, 'Đầu vào tủ/hộp cách ly/thiết bị', 'cái', int(v))
G6 = 'F. Cầu dao cách ly tại chỗ, ổ cắm'
for k, v in sorted(ISOQ.items()):
    boq(G6, k, 'Vỏ polycarbonate/kim loại, tay khóa được (lock-off), IEC 60947-3 AC-23A', 'cái', v)
boq(G6, 'Ổ cắm công nghiệp 1P+N+E 16 A IP66 tại sàn AHU', 'IEC 60309', 'cái', 3)
G7 = 'G. Cảm biến & thiết bị hiện trường'
for k, v in SENSORS.items():
    boq(G7, k, v['spec'], 'cái', v['qty'])
G8 = 'H. Tiếp địa'
for k, (u, v) in EARTH.items():
    boq(G8, k, 'Cu, IEC 60364-5-54', u, v)
G9 = 'I. Lập trình, thử nghiệm, nghiệm thu'
boq(G9, 'Lập trình DDC, chạy thử liên động hút–bù, cân chỉnh biến tần, T&C, hồ sơ hoàn công', '4 tủ DDC, 27 van gió, 15 hood, 3 AHU, 6 CDU', 'lô', 1)
boq(G9, 'Đo điện trở cách điện, thông mạch PE, thử RCD, đo sụt áp/dòng vận hành', 'TCVN 9208 / IEC 60364-6', 'lô', 1)
G10 = 'J. PHƯƠNG ÁN CHỌN THÊM (không cộng vào tổng)'
for o in OPTIONS:
    boq(G10, '[%s] %s' % (o[0], o[1]), 'Tùy chọn kết nối BMS/HMI Phase 1', o[2], o[3], 'OPTION')

# ==============================================================================================
# 13. TRÌNH TỰ ĐIỀU KHIỂN (tóm tắt)
# ==============================================================================================
SEQUENCES = [
    ('AHU – khởi động', 'Lệnh chạy (lịch/BMS/H-O-A) → mở van OA → trễ 60 s → biến tần chạy, giữ lưu lượng gió cấp theo chênh áp quạt (k-factor, PID). AHU-2.03 chạy 24/7. Quạt mất xác nhận (biến tần/airflow) → báo lỗi, khóa sưởi và CDU.'),
    ('AHU – làm lạnh', 'Điều khiển theo nhiệt độ hồi (giới hạn bởi giá trị cao nhất của cảm biến phòng trong vùng): yêu cầu công suất 0–10 V tới DX-kit, phân cấp module CDU (lead/lag, luân phiên 24 h, chống chu kỳ ngắn 5 phút).'),
    ('AHU – khử ẩm', 'RH phòng > 60 % (Instrument > 55 %) → ưu tiên khử ẩm: CDU chạy để T sau coil ≈ 15 °C; nhiệt độ phòng giữ bằng sưởi gia nhiệt lại (SSR điều biến + 2 cấp contactor).'),
    ('AHU – sưởi (đông)', 'T phòng < 20 °C → sưởi theo PID (SSR cấp 1 điều biến, cấp 2–3 đóng tuần tự). Liên động cứng: chỉ cấp điện sưởi khi biến tần chạy + airflow switch đóng + thermostat an toàn bình thường; sau khi tắt sưởi quạt chạy thêm 3 phút. Không chạy sưởi và lạnh ở chế độ thường (chỉ đồng thời khi khử ẩm).'),
    ('AHU – lọc, nước ngưng', 'Công tắc ΔP lọc G4/F8 → báo bẩn. Phao mức cao bơm nước ngưng → cắt cứng DX-kit (CDU), báo động; quạt vẫn chạy.'),
    ('Máy hút ẩm DH-2.01', 'DDC-AHU-2.03 cho phép DH chạy khi RH Instrument > 55 % (trễ 5 %), hoặc khi AHU-2.03 ở non tải không giữ được RH; giám sát chạy/lỗi/bồn đầy.'),
    ('Hút tủ hood ↔ gió bù', 'Người dùng bật công tắc tại tủ → DDC mở van gió bù MD-<hood> → trễ 30 s → đóng contactor quạt tủ → rơ le dòng xác nhận → đèn tại tủ. Tắt: dừng quạt, đóng van sau 30 s.'),
    ('Hood acid ↔ scrubber', '4 hood acid (DIG-FH1/FH2, ST1-FH1, ST2-FH3): yêu cầu chạy → DO khởi động scrubber → chỉ đóng contactor quạt hood khi scrubber báo chạy; scrubber chạy thêm 10 phút sau hood acid cuối cùng tắt. Lỗi scrubber → báo động tại tủ + DDC (quạt hood giữ chạy, người dùng hạ cửa tủ).'),
    ('EAF ↔ van bù nhóm', 'EAF-2.01/2.02/2.03 chạy theo lịch (giờ làm việc; EAF-2.03 theo Instrument 24/7) → mở van bù nhóm MD-G-xx tương ứng; xác nhận dòng gió bằng công tắc ΔP. Instrument: thải bù bằng gió tươi AHU-2.03 (van OA mở khi EAF-2.03 chạy).'),
    ('FAF (gió bù)', 'FAF-2.01/2.02 được phép chạy khi có ≥ 1 van bù mở; biến tần PID giữ áp suất tĩnh ống gió bù (đặt khi T&C, ~250 Pa) nên lưu lượng tự theo số van mở; van OA cửa hút mở trước khi quạt chạy. Lỗi FAF → chỉ báo động (không dừng quạt hút).'),
    ('Báo cháy', 'Tín hiệu từ module rơ le báo cháy (gói PCCC): dừng 3 AHU (cắt cứng enable/STO biến tần + sưởi), dừng FAF, đóng van OA và van bù. Quạt hút EAF/hood/scrubber GIỮ chạy để không thoát hơi hóa chất ra phòng – cần tư vấn PCCC/CĐT xác nhận.'),
    ('Chế độ mùa', 'Cảm biến T/RH ngoài trời (DDC-AHU-2.02, chia sẻ qua mạng) chọn chế độ Hè/Đông/Nồm; mạng RS485 BACnet MS/TP nối 4 tủ DDC + đồng hồ DB.HVAC.'),
]

# ==============================================================================================
# 14. XUẤT JSON
# ==============================================================================================
def circ_json(d):
    r = d.get('route') or {}
    return dict(id=d['id'], from_panel=d['src'], load=d['desc'], tag=d['tag'], kind=d['kind'], phases=d['ph'], U_V=d['U'],
                P_rated_kW=d['P'], eta=r1(d['eta'] * d['eta_vfd'], 3), pf=d['pf'], P_input_kW=r1(d['Pin'], 2), Ib_A=r1(d['Ib'], 2),
                I_cable_A=r1(d['Icab'], 2), demand_factors=dict(zip(['he', 'dong', 'nom'], d['df'])), protection=d['prot'],
                In_A=d['In'], rcd=d['rcd'], cable=d.get('cable') if d['cable'] else None, S_mm2=d.get('S'), PE_mm2=d.get('PE'),
                Iz_A=r1(d.get('Iz', 0), 1), L_horizontal_m=r.get('L_h'), L_vertical_m=r.get('L_v'), L_m=r.get('L'),
                dU_pct=r1(d['du'], 2), dU_cum_pct=r1(d['du_cum'], 2), local_isolator_A=d.get('iso_A'), note=d['note'])
design = dict(
    project=EQS['project'], wp='WP4 – Điện & điều khiển', revision='v1.0', date='2026-11-14',
    scope='Từ DB.HVAC (gồm DB.HVAC) trở đi: tủ DDC, cáp lộ ra, máng/ống, cáp điều khiển, cảm biến. Nguồn cấp tới DB.HVAC thuộc gói điện (ngoài phạm vi).',
    inputs=['01_design_basis/DESIGN_BASIS.md v1.1', '02_calc/equipment_schedule.json (WP1 v4)', '03_geometry/geometry.json (WP2)'],
    parameters=dict(U3=U3, U1=U1, pf_std=PF_STD, pf_vfd=PF_VFD, pf_heater=PF_RES, eta_vfd=ETA_VFD, eta_motor_IE3=ETA_IE3,
                    k_temp_45C=K_TEMP, k_group=K_GROUP, In_over_Ib=K_IN, slack=SLACK, rho=RHO, X_ohm_per_m=XL,
                    dU_max_pct=DU_MAX, dU_feeder_target_pct=DU_FEEDER, tray_BOP_mm=TRAY_Z, spare=SPARE,
                    ampacity_table='IEC 60364-5-52 B.52.12 XLPE method E (≡ TCVN 9207:2012)'),
    panels={p: dict(PANELS[p], xy_mm=list(PANELS[p]['xy']), components=[dict(item=c[0], unit=c[1], qty=c[2]) for c in PANEL_BOM[p]],
                    spare_ways=SPARES.get(p), io=IO_SUM.get(p)) for p in PANELS},
    db_hvac=INCOMING,
    feeders=[dict(id=f['id'], to=f['to'], P_input_kW=r1(f['Pin'], 2), Ib_A=r1(f['Ib'], 1), protection=f['prot'], In_A=f['In'],
                  cable=f['cable'], S_mm2=f['S'], Iz_A=r1(f['Iz'], 1), L_m=f['route']['L'], dU_pct=r1(f['du'], 2)) for f in FEED],
    circuits=[circ_json(d) for d in LOADS],
    cdu_feed_decision='CDU cấp trực tiếp từ DB.HVAC (6 lộ MCCB 4P riêng + ELCB), cầu dao cách ly 4P IP65 ngoài trời tại mỗi module; AHU điều khiển CDU bằng tín hiệu DX-kit (không đóng cắt nguồn qua tủ DDC).',
    io_schedule=[dict(panel=q['panel'], type=q['type'], point=q['desc'], device=q['device'], location=q['loc'],
                      internal=q['internal'], cable=q['cable'] or None, L_m=q['L']) for q in IO],
    io_summary=IO_SUM, io_total=IO_TOT, extra_control_cables=[dict(cable=x['cable'], desc=x['desc'], L_m=x['L']) for x in EXTRA_CTRL],
    sensors=SENSORS, group_make_up_dampers={k: dict(linked_fan=v[0], points=v[1], xy_mm=[r1(XY[k][0], 0), r1(XY[k][1], 0)]) for k, v in GROUP_DAMPERS.items()},
    control_sequences=[dict(topic=a, text=b) for a, b in SEQUENCES],
    containment=dict(nodes={k: list(v) for k, v in N.items()}, segments=[{k: (list(v) if isinstance(v, tuple) else v) for k, v in s.items()} for s in SEGS],
                     tray_m=TRAY_Q, fittings=dict(FIT), conduits_m=CONDQ, flex_m=dict(FLEXQ), glands=dict(GLQ)),
    isolators=dict(ISOQ), earthing={k: dict(unit=u, qty=v) for k, (u, v) in EARTH.items()},
    cable_totals=dict(power_m={k: r1(v, 1) for k, v in PWQ.items()}, control_m={k: r1(v, 1) for k, v in CTQ.items()}),
    options=[dict(code=o[0], item=o[1], unit=o[2], qty=o[3]) for o in OPTIONS],
)
json.dump(design, open(os.path.join(HERE, 'electrical_design.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump(dict(project=EQS['project'], wp='WP4', revision='v1.0', source='05_electrical/build_electrical.py',
               note='Khối lượng đã gồm +15 % chiều dài ngang & đoạn đứng; mục nhóm J là phương án chọn thêm (không cộng). Thân van gió động cơ thuộc WP3; thiết bị (AHU, CDU, quạt, scrubber, DH, bơm) thuộc WP5.',
               lines=BOQ), open(os.path.join(HERE, 'takeoff_electrical.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

# ==============================================================================================
# 15. EXCEL (công thức) + tính lại LibreOffice + kiểm tra
# ==============================================================================================
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
wb = Workbook()
HDR = PatternFill('solid', fgColor='DDEBF7'); BOLD = Font(bold=True); THIN = Side(style='thin', color='999999')
def header(ws, row, cols, widths=None):
    for j, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=j, value=c); x.font = BOLD; x.fill = HDR
        x.alignment = Alignment(wrap_text=True, vertical='center')
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[CL(j)].width = w
def title(ws, t, sub=''):
    ws['A1'] = t; ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = sub

# --- Bảng tra ---
wt = wb.active; wt.title = 'Bảng tra'
title(wt, 'Bảng tra – khả năng mang tải cáp Cu/XLPE nhiều lõi trên khay đục lỗ (IEC 60364-5-52 B.52.12, PP E, 30 °C) ≡ TCVN 9207:2012',
      'Hệ số: nhiệt độ 45 °C (B.52.14) và nhóm (chỉ đạo WP4) dùng chung cho toàn bộ cáp')
header(wt, 4, ['S (mm²)', 'Iz0 3 dây mang tải (A)', 'Iz0 2 dây mang tải (A)', 'OD cáp 4 lõi (mm)'], [12, 22, 22, 18])
for i, S in enumerate(SIZES):
    wt.cell(row=5 + i, column=1, value=S); wt.cell(row=5 + i, column=2, value=IZ3[S]); wt.cell(row=5 + i, column=3, value=IZ2[S]); wt.cell(row=5 + i, column=4, value=OD4[S])
TB_END = 4 + len(SIZES)
prm = [('k nhiệt độ 45 °C', K_TEMP), ('k nhóm', K_GROUP), ('Hệ số In/Ib', K_IN), ('Hệ số dư chiều dài ngang', SLACK),
       ('ρ Cu (Ω·mm²/m)', RHO), ('X (Ω/m)', XL), ('ΔU cho phép (%)', DU_MAX), ('U 3 pha (V)', U3), ('U 1 pha (V)', U1), ('η biến tần', ETA_VFD)]
wt['F4'] = 'Tham số'; wt['G4'] = 'Giá trị'; wt['F4'].font = BOLD; wt['G4'].font = BOLD
wt.column_dimensions['F'].width = 30
for i, (k, v) in enumerate(prm):
    wt.cell(row=5 + i, column=6, value=k); wt.cell(row=5 + i, column=7, value=v)
P_KT, P_KG, P_KIN, P_SL, P_RHO, P_X, P_DU, P_U3, P_U1, P_EV = ["'Bảng tra'!$G$%d" % (5 + i) for i in range(10)]
TAB = "'Bảng tra'!$A$5:$C$%d" % TB_END

# --- Phụ tải ---
wp = wb.create_sheet('Phụ tải')
title(wp, 'Bảng phụ tải HVAC – DB.HVAC và các tủ DDC (P đặt theo WP1; P vào = P/(η động cơ·η biến tần))',
      'Hệ số đồng thời theo 3 kịch bản: Hè / Đông / Nồm (sưởi và lạnh không đồng thời trừ chế độ khử ẩm)')
cols = ['Mã mạch', 'Tủ nguồn', 'Phụ tải', 'Tag', 'Loại', 'Số pha', 'U (V)', 'P đặt (kW)', 'η động cơ', 'η biến tần', 'cosφ lưới',
        'P vào (kW)', 'Ib tính (A)', 'FLA WP1 (A)', 'Ib thiết kế (A)', 'cosφ động cơ', 'I tính cáp (A)',
        'Kđt Hè', 'Kđt Đông', 'Kđt Nồm', 'P Hè (kW)', 'P Đông (kW)', 'P Nồm (kW)', 'S Hè (kVA)', 'S Đông (kVA)', 'S Nồm (kVA)']
header(wp, 4, cols, [16, 14, 52, 13, 9, 7, 7, 9, 9, 9, 9, 10, 10, 10, 11, 10, 10, 8, 8, 8, 10, 10, 10, 10, 10, 10])
ROW_OF = {}
r0 = 5
for i, d in enumerate(LOADS):
    r = r0 + i; ROW_OF[d['id']] = r
    vals = [d['id'], d['src'], d['desc'], d['tag'], d['kind'], d['ph'], d['U'], d['P'], d['eta'], d['eta_vfd'], d['pf']]
    for j, v in enumerate(vals, 1):
        wp.cell(row=r, column=j, value=v)
    wp.cell(row=r, column=12, value='=H%d/(I%d*J%d)' % (r, r, r))
    wp.cell(row=r, column=13, value='=IF(F%d=3,L%d*1000/(SQRT(3)*G%d*K%d),L%d*1000/(G%d*K%d))' % (r, r, r, r, r, r, r))
    wp.cell(row=r, column=14, value=d['fla'] if d['fla'] else 0)
    wp.cell(row=r, column=15, value='=MAX(M%d,N%d)' % (r, r))
    wp.cell(row=r, column=16, value=d['pf_m'] if d['kind'] == 'VFD' else d['pf'])
    wp.cell(row=r, column=17, value='=IF(E%d="VFD",H%d/I%d*1000/(SQRT(3)*G%d*P%d),O%d)' % (r, r, r, r, r, r))
    for k in range(3):
        wp.cell(row=r, column=18 + k, value=d['df'][k])
        wp.cell(row=r, column=21 + k, value='=L%d*%s%d' % (r, CL(18 + k), r))
        wp.cell(row=r, column=24 + k, value='=L%d*%s%d/K%d' % (r, CL(18 + k), r, r))
rl = r0 + len(LOADS) - 1
rt = rl + 2
wp.cell(row=rt, column=3, value='TỔNG DB.HVAC (đặt / theo kịch bản)').font = BOLD
wp.cell(row=rt, column=8, value='=SUM(H%d:H%d)' % (r0, rl))
wp.cell(row=rt, column=12, value='=SUM(L%d:L%d)' % (r0, rl))
for c in range(21, 27):
    wp.cell(row=rt, column=c, value='=SUM(%s%d:%s%d)' % (CL(c), r0, CL(c), rl))
wp.cell(row=rt + 1, column=3, value='Dòng theo kịch bản I = S/(√3·U) (A)')
for c in range(24, 27):
    wp.cell(row=rt + 1, column=c, value='=%s%d*1000/(SQRT(3)*%s)' % (CL(c), rt, P_U3))
wp.cell(row=rt + 2, column=3, value='Nhu cầu lớn nhất (kVA) / dòng Ib max (A)').font = BOLD
wp.cell(row=rt + 2, column=24, value='=MAX(X%d:Z%d)' % (rt, rt))
wp.cell(row=rt + 2, column=25, value='=X%d*1000/(SQRT(3)*%s)' % (rt + 2, P_U3))
wp.cell(row=rt + 3, column=3, value='Tổng S đặt (kVA) / dòng đặt (A)')
wp.cell(row=rt + 3, column=24, value='=SUMPRODUCT(L%d:L%d/K%d:K%d)' % (r0, rl, r0, rl))
wp.cell(row=rt + 3, column=25, value='=X%d*1000/(SQRT(3)*%s)' % (rt + 3, P_U3))
wp.cell(row=rt + 4, column=3, value='Hệ số nhu cầu chung = P max kịch bản / P vào đặt')
wp.cell(row=rt + 4, column=24, value='=MAX(U%d:W%d)/L%d' % (rt, rt, rt))
XL_DEM = (rt + 2, 24); XL_IMAX = (rt + 2, 25)
rf = rt + 7
wp.cell(row=rf - 1, column=1, value='Cáp lộ tổng DB.HVAC → tủ DDC (Ib = Σ Ib thiết kế các mạch trong tủ, cộng số học)').font = BOLD
header(wp, rf, ['Mã lộ', 'Tủ nguồn', 'Đến tủ', '', '', 'Số pha', 'U (V)', 'Σ P đặt (kW)', '', '', 'cosφ', 'Σ P vào (kW)', '', '', 'Ib (A)'])
FROW = {}
for i, f in enumerate(FEED):
    r = rf + 1 + i; FROW[f['to']] = r
    wp.cell(row=r, column=1, value=f['id']); wp.cell(row=r, column=2, value='DB.HVAC'); wp.cell(row=r, column=3, value=f['to'])
    wp.cell(row=r, column=6, value=3); wp.cell(row=r, column=7, value=U3); wp.cell(row=r, column=11, value=f['pf'])
    wp.cell(row=r, column=8, value='=SUMIF($B$%d:$B$%d,C%d,$H$%d:$H$%d)' % (r0, rl, r, r0, rl))
    wp.cell(row=r, column=12, value='=SUMIF($B$%d:$B$%d,C%d,$L$%d:$L$%d)' % (r0, rl, r, r0, rl))
    wp.cell(row=r, column=15, value='=SUMIF($B$%d:$B$%d,C%d,$O$%d:$O$%d)' % (r0, rl, r, r0, rl))
wp.freeze_panes = 'D5'

# --- Sơ đồ nguyên lý (dạng bảng) ---
ws = wb.create_sheet('Sơ đồ nguyên lý')
title(ws, 'Sơ đồ nguyên lý DB.HVAC & tủ DDC (dạng bảng) – xem WP4_single_line.png',
      'Nguồn đến: %s (gói điện cấp, %s). Thanh cái: %s' % (INCOMING['incoming_cable'], 'ngoài phạm vi HVAC', INCOMING['busbar']))
header(ws, 4, ['Tủ', 'Lộ', 'Thiết bị bảo vệ', 'In (A)', 'Bảo vệ dòng rò', 'Cáp', 'Đến', 'Cách ly tại chỗ', 'Ghi chú'], [14, 18, 46, 8, 34, 40, 52, 34, 50])
r = 5
ws.cell(row=r, column=1, value='DB.HVAC'); ws.cell(row=r, column=2, value='Lộ tổng'); ws.cell(row=r, column=3, value=INCOMING['MCCB'])
ws.cell(row=r, column=4, value=In_db); ws.cell(row=r, column=6, value=INCOMING['incoming_cable'] + ' (gói điện)')
ws.cell(row=r, column=7, value='Từ tủ hạ thế tổng (gói điện)'); ws.cell(row=r, column=9, value='Đồng hồ đa năng Modbus, SPD Type 2, đèn pha'); r += 1
for p in ['DB.HVAC'] + FEEDERS:
    rows = [d for d in LOADS if d['src'] == p]
    if p == 'DB.HVAC':
        rows = rows + FEED
    for d in rows:
        ws.cell(row=r, column=1, value=p); ws.cell(row=r, column=2, value=d['id']); ws.cell(row=r, column=3, value=d['prot'])
        ws.cell(row=r, column=4, value=d['In']); ws.cell(row=r, column=5, value=d.get('rcd', ''))
        ws.cell(row=r, column=6, value=d['cable'] if d.get('cable') and d['cable'] is not True else ('(trong tủ)' if d.get('kind') == 'CTRL' else d.get('cable')))
        ws.cell(row=r, column=7, value=d.get('desc'))
        ws.cell(row=r, column=8, value=('%s %d A IP65' % ('Cách ly', d['iso_A'])) if d.get('iso_A') else '')
        ws.cell(row=r, column=9, value=d.get('note', ''))
        r += 1
    ws.cell(row=r, column=1, value=p); ws.cell(row=r, column=2, value='Dự phòng'); ws.cell(row=r, column=3, value=SPARES[p]['spec'])
    ws.cell(row=r, column=7, value='%d lộ dự phòng (20 %% của %d lộ)' % (SPARES[p]['n_spare'], SPARES[p]['n_used'])); r += 2

# --- Chọn cáp & sụt áp ---
wc = wb.create_sheet('Chọn cáp & sụt áp')
title(wc, 'Chọn cáp theo khả năng mang tải và kiểm tra sụt áp (Iz = Iz0·k_nhiệt·k_nhóm ≥ In; ΔU tích lũy từ DB.HVAC ≤ 4 %)',
      'L tính = L ngang × 1,15 + L đứng; ΔU = √3·I·L·(ρ/S·cosφ + X·sinφ)/U (3P), 2·I·L·(…)/U (1P)')
cc = ['Mã mạch', 'Từ tủ', 'Đến', 'Số pha', 'I tính cáp (A)', 'Ib (A)', 'In (A)', 'In ≥ Ib', 'Ký hiệu cáp', 'S (mm²)', 'PE (mm²)',
      'PE yêu cầu (mm²)', 'Kiểm PE', 'Iz0 (A)', 'Iz (A)', 'Iz ≥ In', 'L ngang (m)', 'L đứng (m)', 'L tính (m)', 'cosφ',
      'ΔU mạch (%)', 'ΔU nguồn (%)', 'ΔU tích lũy (%)', 'Kiểm ΔU']
header(wc, 4, cc, [18, 14, 46, 6, 9, 9, 7, 8, 40, 8, 8, 9, 8, 8, 8, 8, 9, 9, 9, 7, 9, 9, 10, 9])
rows_c = [d for d in LOADS if d['cable']]
c0 = 5; FIRST_FEED = c0 + len(rows_c) + 2
def cable_row(r, d, feeder=False):
    pr = 'Phụ tải'
    if feeder:
        src_r = FROW[d['to']]
        wc.cell(row=r, column=1, value=d['id']); wc.cell(row=r, column=2, value='DB.HVAC'); wc.cell(row=r, column=3, value=d['to'])
        wc.cell(row=r, column=5, value="='%s'!O%d" % (pr, src_r)); wc.cell(row=r, column=6, value="='%s'!O%d" % (pr, src_r))
        wc.cell(row=r, column=20, value=d['pf'])
    else:
        src_r = ROW_OF[d['id']]
        wc.cell(row=r, column=1, value=d['id']); wc.cell(row=r, column=2, value=d['src']); wc.cell(row=r, column=3, value=d['desc'])
        wc.cell(row=r, column=5, value="='%s'!Q%d" % (pr, src_r)); wc.cell(row=r, column=6, value="='%s'!O%d" % (pr, src_r))
        wc.cell(row=r, column=20, value="='%s'!P%d" % (pr, src_r))
    wc.cell(row=r, column=4, value=d['ph'])
    wc.cell(row=r, column=7, value=d['In'])
    wc.cell(row=r, column=8, value='=IF(G%d>=F%d,"ĐẠT","KHÔNG ĐẠT")' % (r, r))
    wc.cell(row=r, column=9, value=d['cable'])
    wc.cell(row=r, column=10, value=d['S']); wc.cell(row=r, column=11, value=d['PE'])
    wc.cell(row=r, column=12, value='=IF(J%d<=16,J%d,IF(J%d<=35,16,J%d/2))' % (r, r, r, r))
    wc.cell(row=r, column=13, value='=IF(K%d>=L%d,"ĐẠT","KHÔNG ĐẠT")' % (r, r))
    wc.cell(row=r, column=14, value='=VLOOKUP(J%d,%s,IF(D%d=3,2,3),0)' % (r, TAB, r))
    wc.cell(row=r, column=15, value='=N%d*%s*%s' % (r, P_KT, P_KG))
    wc.cell(row=r, column=16, value='=IF(O%d>=G%d,"ĐẠT","KHÔNG ĐẠT")' % (r, r))
    wc.cell(row=r, column=17, value=d['route']['L_h']); wc.cell(row=r, column=18, value=d['route']['L_v'])
    wc.cell(row=r, column=19, value='=ROUND(Q%d*%s+R%d,1)' % (r, P_SL, r))
    wc.cell(row=r, column=21, value='=IF(D%d=3,SQRT(3),2)*E%d*S%d*(%s/J%d*T%d+%s*SQRT(1-T%d^2))/IF(D%d=3,%s,%s)*100' %
            (r, r, r, P_RHO, r, r, P_X, r, r, P_U3, P_U1))
    if feeder:
        wc.cell(row=r, column=22, value=0)
    else:
        wc.cell(row=r, column=22, value='=SUMIF($C$%d:$C$%d,B%d,$U$%d:$U$%d)' % (FIRST_FEED, FIRST_FEED + len(FEED) - 1, r, FIRST_FEED, FIRST_FEED + len(FEED) - 1))
    wc.cell(row=r, column=23, value='=U%d+V%d' % (r, r))
    wc.cell(row=r, column=24, value='=IF(W%d<=%s,"ĐẠT","KHÔNG ĐẠT")' % (r, P_DU))
for i, d in enumerate(rows_c):
    cable_row(c0 + i, d)
wc.cell(row=FIRST_FEED - 1, column=1, value='Cáp lộ tổng DB.HVAC → tủ DDC').font = BOLD
for i, f in enumerate(FEED):
    cable_row(FIRST_FEED + i, f, feeder=True)
C_LAST = FIRST_FEED + len(FEED) - 1
rr = C_LAST + 2
wc.cell(row=rr, column=1, value='Số mạch KHÔNG ĐẠT (In/PE/Iz/ΔU)').font = BOLD
wc.cell(row=rr, column=5, value='=COUNTIF(H%d:H%d,"KHÔNG ĐẠT")+COUNTIF(M%d:M%d,"KHÔNG ĐẠT")+COUNTIF(P%d:P%d,"KHÔNG ĐẠT")+COUNTIF(X%d:X%d,"KHÔNG ĐẠT")' %
        ((c0, C_LAST) * 4))
wc.cell(row=rr + 1, column=1, value='ΔU tích lũy lớn nhất (%)').font = BOLD
wc.cell(row=rr + 1, column=5, value='=MAX(W%d:W%d)' % (c0, C_LAST))
XL_FAIL = (rr, 5); XL_DUMAX = (rr + 1, 5)
rr += 3
wc.cell(row=rr, column=1, value='Tổng chiều dài cáp động lực theo ký hiệu (m)').font = BOLD
CAB_SUM_ROW = {}
for j, k in enumerate(sorted(PWQ, key=skey)):
    wc.cell(row=rr + 1 + j, column=9, value=k)
    wc.cell(row=rr + 1 + j, column=19, value='=SUMIF($I$%d:$I$%d,I%d,$S$%d:$S$%d)' % (c0, C_LAST, rr + 1 + j, c0, C_LAST))
    CAB_SUM_ROW[k] = rr + 1 + j
wc.freeze_panes = 'D5'

# --- I/O ---
wi = wb.create_sheet('IO')
title(wi, 'Danh mục điểm I/O DDC và cáp điều khiển/tín hiệu (L tính = L ngang × 1,15 + đứng)', 'Điểm "trong tủ" không cần cáp hiện trường')
header(wi, 4, ['STT', 'Tủ', 'Loại', 'Điểm', 'Thiết bị', 'Vị trí', 'Trong tủ', 'Cáp', 'L (m)'], [6, 14, 6, 60, 40, 16, 8, 52, 8])
i0 = 5
for i, q in enumerate(IO):
    r = i0 + i
    for j, v in enumerate([i + 1, q['panel'], q['type'], q['desc'], q['device'], q['loc'] or '', 'x' if q['internal'] else '', q['cable'], q['L']], 1):
        wi.cell(row=r, column=j, value=v)
iL = i0 + len(IO) - 1
r = iL + 1
for x in EXTRA_CTRL:
    r += 1
    for j, v in enumerate(['', 'Mạng/liên kết', 'NET', x['desc'], '', '', '', x['cable'], x['L']], 1):
        wi.cell(row=r, column=j, value=v)
iE = r
r += 2
wi.cell(row=r, column=2, value='TỔNG HỢP I/O').font = BOLD
header(wi, r + 1, ['', 'Tủ', 'DI', 'DO', 'AI', 'AO', 'Tổng', 'DI+20%', 'DO+20%', 'AI+20%', 'AO+20%'])
IO_ROW = {}
for j, pn in enumerate(list(IO_SUM) + ['TỔNG']):
    rr = r + 2 + j; IO_ROW[pn] = rr
    wi.cell(row=rr, column=2, value=pn)
    for k, t in enumerate(['DI', 'DO', 'AI', 'AO']):
        if pn == 'TỔNG':
            wi.cell(row=rr, column=3 + k, value='=SUM(%s%d:%s%d)' % (CL(3 + k), r + 2, CL(3 + k), rr - 1))
            wi.cell(row=rr, column=8 + k, value='=SUM(%s%d:%s%d)' % (CL(8 + k), r + 2, CL(8 + k), rr - 1))
        else:
            wi.cell(row=rr, column=3 + k, value='=COUNTIFS($B$%d:$B$%d,$B%d,$C$%d:$C$%d,"%s")' % (i0, iL, rr, i0, iL, t))
            wi.cell(row=rr, column=8 + k, value='=ROUNDUP(%s%d*1.2,0)' % (CL(3 + k), rr))
    wi.cell(row=rr, column=7, value='=SUM(C%d:F%d)' % (rr, rr))
r = rr + 2
wi.cell(row=r, column=2, value='TỔNG CÁP ĐIỀU KHIỂN THEO LOẠI (m)').font = BOLD
CT_SUM_ROW = {}
for j, k in enumerate(CTQ):
    wi.cell(row=r + 1 + j, column=8, value=k)
    wi.cell(row=r + 1 + j, column=9, value='=SUMIF($H$%d:$H$%d,H%d,$I$%d:$I$%d)' % (i0, iE, r + 1 + j, i0, iE))
    CT_SUM_ROW[k] = r + 1 + j
r = r + 2 + len(CTQ)
wi.cell(row=r, column=2, value='DANH MỤC CẢM BIẾN / THIẾT BỊ HIỆN TRƯỜNG').font = BOLD
for j, (k, v) in enumerate(SENSORS.items()):
    wi.cell(row=r + 1 + j, column=4, value=k); wi.cell(row=r + 1 + j, column=5, value=v['spec']); wi.cell(row=r + 1 + j, column=9, value=v['qty'])
wi.freeze_panes = 'A5'

# --- Máng cáp ---
wm = wb.create_sheet('Máng cáp')
title(wm, 'Mạng máng cáp (tọa độ WP2, mm; trong nhà BOP +3,6 m; ngoài nhà ốp tường A có nắp) – độ đầy ≤ 40 %',
      'Mỗi đoạn có máng động lực và máng điều khiển riêng (đi song song) khi có cáp tương ứng')
header(wm, 4, ['Đoạn', 'Tên', 'Môi trường', 'Từ nút', 'X1', 'Y1', 'Đến nút', 'X2', 'Y2', 'Đứng thêm (m)', 'L (m)', 'Số cáp ĐL',
               'Máng ĐL', 'Σ tiết diện cáp ĐL (mm²)', 'Độ đầy ĐL (%)', 'Số cáp ĐK', 'Máng ĐK', 'Σ tiết diện cáp ĐK (mm²)', 'Độ đầy ĐK (%)'],
       [6, 40, 9, 9, 8, 8, 9, 8, 8, 9, 8, 8, 32, 12, 10, 8, 32, 12, 10])
m0 = 5
for i, s in enumerate(SEGS):
    r = m0 + i
    vals = [s['i'] + 1, s['name'], s['env'], s['a'], s['xa'], s['ya'], s['b'], s['xb'], s['yb'], s['vz']]
    for j, v in enumerate(vals, 1):
        wm.cell(row=r, column=j, value=v)
    wm.cell(row=r, column=11, value='=ROUND((ABS(H%d-E%d)+ABS(I%d-F%d))/1000+J%d,2)' % (r, r, r, r, r))
    wm.cell(row=r, column=12, value=s['n_pw'])
    if 'pw_tray' in s and s['env'] != 'wall':
        wm.cell(row=r, column=13, value=tray_desc(s['pw_tray'], s['env']))
        wm.cell(row=r, column=14, value=s['pw_area'])
        wm.cell(row=r, column=15, value='=ROUND(N%d/(%d*%d)*100,1)' % (r, s['pw_tray'][0], s['pw_tray'][1]))
    elif s['env'] == 'wall':
        wm.cell(row=r, column=13, value='(ống lồng xuyên tường)')
    wm.cell(row=r, column=16, value=s['n_ct'])
    if 'ct_tray' in s and s['env'] != 'wall':
        wm.cell(row=r, column=17, value=tray_desc(s['ct_tray'], s['env']))
        wm.cell(row=r, column=18, value=s['ct_area'])
        wm.cell(row=r, column=19, value='=ROUND(R%d/(%d*%d)*100,1)' % (r, s['ct_tray'][0], s['ct_tray'][1]))
mL = m0 + len(SEGS) - 1
r = mL + 2
wm.cell(row=r, column=2, value='TỔNG THEO LOẠI MÁNG (m)').font = BOLD
TRAY_SUM_ROW = {}
for j, k in enumerate(TRAY_Q):
    rr = r + 1 + j
    wm.cell(row=rr, column=2, value=k)
    wm.cell(row=rr, column=11, value='=SUMIF($M$%d:$M$%d,B%d,$K$%d:$K$%d)+SUMIF($Q$%d:$Q$%d,B%d,$K$%d:$K$%d)' % (m0, mL, rr, m0, mL, m0, mL, rr, m0, mL))
    TRAY_SUM_ROW[k] = rr
rr += 1
wm.cell(row=rr, column=2, value='Độ đầy lớn nhất (%)')
wm.cell(row=rr, column=11, value='=MAX(O%d:O%d,S%d:S%d)' % (m0, mL, m0, mL))
XL_FILLMAX = (rr, 11)

# --- Tổng hợp khối lượng ---
wq = wb.create_sheet('Tổng hợp khối lượng')
title(wq, 'Tổng hợp khối lượng WP4 – Điện & điều khiển HVAC (đầu vào WP5/BOQ)',
      'Khối lượng cáp/máng lấy bằng công thức từ các sheet tính; nhóm J là phương án chọn thêm, không cộng')
header(wq, 4, ['STT', 'Nhóm', 'Hạng mục', 'Quy cách', 'ĐVT', 'Khối lượng', 'Ghi chú'], [6, 30, 70, 70, 7, 11, 40])
for i, b in enumerate(BOQ):
    r = 5 + i
    for j, v in enumerate([b['no'], b['group'], b['item'], b['spec'], b['unit']], 1):
        wq.cell(row=r, column=j, value=v)
    q = b['qty']
    if b['group'] == G2:
        k = b['item'][4:]
        q = "=ROUNDUP('Chọn cáp & sụt áp'!S%d,0)" % CAB_SUM_ROW[k]
    elif b['group'] == G3:
        q = "=ROUNDUP('IO'!I%d,0)" % CT_SUM_ROW[b['item']]
    elif b['group'] == G4 and b['item'] in TRAY_SUM_ROW:
        q = "=ROUNDUP('Máng cáp'!K%d,0)" % TRAY_SUM_ROW[b['item']]
    wq.cell(row=r, column=6, value=q); wq.cell(row=r, column=7, value=b['note'])
for wsx in wb.worksheets:
    for row in wsx.iter_rows():
        for c in row:
            if isinstance(c.value, str) and not c.value.startswith('='):
                c.alignment = Alignment(wrap_text=True, vertical='top')
tmpd = tempfile.mkdtemp()
raw = os.path.join(tmpd, 'WP4_Electrical.xlsx')
wb.save(raw)
XLSX = os.path.join(HERE, 'WP4_Electrical.xlsx')
try:
    subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', HERE, raw], check=True,
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=240)
    recalc = True
except Exception as ex:
    shutil.copy(raw, XLSX); recalc = False
    print('LibreOffice recalc FAILED:', ex)
from openpyxl import load_workbook
wv = load_workbook(XLSX, data_only=True)
ERR = []
for wsx in wv.worksheets:
    for row in wsx.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith('#') or c.value.startswith('Err:')):
                ERR.append('%s!%s=%s' % (wsx.title, c.coordinate, c.value))
wcv = wv['Chọn cáp & sụt áp']; wpv = wv['Phụ tải']; wmv = wv['Máng cáp']
XCHK = dict(errors=len(ERR), error_cells=ERR[:20], recalculated=recalc,
            fail_count=wcv.cell(*XL_FAIL).value, dU_max_pct=wcv.cell(*XL_DUMAX).value,
            demand_kVA=wpv.cell(*XL_DEM).value, Ib_max_A=wpv.cell(*XL_IMAX).value, tray_fill_max=wmv.cell(*XL_FILLMAX).value)
assert XCHK['errors'] == 0, ERR[:10]
assert XCHK['fail_count'] == 0
assert abs(XCHK['demand_kVA'] - dmax['kVA']) < 0.2, (XCHK['demand_kVA'], dmax['kVA'])
assert abs(XCHK['dU_max_pct'] - max(d['du_cum'] for d in LOADS if d['cable'])) < 0.02
print('Excel check:', XCHK)

# ==============================================================================================
# 16. HÌNH: SƠ ĐỒ NGUYÊN LÝ & MẶT BẰNG MÁNG CÁP
# ==============================================================================================
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon

def short_cable(c):
    return c.replace('CXV/S (màn chắn Cu, cáp biến tần)', 'CXV/S').replace(' + E', '+E')

def board(ax, x0, y, name, items, w=1.25, color='#1f4e79', sub=''):
    n = len(items)
    ax.plot([x0 - 0.3, x0 + (n - 1) * w + 0.3], [y, y], color=color, lw=6, solid_capstyle='butt')
    ax.text(x0 - 0.5, y, name, ha='right', va='center', fontsize=10, weight='bold', color=color)
    if sub:
        ax.text(x0 - 0.5, y - 0.45, sub, ha='right', va='center', fontsize=6.5, color=color)
    for i, it in enumerate(items):
        x = x0 + i * w
        ax.plot([x, x], [y, y - 0.5], color='k', lw=1)
        ax.add_patch(Rectangle((x - 0.18, y - 0.95), 0.36, 0.45, fc='white', ec='k', lw=0.9))
        ax.plot([x - 0.12, x + 0.12], [y - 0.9, y - 0.55], color='k', lw=0.9)
        ax.plot([x, x], [y - 0.95, y - 1.6], color='k', lw=1, ls='--' if it.get('spare') else '-')
        ax.text(x, y - 1.7, it['txt'], rotation=90, ha='center', va='top', fontsize=7.2, family='DejaVu Sans')
        if it.get('rcd'):
            ax.text(x + 0.22, y - 0.75, 'RCD', fontsize=4.5, color='#c00000')
    return x0 + (n - 1) * w

def item_of(d):
    if d.get('kind') == 'FEEDER':
        return dict(txt='%s %dA\n%s\n%.0f m  ΔU %.2f%%\n→ %s' % ('MCCB 4P', d['In'], short_cable(d['cable']), d['route']['L'], d['du'], d['to']))
    if not d['cable']:
        return dict(txt='MCB 1P+N 6A\n(trong tủ)\n%s' % d['desc'][:34])
    pro = d['prot'].replace(', Icu ≥ 25 kA', '').replace(', 10 kA', '').replace(' 10 kA', '').replace(' (C)', '').replace('+ biến tần', '+VFD').replace('+ contactor AC3', '+K')
    return dict(txt='%s\n%s\n%.0f m  ΔU %.2f%%\n%s %s kW' % (pro[:34], short_cable(d['cable']), d['route']['L'], d['du_cum'], d['tag'], vn(d['P'], 2)),
                rcd=bool(d['rcd']))

fig, ax = plt.subplots(figsize=(36, 21))
ax.set_xlim(-7, 41); ax.set_ylim(-19.5, 3.6); ax.axis('off')
dbi = [item_of(d) for d in LOADS if d['src'] == 'DB.HVAC'] + [item_of(f) for f in FEED] + \
      [dict(txt='Dự phòng\n' + SPARES['DB.HVAC']['spec'].split(',')[0], spare=True) for _ in range(SPARES['DB.HVAC']['n_spare'])]
ax.text(9, 3.2, 'Nguồn từ tủ hạ thế tổng (gói điện – ngoài phạm vi HVAC): %s' % INCOMING['incoming_cable'], ha='center', fontsize=10, color='#666666')
ax.plot([9, 9], [2.9, 1.6], color='#666666', lw=2, ls='--')
ax.add_patch(Rectangle((8.6, 1.0), 0.8, 0.6, fc='#fff2cc', ec='k'))
ax.text(9.6, 1.3, '%s\nĐồng hồ đa năng Modbus · SPD T2 40 kA · đèn pha' % INCOMING['MCCB'], fontsize=8, va='center')
ax.plot([9, 9], [1.0, 0.3], color='k', lw=2)
board(ax, 0, 0.3, 'DB.HVAC', dbi, w=1.35, sub='Thanh cái %d A · Nhu cầu max %s kVA (%s A)' % (BUSBAR, vn(dmax['kVA'], 1), vn(Ib_db, 0)))
yrow = -6.6
x = 0
for p in ['DDC-AHU-2.01', 'DDC-AHU-2.02', 'DDC-AHU-2.03']:
    its = [item_of(d) for d in LOADS if d['src'] == p] + [dict(txt='Dự phòng\n' + SPARES[p]['spec'][:20], spare=True) for _ in range(SPARES[p]['n_spare'])]
    xe = board(ax, x + 1.5, yrow, p, its, w=1.15, color='#375623', sub='Lộ %s' % FEED_BY[p]['cable'])
    x = xe + 2.6
its = [item_of(d) for d in LOADS if d['src'] == 'DDC-FAN'] + [dict(txt='Dự phòng\n' + SPARES['DDC-FAN']['spec'][:24], spare=True) for _ in range(SPARES['DDC-FAN']['n_spare'])]
board(ax, 1.5, -13.0, 'DDC-FAN', its, w=1.35, color='#7f3f00', sub='Lộ %s' % FEED_BY['DDC-FAN']['cable'])
ax.text(-6.5, -19.3, 'Ghi chú: ΔU ghi tại mỗi lộ cuối là ΔU tích lũy từ DB.HVAC (≤ 4 %). Cáp động cơ biến tần dùng CXV/S màn chắn, ốc siết EMC. '
        'Lộ CDU có ELCB 300 mA theo hãng; nhóm sưởi RCD 300 mA; tải 1 pha RCBO 30 mA. Mỗi động cơ/cụm sưởi/CDU có cầu dao cách ly IP65 tại chỗ.\n'
        'Liên động cứng: báo cháy → STO/enable biến tần AHU + sưởi + FAF; airflow + thermostat an toàn → cuộn hút contactor/SSR sưởi; phao nước ngưng → DX-kit.',
        fontsize=10)
ax.set_title('WP4 – SƠ ĐỒ NGUYÊN LÝ CẤP ĐIỆN HVAC PHASE 2 (DB.HVAC → tủ DDC → thiết bị)', fontsize=16, weight='bold')
fig.savefig(os.path.join(HERE, 'WP4_single_line.png'), dpi=90, bbox_inches='tight')
plt.close(fig)

# --- mặt bằng máng cáp ---
CAL = GEO['calibration']
ax_, bx_ = CAL['x_mm = a*x_pt + b']; ay_, by_ = CAL['y_mm = a*top_pt + b']
fig, ax = plt.subplots(figsize=(24, 22))
try:
    import pypdfium2 as pdfium
    from PIL import ImageEnhance
    doc = pdfium.PdfDocument(PDF)
    sc = 2
    img = doc[6].render(scale=sc).to_pil().convert('L').convert('RGB')
    img = ImageEnhance.Brightness(img).enhance(1.3)
    W, H = img.size
    ext = [bx_, ax_ * W / sc + bx_, ay_ * H / sc + by_, by_]
    ax.imshow(img, extent=ext, alpha=0.35)
except Exception as ex:
    print('background skipped:', ex)
for r in GEO['rooms']:
    ax.add_patch(Polygon(r['polygon_mm'], closed=True, fc='#1f77b4', alpha=0.05, ec='#1f77b4', lw=1))
    ax.text(r['centroid_mm'][0], r['centroid_mm'][1] - 600, r['id'], fontsize=8, color='#08306b', ha='center', alpha=0.8)
tcol = {(100, 50): '#2ca02c', (200, 100): '#d62728', (300, 100): '#9467bd', (400, 100): '#8c564b'}
for s in SEGS:
    xa, ya, xb, yb = s['xa'], s['ya'], s['xb'], s['yb']
    if s['env'] == 'wall':
        ax.plot([xa, xb], [ya, yb], color='k', lw=5, ls=':'); continue
    if 'pw_tray' in s:
        ax.plot([xa, xb], [ya, yb], color=tcol[s['pw_tray']], lw=5 if s['pw_tray'][0] >= 200 else 3.5, ls='--' if s['env'] == 'out' else '-', zorder=4,
                solid_capstyle='butt')
    if 'ct_tray' in s:
        off = 180
        dx, dy = (0, off) if ya == yb else (off, 0)
        ax.plot([xa + dx, xb + dx], [ya + dy, yb + dy], color='#17becf', lw=2, ls='--' if s['env'] == 'out' else '-', zorder=4)
for d in LOADS:
    if not d['cable']:
        continue
    nn = d['route']['node']; xy = d['xy']
    ax.plot([N[nn][0], xy[0], xy[0]], [N[nn][1], N[nn][1], xy[1]], color='#ff7f0e', lw=1, zorder=3)
    ax.plot(xy[0], xy[1], 's' if d['kind'] != 'DOL' else 'o', ms=5, color='#ff7f0e', mec='k', zorder=5)
for q in IO:
    if q['internal'] or not q['cable'] or q['loc'] in (None, 'FA'):
        continue
    xy = loc_xy(q['loc'])
    if q['loc'].startswith('room:'):
        ax.plot(xy[0], xy[1], '^', ms=7, color='#17becf', mec='k', zorder=5)
    elif q['loc'].startswith('MD-G') or (q['type'] == 'DO' and q['loc'].startswith('hood:')):
        ax.plot(xy[0] + 250, xy[1] + 250, 'D', ms=4.5, color='#bcbd22', mec='k', zorder=5)
for p, v in PANELS.items():
    x, y = v['xy']
    w, dpt = (1200 if p in ('DB.HVAC', 'DDC-FAN') else 1000), (500 if p in ('DB.HVAC', 'DDC-FAN') else 400)
    rx = 16185 - dpt if x > XC else 14385
    ax.add_patch(Rectangle((rx, y - w / 2), dpt, w, fc='#000000' if p == 'DB.HVAC' else '#555555', ec='k', zorder=6))
    ax.text(rx + (dpt + 250 if x > XC else -250), y, p, fontsize=8, weight='bold', ha='left' if x > XC else 'right', va='center', zorder=7,
            bbox=dict(fc='#ffffcc', ec='k', lw=0.5, pad=1.5))
for eid in ['AHU-2.01', 'AHU-2.02', 'AHU-2.03', 'FAF-2.01', 'FAF-2.02', 'EAF-2.01', 'EAF-2.02', 'EAF-2.03', 'DH-2.01', 'SCR-2.01', 'CDU-YARD']:
    q = GEQ[eid]; x0, y0, x1, y1 = q['rect_mm']
    ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fill=False, ec='#0b6e4f', lw=1.2, ls='-.', zorder=2))
    ax.text(q['center_mm'][0], q['center_mm'][1] + 150, eid, fontsize=7, color='#0b6e4f', ha='center', weight='bold', zorder=6)
for m, (x, y) in CDU_XY.items():
    ax.add_patch(Rectangle((x - 450, y - 400), 900, 800, fc='#ccccff', ec='k', zorder=5)); ax.text(x + 550, y, m, fontsize=6.5, va='center')
ax.axvline(28000, color='k', lw=2)
import matplotlib.lines as mlines
hs = [mlines.Line2D([], [], color=c, lw=5, label='Máng động lực %d×%d' % k) for k, c in tcol.items() if any(s.get('pw_tray') == k for s in SEGS)]
hs += [mlines.Line2D([], [], color='#17becf', lw=2, label='Máng điều khiển 100×50 (đi song song)'),
       mlines.Line2D([], [], color='k', lw=1.5, ls='--', label='Ngoài trời: máng HDG có nắp'),
       mlines.Line2D([], [], color='k', lw=4, ls=':', label='Xuyên tường A (ống lồng)'),
       mlines.Line2D([], [], color='#ff7f0e', lw=1, marker='s', label='Ống EMT/GI xuống thiết bị'),
       mlines.Line2D([], [], color='#17becf', marker='^', ls='', mec='k', label='Cảm biến T/RH phòng'),
       mlines.Line2D([], [], color='#bcbd22', marker='D', ls='', mec='k', label='Van gió bù động cơ')]
ax.legend(handles=hs, loc='lower left', fontsize=10, framealpha=0.95)
ax.set_xlim(-1500, 33000); ax.set_ylim(-16000, 20500); ax.set_aspect('equal'); ax.grid(True, ls=':', lw=0.4)
ax.set_title('WP4 – Mặt bằng tuyến máng cáp & cấp điện HVAC (mm, gốc E/3; máng BOP +3,6 m) – nền DW-A-03-06', fontsize=15)
fig.savefig(os.path.join(HERE, 'WP4_containment.png'), dpi=85, bbox_inches='tight')
plt.close(fig)

# ==============================================================================================
# 17. BÁO CÁO
# ==============================================================================================
def tbl(head, rows):
    s = '| ' + ' | '.join(head) + ' |\n|' + '---|' * len(head) + '\n'
    for r in rows:
        s += '| ' + ' | '.join(str(x) for x in r) + ' |\n'
    return s
pw_rows = [(k, vn(PWQ[k], 0)) for k in sorted(PWQ, key=skey)]
ct_rows = [(k, vn(v, 0)) for k, v in CTQ.items()]
R = []
R.append('# WP4 – ĐIỆN & ĐIỀU KHIỂN HVAC – SGS HẢI PHÒNG, KHU H&N PHASE 2 (NHÀ XƯỞNG 3)\n')
R.append('> v1.0 – 14/11/2026. Phụ thuộc: `01_design_basis/DESIGN_BASIS.md` v1.1, WP1 `02_calc/equipment_schedule.json` (v4), WP2 `03_geometry/geometry.json`.\n'
         '> Kết quả: `electrical_design.json`, `takeoff_electrical.json`, `WP4_Electrical.xlsx` (7 sheet: Bảng tra, Phụ tải, Sơ đồ nguyên lý, Chọn cáp & sụt áp, IO – Excel không cho ký tự “/” trong tên sheet, Máng cáp, Tổng hợp khối lượng; công thức, đã tính lại bằng LibreOffice – %d lỗi), '
         '`WP4_single_line.png`, `WP4_containment.png`. Chạy lại: `python3 -I 05_electrical/build_electrical.py`.\n' % XCHK['errors'])
R.append('## 1. Phạm vi\n\nTheo DESIGN_BASIS §0 và §4.6: **từ tủ tổng DB.HVAC trở đi** (gồm vỏ tủ, MCCB tổng, thanh cái, các lộ ra), các tủ DDC/động lực riêng, '
         'toàn bộ cáp lộ ra, máng/ống, cáp điều khiển, cảm biến, lập trình & T&C. **Ngoài phạm vi**: cáp và thiết bị cấp nguồn tới DB.HVAC, thanh tiếp địa chính tòa nhà, '
         'hệ báo cháy (chỉ nhận tín hiệu từ module rơ le của gói PCCC), thiết bị (AHU, CDU, quạt, scrubber, DH, bơm – WP5), thân van gió động cơ (WP3), quạt tủ hood (gói nội thất).\n')
R.append('## 2. Vị trí tủ (giả định WP4)\n\nKhông có phòng điện/“ELE”/“Tủ IT” trong khu Phase 2 theo WP2 (13 phòng đều là phòng chức năng; khu trái trên/phải trên là ngoài phạm vi). '
         'Do đó các tủ đặt **tại lối đi kỹ thuật giữa 2 khối** (X 14.385…16.185, thông thủy 1,8 m, không điều hòa, không phải phòng thí nghiệm), cao độ sàn, '
         'tiếp cận trực tiếp – đúng vị trí gợi ý DB.HVAC của WP2 (X ≈ 16.400, Y ≈ 2.200). Biến tần/contactor nằm trong tủ ở sàn (không đặt trên sàn AHU) để bảo trì an toàn; '
         'trên sàn thiết bị chỉ có cầu dao cách ly + ổ cắm bảo trì.\n')
R.append(tbl(['Tủ', 'Tọa độ tâm (mm)', 'Lắp đặt', 'Kích thước', 'IP', 'Vị trí'],
             [(p, '%d; %d' % v['xy'], v['mount'], v['dims'], v['ip'], v['wall']) for p, v in PANELS.items()]))
R.append('\nKhoảng trống trước tủ đặt sàn còn ≥ 1,3 m (lối đi 1,8 m − tủ 0,5 m), tủ treo tường hai bên lệch nhau theo Y (không đối diện) – thỏa IEC 60364-7-729 (≥ 0,7 m). '
         'Cần CĐT xác nhận lối đi được phép đặt tủ (lối thoát nạn?) – vấn đề mở #1.\n')
R.append('## 3. Phương pháp & giả định\n')
R.append(tbl(['#', 'Giả định / tham số', 'Giá trị', 'Nguồn'], [
    (1, 'Điện áp', '380/220 V, 3P+N+PE, 50 Hz, TN-S', 'Lưới VN; Phase 1'),
    (2, 'cosφ', 'Động cơ/CDU 0,85; tải qua biến tần 0,95 (phía lưới); sưởi điện trở 1,0', 'Chỉ đạo WP4; vật lý tải điện trở'),
    (3, 'Hiệu suất', 'Động cơ IE3 4 cực (0,75 kW 82,5 % … 15 kW 92,1 %); biến tần 97 %; quạt hood 85 %; scrubber 88 %', 'IEC 60034-30-1; WP1 FLA'),
    (4, 'P CDU', 'Điện vào theo WP1 (COP 3,0): 16,54 / 13,43 / 13,45 kW/module; Ib = max(tính, FLA WP1)', 'equipment_schedule.json'),
    (5, 'Chọn bảo vệ', 'In ≥ 1,25·Ib (MCB/MCCB); động cơ trực tiếp: MPCB theo dải dòng; lộ tổng tủ DDC ≥ 1,25·ΣIb và ≥ 1,6·In lớn nhất phía sau (chọn lọc)', 'IEC 60364-4-43; thực hành'),
    (6, 'Khả năng mang tải cáp', 'Cu/XLPE/PVC (CXV) nhiều lõi trên khay đục lỗ – PP E, Bảng B.52.12', 'IEC 60364-5-52 ≡ TCVN 9207:2012'),
    (7, 'Hệ số hiệu chỉnh', 'Nhiệt độ 45 °C trên trần: 0,87; nhóm: 0,80 → Iz = 0,696·Iz0', 'B.52.14; chỉ đạo WP4 (B.52.17: mạch < 30 % Iz bỏ qua)'),
    (8, 'Chiều dài', 'Theo mạng máng (tọa độ WP2) + đoạn xuống thiết bị (Manhattan); ngang × 1,15 + đứng (máng đứng nóc tủ 1,5 m, trong tủ 1 m, xuống thiết bị trên trần 1 m, cảm biến tường 2,6 m, CDU 3,1 + 1 m, scrubber 4,1 m)', 'WP2; chỉ đạo WP4'),
    (9, 'Sụt áp', 'ΔU tích lũy DB.HVAC → tải ≤ 4 %; cáp lộ tổng mục tiêu ≤ 1,5 %; ρ = 0,0225 Ω·mm²/m, X = 0,08 mΩ/m', 'IEC 60364-5-52 G.52'),
    (10, 'Tiết diện PE', 'S ≤ 16: PE = S; 16 < S ≤ 35: 16; S > 35: S/2 (lõi PE trong cáp)', 'IEC 60364-5-54 Bảng 54.2'),
    (11, 'Máng cáp', 'BOP +3,6 m (như Phase 1 F3-SD-AC-013 “CABLE TRAY … BOP=FFL3600”); máng động lực và máng điều khiển riêng; độ đầy ≤ 40 %', 'Phase 1; thực hành VN'),
    (12, 'Dự phòng', '20 % số lộ mỗi tủ (làm tròn lên), 20 % I/O mỗi DDC', 'Chỉ đạo WP4'),
    (13, 'Ik tại DB.HVAC', 'Chưa có số liệu → MCCB tổng Icu ≥ 36 kA, lộ ra ≥ 25 kA, tủ DDC 10 kA', 'Giả định – gói điện xác nhận'),
    (14, 'Quạt EAF', 'Khởi động trực tiếp (MPCB + contactor) theo chỉ đạo WP4 (lưu lượng cố định, cân chỉnh bằng van). WP1 ghi “biến tần” – xem vấn đề mở #4', 'Chỉ đạo WP4'),
    (15, 'Quạt AHU', 'Động cơ AC + biến tần trong tủ DDC (nếu chọn quạt EC thì bỏ biến tần, giữ tín hiệu 0–10 V)', 'WP1: “quạt plug EC/biến tần”'),
    (16, 'Bản vẽ điện Phase 1', 'Thư mục `00_inputs/phase1_dwg` chỉ có bản vẽ AC (không có F3-SD-EE-002/011/012). Tham khảo được: máng cáp 400×200 BOP FFL3600, “MÀN HÌNH HMI” (F3-SD-AC-017), DB-27…30', 'phase1_dwg/*.txt'),
]))
R.append('\n## 4. Phụ tải & DB.HVAC\n')
R.append('CDU **cấp trực tiếp từ DB.HVAC** (6 lộ MCCB 4P + ELCB 300 mA theo hãng), cầu dao cách ly 4P IP65 ngoài trời tại mỗi module. Lý do: tải lớn nhất (90 kW), ở ngoài nhà, '
         'được AHU điều khiển qua DX-kit bằng tín hiệu (enable/0–10 V/trạng thái/lỗi) – không cần đóng cắt nguồn trong tủ DDC; tủ DDC nhỏ gọn, dễ chọn lọc. '
         'Scrubber: HVAC cấp nguồn tới cầu dao cách ly + tủ điều khiển trọn bộ của nhà cung cấp; DDC-FAN trao đổi lệnh chạy/trạng thái/lỗi.\n')
R.append(tbl(['Chỉ tiêu', 'Giá trị'], [
    ('Tổng công suất đặt (định mức, không gồm ổ cắm/điều khiển)', '%s kW (WP1: 191,7 kW)' % vn(P_conn_rated, 1)),
    ('Tổng công suất điện vào (kể hiệu suất) / S đặt', '%s kW / %s kVA (%s A)' % (vn(P_conn_in, 1), vn(S_conn, 1), vn(INCOMING['I_conn_A'], 0))),
] + [('Kịch bản: ' + c['case'], '%s kW / %s kVA / %s A' % (vn(c['kW'], 1), vn(c['kVA'], 1), vn(c['I_A'], 0))) for c in cases] + [
    ('**Nhu cầu lớn nhất (thiết kế)**', '**%s kW / %s kVA / Ib = %s A** (%s)' % (vn(dmax['kW'], 1), vn(dmax['kVA'], 1), vn(Ib_db, 0), dmax['case'])),
    ('Hệ số nhu cầu chung', vn(INCOMING['demand_factor_overall'], 2)),
    ('MCCB tổng', INCOMING['MCCB']),
    ('Thanh cái', INCOMING['busbar']),
    ('Cáp cấp nguồn khuyến nghị (thông tin cho gói điện)', INCOMING['incoming_cable']),
    ('Ghi chú cáp nguồn', INCOMING['incoming_cable_note']),
    ('Lộ ra DB.HVAC', '%d lộ dùng + %d lộ dự phòng (%s)' % (SPARES['DB.HVAC']['n_used'], SPARES['DB.HVAC']['n_spare'], SPARES['DB.HVAC']['spec'])),
]))
R.append('\nHệ số đồng thời theo kịch bản (Hè / Đông / Nồm): CDU-2.01, 2.02: 1 / 0 / 0,7; CDU-2.03 (Instrument tải thiết bị lớn, chạy cả mùa đông): 1 / 0,5 / 0,7; '
         'sưởi: 0 / 1 / 1 (sưởi gia nhiệt lại khi khử ẩm); DH-2.01: 1 / 0,5 / 1; quạt, quạt hood, scrubber, bơm, điều khiển: 1; ổ cắm bảo trì 0,2. '
         'Kịch bản Nồm (khử ẩm + gia nhiệt lại) là bất lợi nhất; sưởi và lạnh không cộng 100 %/100 %.\n')
R.append('\n### 4.1 Lộ ra DB.HVAC và cáp lộ tổng tủ DDC\n')
R.append(tbl(['Lộ', 'Phụ tải', 'Ib (A)', 'Bảo vệ', 'Cáp', 'L (m)', 'ΔU (%)'],
             [(d['id'], d['desc'][:60], vn(d['Ib'], 1), d['prot'], d['cable'], vn(d['route']['L'], 1), vn(d['du'], 2)) for d in LOADS if d['src'] == 'DB.HVAC'] +
             [(f['id'], f['desc'], vn(f['Ib'], 1), f['prot'], f['cable'], vn(f['route']['L'], 1), vn(f['du'], 2)) for f in FEED]))
R.append('\n### 4.2 Mạch cuối trong tủ DDC\n')
R.append(tbl(['Mạch', 'Phụ tải', 'Ib / I cáp (A)', 'Bảo vệ', 'Dòng rò', 'Cáp', 'L (m)', 'ΔU tích lũy (%)', 'Cách ly'],
             [(d['id'], d['desc'][:58], '%s / %s' % (vn(d['Ib'], 1), vn(d['Icab'], 1)), d['prot'], d['rcd'] or '–', d['cable'] if d['cable'] else '(trong tủ)',
               vn(d['route']['L'], 1) if d['cable'] else '–', vn(d['du_cum'], 2), ('%d A' % d['iso_A']) if d.get('iso_A') else '–')
              for d in LOADS if d['src'] != 'DB.HVAC']))
R.append('\nKiểm tra (Excel, sau khi LibreOffice tính lại): %d mạch không đạt; ΔU tích lũy lớn nhất %s %% (≤ 4 %%); độ đầy máng lớn nhất %s %% (≤ 40 %%).\n'
         % (XCHK['fail_count'], vn(XCHK['dU_max_pct'], 2), vn(XCHK['tray_fill_max'], 1)))
R.append('## 5. Tủ & thành phần chính\n')
for p in PANELS:
    R.append('**%s** – %s, %s, %s.\n' % (p, PANELS[p]['dims'], PANELS[p]['mount'], PANELS[p]['ip']))
    R.append(''.join('- %s × %s %s\n' % (c[0], c[2], c[1]) for c in PANEL_BOM[p]))
    R.append('\n')
R.append('## 6. I/O DDC & cảm biến\n')
R.append(tbl(['Tủ', 'DI', 'DO', 'AI', 'AO', 'Tổng', 'Chọn DDC (+20 %)'],
             [(p, c['DI'], c['DO'], c['AI'], c['AO'], c['total'], '%d/%d/%d/%d' % (c['DI_spare'], c['DO_spare'], c['AI_spare'], c['AO_spare'])) for p, c in IO_SUM.items()] +
             [('**Tổng**', IO_TOT['DI'], IO_TOT['DO'], IO_TOT['AI'], IO_TOT['AO'], IO_TOT['total'], '')]))
R.append('\nChi tiết từng điểm (mô tả, thiết bị, vị trí, cáp, chiều dài) trong sheet `IO` và `electrical_design.json → io_schedule`. '
         'Liên động an toàn được **đi dây cứng** (không phụ thuộc DDC): báo cháy → STO/enable biến tần + contactor sưởi; airflow switch + thermostat an toàn + quạt chạy → mạch cuộn hút sưởi; phao mức cao bơm ngưng → DX-kit.\n')
R.append(tbl(['Cảm biến / thiết bị hiện trường', 'Quy cách', 'SL'], [(k, v['spec'], v['qty']) for k, v in SENSORS.items()]))
R.append('\nVan gió bù nhóm theo quạt EAF: ' + '; '.join('%s ← %s (%s)' % (k, v[0], ', '.join(v[1])) for k, v in GROUP_DAMPERS.items()) + '.\n')
R.append('\n## 7. Trình tự điều khiển (tóm tắt)\n')
R.append(''.join('- **%s**: %s\n' % (a, b) for a, b in SEQUENCES))
R.append('\n## 8. Máng cáp, ống, cách ly\n')
R.append('Tuyến (xem `WP4_containment.png`): máng chính dọc lối đi X = 15.285 (Y −4.800…16.000) nối 5 tủ; khối trái: nhánh AHU-2.01 (Y 5.150), nhánh bắc Y 11.000 → '
         'FAF-2.01, hood ST1/Digestion (X 7.000), nhánh EAF-2.01/Kjeldahl (Y 16.000), nhánh nam Y −1.200 → AHU-2.03, DH-2.01, EAF-2.03, Sample prep/Coding/Dividing; '
         'khối phải: máng chính Y 2.200 qua Office → ST2 → dọc tường A (X 27.300), nhánh Standard 19,72 (X 18.000), nhánh hood ST2 (X 21.700 và Y 9.000). '
         'Ngoài nhà: xuyên tường A tại Y −5.500 (ống lồng) → máng HDG có nắp ốp tường +3,6 m vượt ram dốc → máng đứng xuống bãi CDU → máng HDG có nắp dọc bãi, ống GI ren tới từng cầu dao cách ly; '
         'phía bắc: xuyên tường A tại Y 9.300 → máng HDG có nắp ốp tường tới scrubber (Y 15.500) → ống GI xuống tủ scrubber.\n')
R.append(tbl(['Loại máng', 'Chiều dài (m)'], [(k, vn(v, 1)) for k, v in TRAY_Q.items()] + [('**Tổng**', vn(L_tray_total, 1))]))
R.append('\n' + tbl(['Phụ kiện máng', 'SL'], list(FIT.items())))
R.append('\n' + tbl(['Ống luồn / ống mềm', 'm'], [(k, vn(v, 0)) for k, v in sorted(CONDQ.items())] + [(k, vn(v, 0)) for k, v in sorted(FLEXQ.items())]))
R.append('\n' + tbl(['Cầu dao cách ly tại chỗ', 'SL'], sorted(ISOQ.items())))
R.append('\n' + tbl(['Ốc siết cáp', 'SL'], sorted(GLQ.items())))
R.append('\n## 9. Khối lượng cáp\n')
R.append(tbl(['Cáp động lực', 'm'], pw_rows + [('**Tổng**', vn(sum(PWQ.values()), 0))]))
R.append('\n' + tbl(['Cáp điều khiển / tín hiệu / mạng', 'm'], ct_rows + [('**Tổng**', vn(sum(CTQ.values()), 0))]))
R.append('\n## 10. Tiếp địa\n\nPE đi chung trong mỗi cáp (lõi PE theo §3 #10); thanh PE mỗi tủ DDC nối về DB.HVAC qua lõi PE cáp lộ tổng; DB.HVAC nối thanh tiếp địa chính tòa nhà (MET). '
         'Máng cáp động lực có dây tiếp địa CV 1×25 mm² chạy suốt + cầu nối CV 6 mm² tại mỗi mối nối; vỏ CDU/scrubber nối đẳng thế CV 16 mm².\n')
R.append(tbl(['Hạng mục', 'ĐVT', 'SL'], [(k, u, vn(v, 0)) for k, (u, v) in EARTH.items()]))
R.append('\n## 11. Phương án chọn thêm – kết nối BMS/HMI Phase 1 (không cộng vào tổng)\n')
R.append(tbl(['Mã', 'Hạng mục', 'ĐVT', 'SL'], OPTIONS))
R.append('\n## 12. Vấn đề còn mở\n')
issues = [
    'Vị trí tủ trên lối đi giữa 2 khối: CĐT/PCCC xác nhận lối đi không phải lối thoát nạn bắt buộc chiều rộng ≥ 1,8 m; nếu không, chuyển DB.HVAC + DDC-FAN vào khu ngoài phạm vi phải trên (CST) hoặc làm hốc tủ âm vách Office.',
    'Gói điện cung cấp: Ik tại DB.HVAC, vị trí tủ hạ thế tổng, chiều dài & sụt áp cáp nguồn (khuyến nghị ≤ 1 %%), có cho phép %s A dự kiến không (kiểm tra máy biến áp hiện hữu).' % vn(Ib_db, 0),
    'Hãng CDU: yêu cầu ELCB (loại/độ nhạy), cấu hình DX-kit (số tín hiệu, 0–10 V hay Modbus), cáp truyền tín hiệu – chốt khi chọn hãng (WP5). Nếu CDU-2.01 phải 4 module (WP1 ghi chú) → thêm 1 lộ MCCB + cáp + cách ly (đã có lộ dự phòng 63 A).',
    'Quạt EAF: WP1 ghi “biến tần”, chỉ đạo WP4 dùng khởi động trực tiếp. Nếu cần biến tần để cân chỉnh lưu lượng → thay MPCB+contactor bằng biến tần 0,75–1,1 kW trong DDC-FAN (cáp đổi sang loại màn chắn, cùng tiết diện).',
    'Quạt tủ hood (gói nội thất): xác nhận quạt 380 V/3P 0,6 kW, vị trí quạt (nóc tủ/trên trần) và việc dùng hộp công tắc của HVAC thay công tắc đi kèm tủ; tủ có cảm biến sash báo động (không vào DDC).',
    'Báo cháy: tư vấn PCCC xác nhận chiến lược – WP4 dừng AHU/FAF, GIỮ chạy quạt hút hóa chất; module rơ le báo cháy và cáp từ tủ báo cháy tới module thuộc gói PCCC.',
    'Tuyến máng ngoài nhà phía bắc tới scrubber đi ốp tường A (không đi qua trần khu CST ngoài phạm vi) – cần CĐT đồng ý; vị trí tủ điều khiển scrubber theo nhà cung cấp.',
    'Cao độ máng +3,6 m trùng vùng ống gió Phase 2 (+3,3…+4,1 m) – phối hợp với WP3 tại giao cắt; tại khu có ống Phase 1 (+4,2…+6,2) giữ nguyên.',
    'Ổ cắm bảo trì và chiếu sáng sàn thiết bị trên trần: WP4 chỉ đưa ổ cắm tại 3 AHU; chiếu sáng bảo trì do gói điện (nếu cần).',
    'Kết nối HMI Phase 1 (phương án OP-1…OP-5): cần thông tin giao thức/hãng hệ điều khiển Phase 1 và vị trí màn hình HMI.',
    'ST1-CC2 (tủ hóa chất thứ 2 ở ST1) – nếu có hút, đã nằm trong van nhóm MD-G-ST1 (không đổi I/O); nếu tách van riêng +1 DO.',
]
R.append(''.join('%d. %s\n' % (i + 1, s) for i, s in enumerate(issues)))
R.append('\n## 13. Kiểm tra\n\n- Excel: %d ô lỗi sau tính lại LibreOffice; 0 mạch “KHÔNG ĐẠT” (In ≥ Ib, PE, Iz ≥ In, ΔU ≤ 4 %%).\n'
         '- Nhu cầu max Excel = %s kVA, Python = %s kVA (khớp).\n- Σ quạt tủ hood = 15 điểm (4 acid) khớp WP1/WP2; 6 module CDU khớp WP1.\n'
         '- Mọi I/O có vị trí → có chiều dài cáp; mọi mạch có cáp → có tuyến máng.\n' % (XCHK['errors'], vn(XCHK['demand_kVA'], 1), vn(dmax['kVA'], 1)))
open(os.path.join(HERE, 'WP4_report.md'), 'w', encoding='utf-8').write('\n'.join(R))

print('DB.HVAC: In=%d Ir=%d, demand %.1f kW / %.1f kVA / %.0f A (%s); connected rated %.1f kW, input %.1f kW' %
      (In_db, Ir_set, dmax['kW'], dmax['kVA'], Ib_db, dmax['case'], P_conn_rated, P_conn_in))
print('Incoming cable:', INCOMING['incoming_cable'])
for f in FEED:
    print(f['id'], round(f['Ib'], 1), f['In'], f['cable'], f['route']['L'], round(f['du'], 2))
print('IO', IO_SUM, IO_TOT)
print('Power cable', {k: round(v) for k, v in PWQ.items()})
print('Control cable', {k: round(v) for k, v in CTQ.items()})
print('Tray', {k: round(v, 1) for k, v in TRAY_Q.items()})
print('max dU', max(d['du_cum'] for d in LOADS if d['cable']))
