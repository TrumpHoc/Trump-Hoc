#!/usr/bin/env python3
"""WP2 - Mô hình hình học khu H&N Phase 2 (Nhà xưởng 3, SGS Hải Phòng).

Chạy:  python3 -I extract_geometry.py   (từ bất kỳ thư mục nào)
Đầu vào : ../00_inputs/261109_Cai_tao_khu_HN_xuong_3.pdf (trang 5, 6, 7)
Đầu ra  : geometry.json, plan_overlay_p7.png, plan_rooms_p6.png (cùng thư mục script)

Hệ tọa độ: mm, gốc tại giao trục E/3, X dương sang phải (về trục A), Y dương lên trên (về trục 5).

Phương pháp
1. Hiệu chuẩn: lấy các đường lưới (đường thẳng dài, nét đứt) và chữ ký hiệu trục trên trang 6
   bằng pdfplumber, hồi quy tuyến tính pt -> mm theo khoảng cách lưới 7000 / 7700.
2. Phòng: tọa độ tâm vách đo từ vector (trang 6), mặt trong vách lấy theo kích thước thông thủy
   ghi trên DW-A-03-04 (trang 5); kiểm tra lại: tâm vách đo được phải nằm giữa 2 mặt trong
   (sai số <= 60 mm) và nhãn tên phòng (trang 6) phải nằm trong đa giác.
3. Điểm hút: tâm ký hiệu thiết bị đo trên trang 7 (vector), mỗi điểm có "neo" là nhãn chữ trên bản vẽ,
   script kiểm tra chữ trong khung neo có chứa đủ ký tự mong đợi.
4. Thiết bị đề xuất, louver, điểm xả, khoảng cách Manhattan.
"""
import json, math, os, sys, unicodedata, collections

import pdfplumber
import pypdfium2 as pdfium

HERE = os.path.dirname(os.path.abspath(__file__))
PDF = os.path.join(HERE, '..', '00_inputs', '261109_Cai_tao_khu_HN_xuong_3.pdf')

# ----------------------------------------------------------------------------------------------
# 1. CALIBRATION
# ----------------------------------------------------------------------------------------------
GRID_X = {'E': 0, 'D': 7000, 'C': 14000, 'B': 21000, 'A': 28000}
GRID_Y = {'3': 0, '4': 7700, '5': 15400}


def linfit(xs, ys):
    n = len(xs); mx = sum(xs) / n; my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs); sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    a = sxy / sxx; b = my - a * mx
    res = [y - (a * x + b) for x, y in zip(xs, ys)]
    return a, b, res


def grid_lines(page):
    """Long thin vertical/horizontal lines (grey dashed grid) – total length per coordinate."""
    V = collections.defaultdict(float); H = collections.defaultdict(float)
    for l in page.lines:
        if l['x0'] > 1000:
            continue
        if abs(l['x0'] - l['x1']) < 0.05:
            V[round(l['x0'], 2)] += abs(l['bottom'] - l['top'])
        elif abs(l['top'] - l['bottom']) < 0.05:
            H[round(l['top'], 2)] += abs(l['x1'] - l['x0'])
    return V, H


def calibrate(pdf):
    out = {}
    p = pdf.pages[5]  # page 6
    V, H = grid_lines(p)
    # grid label text (big font 11.7) -> approximate positions, then snap to the nearest long line
    labels = {}
    for c in p.chars:
        if 11 < c['size'] < 12.5 and c['x0'] < 1000 and c['text'].strip():
            labels.setdefault(c['text'], []).append(((c['x0'] + c['x1']) / 2, (c['top'] + c['bottom']) / 2))
    xs_pt, xs_mm, ys_pt, ys_mm, detail = [], [], [], [], {}
    for g, mm in GRID_X.items():
        lx = labels[g][0][0]
        cand = [(abs(k - lx), k) for k, v in V.items() if v > 500 and abs(k - lx) < 3]
        k = min(cand)[1]
        xs_pt.append(k); xs_mm.append(mm); detail[g] = {'label_x_pt': round(lx, 2), 'line_x_pt': k}
    for g, mm in GRID_Y.items():
        ly = labels[g][0][1]
        cand = [(abs(k - ly), k) for k, v in H.items() if v > 500 and abs(k - ly) < 3]
        k = min(cand)[1]
        ys_pt.append(k); ys_mm.append(mm); detail[g] = {'label_y_pt': round(ly, 2), 'line_y_pt': k}
    ax, bx, rx = linfit(xs_pt, xs_mm)
    ay, by, ry = linfit(ys_pt, ys_mm)
    out['page'] = 6
    out['grid_detail_pt'] = detail
    out['sx_mm_per_pt'] = round(ax, 4)
    out['sy_mm_per_pt'] = round(-ay, 4)
    out['x_mm = a*x_pt + b'] = [round(ax, 5), round(bx, 2)]
    out['y_mm = a*top_pt + b'] = [round(ay, 5), round(by, 2)]
    out['residual_x_mm'] = [round(r, 1) for r in rx]
    out['residual_y_mm'] = [round(r, 1) for r in ry]
    out['grid_spacing_check_mm'] = {
        'E-D,D-C,C-B,B-A': [round((xs_pt[i + 1] - xs_pt[i]) * ax, 1) for i in range(4)],
        '3-4,4-5': [round((ys_pt[i + 1] - ys_pt[i]) * ay, 1) for i in range(2)]}
    out['implied_paper_scale'] = '1:%d (A3, bản vẽ không theo tỷ lệ chuẩn)' % round(ax / 0.35278)
    # same grid on page 5 and 7 ?
    same = {}
    for pg in (4, 6):
        V2, H2 = grid_lines(pdf.pages[pg])
        okx = all(any(abs(k - x) < 0.05 and v > 500 for k, v in V2.items()) for x in xs_pt)
        oky = all(any(abs(k - y) < 0.05 and v > 500 for k, v in H2.items()) for y in ys_pt)
        same['page%d' % (pg + 1)] = bool(okx and oky)
    out['same_grid_on_pages'] = same
    return out, (ax, bx, ay, by)


# ----------------------------------------------------------------------------------------------
PDF_OBJ = pdfplumber.open(PDF)
CAL, (AX, BX, AY, BY) = calibrate(PDF_OBJ)


def to_mm(x_pt, y_pt):
    return (AX * x_pt + BX, AY * y_pt + BY)


def to_pt(X, Y):
    return ((X - BX) / AX, (Y - BY) / AY)


def r0(v):
    return int(round(v))


# ----------------------------------------------------------------------------------------------
# 2. ROOMS
# ----------------------------------------------------------------------------------------------
# Wall centre lines measured from page-6 vectors (pt).  Name -> (axis, pt coordinate)
WALL_CL_PT = {
    'xL_left_block_outer': ('x', 151.3), 'x_dividing_coding': ('x', 207.5),
    'x_dining_digestion': ('x', 278.4), 'x_prep_balance_instr': ('x', 350.1),
    'x_digestion_kjeldahl': ('x', 376.5), 'x_instr_reporting': ('x', 451.4),
    'xR_left_block_outer': ('x', 503.4), 'xL_right_block_outer': ('x', 552.4),
    'x_std_std': ('x', 629.2), 'x_std_st2_office_st2': ('x', 678.2), 'xR_right_block_A_side': ('x', 856.3),
    'y_top': ('y', 57.4), 'y_kj_st1': ('y', 164.7), 'y_cst_office_st2': ('y', 296.5),
    'y_st1_prep_balance': ('y', 372.7), 'y_balance_instr': ('y', 459.2), 'y_office_std': ('y', 530.3),
    'y_prep_coding': ('y', 584.6), 'y_bottom': ('y', 711.8)}
# Inner faces (mm) – chained from DW-A-03-04 clear dimensions, panel 50 mm
T = 50
XF = {}  # x faces
XF['L0'] = 1095                     # inner face of left outer wall (centre 1070)
XF['div_R'] = XF['L0'] + 2100        # Sample dividing 2100
XF['cod_L'] = XF['div_R'] + T
XF['cod_R'] = XF['cod_L'] + 5305     # Sample coding 5305
XF['ins_L'] = XF['cod_R'] + T
XF['ins_R'] = XF['ins_L'] + 3785     # Instrument 3785
XF['rep_L'] = XF['ins_R'] + T
XF['rep_R'] = XF['rep_L'] + 1900     # Reporting 1900  -> 14335
XF['prep_R'] = XF['L0'] + 7455       # Sample prep 7455 -> 8550
XF['bal_L'] = XF['prep_R'] + T
XF['bal_R'] = XF['bal_L'] + 5735     # Balance 5735 -> 14335
XF['st1_R'] = XF['rep_R']
XF['st1_L'] = XF['st1_R'] - 8440     # ST1 8440 -> 5895
XF['dig_L'] = XF['st1_L']
XF['kj_R'] = XF['st1_R']
XF['kj_L'] = XF['kj_R'] - 4740       # Kjeldahl 4740
XF['dig_R'] = XF['kj_L'] - T         # Digestion width -> 3650
XF['RB_L'] = 16235                   # right block inner face (centre 16210)
XF['std1_R'] = XF['RB_L'] + 2900     # Standard 19.72: 2900
XF['std2_L'] = XF['std1_R'] + T
XF['std2_R'] = XF['std2_L'] + 1750   # Standard 11.90: 1750
XF['off_R'] = XF['RB_L'] + 4700      # Office 4701
XF['st2_L'] = XF['std2_R'] + T
XF['st2_R'] = XF['st2_L'] + 6720     # ST2 6720
YF = {}
YF['B0'] = -5975                     # inner face of bottom wall (centre -6000)
YF['cod_T'] = YF['B0'] + 4750
YF['prep_B'] = YF['cod_T'] + T
YF['prep_T'] = YF['prep_B'] + 7950
YF['ins_T'] = YF['B0'] + 9485
YF['bal_B'] = YF['ins_T'] + T
YF['bal_T'] = YF['bal_B'] + 3215
YF['st1_B'] = YF['bal_T'] + T
YF['st1_T'] = YF['st1_B'] + 7800
YF['kj_B'] = YF['st1_T'] + T
YF['kj_T'] = YF['kj_B'] + 4000       # -> 18675
YF['std_T'] = YF['B0'] + 6800
YF['off_B'] = YF['std_T'] + T
YF['off_T'] = YF['off_B'] + 8775
YF['st2_T'] = YF['B0'] + 15625

ROOMS_DEF = [
    # id, name, schedule m2, x0,y0,x1,y1, label page6 bbox(pt) for check, AHU zone, group
    ('DIG', 'Digestion room', 14.60, XF['dig_L'], YF['kj_B'], XF['dig_R'], YF['kj_T'], 'Hóa'),
    ('KJE', 'Kjeldahl room', 18.88, XF['kj_L'], YF['kj_B'], XF['kj_R'], YF['kj_T'], 'Hóa'),
    ('ST1', 'Sample treatment 1', 65.70, XF['st1_L'], YF['st1_B'], XF['st1_R'], YF['st1_T'], 'Hóa'),
    ('BAL', 'Balance room', 18.35, XF['bal_L'], YF['bal_B'], XF['bal_R'], YF['bal_T'], 'Thường'),
    ('PRE', 'Sample preparation', 58.99, XF['L0'], YF['prep_B'], XF['prep_R'], YF['prep_T'], 'Thường (có hút)'),
    ('COD', 'Sample coding', 25.20, XF['cod_L'], YF['B0'], XF['cod_R'], YF['cod_T'], 'Thường'),
    ('DIV', 'Sample dividing', 9.98, XF['L0'], YF['B0'], XF['div_R'], YF['cod_T'], 'Thường'),
    ('INS', 'Instrument room', 35.82, XF['ins_L'], YF['B0'], XF['ins_R'], YF['ins_T'], 'Thường (có hút)'),
    ('REP', 'Reporting room', 18.02, XF['rep_L'], YF['B0'], XF['rep_R'], YF['ins_T'], 'Thường'),
    ('OFF', 'Office', 40.96, XF['RB_L'], YF['off_B'], XF['off_R'], YF['off_T'], 'Thường'),
    ('STD20', 'Standard (19,72 m²)', 19.72, XF['RB_L'], YF['B0'], XF['std1_R'], YF['std_T'], 'Hóa'),
    ('STD12', 'Standard (11,90 m²)', 11.90, XF['std2_L'], YF['B0'], XF['std2_R'], YF['std_T'], 'Hóa'),
    ('ST2', 'Sample treatment 2', 103.64, XF['st2_L'], YF['B0'], XF['st2_R'], YF['st2_T'], 'Hóa'),
]
AHU_ZONE = {'DIG': 'AHU-2.01', 'KJE': 'AHU-2.01', 'ST1': 'AHU-2.01', 'BAL': 'AHU-2.01', 'PRE': 'AHU-2.01',
            'COD': 'AHU-2.01', 'DIV': 'AHU-2.01', 'REP': 'AHU-2.01', 'INS': 'AHU-2.03 (phương án) / AHU-2.01',
            'OFF': 'AHU-2.02', 'STD20': 'AHU-2.02', 'STD12': 'AHU-2.02', 'ST2': 'AHU-2.02'}
# Page-6 area label of each room (char multiset matched, vertical labels have scrambled char order)
LABEL_KEY = {'DIG': '14.60m2', 'KJE': '18.88m2', 'ST1': '65.70m2', 'BAL': '18.35m2', 'PRE': '58.99m2',
             'COD': '25.20m2', 'DIV': '9.98m2', 'INS': '35.82m2', 'REP': '18.02m2', 'OFF': '40.96m2',
             'ST2': '103.64m2', 'STD20': '19.72m2', 'STD12': '11.90m2'}

# Non-scope / surrounding zones (approx., mm)
ZONES = {
    'DINING_nonscope': {'name': 'Khu hiện hữu không ghi tên (trái trên, ~DIR nhà ăn trên DW-A-03-01) – ngoài phạm vi',
                        'rect': [XF['L0'], YF['st1_B'], XF['st1_L'] - T, YF['kj_T']]},
    'CST_nonscope': {'name': 'Khu hiện hữu không ghi tên (phải trên, ~CST kho hóa chất trên DW-A-03-01) – ngoài phạm vi',
                     'rect': [XF['RB_L'], YF['st2_T'] + T, XF['st2_R'], YF['kj_T']]},
    'CORRIDOR': {'name': 'Lối đi/không gian xưởng giữa khối trái (E–C) và khối phải (B–A), không điều hòa',
                 'rect': [XF['rep_R'] + T, -6000, XF['RB_L'] - T, 18700]},
}
X_WALL_A_INNER = 27900   # assumed inner face of external wall A (grid A = 28000)


def text_clusters(page):
    """Group chars into labels (used for verification only)."""
    ch = [c for c in page.chars if c['x0'] < 1000]
    n = len(ch); par = list(range(n))

    def f(i):
        while par[i] != i:
            par[i] = par[par[i]]; i = par[i]
        return i
    B = collections.defaultdict(list)
    for i, c in enumerate(ch):
        B[(int(c['x0'] // 5), int(c['top'] // 5))].append(i)
    for i, c in enumerate(ch):
        g = max(1.2 * min(c['size'], 8) * 0.4, 0.6)
        bx, by = int(c['x0'] // 5), int(c['top'] // 5)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in B[(bx + dx, by + dy)]:
                    if j <= i:
                        continue
                    d = ch[j]
                    if c['upright'] != d['upright'] or abs(c['size'] - d['size']) > 0.6 * max(c['size'], d['size']):
                        continue
                    if d['x0'] - c['x1'] < g and c['x0'] - d['x1'] < g and d['top'] - c['bottom'] < g and c['top'] - d['bottom'] < g:
                        par[f(i)] = f(j)
    G = collections.defaultdict(list)
    for i in range(n):
        G[f(i)].append(ch[i])
    out = []
    for g in G.values():
        s = ''.join(c['text'] for c in sorted(g, key=lambda c: (round(c['top']), c['x0'])))
        out.append({'x0': min(c['x0'] for c in g), 'x1': max(c['x1'] for c in g),
                    'top': min(c['top'] for c in g), 'bottom': max(c['bottom'] for c in g), 'text': s})
    return out


def poly_area(P):
    return 0.5 * abs(sum(P[i][0] * P[(i + 1) % len(P)][1] - P[(i + 1) % len(P)][0] * P[i][1] for i in range(len(P))))


def in_rect(x, y, r, tol=0):
    return r[0] - tol <= x <= r[2] + tol and r[1] - tol <= y <= r[3] + tol


def build_rooms():
    p6 = PDF_OBJ.pages[5]
    cl6 = text_clusters(p6)
    rooms = []
    for rid, name, sched, x0, y0, x1, y1, grp in ROOMS_DEF:
        P = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]
        A = poly_area(P) / 1e6
        # gross (to wall centre lines)
        Ag = (x1 - x0 + T) * (y1 - y0 + T) / 1e6
        # label check
        need = collections.Counter(LABEL_KEY[rid])
        lab = [c for c in cl6 if collections.Counter(c['text'].replace(' ', '')) == need]
        lab_ok = None; lab_mm = None
        if lab:
            c = lab[0]
            lx, ly = to_mm((c['x0'] + c['x1']) / 2, (c['top'] + c['bottom']) / 2)
            lab_mm = [r0(lx), r0(ly)]
            lab_ok = in_rect(lx, ly, [x0, y0, x1, y1])
        # independent check: faces from vector-measured wall centre lines (page 6) +/- 25 mm
        def vf(v, axis, sgn):
            cls = [to_mm(c, 0)[0] if a == 'x' else to_mm(0, c)[1] for a, c in WALL_CL_PT.values() if a == axis]
            cl = min(cls, key=lambda q: abs(q - (v - sgn * T / 2)))
            return cl + sgn * T / 2
        vx0, vx1 = vf(x0, 'x', +1), vf(x1, 'x', -1)
        vy0, vy1 = vf(y0, 'y', +1), vf(y1, 'y', -1)
        Av = (vx1 - vx0) * (vy1 - vy0) / 1e6
        rooms.append({'id': rid, 'name': name, 'group': grp, 'ahu_zone': AHU_ZONE[rid],
                      'area_vector_only_m2': round(Av, 2), 'area_vector_only_diff_pct': round((Av - sched) / sched * 100, 2),
                      'polygon_mm': [[r0(a), r0(b)] for a, b in P],
                      'clear_dims_mm': [r0(x1 - x0), r0(y1 - y0)],
                      'area_model_m2': round(A, 2), 'area_gross_wall_cl_m2': round(Ag, 2),
                      'area_schedule_m2': sched, 'area_diff_pct': round((A - sched) / sched * 100, 2),
                      'area_check_5pct': abs(A - sched) / sched <= 0.05,
                      'centroid_mm': [r0((x0 + x1) / 2), r0((y0 + y1) / 2)],
                      'bbox_mm': [r0(x0), r0(y0), r0(x1), r0(y1)],
                      'page6_label_mm': lab_mm, 'page6_label_inside': lab_ok,
                      'volume_m3_at_ceiling_3m': round(A * 3.0, 1)})
    return rooms


def wall_cl_check():
    """Compare measured wall centre lines (vectors) with the face model."""
    out = {}
    pairs = {'xL_left_block_outer': XF['L0'] - T / 2, 'x_dividing_coding': XF['div_R'] + T / 2,
             'x_dining_digestion': XF['st1_L'] - T / 2, 'x_prep_balance_instr': XF['prep_R'] + T / 2,
             'x_digestion_kjeldahl': XF['dig_R'] + T / 2, 'x_instr_reporting': XF['ins_R'] + T / 2,
             'xR_left_block_outer': XF['rep_R'] + T / 2, 'xL_right_block_outer': XF['RB_L'] - T / 2,
             'x_std_std': XF['std1_R'] + T / 2, 'x_std_st2_office_st2': XF['std2_R'] + T / 2,
             'xR_right_block_A_side': XF['st2_R'] + T / 2,
             'y_top': YF['kj_T'] + T / 2, 'y_kj_st1': YF['st1_T'] + T / 2, 'y_cst_office_st2': YF['st2_T'] + T / 2,
             'y_st1_prep_balance': YF['prep_T'] + T / 2, 'y_balance_instr': YF['ins_T'] + T / 2,
             'y_office_std': YF['std_T'] + T / 2, 'y_prep_coding': YF['cod_T'] + T / 2, 'y_bottom': YF['B0'] - T / 2}
    p6 = PDF_OBJ.pages[5]
    V, H = grid_lines(p6)
    for k, (ax, v) in WALL_CL_PT.items():
        meas = to_mm(v, 0)[0] if ax == 'x' else to_mm(0, v)[1]
        # presence of a wall line in the vectors near v
        D = V if ax == 'x' else H
        present = any(abs(kk - v) < 0.6 and L > 30 for kk, L in D.items())
        out[k] = {'measured_cl_mm': r0(meas), 'model_cl_mm': r0(pairs[k]), 'diff_mm': r0(meas - pairs[k]),
                  'vector_line_present': present}
    return out


# ----------------------------------------------------------------------------------------------
# Exposed walls
# ----------------------------------------------------------------------------------------------
def classify_outside(x, y):
    if x > XF['st2_R'] + 10:
        return 'external_wall_A'
    for zid, z in ZONES.items():
        if in_rect(x, y, z['rect'], tol=60):
            return zid
    return 'factory_unconditioned'


def interval_sub(a, b, cuts):
    segs = [(a, b)]
    for c0, c1 in cuts:
        new = []
        for s0, s1 in segs:
            if c1 <= s0 or c0 >= s1:
                new.append((s0, s1)); continue
            if c0 > s0:
                new.append((s0, c0))
            if c1 < s1:
                new.append((c1, s1))
        segs = new
    return [s for s in segs if s[1] - s[0] > 120]   # drop wall-thickness slivers at T-junctions


def exposed_walls(rooms):
    R = {r['id']: r['bbox_mm'] for r in rooms}
    res = {}
    for rid, (x0, y0, x1, y1) in R.items():
        edges = {'W': ('x', x0, y0, y1, -1), 'E': ('x', x1, y0, y1, +1),
                 'S': ('y', y0, x0, x1, -1), 'N': ('y', y1, x0, x1, +1)}
        segs_out = []; shared = []
        for side, (ax, c, a, b, sgn) in edges.items():
            cuts = []
            for oid, (u0, v0, u1, v1) in R.items():
                if oid == rid:
                    continue
                if ax == 'x':
                    oc = u1 if sgn < 0 else u0
                    if abs(abs(oc - c) - T) <= 30:
                        lo, hi = max(a, v0), min(b, v1)
                        if hi - lo > 1:
                            cuts.append((lo, hi)); shared.append({'side': side, 'with': oid, 'length_mm': r0(hi - lo)})
                else:
                    oc = v1 if sgn < 0 else v0
                    if abs(abs(oc - c) - T) <= 30:
                        lo, hi = max(a, u0), min(b, u1)
                        if hi - lo > 1:
                            cuts.append((lo, hi)); shared.append({'side': side, 'with': oid, 'length_mm': r0(hi - lo)})
            for s0, s1 in interval_sub(a, b, cuts):
                m = (s0 + s1) / 2
                px, py = (c + sgn * 400, m) if ax == 'x' else (m, c + sgn * 400)
                kind = classify_outside(px, py)
                segs_out.append({'side': side, 'from_mm': [r0(c), r0(s0)] if ax == 'x' else [r0(s0), r0(c)],
                                 'to_mm': [r0(c), r0(s1)] if ax == 'x' else [r0(s1), r0(c)],
                                 'length_mm': r0(s1 - s0), 'faces': kind})
        tot = collections.defaultdict(int)
        for s in segs_out:
            tot[s['faces']] += s['length_mm']
        res[rid] = {'shared_walls': shared, 'exposed_segments': segs_out,
                    'exposed_length_mm_by_type': dict(tot),
                    'exposed_wall_area_m2_by_type_h3m': {k: round(v * 3.0 / 1000, 2) for k, v in tot.items()}}
    return res


# ----------------------------------------------------------------------------------------------
# 3. EXHAUST POINTS
# ----------------------------------------------------------------------------------------------
# (id, room, type, acid, symbol centre pt (page7), facing, anchor bbox pt or None, expected letters, note, flow m3/h)
Q = {'hood': 2270, 'canopy': 1100, 'cabinet': 1300, 'sink_hood': 300, 'arm_hood': 800, 'instr_vent': 150}
EXH = [
    ('DIG-FH1', 'DIG', 'fume_hood_1600', True, (299.6, 68.0), 'S', (280.0, 66.8, 301.9, 74.8), 'HOODACID', 'Hood acid đặt trên bàn 1600×800 sát vách trên'),
    ('DIG-FH2', 'DIG', 'fume_hood_1600', True, (347.3, 68.0), 'S', (327.0, 66.8, 348.9, 74.8), 'HOODACID', 'Hood acid 1600 đặt trên bàn 2000×800 sát vách trên'),
    ('DIG-CH1', 'DIG', 'canopy', False, (348.6, 152.7), 'N', (326.1, 146.0, 349.4, 161.0), 'CHPHOODLÒNUNG', 'Chụp treo 2000×800 trên bàn lò nung'),
    ('KJE-CH1', 'KJE', 'canopy', False, (433.2, 152.7), 'N', (416.5, 146.5, 439.8, 158.6), 'CHPHOOD', 'Chụp treo 2000×800 trên bàn (máy Kjeldahl)'),
    ('ST1-FH1', 'ST1', 'fume_hood_1600', True, (289.0, 185.8), 'E', (283.0, 166.4, 291.0, 188.3), 'HOODACID', 'Hood acid, vách trái'),
    ('ST1-FH2', 'ST1', 'fume_hood_1600', False, (289.0, 228.3), 'E', (282.5, 216.7, 286.2, 241.1), 'HOOD', 'Hood hữu cơ (lọc than)'),
    ('ST1-FH3', 'ST1', 'fume_hood_1600', False, (289.0, 270.7), 'E', (282.1, 251.9, 285.8, 276.3), 'HOOD', 'Hood hữu cơ (lọc than)'),
    ('ST1-CC1', 'ST1', 'chemical_cabinet', False, (493.3, 180.6), 'W', (491.1, 172.3, 496.2, 190.0), 'CHEMICAL', 'Tủ hóa chất 2000×650 (mua mới), vách phải'),
    ('ST1-CC2', 'ST1', 'chemical_cabinet', False, (359.0, 362.7), 'N', (349.6, 360.5, 368.4, 365.5), 'CHEMICAL', 'Tủ hóa chất 2000×650 (mua mới), vách dưới – CẦN XÁC NHẬN có hút (cơ sở TK chỉ 1 tủ)'),
    ('ST2-FH1', 'ST2', 'fume_hood_1600', False, (729.8, 307.0), 'S', (715.9, 302.8, 743.7, 308.1), 'FUMEHOOD1600', 'Hàng 3 hood sát vách trên'),
    ('ST2-FH2', 'ST2', 'fume_hood_1600', False, (772.2, 307.0), 'S', (758.3, 302.8, 786.1, 308.1), 'FUMEHOOD1600', ''),
    ('ST2-FH3', 'ST2', 'fume_hood_1600', True, (814.7, 307.0), 'S', (800.8, 302.8, 828.6, 308.1), 'FUMEHOOD1600', 'GIẢ ĐỊNH là hood acid (bản vẽ không phân biệt; chọn hood gần trục A/scrubber nhất)'),
    ('ST2-FH4', 'ST2', 'fume_hood_1600', False, (689.4, 357.2), 'E', (684.5, 342.8, 689.8, 370.6), 'FUMEHOOD1600', 'Vách trái ST2'),
    ('ST2-FH5', 'ST2', 'fume_hood_1600', False, (689.4, 544.0), 'E', (684.5, 529.6, 689.8, 557.4), 'FUMEHOOD1600', 'Vách trái ST2'),
    ('ST2-FH6', 'ST2', 'fume_hood_1600', False, (689.4, 586.5), 'E', (684.5, 572.0, 689.8, 599.8), 'FUMEHOOD1600', 'Vách trái ST2'),
    ('ST2-SK1', 'ST2', 'sink_hood', False, (744.8, 376.0), 'E', (742.7, 368.2, 747.8, 385.0), 'BỒNRỬA', 'Bồn rửa đôi 1400×800 (đầu bàn đảo) – chụp hút'),
    ('ST2-SK2', 'ST2', 'sink_hood', False, (744.8, 508.2), 'E', (742.7, 500.6, 747.8, 517.5), 'BỒNRỬA', 'Bồn rửa đôi 1400×800 (đầu bàn đảo) – chụp hút'),
    ('ST2-CC1', 'ST2', 'chemical_cabinet', False, (845.5, 516.0), 'W', (839.5, 514.3, 851.8, 519.2), 'TỦHÓACHẤT', 'Tủ hóa chất 800×800, vách phải (trục A)'),
    ('ST2-CC2', 'ST2', 'chemical_cabinet', False, (845.5, 567.9), 'W', (839.5, 566.0, 851.8, 570.9), 'TỦHÓACHẤT', 'Tủ hóa chất 800×800, vách phải (trục A)'),
    ('STD20-FH1', 'STD20', 'fume_hood_1600', False, (618.0, 550.5), 'W', (617.7, 538.4, 623.0, 566.2), 'FUMEHOOD1600', 'Sát vách ngăn 2 phòng Standard'),
    ('STD20-FH2', 'STD20', 'fume_hood_1600', False, (618.0, 592.1), 'W', (617.7, 580.8, 623.0, 608.6), 'FUMEHOOD1600', ''),
    ('STD20-FH3', 'STD20', 'fume_hood_1600', False, (618.0, 688.3), 'W', (617.7, 676.2, 623.0, 704.0), 'FUMEHOOD1600', 'Góc dưới phải'),
    ('STD20-FH4', 'STD20', 'fume_hood_1600', False, (563.1, 688.3), 'E', (558.6, 674.6, 563.9, 702.4), 'FUMEHOOD1600', 'Góc dưới trái'),
    ('STD20-SK1', 'STD20', 'sink_hood', False, (563.0, 551.5), 'E', None, '', 'Bồn rửa đôi (ký hiệu, không nhãn) góc trên trái – chụp hút'),
    ('STD12-CC1', 'STD12', 'chemical_cabinet', False, (662.5, 694.6), 'W', (658.7, 687.3, 664.2, 700.5), 'TỦKHÔ', '"TỦ KHÔ" = tủ khô chứa hóa chất 1000×600 có hút (DM nội thất 9.6)'),
    ('PRE-AH1', 'PRE', 'arm_hood', False, (206.9, 491.5), '-', None, '', 'GIẢ ĐỊNH: arm hood treo trên đảo bàn trái (bản vẽ không thể hiện)'),
    ('PRE-AH2', 'PRE', 'arm_hood', False, (277.1, 491.5), '-', None, '', 'GIẢ ĐỊNH: arm hood treo trên đảo bàn phải (bản vẽ không thể hiện)'),
]
for i, (xx, lab) in enumerate([(380.7, 378.3), (420.4, 418.0)]):
    for j, (yy, ly) in enumerate([(525.0, 515.9), (578.0, 568.9), (631.1, 621.8), (684.2, 674.8)]):
        EXH.append(('INS-V%d' % (i * 4 + j + 1), 'INS', 'instr_vent', False, (xx, yy), '-',
                    (lab, ly, lab + 5.1, ly + 19.5), 'BÀNĐỂMÁY',
                    'Điểm hút thiết bị (snorkel/ống mềm) trên bàn để máy 2000×1000 (8 bàn = 8 điểm)'))
TYPE_KEY = {'fume_hood_1600': 'hood', 'canopy': 'canopy', 'chemical_cabinet': 'cabinet', 'sink_hood': 'sink_hood',
            'arm_hood': 'arm_hood', 'instr_vent': 'instr_vent'}


def norm(s):
    return [ch for ch in unicodedata.normalize('NFC', s.upper()) if not ch.isspace()]


def build_exhaust(rooms):
    p7 = PDF_OBJ.pages[6]
    R = {r['id']: r['bbox_mm'] for r in rooms}
    out = []
    for eid, rid, typ, acid, (xp, yp), face, anc, exp, note in EXH:
        X, Y = to_mm(xp, yp)
        ver = None
        if anc:
            x0, t0, x1, b0 = anc
            got = collections.Counter(ch for c in p7.chars if x0 - 0.8 <= (c['x0'] + c['x1']) / 2 <= x1 + 0.8
                                      and t0 - 0.8 <= (c['top'] + c['bottom']) / 2 <= b0 + 0.8 for ch in norm(c['text']))
            need = collections.Counter(norm(exp))
            ver = all(got[k] >= v for k, v in need.items())
        room_hit = [k for k, bb in R.items() if in_rect(X, Y, bb, tol=5)]
        disch = 'scrubber' if acid else ({'fume_hood_1600': 'thải thẳng (lọc than tại tủ)'}.get(typ, 'quạt EAF HVAC'))
        out.append({'id': eid, 'room': rid, 'type': typ, 'acid': acid, 'xy_mm': [r0(X), r0(Y)],
                    'facing': face, 'flow_m3h_basis': Q[TYPE_KEY[typ]], 'discharge_path': disch,
                    'pdf_page7_pt': [xp, yp], 'label_verified': ver, 'inside_room_polygon': rid in room_hit,
                    'note': note})
    return out


BASIS = {  # DESIGN_BASIS §3: hood(acid), canopy, cabinet, sink, other
    'DIG': {'fume_hood_1600': 2, 'acid': 2, 'canopy': 1}, 'KJE': {'canopy': 1},
    'ST1': {'fume_hood_1600': 3, 'acid': 1, 'chemical_cabinet': 1},
    'ST2': {'fume_hood_1600': 6, 'acid': 1, 'chemical_cabinet': 3, 'sink_hood': 3},
    'STD20': {'fume_hood_1600': 3}, 'STD12': {'fume_hood_1600': 1},
    'PRE': {'arm_hood': 2}, 'INS': {'instr_vent': 8}}


def reconcile(exh):
    cnt = collections.defaultdict(lambda: collections.Counter())
    for e in exh:
        cnt[e['room']][e['type']] += 1
        if e['acid']:
            cnt[e['room']]['acid'] += 1
    rows = []
    keys = ['fume_hood_1600', 'acid', 'canopy', 'chemical_cabinet', 'sink_hood', 'arm_hood', 'instr_vent']
    for rid in ['DIG', 'KJE', 'ST1', 'ST2', 'STD20', 'STD12', 'PRE', 'INS']:
        b = BASIS.get(rid, {})
        row = {'room': rid}
        for k in keys:
            d, bb = cnt[rid][k], b.get(k, 0)
            if d or bb:
                row[k] = {'drawing': d, 'basis': bb, 'match': d == bb}
        rows.append(row)
    tot = {}
    for k in keys:
        tot[k] = {'drawing': sum(cnt[r][k] for r in cnt), 'basis': sum(BASIS[r].get(k, 0) for r in BASIS)}
    grp = ['ST2', 'STD20', 'STD12']
    tot_grp = {k: {'drawing': sum(cnt[r][k] for r in grp), 'basis': sum(BASIS[r].get(k, 0) for r in grp)} for k in keys}
    return {'per_room': rows, 'total': tot, 'group_ST2_plus_2_Standard': tot_grp}


# ----------------------------------------------------------------------------------------------
# 4. DOORS (centre on wall, mm) – read from door-swing arcs on page 6 and labels on page 5
# ----------------------------------------------------------------------------------------------
def door(i, a, b, wall_axis, X, Y, w, kind, note=''):
    return {'id': i, 'between': [a, b], 'wall': wall_axis, 'center_mm': [r0(X), r0(Y)], 'width_mm': w,
            'kind': kind, 'note': note}


def build_doors():
    m = lambda xp, yp: to_mm(xp, yp)
    D = [
        door('D-01', 'DIG', 'KJE', 'x', 9571, m(0, 111.0)[1], 900, 'cánh đơn', 'Digestion chỉ đi qua Kjeldahl'),
        door('D-02', 'KJE', 'ST1', 'y', m(391.8, 0)[0], 14650, 1000, 'cánh đơn D4.5 (tận dụng)'),
        door('D-03', 'KJE', 'CORRIDOR', 'x', 14360, 16950, 1400, 'cửa đôi 1,4 m (giữ hiện trạng)'),
        door('D-04', 'ST1', 'CORRIDOR', 'x', 14360, m(0, 340.4)[1], 1400, 'cửa đôi lệch cánh (hiện trạng)'),
        door('D-05', 'BAL', 'CORRIDOR', 'x', 14360, m(0, 414.0)[1], 1400, 'cửa đôi lệch cánh (hiện trạng)'),
        door('D-06', 'REP', 'CORRIDOR', 'x', 14360, m(0, 484.5)[1], 1400, 'cửa đôi lệch cánh (hiện trạng)'),
        door('D-07', 'INS', 'REP', 'x', 12410, 2760, 1400, 'D_1.4x2.2 (làm mới)', 'Instrument vào qua Reporting'),
        door('D-08', 'BAL', 'PRE', 'x', 8575, m(0, 417.0)[1], 1000, 'cửa lùa D_1x2.2 (làm mới)'),
        door('D-09', 'PRE', 'FACTORY_W', 'x', 1070, m(0, 414.0)[1], 1400, 'cửa đôi lệch cánh (hiện trạng)'),
        door('D-10', 'PRE', 'COD', 'y', m(329.8, 0)[0], -1200, 1000, 'D_1x2.2 (tận dụng)'),
        door('D-11', 'COD', 'FACTORY_S', 'y', m(325.0, 0)[0], -6000, 1400, 'D_1.4x2.2 (làm mới)'),
        door('D-12', 'DIV', 'FACTORY_W', 'x', 1070, m(0, 640.8)[1], 1000, 'D_1x2.2 (tận dụng), đổi chiều mở'),
        door('PB-01', 'COD', 'DIV', 'x', 3220, m(0, 616.0)[1], 800, 'cửa inox 800×800 kiểu passbox, đáy cách sàn 0,9–1,0 m'),
        door('D-13', 'OFF', 'CORRIDOR', 'x', 16210, m(0, 415.7)[1], 1600, 'cửa đôi 1,6 m'),
        door('D-14', 'OFF', 'ST2', 'x', 20960, m(0, 488.0)[1], 1000, 'cửa lùa 1 m D_1x2.2 (làm mới)'),
        door('D-15', 'STD20', 'CORRIDOR', 'x', 16210, m(0, 635.0)[1], 1400, 'cửa đôi lệch cánh'),
        door('D-16', 'STD20', 'STD12', 'x', 19160, m(0, 635.0)[1], 1400, 'D_1.4x2.2 (làm mới)'),
        door('D-17', 'STD12', 'ST2', 'x', 20960, m(0, 635.0)[1], 1400, 'D_1.4x2.2 (làm mới)'),
        door('D-18', 'ST2', 'FACTORY_S', 'y', m(723.3, 0)[0], -6000, 1600, 'cửa đôi 1,6 m mở 2 chiều D_1.6x2.2 (làm mới)'),
        # non-scope / external
        door('D-N1', 'DINING_nonscope', 'FACTORY_W', 'x', 1070, m(0, 221.0)[1], 1400, 'hiện trạng', 'ngoài phạm vi'),
        door('D-N2', 'CST_nonscope', 'CORRIDOR', 'x', 16210, m(0, 266.7)[1], 1700, 'cửa đôi hiện trạng', 'ngoài phạm vi'),
        door('D-N3', 'CST_nonscope', 'FACTORY_N', 'y', m(826.0, 0)[0], 18700, 1700, 'cửa đôi hiện trạng', 'ngoài phạm vi'),
        door('D-N4', 'CST_nonscope', 'OUTSIDE_A', 'x', 28000, m(0, 189.2)[1], 900, 'cửa ra ngoài trên tường trục A', 'giữ lối ra – không đặt thiết bị trước cửa'),
        door('D-N5', 'ST2 (khe tường A)', 'OUTSIDE_A', 'x', 28000, -850, 900, 'EXISTING D3 (theo DW-A-03-01)', 'cần khảo sát: có còn sau cải tạo?'),
    ]
    return D


# ----------------------------------------------------------------------------------------------
# 5. EQUIPMENT PROPOSAL
# ----------------------------------------------------------------------------------------------
EQUIP = [
    # id, type, centre (X,Y), footprint [dx,dy], z-range/level, location note, serves
    ('AHU-2.01', 'AHU', (11700, 5150), (4800, 2000), '+3,3…+5,3 (sàn kết cấu trên trần)',
     'Trên trần Balance room, sát lối đi giữa 2 khối (trục C)', ['DIG', 'KJE', 'ST1', 'BAL', 'PRE', 'COD', 'DIV', 'REP', 'INS*']),
    ('AHU-2.02', 'AHU', (24600, -3300), (4800, 2000), '+3,3…+5,3',
     'Trên trần ST2 (khu "đặt máy tương lai"), cách tường A ~3 m', ['OFF', 'STD20', 'STD12', 'ST2']),
    ('AHU-2.03', 'AHU (phương án)', (10500, -3800), (3000, 1600), '+3,3…+5,0',
     'Trên trần Instrument room (nếu WP1 chọn tách AHU)', ['INS']),
    ('DH-2.01', 'Máy hút ẩm', (11900, -1200), (1200, 800), '+3,3… (dạng nối ống)', 'Trên trần Instrument, cạnh AHU-2.03', ['INS']),
    ('FAF-2.01', 'Quạt cấp bù FAFE (G4+F8)', (13000, 11000), (1500, 900), '+3,3…+4,3',
     'Trên trần ST1 sát lối đi', ['DIG', 'KJE', 'ST1']),
    ('FAF-2.02', 'Quạt cấp bù FAFE (G4+F8)', (24600, 400), (1500, 900), '+3,3…+4,3',
     'Trên trần ST2 gần AHU-2.02 và louver LV-2.02', ['ST2', 'STD20', 'STD12']),
    ('EAF-2.01', 'Quạt hút EAF (SS304)', (13000, 16000), (1200, 900), '+3,3…+4,3',
     'Trên trần Kjeldahl sát lối đi', ['DIG-CH1', 'KJE-CH1', 'ST1-CC1', 'ST1-CC2']),
    ('EAF-2.02', 'Quạt hút EAF (SS304)', (26500, 5000), (1200, 900), '+3,3…+4,3',
     'Trên trần ST2 phía tường A', ['ST2-SK1', 'ST2-SK2', 'ST2-CC1', 'ST2-CC2', 'STD20-SK1', 'STD12-CC1']),
    ('EAF-2.03', 'Quạt hút EAF (GI/SS)', (8575, 1500), (1200, 900), '+3,3…+4,3',
     'Trên vách Sample prep/Instrument', ['PRE-AH1', 'PRE-AH2'] + ['INS-V%d' % i for i in range(1, 9)]),
    ('SCR-2.01', 'Scrubber ướt (PP) – ngoài nhà', (30000, 16250), (2000, 2500), 'mặt đất ngoài nhà, ống xả ≥ đỉnh mái',
     'Ngoài tường trục A, phía bắc cửa CST (D-N4), cách FPS > 6 m', ['DIG-FH1', 'DIG-FH2', 'ST1-FH1', 'ST2-FH3']),
    ('CDU-YARD', 'Bãi dàn nóng CDU-2.01/2.02/(2.03)', (29500, -12000), (2000, 6000), 'mặt đất (móng: xây dựng)',
     'Ngoài tường A, phía nam bậc/ram dốc hiện hữu (Y -5,9…-7,6 m)', ['AHU-2.01', 'AHU-2.02', 'AHU-2.03']),
    ('DB.HVAC', 'Tủ điện tổng HVAC (gợi ý cho WP4)', (16400, 2200), (800, 400), 'sàn',
     'Vách khối phải phía lối đi (giữa D-13 và D-15)', []),
]
POINTS = [
    # louvers / discharge / drain on wall A (X = 28000)
    ('LV-2.01', 'louver_OA', (28000, -2400), '+4,0…+5,5', 'Gió tươi AHU-2.01 + FAF-2.01 (+AHU-2.03 nếu gộp)'),
    ('LV-2.02', 'louver_OA', (28000, -3700), '+4,0…+5,5', 'Gió tươi AHU-2.02 + FAF-2.02'),
    ('LV-2.03', 'louver_OA', (28000, -5000), '+4,0…+5,5', 'Gió tươi AHU-2.03 (phương án)'),
    ('DP-2.00', 'discharge', (30700, 17300), 'ống xả scrubber, đỉnh ≥ mái +3 m', 'Khí acid sau scrubber'),
    ('DP-2.01', 'discharge', (28000, 18200), 'xuyên tường A tại +4,5 → ống đứng vượt mái', 'Hood hữu cơ ST1 (2) + EAF-2.01'),
    ('DP-2.02', 'discharge', (28000, 12000), 'xuyên tường A tại +4,5 → ống đứng vượt mái', 'Hood hữu cơ ST2 (5) + Standard (4) + EAF-2.02'),
    ('DP-2.03', 'discharge', (28000, 9600), 'xuyên tường A tại +4,5 → ống đứng vượt mái', 'EAF-2.03 (arm hood + Instrument)'),
    ('CD-2.01', 'condensate_drain', (28100, -8500), 'ngoài nhà, cao độ đất', 'Điểm thoát nước ngưng giả định (rãnh/ga thoát nước mưa ngoài tường A)'),
]
OBSTACLES = [  # outside wall A (from DW-A-03-01 page 2, scale 130.4 mm/pt, verified grid)
    ('FPS', 'Phòng FUMI & PCO (kho hiện hữu ngoài tường A)', [28390, 640, 31010, 8650]),
    ('RAMP', 'Bậc/ram dốc hiện hữu ngoài tường A', [28000, -7600, 30000, -5900]),
]


def build_equipment():
    eq = []
    for i, t, c, fp, z, note, serves in EQUIP:
        eq.append({'id': i, 'type': t, 'center_mm': list(c), 'footprint_mm': list(fp),
                   'rect_mm': [c[0] - fp[0] // 2, c[1] - fp[1] // 2, c[0] + fp[0] // 2, c[1] + fp[1] // 2],
                   'level': z, 'location': note, 'serves': serves})
    pts = [{'id': i, 'type': t, 'xy_mm': list(c), 'level': z, 'note': n} for i, t, c, z, n in POINTS]
    obs = [{'id': i, 'name': n, 'rect_mm': r} for i, n, r in OBSTACLES]
    return eq, pts, obs


def manh(a, b):
    return abs(a[0] - b[0]) + abs(a[1] - b[1])


def build_distances(rooms, exh, eq, pts):
    C = {r['id']: r['centroid_mm'] for r in rooms}
    E = {e['id']: e['xy_mm'] for e in exh}
    Q_ = {e['id']: e['center_mm'] for e in eq}
    P = {p['id']: p['xy_mm'] for p in pts}
    ALL = {}
    ALL.update({('room:' + k): v for k, v in C.items()})
    ALL.update({('exh:' + k): v for k, v in E.items()})
    ALL.update({('pt:' + k): v for k, v in P.items()})
    matrix = {}
    for qid, qc in Q_.items():
        matrix[qid] = {k: r0(manh(qc, v)) for k, v in ALL.items()}
    # assigned runs (for WP3 take-off).  L_est = 1.10*Manhattan + vertical allowance
    runs = []

    def add(frm, to, sysname, vert=1500):
        a = Q_.get(frm) or P.get(frm) or E.get(frm) or C.get(frm)
        b = Q_.get(to) or P.get(to) or E.get(to) or C.get(to)
        L = manh(a, b)
        runs.append({'system': sysname, 'from': frm, 'to': to, 'L_manhattan_m': round(L / 1000, 1),
                     'L_est_m': round((1.10 * L + vert) / 1000, 1)})
    for r in ['DIG', 'KJE', 'ST1', 'BAL', 'PRE', 'COD', 'DIV', 'REP']:
        add('AHU-2.01', r, 'SA/RA AHU-2.01 (tới tâm phòng)')
    add('AHU-2.01', 'INS', 'SA/RA AHU-2.01 (nếu không tách AHU-2.03)')
    for r in ['OFF', 'STD20', 'STD12', 'ST2']:
        add('AHU-2.02', r, 'SA/RA AHU-2.02 (tới tâm phòng)')
    add('AHU-2.03', 'INS', 'SA/RA AHU-2.03')
    add('DH-2.01', 'INS', 'Máy hút ẩm')
    add('LV-2.01', 'AHU-2.01', 'OA', 500); add('LV-2.01', 'FAF-2.01', 'OA', 500)
    add('LV-2.02', 'AHU-2.02', 'OA', 500); add('LV-2.02', 'FAF-2.02', 'OA', 500)
    add('LV-2.03', 'AHU-2.03', 'OA', 500)
    for a in ['AHU-2.01', 'AHU-2.02', 'AHU-2.03']:
        add(a, 'CDU-YARD', 'Ống gas AHU→CDU', 4500)
        add(a, 'CD-2.01', 'Nước ngưng', 3500)
    add('DH-2.01', 'CD-2.01', 'Nước ngưng', 3500)
    for e in exh:
        if e['room'] in ('DIG', 'KJE', 'ST1'):
            add('FAF-2.01', e['id'], 'Gió bù FAFE')
        if e['room'] in ('ST2', 'STD20', 'STD12'):
            add('FAF-2.02', e['id'], 'Gió bù FAFE')
        if e['acid']:
            add(e['id'], 'SCR-2.01', 'Hút acid → scrubber (PP)', 2500)
        elif e['type'] == 'fume_hood_1600':
            add(e['id'], 'DP-2.01' if e['room'] == 'ST1' else 'DP-2.02', 'Hood hữu cơ → xả (SS304)', 2500)
    for q in eq:
        if q['id'].startswith('EAF'):
            for s in q['serves']:
                add(s, q['id'], 'Hút EAF nhánh', 2000)
    add('EAF-2.01', 'DP-2.01', 'Hút EAF → xả', 1500)
    add('EAF-2.02', 'DP-2.02', 'Hút EAF → xả', 1500)
    add('EAF-2.03', 'DP-2.03', 'Hút EAF → xả', 1500)
    # separation intake ↔ exhaust
    sep = []
    for lv in [p for p in pts if p['type'] == 'louver_OA']:
        for dp in [p for p in pts if p['type'] == 'discharge']:
            d = math.dist(lv['xy_mm'], dp['xy_mm'])
            sep.append({'louver': lv['id'], 'discharge': dp['id'], 'dist_m': round(d / 1000, 1), 'ok_ge_10m': d >= 10000})
    return matrix, runs, sep


# ----------------------------------------------------------------------------------------------
# 6. PLOTS
# ----------------------------------------------------------------------------------------------
def render_bg(page_index, scale=2):
    doc = pdfium.PdfDocument(PDF)
    img = doc[page_index].render(scale=scale).to_pil().convert('L').convert('RGB')
    W, H = img.size
    # extent in mm: pixel (0,0) = pt (0,0)
    x0, y0 = to_mm(0, 0); x1, y1 = to_mm(W / scale, H / scale)
    return img, [x0, x1, y1, y0]


def plot_all(rooms, exh, eq, pts, obs, doors, expo):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle, Polygon
    from PIL import ImageEnhance
    colors = {'fume_hood_1600': '#d62728', 'canopy': '#ff7f0e', 'chemical_cabinet': '#9467bd',
              'sink_hood': '#17becf', 'arm_hood': '#8c564b', 'instr_vent': '#2ca02c'}
    # ---------- overlay on page 7 (all) ----------
    img, ext = render_bg(6)
    img = ImageEnhance.Brightness(img).enhance(1.25)
    fig, ax = plt.subplots(figsize=(22, 17))
    ax.imshow(img, extent=ext, alpha=0.45)
    for r in rooms:
        ax.add_patch(Polygon(r['polygon_mm'], closed=True, fill=True, fc='#1f77b4', alpha=0.07, ec='#1f77b4', lw=1.6))
        cx, cy = r['centroid_mm']
        ax.text(cx, cy, '%s\n%.2f m²' % (r['id'], r['area_model_m2']), ha='center', va='center', fontsize=8,
                color='#08306b', weight='bold', bbox=dict(fc='white', ec='none', alpha=0.6, pad=1))
    for zid, z in ZONES.items():
        x0, y0, x1, y1 = z['rect']
        ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fill=False, hatch='//', ec='#999999', lw=0.8))
        ax.text((x0 + x1) / 2, (y0 + y1) / 2, {'DINING_nonscope': 'NGOÀI PHẠM VI\n(trái trên)',
                                                'CST_nonscope': 'NGOÀI PHẠM VI\n(phải trên, ~CST)',
                                                'CORRIDOR': 'LỐI ĐI\nXƯỞNG'}[zid],
                ha='center', va='center', fontsize=8, color='#555555', rotation=90 if zid == 'CORRIDOR' else 0)
    for e in exh:
        x, y = e['xy_mm']
        mk = '^' if e['acid'] else 'o'
        ax.plot(x, y, mk, ms=9, mfc=colors[e['type']], mec='k', mew=0.8, zorder=5)
        ax.text(x + 150, y + 150, e['id'], fontsize=6, color=colors[e['type']], zorder=6)
    for q in eq:
        x0, y0, x1, y1 = q['rect_mm']
        col = {'AHU': '#0b6e4f', 'AHU (phương án)': '#0b6e4f'}.get(q['type'], '#b8860b')
        if q['id'].startswith('SCR'):
            col = '#c00000'
        if q['id'].startswith('CDU'):
            col = '#555599'
        if q['id'].startswith('DB'):
            col = '#000000'
        ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fc=col, alpha=0.35, ec=col, lw=1.5,
                               ls='--' if 'phương án' in q['type'] else '-', zorder=4))
        ax.text(q['center_mm'][0], q['center_mm'][1], q['id'], ha='center', va='center', fontsize=7, weight='bold', zorder=6)
    for p in pts:
        x, y = p['xy_mm']
        st = {'louver_OA': ('s', '#1f77b4'), 'discharge': ('X', '#c00000'), 'condensate_drain': ('D', '#17becf')}[p['type']]
        ax.plot(x, y, st[0], ms=11, color=st[1], mec='k', zorder=7)
        ax.text(x + 300, y, p['id'], fontsize=7, color=st[1], va='center', weight='bold', zorder=7)
    for o in obs:
        x0, y0, x1, y1 = o['rect_mm']
        ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fill=False, ec='k', hatch='xx', lw=1))
        ax.text((x0 + x1) / 2, (y0 + y1) / 2, o['id'], ha='center', va='center', fontsize=8, rotation=90)
    for d in doors:
        x, y = d['center_mm']
        ax.plot(x, y, 'P', ms=6, color='#2ca02c', zorder=6)
        ax.text(x, y - 380, d['id'], fontsize=5, color='#2ca02c', ha='center')
    # intake / exhaust separation circle
    for lv in [p for p in pts if p['id'] == 'LV-2.01']:
        ax.add_patch(plt.Circle(lv['xy_mm'], 10000, fill=False, ls=':', ec='#1f77b4'))
    ax.axvline(28000, color='k', lw=2.5)
    ax.text(28150, 19500, 'TƯỜNG NGOÀI TRỤC A', fontsize=9, weight='bold')
    import matplotlib.lines as mlines
    hs = [mlines.Line2D([], [], marker='o', ls='', mfc=c, mec='k', label=k) for k, c in colors.items()]
    hs.append(mlines.Line2D([], [], marker='^', ls='', mfc='#d62728', mec='k', label='hood ACID'))
    hs += [mlines.Line2D([], [], marker='s', ls='', color='#1f77b4', label='louver gió tươi'),
           mlines.Line2D([], [], marker='X', ls='', color='#c00000', label='điểm xả'),
           mlines.Line2D([], [], marker='D', ls='', color='#17becf', label='thoát nước ngưng'),
           mlines.Line2D([], [], marker='P', ls='', color='#2ca02c', label='cửa')]
    ax.legend(handles=hs, loc='lower left', fontsize=8, framealpha=0.9)
    ax.set_xlim(-2500, 34500); ax.set_ylim(-15800, 21000)
    ax.set_aspect('equal'); ax.grid(True, ls=':', lw=0.4)
    for g, X in GRID_X.items():
        ax.text(X, 20500, g, ha='center', fontsize=12, weight='bold', color='m')
    for g, Y in GRID_Y.items():
        ax.text(-2200, Y, g, va='center', fontsize=12, weight='bold', color='m')
    ax.set_title('WP2 – Mô hình hình học H&N Phase 2 trên nền DW-A-03-06 (mm, gốc E/3)', fontsize=14)
    fig.savefig(os.path.join(HERE, 'plan_overlay_p7.png'), dpi=110, bbox_inches='tight')
    plt.close(fig)

    # ---------- rooms + exposed walls on page 6 ----------
    img, ext = render_bg(5)
    fig, ax = plt.subplots(figsize=(20, 15))
    ax.imshow(img, extent=ext, alpha=0.5)
    ecol = {'external_wall_A': '#c00000', 'factory_unconditioned': '#ff7f0e', 'CORRIDOR': '#e6a800',
            'DINING_nonscope': '#8c564b', 'CST_nonscope': '#9467bd'}
    for r in rooms:
        ax.add_patch(Polygon(r['polygon_mm'], closed=True, fc='#1f77b4', alpha=0.08, ec='#1f77b4', lw=1))
        cx, cy = r['centroid_mm']
        ax.text(cx, cy, '%s\nmô hình %.2f / BV %.2f m²\n(%+.1f%%)' % (r['id'], r['area_model_m2'], r['area_schedule_m2'], r['area_diff_pct']),
                ha='center', va='center', fontsize=7, bbox=dict(fc='white', ec='none', alpha=0.7))
        for s in expo[r['id']]['exposed_segments']:
            (xa, ya), (xb, yb) = s['from_mm'], s['to_mm']
            ax.plot([xa, xb], [ya, yb], color=ecol[s['faces']], lw=4, alpha=0.8)
    import matplotlib.lines as mlines
    ax.legend(handles=[mlines.Line2D([], [], color=c, lw=4, label=k) for k, c in ecol.items()], loc='lower left', fontsize=9)
    ax.set_xlim(-1500, 31000); ax.set_ylim(-8000, 20500); ax.set_aspect('equal')
    ax.set_title('WP2 – Đa giác phòng (mặt trong vách) & tường tiếp xúc không gian không điều hòa (nền DW-A-03-05)', fontsize=13)
    fig.savefig(os.path.join(HERE, 'plan_rooms_p6.png'), dpi=100, bbox_inches='tight')
    plt.close(fig)


# ----------------------------------------------------------------------------------------------
def main():
    rooms = build_rooms()
    wcl = wall_cl_check()
    expo = exposed_walls(rooms)
    for r in rooms:
        r['exposure'] = expo[r['id']]
    exh = build_exhaust(rooms)
    rec = reconcile(exh)
    doors = build_doors()
    eq, pts, obs = build_equipment()
    matrix, runs, sep = build_distances(rooms, exh, eq, pts)
    G = {
        'meta': {'project': 'SGS Hải Phòng – H&N Phase 2 (Nhà xưởng 3)', 'wp': 'WP2 geometry', 'date': '2026-11-14',
                 'source_pdf': '00_inputs/261109_Cai_tao_khu_HN_xuong_3.pdf (trang 5 DW-A-03-04, 6 DW-A-03-05, 7 DW-A-03-06, 2 DW-A-03-01)',
                 'coords': 'mm; gốc giao trục E/3; X → trục A (phải); Y → trục 5 (lên)',
                 'levels': {'FFL': 0, 'ceiling': 3000, 'equipment_zone_above_ceiling': [3300, 6000],
                            'phase1_ducts_BOD': [4200, 6200], 'roof': 'chưa biết'},
                 'wall_panel_mm': T},
        'calibration': CAL,
        'wall_centreline_check': wcl,
        'rooms': rooms,
        'rooms_total_m2': {'model': round(sum(r['area_model_m2'] for r in rooms), 2),
                           'schedule': round(sum(r['area_schedule_m2'] for r in rooms), 2)},
        'zones_nonscope': ZONES,
        'external_wall_A': {'x_mm': 28000, 'room_side_inner_face_x_mm': XF['st2_R'],
                            'note': 'Vách panel ST2 đặt cách tường A ~0,3 m; tường A là tường bao ngoài nhà'},
        'exhaust_points': exh,
        'exhaust_reconciliation': rec,
        'doors': doors,
        'equipment_proposed': eq,
        'louvers_discharge_drain': pts,
        'outdoor_obstacles': obs,
        'intake_exhaust_separation': sep,
        'runs_for_takeoff': runs,
        'distance_matrix_manhattan_mm': matrix,
    }
    with open(os.path.join(HERE, 'geometry.json'), 'w', encoding='utf-8') as f:
        json.dump(G, f, ensure_ascii=False, indent=1)
    plot_all(rooms, exh, eq, pts, obs, doors, expo)
    # console summary
    print('scale', CAL['sx_mm_per_pt'], CAL['sy_mm_per_pt'], 'resid', CAL['residual_x_mm'], CAL['residual_y_mm'], CAL['same_grid_on_pages'])
    for r in rooms:
        print('%-6s %7.2f %7.2f %+6.2f%% vec %7.2f %+6.2f%% label_in=%s' % (r['id'], r['area_model_m2'], r['area_schedule_m2'], r['area_diff_pct'], r['area_vector_only_m2'], r['area_vector_only_diff_pct'], r['page6_label_inside']))
    print('wall cl', {k: v['diff_mm'] for k, v in wcl.items()})
    print('labels verified', sum(1 for e in exh if e['label_verified']), '/', sum(1 for e in exh if e['label_verified'] is not None),
          'inside', all(e['inside_room_polygon'] for e in exh))
    bad = [e['id'] for e in exh if e['label_verified'] is False or not e['inside_room_polygon']]
    print('BAD', bad)
    print(json.dumps(rec['total'], ensure_ascii=False))
    for s in sep:
        print(s)


if __name__ == '__main__':
    main()
