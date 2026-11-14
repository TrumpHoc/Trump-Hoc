import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
import os
OUT=os.environ.get('OUT','/tmp/claude-0/-home-user-Trump-Hoc/51dcfdd4-dd01-598e-b524-4827d19677dc/scratchpad/HVAC_SGS_HN_Phase2_TinhTai.xlsx')
wb=openpyxl.Workbook()
B=Font(bold=True); INP=PatternFill('solid',fgColor='FFF2CC'); HDR=PatternFill('solid',fgColor='D9E1F2'); TOT=PatternFill('solid',fgColor='E2EFDA')
th=Side(style='thin',color='999999'); BD=Border(left=th,right=th,top=th,bottom=th)
wrap=Alignment(wrap_text=True,vertical='center')

# ---------- Sheet 1: design basis ----------
s=wb.active; s.title='1_CoSo'
s['A1']='CƠ SỞ THIẾT KẾ – HVAC CẢI TẠO KHU H&N PHASE 2 (XƯỞNG 3, SGS HẢI PHÒNG)'; s['A1'].font=Font(bold=True,size=13)
s['A2']='Ô vàng = số liệu đầu vào (sửa được). Ô trắng = công thức.'
rows=[
('P','Áp suất khí quyển',101.325,'kPa','Mực nước biển'),
('To','Ngoài trời mùa hè – nhiệt độ khô',37.8,'°C','Theo Phase 1 (F3-SD-AC-004)'),
('RHo','Ngoài trời mùa hè – độ ẩm',0.628,'-','Theo Phase 1 (F3-SD-AC-004)'),
('Tw','Ngoài trời mùa đông – nhiệt độ khô',9.8,'°C','Theo Phase 1 (F3-SD-AC-004)'),
('RHw','Ngoài trời mùa đông – độ ẩm',0.71,'-','Theo Phase 1'),
('Tc','Không khí sau coil',15,'°C','Theo Phase 1 (F3-SD-AC-004): 15°C/97%'),
('RHc','Độ ẩm sau coil',0.97,'-','Theo Phase 1'),
('Tsup','Nhiệt độ gió cấp vào phòng',16,'°C','Bằng nhiệt độ sau coil + ~1K nhiệt quạt'),
('rho','Khối lượng riêng không khí',1.2,'kg/m³',''),
('cp','Nhiệt dung riêng không khí',1.006,'kJ/kg.K',''),
('Tsur','Nhiệt độ không gian xưởng bao quanh (không ĐH)',35,'°C','Giả định – cần đo/xác nhận'),
('Tpl','Nhiệt độ plenum trên trần panel (dưới mái tôn)',45,'°C','Giả định'),
('Uw','U vách panel PU/EPS 50mm',0.45,'W/m².K','λ≈0.022 W/m.K + màng'),
('Uc','U trần panel PU 50mm',0.42,'W/m².K',''),
('H','Chiều cao thông thủy phòng',3.0,'m','Vách panel cao 3m (bản vẽ DW-A-03-04) – cần xác nhận cao độ trần'),
('LPD','Mật độ chiếu sáng',12,'W/m²','LED, ~500 lux phòng thí nghiệm'),
('qps','Nhiệt hiện / người',75,'W','Công việc nhẹ đứng/ngồi, 24°C (ASHRAE Fund. Ch.18)'),
('qpl','Nhiệt ẩn / người',55,'W',''),
('OAp','Gió tươi tối thiểu / người',25,'m³/h.người','TCVN 5687:2024 (văn phòng/PTN) – kiểm tra lại bản hiện hành'),
('FH','Lưu lượng hút tủ hood (CAV)',2270,'m³/h/tủ','Theo IFC tư vấn DW-AC-00-103/106: FUME HOOD 2,270 CMH/set (≈1,300 CFM; DW-AC-01-311)'),
('CAN','Lưu lượng chụp hút treo (lò nung Digestion, Kjeldahl)',1100,'m³/h/chụp','Theo CĐT – file thiết bị, sheet HSE mục 2.2–2.3: ~1,100 m³/h (IFC MIN lab: canopy 2000x750 = 2,200)'),
('CAB','Lưu lượng hút tủ hóa chất',1300,'m³/h/tủ','Theo IFC tư vấn DW-AC-00-103/01-311: CHEMICAL CABINET 1200x450x1800 – 1,300 CMH/set (cần xác nhận cho tủ Phase 2)'),
('SNK','Lưu lượng chụp hút bồn rửa',300,'m³/h/chụp','Giả định – IFC không có hạng mục tương ứng (cánh tay hút ARM HOOD = 800 CMH)'),
('MU','Tỷ lệ gió bù hút được xử lý lạnh (phòng Hệ 1)',0.0,'-','1.0 = FAHU xử lý 100%; 0 = cấp bù không xử lý (quạt FAFE + lọc G4/F8, cấp trước chụp) như tư vấn làm ở MIN lab/Phase 1'),
('SPL','Tỷ lệ gió bù không xử lý tràn vào vùng phòng',0.0,'-','Theo IFC/Phase 1: AHU không tính tải gió bù (=0). Xem độ nhạy 10–30% ở sheet 4_KichBan'),
('OBW','Hệ số gió bù FAFE / gió thải',1.0,'-','Phase 1: FAFE 3,100 / EAFE 3,050 CMH ≈ 1.0'),
('QB','Số lượng thiết bị tính tải: 1 = khi mua đủ (SL 2027 + bổ sung); 0 = chỉ SL 2027',1,'-','Theo file thiết bị CĐT HAI_PHONG_LAB_30102027 (sheet Equipment (Power))'),
('DIV','Hệ số đồng thời tủ hood',1.0,'-','1.0 = CAV chạy đồng thời; 0.6–0.7 nếu VAV + cảm biến sash'),
('SF','Hệ số an toàn tải phòng',0.10,'-',''),
]
s.append([]); s.append(['Ký hiệu','Thông số','Giá trị','Đơn vị','Nguồn / ghi chú'])
for c in s[4]: c.font=B; c.fill=HDR
name={}
import json as _j,os as _o
OVR=_j.loads(_o.environ.get('OVR','{}'))
for i,(k,t,v,u,n) in enumerate(rows):
    r=5+i; s.cell(r,1,k); s.cell(r,2,t); c=s.cell(r,3,OVR.get(k,v)); c.fill=INP; s.cell(r,4,u); s.cell(r,5,n); name[k]=f"'1_CoSo'!$C${r}"
r=5+len(rows)+1
s.cell(r,1,'TRẠNG THÁI KHÔNG KHÍ (tính tự động)').font=B; r+=1
s.cell(r,1,'Điểm');s.cell(r,2,'T (°C)');s.cell(r,3,'RH');s.cell(r,4,'Psat (kPa)');s.cell(r,5,'W (kg/kg)');s.cell(r,6,'h (kJ/kg)')
for c in s[r]: c.font=B; c.fill=HDR
def state(rr,label,T,RH):
    s.cell(rr,1,label); s.cell(rr,2,f'={T}'); s.cell(rr,3,f'={RH}')
    s.cell(rr,4,f'=0.61094*EXP(17.625*B{rr}/(B{rr}+243.04))')
    s.cell(rr,5,f"=0.622*C{rr}*D{rr}/({name['P']}-C{rr}*D{rr})")
    s.cell(rr,6,f'=1.006*B{rr}+E{rr}*(2501+1.86*B{rr})')
r+=1; state(r,'Ngoài trời hè',name['To'],name['RHo']); name['Wo']=f"'1_CoSo'!$E${r}"; name['ho']=f"'1_CoSo'!$F${r}"
r+=1; state(r,'Sau coil',name['Tc'],name['RHc']); name['Wc']=f"'1_CoSo'!$E${r}"; name['hc']=f"'1_CoSo'!$F${r}"
r+=1; state(r,'Ngoài trời đông',name['Tw'],name['RHw'])
r+=2
notes=['GHI CHÚ:',
'- Điều kiện ngoài trời 37.8°C/62.8% lấy đúng như Phase 1 để đồng bộ. Mức này khá khắt khe (h≈106 kJ/kg). Nếu CĐT/TVGS cho dùng số liệu theo QCVN 02:2022/BXD (trạm Phù Liễn) thì tải gió tươi giảm đáng kể – cần chốt.',
'- Công thức Psat theo Magnus (sai số <0.4% trong dải 0–50°C).',
'- Tải gió tươi tính từ trạng thái ngoài trời về trạng thái sau coil (đã gồm cả hiện và ẩn).']
for n in notes: s.cell(r,1,n); r+=1
for col,w in zip('ABCDEF',[12,48,12,14,80,12]): s.column_dimensions[col].width=w

# ---------- Sheet 2: room loads ----------
t=wb.create_sheet('2_TaiPhong')
t['A1']='TÍNH TẢI LẠNH & LƯU LƯỢNG GIÓ TỪNG PHÒNG – PHASE 2'; t['A1'].font=Font(bold=True,size=13)
t['A2']='Diện tích theo DW-A-03-05. Số tủ hood/chụp theo DW-A-03-06 + file nội thất 10.11.26. Lưu lượng hút/thiết bị theo IFC tư vấn (DW-AC-00-103/106, 01-305/311). Ô vàng = cần xác nhận.'
hdr=['STT','Phòng','Hệ','Diện tích (m²)','Cao (m)','Thể tích (m³)','T đặt (°C)','RH đặt','W phòng (kg/kg)','h phòng (kJ/kg)',
'Số người (CĐT)','Nhiệt thiết bị vào phòng (kW) – từ sheet 2b','Dài vách tiếp giáp KV không ĐH (m)',
'Q vách (W)','Q trần (W)','Q đèn (W)','Q người hiện (W)','Q thiết bị (W)','RSH tổng (W)','RLH (W)',
'Tủ hood 1600 (cái)','Chụp treo (cái)','Tủ hóa chất có hút (cái)','Chụp bồn rửa (cái)','Gió thải (m³/h)','ACH thải',
'Gió tươi theo người (m³/h)','Gió tươi qua coil (m³/h)','Gió cấp cần cho tải hiện (m³/h)','Gió cấp thiết kế (m³/h)','ACH cấp','Gió hồi tuần hoàn (m³/h)',
'Tải coil – gió tươi (kW)','Tải coil – gió hồi (kW)','Tổng tải lạnh coil (kW)','Gia nhiệt lại (kW)','Kiểm tra khử ẩm','kW/m²','BTU/h','Gió bù FAFE không xử lý (m³/h)','Khối AHU','Hood acid (cái)','Hút → Scrubber (m³/h)','Hút → quạt EAFE (m³/h)']
t.append([]); t.append(hdr)
for c in t[4]: c.font=B; c.fill=HDR; c.alignment=wrap; c.border=BD
t.row_dimensions[4].height=75
rooms=[
# name, sys, area, T, RH, ppl, eq(unused), wall, FH, CAN, CAB, SNK, AHU, acidFH
('Digestion room',1,14.60,25,0.60,2,0,3.7,2,1,0,0,'AHU-2.01',2),
('Kjeldahl room',1,18.88,25,0.60,2,0,8.8,0,1,0,0,'AHU-2.01',0),
('Sample treatment 1',1,65.70,25,0.60,10,0,15.8,3,0,1,0,'AHU-2.01',1),
('Sample treatment 2',1,103.64,25,0.60,10,0,29.3,6,0,3,3,'AHU-2.02',1),
('Standard (19.72 m²)',1,19.72,25,0.60,2,0,9.9,3,0,0,0,'AHU-2.02',0),
('Standard (11.90 m²)',1,11.90,25,0.60,2,0,1.7,1,0,0,0,'AHU-2.02',0),
('Balance room',2,18.35,25,0.60,4,0,3.3,0,0,0,0,'AHU-2.01',0),
('Sample preparation',2,58.99,25,0.60,10,0,8.0,0,0,0,0,'AHU-2.01',0),
('Sample coding',2,25.20,25,0.60,2,0,5.3,0,0,0,0,'AHU-2.01',0),
('Sample dividing',2,9.98,25,0.60,1,0,7.0,0,0,0,0,'AHU-2.01',0),
('Instrument room',2,35.82,25,0.55,10,0,3.9,0,0,0,0,'AHU-2.01',0),
('Reporting room',2,18.02,25,0.60,10,0,11.6,0,0,0,0,'AHU-2.01',0),
('Office',2,40.96,25,0.60,10,0,13.7,0,0,0,0,'AHU-2.02',0),
]
r0=5
N=name
for i,(nm,sy,A,T,RH,pp,eq,wl,fh,can,cab,snk,ahu,acid) in enumerate(rooms):
    r=r0+i; R=str(r)
    vals={1:i+1,2:nm,3:sy,4:A,7:T,8:RH,11:pp,13:wl,21:fh,22:can,23:cab,24:snk,41:ahu,42:acid}
    for k,v in vals.items(): t.cell(r,k,v)
    for k in (3,7,8,11,13,21,22,23,24,41,42): t.cell(r,k).fill=INP
    f={12:f"=SUMIF('2b_ThietBi'!$B$5:$B$200,B{R},'2b_ThietBi'!$K$5:$K$200)/1000",5:f"={N['H']}",6:f"=D{R}*E{R}",
    9:f"=0.622*H{R}*0.61094*EXP(17.625*G{R}/(G{R}+243.04))/({N['P']}-H{R}*0.61094*EXP(17.625*G{R}/(G{R}+243.04)))",
    10:f"=1.006*G{R}+I{R}*(2501+1.86*G{R})",
    14:f"=M{R}*E{R}*{N['Uw']}*({N['Tsur']}-G{R})",15:f"=D{R}*{N['Uc']}*({N['Tpl']}-G{R})",16:f"=D{R}*{N['LPD']}",
    17:f"=K{R}*{N['qps']}",18:f"=L{R}*1000",19:f"=(N{R}+O{R}+P{R}+Q{R}+R{R}+IF(C{R}=1,(1-{N['MU']})*{N['SPL']}*Y{R}*{N['rho']}*{N['cp']}*({N['To']}-G{R})/3.6,0))*(1+{N['SF']})",20:f"=(K{R}*{N['qpl']}+IF(C{R}=1,(1-{N['MU']})*{N['SPL']}*Y{R}*{N['rho']}/3.6*({N['Wo']}-I{R})*2501,0))*(1+{N['SF']})",
    25:f"=(U{R}*{N['FH']}+V{R}*{N['CAN']}+W{R}*{N['CAB']}+X{R}*{N['SNK']})*{N['DIV']}",26:f"=Y{R}/F{R}",
    27:f"=K{R}*{N['OAp']}",28:f"=IF(C{R}=1,MAX(Y{R}*{N['MU']},AA{R}),MAX(Y{R},AA{R}))",40:f"=IF(C{R}=1,Y{R}*(1-{N['MU']})*{N['OBW']},0)",43:f"=AP{R}*{N['FH']}*{N['DIV']}",44:f"=Y{R}-AQ{R}",
    29:f"=S{R}/({N['rho']}*{N['cp']}*1000*(G{R}-{N['Tsup']}))*3600",
    30:f"=MAX(AC{R},T{R}/({N['rho']}/3600*(I{R}-{N['Wc']})*2501*1000),AB{R},6*F{R})",31:f"=AD{R}/F{R}",32:f"=AD{R}-AB{R}",
    33:f"=AB{R}*{N['rho']}/3600*({N['ho']}-{N['hc']})",34:f"=AF{R}*{N['rho']}/3600*(J{R}-{N['hc']})",35:f"=AG{R}+AH{R}",
    36:f"=MAX(0,AD{R}*{N['rho']}*{N['cp']}*(G{R}-{N['Tsup']})/3600-S{R}/1000)",
    37:f'=IF(AD{R}*{N["rho"]}/3600*(I{R}-{N["Wc"]})*2501*1000>=T{R},"OK","KHÔNG ĐẠT")',
    38:f"=AI{R}/D{R}",39:f"=AI{R}*3412.14"}
    for k,v in f.items(): t.cell(r,k,v)
rl=r0+len(rooms)-1; rt=rl+1
t.cell(rt,2,'TỔNG').font=B
for k in (4,6,11,12,14,15,16,17,18,19,20,21,22,23,24,25,27,28,29,30,32,33,34,35,36,39,40,42,43,44):
    t.cell(rt,k,f'=SUM({L(k)}{r0}:{L(k)}{rl})').font=B
t.cell(rt,38,f'=AI{rt}/D{rt}').font=B
for c in t[rt]: c.fill=TOT
fmt={12:'0.0',4:'0.00',6:'0.0',8:'0%',9:'0.00000',10:'0.0',14:'#,##0',15:'#,##0',16:'#,##0',17:'#,##0',18:'#,##0',19:'#,##0',20:'#,##0',25:'#,##0',26:'0',27:'#,##0',28:'#,##0',29:'#,##0',30:'#,##0',31:'0',32:'#,##0',33:'0.0',34:'0.0',35:'0.0',36:'0.0',38:'0.00',39:'#,##0',40:'#,##0',43:'#,##0',44:'#,##0'}
for r in range(r0,rt+1):
    for k,fm in fmt.items(): t.cell(r,k).number_format=fm
    for k in range(1,45): t.cell(r,k).border=BD
t.column_dimensions['B'].width=22
for k in range(3,45): t.column_dimensions[L(k)].width=11
t.freeze_panes='C5'
r=rt+2
for n in ['Hệ 1 = phòng hóa có tủ hood/chụp hút; Hệ 2 = phòng không hút độc. Theo IFC/Phase 1: gió bù hút = quạt FAFE (lọc G4/F8, không làm lạnh) cấp trước tủ/chụp; AHU DX hồi gió xử lý tải phòng + gió tươi cho người.',
'Gió cấp thiết kế = MAX(gió cần cho tải hiện; gió cần cho tải ẩn; gió tươi qua coil; 6 ACH tối thiểu).',
'Gia nhiệt lại: khi gió bù hút lớn hơn gió cần cho tải hiện, gió 14°C làm lạnh quá mức → cần sưởi lại (điện/hot-gas) hoặc cấp gió ấm hơn.',
'Kiểm tra khử ẩm: so khả năng khử ẩm của gió cấp (W phòng – W sau coil) với tải ẩn phòng.',
'Vách tiếp giáp: chiều dài đo sơ bộ trên DW-A-03-05 (tỉ lệ ~22.5mm/px), coi không gian xưởng bao quanh là không điều hòa. Cạnh trục A của ST2 là tường ngoài nhà.']:
    t.cell(r,2,n); r+=1

# ---------- Sheet 2b: equipment ----------
e=wb.create_sheet('2b_ThietBi')
e['A1']='DANH MỤC THIẾT BỊ & NHIỆT TỎA VÀO PHÒNG'; e['A1'].font=Font(bold=True,size=13)
e['A2']='Nguồn: file CĐT "HAI_PHONG_LAB_30102027 - Equipment.xlsx", sheet Equipment (Power). Phòng lấy theo cột K (vị trí); riêng dòng "Tủ hút/chụp hút" lấy theo cột B. Kho khí (máy nén, máy sấy khí) không thuộc phòng điều hòa – không tính.'
e['A3']='Hệ số tỏa nhiệt (ô vàng) = giả định Tiến Đạt: gộp hệ số sử dụng & phần nhiệt tủ hút/chụp mang đi (tham khảo ASHRAE Fundamentals Ch.18). Cần CĐT xác nhận chế độ sử dụng.'
e.append(['Dòng CĐT','Phòng','Thiết bị','Công suất (W/cái)','SL 2027','SL bổ sung (khi mua đủ)','SL tính tải','Công suất tổng (W)','Nhóm','Hệ số tỏa nhiệt vào phòng','Nhiệt vào phòng (W)','Khí tươi (CĐT)','Khí hút (CĐT)'])
for c in e[4]: c.font=B; c.fill=HDR; c.alignment=wrap; c.border=BD
e.row_dimensions[4].height=45
import json as _js
EQ=_js.load(open('/tmp/claude-0/equip.json'))
EQ+= [['—','Standard (11.90 m²)','Tủ lạnh (3 cái theo mặt bằng DW-A-03-06) – giả định, không có trong danh mục CĐT',485,3,0,'','Tủ lạnh/đông, chiller',0.5,'',''],
      ['—','Standard (19.72 m²)','Thiết bị nhỏ trên bàn – giả định, không có trong danh mục CĐT',1000,1,0,'','Khác',0.5,'','']]
for k,(rw,room,dev,w,q1,q2,ph,cat,fct,o,pp) in enumerate(EQ):
    r=5+k; R=str(r)
    for col,v in zip('ABCDEFIJLM',[rw,room,dev,w,q1,q2,cat,fct,o,pp]): e[col+R]=v
    e['G'+R]=f"=E{R}+{N['QB']}*F{R}"; e['H'+R]=f'=D{R}*G{R}'; e['K'+R]=f'=H{R}*J{R}'
    e['J'+R].fill=INP
    for col in 'ABCDEFGHIJKLM': e[col+R].border=BD
    for col,fm in (('D','#,##0'),('H','#,##0'),('K','#,##0'),('J','0.00')): e[col+R].number_format=fm
re_=5+len(EQ)
e[f'C{re_}']='TỔNG'; e[f'C{re_}'].font=B
for col in 'HK': e[f'{col}{re_}']=f'=SUM({col}5:{col}{re_-1})'; e[f'{col}{re_}'].font=B; e[f'{col}{re_}'].number_format='#,##0'
r2=re_+2
e[f'B{r2}']='TỔNG THEO PHÒNG'; e[f'B{r2}'].font=B
e.cell(r2+1,2,'Phòng'); e.cell(r2+1,3,'Công suất điện (kW)'); e.cell(r2+1,4,'Nhiệt vào phòng (kW)')
for c in e[r2+1][1:4]: c.font=B; c.fill=HDR
for k,rm in enumerate([x[0] for x in rooms]):
    rr=r2+2+k; e.cell(rr,2,rm)
    e.cell(rr,3,f'=SUMIF($B$5:$B${re_-1},B{rr},$H$5:$H${re_-1})/1000').number_format='0.0'
    e.cell(rr,4,f'=SUMIF($B$5:$B${re_-1},B{rr},$K$5:$K${re_-1})/1000').number_format='0.0'
for col,w in zip('ABCDEFGHIJKLM',[8,22,48,11,8,11,9,12,30,11,12,9,9]): e.column_dimensions[col].width=w
e.freeze_panes='C5'

# ---------- Sheet 3: system summary ----------
p=wb.create_sheet('3_TongHop')
p['A1']='TỔNG HỢP THEO KHỐI AHU & HỆ HÚT – PHƯƠNG ÁN THEO IFC/PHASE 1'; p['A1'].font=Font(bold=True,size=13)
TP="'2_TaiPhong'!"
rng=lambda col: f"{TP}${col}${r0}:${col}${rl}"
p.append([]); p.append(['Hạng mục','AHU-2.01 (khối trái: Digestion…Reporting)','AHU-2.02 (khối phải: Office, ST2, Standard)','Tổng','Đơn vị'])
for c in p[3]: c.font=B; c.fill=HDR; c.alignment=wrap
items=[('Diện tích phục vụ','D','m²','#,##0.0'),('Gió cấp AHU','AD','m³/h','#,##0'),('Gió tươi qua AHU (cho người)','AB','m³/h','#,##0'),('Gió hồi','AF','m³/h','#,##0'),
('Tải lạnh coil – gió tươi','AG','kW','#,##0.0'),('Tải lạnh coil – gió hồi','AH','kW','#,##0.0'),('TỔNG TẢI LẠNH AHU','AI','kW','#,##0.0'),('Gia nhiệt lại','AJ','kW','#,##0.0'),
('Gió thải hút độc tổng','Y','m³/h','#,##0'),('  – về Scrubber (hood acid)','AQ','m³/h','#,##0'),('  – về quạt hút EAFE (hood hữu cơ có lọc than, chụp, tủ hóa chất, bồn rửa)','AR','m³/h','#,##0'),('Gió bù FAFE (lọc G4/F8, không xử lý lạnh)','AN','m³/h','#,##0')]
for i,(lab,col,u,fm) in enumerate(items):
    r=4+i; p.cell(r,1,lab)
    p.cell(r,2,f'=SUMIF({rng("AO")},"AHU-2.01",{rng(col)})'); p.cell(r,3,f'=SUMIF({rng("AO")},"AHU-2.02",{rng(col)})'); p.cell(r,4,f'=B{r}+C{r}'); p.cell(r,5,u)
    for k in (2,3,4): p.cell(r,k).number_format=fm
    if col=='AI':
        for k in range(1,6): p.cell(r,k).font=B; p.cell(r,k).fill=TOT
r=4+len(items)
p.cell(r,1,'Tổng tải lạnh AHU (TR)'); [p.cell(r,k,f'={L(k)}10/3.517') for k in (2,3,4)]; p.cell(r,5,'TR'); r+=1
p.cell(r,1,'Tổng tải lạnh AHU (BTU/h)'); [p.cell(r,k,f'={L(k)}10*3412.14') for k in (2,3,4)]; p.cell(r,5,'BTU/h'); r+=1
p.cell(r,1,'Tải lạnh chọn máy (+10% dự phòng)'); [p.cell(r,k,f'={L(k)}10*1.1') for k in (2,3,4)]; p.cell(r,5,'kW'); r+=1
p.cell(r,1,'ACH cấp trung bình'); [p.cell(r,k,f"={L(k)}5/({L(k)}4*{N['H']})") for k in (2,3,4)]; p.cell(r,5,'lần/h'); r+=1
for rr in range(16,r):
    for k in (2,3,4): p.cell(rr,k).number_format='#,##0.0'
r+=1
p.cell(r,1,'SO SÁNH VỚI PHASE 1 (F3-SD-AC-004, IFC DW-AC-00-109)').font=B; r+=1
for a,b in [('HN-AHU-1.01','12,459 m³/h @ESP 800 Pa; gió tươi 1,200 m³/h; 90.05 kW'),('HN-AHU-1.02','12,010 m³/h @ESP 800 Pa; gió tươi 2,000 m³/h; 91.21 kW'),
('Autoclave (phòng rửa)','EAFE 3,050 m³/h ×2 + FAFE bù 3,100 m³/h ×2 (in-line, lọc G4/F8, không xử lý lạnh)'),
('Phase 1 tổng','≈181 kW / ≈387 m² (≈0.47 kW/m²; phòng sạch ISO 7/8, 20–40 ACH)')]:
    p.cell(r,1,a); p.cell(r,2,b); r+=1
r+=1
p.cell(r,1,'SO SÁNH VỚI BẢNG MÁY HÚT ẨM (ảnh gửi kèm)').font=B; r+=1
p.cell(r,1,'PA Chia: 3×HD-45BE + 4×HD-60B + 2×HD-150B + 1×HD-192B + 2×HD-360B + 1×HD-504B'); p.cell(r,4,'=3*45+4*60+2*150+192+2*360+504'); p.cell(r,5,'lít/ngày (nếu số model = lít/ngày)'); r+=1
p.cell(r,1,'PA Gộp: 2×HCD-360B + 3×HCD-504B'); p.cell(r,4,'=2*360+3*504'); p.cell(r,5,'lít/ngày'); r+=1
p.cell(r,1,'Ẩm do gió bù FAFE nếu 10% tràn vào phòng (hè)'); p.cell(r,4,f"=D15*0.1*{N['rho']}*({N['Wo']}-0.0119)*24"); p.cell(r,5,'lít/ngày'); p.cell(r,4).number_format='#,##0'; r+=1
p.cell(r,1,'→ Máy hút ẩm có thể dùng làm giải pháp bổ trợ cho phòng hóa khi gió bù FAFE tràn vào phòng (mùa nồm), không thay AHU.')
p.column_dimensions['A'].width=62; [p.column_dimensions[c].__setattr__('width',28) for c in 'BCD']; p.column_dimensions['E'].width=30

# ---------- Sheet 4: sensitivity ----------
q=wb.create_sheet('4_KichBan')
q['A1']='KỊCH BẢN GIÓ BÙ – kết quả tính lại toàn bộ bảng với từng bộ thông số (giá trị tĩnh)'; q['A1'].font=Font(bold=True,size=13)
q['A2']='Muốn xem chi tiết kịch bản nào: đổi ô DIV / MU / SPL ở sheet 1_CoSo, các sheet 2–3 tự tính lại.'
q.append([]); q.append(['Kịch bản','DIV','MU','SPL','Gió thải (m³/h)','Gió tươi qua AHU (m³/h)','Gió bù FAFE (m³/h)','Tổng tải lạnh AHU (kW)','(TR)','Gia nhiệt lại (kW)','Ghi chú'])
for c in q[4]: c.font=B; c.fill=HDR; c.alignment=wrap
import json,os
if os.path.exists('/tmp/claude-0/scen.json'):
    for row in json.load(open('/tmp/claude-0/scen.json')): q.append(row)
for r in q.iter_rows(min_row=5):
    for c,fm in zip(r,[None,'0%','0%','0%','#,##0','#,##0','#,##0','#,##0','#,##0','#,##0',None]):
        if fm: c.number_format=fm
q.column_dimensions['A'].width=48; q.column_dimensions['K'].width=70
for c in 'BCDEFGHIJ': q.column_dimensions[c].width=13
# ---------- Sheet 5: open items ----------
o=wb.create_sheet('5_CanXacNhan')
o.append(['STT','Nội dung cần CĐT/NCC xác nhận','Ảnh hưởng'])
for c in o[1]: c.font=B; c.fill=HDR
items=[
('Lưu lượng hút: đang lấy theo IFC MIN lab (hood 2,270; chụp 2,200; tủ hóa chất 1,300 m³/h). IFC KHÔNG có bản vẽ cho phòng Phase 2 → cần tư vấn xác nhận áp dụng cho tủ hood 1600 & tủ hóa chất Phase 2 (theo vận tốc 0.3 m/s, cửa 700mm: hood 1600 chỉ ≈1,100 m³/h)','Quyết định >90% tải lạnh'),
('Phương án gió bù: xử lý trung tâm (FAHU) hay cấp bù không xử lý bằng quạt FAFE như tư vấn làm ở MIN lab – xem sheet 4_KichBan','Tải lạnh 700–1,100 kW và khả năng giữ RH'),
('Hood acid (Digestion, ST1, ST2) có đấu về scrubber như MIN lab (MI-SCR) không? Công suất scrubber hiện hữu còn dư?','Hệ hút, cột áp quạt'),
('Số chụp hút treo: Digestion & Kjeldahl – bản vẽ 1 chụp/phòng, file nội thất ghi 2 chụp/phòng','±2,700 m³/h/chụp ≈ ±60 kW'),
('Phân bổ 10 tủ hood khu bên phải giữa Sample treatment 2 và 2 phòng Standard','Phân hệ, ống gió'),
('Công suất điện thiết bị: lò nung (2 CB 3 pha), máy chưng cất đạm, tủ sấy, thiết bị Instrument room (GC/HPLC/AAS/ICP?), tủ lạnh/tủ đông','Tải hiện phòng'),
('Instrument room có thiết bị cần hút cục bộ (snorkel AAS/ICP) không?','Gió thải, gió bù'),
('Điểm thiết kế: đề xuất 24±1°C / RH ≤ 55–60% (Instrument 23°C / ≤ 55%) trong dải 18–25°C, 30–70% của CĐT','Điểm sau coil, tải ẩn'),
('Điều kiện ngoài trời: giữ 37.8°C/62.8% như Phase 1 hay dùng QCVN 02:2022/BXD','Tải gió tươi ±20–30%'),
('Nguồn lạnh: hệ độc lập (DX AHU/FAHU như Phase 1) hay đấu chung; công suất còn dư của ODU Phase 1 (nếu có)','Phương án hệ'),
('Bảng máy hút ẩm HD/HCD: xác nhận hãng/model, ý nghĩa cột BTU (tải phòng hay nhiệt máy thải ra)','Nhiệt máy hút ẩm thải ra phòng phải cộng vào tải ĐH'),
('Cao độ trần panel thực tế và không gian trên trần để đi ống gió/đặt FAHU','Bố trí thiết bị'),
('Áp suất phòng: các phòng hóa âm áp so với hành lang/phòng lân cận (như Phase 1 dùng ±15 Pa)?','Cân bằng gió'),
]
for i,(a,b) in enumerate(items): o.append([i+1,a,b])
o.column_dimensions['B'].width=110; o.column_dimensions['C'].width=45
for row in o.iter_rows(min_row=2):
    for c in row: c.alignment=wrap
wb.save(OUT); print(OUT)
