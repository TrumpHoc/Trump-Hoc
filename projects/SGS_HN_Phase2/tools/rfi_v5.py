import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
OUT='/tmp/claude-0/-home-user-Trump-Hoc/51dcfdd4-dd01-598e-b524-4827d19677dc/scratchpad/RFI_HVAC_SGS_HN_Phase2.xlsx'
F='Arial'
def f(b=False,sz=10,c='000000',i=False): return Font(name=F,bold=b,size=sz,color=c,italic=i)
HDR=PatternFill('solid',fgColor='1F4E78'); GRP=PatternFill('solid',fgColor='D9E1F2'); FILL=PatternFill('solid',fgColor='FFF2CC')
th=Side(style='thin',color='808080'); BD=Border(left=th,right=th,top=th,bottom=th)
W=Alignment(wrap_text=True,vertical='top'); C=Alignment(horizontal='center',vertical='top',wrap_text=True)
CDT='File thiết bị CĐT "HAI_PHONG_LAB_30102027 - Equipment" '
wb=openpyxl.Workbook(); ws=wb.active; ws.title='RFI'; ws.sheet_view.showGridLines=False
NC=6
HR=1
heads=['STT','Nội dung cần làm rõ','Căn cứ / tài liệu tham chiếu','Ảnh hưởng nếu chưa làm rõ','Giả định tạm thời của GGV','Phản hồi của CĐT']
for k,h in enumerate(heads,1):
    c=ws.cell(HR,k,h); c.font=f(True,10,'FFFFFF'); c.fill=HDR; c.alignment=C; c.border=BD
ws.row_dimensions[HR].height=32
G=[('I','CƠ SỞ THIẾT KẾ & PHẠM VI',[
('Xác nhận phạm vi HVAC Phase 2 gồm 13 phòng, tổng diện tích 441,76 m² theo DW-A-03-05: Digestion, Kjeldahl, Sample treatment 1 & 2, Standard ×2, Balance, Sample preparation, Sample coding, Sample dividing, Instrument, Reporting, Office. Phạm vi có bao gồm hút khí độc, scrubber không?',
 'DW-A-03-05, DW-A-03-06','Xác định khối lượng & phạm vi chào giá','Bao gồm: điều hòa, thông gió, hút khí độc (đến đầu chờ tủ hood/chụp), scrubber'),
('Điều kiện trong phòng: CĐT đã xác nhận dải 18–25°C, RH 30–70% (Instrument RH <60%). Đề nghị chốt điểm đặt (đề xuất 25°C như Phase 1) và cho biết có yêu cầu sưởi để giữ ≥18°C, tạo ẩm để giữ RH ≥30% vào mùa đông không.',
 CDT+'sheet Equipment (Power) cột L–M; DW-A-03-05; Phase 1 F3-SD-AC-004','Có hay không sưởi/tạo ẩm, điểm chọn coil','25°C; RH 40–70% (Instrument ≤55%); không tạo ẩm; có sưởi điện'),
('Điều kiện thiết kế ngoài trời: đề nghị xác nhận áp dụng như Phase 1 (hè 37,8°C / 62,8%; đông 9,8°C / 71%).',
 'Phase 1 F3-SD-AC-004 / IFC DW-AC-00-109','Tải gió tươi, chọn máy','Áp dụng như Phase 1'),
('Phase 2 có yêu cầu cấp sạch (ISO) hoặc số lần trao đổi gió tối thiểu không? Yêu cầu cấp lọc (G4/F8 như Phase 1?).',
 'Phase 1: ISO 7 (~35 ACH), ISO 8 (~20 ACH)','Lưu lượng gió, cột áp, cấp lọc','Không yêu cầu ISO; tối thiểu 6 ACH; lọc G4 + F8'),
]),
('II','TỦ HOOD, CHỤP HÚT & THIẾT BỊ PHÒNG THÍ NGHIỆM',[
('Lưu lượng hút thiết kế của tủ hood 1600 Phase 2. IFC MIN lab dùng 2.270 m³/h/bộ (≈1.300 CFM); theo ghi chú IFC "cửa cao 700 mm, vận tốc qua cửa 0,3 m/s" thì tủ 1600 chỉ cần khoảng 1.100 m³/h. Đề nghị cung cấp data sheet tủ hood: lưu lượng, cột áp, đường kính đầu ra, chiều cao mở cửa.',
 'IFC DW-AC-00-103/106, DW-AC-01-305/311; Danh mục nội thất 10.11.26 (mục 5.7, 6.4, 9.1, 9.12: vận tốc 0,3–0,5 m/s)','Quyết định lưu lượng hút, gió bù, kích thước quạt & ống; sai lệch tới khoảng 2 lần','2.270 m³/h/tủ theo IFC'),
('Tủ hood lưu lượng cố định (CAV) hay thay đổi (VAV)? Cảm biến cửa kính (sash) chỉ để báo động hay có điều khiển lưu lượng? Có van VAV phía hút/cấp không?',
 'Danh mục nội thất mục 9.1, 9.12: "cảm biến khi kéo quá vạch làm việc sẽ báo động"','Lưu lượng đồng thời, điều khiển','CAV, cảm biến sash chỉ báo động'),
('Số tủ hút/chụp hút khu Sample treatment 2: file thiết bị CĐT ghi 13 bộ (quạt 600 W, 3 pha); mặt bằng nội thất thể hiện 10 tủ hood (9 hữu cơ + 1 acid) + 3 tủ hóa chất có hút + 3 chụp bồn rửa = 16 điểm hút (gồm cả 2 phòng Standard). Đề nghị xác nhận danh sách điểm hút từng phòng và vị trí tủ acid.',
 CDT+'dòng 42; DW-A-03-06; Danh mục nội thất 9.1, 9.5–9.7, 9.12','Phân vùng hút, scrubber, cân bằng gió từng phòng','ST2: 6 hood (1 acid) + 3 tủ hóa chất + 3 chụp bồn rửa; Standard 19,72 m²: 3 hood; Standard 11,90 m²: 1 hood'),
('Chụp hút: theo HSE, chụp lò nung Digestion và hút Kjeldahl ~1.100 m³/h; danh mục thiết bị ghi Digestion 3 tủ/chụp (= 2 hood acid + 1 chụp lò nung), Kjeldahl 1. Đề nghị xác nhận: (a) Kjeldahl chỉ 1 chụp (danh mục nội thất mục 7.4 ghi 2); (b) kích thước, cao độ chụp; (c) lưu lượng 1.100 m³/h là cho mỗi chụp.',
 CDT+'sheet HSE mục 2.2–2.3; dòng 8, 23; Danh mục nội thất 6.5, 7.4','±1.100 m³/h mỗi chụp','Digestion: 2 hood acid + 1 chụp lò nung 1.100 m³/h; Kjeldahl: 1 chụp 1.100 m³/h'),
('Tủ hóa chất có hút (ST1: 1 tủ 1200×650×2020; ST2: 1 tủ 1000×600×2200 + 2 tủ 800×800×2200): lưu lượng hút? IFC MIN lab áp 1.300 m³/h cho tủ 1200×450×1800, khá lớn so với tủ hóa chất thông thường.',
 'Danh mục nội thất 5.12, 9.6, 9.7; IFC DW-AC-00-103, 01-311','Tổng 4 tủ: 400–5.200 m³/h','1.300 m³/h/tủ theo IFC'),
('Chụp hút bồn rửa dụng cụ ST2 (3 bộ – mục 9.5): lưu lượng, kích thước đầu nối, có đấu vào hệ hút chung không?',
 'Danh mục nội thất 9.5','Lưu lượng hút, ống','300 m³/h/bộ, đấu hệ hút chung'),
('Chế độ sử dụng thiết bị gia nhiệt (để tính nhiệt tỏa vào phòng): Sample treatment 1 có 20 bếp hồng ngoại × 2 kW, 25 bếp đun bình cầu × 350 W, 3 tủ sấy × 3,4 kW, 6 bể ổn nhiệt lạnh × 1,8 kW (tổng ~94 kW khi mua đủ). Đề nghị cho biết: số thiết bị chạy đồng thời, thiết bị nào đặt trong tủ hút; lò nung Digestion (2 × 4,5 kW) chạy liên tục hay theo mẻ.',
 CDT+'dòng 7–11, 46–66','Nhiệt tỏa vào phòng Sample treatment 1 – phòng tải lớn nhất cùng Instrument','Bếp, Soxhlet, Digestion system, lò nung: 20% nhiệt vào phòng; tủ sấy/bể ổn nhiệt: 50%; động cơ: 30%'),
('Instrument room: công suất thiết bị khi mua đủ ~52 kW (LC-MS/MS 3 × 5,5 kW; HS-GC-MS/MS 3 × 6,1 kW; HPAEC 2 × 3,9 kW; GC-FID 2 × 2,4 kW; UPLC 2 × 1,5 kW; máy sinh N₂/H₂). Đề nghị: (a) xác nhận số lượng giai đoạn 2027 và khi mua đủ; (b) CĐT ghi "Khí hút: Có" – cho biết lưu lượng hút cho từng máy (ống xả bơm chân không MS, split vent GC, snorkel); (c) yêu cầu ổn định nhiệt độ/độ ẩm của hãng thiết bị.',
 CDT+'dòng 13–19; IFC MIN lab: AAS 540, ICP 360 m³/h','Tải lạnh lớn nhất trong một phòng; điểm hút thiết bị','Tính theo SL khi mua đủ; 150 m³/h/máy cần hút điểm'),
('Sample preparation: CĐT ghi "Khí hút: Có" (máy nghiền ZM200, máy xay, máy hút bụi). Mặt bằng không có điểm hút. Đề nghị cho biết vị trí, loại điểm hút (chụp/cánh tay hút/hút bụi) và lưu lượng.',
 CDT+'dòng 28–37','Thêm hệ hút + gió bù','2 cánh tay hút × 800 m³/h (theo IFC ARM HOOD)'),
]),
('III','HỆ HÚT KHÍ THẢI & XỬ LÝ',[
('Scrubber: HSE yêu cầu tủ hút vô cơ đấu vào scrubber ướt xây mới, nước thải scrubber đấu vào hệ xử lý nước thải. Đề nghị xác nhận là lắp mới hoàn toàn hay đấu nối chung vào hệ scrubber cũ.',
 CDT+'sheet HSE mục 2, 2.1 (H₂SO₄ 6 L/ngày Digestion; ST1 H₂SO₄ 2, HCl 0,5, HNO₃ 0,8; ST2 HCl 0,5 L/ngày); IFC DW-AC-00-106, DW-AC-01-311','~9.100 m³/h khí acid; thiết bị, tuyến ống, đấu nối nước','Lắp mới hoàn toàn cho Phase 2; ống PP'),
('Tủ hút hữu cơ: HSE yêu cầu lọc than hoạt tính đồng bộ theo từng tủ. Danh mục thiết bị ghi mỗi tủ hút/chụp có quạt riêng 600 W, 3 pha. Đề nghị xác nhận: quạt riêng lắp sẵn theo máy hay quạt hút tập trung (như IFC MI-EAFE)?',
 CDT+'sheet HSE mục 2; dòng 8, 23, 42, 51','Số lượng quạt, ống, điện động lực 3 pha','Quạt riêng lắp sẵn theo máy'),
]),
('IV','GIÓ BÙ & ĐIỀU HÒA',[
('Phương án gió bù: GGV dự kiến theo IFC/Phase 1 – quạt cấp gió tươi FAFE (lọc G4/F8, không xử lý lạnh) cấp ngay trước tủ hood/chụp, lưu lượng ≈ gió thải; điều hòa chỉ xử lý tải phòng + gió tươi cho người. Đề nghị CĐT xác nhận chấp nhận phương án này, kể cả rủi ro RH phòng hóa vượt 70% vào mùa nồm/ngày nóng ẩm khi một phần gió bù tràn vào phòng.',
 'IFC DW-AC-01-305 (ghi chú: nhà thầu tự tính phương án cấp bù); Phase 1: HN-FAFE 3.100 m³/h bù autoclave 3.050 m³/h','Thay đổi công suất lạnh nhiều lần nếu yêu cầu xử lý lạnh toàn bộ gió bù','Theo IFC/Phase 1: gió bù không xử lý lạnh'),
('Hệ điều hòa: sử dụng AHU (DX như Phase 1) hay lắp điều hòa âm trần?',
 'Phase 1: HN-AHU-1.01/1.02 + HN-ODU-AHU (F3-SD-AC-017)','Chọn thiết bị, hệ ống gió','AHU DX như Phase 1'),
('Trong trường hợp dùng AHU, cho phép hồi gió từ các phòng hóa (có tủ hood) về AHU dùng chung với văn phòng/phòng thường không? Hay yêu cầu tách AHU riêng hoặc không hồi gió từ phòng hóa?',
 'Phase 1 hồi gió từ phòng vi sinh (F3-SD-AC-010)','Số lượng AHU, chất lượng không khí','Cho phép hồi gió như Phase 1'),
('Máy hút ẩm: CĐT ghi Instrument room cần máy hút ẩm (có lỗ thoát nước). Đề nghị xác nhận máy hút ẩm chỉ cho Instrument room hay cả các phòng hóa.',
 CDT+'dòng 11 (ghi chú)','Phạm vi, tải điều hòa, thoát nước','Chỉ Instrument room'),
]),
]
G=[G[0],G[3],G[2],G[1]]
G=[(rn,)+g[1:] for rn,g in zip(['I','II','III','IV'],G)]
r=HR+1; n=0
for code,title,items in G:
    ws.cell(r,1,code).font=f(True); ws.cell(r,1).alignment=C; ws.cell(r,2,title).font=f(True)
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=NC)
    for k in range(1,NC+1): ws.cell(r,k).fill=GRP; ws.cell(r,k).border=BD
    r+=1
    for q,ref,imp,ass in items:
        n+=1
        for k,v in enumerate([n,q,ref,imp,ass,''],1):
            c=ws.cell(r,k,v); c.font=f(); c.border=BD; c.alignment=C if k==1 else W
            if k>=6: c.fill=FILL
        r+=1
for k,v in {'A':6,'B':66,'C':38,'D':28,'E':36,'F':50}.items(): ws.column_dimensions[k].width=v
ws.freeze_panes='B2'
ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A3; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
ws.sheet_properties.pageSetUpPr.fitToPage=True; ws.print_title_rows='1:1'
wb.save(OUT); print(n)
