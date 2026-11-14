import json, os, subprocess, openpyxl
W='/tmp/claude-0/scen'; os.makedirs(W,exist_ok=True)
if os.path.exists('/tmp/claude-0/scen.json'): os.remove('/tmp/claude-0/scen.json')
SC=[('1. Theo IFC/Phase 1: bù FAFE không xử lý, AHU không tính tải gió bù – thiết bị khi mua đủ',1,0,0,'Phương án cơ sở',1),
('1b. Như (1) – chỉ thiết bị năm 2027',1,0,0,'Tham khảo giai đoạn đầu',0),
('2. Như (1), giả định 10% gió bù tràn vào phòng',1,0,0.1,'Độ nhạy – thực tế cấp gió trước chụp luôn có phần tràn',1),
('3. Như (1), giả định 20% gió bù tràn vào phòng',1,0,0.2,'Độ nhạy',1),
('4. Tham khảo: gió bù xử lý lạnh 100% (FAHU)',1,1,0,'Không theo IFC – để so sánh',1)]
rows=[]
for lab,div,mu,spl,note,qb in SC:
    out=f'{W}/s.xlsx'
    env=dict(os.environ,OVR=json.dumps({'DIV':div,'MU':mu,'SPL':spl,'QB':qb}),OUT=out)
    subprocess.run(['python3','/tmp/claude-0/scripts/build.py'],env=env,check=True,capture_output=True)
    subprocess.run(['soffice','--headless','--calc','--convert-to','xlsx','--outdir',W+'/r',out],capture_output=True)
    wb=openpyxl.load_workbook(W+'/r/s.xlsx',data_only=True); p=wb['3_TongHop']
    v=lambda r:p.cell(r,4).value
    rows.append([lab,div,mu,spl,v(12),v(6),v(15),v(10),v(10)/3.517,v(11),note])
    print(lab, [round(x) for x in (v(12),v(6),v(15),v(10),v(11))], round(p.cell(10,2).value,1), round(p.cell(10,3).value,1))
json.dump(rows,open('/tmp/claude-0/scen.json','w'))
