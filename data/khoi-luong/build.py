# Tính khối lượng công việc theo người từ file thực tế BULK và SPL (mẫu).
# Chạy từ thư mục gốc repo: python3 data/khoi-luong/build.py
# Kết quả: data/khoi-luong/thuc-te.json (dữ liệu gốc đã gom) và data/khoi-luong/hien-tai.json (tải hiện tại theo người).
BULK = "data/nguon/Thuc-te-BULK-19FW-27SS-20261012.xlsx"
SPL = "data/nguon/Thuc-te-mau-SPL-23FW-27SS-20261012.xlsx"
OUT = "data/khoi-luong"
import openpyxl, json, re, collections
def norm(n): return str(n).strip()
wbB=openpyxl.load_workbook(BULK,data_only=True)
bulk=[]
for sn in wbB.sheetnames[4:]:
    for r in wbB[sn].iter_rows(min_row=4,values_only=True):
        if not r[1] or not r[2] or not isinstance(r[3],(int,float)): continue
        bulk.append(dict(nguoi=norm(sn),mua=norm(r[1]),brand=norm(r[2]),sl=r[3] or 0,ma=r[4] or 0,giay=r[5],ngay=r[6] or 0,ghichu=r[7]))
wbS=openpyxl.load_workbook(SPL,data_only=True)
spl=[]
for sn in wbS.sheetnames[5:]:
    for r in wbS[sn].iter_rows(min_row=13,values_only=True):
        if not r[1] or not r[2]: continue
        spl.append(dict(nguoi=norm(sn),mua=norm(r[1]),brand=norm(r[2]),loai='SMS' if 'SMS' in str(r[3]) else '1st-3rd',sl=r[4] or 0,ma=r[5] or 0,giay=r[6],ngay=r[7] or 0,ghichu=r[8]))
t=wbB['通訳作業時間']
tr=[]
names={c:norm(t.cell(5,c).value) for c in range(3,11)}
for r in range(6,85):
    br=t.cell(r,2).value
    if not br: continue
    tr.append(dict(brand=norm(br),nguoi=[names[c] for c in names if str(t.cell(r,c).value or '').strip()=='●'],gio_ma=t.cell(r,22).value,trang=t.cell(r,23).value,ty_le_dich=t.cell(r,11).value))
json.dump(dict(bulk=bulk,spl=spl,dich=tr),open(OUT+'/thuc-te.json','w'),ensure_ascii=False,indent=0,default=str)



d=dict(bulk=bulk,spl=spl,dich=tr)
def key(b):
    b=str(b).replace('　',' ').strip().upper()
    b=b.replace('TOYAMA','富山').replace('HAT+CAP...','HAT +').replace('HAT+','HAT +')
    b=re.sub(r'\s+','',b)
    return {'TNF別注':'TNF別注東京','TETON':'TETONBROS','MACKINTOSHLONDON':'MACKINTOSH','DESCENTE':'DESCENT'}.get(b,b)
def table(path, sheet, r0, r1):
    ws=openpyxl.load_workbook(path,data_only=True)[sheet]; t={}
    for r in range(r0,r1+1):
        b=ws.cell(r,2).value; v=ws.cell(r,22).value
        if b and isinstance(v,(int,float)): t.setdefault(key(b),v)
    return t
TB=table(BULK,'通訳作業時間',6,84)
TS=table(SPL,'APIC 翻訳業務時間',6,53)
TS.setdefault('TNF別注富山',TS.get('TNF別注富山',3.33))
# 1pcs縫製期間 theo 平均秒数 (bảng trong file SPL)
def ngay_mau(sec): 
    for lim,dd in ((3000,1),(8000,2),(10000,3),(12000,4),(15000,5)):
        if sec<=lim: return dd
    return 6
# ước tính giây cho mẫu 27SS chưa có: lấy giây TB của brand mùa 26SS, không có thì mùa gần nhất
sec_hist=collections.defaultdict(dict)
for r in d['spl']:
    if r['giay']: sec_hist[key(r['brand'])][r['mua']]=float(r['giay'])
def est_sec(b):
    h=sec_hist.get(key(b),{})
    for m in ('26SS','25SS','26FW','25FW'):
        if m in h: return h[m]
    return None
CUR={'bulk':['26FW','27SS'],'spl':['27SS']}
# Hệ số độ khó theo brand (định tính, do quản lý đánh giá 10/2026). Chỉ dùng khi tính chỉ số tải; số mã/ngày hiển thị vẫn là số gốc.
# ALLTERRAIN, DESCENTE: tài liệu, rập không rõ, phải hỏi kỹ thuật nhiều; TIJ chuẩn bị yếu, hay thiếu NPL.
HE_SO_KHO={'ALLTERRAIN':1.3,'DESCENT':1.3,'DESCENTEGOLF':1.3}
P=collections.defaultdict(lambda: collections.defaultdict(float)); brands=collections.defaultdict(set); miss=set(); notes=[]
for r in d['bulk']:
    if r['mua'] not in CUR['bulk']: continue
    p=r['nguoi']; m=r['mua']
    P[p][f'bulk_{m}_sl']+=r['sl']; P[p][f'bulk_{m}_ma']+=r['ma']; P[p][f'bulk_{m}_ngay']+=r['ngay']
    h=TB.get(key(r['brand']))
    if h is None: miss.add(('B',r['brand']))
    if m=='27SS': P[p]['gio_dich']+=r['ma']*(h or 0)
    k=HE_SO_KHO.get(key(r['brand']),1.0)
    P[p]['w_bulk']+=r['ngay']*k
    if m=='27SS': P[p]['w_ma']+=r['ma']*k; P[p]['w_dich']+=r['ma']*(h or 0)*k
    brands[p].add(key(r['brand']))
for r in d['spl']:
    if r['mua'] not in CUR['spl']: continue
    p=r['nguoi']; ngay=r['ngay']
    if not r['giay']:
        s=est_sec(r['brand'])
        if s: ngay=r['sl']*ngay_mau(s); notes.append((p,r['brand'],r['loai'],r['sl'],s,ngay))
    P[p]['spl_sl']+=r['sl']; P[p]['spl_ma']+=r['ma']; P[p]['spl_ngay']+=ngay
    h=TS.get(key(r['brand'])) or TB.get(key(r['brand']))
    if h is None: miss.add(('S',r['brand']))
    P[p]['gio_dich']+=r['ma']*(h or 0); brands[p].add(key(r['brand']))
    k=HE_SO_KHO.get(key(r['brand']),1.0)
    P[p]['w_ma']+=r['ma']*k; P[p]['w_mau']+=ngay*k; P[p]['w_dich']+=r['ma']*(h or 0)*k
for p in P: P[p]['so_brand']=len(brands[p])
cols=['bulk_26FW_ma','bulk_26FW_sl','bulk_26FW_ngay','bulk_27SS_ma','bulk_27SS_sl','bulk_27SS_ngay','spl_ma','spl_sl','spl_ngay','gio_dich','so_brand']
# chỉ số tải tương đối: TB phần chia của 4 nhóm yếu tố (mã hàng, ngày chuyền bulk, ngày may mẫu, giờ dịch), 1.0 = trung bình nhóm
# Chỉ số tải chỉ so trong nhóm MD đang làm đủ kỳ.
# H.Trang thuộc nhóm CAD (chỉ phụ HAT&CAP, 141); Khanh nghỉ thai sản, đi làm lại từ giữa 10/2026.
CAD = {'H.Trang'}
KHONG_SO = {'Khanh'}
people=[p for p in P if p not in CAD|KHONG_SO]
def share(f):
    tot=sum(f(P[p]) for p in people); n=len(people)
    return {p:(f(P[p])/tot*n if tot else 0) for p in people}
S1=share(lambda x:x['w_ma'])
S2=share(lambda x:x['w_bulk'])
S3=share(lambda x:x['w_mau'])
S4=share(lambda x:x['w_dich'])
# Trọng số 4 yếu tố (mã, ngày chuyền bulk, ngày may mẫu, giờ dịch). Mặc định chia đều.
PHUONG_AN = {'deu': (.25,.25,.25,.25), 'uu_tien_dich': (.2,.2,.2,.4), 'uu_tien_bulk': (.2,.4,.2,.2), 'uu_tien_mau': (.2,.2,.4,.2)}
for p in people:
    f=(S1[p],S2[p],S3[p],S4[p])
    P[p]['cs_ma']=round(S1[p],2); P[p]['cs_bulk']=round(S2[p],2); P[p]['cs_mau']=round(S3[p],2); P[p]['cs_dich']=round(S4[p],2)
    P[p]['chi_so_theo_trong_so']={k:round(sum(w*x for w,x in zip(ws,f)),2) for k,ws in PHUONG_AN.items()}
    P[p]['chi_so']=P[p]['chi_so_theo_trong_so']['deu']
for p in P:
    if p in CAD: P[p]['nhom']='CAD (không tính chỉ số MD)'
    elif p in KHONG_SO: P[p]['nhom']='MD – vừa đi làm lại, không tính chỉ số'
    else: P[p]['nhom']='MD'
print('person'.ljust(8),' '.join(c[-12:].rjust(12) for c in cols+['cs_ma','cs_bulk','cs_mau','cs_dich','chi_so']))
for p in sorted(people,key=lambda p:-P[p]['chi_so']):
    print(p, P[p]['chi_so_theo_trong_so'])
    print(p.ljust(8),' '.join(str(round(P[p][c],1)).rjust(12) for c in cols+['cs_ma','cs_bulk','cs_mau','cs_dich','chi_so']))
print('MISS',miss); print('EST',notes)
json.dump({p:dict(v,brands=sorted(brands[p])) for p,v in P.items()},open(OUT+'/hien-tai.json','w'),ensure_ascii=False,indent=1)
