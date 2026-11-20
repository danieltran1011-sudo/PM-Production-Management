# Mô phỏng chỉ số tải khi chuyển brand giữa các MD. Chạy từ thư mục gốc repo: python3 data/khoi-luong/mo-phong.py
# Sửa danh sách chuyển ở cuối file: (người cũ, người mới, brand, "bulk"/"spl"/"all", mùa hoặc None).
import json, collections
# chạy lại phần tính từ build.py nhưng cho phép chuyển brand
src=open('data/khoi-luong/build.py').read()
cut=src.index('# Chỉ số tải chỉ so')
ns={}; exec(src[:cut].replace("print(","(lambda *a,**k:None)("), ns)
def run(moves, label):
    d=json.load(open('data/khoi-luong/thuc-te.json'))
    for r in d['bulk']+d['spl']:
        for (frm,to,brand,kind,mua) in moves:
            if r['nguoi']==frm and ns['key'](r['brand'])==ns['key'](brand) and (kind=='all' or (kind=='bulk' and r in d['bulk']) or (kind=='spl' and r in d['spl'])) and (mua is None or r['mua']==mua):
                r['nguoi']=to
    P=collections.defaultdict(lambda: collections.defaultdict(float))
    for r in d['bulk']:
        if r['mua'] not in ('26FW','27SS'): continue
        p=r['nguoi']; P[p]['ma']+= r['ma'] if r['mua']=='27SS' else 0; P[p]['bulk']+=r['ngay']
        if r['mua']=='27SS': P[p]['dich']+=r['ma']*(ns['TB'].get(ns['key'](r['brand'])) or 0)
    for r in d['spl']:
        if r['mua']!='27SS': continue
        p=r['nguoi']; ngay=r['ngay']
        if not r['giay']:
            s=ns['est_sec'](r['brand']); ngay=r['sl']*ns['ngay_mau'](s) if s else ngay
        P[p]['ma']+=r['ma']; P[p]['mau']+=ngay
        P[p]['dich']+=r['ma']*((ns['TS'].get(ns['key'](r['brand'])) or ns['TB'].get(ns['key'](r['brand']))) or 0)
    people=['Cuong','Luan','Hien','Q.Huy','B.Trang','Nhu']
    cap={'Khanh':8/9,'Nhu':0.7}  # Khanh 8 giờ/ngày so với 9 giờ; Nhu ~30% việc văn phòng
    out={}
    ds=people+['Khanh']  # trung bình theo tổng công suất, nên chuyển việc không làm đổi chỉ số người khác
    for p in ds:
        v=0
        for k in ('ma','bulk','mau','dich'):
            avg=sum(P[q][k] for q in ds)/sum(cap.get(q,1) for q in ds)
            v+=P[p][k]/avg/4
        out[p]=round(v/cap.get(p,1),2)
    print(label, out)
run([], 'Hiện tại (Nhu, Khanh đã quy đổi theo công suất):')
run([('Cuong','Khanh','NW','all',None)], 'Chuyển NEUTRALWORKS Cuong→Khanh:')
