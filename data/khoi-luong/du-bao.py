# Dự báo tải nhóm MD cho giai đoạn tới (khoảng 11/2026 – 2/2027):
#   - Bulk 27SS: số liệu thực tế trong file BULK.
#   - Mẫu triển lãm (SMS) 27FW: chưa có số liệu, ước bằng SMS 26FW của cùng brand.
#   - Mẫu 1st 28SS: chưa có số liệu, ước bằng 1st–3rd 27SS của cùng brand.
# Brand được gán cho người đang phụ trách hiện tại (bảng 22/7/2026 + quyết định 10/2026).
# Công suất theo giờ/ngày có tăng ca: mọi người 9 giờ (8 + 1 tăng ca), Khanh 8 giờ (đến 16h, tăng ca đến 17h).
# Chạy từ thư mục gốc repo: python3 data/khoi-luong/du-bao.py
import json, collections
src = open('data/khoi-luong/build.py').read()
ns = {}; exec(src[:src.index('# Chỉ số tải chỉ so')].replace("print(", "(lambda *a, **k: None)("), ns)
key, TB, TS, est_sec, ngay_mau = ns['key'], ns['TB'], ns['TS'], ns['est_sec'], ns['ngay_mau']

CHU = {  # brand (đã chuẩn hóa) -> người phụ trách hiện tại
 'HH': 'Cuong', 'TNF(東京企画)': 'Cuong', 'NW': 'Khanh',
 'GOLDWIN': 'Q.Huy', 'CANTERBURY': 'Q.Huy', 'NANGA': 'Q.Huy', 'TNF別注東京': 'Q.Huy',
 'TNF別注富山': 'Nhu', 'TNF(GTC企画)': 'Luan',
 'TNF(東京企画、GTC企画共通)サンプル': 'B.Trang', 'YONEX': 'B.Trang', 'KASCO': 'B.Trang', 'FOXFIRE': 'B.Trang', 'TATRAS': 'B.Trang',
 'HAT+': 'H.Trang', '141': 'H.Trang',
}
GIO_NGAY = {'Khanh': 8}; GIO_CHUAN = 9
MD = ['Cuong', 'Luan', 'Hien', 'Q.Huy', 'B.Trang', 'Nhu', 'Khanh']

def tai(chia_tnf_sms_cho_khanh=0.0):
    d = json.load(open('data/khoi-luong/thuc-te.json'))
    P = collections.defaultdict(lambda: collections.defaultdict(float))
    def add(p, ma, bulk, mau, dich):
        P[p]['ma'] += ma; P[p]['bulk'] += bulk; P[p]['mau'] += mau; P[p]['dich'] += dich
    for r in d['bulk']:
        if r['mua'] != '27SS': continue
        p = CHU.get(key(r['brand']), r['nguoi'])
        add(p, r['ma'], r['ngay'], 0, r['ma'] * (TB.get(key(r['brand'])) or 0))
    for r in d['spl']:
        if not ((r['mua'] == '26FW' and r['loai'] == 'SMS') or (r['mua'] == '27SS' and r['loai'] == '1st-3rd')): continue
        b = key(r['brand']); p = CHU.get(b, r['nguoi'])
        ngay = r['ngay']
        if not r['giay']:
            s = est_sec(r['brand']); ngay = r['sl'] * ngay_mau(s) if s else ngay
        h = (TS.get(b) or TB.get(b) or 0) * r['ma']
        if b == 'TNF(東京企画、GTC企画共通)サンプル' and r['loai'] == 'SMS' and chia_tnf_sms_cho_khanh:
            k = chia_tnf_sms_cho_khanh
            add('Khanh', r['ma'] * k, 0, ngay * k, h * k); add(p, r['ma'] * (1 - k), 0, ngay * (1 - k), h * (1 - k))
        else:
            add(p, r['ma'], 0, ngay, h)
    # Mức trung bình = tổng việc của 7 người chia cho tổng công suất (Khanh = 8/9 người),
    # nên chuyển việc giữa hai người không làm thay đổi chỉ số của người khác.
    cap = {p: GIO_NGAY.get(p, GIO_CHUAN) / GIO_CHUAN for p in MD}
    tb = {k: sum(P[p][k] for p in MD) / sum(cap.values()) for k in ('ma', 'bulk', 'mau', 'dich')}
    out = {}
    for p in MD:
        cs = sum(P[p][k] / tb[k] for k in tb) / 4
        out[p] = dict(chi_so=round(cs / cap[p], 2), ma=round(P[p]['ma']), ngay_bulk=round(P[p]['bulk']), ngay_mau=round(P[p]['mau']), gio_dich=round(P[p]['dich']))
    return out

if __name__ == '__main__':
    kq = {'B.Trang làm hết mẫu TNF SMS 27FW': tai(0), 'Khanh chia 1/3 mẫu TNF SMS 27FW': tai(1/3), 'Khanh chia 1/2 mẫu TNF SMS 27FW': tai(1/2)}
    for k, v in kq.items():
        print('==', k)
        for p, x in v.items(): print(' ', p.ljust(8), x)
    json.dump(kq, open('data/khoi-luong/du-bao.json', 'w'), ensure_ascii=False, indent=1)
