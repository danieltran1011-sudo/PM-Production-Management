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

# Hệ số độ khó theo brand (định tính, do quản lý đánh giá). 1,0 = bình thường.
# ALLTERRAIN, DESCENTE: tài liệu, rập không rõ, phải hỏi kỹ thuật nhiều; TIJ chuẩn bị yếu, hay thiếu NPL.
HE_SO_KHO = {'ALLTERRAIN': 1.3, 'DESCENT': 1.3, 'DESCENTEGOLF': 1.3}

# Sheet Hien trong file SPL thiếu mẫu 26FW (sheet tổng hợp ghi 36 cái, 8 mã, không có brand).
# Bù vào làm mẫu ước cho SMS 27FW, giả định là ALLTERRAIN (13.131 giây ≈ 5 ngày/cái).
BU_MAU = [dict(nguoi='Hien', mua='26FW', brand='ALLTERRAIN', loai='SMS', sl=36, ma=8, giay=13131, ngay=180, ghichu='bù từ sheet tổng hợp')]
MD = ['Cuong', 'Luan', 'Hien', 'Q.Huy', 'B.Trang', 'Nhu', 'Khanh']

def tai(chia_tnf_sms_cho_khanh=0.0, he_so_kho=True, bu_mau=True, chia=(), giao_them=None, chuyen_phan=(), chu_moi=None):
    """chia: [(brand, 'bulk'/'spl', người cũ, người mới, tỷ lệ)] – chia theo lô/mã của một brand.
    giao_them: {người: phần thời gian dành cho việc giao thêm, vd 0.1}.
    chuyen_phan: [(người cũ, người mới, tỷ lệ)] – chuyển một phần việc theo chức năng (vd. dịch, báo cáo),
    người cũ vẫn giữ brand; tỷ lệ tính trên phần việc brand của người cũ."""
    d = json.load(open('data/khoi-luong/thuc-te.json'))
    if bu_mau: d['spl'] = d['spl'] + BU_MAU
    chu = dict(CHU, **(chu_moi or {}))
    P = collections.defaultdict(lambda: collections.defaultdict(float))
    def add1(p, b, ma, bulk, mau, dich):
        k = HE_SO_KHO.get(b, 1.0) if he_so_kho else 1.0
        P[p]['ma'] += ma * k; P[p]['bulk'] += bulk * k; P[p]['mau'] += mau * k; P[p]['dich'] += dich * k
    def add(p, ma, bulk, mau, dich, b=None, loai=None):
        if b is None: b = add.b
        for (cb, ck, tu, den, tl) in chia:
            if cb == b and ck == loai and tu == p:
                add1(den, b, ma * tl, bulk * tl, mau * tl, dich * tl)
                ma, bulk, mau, dich = ma * (1 - tl), bulk * (1 - tl), mau * (1 - tl), dich * (1 - tl)
        add1(p, b, ma, bulk, mau, dich)
    for r in d['bulk']:
        if r['mua'] != '27SS': continue
        b = key(r['brand']); p = chu.get(b, r['nguoi'])
        add(p, r['ma'], r['ngay'], 0, r['ma'] * (TB.get(b) or 0), b, 'bulk')
    for r in d['spl']:
        if not ((r['mua'] == '26FW' and r['loai'] == 'SMS') or (r['mua'] == '27SS' and r['loai'] == '1st-3rd')): continue
        b = key(r['brand']); p = chu.get(b, r['nguoi'])
        ngay = r['ngay']
        if not r['giay']:
            s = est_sec(r['brand']); ngay = r['sl'] * ngay_mau(s) if s else ngay
        h = (TS.get(b) or TB.get(b) or 0) * r['ma']
        if b == 'TNF(東京企画、GTC企画共通)サンプル' and r['loai'] == 'SMS' and chia_tnf_sms_cho_khanh:
            k = chia_tnf_sms_cho_khanh
            add('Khanh', r['ma'] * k, 0, ngay * k, h * k, b, 'spl'); add(p, r['ma'] * (1 - k), 0, ngay * (1 - k), h * (1 - k), b, 'spl')
        else:
            add(p, r['ma'], 0, ngay, h, b, 'spl')
    # Mức trung bình = tổng việc của 7 người chia cho tổng công suất (Khanh = 8/9 người),
    # nên chuyển việc giữa hai người không làm thay đổi chỉ số của người khác.
    cap = {p: GIO_NGAY.get(p, GIO_CHUAN) / GIO_CHUAN for p in MD}
    tb = {k: sum(P[p][k] for p in MD) / sum(cap.values()) for k in ('ma', 'bulk', 'mau', 'dich')}
    cs0 = {p: sum(P[p][k] / tb[k] for k in tb) / 4 * cap[p] for p in MD}  # tải tuyệt đối (1,0 = 1 người trung bình)
    for (tu, den, tl) in chuyen_phan:
        cs0[den] += cs0[tu] * tl; cs0[tu] *= (1 - tl)
    out = {}
    for p in MD:
        cs = cs0[p] / cap[p]
        con_lai = 1 - (giao_them or {}).get(p, 0)   # phần thời gian còn cho brand
        out[p] = dict(chi_so=round(cs / cap[p] / con_lai, 2), ma=round(P[p]['ma']), ngay_bulk=round(P[p]['bulk']), ngay_mau=round(P[p]['mau']), gio_dich=round(P[p]['dich']))
    return out

def in_kq(kq):
    for k, v in kq.items():
        print(k); print('   ', {p: x['chi_so'] for p, x in v.items()})

if __name__ == '__main__':
    K = 1/3  # Khanh nhận 1/3 mẫu TNF triển lãm của B.Trang (đã quyết)
    # Quyết định 10/2026: Hien, Huy, Nhu giữ nguyên. Cuong chỉ chuyển NEUTRALWORKS (→ Khanh), giữ HELLY HANSEN (chủ lực)
    # và TNF 東京企画, nhận thêm điều phối chung (xuất hàng cả nhóm, quản lý chung may mẫu). Luan nhận việc quản lý.
    # Thời gian việc giao thêm chưa biết, thử 2 mức.
    kq = {}
    for ten, giao in (('điều phối/quản lý ~1 giờ/ngày (Cuong 10%, Luan 10%)', {'Cuong': 0.1, 'Luan': 0.1}),
                      ('điều phối ~2 giờ/ngày (Cuong 20%, Luan 10%)', {'Cuong': 0.2, 'Luan': 0.1})):
        kq[f'0. Chưa chia việc Luan – {ten}'] = tai(K, giao_them=giao)
        kq[f'1. Khanh làm hậu cần TNF 富山企画 (~15% việc Luan) – {ten}'] = tai(K, giao_them=giao, chuyen_phan=[('Luan', 'Khanh', 0.15)])
        kq[f'2. + Khanh hỗ trợ dịch tài liệu (~25% việc Luan) – {ten}'] = tai(K, giao_them=giao, chuyen_phan=[('Luan', 'Khanh', 0.25)])
    in_kq(kq)
    json.dump(kq, open('data/khoi-luong/du-bao.json', 'w'), ensure_ascii=False, indent=1)
