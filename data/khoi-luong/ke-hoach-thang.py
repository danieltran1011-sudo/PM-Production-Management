# Tải nhóm MD theo tháng (11/2026 – 2/2027) từ kế hoạch thật:
#   - Sản xuất: sheet PLANTABLE (原価) 2026年 trong file P-Plan 26/10/2026 (đã trích ra plantable-10-2026-3-2027.json).
#     MD của mỗi mã lấy theo người theo dõi mẫu PPS của mã đó trong file mẫu; không có thì theo brand.
#   - Mẫu: đơn mẫu còn tồn (chưa xong, hạn ≤ 10/2026) dồn vào 11/2026, cộng số đơn mẫu cùng tháng năm trước
#     (11/2025 – 2/2026) của cùng brand, gán cho người đang phụ trách brand đó.
# Chạy từ thư mục gốc repo: python3 data/khoi-luong/ke-hoach-thang.py
import json, collections, datetime as dt, re

THANG = ['2026-11', '2026-12', '2027-01', '2027-02']
MD = ['Cuong', 'Luan', 'Khanh', 'B.Trang', 'Q.Huy', 'Hien', 'Nhu']
GIO = {'Khanh': 8}; GIO_CHUAN = 9                     # giờ/ngày kể cả tăng ca
HE_SO_KHO = {'ALLTERRAIN': 1.3, 'DESCENT': 1.3, 'DCC': 1.3}
GIO_DICH = {'TNF': 3.83, 'TNF別注': 3.33, 'HH': 4.33, 'NW': 3.83, 'GOLDWIN': 1.75, 'NANGA': 2.5, 'CANTERBURY': 1.25, 'TETON': 3.33,
            'ALLTERRAIN': 3.83, 'DESCENT': 2.83, 'DCC': 2.83, 'MONTBELL': 4.0, '141': 3.83, 'KASCO': 1.83, 'FOXFIRE': 1.42, 'TATRAS': 1.42,
            'YONEX': 1.42, 'UTMF': 3.33}   # giờ dịch / mã (bảng 通訳作業時間)

# Trọng số theo loại mẫu: (theo dõi, dịch). Mẫu 1st/0次/đề án cần dịch trọn bộ tài liệu; 2nd dịch bản sửa;
# SMS, PPS chủ yếu theo dõi tiến độ, dịch góp ý; ảnh, lưu mẫu, test chỉ theo dõi nhẹ.
def trong_so_mau(loai):
    l = str(loai or '').upper()
    if l.startswith(('1ST', '0次', '提案', '展開')): return 1.0, 1.0
    if l.startswith(('2ND', '3RD', '4TH', '5TH')): return 0.8, 0.5
    if 'SMS' in l: return 0.5, 0.2
    if 'PPS' in l or 'MOCK' in l: return 0.6, 0.2
    return 0.3, 0.1

def nhom_brand(b):
    b = re.sub(r'\s+', '', str(b or '')).upper()
    for k, v in (('TNF別注', 'TNF別注'), ('TNF', 'TNF'), ('HELLY', 'HH'), ('HH', 'HH'), ('NEUTRAL', 'NW'), ('N/W', 'NW'), ('NW', 'NW'),
                 ('GOLDWIN', 'GOLDWIN'), ('NANGA', 'NANGA'), ('CANTERBURY', 'CANTERBURY'), ('TETON', 'TETON'), ('ALLTERRAIN', 'ALLTERRAIN'),
                 ('DESCENT', 'DESCENT'), ('DCC', 'DCC'), ('MONT', 'MONTBELL'), ('SNOWPEAK', 'SNOWPEAK'), ('141', '141'), ('KASCO', 'KASCO'), ('FOXFIRE', 'FOXFIRE'),
                 ('TATRAS', 'TATRAS'), ('YONEX', 'YONEX'), ('UTMF', 'UTMF')):
        if b.startswith(k): return v
    return b

def tai(nw_cho_khanh=True, tnf_tokyo_cho_khanh_tu='2026-11', khanh_mau_tnf=1/3, khanh_mau_tu='2026-12',
        giao_them={'Cuong': 0.1, 'Luan': 0.1}, ho_tro_luan={}, ho_tro={}, he_so_kho=True):
    """ho_tro_luan: {tháng: tỷ lệ việc của Luan do Khanh làm (hậu cần, dịch)}
    ho_tro: {tháng: [(người được hỗ trợ, người hỗ trợ, tỷ lệ việc chuyển)]} – hỗ trợ tạm thời theo tháng"""
    pt = json.load(open('data/khoi-luong/plantable-10-2026-3-2027.json'))['rows']
    mau = json.load(open('data/khoi-luong/mau-2025-2026.json'))
    # Mã TNF 東京企画 của Cuong: mã TNF (không phải 別注) đang gán cho Cuong
    P = {m: collections.defaultdict(lambda: collections.defaultdict(float)) for m in THANG}
    def k_(b): return HE_SO_KHO.get(b, 1.0) if he_so_kho else 1.0
    dau_tien = {}; dang_chay = set()
    for x in sorted(pt, key=lambda r: r['start']):
        if (x['ma'], x['nguoi']) not in dau_tien and x['start'] >= THANG[0] + '-01': dau_tien[(x['ma'], x['nguoi'])] = x['ma']
        if x['start'] < THANG[0] + '-01': dau_tien.setdefault((x['ma'], x['nguoi']), None)   # mã đã vào chuyền trước 11/2026
    for x in pt:
        b = nhom_brand(x['brand']); p = x['nguoi']
        if p == '?' and b in ('DCC', 'SNOWPEAK'): p = 'Hien'   # Snowpeak: quản lý xác nhận Hien theo dõi
        if nw_cho_khanh and b == 'NW' and p == 'Cuong': p = 'Khanh'
        s, e = dt.date.fromisoformat(x['start']), dt.date.fromisoformat(x['end'])
        for m in THANG:
            q = p
            if tnf_tokyo_cho_khanh_tu and b == 'TNF' and p == 'Cuong' and m >= tnf_tokyo_cho_khanh_tu: q = 'Khanh'
            y, mo = map(int, m.split('-')); a = dt.date(y, mo, 1); z = dt.date(y + (mo == 12), mo % 12 + 1, 1) - dt.timedelta(days=1)
            ov = (min(e, z) - max(s, a)).days + 1
            if ov > 0:
                P[m][q]['ngay_chuyen'] += ov * k_(b)
                if (m, q, x['ma']) not in dang_chay: dang_chay.add((m, q, x['ma'])); P[m][q]['ma_chay'] += 1
            if a <= s <= z and x['ma'] == dau_tien.get((x['ma'], p)):   # một mã chia nhiều lô/chuyền: chỉ tính lần vào chuyền đầu tiên
                dau_tien[(x['ma'], p)] = None
                P[m][q]['ma_vao'] += k_(b); P[m][q]['dich'] += GIO_DICH.get(b, 2.5) * k_(b); P[m][q]['ma_vao_goc'] += 1
    # mẫu: tồn + cùng kỳ năm trước
    def chu_mau(p, b):
        if p == 'Khanh': return 'Q.Huy' if b == 'GOLDWIN' else 'B.Trang'      # thời gian Khanh nghỉ: mẫu TNF → B.Trang, Goldwin → Huy
        if p == 'Cuong' and b == 'NW' and nw_cho_khanh: return 'Khanh'
        if p == 'Cuong' and b == 'YONEX': return 'B.Trang'
        return p
    def them_mau(m, p, b, loai, n=1.0):
        w, wd = trong_so_mau(loai)
        if p == 'B.Trang' and b == 'TNF' and khanh_mau_tnf and m >= khanh_mau_tu:
            P[m]['Khanh']['mau'] += n * khanh_mau_tnf * w; P[m]['Khanh']['dich'] += n * khanh_mau_tnf * wd * GIO_DICH['TNF']; P[m]['Khanh']['don_mau'] += n * khanh_mau_tnf
            n = n * (1 - khanh_mau_tnf)
        P[m][p]['mau'] += n * w * k_(b); P[m][p]['dich'] += n * wd * GIO_DICH.get(b, 2.5) * k_(b); P[m][p]['don_mau'] += n
    for x in mau:
        b = nhom_brand(x['brand']); p = chu_mau(x['nguoi'], b)
        if p in ('H.Trang',): continue
        if x['sheet'] == 'サンプル表2026' and not x['xong'] and x['due'] < '2026-11-01' and x['due'] >= '2026-06-01':
            them_mau('2026-11', p, b, x['loai'])          # đơn mẫu còn tồn → dồn vào 11/2026
        for m in THANG:
            y, mo = map(int, m.split('-')); ky_truoc = f'{y-1}-{mo:02d}'
            if x['due'][:7] == ky_truoc: them_mau(m, p, b, x['loai'])
    out = {}
    cap = {p: GIO.get(p, GIO_CHUAN) / GIO_CHUAN for p in MD}
    # Mốc 1,0 = khối lượng trung bình 1 người / 1 tháng trong cả giai đoạn 11/2026 – 2/2027
    # (cùng một mốc cho mọi tháng, nên tháng bận và tháng rảnh so được với nhau).
    tb = {k: sum(P[m][p][k] for m in THANG for p in MD) / len(THANG) / sum(cap.values()) for k in ('ma_vao', 'ngay_chuyen', 'mau', 'dich')}
    for m in THANG:
        tuyet_doi = {p: sum(P[m][p][k] / tb[k] for k in tb if tb[k]) / 4 for p in MD}
        tl = ho_tro_luan.get(m, 0)
        tuyet_doi['Khanh'] += tuyet_doi['Luan'] * tl; tuyet_doi['Luan'] *= 1 - tl
        for (tu, den, r) in ho_tro.get(m, []):
            tuyet_doi[den] += tuyet_doi[tu] * r; tuyet_doi[tu] *= 1 - r
        out[m] = {p: dict(chi_so=round(tuyet_doi[p] / cap[p] / (1 - giao_them.get(p, 0)), 2),
                          ma_vao_chuyen=int(P[m][p]['ma_vao_goc']), ma_dang_chay=int(P[m][p]['ma_chay']),
                          don_mau=round(P[m][p]['don_mau']), gio_dich=round(P[m][p]['dich'])) for p in MD}
    return out

if __name__ == '__main__':
    L1 = {'2026-11': 0.25, '2026-12': 0.25, '2027-01': 0.35, '2027-02': 0.35}
    PA = {
     'A. Phương án cũ: NW → Khanh ngay; TNF 東京企画 và 1/3 mẫu TNF → Khanh từ 12/2026': tai(tnf_tokyo_cho_khanh_tu='2026-12'),
     'B. A + Khanh hỗ trợ Luan (~25% tháng 11–12, ~35% tháng 1–2)': tai(tnf_tokyo_cho_khanh_tu='2026-12', ho_tro_luan=L1),
     'C. B + TNF 東京企画 → Khanh từ 11/2026 (đã quyết 28/10/2026)': tai(tnf_tokyo_cho_khanh_tu='2026-11', ho_tro_luan=L1),
     'D. C + hỗ trợ theo tháng: 12/2026 Huy hỗ trợ Cuong 25%; 1/2027 B.Trang hỗ trợ Huy 20%':
        tai(tnf_tokyo_cho_khanh_tu='2026-11', ho_tro_luan=L1, ho_tro={'2026-12': [('Cuong', 'Q.Huy', 0.25)], '2027-01': [('Q.Huy', 'B.Trang', 0.2)]}),
     'E. D + 12/2026 Nhu hỗ trợ thêm Cuong 15%':
        tai(tnf_tokyo_cho_khanh_tu='2026-11', ho_tro_luan=L1, ho_tro={'2026-12': [('Cuong', 'Q.Huy', 0.25), ('Cuong', 'Nhu', 0.15)], '2027-01': [('Q.Huy', 'B.Trang', 0.2)]}),
    }
    for k, v in PA.items():
        print('==', k)
        for m in THANG: print(' ', m, {p: x['chi_so'] for p, x in v[m].items()})
    print('-- số liệu gốc (phương án A): mã vào chuyền, mã đang chạy, đơn mẫu, giờ dịch')
    for m in THANG: print(' ', m, {p: (x['ma_vao_chuyen'], x['ma_dang_chay'], x['don_mau'], x['gio_dich']) for p, x in PA[list(PA)[0]][m].items()})
    json.dump(PA, open('data/khoi-luong/ke-hoach-thang.json', 'w'), ensure_ascii=False, indent=1)
