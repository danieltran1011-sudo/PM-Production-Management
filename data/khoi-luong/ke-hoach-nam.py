# Tải nhóm MD theo tháng cho cả vòng 12 tháng (11/2026 – 10/2027), tính chồng giữa mẫu và sản xuất của hai mùa.
#   Vòng mùa (quản lý mô tả): FW: mẫu phát triển 1st–3rd từ tháng 7 → SMS tháng 10–11 → sản xuất tháng 4–9 năm sau.
#                             SS: mẫu phát triển từ tháng 1 → SMS tháng 4–5 → sản xuất tháng 10 – 3 năm sau.
#   - Sản xuất mùa đã có kế hoạch (đến 27SS): lịch thật trong Plantable 26/10/2026.
#   - Sản xuất mùa chưa có kế hoạch (27FW, 28SS): lịch của mùa tương ứng năm trước (26FW, 27SS) dời +1 năm, gán người đang phụ trách.
#   - Mẫu: đơn còn tồn dồn vào 11/2026 + số đơn cùng tháng năm trước (file mẫu 2025–2026), gán người đang phụ trách.
#   - Công suất: số ngày làm việc thực của từng người (thứ 2–7, nghỉ thứ 7 theo nhóm, nghỉ lễ) × giờ/ngày.
# Chạy: python3 data/khoi-luong/ke-hoach-nam.py  → data/khoi-luong/ke-hoach-nam.json
import json, collections, datetime as dt, re, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lich_lam_viec import ngay_lam, ngay_xuong

THANG = ['2026-11', '2026-12', '2027-01', '2027-02', '2027-03', '2027-04', '2027-05', '2027-06', '2027-07', '2027-08', '2027-09', '2027-10']
MUA_CO_KE_HOACH = (27, 0)                          # mùa đến 27SS đã có kế hoạch thật; mùa sau dùng mùa năm trước
MD = ['Cuong', 'Luan', 'Khanh', 'B.Trang', 'Q.Huy', 'Hien', 'Nhu']
GIO = {'Khanh': 8}; GIO_CHUAN = 9
HE_SO_KHO = {'ALLTERRAIN': 1.3, 'DESCENT': 1.3, 'DCC': 1.3}
GIO_DICH = {'TNF': 3.83, 'TNF別注': 3.33, 'HH': 4.33, 'NW': 3.83, 'GOLDWIN': 1.75, 'NANGA': 2.5, 'CANTERBURY': 1.25, 'TETON': 3.33,
            'ALLTERRAIN': 3.83, 'DESCENT': 2.83, 'DCC': 2.83, 'MONTBELL': 4.0, '141': 3.83, 'KASCO': 1.83, 'FOXFIRE': 1.42, 'TATRAS': 1.42,
            'YONEX': 1.42, 'UTMF': 3.33}

# Giai đoạn theo vòng mùa (để đọc bảng; không dùng trong tính toán)
def giai_doan(t):
    y, m = int(t[2:4]), int(t[5:])
    g = []
    if 7 <= m <= 11: g.append(f'Mẫu 1st–3rd {y + 1}FW')
    if m in (10, 11): g.append(f'SMS {y + 1}FW')
    if 1 <= m <= 5: g.append(f'Mẫu 1st–3rd {y + 1}SS')
    if m in (4, 5): g.append(f'SMS {y + 1}SS')
    if m <= 3: g.append(f'Sản xuất {y}SS')
    if 4 <= m <= 9: g.append(f'Sản xuất {y}FW')
    if m >= 10: g.append(f'Sản xuất {y + 1}SS')
    if m in (2, 3): g.append(f'PPS, chuẩn bị {y}FW')
    if m in (8, 9): g.append(f'PPS, chuẩn bị {y + 1}SS')
    return ', '.join(g)

def nhom_brand(b):
    b = re.sub(r'\s+', '', str(b or '')).upper()
    for k, v in (('TNF別注', 'TNF別注'), ('TNF', 'TNF'), ('HELLY', 'HH'), ('HH', 'HH'), ('NEUTRAL', 'NW'), ('N/W', 'NW'), ('NW', 'NW'),
                 ('GOLDWIN', 'GOLDWIN'), ('NANGA', 'NANGA'), ('CANTERBURY', 'CANTERBURY'), ('TETON', 'TETON'), ('ALLTERRAIN', 'ALLTERRAIN'),
                 ('DESCENT', 'DESCENT'), ('DCC', 'DCC'), ('MONT', 'MONTBELL'), ('SNOWPEAK', 'SNOWPEAK'), ('141', '141'), ('KASCO', 'KASCO'),
                 ('FOXFIRE', 'FOXFIRE'), ('TATRAS', 'TATRAS'), ('YONEX', 'YONEX'), ('UTMF', 'UTMF')):
        if b.startswith(k): return v
    return b

def mua_so(m):
    m = str(m or '').replace('"', '').strip()
    return (int(m[:2]), 0 if m.endswith('SS') else 1) if m[:2].isdigit() else (0, 0)

def chu_san_xuat(p, b, mua):
    """Người đang phụ trách (sau các quyết định 10/2026) cho một lô sản xuất."""
    if b == 'NW': return 'Khanh'
    if b == 'GOLDWIN': return 'Q.Huy'
    if b == 'YONEX': return 'B.Trang' if mua_so(mua) >= (27, 0) else 'Cuong'
    if b == 'TNF' and p == 'Cuong': return 'Khanh'          # TNF 東京企画 → Khanh từ 11/2026
    if b == 'TNF' and p in ('Khanh', 'B.Trang', '?'): return 'Luan'
    if p == '?' and b in ('DCC', 'SNOWPEAK', 'ALLTERRAIN', 'DESCENT', 'TETON', 'MONTBELL'): return 'Hien'
    if p == 'Khanh': return 'Luan'
    return p

def chu_mau(p, b, thang, khanh_mau_tnf_tu):
    if p == 'Khanh': return 'Q.Huy' if b == 'GOLDWIN' else 'B.Trang'   # mẫu Khanh năm trước: TNF → B.Trang, Goldwin → Huy
    if b == 'NW': return 'Khanh'
    if b == 'YONEX': return 'B.Trang'
    if p == 'Cuong' and b == 'TNF': return 'Khanh'
    return p

def trong_so_mau(loai):
    l = str(loai or '').upper()
    if l.startswith(('1ST', '0次', '提案', '展開', 'MÂU 0', 'DE AN')): return 1.0, 1.0
    if l.startswith(('2ND', '3RD', '4TH', '5TH')): return 0.8, 0.5
    if 'SMS' in l: return 0.5, 0.2
    if 'PPS' in l or 'MOCK' in l: return 0.6, 0.2
    return 0.3, 0.1

def tai(khanh_mau_tnf=1/3, khanh_mau_tnf_tu='2026-12', giao_them={'Cuong': 0.1, 'Luan': 0.1},
        ho_tro_luan={'2026-11': 0.25, '2026-12': 0.25}, ho_tro_luan_sau=0.35, he_so_kho=True):
    pt = json.load(open('data/khoi-luong/plantable-11-2025-4-2027.json'))['rows']
    mau = json.load(open('data/khoi-luong/mau-2025-2026.json'))
    P = {m: collections.defaultdict(lambda: collections.defaultdict(float)) for m in THANG}
    k_ = lambda b: HE_SO_KHO.get(b, 1.0) if he_so_kho else 1.0
    # lịch sản xuất theo tháng: thật đến 3/2027, sau đó dùng năm trước dời +1 năm
    lich = []
    for x in pt:
        s, e = dt.date.fromisoformat(x['start']), dt.date.fromisoformat(x['end'])
        b = nhom_brand(x['brand']); ms = mua_so(x['mua'])
        if ms <= MUA_CO_KE_HOACH and e >= dt.date(2026, 11, 1):
            lich.append((x, b, s, e, x['mua'], 'thật'))                         # mùa đã có kế hoạch
        if (ms[0] + 1, ms[1]) > MUA_CO_KE_HOACH and ms <= MUA_CO_KE_HOACH:    # 26FW → 27FW, 27SS → 28SS
            mua1 = f"{ms[0] + 1}{'FW' if ms[1] else 'SS'}"
            s1, e1 = s + dt.timedelta(days=364), e + dt.timedelta(days=364)   # dời 52 tuần (giữ thứ trong tuần)
            if e1 >= dt.date(2026, 11, 1) and s1 <= dt.date(2027, 10, 31):
                lich.append((x, b, s1, e1, mua1, 'năm trước'))
    da_vao = set()
    for (x, b, s, e, mua, nguon) in sorted(lich, key=lambda t: t[2]):
        p = chu_san_xuat(x['nguoi'], b, mua)
        if p not in MD: continue
        for m in THANG:
            y, mo = map(int, m.split('-')); a = dt.date(y, mo, 1); z = dt.date(y + (mo == 12), mo % 12 + 1, 1) - dt.timedelta(days=1)
            ov = ngay_xuong(max(s, a), min(e, z)) if s <= z and e >= a else 0
            if ov > 0:
                P[m][p]['ngay_chuyen'] += ov * k_(b)
                P[m][p]['_chay'] += 0; P[m][p].setdefault('ma_set', 0)
            key = (x['ma'], p, nguon)
            if a <= s <= z and key not in da_vao:
                da_vao.add(key)
                P[m][p]['ma_vao'] += k_(b); P[m][p]['dich'] += GIO_DICH.get(b, 2.5) * k_(b); P[m][p]['ma_vao_goc'] += 1
    # mẫu
    def them_mau(m, p, b, loai):
        w, wd = trong_so_mau(loai); n = 1.0
        if p == 'B.Trang' and b == 'TNF' and khanh_mau_tnf and m >= khanh_mau_tnf_tu:
            P[m]['Khanh']['mau'] += n * khanh_mau_tnf * w; P[m]['Khanh']['dich'] += n * khanh_mau_tnf * wd * GIO_DICH['TNF']; P[m]['Khanh']['don_mau'] += khanh_mau_tnf
            n = 1 - khanh_mau_tnf
        P[m][p]['mau'] += n * w * k_(b); P[m][p]['dich'] += n * wd * GIO_DICH.get(b, 2.5) * k_(b); P[m][p]['don_mau'] += n
    for x in mau:
        b = nhom_brand(x['brand'])
        if x['nguoi'] == 'H.Trang': continue
        if x['sheet'] == 'サンプル表2026' and not x['xong'] and '2026-06-01' <= x['due'] < '2026-11-01':
            them_mau('2026-11', chu_mau(x['nguoi'], b, '2026-11', khanh_mau_tnf_tu), b, x['loai'])
        for m in THANG:
            y, mo = map(int, m.split('-'))
            if x['due'][:7] == f'{y - 1}-{mo:02d}':
                p = chu_mau(x['nguoi'], b, m, khanh_mau_tnf_tu)
                if p in MD: them_mau(m, p, b, x['loai'])
    # công suất: ngày làm việc × giờ/ngày (đơn vị: ngày của người làm 9 giờ)
    cs = {m: {p: ngay_lam(m, p) * GIO.get(p, GIO_CHUAN) / GIO_CHUAN for p in MD} for m in THANG}
    yeu_to = ('ma_vao', 'ngay_chuyen', 'mau', 'dich')
    tong_cs = sum(cs[m][p] * (1 - giao_them.get(p, 0)) for m in THANG for p in MD)
    tb = {k: sum(P[m][p][k] for m in THANG for p in MD) / tong_cs for k in yeu_to}   # mức trung bình / 1 ngày công
    out = {}
    for m in THANG:
        tai_ = {p: sum(P[m][p][k] / tb[k] for k in yeu_to) / 4 for p in MD}          # đơn vị: ngày công trung bình
        tl = ho_tro_luan.get(m, ho_tro_luan_sau)
        tai_['Khanh'] += tai_['Luan'] * tl; tai_['Luan'] *= 1 - tl
        out[m] = {p: dict(chi_so=round(tai_[p] / (cs[m][p] * (1 - giao_them.get(p, 0))), 2), ngay_lam=ngay_lam(m, p),
                          ma_vao_chuyen=int(P[m][p]['ma_vao_goc']), ngay_chuyen=round(P[m][p]['ngay_chuyen']),
                          don_mau=round(P[m][p]['don_mau']), gio_dich=round(P[m][p]['dich'])) for p in MD}
        out[m]['_giai_doan'] = giai_doan(m)
    return out

if __name__ == '__main__':
    r = tai()
    print('tháng   ' + ' '.join(p[:7].rjust(7) for p in MD) + '  giai đoạn')
    for m in THANG:
        print(m, ' '.join(f"{r[m][p]['chi_so']:.2f}".rjust(7) for p in MD), ' ', r[m]['_giai_doan'])
    print('-- mã vào chuyền / ngày chuyền / đơn mẫu / ngày làm')
    for m in THANG:
        print(m, ' '.join(f"{r[m][p]['ma_vao_chuyen']}/{r[m][p]['ngay_chuyen']}/{r[m][p]['don_mau']}/{r[m][p]['ngay_lam']}".rjust(14) for p in MD))
    json.dump(r, open('data/khoi-luong/ke-hoach-nam.json', 'w'), ensure_ascii=False, indent=1)
