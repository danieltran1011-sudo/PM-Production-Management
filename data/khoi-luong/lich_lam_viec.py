# Dùng: from lich_lam_viec import ngay_lam, ngay_xuong (chạy script trong data/khoi-luong). Chạy trực tiếp để in số ngày làm việc từng tháng.
# Lịch làm việc TIVN-APIC: thứ 2 – thứ 7, mỗi kỳ lương nghỉ 2 thứ 7 theo nhóm (file 土曜休日 2026, đến 3/2027).
# Sau 3/2027 chưa có lịch: giả định mỗi tháng mỗi nhóm nghỉ 2 thứ 7 (luân phiên B, A, B, A).
# Ngày lễ: chưa có trong file, dùng lịch nghỉ lễ Việt Nam (Tết 2027 là giả định, cần lịch công ty).
import datetime as dt, json, collections, openpyxl

NHOM = {'Cuong': 'B', 'Hien': 'B', 'Luan': 'A', 'Khanh': 'A', 'Nhu': 'A', 'Q.Huy': 'A', 'B.Trang': 'A'}
NGAY_LE = {  # nghỉ lễ (không làm), * = giả định
    '2027-01-01': 'Tết Dương lịch',
    **{f'2027-02-{d:02d}': 'Tết Nguyên đán (giả định 4–12/2)*' for d in range(4, 13)},
    '2027-04-16': 'Giỗ Tổ Hùng Vương (10/3 âm lịch, giả định)*',
    '2027-04-30': 'Giải phóng miền Nam', '2027-05-01': 'Quốc tế Lao động',
    '2027-09-02': 'Quốc khánh', '2027-09-03': 'Quốc khánh (nghỉ thêm)',
}

def doc_lich(path='data/nguon/Lich-nghi-thu-7-TIVN-APIC-2026-20260811.xlsx'):
    ws = openpyxl.load_workbook(path, data_only=True)['土曜休日一覧']
    t7 = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if isinstance(r[0], dt.datetime) and r[5] in ('Làm', 'Nghỉ'):
            t7[r[0].date().isoformat()] = {'A': r[5] == 'Làm', 'B': r[6] == 'Làm'}
    return t7

T7 = doc_lich()

def lam_t7(d, nhom):
    k = d.isoformat()
    if k in T7: return T7[k][nhom]
    # ngoài lịch: trong tháng, thứ 7 thứ 1 và 3 nhóm B nghỉ, thứ 2 và 4 nhóm A nghỉ, thứ 7 thứ 5 cả hai làm
    n = (d.day - 1) // 7 + 1
    if n == 5: return True
    return not ((nhom == 'B' and n in (1, 3)) or (nhom == 'A' and n in (2, 4)))

def ngay_lam(thang, nguoi=None):
    """Số ngày làm việc trong tháng 'YYYY-MM' của một người (theo nhóm thứ 7); nguoi=None: lịch nhà máy (thứ 2–7)."""
    y, m = map(int, thang.split('-')); d = dt.date(y, m, 1); n = 0
    while d.month == m:
        if d.isoformat() not in NGAY_LE and d.weekday() < 6:
            if d.weekday() < 5 or nguoi is None or lam_t7(d, NHOM.get(nguoi, 'A')): n += 1
        d += dt.timedelta(days=1)
    return n

def ngay_xuong(a, b):
    """Số ngày làm việc của xưởng (thứ 2–7, trừ lễ) trong đoạn [a, b]."""
    n = 0; d = a
    while d <= b:
        if d.weekday() < 6 and d.isoformat() not in NGAY_LE: n += 1
        d += dt.timedelta(days=1)
    return n

if __name__ == '__main__':
    TH = ['2026-11', '2026-12', '2027-01', '2027-02', '2027-03', '2027-04', '2027-05', '2027-06', '2027-07', '2027-08', '2027-09', '2027-10']
    print('tháng   ', ' '.join(t[5:] for t in TH))
    for p in ['Cuong', 'Hien', 'Luan']:
        print(p.ljust(8), ' '.join(str(ngay_lam(t, p)).rjust(2) for t in TH))
    print('xưởng   ', ' '.join(str(ngay_lam(t)).rjust(2) for t in TH))
