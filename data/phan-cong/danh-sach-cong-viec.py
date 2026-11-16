# Tạo file Excel "Danh sách công việc đang phụ trách" cho một MD, điền sẵn từ Plantable và file mẫu để người đó xác nhận, bổ sung.
# Chạy từ thư mục gốc repo:
#   python3 data/phan-cong/danh-sach-cong-viec.py Hien "Hiền" 2026-11-16 docs/tai-ve/Danh-sach-cong-viec-Hien-20261116.xlsx
# Tham số: tên trong dữ liệu (Hien, Cuong, Luan, Khanh, B.Trang, Q.Huy, Nhu), tên hiển thị, ngày lập, file xuất.
import json, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

ma_nguoi, ten, ngay, out = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
hom_nay = dt.date.fromisoformat(ngay)
pt = json.load(open('data/khoi-luong/plantable-11-2025-4-2027.json'))['rows']
mau = json.load(open('data/khoi-luong/mau-2025-2026.json'))

F = 'Arial'
HEAD = PatternFill('solid', fgColor='DCE6F1'); DIEN = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='999999'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()

def sheet(title, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    return ws

def bang(ws, r0, cols, rows, n_dien, them_dong=0):
    """cols: [(tên cột, độ rộng)]; n_dien: số cột cuối để người phụ trách điền (tô vàng)."""
    for j, (c, w) in enumerate(cols, 1):
        x = ws.cell(r0, j, c); x.font = Font(name=F, bold=True, size=10); x.fill = DIEN if j > len(cols) - n_dien else HEAD
        x.alignment = Alignment(wrap_text=True, vertical='center'); x.border = BOX
        ws.column_dimensions[L(j)].width = w
    for i, row in enumerate(list(rows) + [[None] * (len(cols) - n_dien)] * them_dong, r0 + 1):
        for j in range(1, len(cols) + 1):
            v = row[j - 1] if j - 1 < len(row) else None
            x = ws.cell(i, j, v); x.font = Font(name=F, size=10); x.border = BOX
            x.alignment = Alignment(wrap_text=True, vertical='top')
            if j > len(cols) - n_dien: x.fill = DIEN
    ws.freeze_panes = ws.cell(r0 + 1, 1)
    return r0 + len(rows) + them_dong

def tieu_de(ws, t, ghi_chu):
    ws.cell(1, 1, t).font = Font(name=F, bold=True, size=13)
    ws.cell(2, 1, ghi_chu).font = Font(name=F, italic=True, size=9, color='555555')

# 0. Hướng dẫn
ws = sheet('Hướng dẫn', True)
ws.column_dimensions['A'].width = 110
dong = [
    (f'Danh sách công việc đang phụ trách – {ten}', True),
    (f'Ngày lập: {hom_nay:%d/%m/%Y}', False),
    ('', False),
    ('Mục đích', True),
    ('Nắm đúng khối lượng việc đang làm để sắp xếp lại việc trong nhóm trong thời gian tới (có thành viên nghỉ thai sản, lịch đơn hàng thay đổi).', False),
    ('Đây không phải bảng đánh giá. Số liệu giúp quản lý chia việc hợp lý và đề xuất hỗ trợ khi cần.', False),
    ('', False),
    ('Cách điền', True),
    ('1. Các cột nền xanh đã điền sẵn từ Plantable (16/11/2026) và file mẫu (13/11/2026). Nếu sai hoặc thiếu, sửa trực tiếp hoặc thêm dòng.', False),
    ('2. Điền các cột nền vàng: giai đoạn hiện tại, việc còn lại, vấn đề đang vướng, số giờ ước mỗi tuần.', False),
    ('3. Lô / đơn đã xong: chọn giai đoạn "Đã xong", không cần điền thêm.', False),
    ('4. Sheet "Việc khác": các việc không theo mã (báo cáo, họp, liên lạc, xử lý sự cố, hỗ trợ người khác…). Ghi số giờ ước mỗi tuần.', False),
    ('5. Giờ ước mức 30 phút là đủ, không cần chính xác.', False),
    ('', False),
    ('Hạn gửi lại: ……/……/2026', True),
]
for i, (t, b) in enumerate(dong, 1):
    c = ws.cell(i, 1, t); c.font = Font(name=F, bold=b, size=12 if i == 1 else 10); c.alignment = Alignment(wrap_text=True)

GD_SX = '"Chuẩn bị (tài liệu, NPL),PPS / duyệt mẫu,Đang may,Hoàn thiện / kiểm hàng,Xuất hàng,Đã xong"'
GD_MAU = '"Chờ tài liệu,Đang chuẩn bị,Đang may,Gửi khách / chờ góp ý,Đã xong"'

# 1. Đơn hàng sản xuất
ws = sheet('Đơn hàng sản xuất')
tieu_de(ws, f'Đơn hàng sản xuất – {ten}', 'Lô có ngày kết thúc từ 1/11/2026 trở đi (Plantable 16/11/2026). Cột vàng: người phụ trách điền.')
rows = sorted([x for x in pt if x['nguoi'] == ma_nguoi and x['end'] >= '2026-11-01'], key=lambda x: (x['start'], x['brand']))
cols = [('Brand', 14), ('Mùa', 7), ('Hợp đồng', 13), ('Mã hàng', 18), ('Số lượng', 9), ('Xưởng / line', 10), ('Bắt đầu may', 11), ('Kết thúc', 11),
        ('Giai đoạn hiện tại', 18), ('Việc còn lại', 30), ('Vấn đề đang vướng (NPL, rập, tài liệu…)', 30), ('Giờ / tuần (ước)', 9)]
data = [[x['brand'], x['mua'], x['hd'], x['ma'], x['sl'], f"{x['xuong']}-{x['line']}", dt.date.fromisoformat(x['start']), dt.date.fromisoformat(x['end'])] for x in rows]
end = bang(ws, 4, cols, data, 4, them_dong=10)
for r in range(5, end + 1):
    for c in (7, 8): ws.cell(r, c).number_format = 'DD/MM/YYYY'
dv = DataValidation(type='list', formula1=GD_SX, allow_blank=True); ws.add_data_validation(dv); dv.add(f'I5:I{end}')
n_sx = len(data)

# 2. Mẫu
ws = sheet('Mẫu')
tieu_de(ws, f'Đơn mẫu – {ten}', 'Đơn mẫu chưa xong trong file mẫu (sheet サンプル表2026), cộng các đơn mẫu khác đang làm (thêm dòng). Cột vàng: người phụ trách điền.')
rows = sorted([x for x in mau if x['nguoi'] == ma_nguoi and not x['xong'] and x['sheet'] == 'サンプル表2026'], key=lambda x: (x['due'], x['brand']))
cols = [('Brand', 16), ('Mùa', 7), ('Loại mẫu', 14), ('Số lượng', 9), ('Hạn', 11),
        ('Giai đoạn hiện tại', 18), ('Việc còn lại', 30), ('Vấn đề đang vướng', 30), ('Giờ / tuần (ước)', 9)]
data = [[x['brand'], x['mua'], x['loai'], x['sl'], dt.date.fromisoformat(x['due'])] for x in rows]
end = bang(ws, 4, cols, data, 4, them_dong=15)
for r in range(5, end + 1): ws.cell(r, 5).number_format = 'DD/MM/YYYY'
dv = DataValidation(type='list', formula1=GD_MAU, allow_blank=True); ws.add_data_validation(dv); dv.add(f'F5:F{end}')
n_mau = len(data)

# 3. Việc khác
ws = sheet('Việc khác')
tieu_de(ws, f'Việc không theo mã – {ten}', 'Các việc định kỳ hoặc phát sinh không gắn với một mã hàng. Sửa, xóa hoặc thêm dòng cho đúng thực tế.')
cols = [('Việc', 42), ('Tần suất (ngày / tuần / tháng / phát sinh)', 18), ('Giờ / tuần (ước)', 10), ('Ghi chú', 40)]
goi_y = [['Báo cáo tuần / tháng'], ['Họp (nội bộ, với TIJ, với nhà máy)'], ['Dịch tài liệu, email ngoài phần mã hàng'],
         ['Liên lạc TIJ / khách hàng (hỏi, xác nhận)'], ['Xử lý sự cố (vải lỗi, NPL thiếu, kỹ thuật)'], ['Theo dõi xuất hàng, chứng từ'],
         ['Hỗ trợ người khác trong nhóm'], ['Khác: …']]
bang(ws, 4, cols, goi_y, 3, them_dong=4)

# 4. Tổng hợp (công thức)
ws = sheet('Tổng hợp')
tieu_de(ws, f'Tổng hợp – {ten}', 'Tự tính từ các sheet khác.')
items = [('Số lô sản xuất (điền sẵn)', n_sx), ('Số lô chưa xong', "=COUNTA('Đơn hàng sản xuất'!D5:D500)-COUNTIF('Đơn hàng sản xuất'!I5:I500,\"Đã xong\")"),
         ('Số đơn mẫu (điền sẵn)', n_mau), ('Số đơn mẫu chưa xong', "=COUNTA('Mẫu'!A5:A500)-COUNTIF('Mẫu'!F5:F500,\"Đã xong\")"),
         ('Giờ / tuần – đơn hàng sản xuất', "=SUM('Đơn hàng sản xuất'!L5:L500)"), ('Giờ / tuần – mẫu', "=SUM('Mẫu'!I5:I500)"),
         ('Giờ / tuần – việc khác', "=SUM('Việc khác'!C5:C500)"), ('Tổng giờ / tuần', '=SUM(B8:B10)'),
         ('Giờ làm chuẩn / tuần (5,5 ngày × 8 giờ, tham khảo)', 44)]
for i, (k, v) in enumerate(items, 4):
    a = ws.cell(i, 1, k); b = ws.cell(i, 2, v)
    for x in (a, b): x.font = Font(name=F, size=10, bold=(k == 'Tổng giờ / tuần')); x.border = BOX
    if isinstance(v, (int, float)): b.font = Font(name=F, size=10, color='0000FF')
ws.column_dimensions['A'].width = 48; ws.column_dimensions['B'].width = 12

wb.save(out)
print('ok', out, f'{n_sx} lô sản xuất, {n_mau} đơn mẫu')
