# Tạo file Excel "Danh sách công việc đang phụ trách" cho từng MD, điền sẵn từ Plantable và file mẫu để người đó xác nhận, bổ sung.
# Người phụ trách lấy theo phân công mới nhất (bảng chốt 5/11/2026, quy tắc trong data/khoi-luong/ke-hoach-nam.py), không theo tên ghi trong Plantable.
# Chạy từ thư mục gốc repo:
#   python3 data/phan-cong/danh-sach-cong-viec.py 2026-11-16                      # cả nhóm MD → docs/tai-ve/danh-sach-cong-viec/
#   python3 data/phan-cong/danh-sach-cong-viec.py 2026-11-16 docs/tai-ve/danh-sach-cong-viec Hien Khanh   # vài người
#   thêm --ja để xuất bản tiếng Nhật (Danh-sach-cong-viec-<tên>-<ngày>-JA.xlsx; không tạo file tổng hợp)
import json, sys, os, re, collections, importlib.util, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

ARGS = [a for a in sys.argv[1:] if not a.startswith('--')]
JA = '--ja' in sys.argv                     # xuất bản tiếng Nhật (tên file thêm -JA)
ngay = ARGS[0]
OUT = ARGS[1] if len(ARGS) > 1 else 'docs/tai-ve/danh-sach-cong-viec'
TEN = {'Cuong': ('Cường', 'Cuong'), 'Luan': ('Luân', 'Luan'), 'Khanh': ('Khánh', 'Khanh'), 'B.Trang': ('Bích Trang', 'Bich-Trang'),
       'Q.Huy': ('Huy', 'Huy'), 'Hien': ('Hiền', 'Hien'), 'Nhu': ('Như', 'Nhu')}
NGUOI = ARGS[2:] or list(TEN)
TEN_JA = {'Cuong': 'クォン（Cuong）', 'Luan': 'ルァン（Luan）', 'Khanh': 'カィン（Khanh）', 'B.Trang': 'チャン（Bich Trang）',
          'Q.Huy': 'フィ（Huy）', 'Hien': 'ヒィン（Hien）', 'Nhu': 'ヌ（Nhu）'}
hom_nay = dt.date.fromisoformat(ngay)
pt = json.load(open('data/khoi-luong/plantable-11-2025-4-2027.json'))['rows']
mau = json.load(open('data/khoi-luong/mau-2025-2026.json'))
spec = importlib.util.spec_from_file_location('khn', 'data/khoi-luong/ke-hoach-nam.py'); kh = importlib.util.module_from_spec(spec); spec.loader.exec_module(kh)
ma_ = lambda m: re.sub(r'[\s\.]+$', '', str(m or '').strip().upper())
hien_thi = lambda p: TEN.get(p, (p,))[0] if p != '?' else 'chưa rõ'

# Phân công lô sản xuất (lô kết thúc từ 1/11/2026)
SX = []
for x in pt:
    if x['end'] < '2026-11-01': continue
    p = kh.chu_san_xuat(x['nguoi'], kh.nhom_brand(x['brand']), x['mua'], x.get('tij'))
    SX.append(dict(x, moi=p, ghi='' if p == x['nguoi'] else f"Nhận từ {hien_thi(x['nguoi'])}"))
MA_SGM = {ma_(x['ma']) for x in pt if kh.nhom_brand(x['brand']) == 'TNF' and x.get('tij') == 'SGM'}

# Phân công đơn mẫu còn mở (sheet サンプル表2026): người đang làm làm tiếp, trừ các phần đã chuyển
MAU = []
for x in mau:
    if x['xong'] or x['sheet'] != 'サンプル表2026': continue
    b, p = kh.nhom_brand(x['brand']), x['nguoi']
    if b == '141': p = 'Khanh'
    elif p == 'H.Trang': continue                                           # nhóm CAD
    elif p == 'Luan' and b == 'TNF' and ma_(x.get('ma')) in MA_SGM: p = 'Khanh'   # TNF 富山企画 phần SGM
    elif (b == 'NW' or (p == 'Cuong' and b == 'TNF')) and kh.mua_so(x['mua']) >= (28, 0): p = 'Khanh'   # NW, TNF 東京企画 từ mùa sau 27FW
    MAU.append(dict(x, moi=p, ghi='' if p == x['nguoi'] else f"Nhận từ {hien_thi(x['nguoi'])}"))

F = 'Meiryo UI' if JA else 'Arial'
HEAD = PatternFill('solid', fgColor='DCE6F1'); DIEN = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='999999'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

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


JA_TU = {
    'Mục đích': '目的', 'Cách điền': '記入方法',
    'Nắm đúng khối lượng việc đang làm để sắp xếp lại việc trong nhóm trong thời gian tới (có thành viên nghỉ thai sản, lịch đơn hàng thay đổi).':
        '今後のチーム内の業務再配分（産休者の発生、受注スケジュールの変更）に向けて、現在の業務量を正確に把握するための資料です。',
    'Đây không phải bảng đánh giá. Số liệu giúp quản lý chia việc hợp lý và đề xuất hỗ trợ khi cần.':
        '評価のための表ではありません。業務を適切に配分し、必要に応じて支援を検討するために使います。',
    '1. Các cột nền xanh đã điền sẵn từ Plantable (16/11/2026) và file mẫu (13/11/2026), theo phân công brand chốt 5/11/2026. Cột "Phân công" ghi lô / đơn nhận từ người khác. Nếu sai hoặc thiếu, sửa trực tiếp hoặc thêm dòng.':
        '1. 青色の列は、Plantable（2026/11/16）とサンプル表（2026/11/13）から、2026/11/5確定のブランド担当に基づいて事前入力しています。「担当変更」列は他の担当者から引き継いだロット・オーダーです。誤りや漏れがあれば直接修正するか、行を追加してください。',
    '2. Điền các cột nền vàng: giai đoạn hiện tại, việc còn lại, vấn đề đang vướng, số giờ ước mỗi tuần.':
        '2. 黄色の列（現在の工程、残作業、課題、週あたりの時間）を記入してください。',
    '3. Lô / đơn đã xong: chọn giai đoạn "Đã xong", không cần điền thêm.':
        '3. 完了したロット・オーダーは、工程で「完了」を選択してください。他の記入は不要です。',
    '4. Sheet "Việc khác": các việc không theo mã (báo cáo, họp, liên lạc, xử lý sự cố, hỗ trợ người khác…). Ghi số giờ ước mỗi tuần.':
        '4. 「その他業務」シート：品番に紐づかない業務（報告、会議、連絡、トラブル対応、他メンバーの支援など）。週あたりの時間を記入してください。',
    '5. Giờ ước mức 30 phút là đủ, không cần chính xác.': '5. 時間は30分単位の概算で構いません。',
    'Hạn gửi lại: ……/……/2026': '提出期限：2026年……月……日',
    'Lô có ngày kết thúc từ 1/11/2026 trở đi (Plantable 16/11/2026). Cột vàng: người phụ trách điền.':
        '2026/11/1以降に縫製終了するロット（Plantable 2026/11/16）。黄色の列：担当者が記入。',
    'Đơn mẫu chưa xong trong file mẫu (sheet サンプル表2026), cộng các đơn mẫu khác đang làm (thêm dòng). Cột vàng: người phụ trách điền.':
        'サンプル表（サンプル表2026）の未完了オーダー。その他進行中のサンプルは行を追加してください。黄色の列：担当者が記入。',
    'Các việc định kỳ hoặc phát sinh không gắn với một mã hàng. Sửa, xóa hoặc thêm dòng cho đúng thực tế.':
        '品番に紐づかない定常・臨時の業務。実態に合わせて修正・削除・追加してください。',
    'Tự tính từ các sheet khác.': '他のシートから自動計算します。',
    'Brand': 'ブランド', 'Mùa': 'シーズン', 'Hợp đồng': '契約番号', 'Mã hàng': '品番', 'Số lượng': '数量', 'Line': 'ライン',
    'Bắt đầu may': '縫製開始', 'Kết thúc': '縫製終了', 'Phân công': '担当変更', 'Giai đoạn hiện tại': '現在の工程', 'Việc còn lại': '残作業',
    'Vấn đề đang vướng (NPL, rập, tài liệu…)': '課題（資材・パターン・資料など）', 'Vấn đề đang vướng': '課題', 'Giờ / tuần (ước)': '時間／週（概算）',
    'Tên hàng': '品名', 'Loại mẫu': 'サンプル種類', 'Hạn': '納期',
    'Việc': '業務', 'Tần suất (ngày / tuần / tháng / phát sinh)': '頻度（日次／週次／月次／不定期）', 'Ghi chú': '備考',
    'Báo cáo tuần / tháng': '週次・月次報告', 'Họp (nội bộ, với TIJ, với nhà máy)': '会議（社内・TIJ・工場）',
    'Dịch tài liệu, email ngoài phần mã hàng': '品番以外の翻訳（資料・メール）', 'Liên lạc TIJ / khách hàng (hỏi, xác nhận)': 'TIJ・顧客との連絡（問い合わせ・確認）',
    'Xử lý sự cố (vải lỗi, NPL thiếu, kỹ thuật)': 'トラブル対応（生地不良・資材不足・技術）', 'Theo dõi xuất hàng, chứng từ': '出荷フォロー・書類',
    'Hỗ trợ người khác trong nhóm': 'チーム内の他メンバーの支援', 'Khác: …': 'その他：…',
    'Số lô sản xuất (điền sẵn)': '量産ロット数（事前入力）', 'Số lô chưa xong': '未完了ロット数', 'Số đơn mẫu (điền sẵn)': 'サンプルオーダー数（事前入力）',
    'Số đơn mẫu chưa xong': '未完了サンプル数', 'Giờ / tuần – đơn hàng sản xuất': '時間／週 – 量産', 'Giờ / tuần – mẫu': '時間／週 – サンプル',
    'Giờ / tuần – việc khác': '時間／週 – その他業務', 'Tổng giờ / tuần': '合計時間／週', 'Giờ làm chuẩn / tuần (5,5 ngày × 8 giờ, tham khảo)': '標準勤務時間／週（5.5日×8時間、参考）',
}
JA_SHEET = {'Hướng dẫn': '説明', 'Đơn hàng sản xuất': '量産オーダー', 'Mẫu': 'サンプル', 'Việc khác': 'その他業務', 'Tổng hợp': '集計'}
JA_TIEUDE = {'Danh sách công việc đang phụ trách': '担当業務リスト', 'Đơn hàng sản xuất': '量産オーダー', 'Đơn mẫu': 'サンプル',
             'Việc không theo mã': 'その他業務', 'Tổng hợp': '集計'}
JA_CHON = {'Chuẩn bị (tài liệu / NPL)': '準備（資料・資材）', 'PPS / duyệt mẫu': 'PPS・サンプル承認', 'Đang may': '縫製中',
           'Hoàn thiện / kiểm hàng': '仕上げ・検品', 'Xuất hàng': '出荷', 'Đã xong': '完了', 'Chờ tài liệu': '資料待ち',
           'Đang chuẩn bị': '準備中', 'Gửi khách / chờ góp ý': '客先送付・コメント待ち'}

def dich_ja(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if c.number_format == 'DD/MM/YYYY': c.number_format = 'YYYY/MM/DD'   # cách ghi ngày của Nhật
                if not isinstance(v, str): continue
                if v.startswith('='):
                    for a, b in list(JA_SHEET.items()) + list(JA_CHON.items()): v = v.replace(f"'{a}'!", f"'{b}'!").replace(f'"{a}"', f'"{b}"')
                elif v in JA_TU: v = JA_TU[v]
                elif v.startswith('Ngày lập: '): v = f'作成日：{hom_nay:%Y/%m/%d}'
                elif v.startswith('Nhận từ '): v = v[len('Nhận từ '):] + 'から引継ぎ'
                elif re.fullmatch(r'(Áo|Quần) (\d) lớp', v): v = re.sub(r'(Áo|Quần) (\d) lớp', lambda m: ('上衣' if m[1] == 'Áo' else 'パンツ') + f'（{m[2]}層）', v)
                elif ' – ' in v and v.split(' – ')[0] in JA_TIEUDE: v = JA_TIEUDE[v.split(' – ')[0]] + ' – ' + v.split(' – ', 1)[1]
                c.value = v
        for dv in ws.data_validations.dataValidation:
            ds = dv.formula1.strip('"').split(',')
            dv.formula1 = '"' + ','.join(JA_CHON.get(x, x) for x in ds) + '"'
    for ws in wb.worksheets: ws.title = JA_SHEET.get(ws.title, ws.title)

def tao(ma_nguoi, ten, out):
    wb = Workbook()
    def sheet(title, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        return ws
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
        ('1. Các cột nền xanh đã điền sẵn từ Plantable (16/11/2026) và file mẫu (13/11/2026), theo phân công brand chốt 5/11/2026. Cột "Phân công" ghi lô / đơn nhận từ người khác. Nếu sai hoặc thiếu, sửa trực tiếp hoặc thêm dòng.', False),
        ('2. Điền các cột nền vàng: giai đoạn hiện tại, việc còn lại, vấn đề đang vướng, số giờ ước mỗi tuần.', False),
        ('3. Lô / đơn đã xong: chọn giai đoạn "Đã xong", không cần điền thêm.', False),
        ('4. Sheet "Việc khác": các việc không theo mã (báo cáo, họp, liên lạc, xử lý sự cố, hỗ trợ người khác…). Ghi số giờ ước mỗi tuần.', False),
        ('5. Giờ ước mức 30 phút là đủ, không cần chính xác.', False),
        ('', False),
        ('Hạn gửi lại: ……/……/2026', True),
    ]
    for i, (t, b) in enumerate(dong, 1):
        c = ws.cell(i, 1, t); c.font = Font(name=F, bold=b, size=12 if i == 1 else 10); c.alignment = Alignment(wrap_text=True)

    GD_SX = '"Chuẩn bị (tài liệu / NPL),PPS / duyệt mẫu,Đang may,Hoàn thiện / kiểm hàng,Xuất hàng,Đã xong"'
    GD_MAU = '"Chờ tài liệu,Đang chuẩn bị,Đang may,Gửi khách / chờ góp ý,Đã xong"'

    # 1. Đơn hàng sản xuất
    ws = sheet('Đơn hàng sản xuất')
    tieu_de(ws, f'Đơn hàng sản xuất – {ten}', 'Lô có ngày kết thúc từ 1/11/2026 trở đi (Plantable 16/11/2026). Cột vàng: người phụ trách điền.')
    rows = sorted([x for x in SX if x['moi'] == ma_nguoi], key=lambda x: (x['start'], x['brand']))
    cols = [('Brand', 14), ('Mùa', 7), ('Hợp đồng', 13), ('Mã hàng', 18), ('Số lượng', 9), ('Part', 6), ('Line', 6), ('Bắt đầu may', 11), ('Kết thúc', 11), ('Phân công', 14),
            ('Giai đoạn hiện tại', 18), ('Việc còn lại', 30), ('Vấn đề đang vướng (NPL, rập, tài liệu…)', 30), ('Giờ / tuần (ước)', 9)]
    data = [[x['brand'], x['mua'], x['hd'], x['ma'], x['sl'], x['xuong'].replace('AP', ''), x['line'], dt.date.fromisoformat(x['start']), dt.date.fromisoformat(x['end']), x['ghi']] for x in rows]
    end = bang(ws, 4, cols, data, 4, them_dong=10)
    for r in range(5, end + 1):
        for c in (8, 9): ws.cell(r, c).number_format = 'DD/MM/YYYY'
    dv = DataValidation(type='list', formula1=GD_SX, allow_blank=True); ws.add_data_validation(dv); dv.add(f'K5:K{end}')
    n_sx = len(data)

    # 2. Mẫu
    ws = sheet('Mẫu')
    tieu_de(ws, f'Đơn mẫu – {ten}', 'Đơn mẫu chưa xong trong file mẫu (sheet サンプル表2026), cộng các đơn mẫu khác đang làm (thêm dòng). Cột vàng: người phụ trách điền.')
    rows = [x for x in MAU if x['moi'] == ma_nguoi]
    rows = sorted(rows, key=lambda x: (x['due'], x['brand'], x.get('ma', '')))
    cols = [('Brand', 16), ('Mùa', 7), ('Mã hàng', 16), ('Tên hàng', 22), ('Loại mẫu', 12), ('Số lượng', 9), ('Hạn', 11), ('Phân công', 14),
            ('Giai đoạn hiện tại', 18), ('Việc còn lại', 30), ('Vấn đề đang vướng', 30), ('Giờ / tuần (ước)', 9)]
    data = [[x['brand'], x['mua'], x.get('ma', ''), x.get('ten', ''), x['loai'], x['sl'], dt.date.fromisoformat(x['due']), x['ghi']] for x in rows]
    end = bang(ws, 4, cols, data, 4, them_dong=15)
    for r in range(5, end + 1): ws.cell(r, 7).number_format = 'DD/MM/YYYY'
    dv = DataValidation(type='list', formula1=GD_MAU, allow_blank=True); ws.add_data_validation(dv); dv.add(f'I5:I{end}')
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
    items = [('Số lô sản xuất (điền sẵn)', n_sx), ('Số lô chưa xong', "=COUNTA('Đơn hàng sản xuất'!D5:D500)-COUNTIF('Đơn hàng sản xuất'!K5:K500,\"Đã xong\")"),
             ('Số đơn mẫu (điền sẵn)', n_mau), ('Số đơn mẫu chưa xong', "=COUNTA('Mẫu'!A5:A500)-COUNTIF('Mẫu'!I5:I500,\"Đã xong\")"),
             ('Giờ / tuần – đơn hàng sản xuất', "=SUM('Đơn hàng sản xuất'!N5:N500)"), ('Giờ / tuần – mẫu', "=SUM('Mẫu'!L5:L500)"),
             ('Giờ / tuần – việc khác', "=SUM('Việc khác'!C5:C500)"), ('Tổng giờ / tuần', '=SUM(B8:B10)'),
             ('Giờ làm chuẩn / tuần (5,5 ngày × 8 giờ, tham khảo)', 44)]
    for i, (k, v) in enumerate(items, 4):
        a = ws.cell(i, 1, k); b = ws.cell(i, 2, v)
        for x in (a, b): x.font = Font(name=F, size=10, bold=(k == 'Tổng giờ / tuần')); x.border = BOX
        if isinstance(v, (int, float)): b.font = Font(name=F, size=10, color='0000FF')
    ws.column_dimensions['A'].width = 48; ws.column_dimensions['B'].width = 12

    if JA: dich_ja(wb)
    wb.save(out)
    return n_sx, n_mau

os.makedirs(OUT, exist_ok=True)
tong = []
for p in NGUOI:
    ten, tep = TEN[p]
    if JA: ten = TEN_JA[p]
    out = f"{OUT}/Danh-sach-cong-viec-{tep}-{hom_nay:%Y%m%d}{'-JA' if JA else ''}.xlsx"
    n_sx, n_mau = tao(p, ten, out)
    nhan = sum(1 for x in SX if x['moi'] == p and x['ghi']) + sum(1 for x in MAU if x['moi'] == p and x['ghi'])
    giao = sum(1 for x in SX if x['nguoi'] == p and x['moi'] != p) + sum(1 for x in MAU if x['nguoi'] == p and x['moi'] != p)
    tong.append((ten, n_sx, n_mau, nhan, giao, os.path.basename(out)))
    print('ok', out, f'{n_sx} lô sản xuất, {n_mau} đơn mẫu, nhận {nhan}, chuyển đi {giao}')

# File tổng hợp cho quản lý: số lô, số đơn mẫu từng người; lô chưa rõ người phụ trách
if JA: sys.exit(0)
wb = Workbook(); ws = wb.active; ws.title = 'Tổng hợp'
ws.cell(1, 1, f'Danh sách công việc nhóm MD – tổng hợp ({hom_nay:%d/%m/%Y})').font = Font(name=F, bold=True, size=13)
ws.cell(2, 1, 'Lô sản xuất kết thúc từ 1/11/2026 (Plantable 16/11/2026), đơn mẫu còn mở (file mẫu 13/11/2026), theo phân công chốt 5/11/2026.').font = Font(name=F, italic=True, size=9, color='555555')
for j, (c, w) in enumerate([('Người', 14), ('Lô sản xuất', 12), ('Đơn mẫu còn mở', 14), ('Lô / đơn nhận từ người khác', 16), ('Lô / đơn chuyển cho người khác', 16), ('File', 40)], 1):
    x = ws.cell(4, j, c); x.font = Font(name=F, bold=True, size=10); x.fill = HEAD; x.border = BOX; x.alignment = Alignment(wrap_text=True)
    ws.column_dimensions[L(j)].width = w
for i, r in enumerate(tong, 5):
    for j, v in enumerate(r, 1):
        x = ws.cell(i, j, v); x.font = Font(name=F, size=10); x.border = BOX
n = 5 + len(tong)
ws.cell(n, 1, 'Tổng').font = Font(name=F, bold=True, size=10)
for j in (2, 3): ws.cell(n, j, f'=SUM({L(j)}5:{L(j)}{n - 1})').font = Font(name=F, bold=True, size=10)
ws2 = wb.create_sheet('Chưa rõ người phụ trách')
ws2.cell(1, 1, 'Lô sản xuất chưa gán được MD (Plantable không ghi, không có mẫu PPS, brand chưa có quy tắc)').font = Font(name=F, bold=True, size=12)
cr = sorted([x for x in SX if x['moi'] not in TEN], key=lambda x: (x['brand'], x['start']))
for j, (c, w) in enumerate([('Brand', 16), ('Mùa', 7), ('TIJ', 6), ('Hợp đồng', 13), ('Mã hàng', 18), ('Số lượng', 9), ('Part', 6), ('Line', 6), ('Bắt đầu', 11), ('Kết thúc', 11), ('Người phụ trách (quản lý điền)', 18)], 1):
    x = ws2.cell(3, j, c); x.font = Font(name=F, bold=True, size=10); x.fill = DIEN if j == 11 else HEAD; x.border = BOX
    ws2.column_dimensions[L(j)].width = w
for i, x in enumerate(cr, 4):
    for j, v in enumerate([x['brand'], x['mua'], x.get('tij'), x['hd'], x['ma'], x['sl'], x['xuong'].replace('AP', ''), x['line'], dt.date.fromisoformat(x['start']), dt.date.fromisoformat(x['end']), None], 1):
        c = ws2.cell(i, j, v); c.font = Font(name=F, size=10); c.border = BOX
        if j in (9, 10): c.number_format = 'DD/MM/YYYY'
        if j == 11: c.fill = DIEN
ws2.freeze_panes = 'A4'
out = f"{OUT}/Tong-hop-danh-sach-cong-viec-MD-{hom_nay:%Y%m%d}.xlsx"; wb.save(out)
print('ok', out, f'{len(cr)} lô chưa rõ người phụ trách')
