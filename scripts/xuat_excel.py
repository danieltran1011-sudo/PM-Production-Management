"""Xuất file Excel lead time sản xuất (song ngữ Nhật–Việt) để thảo luận với TIJ.

Chạy: python3 scripts/xuat_excel.py [đường_dẫn_ra.xlsx]
Mặc định ghi ra docs/Leadtime-TIJ.xlsx. Mọi số ngày đều là công thức Excel tính từ
sheet 標準日数 (thời gian chuẩn), cùng quy tắc với workflow/schedule.js.
"""
import sys
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "docs/Leadtime-TIJ.xlsx"

FONT = "Arial"
INK = "17203B"
MUTED = "5B6482"
ACCENT = "2B3F9E"
ACCENT_SOFT = "E6EAFA"
HEAD = "EEF0F6"
GOOD = "E1F3E9"
GOOD_INK = "1D7A4B"
BAD = "FBE6E4"
BAD_INK = "B3261E"
INPUT_FILL = "FFF7CC"
DEPT = {
    "bocvac": "8C6A3F", "qckho": "13827C", "cat": "3949AB", "kythuat": "A8436F",
    "giacong": "B87A0B", "chuyen": "2D8653", "xnk": "5B6B8C",
}

thin = Side(style="thin", color="D8DDE8")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def put(ws, ref, value, font=None, fill_=None, align=None, border=True, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = font or f()
    if fill_:
        c.fill = fill(fill_)
    if align:
        c.alignment = align
    if border:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


def title(ws, text_ja, text_vi):
    ws["A1"] = text_ja
    ws["A1"].font = f(16, True, ACCENT)
    ws["A2"] = text_vi
    ws["A2"].font = f(11, False, MUTED)


wb = Workbook()

# ---------------------------------------------------------------- 標準日数
std = wb.active
std.title = "標準日数"
title(std, "標準日数（稼働日）", "Thời gian chuẩn (ngày làm việc) — ô nền vàng, chữ xanh là ô nhập; sửa ở đây thì mọi sheet tự tính lại")
heads = ["キー\nKhóa", "工程 / Bước", "部署 / Bộ phận", "< 8,000秒\n< 8.000s", "8,000–12,000秒\n8.000–12.000s",
         "> 12,000秒\n> 12.000s", "根拠 / Căn cứ"]
for i, h in enumerate(heads):
    put(std, f"{get_column_letter(i + 1)}4", h, f(10, True), HEAD, CENTER)
STD_ROWS = [
    ("xnk", "通関・工場への輸送（暦日）\nThông quan, chuyển về nhà máy (ngày lịch)", "輸出入 / XNK", [7, 7, 7], "最大7日（赤線含む）/ Tối đa 7 ngày kể cả luồng đỏ"),
    ("khui", "荷降ろし・開梱検品\nXuống hàng, khui thùng kiểm tra", "荷役 / Bốc vác", [1, 1, 1], "会議 / Họp"),
    ("qc", "品質検査・有効幅確認\nKiểm chất lượng, báo khổ", "倉庫品質管理 / Chất lượng kho", [1, 1, 1], "会議 / Họp"),
    ("test", "生地テスト（収縮率）\nTest vải (co rút)", "技術課 / Kỹ thuật", [1, 1, 1], "会議：テスト＋型紙微調整 1〜2日 / Họp: test + chỉnh rập 1–2 ngày"),
    ("xa", "放反（リラックス）\nXả vải", "裁断課 / Phòng cắt", [1, 1, 1], "会議 / Họp"),
    ("nlTruocCat", "原材料準備完了（裁断開始の何日前）\nNguyên liệu sẵn sàng trước khi cắt (ngày)", "裁断課 / Phòng cắt", [4, 4, 4], "会議：最低4日前 / Họp: tối thiểu 4 ngày"),
    ("rap", "型紙調整\nĐiều chỉnh rập", "技術課 / Kỹ thuật", [1, 2, 2], "会議：12,000秒超は型紙＋マーカー最低3日 / Họp: trên 12.000s rập + sơ đồ ≥ 3 ngày"),
    ("sodo", "マーカー作成・型紙出力（型紙調整2日目から）\nĐi sơ đồ + in rập (từ ngày thứ 2 của chỉnh rập)", "技術課 / Kỹ thuật", [1, 2, 2], "会議 / Họp"),
    ("caitien", "改良パターン（裁断と並行）\nRập cải tiến (song song với cắt)", "技術課 / Kỹ thuật", [3, 3, 3], "Excel 20260722"),
    ("cat", "裁断・ナンバリング・束ね\nCắt, đánh số, bó", "裁断課 / Phòng cắt", [2, 2, 2], "会議 / Họp"),
    ("gcHD", "プリント外注 Hạnh Đan（加工2〜3日＋輸送）\nIn gia công ngoài Hạnh Đan (2–3 ngày + vận chuyển)", "外注加工 / Gia công ngoài", [4, 4, 4], "会議：ライン投入4日前に裁断完了 / Họp: cắt xong trước vào chuyền 4 ngày"),
    ("gcVP", "プリント外注 Vĩnh Phát（加工4〜5日＋輸送）\nIn gia công ngoài Vĩnh Phát (4–5 ngày + vận chuyển)", "外注加工 / Gia công ngoài", [6, 6, 6], "会議 / Họp. 輸送：夕方発送→翌日昼着、帰りも同様 / Chiều gửi, trưa hôm sau tới; chiều về tương tự"),
    ("theuNM", "刺繍（社内）\nThêu trong nhà máy", "裁断課 / Phòng cắt", [1, 1, 1], "社内で実施、繁忙期のみ外注 / Làm trong nhà máy, chỉ gửi ngoài khi cao điểm"),
    ("hop", "技術資料読込・展開用型紙・展開会議\nĐọc tài liệu, rập triển khai, họp triển khai", "縫製ライン / Chuyền", [2, 2, 2], "ライン投入前に完了 / Xong trước khi vào chuyền"),
    ("tkMoi", "START後のライン展開（新規）\nTriển khai sau START (mã mới)", "縫製ライン / Chuyền", [5, 7, 8], "技術者1人 1,000〜1,200秒/日 / 1 kỹ thuật ≈ 1.000–1.200 giây/ngày"),
    ("tkLap", "START後のライン展開（リピート）\nTriển khai sau START (lặp lại)", "縫製ライン / Chuyền", [3, 5, 7], "Excel 20260722"),
]
ROW = {}
for i, (key, name, dept, vals, src) in enumerate(STD_ROWS):
    r = 5 + i
    ROW[key] = r
    put(std, f"A{r}", key, f(9, color=MUTED), None, Alignment(vertical="top"))
    put(std, f"B{r}", name, f(), None, WRAP)
    put(std, f"C{r}", dept, f(9, color=MUTED), None, WRAP)
    for j, v in enumerate(vals):
        put(std, f"{'DEF'[j]}{r}", v, f(10, True, "0000FF"), INPUT_FILL, CENTER, fmt="0")
    put(std, f"G{r}", src, f(9, color=MUTED), None, WRAP)
last_std = 4 + len(STD_ROWS)

# Thiết lập chung
SET = last_std + 3
std[f"A{SET - 1}"] = "設定 / Thiết lập"
std[f"A{SET - 1}"].font = f(12, True, ACCENT)
SETTINGS = [
    ("old", "旧基準：工場着はSTARTの何日前\nChuẩn cũ: 工場着 trước START (ngày)", 14, "現行基準 / Chuẩn đang dùng"),
    ("new", "新基準：工場着はSTARTの何日前\nChuẩn mới: 工場着 trước START (ngày)", 21, "提案 / Đề xuất"),
    ("sun", "日曜休み（1＝あり、0＝なし）\nNghỉ Chủ nhật (1 = có, 0 = không)", 1, "暦日換算に使用 / Dùng để đổi sang ngày lịch"),
    ("theu", "刺繍（社内）あり（1＝あり、0＝なし）\nCó thêu trong nhà máy (1 = có, 0 = không)", 0, "刺繍は展開準備と並行のため日数は変わらない / Thêu song song với chuẩn bị nên không đổi số ngày"),
]
SROW = {}
for i, (key, name, val, note) in enumerate(SETTINGS):
    r = SET + i
    SROW[key] = r
    put(std, f"A{r}", key, f(9, color=MUTED))
    std.merge_cells(f"B{r}:C{r}")
    put(std, f"B{r}", name, f(), None, WRAP)
    put(std, f"D{r}", val, f(10, True, "0000FF"), INPUT_FILL, CENTER, fmt="0")
    std.merge_cells(f"E{r}:G{r}")
    put(std, f"E{r}", note, f(9, color=MUTED), None, WRAP)
for col, w in zip("ABCDEFG", [11, 46, 22, 13, 15, 13, 52]):
    std.column_dimensions[col].width = w
std.freeze_panes = "B5"


def S(key):
    """Ô thiết lập trên sheet 標準日数 (tham chiếu tuyệt đối)."""
    return f"'標準日数'!$D${SROW[key]}"


def STDV(key, band_cell):
    r = ROW[key]
    return f"INDEX('標準日数'!$D${r}:$F${r},{band_cell})"


# ---------------------------------------------------------------- 計算
calc = wb.create_sheet("計算")
title(calc, "計算（稼働日ベース、工場着日＝1日目）", "Tính toán (theo ngày làm việc, ngày hàng về nhà máy = ngày 1). Không cần sửa sheet này.")
CALC_HEAD = [
    ("キー\nKhóa", 22), ("標準作業時間\nSAH", 15), ("区分\nMức", 6), ("プリント\nIn", 11), ("外注日数\nNgày gia công", 10),
    ("開梱 始\nKhui bắt đầu", 8), ("開梱 終\nKhui xong", 8), ("品質検査 終\nKiểm CL xong", 9), ("生地テスト 終\nTest xong", 9),
    ("放反 終\nXả xong", 8), ("型紙 始\nRập bắt đầu", 8), ("型紙 終\nRập xong", 8), ("マーカー 始\nSơ đồ bắt đầu", 9),
    ("マーカー 終\nSơ đồ xong", 9), ("裁断 始\nCắt bắt đầu", 8), ("裁断 終\nCắt xong", 8), ("改良P 終\nRập cải tiến xong", 9),
    ("外注 終\nIn xong", 8), ("刺繍 終\nThêu xong", 8), ("展開準備 終\nChuẩn bị xong", 9), ("START\n(日目 / ngày thứ)", 10),
    ("稼働日数\nNgày làm việc", 10), ("暦日数（最大）\nNgày lịch (tối đa)", 11), ("旧基準 余裕\nDư so với cũ", 10),
    ("新基準 余裕\nDư so với mới", 10),
]
for i, (h, w) in enumerate(CALC_HEAD):
    col = get_column_letter(i + 1)
    put(calc, f"{col}4", h, f(9, True), HEAD, CENTER)
    calc.column_dimensions[col].width = w
calc.row_dimensions[4].height = 42
BANDS = ["< 8,000秒", "8,000–12,000秒", "> 12,000秒"]
PRINTS = ["なし / Không", "Hạnh Đan", "Vĩnh Phát"]
CROW = {}
r = 5
for b, band in enumerate(BANDS):
    for p in PRINTS:
        CROW[(b, p)] = r
        c = lambda k: STDV(k, f"$C{r}")
        formulas = {
            "A": f'=B{r}&"|"&D{r}',
            "B": band,
            "C": b + 1,
            "D": p,
            "E": f'=IF(D{r}="Hạnh Đan",{c("gcHD")},IF(D{r}="Vĩnh Phát",{c("gcVP")},0))',
            "F": 1,
            "G": f"=F{r}+{c('khui')}-1",
            "H": f"=G{r}+{c('qc')}",
            "I": f"=G{r}+{c('test')}",
            "J": f"=H{r}+{c('xa')}",
            "K": f"=I{r}+1",
            "L": f"=K{r}+{c('rap')}-1",
            "M": f"=K{r}+1",
            "N": f"=M{r}+{c('sodo')}-1",
            "O": f"=MAX(J{r}+1,N{r}+1,F{r}+{c('nlTruocCat')})",
            "P": f"=O{r}+{c('cat')}-1",
            "Q": f"=O{r}+{c('caitien')}-1",
            "R": f"=P{r}+E{r}",
            "S": f"=IF({S('theu')}=1,P{r}+{c('theuNM')},P{r})",
            "T": f"=P{r}+{c('hop')}",
            "U": f"=MAX(R{r},S{r},T{r})+1",
            "V": f"=U{r}-1",
            "W": f"=IF({S('sun')}=1,V{r}+INT(V{r}/6)+1,V{r})",
            "X": f"={S('old')}-W{r}",
            "Y": f"={S('new')}-W{r}",
        }
        for col, v in formulas.items():
            is_input = col in ("B", "C", "D", "F")
            put(calc, f"{col}{r}", v, f(9, color="0000FF" if is_input else INK), None,
                CENTER if col not in ("A", "B", "D") else Alignment(vertical="center"), fmt="0" if col not in "ABD" else None)
        r += 1
CALC_LAST = r - 1
note_r = CALC_LAST + 2
NOTES_CALC = [
    "・暦日数（最大）＝稼働日数＋日曜日の数。工場着日が日曜でも開梱は翌稼働日となるため、最悪の曜日で 稼働日数＋INT(稼働日数/6)＋1。",
    "  Ngày lịch (tối đa) = ngày làm việc + số Chủ nhật. Nếu hàng về nhà máy đúng Chủ nhật thì hôm sau mới khui, nên trường hợp xấu nhất là ngày làm việc + INT(ngày làm việc/6) + 1.",
    "・マーカー作成は型紙調整の2日目から開始。裁断は放反・マーカー完了後、かつ開梱から「原材料準備完了」日数以降。",
    "  Đi sơ đồ bắt đầu từ ngày thứ 2 của chỉnh rập. Cắt bắt đầu khi xả vải và sơ đồ xong, và không sớm hơn số ngày “nguyên liệu sẵn sàng trước khi cắt” tính từ ngày khui.",
    "・START（ライン投入）＝外注プリント・刺繍・展開準備がすべて完了した翌稼働日。",
    "  START (vào chuyền) = ngày làm việc kế tiếp sau khi in gia công ngoài, thêu và chuẩn bị triển khai đều xong.",
    "・同じ規則で作成したWebツール（workflow/schedule.js）の計算結果と一致することを確認済み。",
    "  Đã đối chiếu khớp với công cụ web cùng quy tắc (workflow/schedule.js).",
]
for i, t in enumerate(NOTES_CALC):
    calc[f"A{note_r + i}"] = t
    calc[f"A{note_r + i}"].font = f(9, color=MUTED)
calc.freeze_panes = "E5"

CALC_W = f"'計算'!$W$5:$W${CALC_LAST}"


def CREF(col, b, p):
    return f"'計算'!${col}${CROW[(b, p)]}"


# ---------------------------------------------------------------- 概要
sm = wb.create_sheet("概要", 0)
title(sm, "生産工程リードタイム ― 原副資材納期基準の見直し（TIJ協議資料）",
      "Lead time sản xuất — đề nghị đổi mốc giao nguyên phụ liệu (tài liệu thảo luận với TIJ)")
sm["A3"] = "START＝ライン投入（準備完了）。ライン展開（先頭5枚まで）はSTART後。 / START = ngày vào chuyền (chuẩn bị xong). Triển khai chuyền đến 5 pcs đầu diễn ra sau START."
sm["A3"].font = f(10, color=MUTED)

put(sm, "A5", "基準 / Chuẩn", f(10, True), HEAD, CENTER)
put(sm, "B5", "ETA（港着）\nVề cảng HCM", f(10, True), HEAD, CENTER)
put(sm, "C5", "工場着\nVề nhà máy", f(10, True), HEAD, CENTER)
put(sm, "D5", "START", f(10, True), HEAD, CENTER)
xnk = f"INDEX('標準日数'!$D${ROW['xnk']}:$F${ROW['xnk']},1)"
for r_, (lab, key) in zip((6, 7), (("旧基準 / Chuẩn cũ", "old"), ("新基準 / Chuẩn mới（提案 / đề xuất）", "new"))):
    put(sm, f"A{r_}", lab, f(10, True), None, Alignment(vertical="center"))
    put(sm, f"B{r_}", f'="START −"&({S(key)}+{xnk})&"日 / ngày"', f(11, True, "008000"), None, CENTER)
    put(sm, f"C{r_}", f'="START −"&{S(key)}&"日 / ngày"', f(11, True, "008000"), None, CENTER)
    put(sm, f"D{r_}", "D", f(11, True), None, CENTER)
sm["A7"].fill = fill(ACCENT_SOFT)

sm["A9"] = "ケース別 所要日数（工場着→START、暦日・最大）/ Số ngày cần từ 工場着 đến START (ngày lịch, tối đa)"
sm["A9"].font = f(12, True, ACCENT)
put(sm, "A10", "標準作業時間\nSAH", f(10, True), HEAD, CENTER)
sm.merge_cells("A10:A11")
groups = [("B", "プリントなし / Không in"), ("E", "プリント Hạnh Đan / In tại Hạnh Đan"), ("H", "プリント Vĩnh Phát / In tại Vĩnh Phát")]
for start_col, lab in groups:
    ci = ord(start_col) - 64
    end_col = get_column_letter(ci + 2)
    sm.merge_cells(f"{start_col}10:{end_col}10")
    put(sm, f"{start_col}10", lab, f(10, True), HEAD, CENTER)
    for k, sub in enumerate(["所要日数\nCần", "旧基準 余裕\nDư (cũ)", "新基準 余裕\nDư (mới)"]):
        put(sm, f"{get_column_letter(ci + k)}11", sub, f(9, True), HEAD, CENTER)
sm.row_dimensions[11].height = 30
for b, band in enumerate(BANDS):
    r_ = 12 + b
    put(sm, f"A{r_}", band, f(10, True), None, CENTER)
    for g, p in enumerate(PRINTS):
        ci = 2 + g * 3
        put(sm, f"{get_column_letter(ci)}{r_}", f"={CREF('W', b, p)}", f(11, True, "008000"), None, CENTER, fmt='0"日"')
        put(sm, f"{get_column_letter(ci + 1)}{r_}", f"={CREF('X', b, p)}", f(10, False, "008000"), None, CENTER, fmt='+0;−0;0')
        put(sm, f"{get_column_letter(ci + 2)}{r_}", f"={CREF('Y', b, p)}", f(10, False, "008000"), None, CENTER, fmt='+0;−0;0')
for col in ("C", "D", "F", "G", "I", "J"):
    rng = f"{col}12:{col}14"
    sm.conditional_formatting.add(rng, FormulaRule(formula=[f"{col}12<0"], fill=fill(BAD), font=Font(name=FONT, color=BAD_INK, bold=True)))
    sm.conditional_formatting.add(rng, FormulaRule(formula=[f"{col}12>=0"], fill=fill(GOOD), font=Font(name=FONT, color=GOOD_INK)))
sm["A15"] = "余裕＝基準日数−所要日数（マイナスはSTARTに間に合わない）/ Dư = số ngày theo chuẩn − số ngày cần (âm là không kịp START). 刺繍は社内で展開準備と並行のため日数に影響しない / Thêu làm trong nhà máy, song song với chuẩn bị nên không ảnh hưởng số ngày."
sm["A15"].font = f(9, color=MUTED)

sm["A17"] = "要点 / Điểm chính"
sm["A17"].font = f(12, True, ACCENT)
KPI = [
    ("所要日数の範囲（1品番のみ準備）\nSố ngày cần (chỉ chuẩn bị 1 mã)", f'=MIN({CALC_W})&"〜"&MAX({CALC_W})&"日 / ngày"'),
    ("旧基準で間に合うケース\nSố trường hợp kịp với chuẩn cũ", f'=COUNTIF({CALC_W},"<="&{S("old")})&" / 9"'),
    ("新基準で間に合うケース\nSố trường hợp kịp với chuẩn mới", f'=COUNTIF({CALC_W},"<="&{S("new")})&" / 9"'),
    ("旧基準の最大不足\nThiếu nhiều nhất với chuẩn cũ", f'=MAX(0,MAX({CALC_W})-{S("old")})&"日 / ngày"'),
    ("新基準の最少予備日\nDự phòng ít nhất với chuẩn mới", f'={S("new")}-MAX({CALC_W})&"日 / ngày"'),
    ("10〜12日展開が可能なケース\nSố trường hợp làm được trong 10–12 ngày", f'=COUNTIF({CALC_W},"<=12")&" / 9"'),
]
for i, (lab, fx) in enumerate(KPI):
    r_ = 18 + i
    sm.merge_cells(f"A{r_}:D{r_}")
    put(sm, f"A{r_}", lab, f(10), None, WRAP)
    sm.merge_cells(f"E{r_}:G{r_}")
    put(sm, f"E{r_}", fx, f(12, True, "008000"), None, CENTER)
    sm.row_dimensions[r_].height = 30

sm["A25"] = "結論 / Kết luận"
sm["A25"].font = f(12, True, ACCENT)
concl_ja = (f'="工場着からライン投入まで、1品番のみを準備する場合でも"&MIN({CALC_W})&"〜"&MAX({CALC_W})&"日必要。'
            f'旧基準（工場着"&{S("old")}&"日前）で間に合うのは9ケース中"&COUNTIF({CALC_W},"<="&{S("old")})&"ケースで、'
            f'Vĩnh Phát外注プリント品は最大"&MAX(0,MAX({CALC_W})-{S("old")})&"日不足。'
            f'新基準（工場着"&{S("new")}&"日前）なら全ケースで"&({S("new")}-MAX({CALC_W}))&"日以上の予備日を確保でき、ETA遅延・不良・同時進行品番による待ちを吸収できる。"')
concl_vi = (f'="Từ 工場着 đến vào chuyền, dù chỉ chuẩn bị 1 mã, vẫn cần "&MIN({CALC_W})&"–"&MAX({CALC_W})&" ngày. '
            f'Chuẩn cũ (工場着 trước "&{S("old")}&" ngày) chỉ đáp ứng "&COUNTIF({CALC_W},"<="&{S("old")})&"/9 trường hợp; mã in tại Vĩnh Phát thiếu tới "&MAX(0,MAX({CALC_W})-{S("old")})&" ngày. '
            f'Chuẩn mới (工場着 trước "&{S("new")}&" ngày) cho mọi trường hợp dự phòng ít nhất "&({S("new")}-MAX({CALC_W}))&" ngày để hấp thụ ETA trễ, vải lỗi và thời gian chờ khi nhiều mã chạy song song."')
for r_, fx, bold in ((26, concl_ja, True), (27, concl_vi, False)):
    sm.merge_cells(f"A{r_}:J{r_}")
    put(sm, f"A{r_}", fx, f(10, bold), ACCENT_SOFT, WRAP)
    sm.row_dimensions[r_].height = 48

sm["A29"] = "10〜12日での展開要求：プリントなし・Hạnh Đanの一部ケースのみ物理的に可能。予備日はほぼゼロで、ETA遅れや不良が1つでもあれば間に合わない。Vĩnh Phát外注品は不可。"
sm["A30"] = "Yêu cầu triển khai trong 10–12 ngày: chỉ làm được ở một số trường hợp không in hoặc in tại Hạnh Đan, gần như không có dự phòng; chỉ cần ETA trễ hoặc vải lỗi là không kịp. Mã in tại Vĩnh Phát thì không thể."
sm["A29"].font = f(10, True)
sm["A30"].font = f(10, color=MUTED)
sm["A32"] = "根拠：「Rà soát lại thời gian cần để vào chuyền – 20260722」およびリードタイム再検討会議。標準日数はシート「標準日数」で変更可能。 / Căn cứ: file rà soát 20260722 và cuộc họp rà soát lead time. Sửa số ngày chuẩn ở sheet 標準日数."
sm["A32"].font = f(9, color=MUTED, italic=True)
sm.column_dimensions["A"].width = 30
for col in "BCDEFGHIJ":
    sm.column_dimensions[col].width = 12
# Dòng chữ dài: gộp ô A:J và cho xuống dòng để không tràn khi in
for r_, h in ((3, 30), (15, 30), (29, 30), (30, 30), (32, 30)):
    sm.merge_cells(f"A{r_}:J{r_}")
    sm[f"A{r_}"].alignment = WRAP
    sm.row_dimensions[r_].height = h
for r_ in (5, 6, 7):
    sm.row_dimensions[r_].height = 32
    sm[f"A{r_}"].alignment = Alignment(wrap_text=True, vertical="center")
for col in "BC":
    sm.column_dimensions[col].width = 16

# ---------------------------------------------------------------- フロー (Gantt)
fl = wb.create_sheet("フロー", 1)
title(fl, "業務フロー（工場着→START、稼働日ベース）", "Flowchart các bước từ 工場着 đến START (theo ngày làm việc). Chọn điều kiện ở ô nền vàng.")
put(fl, "A4", "標準作業時間 / SAH", f(10, True), HEAD)
put(fl, "A5", "プリント / In", f(10, True), HEAD)
put(fl, "A6", "品番区分 / Loại mã", f(10, True), HEAD)
put(fl, "B4", BANDS[2], f(10, True, "0000FF"), INPUT_FILL, CENTER)
put(fl, "B5", PRINTS[2], f(10, True, "0000FF"), INPUT_FILL, CENTER)
put(fl, "B6", "新規 / Mới", f(10, True, "0000FF"), INPUT_FILL, CENTER)
fl.merge_cells("B4:C4"); fl.merge_cells("B5:C5"); fl.merge_cells("B6:C6")
dv1 = DataValidation(type="list", formula1='"' + ",".join(BANDS) + '"', allow_blank=False)
dv2 = DataValidation(type="list", formula1='"' + ",".join(PRINTS) + '"', allow_blank=False)
dv3 = DataValidation(type="list", formula1='"新規 / Mới,リピート / Lặp lại"', allow_blank=False)
for dv, ref in ((dv1, "B4"), (dv2, "B5"), (dv3, "B6")):
    fl.add_data_validation(dv)
    dv.add(ref)
put(fl, "E4", "計算行 / Dòng tính", f(9, color=MUTED), None)
put(fl, "F4", "=MATCH(B4&\"|\"&B5,'計算'!$A$5:$A$%d,0)" % CALC_LAST, f(9, color=MUTED), None, CENTER)
put(fl, "E5", "区分 / Mức", f(9, color=MUTED), None)
put(fl, "F5", "=MATCH(B4,{\"%s\",\"%s\",\"%s\"},0)" % tuple(BANDS), f(9, color=MUTED), None, CENTER)
put(fl, "E6", "START（日目）", f(9, color=MUTED), None)


def CL(col):
    return f"INDEX('計算'!${col}$5:${col}${CALC_LAST},$F$4)"


put(fl, "F6", f"={CL('U')}", f(10, True), None, CENTER)
put(fl, "H4", "工場着→START", f(9, color=MUTED), None)
put(fl, "I4", f'={CL("V")}&" 稼働日 / ngày làm việc"', f(10, True, ACCENT), None)
fl.merge_cells("I4:N4")
put(fl, "H5", "暦日（最大）", f(9, color=MUTED), None)
put(fl, "I5", f'={CL("W")}&" 日 / ngày lịch"', f(10, True, ACCENT), None)
fl.merge_cells("I5:N5")

NDAYS = 30
FIRST_DAY_COL = 6  # F
HR1, HR2 = 8, 9
put(fl, f"A{HR2}", "部署 / Bộ phận", f(9, True), HEAD, CENTER)
put(fl, f"B{HR2}", "工程 / Bước", f(9, True), HEAD, CENTER)
put(fl, f"C{HR2}", "日数\nNgày", f(9, True), HEAD, CENTER)
put(fl, f"D{HR2}", "開始\nBắt đầu", f(9, True), HEAD, CENTER)
put(fl, f"E{HR2}", "終了\nKết thúc", f(9, True), HEAD, CENTER)
put(fl, f"A{HR1}", "稼働日（工場着＝1）/ Ngày làm việc (ngày hàng về = 1)", f(9, color=MUTED), None, None, border=False)
for d in range(1, NDAYS + 1):
    col = get_column_letter(FIRST_DAY_COL + d - 1)
    put(fl, f"{col}{HR2}", d, f(8, True), HEAD, CENTER, fmt="0")
    put(fl, f"{col}{HR1}", f'=IF({col}${HR2}=$F$6,"START","D"&IF({col}${HR2}<$F$6,"−","+")&ABS({col}${HR2}-$F$6))',
        f(7, color=MUTED), None, CENTER)
    fl.column_dimensions[col].width = 4.6
fl.row_dimensions[HR1].height = 18

tk = f"INDEX(IF($B$6=\"新規 / Mới\",'標準日数'!$D${ROW['tkMoi']}:$F${ROW['tkMoi']},'標準日数'!$D${ROW['tkLap']}:$F${ROW['tkLap']}),$F$5)"
GANTT = [
    ("bocvac", "荷役\nBốc vác", "荷降ろし・開梱検品\nXuống hàng, khui thùng", f"={CL('F')}", f"={CL('G')}"),
    ("qckho", "倉庫品質管理\nChất lượng kho", "品質検査・有効幅確認\nKiểm chất lượng, báo khổ", f"={CL('G')}+1", f"={CL('H')}"),
    ("kythuat", "技術課\nKỹ thuật", "生地テスト\nTest vải", f"={CL('G')}+1", f"={CL('I')}"),
    ("cat", "裁断課\nPhòng cắt", "放反\nXả vải", f"={CL('H')}+1", f"={CL('J')}"),
    ("kythuat", "技術課\nKỹ thuật", "型紙調整\nĐiều chỉnh rập", f"={CL('K')}", f"={CL('L')}"),
    ("kythuat", "技術課\nKỹ thuật", "マーカー作成・型紙出力\nĐi sơ đồ + in rập", f"={CL('M')}", f"={CL('N')}"),
    ("cat", "裁断課\nPhòng cắt", "裁断・ナンバリング・束ね\nCắt, đánh số, bó", f"={CL('O')}", f"={CL('P')}"),
    ("kythuat", "技術課\nKỹ thuật", "改良パターン（並行）\nRập cải tiến (song song)", f"={CL('O')}", f"={CL('Q')}"),
    ("giacong", "外注加工\nGia công ngoài", "プリント（外注・輸送含む）\nIn gia công ngoài (gồm vận chuyển)", f"={CL('P')}+1", f"={CL('R')}"),
    ("cat", "裁断課\nPhòng cắt", "刺繍（社内）\nThêu trong nhà máy", f"={CL('P')}+1", f"={CL('S')}"),
    ("chuyen", "縫製ライン\nChuyền", "技術資料読込・展開準備\nĐọc tài liệu, chuẩn bị triển khai", f"={CL('P')}+1", f"={CL('T')}"),
    ("start", "", "START＝ライン投入（準備完了）\nSTART = vào chuyền", "=$F$6", "=$F$6"),
    ("chuyen", "縫製ライン\nChuyền", "ライン展開（先頭5枚まで）\nTriển khai chuyền đến 5 pcs đầu", "=$F$6", f"=$F$6+{tk}-1"),
]
G0 = HR2 + 1
last_col = get_column_letter(FIRST_DAY_COL + NDAYS - 1)
for i, (dept, dlab, name, s_, e_) in enumerate(GANTT):
    r_ = G0 + i
    put(fl, f"A{r_}", dlab, f(8, color=MUTED), None, WRAP)
    put(fl, f"B{r_}", name, f(9, dept == "start"), ACCENT_SOFT if dept == "start" else None, WRAP)
    put(fl, f"D{r_}", s_, f(9), None, CENTER, fmt="0")
    put(fl, f"E{r_}", e_, f(9), None, CENTER, fmt="0")
    put(fl, f"C{r_}", f"=IF(E{r_}>=D{r_},E{r_}-D{r_}+1,0)", f(9, True), None, CENTER, fmt='0;-0;"—"')
    fl.row_dimensions[r_].height = 30
    for d in range(1, NDAYS + 1):
        put(fl, f"{get_column_letter(FIRST_DAY_COL + d - 1)}{r_}", None, f(8), None, CENTER)
    rng = f"F{r_}:{last_col}{r_}"
    color = ACCENT if dept == "start" else DEPT[dept]
    fl.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(F${HR2}>=$D{r_},F${HR2}<=$E{r_},$C{r_}>0)"], fill=fill(color)))
G_LAST = G0 + len(GANTT) - 1
fl.conditional_formatting.add(f"F{HR1}:{last_col}{G_LAST}", FormulaRule(formula=[f"F${HR2}=$F$6"], fill=fill(ACCENT_SOFT)))
fl.conditional_formatting.add(f"F{HR1}:{last_col}{HR1}", FormulaRule(formula=[f"F${HR2}=$F$6"], font=Font(name=FONT, bold=True, color=ACCENT)))
n_r = G_LAST + 2
FL_NOTES = [
    "・D−n＝STARTまでの稼働日数（日曜・祝日を含まない）。暦日の目安はシート「概要」を参照。",
    "  D−n = số ngày làm việc trước START (không tính Chủ nhật, ngày lễ). Số ngày lịch xem ở sheet 概要.",
    "・刺繍は「標準日数」シートの設定が1のときのみ表示。プリントなしを選ぶと外注行は空欄。",
    "  Dòng thêu chỉ hiện khi thiết lập ở sheet 標準日数 = 1. Chọn “Không in” thì dòng gia công ngoài để trống.",
]
for i, t in enumerate(FL_NOTES):
    fl[f"A{n_r + i}"] = t
    fl[f"A{n_r + i}"].font = f(9, color=MUTED)
fl.column_dimensions["A"].width = 22
fl.column_dimensions["B"].width = 34
for col in "CDE":
    fl.column_dimensions[col].width = 7.5
fl.freeze_panes = f"F{G0}"

# ---------------------------------------------------------------- 理由・依頼
tx = wb.create_sheet("理由・依頼", 2)
title(tx, "基準変更の理由・TIJへの依頼事項（案）", "Lý do thay đổi chuẩn, đề nghị đối với TIJ (bản nháp để thảo luận)")


def block(start_row, heading, items, numbered=True):
    tx[f"A{start_row}"] = heading
    tx[f"A{start_row}"].font = f(12, True, ACCENT)
    put(tx, f"A{start_row + 1}", "No.", f(9, True), HEAD, CENTER)
    put(tx, f"B{start_row + 1}", "日本語", f(9, True), HEAD, CENTER)
    put(tx, f"C{start_row + 1}", "Tiếng Việt", f(9, True), HEAD, CENTER)
    for i, (ja, vi) in enumerate(items):
        r_ = start_row + 2 + i
        put(tx, f"A{r_}", i + 1 if numbered else "", f(10, True, ACCENT), None, Alignment(horizontal="center", vertical="top"))
        put(tx, f"B{r_}", ja, f(10), None, WRAP)
        put(tx, f"C{r_}", vi, f(10), None, WRAP)
        tx.row_dimensions[r_].height = max(45, 15 * (1 + max(len(ja) // 38, len(vi) // 60)))
    return start_row + 2 + len(items) + 1


row_ = 4
row_ = block(row_, "基準変更の理由 / Lý do thay đổi chuẩn", [
    ("ライン数の増加：10年前よりライン数が増え、複数品番の準備が同時に進む。技術課・裁断課の作業が重なり、品番ごとに待ち時間が発生しやすい。",
     "Số chuyền tăng: so với 10 năm trước có nhiều chuyền hơn, nhiều mã được chuẩn bị cùng lúc. Việc của kỹ thuật và phòng cắt chồng lên nhau, mỗi mã dễ phát sinh thời gian chờ."),
    ("難易度の上昇：12,000秒を超える品番は型紙調整＋マーカーで最低3日。ライン展開は技術者1人1日1,000〜1,200秒が限界で、12,000〜15,000秒では7〜8日かかる。",
     "Độ khó tăng: mã trên 12.000 giây cần tối thiểu 3 ngày cho rập + sơ đồ. Mỗi kỹ thuật chỉ triển khai được 1.000–1.200 giây/ngày, mã 12.000–15.000 giây mất 7–8 ngày."),
    ("小ロット化・外注加工：同じ生産量でも品番数が増え、品番ごとにテスト・型紙・マーカー・展開準備が必要。プリントの外注（Vĩnh Phát 4〜5日＋輸送）では裁断完了からライン投入まで6日かかる。",
     "Lô nhỏ, gia công ngoài: cùng sản lượng nhưng nhiều mã hơn, mã nào cũng cần test, rập, sơ đồ, chuẩn bị triển khai. In gia công ngoài (Vĩnh Phát 4–5 ngày + vận chuyển) làm khoảng cách từ cắt xong đến vào chuyền lên 6 ngày."),
    ("品番切替の増加：切替のたびに技術資料の読込・展開用型紙・展開会議が必要。準備不足のままライン投入すると現場での研究時間が延び、効率が下がる。",
     "Chuyển đổi mã nhiều: mỗi lần chuyển đổi đều phải đọc tài liệu, làm rập triển khai, họp triển khai. Vào chuyền khi chưa chuẩn bị đủ thì thời gian tìm hiểu tại chuyền kéo dài, hiệu suất giảm."),
])
row_ = block(row_, "TIJへの依頼事項（案）/ Đề nghị đối với TIJ (bản nháp)", [
    ("原副資材（生地・副資材すべて）のETA（港着）をSTART日の28日前、工場着を21日前までとする。（旧基準：ETA 21日前・工場着 14日前）",
     "Toàn bộ nguyên phụ liệu (vải và phụ liệu): ETA về cảng chậm nhất D−28, về nhà máy chậm nhất D−21. (Chuẩn cũ: ETA D−21, 工場着 D−14.)"),
    ("分納の場合は最終便も同じ期限とする。1点でも欠品があると裁断を開始できない。",
     "Nếu giao nhiều đợt, đợt cuối cũng phải đúng mốc này. Thiếu một loại phụ liệu là không thể bắt đầu cắt."),
    ("ETAの変更・遅延は判明次第すぐに共有。出荷前にパッキングリスト・ロット・反数・生地幅を共有する。",
     "Báo ngay khi ETA thay đổi hoặc trễ. Gửi trước packing list, số lot, số cây vải, khổ vải trước khi xuất hàng."),
    ("10〜12日での展開要求は、プリントなし・低難易度の品番に限り個別に協議する。",
     "Yêu cầu triển khai trong 10–12 ngày chỉ áp dụng cho mã không in, độ khó thấp, và thỏa thuận riêng từng mã."),
    ("テト・祝日が期間内に入る場合は休日日数分を前倒しする。",
     "Nếu khoảng thời gian này có Tết hoặc ngày lễ thì đẩy ETA sớm hơn đúng bằng số ngày nghỉ."),
])
row_ = block(row_, "工場側の取り組み（案）/ Cải tiến phía nhà máy (bản nháp)", [
    ("技術資料の読込・型紙確認をライン投入の2日前までに完了する体制をつくる。",
     "Xây dựng cách làm để đọc tài liệu kỹ thuật, kiểm tra rập xong trước ngày vào chuyền 2 ngày."),
    ("倉庫・物流担当が原材料のETAを正確に管理し、計画チームと共有する。",
     "Kho và logistics quản lý chính xác ETA nguyên liệu, chia sẻ với bộ phận kế hoạch."),
])
tx.column_dimensions["A"].width = 6
tx.column_dimensions["B"].width = 62
tx.column_dimensions["C"].width = 78

# ---------------------------------------------------------------- In ấn
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size, bold=c.font.bold, color=c.font.color, italic=c.font.italic)
wb.active = 0
wb.save(OUT)
print("saved", OUT)
