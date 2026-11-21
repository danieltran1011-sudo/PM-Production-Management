"""APIC 2026年度設備投資 説明資料（ドラフト）を生成する。
出典: APIC2026年度設備投資予算案 20260403.xlsx（優先①台数・単価(USD)）と ChatGPT handoff の数値。
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

OUT = sys.argv[1]
F = "Arial"
BLUE = "0000FF"; GREEN = "008000"
YEL = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
GRY = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="A6A6A6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
USD = '#,##0;(#,##0);"-"'
USD2 = '#,##0.00;(#,##0.00);"-"'
JPY = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM3 = '#,##0.000;(#,##0.000);"-"'

wb = Workbook()

def font(c, bold=False, color="000000", size=10):
    c.font = Font(name=F, bold=bold, color=color, size=size)

def put(ws, ref, v, bold=False, color=None, fmt=None, fill=None, wrap=True, border=True, size=10, align=None):
    c = ws[ref]
    c.value = v
    if color is None:
        color = "000000"
    font(c, bold, color, size)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    c.alignment = Alignment(wrap_text=wrap, vertical="center", horizontal=align)
    if border: c.border = BOX
    return c

def inp(ws, ref, v, fmt=None, note=None, yellow=False):
    c = put(ws, ref, v, color=BLUE, fmt=fmt, fill=YEL if yellow else None)
    if note: c.comment = Comment(note, "draft")
    return c

def header(ws, row, labels, col0=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=t)
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
        c.fill = HDR; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, size=14, border=False, wrap=False)
    if sub: put(ws, "A2", sub, size=9, border=False, wrap=False, color="595959")

def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v

def legend(ws, ref):
    put(ws, ref, "凡例：青字＝入力値／黄色＝工場・関係部署で記入または確認が必要なセル／黒字＝計算式", size=9, border=False, wrap=False, color="595959")

# ---------------------------------------------------------------- パラメータ
P = wb.active; P.title = "パラメータ"
title(P, "パラメータ（前提条件）", "全シートの計算はこのシートの値を参照します。値を変更すると全体が再計算されます。")
legend(P, "A3")
header(P, 5, ["項目", "値", "単位", "出典・備考"])
params = [
    ("為替レート（計画）", 160, "円/USD", "予算案ファイル P30（=160）。O26では150を使用しており、統一要確認", False, JPY),
    ("為替レート（感応度用）", 150, "円/USD", "予算案ファイル O26", False, JPY),
    ("投資目標上限（Suntool除く）", 40000000, "円", "佐藤さん意向：4,000万円以下", False, JPY),
    ("APIC負担額", 50000, "USD", "新規決定事項。対象機材・資産計上先・償却の扱いは要確認", True, USD),
    ("償却年数", 7, "年", "予算案ファイル U列（=金額/7×為替）と同じ前提", False, "0"),
    ("年間稼働日数", 295, "日", "handoff（電力試算の前提）", False, "0"),
    ("1日稼働時間", 8, "時間", "電力試算の前提（2.400kWh/日から逆算）。要確認", True, "0.0"),
    ("電気料金単価", 0.08, "USD/kWh", "handoff。工場の実績単価で要確認", True, "0.000"),
    ("作業者稼働効率", 1.0, "%", "1人年間有効稼働秒の算出用。APIC標準値に置換要", True, PCT),
    ("工員人件費（月額・1人）", None, "USD/月", "省人効果・損益分岐の金額換算用。工場人事データ（社保等込み）で記入", True, USD),
]
for i, (k, v, u, s, y, f) in enumerate(params):
    r = 6 + i
    put(P, f"A{r}", k); inp(P, f"B{r}", v, fmt=f, yellow=y); put(P, f"C{r}", u); put(P, f"D{r}", s)
r = 6 + len(params)
put(P, f"A{r}", "1人年間有効稼働秒", bold=True)
put(P, f"B{r}", "=B12*3600*B11*B14", fmt=JPY); put(P, f"C{r}", "秒/人・年"); put(P, f"D{r}", "＝1日稼働時間×3,600×年間稼働日数×稼働効率")
extra = [
    ("月間稼働日数", 26, "日/月", "比較シートの前提（ダウン充填装置比較と同じ考え方）。APIC実績で要確認", True, "0"),
    ("設備稼働率", 0.6, "%", "比較シートの前提（同上：60%）。APIC実績で要確認", True, PCT),
    ("余裕率", 0.2, "%", "日産能力算出用（同上：20%）", True, PCT),
]
for i, (k, v, u, s_, y, f) in enumerate(extra):
    r = 18 + i
    put(P, f"A{r}", k); inp(P, f"B{r}", v, fmt=f, yellow=y); put(P, f"C{r}", u); put(P, f"D{r}", s_)
put(P, "A21", "1日有効作業秒（1台・1人）", bold=True)
put(P, "B21", "=B12*3600*(1-B20)", fmt=JPY); put(P, "C21", "秒/日"); put(P, "D21", "＝1日稼働時間×3,600×(1－余裕率)")
widths(P, {"A": 30, "B": 16, "C": 12, "D": 70})
# 参照名
RATE, RATE2, TARGET, APIC, YEARS, DAYS, HOURS, ELEC, EFF, LABOR, SECYR = (
    "'パラメータ'!$B$6", "'パラメータ'!$B$7", "'パラメータ'!$B$8", "'パラメータ'!$B$9", "'パラメータ'!$B$10",
    "'パラメータ'!$B$11", "'パラメータ'!$B$12", "'パラメータ'!$B$13", "'パラメータ'!$B$14", "'パラメータ'!$B$15", "'パラメータ'!$B$16")
MDAYS, UTIL, DAYSEC = "'パラメータ'!$B$18", "'パラメータ'!$B$19", "'パラメータ'!$B$21"

# ---------------------------------------------------------------- 総括
S = wb.create_sheet("総括", 0)
title(S, "APIC 2026年度 設備投資 ― 総括（ドラフト）",
      "現状 → 課題 → 数値根拠 → 導入効果 → 費用対効果 → 投資妥当性 → 優先順位　｜　金額はUSD建て、円換算は『パラメータ』の為替レートを使用")
put(S, "A4", "方針", bold=True, border=False)
msgs = [
    "1. 全機材を一括で正当化せず、『必須（安全・顧客要求・生産継続）』『改善（省人・工数・レンタル削減・品質）』『戦略・検証（将来案件・試験）』に区分する。",
    "2. 必要性と数字が明確な機材から順次決裁し、全機材の説明完了を待たずに投資を前進させる（第1次 → 第2次 → 試験後）。",
    "3. 推奨案は投資目標（Suntool除き4,000万円以下）に収める。戦略・検証機材のうち根拠未整備のものは次回以降に回す。",
]
for i, m in enumerate(msgs):
    put(S, f"A{5+i}", m, border=False, wrap=False)

# 明細表
H0 = 18
cols = ["当初No", "設備名", "型式／メーカー", "区分", "投資目的", "根拠（要点）", "当初案台数", "単価(USD)",
        "当初案金額(USD)", "推奨案台数", "推奨案金額(USD)", "必須案金額(USD)", "全体案金額(USD)",
        "推奨案金額(円)", "推奨案 年間償却(円)", "決裁ステップ", "数値根拠の状況", "詳細シート"]
put(S, f"A{H0-1}", "設備別明細（優先①の15機種）", bold=True, border=False, wrap=False)
header(S, H0, cols)
items = [
    (1, "2本針オーバーロック（サーボ・自動糸切り）", "MO-6814／Juki", "必須", "生産継続・レンタル削減",
     "実稼働可能86台 vs ピーク必要112台。第2工場借用・外部レンタルで不足が既に顕在化", 5, 980, 5, "第1次", "試算済（台数データ要最終確認）", "ロックミシン／比較_レンタル購入"),
    (2, "2本針インターロック（サーボ・自動糸切り）", "MO-6816／Juki", "必須", "生産継続",
     "実稼働可能64台 vs ピーク必要70台", 2, 1150, 2, "第1次", "試算済（台数データ要最終確認）", "ロックミシン"),
    (11, "スポットプレス 30×10cm", "PTTG", "必須", "安全（老朽機入替）",
     "旧型（約15年）は安全装置なし・後付け不可・メーカー廃棄推奨。必要総数の計算根拠を整備中", 10, 764, None, "第1次", "データ待ち（必要総数・旧型台数）", "スポットプレス"),
    (12, "温冷カーブプレス（フロントジッパー）", "HTM-8749／Nawon", "必須", "顧客要求（26FW GW）",
     "26FW GW要請。現行シーム機のみでは品質・品位を確保できない品番への対応", 2, 19980, 2, "第1次", "顧客要請（定性）", "―"),
    (4, "パターンシーマ 60×40cm", "HMS-6040／Hikari", "改善", "省人・品質・属人性排除",
     "ポケットファスナー等の工程集約。手作業／治具／PSの3方式比較を整備中", 3, 4450, 3, "第2次", "作成中（工程秒数・拘束時間）", "比較_パターンシーマ"),
    (5, "パターンシーマ 130×95cm（360度回転ヘッド）", "HMS-13095A/RND／Hikari", "改善", "省人・品質・外注/レンタル削減",
     "GW系フロントファスナー指示、キルティング内製化（25FW外注 約3,000USD）、2025年レンタル約3,017USD", 6, 6500, 6, "第2次", "作成中（工程秒数・レンタル）", "比較_パターンシーマ／比較_レンタル購入"),
    (6, "本縫針送りミシン", "DM1968L-NF／DUMA", "改善", "機材不足・レンタル削減",
     "裾縫い・ファスナー付けで各ライン1台。理由書のレンタル削減額をレンタル実績と照合中", 10, 880, 10, "第2次", "要照合（レンタル実績）", "比較_レンタル購入"),
    (7, "テーピングマシン spec-D", "AI-002／H&H", "改善", "省人・品質",
     "長尺パーツのシーム作業。熟練者・2名体制 → 属人性排除。アタッチメント対応不可を確認済", 3, 9700, 3, "第2次", "作成中（省人効果）", "比較_シーム機"),
    (9, "温冷プレス 60×12cm", "CSL-201／H&H", "改善", "機材不足・工数削減",
     "現有30×10cm 26台で不足。プレス面積拡大で工程秒数約1/2の見込み", 6, 4150, 6, "第2次", "要実測（工程秒数）", "比較_温冷プレス"),
    (10, "平プレス 125×83cm", "CS-691／H&H", "改善", "品質・新構造品対応",
     "TNFエアチャンバー等の大型パーツ。プレス精度の検証方法を提示", 2, 13000, 2, "第2次", "検証方法 案作成済", "平プレス検証"),
    (8, "テーピングマシン spec-IJ", "AI-129／H&H", "戦略・検証", "品質（曲線部）",
     "フード等曲線部のシーム精度。1台で評価後、追加判断", 1, 6800, 1, "試験後", "試験導入", "―"),
    (15, "温冷プレス", "HTM-8601／Nawon", "戦略・検証", "品質・営業要望",
     "2025年度1台導入済で評価良好。TNF等高難度品で営業から増設要望", 1, 18333, 1, "第2次", "導入実績あり（定性）", "―"),
    (18, "超音波溶着機 35kHz", "US-512／H&H", "戦略・検証", "品質・提案幅拡大",
     "既存28kHz機で一部素材の溶着不良。サンプル作製を兼ねたトライアル", 1, 6300, 1, "試験後", "試験導入", "―"),
    (13, "温冷カーブプレス（袖）", "HTM-8739／Nawon", "戦略・検証", "品質",
     "袖筒専用テーブル。根拠整備後に次回以降で判断", 1, 19980, 0, "次回以降", "未整備", "―"),
    (14, "温冷プレス（ヘムライン）", "HTM-8709／Nawon", "戦略・検証", "品質",
     "ストレート形状のため同等H&H機（CSL-207 約10,500USD）との比較要。次回以降で判断", 1, 19980, 0, "次回以降", "未整備（代替機比較要）", "―"),
]
first = H0 + 1
for i, it in enumerate(items):
    r = first + i
    no, name, model, cat, purpose, basis, q0, unit, qr, step, status, sheet = it
    put(S, f"A{r}", no, align="center"); put(S, f"B{r}", name); put(S, f"C{r}", model)
    put(S, f"D{r}", cat, align="center"); put(S, f"E{r}", purpose); put(S, f"F{r}", basis)
    inp(S, f"G{r}", q0, fmt="0"); inp(S, f"H{r}", unit, fmt=USD)
    put(S, f"I{r}", f"=G{r}*H{r}", fmt=USD)
    if qr is None:  # スポットプレスは計算シートから
        c = put(S, f"J{r}", "='スポットプレス'!C14", fmt="0", color=GREEN)
        c.comment = Comment("『スポットプレス』シートの追加必要台数を参照", "draft")
    else:
        inp(S, f"J{r}", qr, fmt="0")
    put(S, f"K{r}", f"=J{r}*H{r}", fmt=USD)
    put(S, f"L{r}", f'=IF(D{r}="必須",K{r},0)', fmt=USD)
    put(S, f"M{r}", f"=MAX(G{r},J{r})*H{r}", fmt=USD)
    put(S, f"N{r}", f"=K{r}*{RATE}", fmt=JPY)
    put(S, f"O{r}", f"=N{r}/{YEARS}", fmt=JPY)
    put(S, f"P{r}", step, align="center"); put(S, f"Q{r}", status); put(S, f"R{r}", sheet, align="center")
    if cat == "必須":
        for col in "ABCD": S[f"{col}{r}"].fill = SUB
last = first + len(items) - 1
tr = last + 1
put(S, f"B{tr}", "合計", bold=True, fill=GRY)
for col in "ACDEFHPQR": put(S, f"{col}{tr}", None, fill=GRY)
for col, f in [("G", "0"), ("I", USD), ("J", "0"), ("K", USD), ("L", USD), ("M", USD), ("N", JPY), ("O", JPY)]:
    put(S, f"{col}{tr}", f"=SUM({col}{first}:{col}{last})", bold=True, fmt=f, fill=GRY)
put(S, f"A{tr+1}", "注1：当初案＝予算案ファイル（2026/4/3版）の優先①台数・単価(USD)（VA後単価）。当初案合計は同ファイルN25（267,343USD）と一致。", size=9, border=False, wrap=False)
put(S, f"A{tr+2}", "注2：推奨案台数は本ドラフトの提案値（青字）。スポットプレスは『スポットプレス』シートの必要台数に連動。全体案＝当初案と推奨案の大きい方。", size=9, border=False, wrap=False)
put(S, f"A{tr+3}", "注3：優先②の機材、ダウン充填機（優先度引下げ済）、自動延反機（26年度見送り）、刺繍機は本表の対象外。", size=9, border=False, wrap=False)

# シナリオ集計（上部）
put(S, "A10", "投資シナリオ", bold=True, border=False)
header(S, 11, ["案", "内容", "投資額(USD)", "投資額(円)", "目標上限との差(円)", "APIC負担控除後(円)", "目標判定", "円換算(感応度レート)"])
rng = lambda col: f"{col}{first}:{col}{last}"
scen = [
    ("必須案", "安全・顧客要求・生産継続のみ（第1次決裁対象）", f"=SUM({rng('L')})"),
    ("推奨案", "必須＋改善＋試験導入（根拠未整備の戦略機材は次回以降）", f"=SUM({rng('K')})"),
    ("全体案", "当初案全機種（スポットプレス必要台数を反映）", f"=SUM({rng('M')})"),
    ("参考：当初案", "予算案ファイル（2026/4/3版）優先①そのまま", f"=SUM({rng('I')})"),
]
for i, (n, d, f) in enumerate(scen):
    r = 12 + i
    put(S, f"A{r}", n, bold=True, align="center"); put(S, f"B{r}", d)
    put(S, f"C{r}", f, fmt=USD); put(S, f"D{r}", f"=C{r}*{RATE}", fmt=JPY)
    put(S, f"E{r}", f"={TARGET}-D{r}", fmt=JPY)
    put(S, f"F{r}", f"=MAX(C{r}-{APIC},0)*{RATE}", fmt=JPY)
    put(S, f"G{r}", f'=IF(D{r}<={TARGET},"目標内",IF(F{r}<={TARGET},"APIC負担控除後なら目標内","目標超過"))', align="center")
    put(S, f"H{r}", f"=C{r}*{RATE2}", fmt=JPY)
put(S, "A16", "※『APIC負担控除後』は50,000USDを全額TI投資額から控除した場合の参考値。目標4,000万円がTI負担分を指すか総額を指すかは要確認。", size=9, border=False, wrap=False)
widths(S, {"A": 9, "B": 34, "C": 22, "D": 11, "E": 20, "F": 52, "G": 9, "H": 11, "I": 13, "J": 9, "K": 13,
           "L": 13, "M": 13, "N": 15, "O": 14, "P": 10, "Q": 24, "R": 13})
S.freeze_panes = f"C{H0+1}"
for r in range(first, last + 1):
    S.row_dimensions[r].height = 42

# ---------------------------------------------------------------- 要確認事項
Q = wb.create_sheet("要確認事項")
title(Q, "前回会議（2026/3/25）要確認事項への対応状況", "①～③、⑤、⑦が今回の説明対象。④⑥は参考。")
legend(Q, "A3")
header(Q, 5, ["No.", "要確認事項", "回答方針", "必要データ", "状況", "担当", "期限", "詳細シート"])
qs = [
    ("①", "機材導入で削減できる年間外部レンタル費用。パターンシーマ・シーム機(spec-D)による要員減の数値",
     "レンタル実績(2024～2026)から投資で回避できる分のみを削減額とする。省人化は『工程秒数削減→必要工数削減→余力人員の再配置』として算出（人員削減ではない）",
     "レンタル明細(機種・台数・期間・D/N・ライン)／3方式の工程秒数・作業者拘束時間・年間数量", "作成中", "レンタル／パターンシーマ"),
    ("②", "パターンシーマの動線・工員配置・縫製中の動きを含む運用設計",
     "作業者→セット→自動縫製→自動縫製中に次製品準備→取り出し、の各秒数を実測し、1人当たり担当台数を実測値で設定（『1人で複数台』を前提にしない）",
     "工程別の実測秒数、ライン内配置図", "作成中", "パターンシーマ"),
    ("③", "スポットプレスを最優先で入替。必要総数の計算根拠と現状の運用状況",
     "対象ライン数×必要台数/ライン＋共用分＋予備・保全分＝必要総数。更新分と純増分を分けて提示。安全性を主理由とし、ROIは求めない",
     "旧型の台数・年式・安全装置有無・修理可否・使用ライン、写真、ヒヤリハット", "データ待ち", "スポットプレス"),
    ("④", "ダウン充填機は投入量に鑑み優先度を下げる", "優先①から除外済", "―", "対応済", "―"),
    ("⑤", "平プレス(No.10)のプレス精度の検証方法",
     "水平器による4箇所測定、紙挟みテストによる圧着均一性、温度・圧力確認、メーカー校正を組み合わせ、判定基準・頻度・担当を明記",
     "メーカーの判定基準値・校正方法", "案作成済（基準値確認中）", "平プレス検証"),
    ("⑥", "定常投資と別に48百万円枠を利用できるか", "本資料は定常投資のみを対象", "―", "対象外", "―"),
    ("⑦", "縮絨機の導入検討（社内事例含む）",
     "QN事例確認→対象生地→使用条件→品質改善効果→APIC適用性→小規模試験→導入/見送り判断。本年度は試験後判断とする",
     "QNの対象生地・条件・効果", "事例確認中", "縮絨機"),
    ("新", "APIC負担50,000USDの扱い", "対象機材・資産計上先・償却の扱いを確定し、TI側償却影響を算出", "経理確認", "要確認", "PL影響"),
]
for i, (no, item, policy, data, st, sh) in enumerate(qs):
    r = 6 + i
    put(Q, f"A{r}", no, align="center"); put(Q, f"B{r}", item); put(Q, f"C{r}", policy); put(Q, f"D{r}", data)
    inp(Q, f"E{r}", st); inp(Q, f"F{r}", None, yellow=True); inp(Q, f"G{r}", None, yellow=True); put(Q, f"H{r}", sh)
    Q.row_dimensions[r].height = 60
widths(Q, {"A": 5, "B": 36, "C": 60, "D": 36, "E": 16, "F": 10, "G": 10, "H": 18})

# ---------------------------------------------------------------- ロックミシン
L = wb.create_sheet("ロックミシン")
title(L, "ロックミシン（オーバーロック／インターロック）必要台数と不足", "不足は試算上だけではなく、第2工場借用および外部レンタルで実運用上すでに顕在化している。")
legend(L, "A3")
header(L, 5, ["項目", "オーバーロック", "インターロック", "備考"])
rows = [
    ("通常ライン数", 31, 31, "APIC通常稼働", "in"),
    ("ピーク時ライン数", 35, 35, "繁忙期34～35ライン（軽量品は32人ラインを16人×2に分割）", "in"),
    ("標準台数／ライン（通常品・Seam・Down）", 3, 2, "品種別標準使用台数", "in"),
    ("標準台数／ライン（WB・1枚もの軽量品）", 4, 2, "軽量品は全パーツ端処理のためOL4台", "in"),
    ("軽量品の構成比", 0.2, 0.2, "製品構成の仮定（80%／20%）。実績で要確認", "inpct"),
    ("加重平均台数／ライン", "=B8*(1-B10)+B9*B10", "=C8*(1-C10)+C9*C10", "", "f1"),
    ("通常時必要台数", "=B6*B11", "=C6*C11", "", "f1"),
    ("ピーク時必要台数", "=B7*B11", "=C7*C11", "", "f1"),
    ("現有台数", 102, 68, "予算案ファイル S列", "in"),
    ("老朽化重大故障（修理復旧不可）", 4, 0, "OL：C判定4台。IL：要確認（実稼働63～64台の情報あり）", "iny"),
    ("サンプル固定", 12, 4, "", "in"),
    ("生産使用可能台数", "=B14-B15-B16", "=C14-C15-C16", "", "f0"),
    ("通常時の不足", "=MAX(B12-B17,0)", "=MAX(C12-C17,0)", "通常時にも不足が出る場合、その補填方法の説明が必要", "f1"),
    ("ピーク時の不足", "=MAX(B13-B17,0)", "=MAX(C13-C17,0)", "", "f1"),
    ("今回購入台数", "='総括'!J19", "='総括'!J20", "『総括』推奨案台数", "link"),
    ("購入後のピーク時残不足", "=MAX(B19-B20,0)", "=MAX(C19-C20,0)", "第2工場借用（10～12台）・短期レンタルで対応", "f1"),
]
for i, (k, a, b, note, kind) in enumerate(rows):
    r = 6 + i
    put(L, f"A{r}", k, bold=kind.startswith("f"))
    for col, v in (("B", a), ("C", b)):
        if kind == "in": inp(L, f"{col}{r}", v, fmt="0")
        elif kind == "iny": inp(L, f"{col}{r}", v, fmt="0", yellow=True)
        elif kind == "inpct": inp(L, f"{col}{r}", v, fmt=PCT, yellow=True)
        elif kind == "link": put(L, f"{col}{r}", v, fmt="0", color=GREEN)
        else: put(L, f"{col}{r}", v, fmt=NUM1 if kind == "f1" else "0")
    put(L, f"D{r}", note)
put(L, "A23", "運用上の実績（不足の顕在化）", bold=True, border=False)
header(L, 24, ["実績", "台数", "期間", "備考"])
ev = [("APIC第2工場からの借用", "10～12台", "繁忙期", "無償借用。受注増で継続貸出は困難見込み"),
      ("外部レンタル（ロックミシン）", "約10台", "2024年10～12月", "金額・D/Nは『レンタル』シートで集計")]
for i, e in enumerate(ev):
    for j, v in enumerate(e): put(L, f"{'ABCD'[j]}{25+i}", v)

put(L, "A28", "サーボモーター化の副次効果（電力）", bold=True, border=False)
header(L, 29, ["項目", "クラッチモーター（現行）", "サーボモーター（新）", "備考"])
el = [
    ("縫製時消費電力(W)", 350, 200, "定格はクラッチ400W／サーボ550Wだが、サーボはペダル操作時のみ駆動"),
    ("非縫製時消費電力(W)", 250, 2, "クラッチは非縫製時も常時回転、サーボは待機電力のみ"),
]
for i, (k, a, b, n) in enumerate(el):
    r = 30 + i; put(L, f"A{r}", k); inp(L, f"B{r}", a, fmt="0"); inp(L, f"C{r}", b, fmt="0"); put(L, f"D{r}", n)
put(L, "A32", "縫製稼働率（稼働時間中の縫製時間比率）"); inp(L, "B32", 0.5, fmt=PCT, yellow=True)
put(L, "C32", None); put(L, "D32", "既存試算（2.400kWh/日）は50%相当。実測で要確認（低いほどサーボの効果は大）")
put(L, "A33", "1日消費電力(kWh/台)", bold=True)
put(L, "B33", f"={HOURS}*(B30*$B$32+B31*(1-$B$32))/1000", fmt=NUM3)
put(L, "C33", f"={HOURS}*(C30*$B$32+C31*(1-$B$32))/1000", fmt=NUM3); put(L, "D33", "")
put(L, "A34", "年間削減電力(kWh/台)"); put(L, "B34", f"=(B33-C33)*{DAYS}", fmt=NUM1); put(L, "C34", None); put(L, "D34", "")
put(L, "A35", "年間削減額(USD/台)"); put(L, "B35", f"=B34*{ELEC}", fmt=USD2); put(L, "C35", None); put(L, "D35", "")
put(L, "A36", "年間削減額(USD/購入台数分)", bold=True); put(L, "B36", "=B35*(B20+C20)", fmt=USD2); put(L, "C36", None)
put(L, "D36", "電力は副次効果。主目的は機材不足解消と品質（下記）")
put(L, "A38", "品質・操作性（定性）", bold=True, border=False)
header(L, 39, ["項目", "クラッチモーター（現行）", "サーボモーター（新）", "効果"])
qual = [("ペダルOFF時", "惰性回転で数針進む", "即停止", "縫いズレ低減"),
        ("曲線部の微調整", "難しい（熟練依存）", "容易", "品質安定・属人性低減"),
        ("自動糸切り", "なし", "あり", "工数削減")]
for i, e in enumerate(qual):
    for j, v in enumerate(e): put(L, f"{'ABCD'[j]}{40+i}", v)
widths(L, {"A": 40, "B": 22, "C": 22, "D": 70})

# ---------------------------------------------------------------- スポットプレス
SP = wb.create_sheet("スポットプレス")
title(SP, "スポットプレス 必要台数（安全投資・最優先）", "対象ライン数×必要台数/ライン＋共用分＋予備・保全分＝必要総数。安全上の理由で入替のため、ROIは主根拠としない。")
legend(SP, "A3")
header(SP, 5, ["項目", "入力", "計算", "備考"])
sp = [
    ("対象ライン数", "in", None, "スポットプレスを使用するライン数"),
    ("必要台数／ライン", "in", None, "工程別の使用台数から設定"),
    ("共用分", "in", None, "シーム後プレス等の共用機"),
    ("予備・保全分", "in", None, "故障時の代替"),
    ("積上げ必要総数", None, "=IF(B6*B7+B8+B9>0,B6*B7+B8+B9,0)", "上記入力から自動計算"),
    ("会議時の概算必要総数", "c45", None, "2026/3/25会議での概算（45～50台）。積上げ値が入れば積上げ値を優先"),
    ("採用する必要総数", None, "=IF(C10>0,C10,B11)", ""),
    ("現有（新型30×10cm）", "c15", None, "予算案ファイル R14"),
    ("追加必要台数", None, "=MAX(C12-B13,0)", "『総括』推奨案台数に連動"),
    ("うち更新分（旧型入替）", None, '=IF(B17="","",MIN(B17,C14))', "旧型台数の入力後に計算"),
    ("うち純増分", None, '=IF(C15="","",C14-C15)', ""),
    ("旧型（安全装置なし）台数", "in", None, "廃棄対象。全数の現況確認要"),
    ("当初案との差（台）", None, "=C14-'総括'!G21", "当初案10台"),
]
for i, (k, mode, f, n) in enumerate(sp):
    r = 6 + i
    put(SP, f"A{r}", k, bold=(f is not None))
    if mode == "in": inp(SP, f"B{r}", None, fmt="0", yellow=True)
    elif mode == "c45": inp(SP, f"B{r}", 45, fmt="0", note="会議での口頭概算（下限値）。積上げ根拠で置換する", yellow=True)
    elif mode == "c15": inp(SP, f"B{r}", 15, fmt="0")
    else: put(SP, f"B{r}", None)
    put(SP, f"C{r}", f, fmt="0") if f else put(SP, f"C{r}", None)
    put(SP, f"D{r}", n)
put(SP, "A21", "旧型スポットプレス 現況一覧（工場記入）", bold=True, border=False)
header(SP, 22, ["管理番号", "年式", "安全装置 有/無", "後付け可否", "修理可否", "使用ライン／工程", "故障・ヒヤリハット記録", "写真No."])
for r in range(23, 29):
    for col in "ABCDEFGH": inp(SP, f"{col}{r}", None, yellow=True)
inp(SP, "A23", "例：SP-001"); inp(SP, "B23", "2011年頃"); inp(SP, "C23", "無"); inp(SP, "D23", "不可（メーカー見解）")
inp(SP, "E23", "部品供給なし"); inp(SP, "F23", "L05／シーム後プレス"); inp(SP, "G23", "記入例"); inp(SP, "H23", "P-01")
put(SP, "A30", "安全面の論点：旧型はアームが斜め方向に動作し圧力計算ができない／安全装置なし・後付け不可／メーカー廃棄推奨／圧力ムラによる再プレス増加。", size=9, border=False, wrap=False)
widths(SP, {"A": 26, "B": 12, "C": 14, "D": 60, "E": 14, "F": 20, "G": 24, "H": 10})

# ---------------------------------------------------------------- パターンシーマ
PS = wb.create_sheet("パターンシーマ")
title(PS, "パターンシーマ 3方式比較（手作業／改善治具＋1本針／パターンシーマ）",
      "省人化＝工程秒数削減→必要工数削減→必要人員減→余力人員の再配置→工場全体の生産性向上（人員削減ではない）")
legend(PS, "A3")
put(PS, "A4", "注：全ケースでサイクルタイムが短縮するわけではない（同等または数秒長いケースあり）。主効果は工程集約・作業者拘束時間削減・属人性排除・習熟期間短縮・品質安定。", size=9, border=False, wrap=False)
hd = ["顧客", "部位・工程", "機種", "年間数量(枚)", "手作業 工程数", "手作業 秒/枚", "治具 工程数", "治具 秒/枚",
      "PS 工程数", "PS 秒/枚", "PS 作業者拘束秒/枚", "削減秒/枚（手作業→PS）", "年間削減工数(秒)",
      "省人工数(人)", "拘束時間削減秒/枚", "年間拘束削減(秒)", "拘束ベース再配置可能人員(人)"]
header(PS, 6, hd)
for r in range(7, 22):
    for j in range(11):
        col = "ABCDEFGHIJK"[j]
        inp(PS, f"{col}{r}", None, fmt=None if j < 3 else "#,##0", yellow=(r != 7))
    put(PS, f"L{r}", f'=IF(OR(F{r}="",J{r}=""),"",F{r}-J{r})', fmt=NUM1)
    put(PS, f"M{r}", f'=IF(OR(L{r}="",D{r}=""),"",L{r}*D{r})', fmt=JPY)
    put(PS, f"N{r}", f'=IF(M{r}="","",M{r}/{SECYR})', fmt="0.00")
    put(PS, f"O{r}", f'=IF(OR(F{r}="",K{r}=""),"",F{r}-K{r})', fmt=NUM1)
    put(PS, f"P{r}", f'=IF(OR(O{r}="",D{r}=""),"",O{r}*D{r})', fmt=JPY)
    put(PS, f"Q{r}", f'=IF(P{r}="","",P{r}/{SECYR})', fmt="0.00")
for col, v in zip("ABCEFGHIJ", ["例(実データに置換)", "ファスナー", "中型 60×40", 6, 340, 3, 330, 2, 310]):
    inp(PS, f"{col}7", v)
put(PS, "A22", "合計", bold=True, fill=GRY)
for col in "BCEFGHIJKLO": put(PS, f"{col}22", None, fill=GRY)
for col, f in [("D", JPY), ("M", JPY), ("N", "0.00"), ("P", JPY), ("Q", "0.00")]:
    put(PS, f"{col}22", f"=SUM({col}7:{col}21)", bold=True, fmt=f, fill=GRY)
put(PS, "A23", "計算式：年間削減工数(秒)＝(手作業秒－PS秒)×年間数量／省人工数＝年間削減工数÷1人年間有効稼働秒（『パラメータ』B16）。", size=9, border=False, wrap=False)
put(PS, "A24", "拘束ベース：PS自動縫製中は作業者が拘束されないため、作業者拘束秒で再配置可能人員を別途算出（動線設計後の実測値を使用）。", size=9, border=False, wrap=False)

put(PS, "A26", "運用設計（要確認事項②）：1サイクルの作業者の動き（実測）", bold=True, border=False)
header(PS, 27, ["ステップ", "内容", "秒（実測）", "作業者拘束", "備考"])
steps = [("1", "パーツ・型のセット", "有", ""), ("2", "自動縫製", "無", "この間に次製品の準備"),
         ("3", "次製品の準備（自動縫製中）", "有（並行）", "動線・配置で並行作業が成立するか確認"),
         ("4", "取り出し・検品", "有", "")]
for i, (a, b, d, e) in enumerate(steps):
    r = 28 + i
    put(PS, f"A{r}", a, align="center"); put(PS, f"B{r}", b); inp(PS, f"C{r}", None, fmt="0", yellow=True); put(PS, f"D{r}", d); put(PS, f"E{r}", e)
put(PS, "A32", "サイクルタイム", bold=True); put(PS, "B32", "＝セット＋自動縫製＋取り出し（準備は並行）")
put(PS, "C32", "=C28+C29+C31", fmt="0"); put(PS, "D32", None); put(PS, "E32", None)
put(PS, "A33", "1人当たり担当可能台数（目安）", bold=True); put(PS, "B33", "＝サイクルタイム÷作業者拘束秒（セット＋準備＋取り出し）")
put(PS, "C33", '=IF(C28+C30+C31>0,C32/(C28+C30+C31),"")', fmt="0.0"); put(PS, "D33", None)
put(PS, "E33", "理論値。『1人で複数台』は動線・配置を設計し実測で確認してから採用")
widths(PS, {"A": 18, "B": 34, "C": 13, "D": 11, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10, "J": 10, "K": 12,
            "L": 12, "M": 14, "N": 10, "O": 12, "P": 14, "Q": 14})
PS.freeze_panes = "D7"

# ---------------------------------------------------------------- レンタル
R = wb.create_sheet("レンタル")
title(R, "外部レンタル費・外注費と投資による回避可能額", "投資で回避できる分のみを削減効果とする（レンタル総額の100%を削減額としない）。")
legend(R, "A3")
header(R, 5, ["年", "機種", "台数", "期間", "金額(USD)", "D/N・請求書No.", "ライン／品番", "対応する投資機材", "回避可能率", "回避可能額(USD)"])
rent = [("2025", "大型パターンシーマ", None, None, 3016.54, None, None, "No.5 パターンシーマ130×95"),
        ("2025", "その他機材（内訳確認中）", None, None, "=4890.75-E6", None, None, None),
        ("2024", "ロックミシン", 10, "10～12月", None, None, None, "No.1,2 ロックミシン")]
for i in range(12):
    r = 6 + i
    vals = rent[i] if i < len(rent) else (None,) * 8
    for j, v in enumerate(vals):
        col = "ABCDEFGH"[j]
        if j == 4 and isinstance(v, str):
            put(R, f"{col}{r}", v, fmt=USD2)
        else:
            inp(R, f"{col}{r}", v, fmt=USD2 if j == 4 else None, yellow=(v is None))
    inp(R, f"I{r}", None, fmt=PCT, yellow=True)
    put(R, f"J{r}", f"=IF(I{r}=\"\",0,E{r}*I{r})", fmt=USD2)
put(R, "A18", "合計", bold=True, fill=GRY)
for col in "BCDFGHI": put(R, f"{col}18", None, fill=GRY)
put(R, "E18", "=SUM(E6:E17)", bold=True, fmt=USD2, fill=GRY); put(R, "J18", "=SUM(J6:J17)", bold=True, fmt=USD2, fill=GRY)
put(R, "A19", "出典：2025年分は『APIC_2025年 機械レンタル費用総額.xlsx』の集計値（総額4,890.75USD、うち大型パターンシーマ3,016.54USD）。2024・2026年分は明細を追記。", size=9, border=False, wrap=False)

put(R, "A21", "投資理由書の記載額との照合", bold=True, border=False)
header(R, 22, ["項目", "金額(USD/年)", "備考"])
put(R, "A23", "No.6 本縫針送りミシン 理由書記載のレンタル削減額"); inp(R, "B23", "=112*12*10", fmt=USD)
put(R, "C23", "理由書：112USD×12か月×10台。月額112USD・常時10台レンタルの根拠となる請求書を要確認")
put(R, "A24", "2025年 外部レンタル実績総額"); put(R, "B24", "=4890.75", fmt=USD); put(R, "C24", "")
put(R, "A25", "差異（理由書記載額－実績総額）", bold=True); put(R, "B25", "=B23-B24", fmt=USD, bold=True)
put(R, "C25", "差異がプラスの場合、理由書の削減額が実績を上回るため、別請求（長期レンタル等）の有無を確認し、なければ記載額を修正")
put(R, "A27", "参考：外注費（レンタルではない）", bold=True, border=False)
header(R, 28, ["項目", "金額(USD)", "備考"])
put(R, "A29", "キルティング外注（25FW）"); inp(R, "B29", 3000, fmt=USD)
put(R, "C29", "予算案ファイル R8『約3千USD、納期1か月前後』。大型パターンシーマのレンタル費と重複していないか確認")
widths(R, {"A": 44, "B": 26, "C": 8, "D": 12, "E": 14, "F": 16, "G": 16, "H": 26, "I": 11, "J": 15})
R.column_dimensions["C"].width = 10

# ---------------------------------------------------------------- 平プレス検証
FP = wb.create_sheet("平プレス検証")
title(FP, "No.10 平プレス（125×83cm）プレス精度の検証方法（要確認事項⑤）",
      "大型機のため、圧力の均一性とキャリブレーションを導入時・定期で確認する。")
legend(FP, "A3")
header(FP, 5, ["No.", "検証項目", "検証方法", "判定基準", "確認頻度", "担当"])
fp = [
    ("1", "盤面の水平度", "水平器でプレス盤の4隅（4箇所）を測定", "メーカー基準値（要確認）", "導入時・月1回", "保全"),
    ("2", "圧着の均一性", "紙挟みテスト：4隅＋中央に紙を挟んでプレスし、圧痕・引抜き抵抗の均一性を確認", "全測定点で圧痕が均一であること", "導入時・品番切替時", "品質"),
    ("3", "温度分布", "表面温度計で上下盤の複数点を測定", "設定温度±（メーカー基準値 要確認）℃", "始業時", "ライン"),
    ("4", "圧力", "圧力計の表示値と設定値を確認", "設定値±（メーカー基準値 要確認）", "始業時", "ライン"),
    ("5", "キャリブレーション", "メーカーによる定期校正・保守", "メーカー基準", "年1回（要確認）", "保全／メーカー"),
    ("6", "製品での確認", "接着強度試験・外観（剥離・シワ）確認", "顧客基準", "品番立上げ時", "品質"),
]
for i, row in enumerate(fp):
    r = 6 + i
    for j, v in enumerate(row):
        col = "ABCDEF"[j]
        if "要確認" in v: inp(FP, f"{col}{r}", v, yellow=True)
        else: put(FP, f"{col}{r}", v, align="center" if j == 0 else None)
    FP.row_dimensions[r].height = 36
widths(FP, {"A": 5, "B": 18, "C": 60, "D": 34, "E": 18, "F": 14})

# ---------------------------------------------------------------- 縮絨機
SH = wb.create_sheet("縮絨機")
title(SH, "縮絨機 検討ステップ（要確認事項⑦）", "本年度は購入判断とせず、社内事例確認と小規模試験の後に導入／見送りを判断する。")
header(SH, 5, ["ステップ", "内容", "確認結果", "判定"])
sh = [("1", "QN事例確認：どの生地で、どの条件で使用しているか"), ("2", "対象生地の特定（APICで不安定な生地サプライヤー）"),
      ("3", "使用条件（温度・時間・張力等）"), ("4", "品質改善効果（寸法安定性・外観）"),
      ("5", "APICへの適用性（過去のウール試験では効果不明確）"), ("6", "小規模試験の実施"), ("7", "導入／見送り判断")]
for i, (a, b) in enumerate(sh):
    r = 6 + i
    put(SH, f"A{r}", a, align="center"); put(SH, f"B{r}", b); inp(SH, f"C{r}", None, yellow=True); inp(SH, f"D{r}", None, yellow=True)
widths(SH, {"A": 8, "B": 60, "C": 40, "D": 14})

# ---------------------------------------------------------------- 比較シート（ダウン充填装置比較の定量モデルを踏襲）
from openpyxl.chart import LineChart, Reference

def build_compare(name, ttl, sub, heads, data, cases, pros, cons, notes, na_cols=(), labels=None):
    """heads: 3列の見出し（現行／代替案／新設備案）。data: 行キー -> (C,D,E) 値（数値=入力、'='始まり=式、None=要記入）。"""
    W = wb.create_sheet(name)
    put(W, "A1", ttl, bold=True, size=14, border=False, wrap=False)
    put(W, "A2", sub, size=9, border=False, wrap=False, color="595959")
    put(W, "A3", "凡例：青字＝入力値／黄色＝工場で実測・記入が必要／黒字＝計算式。前提（稼働日・稼働率・余裕率・人件費）は『パラメータ』シート。",
        size=9, border=False, wrap=False, color="595959")
    header(W, 5, ["区分", "項目", heads[0], heads[1], heads[2], "備考"])
    lab = {"name": "方式・機種", "price": "設備価格(USD/台)", "proc": "対応工程", "steps": "工程数",
           "ppl": "1台あたり作業人数(人/台)", "sec": "工程作業時間(秒/枚)", "bind": "作業者拘束時間(秒/枚)"}
    if labels: lab.update(labels)
    order = [("基本情報", "name"), ("", "price"), ("", "proc"), ("", "steps"), ("", "ppl"),
             ("生産性", "sec"), ("", "bind")]
    rowof = {}
    for i, (sec, key) in enumerate(order):
        r = 6 + i; rowof[key] = r
        put(W, f"A{r}", sec, bold=True); put(W, f"B{r}", lab[key])
        vals, note = data.get(key, ((None, None, None), ""))
        for j, col in enumerate("CDE"):
            v = vals[j]
            if col in na_cols and v is None:
                put(W, f"{col}{r}", "－", align="center", fill=GRY)
            elif isinstance(v, str) and v.startswith("="):
                put(W, f"{col}{r}", v, fmt=NUM1 if key in ("sec", "bind") else None)
            elif v is None:
                inp(W, f"{col}{r}", None, fmt=USD if key == "price" else "#,##0", yellow=True)
            else:
                inp(W, f"{col}{r}", v, fmt=USD if key == "price" else ("#,##0" if isinstance(v, (int, float)) else None))
        put(W, f"F{r}", note)
    r_cap, r_cap_p, r_dep = 13, 14, 15
    put(W, f"A{r_cap}", "", bold=True); put(W, f"B{r_cap}", "1台当たり日産能力(枚/台)")
    put(W, f"B{r_cap_p}", "1人当たり日産能力(枚/人)")
    put(W, f"A{r_dep}", "加工コスト比較", bold=True); put(W, f"B{r_dep}", "減価償却費(USD/月・台)")
    for col in "CDE":
        if col in na_cols:
            for rr_ in (r_cap, r_cap_p, r_dep): put(W, f"{col}{rr_}", "－", align="center", fill=GRY)
            continue
        put(W, f"{col}{r_cap}", f'=IF(N({col}{rowof["sec"]})=0,"",{DAYSEC}/{col}{rowof["sec"]})', fmt=NUM1)
        put(W, f"{col}{r_cap_p}", f'=IF(N({col}{rowof["bind"]})=0,"",{DAYSEC}/{col}{rowof["bind"]})', fmt=NUM1)
        put(W, f"{col}{r_dep}", f'=IF({col}{rowof["price"]}="","",N({col}{rowof["price"]})/{YEARS}/12)', fmt=USD2)
    put(W, f"F{r_cap}", "＝1日有効作業秒÷工程作業時間"); put(W, f"F{r_cap_p}", "＝1日有効作業秒÷作業者拘束時間")
    put(W, f"F{r_dep}", "＝設備価格÷償却年数÷12")
    # 損益分岐ケース
    case_rows = []
    base = 20
    caseitems = ["日産目標(枚/日)", "所要設備台数(台)", "装置総額(USD)", "〃 差額(USD) ※現行－各案",
                 "作業人数(人)", "人件費(USD/月)", "〃 差額(USD/月) ※現行－各案", "損益分岐点月数(月)"]
    for k, target in enumerate(cases):
        b = base + k * 8; case_rows.append(b)
        for i, t in enumerate(caseitems):
            put(W, f"A{b+i}", f"損益分岐 ケース{k+1}" if i == 0 else "", bold=True); put(W, f"B{b+i}", t)
        inp(W, f"C{b}", target, fmt="#,##0")
        for col in "DE": put(W, f"{col}{b}", f"=$C${b}", fmt="#,##0")
        for col in "CDE":
            na = col in na_cols
            if na:
                for i in range(1, 8): put(W, f"{col}{b+i}", "－", align="center", fill=GRY)
                continue
            put(W, f"{col}{b+1}", f'=IF({col}{r_cap}="","",ROUNDUP({col}{b}/{col}{r_cap},0))', fmt="0")
            put(W, f"{col}{b+2}", f'=IF({col}{b+1}="","",{col}{b+1}*N({col}{rowof["price"]}))', fmt=USD)
            put(W, f"{col}{b+4}", f'=IF(N({col}{rowof["bind"]})=0,"",ROUNDUP({col}{b}*{col}{rowof["bind"]}/{DAYSEC}*2,0)/2)', fmt="0.0")
            put(W, f"{col}{b+5}", f'=IF(OR({col}{b+4}="",{LABOR}=""),"",{col}{b+4}*{LABOR})', fmt=USD)
            if col == "C":
                for i in (3, 6, 7): put(W, f"C{b+i}", "（基準）", align="center", fill=GRY)
            else:
                put(W, f"{col}{b+3}", f'=IF(OR($C{b+2}="",{col}{b+2}=""),"",$C{b+2}-{col}{b+2})', fmt=USD)
                put(W, f"{col}{b+6}", f'=IF(OR($C{b+5}="",{col}{b+5}=""),"",$C{b+5}-{col}{b+5})', fmt=USD)
                put(W, f"{col}{b+7}",
                    f'=IF(OR({col}{b+3}="",{col}{b+6}=""),"",IF({col}{b+6}<=0,"回収不可",IF({col}{b+3}>=0,0,-{col}{b+3}/({UTIL}*{col}{b+6}))))',
                    fmt=NUM1)
        put(W, f"F{b}", "試算ケース（任意に変更可）"); put(W, f"F{b+4}", "＝ROUNDUP(日産×拘束秒÷1日有効作業秒×2)÷2（0.5人単位）")
        put(W, f"F{b+7}", "＝－装置差額÷(稼働率×人件費差額)")
    # 加工コスト（ケース別）
    put(W, "A16", "試算", bold=True)
    for k, b in enumerate(case_rows):
        r = 16 + k
        if k: put(W, f"A{r}", "")
        put(W, f"B{r}", f'="加工コスト(USD/枚) 日産="&TEXT(C{b},"#,##0")&"枚"')
        for col in "CDE":
            if col in na_cols: put(W, f"{col}{r}", "－", align="center", fill=GRY); continue
            put(W, f"{col}{r}", f'=IF(OR({col}{b+1}="",{col}{b+5}=""),"",(N({col}{r_dep})*{col}{b+1}+{col}{b+5})/({col}{b}*{MDAYS}*{UTIL}))', fmt="0.000")
        put(W, f"F{r}", "＝(減価償却費×所要台数＋人件費)÷(日産×月間稼働日×稼働率)" if k == 0 else "")
    put(W, "A19", ""); put(W, "B19", ""); [put(W, f"{c}19", None) for c in "CDEF"]
    rr = base + len(cases) * 8
    put(W, f"A{rr}", "特徴", bold=True); put(W, f"B{rr}", "長所"); put(W, f"B{rr+1}", "短所"); put(W, f"A{rr+1}", "")
    for j, col in enumerate("CDE"):
        put(W, f"{col}{rr}", pros[j]); put(W, f"{col}{rr+1}", cons[j])
    put(W, f"F{rr}", None); put(W, f"F{rr+1}", None)
    W.row_dimensions[rr].height = 48; W.row_dimensions[rr + 1].height = 48
    for i, n in enumerate(notes):
        put(W, f"A{rr+3+i}", n, size=9, border=False, wrap=False)
    # 月次ランニングコスト累計（ケース2）→ グラフ
    b2 = case_rows[1]
    put(W, "H5", "月", bold=True, fill=SUB, align="center")
    for j, col in enumerate("CDE"):
        put(W, f"{'IJK'[j]}5", f"={col}5", bold=True, fill=SUB)
    for m in range(0, 61):
        r = 6 + m
        put(W, f"H{r}", m, fmt="0", align="center")
        for j, col in enumerate("CDE"):
            tgt = "IJK"[j]
            if col in na_cols: put(W, f"{tgt}{r}", None); continue
            put(W, f"{tgt}{r}", f"=N(${col}${b2+2})+N(${col}${b2+5})*{UTIL}*$H{r}", fmt=USD)
    put(W, "H67", "※ケース2の累計コスト＝装置総額＋人件費×稼働率×月数", size=9, border=False, wrap=False)
    ch = LineChart(); ch.title = "月次ランニングコスト累計（ケース2）"; ch.roundedCorners = False
    ch.y_axis.title = "USD"; ch.x_axis.title = "月"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    for j, col in enumerate("CDE"):
        if col in na_cols: continue
        tcol = 9 + j
        ch.add_data(Reference(W, min_col=tcol, min_row=5, max_row=66), titles_from_data=True)
    ch.set_categories(Reference(W, min_col=8, min_row=6, max_row=66))
    ch.legend.position = "b"; ch.legend.overlay = False; ch.title.overlay = False
    ch.x_axis.title.overlay = False; ch.y_axis.title.overlay = False
    ch.height = 9; ch.width = 18
    W.add_chart(ch, "M5")
    widths(W, {"A": 13, "B": 34, "C": 24, "D": 24, "E": 24, "F": 46, "G": 3, "H": 6, "I": 13, "J": 13, "K": 13})
    W.freeze_panes = "C6"
    return W

build_compare(
    "比較_パターンシーマ", "パターンシーマ 3方式比較（定量モデル）",
    "ダウン充填装置比較と同じ構成：基本情報→生産性→加工コスト→損益分岐（3ケース）→特徴。工程ごと（ファスナー・ポケット等）にシートを複製して使用。",
    ["手作業（現行）", "改善治具＋1本針（代替案）", "パターンシーマ（新設備案）"],
    {"name": (("本縫いミシン・手作業", "改善治具＋1本針", "Hikari HMS-6040（60×40）"), "大型（HMS-13095A/RND）はシートを複製して作成"),
     "price": ((None, None, 4450), "現行は既存機の取得価額（償却済なら0）、治具は製作費"),
     "proc": (("ファスナー（例）", "ファスナー（例）", "ファスナー（例）"), "実データの工程名に置換"),
     "steps": ((6, 3, 2), "handoffの例（手作業6工程・治具3工程・PS2工程）。実データに置換"),
     "ppl": ((1, 1, 1), ""),
     "sec": ((340, 330, 310), "handoffの例（340秒・330秒・310秒）。工程合計秒を実測値に置換"),
     "bind": (("=IF(C11=\"\",\"\",C11*C10)", "=IF(D11=\"\",\"\",D11*D10)", None),
              "PSは自動縫製中は拘束なし：セット＋次製品準備＋取り出しの実測秒（『パターンシーマ』運用設計）")},
    [200, 400, 600],
    ["設備費不要。段取り替え不要", "低コストで工程を集約できる", "工程集約。自動縫製中に次作業が可能。属人性排除・習熟期間短縮・品質安定"],
    ["熟練依存で品質がばらつく。習熟期間が長い", "依然として熟練依存。拘束時間の削減は限定的", "設備費・型製作が必要。サイクルタイムは手作業と同等の場合もある"],
    ["※1) 工程作業時間・拘束時間は同一工程・同一品番で3方式を実測した値を使用する。",
     "※2) 日産能力＝1日有効作業秒（稼働時間×3,600×(1－余裕率)）÷工程作業時間。",
     "※3) 1枚あたり加工コスト＝(減価償却費×設備台数＋人件費)÷(日産×月間稼働日×稼働率)。",
     "※4) 損益分岐点月数＝－(現行装置総額－新装置総額)÷{稼働率×(現行人件費－新人件費)}。",
     "※5) 省人化は人員削減ではなく、余力人員を他工程へ再配置して工場全体の生産性を上げることを指す。"])

build_compare(
    "比較_温冷プレス", "温冷プレス 60×12cm 比較（定量モデル）",
    "シームテープ接着後プレス工程。現行30×10cm（26台）／同型追加購入／60×12cm大型機の比較。",
    ["30×10cm 現行機（現行）", "30×10cm 追加購入（代替案）", "H&H CSL-201 60×12cm（新設備案）"],
    {"name": (("H&H 30×10cm（現有26台）", "H&H 30×10cm 追加", "H&H CSL-201（60×12cm）"), ""),
     "price": ((None, None, 4150), "現行は取得価額（償却済なら0）、代替案は30×10cmの見積単価"),
     "proc": (("シームテープ接着後プレス",) * 3, ""),
     "steps": ((None, '=IF(C9="","",C9)', None), "1箇所あたりのプレス回数。面積が狭いと回数増・シワリスク"),
     "ppl": ((1, 1, 1), ""),
     "sec": ((None, "=C11", "=IF(C11=\"\",\"\",C11/2)"), "新設備案は理由書の見込み（工程秒数約1/2）。実測値で置換"),
     "bind": (("=IF(C11=\"\",\"\",C11*C10)", "=IF(D11=\"\",\"\",D11*D10)", "=IF(E11=\"\",\"\",E11*E10)"), "")},
    [200, 400, 600],
    ["既存機を活用でき追加費用なし", "追加費用が小さい", "1回で広範囲をプレス。工程秒数約1/2見込み。仕上がり均一"],
    ["プレス面積が狭く複数回プレスが必要、シワリスク。台数不足", "回数・シワリスクは解消しない", "設備費"],
    ["※) 計算式・前提は『比較_パターンシーマ』の注記と同じ。現行の工程作業時間（秒/枚）を実測して記入すると全体が計算される。"],
    labels={"steps": "プレス回数(回/箇所)"})

build_compare(
    "比較_シーム機", "テーピングマシン spec-D 比較（定量モデル）",
    "袖・裾等の長尺パーツのシーム工程。現行は熟練者または2名体制。アタッチメント取付による対応は不可を確認済。",
    ["既存シーム機（現行）", "既存機＋アタッチメント（代替案）", "H&H AI-002 spec-D（新設備案）"],
    {"name": (("H&H 既存機（2名体制）", "対応不可（H&H・工場技術で確認済）", "H&H AI-002 spec-D"), ""),
     "price": ((None, None, 9700), "現行は取得価額（償却済なら0）"),
     "proc": (("長尺パーツのシーム", None, "長尺パーツのシーム"), ""),
     "steps": ((1, None, 1), ""),
     "ppl": ((2, None, 1), "現行は2名体制（予算案ファイル R10）"),
     "sec": ((None, None, None), "同一品番・同一パーツで実測"),
     "bind": (("=IF(C11=\"\",\"\",C11*C10)", None, "=IF(E11=\"\",\"\",E11*E10)"), "＝工程作業時間×作業人数")},
    [200, 400, 600],
    ["既存機を活用", "－", "ハンドリング性・視認性向上。1名で作業可能。熟練者依存の削減"],
    ["熟練者限定または2名体制。テープ貼り精度のばらつき", "取付不可のため選択肢にならない", "設備費"],
    ["※) 計算式・前提は『比較_パターンシーマ』の注記と同じ。"],
    na_cols=("D",))

# ---- レンタル vs 購入
RB = wb.create_sheet("比較_レンタル購入")
title(RB, "レンタル継続 vs 購入 比較", "瞬間的にしか使わない機材はレンタルの方が有利な場合もあるため、機種別に回収月数で判断する。")
legend(RB, "A3")
put(RB, "A4", "購入判断の基準（回収月数）"); inp(RB, "B4", 36, fmt="0", yellow=True); put(RB, "C4", "か月以内なら購入有利（基準値は要合意）")
header(RB, 6, ["機種", "購入単価(USD)", "台数", "購入額(USD)", "月額レンタル(USD/台)", "年間レンタル月数", "年間レンタル試算(USD)",
               "年間レンタル実績(USD)", "採用年間レンタル費(USD)", "回収月数(月)", "判定", "出典・備考"])
rb = [
    ("No.6 本縫針送りミシン DUMA", 880, 10, 112, 12, None, "予算案ファイル R9（112USD×12か月×10台）。2025年レンタル総額4,890.75USDを上回るため請求書で要確認"),
    ("No.1 オーバーロック", 980, 5, None, 3, None, "2024年10～12月に約10台を外部レンタル。月額単価を請求書で確認"),
    ("No.2 インターロック", 1150, 2, None, None, None, "レンタル実績の有無を確認"),
    ("No.5 大型パターンシーマ", 6500, 6, None, None, 3016.54, "2025年レンタル実績3,016.54USD（『レンタル』シート）"),
]
for i, (n, pr, q, mo, mm, act, src) in enumerate(rb):
    r = 7 + i
    put(RB, f"A{r}", n); inp(RB, f"B{r}", pr, fmt=USD); inp(RB, f"C{r}", q, fmt="0")
    put(RB, f"D{r}", f"=B{r}*C{r}", fmt=USD)
    inp(RB, f"E{r}", mo, fmt=USD, yellow=(mo is None)); inp(RB, f"F{r}", mm, fmt="0", yellow=(mm is None))
    put(RB, f"G{r}", f'=IF(OR(E{r}="",F{r}=""),"",E{r}*F{r}*C{r})', fmt=USD)
    inp(RB, f"H{r}", act, fmt=USD2, yellow=(act is None))
    put(RB, f"I{r}", f'=IF(H{r}<>"",H{r},G{r})', fmt=USD)
    put(RB, f"J{r}", f'=IF(N(I{r})=0,"",D{r}/(I{r}/12))', fmt=NUM1)
    put(RB, f"K{r}", f'=IF(J{r}="","データ待ち",IF(J{r}<=$B$4,"購入有利","レンタル継続を検討"))', align="center")
    put(RB, f"L{r}", src)
    RB.row_dimensions[r].height = 36
put(RB, "A12", "注：年間レンタル実績が入力されていれば実績を優先し、なければ月額×月数×台数の試算を使う。回収月数＝購入額÷(年間レンタル費÷12)。", size=9, border=False, wrap=False)
put(RB, "A13", "注：購入にはサーボ化による電力削減・品質改善（ロックミシン）などの副次効果もあるが、本表はレンタル費のみで比較。", size=9, border=False, wrap=False)
widths(RB, {"A": 30, "B": 13, "C": 7, "D": 13, "E": 14, "F": 12, "G": 15, "H": 15, "I": 15, "J": 12, "K": 18, "L": 60})

# ---------------------------------------------------------------- P/L影響
PL = wb.create_sheet("PL影響")
title(PL, "当年度P/L影響（シナリオ別）", "P/L影響＝新規設備年間償却－既存設備償却減少額－年間削減効果。プラス＝コスト増（安全・戦略理由で説明）、マイナス＝P/L改善。")
legend(PL, "A3")
header(PL, 5, ["項目", "必須案", "推奨案", "全体案", "備考"])
put(PL, "A6", "投資額(USD)")
for col, src in zip("BCD", ["C12", "C13", "C14"]):
    put(PL, f"{col}6", f"='総括'!{src}", fmt=USD, color=GREEN)
put(PL, "E6", "『総括』シナリオ集計")
put(PL, "A7", "投資額(円)")
for col in "BCD": put(PL, f"{col}7", f"={col}6*{RATE}", fmt=JPY)
put(PL, "A8", "新規設備 年間償却(円)")
for col in "BCD": put(PL, f"{col}8", f"={col}7/{YEARS}", fmt=JPY)
put(PL, "E8", "定額法・償却年数は『パラメータ』")
put(PL, "A9", "参考：APIC負担控除後 年間償却(円)")
for col in "BCD": put(PL, f"{col}9", f"=MAX({col}6-{APIC},0)*{RATE}/{YEARS}", fmt=JPY)
put(PL, "E9", "50,000USDをAPIC資産とする場合の参考値。扱い確定まで本計算には使わない")
put(PL, "A10", "既存設備 償却減少額(円)")
for col in "BCD": inp(PL, f"{col}10", None, fmt=JPY, yellow=True)
put(PL, "E10", "入替で除却する旧型の償却残（経理確認）。償却済なら0")
put(PL, "A11", "年間削減効果：レンタル回避(円)")
for col in "BCD": put(PL, f"{col}11", f"='レンタル'!$J$18*{RATE}", fmt=JPY, color=GREEN)
put(PL, "A12", "年間削減効果：電力(円)")
for col in "BCD": put(PL, f"{col}12", f"='ロックミシン'!$B$36*{RATE}", fmt=JPY, color=GREEN)
put(PL, "A13", "年間削減効果：省人・工数(円)")
put(PL, "B13", 0, fmt=JPY, color=BLUE); put(PL, "B13", 0, fmt=JPY, color=BLUE)
for col in "CD": put(PL, f"{col}13", f"=IF({LABOR}=\"\",0,'パターンシーマ'!$N$22*{LABOR}*12*{RATE})", fmt=JPY)
put(PL, "E13", "必須案はパターンシーマを含まないため0。人件費（月額）は『パラメータ』B15に記入。再配置人員の付加価値で評価する場合は別途")
put(PL, "A14", "年間削減効果：品質（不良削減）(円)")
for col in "BCD": inp(PL, f"{col}14", None, fmt=JPY, yellow=True)
put(PL, "E14", "不良・再加工の実績から算出")
put(PL, "A15", "年間削減効果 合計(円)", bold=True)
for col in "BCD": put(PL, f"{col}15", f"=SUM({col}11:{col}14)", fmt=JPY, bold=True)
put(PL, "A16", "当年度P/L影響(円)", bold=True)
for col in "BCD": put(PL, f"{col}16", f"={col}8-N({col}10)-{col}15", fmt=JPY, bold=True)
put(PL, "A17", "判定", bold=True)
for col in "BCD": put(PL, f"{col}17", f'=IF({col}16>0,"コスト増（安全・戦略理由で説明）","P/L改善")', align="center")
put(PL, "A18", "単純回収年数(年)")
for col in "BCD": put(PL, f"{col}18", f'=IF(N({col}11)+N({col}13)+N({col}14)=0,"効果未算定",{col}7/{col}15)', fmt="0.0")
put(PL, "E18", "＝投資額÷年間削減効果。レンタル・省人・品質効果が未記入の間は表示しない。安全投資（スポットプレス）は回収年数で評価しない")
for r in (11, 12, 15, 16, 17, 18): put(PL, f"E{r}", PL[f"E{r}"].value)
widths(PL, {"A": 36, "B": 22, "C": 22, "D": 22, "E": 70})

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print("saved", OUT)
