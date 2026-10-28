"""Tạo workbook 業務棚卸 của Chanh từ các file nguồn.

Chạy:
  python3 -I nghiep-vu/build_inventory.py --src <thư mục chứa file nguồn> --out <file .xlsx>

Thư mục nguồn cần các file (tên gốc khi upload, tiền tố bất kỳ):
  *Management_Workload_Delegation_Template_JP_NEW.xlsx   (業務棚卸し + báo cáo tuần)
  *2026*TIVN*.xlsx                                        (業務分析表 2026年度上期)
  *20231127_APIC_BTNB*.xlsx                                (業務分析表 2023/2024 + biên bản)
  *Chanh.xlsx                                              (業務分担表 có cột Chanh)
  *260522.xlsx                                             (移管資料 / 週次スケジュール)
  ef36455d-*.xlsx hoặc file 引継ぎ của Inoue               (bàn giao 2023)
File nguồn không được sửa; workbook mới chỉ đọc từ chúng.
"""
import argparse
import glob
import os
import re
import sys
import unicodedata
from collections import Counter, OrderedDict, defaultdict
from copy import copy

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from items import (ACT_TYPES, AMBIGUOUS, CATEGORIES, COVERED_BY, EXISTING, FIT, HANDOVER_MAP, NEW, PROPOSALS,  # noqa: E402
                   ROLE_TYPES, T47_NAMES, TASK_TO_ITEM, TOO_BROAD, TRANSFER_ITEMS)
from name_map import NAME_TO_TASK  # noqa: E402

# ---------------------------------------------------------------- styles
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_HYPER = Font(name=FONT, size=10, color="0563C1", underline="single")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_NOTE = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

STATUS = ["Confirmed", "To Verify", "Not Current"]


def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s)).upper()
    return s.replace(" ", "").replace("　", "").replace("DEVIDNOTE", "DEBITNOTE")


def find(src, pattern):
    hits = sorted(glob.glob(os.path.join(src, pattern)))
    if not hits:
        raise SystemExit(f"Không tìm thấy file nguồn: {pattern}")
    return hits[0]


def parse_hours(v):
    if v is None:
        return None, None
    if isinstance(v, (int, float)):
        return float(v), float(v)
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(v))]
    if not nums:
        return None, None
    return nums[0], nums[-1]


def week_no(w):
    return int(str(w)[-2:]) if w else None


# ---------------------------------------------------------------- helpers to write
def hdr(ws, row, headers, widths=None, height=30):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[CL(i)].width = w


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F_SUB


def put(ws, row, col, value, font=F_BASE, fill=None, fmt=None, align=WRAP, border=True):
    c = ws.cell(row, col, value)
    c.font = font
    c.alignment = align
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    return c


def section(ws, row, text, ncols=8):
    for col in range(1, ncols + 1):
        ws.cell(row, col).fill = FILL_SEC
    c = ws.cell(row, 1, text)
    c.font = F_BOLD


def status_cf(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="Confirmed"'],
                                                   fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="To Verify"'],
                                                   fill=PatternFill("solid", fgColor="FFEB9C")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="Not Current"'],
                                                   fill=PatternFill("solid", fgColor="D9D9D9")))


def dv_list(ws, rng, values):
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


# ---------------------------------------------------------------- load sources
def load_analysis(path, sheet):
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    hdr_row = next(r for r in range(1, 16)
                   if "Chanh" in [ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
    cols = {ws.cell(hdr_row, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(hdr_row, c).value}
    chanh = cols["Chanh"]
    total = next(c for c in range(1, ws.max_column + 1) if ws.cell(hdr_row - 1, c).value == "総計")
    jp_mgr = 7
    items, A, B = [], None, None
    for r in range(hdr_row + 2, ws.max_row + 1):
        A = ws.cell(r, 1).value or A
        B = ws.cell(r, 2).value or B
        n = ws.cell(r, 3).value
        if not isinstance(n, int):
            continue
        items.append(dict(no=n, row=r, A=A, B=B, B_raw=ws.cell(r, 2).value, A_raw=ws.cell(r, 1).value,
                          D=ws.cell(r, 4).value, E=ws.cell(r, 5).value, F=ws.cell(r, 6).value,
                          chanh=ws.cell(r, chanh).value or 0, total=ws.cell(r, total).value or 0,
                          jp=ws.cell(r, jp_mgr).value or 0,
                          chanh_ref=f"{CL(chanh)}{r}", total_ref=f"{CL(total)}{r}"))
    return ws, items, dict(hdr_row=hdr_row, chanh_col=chanh, total_col=total, jp_col=jp_mgr,
                           jp_name=ws.cell(hdr_row, jp_mgr).value)


def seq_match(base, other, kb, ko):
    out, j = {}, 0
    for b in base:
        for k in range(j, len(other)):
            if norm(ko(other[k])) == norm(kb(b)):
                out[id(b)] = other[k]
                j = k + 1
                break
    return out


def load_role_table(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = []
    for r in range(3, ws.max_row + 1):
        c = ws.cell(r, 3).value
        if c:
            rows.append(dict(row=r, A=ws.cell(r, 1).value, B=ws.cell(r, 2).value, C=str(c).strip(),
                             D=ws.cell(r, 4).value, L=ws.cell(r, 12).value, M=ws.cell(r, 13).value,
                             N=ws.cell(r, 14).value))
    return rows


def load_handover(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        last = None
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if not b or str(b).startswith("<"):
                continue
            mark = next((str(ws.cell(r, c).value) for c in range(3, ws.max_column + 1)
                         if str(ws.cell(r, c).value) in ("●", "▲")), "")
            note = next((str(ws.cell(r, c).value) for c in range(3, ws.max_column + 1)
                         if ws.cell(r, c).value and str(ws.cell(r, c).value) not in ("●", "▲")), "")
            key = (ws.title, r)
            if key in HANDOVER_MAP:
                no = HANDOVER_MAP[key]
                last = no
            else:
                no = last
            out.append(dict(sheet=ws.title, row=r, text=str(b).strip(), mark=mark, note=note, no=no))
    return out


def load_self_estimates(wb):
    """Tỷ lệ tải Chanh tự ước tính trong báo cáo tuần."""
    out = []
    for ws in wb.worksheets:
        if not re.match(r"\d{6}", ws.title):
            continue
        # dạng 1: bảng 総工数 (A nhãn, B tỷ lệ)
        rows = []
        for r in range(1, min(ws.max_row, 40) + 1):
            a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
            if isinstance(b, float) and 0 < b < 1 and a:
                rows.append((f"A{r}:B{r}", str(a), f"{b:.0%}"))
        if not rows:
            for r in range(1, min(ws.max_row, 70) + 1):
                a = ws.cell(r, 1).value
                if not a:
                    continue
                s = str(a)
                m = re.search(r"^(.+?)：\s*(約.+[%％])", s)
                if m:
                    rows.append((f"A{r}", m.group(1), m.group(2)))
                elif re.fullmatch(r"約\d+[%％]", s.strip()):
                    lab = next((str(ws.cell(k, 1).value) for k in range(r - 1, r - 4, -1)
                                if ws.cell(k, 1).value), "")
                    rows.append((f"A{r}", lab, s.strip()))
        for ref, lab, val in rows[:6]:
            out.append((ws.title, ref, lab, val))
    return out


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    f_log = find(args.src, "*Management_Workload_Delegation_Template*.xlsx")
    f_26 = find(args.src, "*2026*TIVN*.xlsx")
    f_2324 = find(args.src, "*20231127_APIC_BTNB*.xlsx")
    f_role = find(args.src, "*Chanh.xlsx")
    f_transfer = find(args.src, "*260522.xlsx")
    f_handover = find(args.src, "ef36455d-*.xlsx")

    # --- 業務棚卸し
    wb_log = openpyxl.load_workbook(f_log, data_only=True)
    ws_log = wb_log["業務棚卸し"]
    wb_log_f = openpyxl.load_workbook(f_log)
    n_formulas = sum(1 for row in wb_log_f["業務棚卸し"].iter_rows() for c in row
                     if isinstance(c.value, str) and c.value.startswith("="))
    LOG_LAST = ws_log.max_row
    log = []
    for r in range(2, LOG_LAST + 1):
        v = [ws_log.cell(r, c).value for c in range(1, 19)]
        if not v[4]:
            continue
        name = v[4]
        task = "T47" if name in T47_NAMES else NAME_TO_TASK[name]
        lo, hi = parse_hours(v[7])
        log.append(dict(row=r, week=v[0], wk=week_no(v[0]), name=name, cat=v[5], freq=v[6], h=v[7],
                        lo=lo, hi=hi, K=v[10], L=v[11], N=v[13], task=task,
                        no=TASK_TO_ITEM[task], vals=tuple(v[4:17])))
    weeks = sorted({x["wk"] for x in log if x["wk"]})

    # dòng trùng: dòng nền (không có tuần) giống hệt dòng tuần
    base_rows = {x["vals"]: x["row"] for x in log if not x["week"]}
    by_week = defaultdict(set)
    for x in log:
        if x["week"]:
            by_week[x["wk"]].add(x["vals"])
    for x in log:
        flags = []
        if x["name"] in TOO_BROAD:
            flags.append("Too broad")
        if x["week"] and x["vals"] in base_rows:
            flags.append(f"Exact dup of R{base_rows[x['vals']]}")
            x["dup_of"] = base_rows[x["vals"]]
        if x["week"] and x["vals"] in by_week.get(x["wk"] - 1, set()):
            flags.append("Same as prev week")
            x["carry"] = True
        if x["name"] in AMBIGUOUS:
            flags.append("Mapping?")
        x["flags"] = ", ".join(flags)

    # --- 業務分析表
    ws26, an26, meta26 = load_analysis(f_26, "<更新>APIC(2026年度上期) ")
    _, an24, _ = load_analysis(f_2324, "参考_APIC(2024年度)")
    _, an23, _ = load_analysis(f_2324, "参考_APIC(2023年度) ")
    m24 = seq_match(an26, an24, lambda x: x["D"], lambda x: x["D"])
    m23 = seq_match(an26, an23, lambda x: x["D"], lambda x: x["D"])
    for it in an26:
        it["c24"] = m24.get(id(it), {}).get("chanh")
        it["c23"] = m23.get(id(it), {}).get("chanh")
        it["no24"] = m24.get(id(it), {}).get("no")
    an26_by_no = {it["no"]: it for it in an26}
    chanh_items = [it["no"] for it in an26 if it["chanh"]]
    missing_existing = [k for k in EXISTING if k not in an26_by_no]
    assert not missing_existing, missing_existing
    for k in chanh_items:
        assert k in EXISTING, f"業務分析表 No.{k} có Chanh>0 nhưng chưa định nghĩa trong items.EXISTING"

    # --- 業務分担表 (Chanh)
    role = load_role_table(f_role)
    mr = seq_match(role, an26, lambda x: x["C"], lambda x: x["D"])
    ROLE_EXTRA = {3: 137, 5: 102, 6: 102, 7: 128, 8: 138, 12: 1, 32: 103, 33: 113, 39: 18}
    for x in role:
        x["no"] = ROLE_EXTRA.get(x["row"]) or (mr[id(x)]["no"] if id(x) in mr else None)

    handover = load_handover(f_handover)
    self_est = load_self_estimates(wb_log)

    # --- danh sách mục
    items = OrderedDict()
    for no in sorted(EXISTING):
        cat, vi, act, routine, rtype, fit, prop, dele, note = EXISTING[no]
        a = an26_by_no[no]
        items[no] = dict(no=no, kind="既存（業務分析表2026上期）", A=a["A"], B=a["B"], D=a["D"], E=a["E"],
                         F=a["F"], cat=cat, vi=vi, act=act, routine=routine, role=rtype, fit=fit, prop=prop,
                         dele=dele, note=note, c23=a["c23"], c24=a["c24"], c26=a["chanh"],
                         an_row=a["row"])
    for no in sorted(NEW):
        cat, jp, vi, act, routine, rtype, fit, prop, dele, srch, note = NEW[no]
        items[no] = dict(no=no, kind="新規（業務分析表に無し）", A="", B="", D=jp, E="", F="", cat=cat, vi=vi,
                         act=act, routine=routine, role=rtype, fit=fit, prop=prop, dele=dele,
                         note=note, src_hint=srch, c23=None, c24=None, c26=None, an_row=None)
    for x in log:
        assert x["no"] in items, x

    # bằng chứng khác cho từng mục
    other = defaultdict(list)
    for x in role:
        if x["N"] and x["no"]:
            other[x["no"]].append(f"業務分担表Chanh R{x['row']}{x['N']}")
    for h in handover:
        if h["no"] and h["mark"]:
            other[h["no"]].append(f"引継ぎ（Inoue）{h['sheet']} B{h['row']}{h['mark']}")
    for sh, ref, txt, no in TRANSFER_ITEMS:
        other[no].append(f"{sh} {ref}")
    BRIEF = {1: "Brief §6A", 18: "Brief §6C（約2日/週）", 4: "Brief §6C", 110: "Brief §6B", 128: "Brief §6B",
             8: "Brief §6D", 10: "Brief §6D", 11: "Brief §6D", 7: "Brief §6D", 118: "Brief §6D",
             107: "Brief §6E", 146: "Brief §6E", 127: "Brief §4", 142: "20231127議事録",
             141: "Repo: docs/quy-trinh-lead-time.md, data/nguon/*20260722*"}
    for k, v in BRIEF.items():
        other[k].append(v)
    for no, it in items.items():
        if it.get("src_hint") and it["src_hint"] != "Log":
            pass
        it["other"] = "; ".join(dict.fromkeys(other.get(no, [])))
        recs = [x for x in log if x["no"] == no]
        it["rows"] = ", ".join(f"R{x['row']}" for x in recs)
        it["names"] = sorted({x["name"] for x in recs})
        it["freq"] = Counter(x["freq"] for x in recs).most_common(1)[0][0] if recs else ""
        last = max(recs, key=lambda x: x["row"]) if recs else None
        it["self"] = (f"R{last['row']}: 関与={last['K'] or '-'} / 委譲先={last['L'] or '-'} / "
                      f"推奨={last['N'] or '-'}") if last else ""

    # =============================================================== workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = ["README", "業務一覧_Chanh", "Role_Analysis", "Workload_Analysis", "Missing_Tasks",
             "Duplicate_Review", "Action_Plan", "Advice", "Record_Map", "Audit",
             "Original_業務分析表", "Original_業務棚卸し", "Source_Evidence"]
    S = {n: wb.create_sheet(n) for n in names}
    ML, RM, WL, RA = "'業務一覧_Chanh'", "Record_Map", "Workload_Analysis", "Role_Analysis"

    # ------------------------------------------------ Original_業務棚卸し (copy)
    wo = S["Original_業務棚卸し"]
    for row in ws_log.iter_rows():
        for c in row:
            n = wo.cell(c.row, c.column, c.value)
            if c.has_style:
                n.font, n.fill, n.border = copy(c.font), copy(c.fill), copy(c.border)
                n.alignment, n.number_format = copy(c.alignment), c.number_format
    for k, d in ws_log.column_dimensions.items():
        wo.column_dimensions[k].width = d.width
    wo.freeze_panes = "F2"

    # ------------------------------------------------ Original_業務分析表 (extract)
    wa = S["Original_業務分析表"]
    keep = [1, 2, 3, 4, 5, 6, meta26["jp_col"], meta26["chanh_col"], meta26["total_col"],
            meta26["total_col"] + 2]
    for r in range(1, ws26.max_row + 1):
        for c in keep:
            src = ws26.cell(r, c)
            if src.value is None:
                continue
            n = wa.cell(r, c, src.value)
            n.font = copy(src.font) if src.has_style else F_BASE
            n.alignment = copy(src.alignment) if src.has_style else WRAP
            n.number_format = src.number_format
    for c in keep:
        wa.column_dimensions[CL(c)].width = 14
    wa.column_dimensions["D"].width = 22
    wa.column_dimensions["E"].width = 40
    wa.freeze_panes = "G14"
    note_col = meta26["total_col"] + 4
    wa.cell(1, note_col, "Trích từ sheet '<更新>APIC(2026年度上期) ': chỉ giữ cột A–F, "
            f"{meta26['jp_name']}, Chanh, 総計(工数), 該否 ở đúng vị trí cột/dòng gốc. "
            "Cột của staff khác và cột 金額 được bỏ (repo công khai).").font = F_SUB

    # ------------------------------------------------ Workload_Analysis (params first – referenced everywhere)
    wl = S["Workload_Analysis"]
    title(wl, "Workload_Analysis – Phân tích tải", "Ô vàng chữ xanh = tham số sửa được. Mọi số khác là công thức.")
    section(wl, 4, "1. Tham số", 16)
    params = [
        ("Giờ làm/ngày", 8, "Giả định – xác nhận lại"),
        ("Ngày làm/tuần", 6, "Báo cáo tuần ghi Thứ 2–Thứ 7 (vd 2026/10/19–10/24)"),
        ("Giờ khả dụng/tuần", "=C5*C6", "Công thức"),
        ("Tuần bắt đầu phân tích", 34, "W34–W43 dùng cùng mẫu báo cáo ổn định"),
        ("Tuần kết thúc phân tích", max(weeks), "Tuần mới nhất trong 業務棚卸し"),
        ("Số tuần phân tích", "=C9-C8+1", "Công thức"),
        ("Confirmed nếu xuất hiện trong N tuần gần nhất", 13, "≈ 3 tháng"),
        ("Tuần mốc Confirmed", "=C9-C11+1", "Công thức"),
        ("Plantable theo buổi với Hori (ngày/tuần)", 2, "Brief §6C"),
    ]
    for i, (k, v, n) in enumerate(params):
        r = 5 + i
        put(wl, r, 2, k)
        c = put(wl, r, 3, v, fmt="0.0" if i in (0, 1, 2) else "0")
        if not (isinstance(v, str) and v.startswith("=")):
            c.font, c.fill = F_INPUT, FILL_INPUT
        put(wl, r, 4, n, font=F_SUB)
    P_AVAIL, P_START, P_END, P_NW, P_THR, P_PT = ("Workload_Analysis!$C$7", "Workload_Analysis!$C$8",
                                                  "Workload_Analysis!$C$9", "Workload_Analysis!$C$10",
                                                  "Workload_Analysis!$C$12", "Workload_Analysis!$C$13")
    section(wl, 15, "2. Đề xuất → tỷ lệ giờ giải phóng (giả định, sửa được)", 16)
    hdr(wl, 16, ["", "Đề xuất", "Tỷ lệ giảm", "Ý nghĩa"])
    for i, (p, rate, d) in enumerate(PROPOSALS):
        r = 17 + i
        put(wl, r, 2, p)
        put(wl, r, 3, rate, font=F_INPUT, fill=FILL_INPUT, fmt="0%")
        put(wl, r, 4, d)
    PROP_RNG = f"Workload_Analysis!$B$17:$C${16 + len(PROPOSALS)}"
    CAT0 = 28  # category table first row
    W0 = CAT0 + 24
    CAT_RNG_NO = f"Workload_Analysis!$B${CAT0}:$B${CAT0 + 19}"
    CAT_RNG_NM = f"Workload_Analysis!$C${CAT0}:$C${CAT0 + 19}"

    # ------------------------------------------------ 業務一覧_Chanh
    ml = S["業務一覧_Chanh"]
    title(ml, "業務一覧（Chanh）– 業務分析表2026年度上期ベース＋新規業務（No.103〜）",
          "Mục 1–102: lấy từ 業務分析表 2026年度上期 (mục có FTE của Chanh > 0, cộng No.25). "
          "Mục 103+: việc thực tế ngoài 業務分析表 (từ 業務棚卸し và nguồn khác). Cột Y (vàng) để Chanh sửa trạng thái.")
    MH = ["No.", "区分", "分類\n(業務分析表)", "工程", "作業／業務名", "作業内容", "Mô tả (VI)", "頻度",
          "業務群No", "業務群 (20 nhóm)", "FTE 2023\nChanh", "FTE 2024\nChanh", "FTE 2026上期\nChanh",
          "FTE→h/tuần", "Log: số bản ghi", "Log: số tuần", "Log: tuần đầu", "Log: tuần cuối",
          "Log: giờ ghi\nTB h/tuần (kỳ PT)", "Log: giờ chuẩn hóa\nh/tuần (kỳ PT)", "h/tuần dùng\nphân tích",
          "Căn cứ giờ", "Bằng chứng khác", "Status (auto)", "Status\n(Chanh sửa)", "Status (final)",
          "Giải thích bằng chứng", "Dòng nguồn (Original_業務棚卸し)", "Ghi chú", "_sort", "Giờ đã nằm\ntrong No."]
    hdr(ml, 4, MH, [7, 14, 9, 11, 30, 30, 40, 11, 7, 18, 9, 9, 10, 9, 8, 7, 7, 7, 10, 11, 10, 8, 40,
                    11, 11, 11, 44, 40, 40, 6, 10], height=42)
    first = 5
    ML_ROW = {}
    for i, (no, it) in enumerate(items.items()):
        r = first + i
        ML_ROW[no] = r
        put(ml, r, 1, no, font=F_BOLD, align=CENTER)
        put(ml, r, 2, it["kind"])
        put(ml, r, 3, it["A"] if it["an_row"] else "")
        put(ml, r, 4, it["B"] if it["an_row"] else "")
        if it["an_row"]:
            c = put(ml, r, 5, f"='Original_業務分析表'!D{it['an_row']}", font=F_LINK)
            put(ml, r, 6, f"=IF('Original_業務分析表'!E{it['an_row']}=\"\",\"\",'Original_業務分析表'!E{it['an_row']})",
                font=F_LINK)
            put(ml, r, 8, f"=IF('Original_業務分析表'!F{it['an_row']}=\"\",\"\",'Original_業務分析表'!F{it['an_row']})",
                font=F_LINK)
            ch = CL(meta26["chanh_col"])
            put(ml, r, 13, f"=N('Original_業務分析表'!{ch}{it['an_row']})", font=F_LINK, fmt="0.00")
        else:
            put(ml, r, 5, it["D"])
            put(ml, r, 6, "")
            put(ml, r, 8, it["freq"])
            put(ml, r, 13, "", fmt="0.00")
        put(ml, r, 7, it["vi"])
        put(ml, r, 9, it["cat"], font=F_INPUT, align=CENTER)
        put(ml, r, 10, f"=IFERROR(INDEX({CAT_RNG_NM},MATCH(I{r},{CAT_RNG_NO},0)),\"\")")
        put(ml, r, 11, it["c23"] if it["c23"] is not None else "", fmt="0.00")
        put(ml, r, 12, it["c24"] if it["c24"] is not None else "", fmt="0.00")
        put(ml, r, 14, f"=IF(ISNUMBER(M{r}),M{r}*{P_AVAIL},0)", fmt="0.0")
        rng = f"{RM}!$K$4:$K${3 + len(log)}"
        put(ml, r, 15, f"=COUNTIFS({rng},A{r},{RM}!$P$4:$P${3 + len(log)},\"Y\")", fmt="0")
        put(ml, r, 16, f"=SUMIFS({RM}!$Q$4:$Q${3 + len(log)},{rng},A{r})", fmt="0")
        put(ml, r, 17, f"=IF(O{r}=0,\"\",_xlfn.MINIFS({RM}!$C$4:$C${3 + len(log)},{rng},A{r},{RM}!$P$4:$P${3 + len(log)},\"Y\"))", fmt="0")
        put(ml, r, 18, f"=IF(O{r}=0,\"\",_xlfn.MAXIFS({RM}!$C$4:$C${3 + len(log)},{rng},A{r},{RM}!$P$4:$P${3 + len(log)},\"Y\"))", fmt="0")
        put(ml, r, 19, f"=SUMIFS({RM}!$J$4:$J${3 + len(log)},{rng},A{r},{RM}!$T$4:$T${3 + len(log)},\"Y\")/{P_NW}", fmt="0.0")
        put(ml, r, 20, f"=SUMIFS({RM}!$S$4:$S${3 + len(log)},{rng},A{r},{RM}!$T$4:$T${3 + len(log)},\"Y\")/{P_NW}", fmt="0.0")
        put(ml, r, 21, f"=IF(T{r}>0,T{r},IF(AE{r}<>\"\",0,N{r}))", fmt="0.0")
        put(ml, r, 22, f"=IF(T{r}>0,\"Log\",IF(AE{r}<>\"\",\"→No.\"&AE{r},IF(N{r}>0,\"FTE\",\"—\")))", align=CENTER)
        put(ml, r, 31, COVERED_BY.get(no, ""), font=F_INPUT, align=CENTER)
        put(ml, r, 23, it["other"])
        put(ml, r, 24, f"=IF(O{r}>0,IF(R{r}>={P_THR},\"Confirmed\",\"To Verify\"),\"To Verify\")", align=CENTER)
        ov = ""
        if it["prop"] == "対象外":
            ov = "Not Current"
        elif no == 141:
            ov = "Confirmed"  # có commit trong repo (2026-10-23) và file rà soát 20260722
        put(ml, r, 25, ov, font=F_INPUT, fill=FILL_INPUT, align=CENTER)
        put(ml, r, 26, f"=IF(Y{r}<>\"\",Y{r},X{r})", align=CENTER)
        ev = []
        if it["an_row"]:
            ev.append(f"業務分析表2026上期 No.{no}: Chanh {it['c26']:.2f} FTE")
        if it["names"]:
            ev.append(f"業務棚卸し: {len(it['names'])} tên, xem cột AB")
        if not it["names"]:
            ev.append("Không có trong 業務棚卸し (log tuần) → cần Chanh xác nhận còn làm")
        if ov == "Not Current":
            ev.append("Không phải tải hiện tại (lịch sử/dự kiến)")
        elif ov:
            ev.append("Ghi đè Confirmed: repo có commit 2026-10-23 và file rà soát 20260722, nhưng chưa ghi trong 業務棚卸し")
        put(ml, r, 27, "; ".join(ev))
        put(ml, r, 28, it["rows"])
        put(ml, r, 29, it["note"])
        put(ml, r, 30, f"=IF(T{r}>0,T{r}+ROW()/1000000,0)", fmt="0.000000")
    ML_LAST = first + len(items) - 1
    tot = ML_LAST + 1
    put(ml, tot, 5, "Tổng", font=F_BOLD)
    for col in (13, 14, 15, 19, 20, 21):
        put(ml, tot, col, f"=SUM({CL(col)}{first}:{CL(col)}{ML_LAST})", font=F_BOLD,
            fmt="0.00" if col == 13 else "0.0")
    ml.freeze_panes = "F5"
    ml.auto_filter.ref = f"A4:{CL(len(MH))}{ML_LAST}"
    status_cf(ml, f"X{first}:Z{ML_LAST}", f"X{first}")
    dv_list(ml, f"Y{first}:Y{ML_LAST}", STATUS)
    ml.column_dimensions["AD"].hidden = True
    ml.cell(first, 25).comment = Comment("Chanh điền Confirmed / To Verify / Not Current để ghi đè trạng thái tự động.", "Claude")

    def MLC(col):
        return f"{ML}!${col}${first}:${col}${ML_LAST}"

    # ------------------------------------------------ Record_Map
    rm = S["Record_Map"]
    title(rm, "Record_Map – Mỗi bản ghi trong 業務棚卸し → No. mục",
          "Cột B–G, V–X lấy trực tiếp từ Original_業務棚卸し. Cột H, I, K (chữ xanh) là giá trị nhập: giờ tách từ chuỗi, No. gán. "
          "Sửa K để gán lại; mọi bảng tổng hợp tự cập nhật.")
    RH = ["Dòng gốc", "報告週", "Tuần", "業務名", "分類 (gốc)", "頻度 (gốc)", "想定時間 (gốc)", "h thấp", "h cao",
          "h giữa", "No.", "Tên mục", "業務群", "Loại bản ghi", "Cờ chất lượng", "Tính vào tải tuần",
          "Lần đầu trong tuần", "Tổng h giữa của tuần", "h chuẩn hóa", "Trong kỳ PT", "自分関与必要",
          "委譲先候補", "推奨アクション", "Ghi chú gán"]
    hdr(rm, 3, RH, [8, 10, 6, 34, 16, 10, 10, 7, 7, 7, 6, 30, 7, 10, 22, 8, 8, 9, 8, 7, 12, 16, 26, 36])
    n = len(log)
    for i, x in enumerate(log):
        r = 4 + i
        c = put(rm, r, 1, x["row"], font=F_HYPER, align=CENTER)
        c.hyperlink = f"#'Original_業務棚卸し'!E{x['row']}"
        o = f"'Original_業務棚卸し'!"
        put(rm, r, 2, f"=IF({o}A{x['row']}=\"\",\"\",{o}A{x['row']})", font=F_LINK)
        put(rm, r, 3, f"=IF(B{r}=\"\",\"\",VALUE(RIGHT(B{r},2)))", fmt="0")
        put(rm, r, 4, f"={o}E{x['row']}", font=F_LINK)
        put(rm, r, 5, f"={o}F{x['row']}", font=F_LINK)
        put(rm, r, 6, f"={o}G{x['row']}", font=F_LINK)
        put(rm, r, 7, f"={o}H{x['row']}", font=F_LINK)
        put(rm, r, 8, x["lo"], font=F_INPUT, fmt="0.0")
        put(rm, r, 9, x["hi"], font=F_INPUT, fmt="0.0")
        put(rm, r, 10, f"=(H{r}+I{r})/2", fmt="0.0")
        put(rm, r, 11, x["no"], font=F_INPUT, align=CENTER)
        put(rm, r, 12, f"=IFERROR(INDEX({MLC('E')},MATCH(K{r},{MLC('A')},0)),\"?\")")
        put(rm, r, 13, f"=IFERROR(INDEX({MLC('I')},MATCH(K{r},{MLC('A')},0)),\"?\")", align=CENTER)
        put(rm, r, 14, "Weekly" if x["week"] else "Foundation")
        put(rm, r, 15, x["flags"])
        put(rm, r, 16, "Y" if x["week"] else "N", align=CENTER)
        put(rm, r, 17, f"=IF(AND(P{r}=\"Y\",COUNTIFS($K$4:K{r},K{r},$C$4:C{r},C{r},$P$4:P{r},\"Y\")=1),1,0)")
        put(rm, r, 18, f"=IF(P{r}=\"Y\",SUMIFS($J$4:$J${3 + n},$C$4:$C${3 + n},C{r},$P$4:$P${3 + n},\"Y\"),\"\")", fmt="0.0")
        put(rm, r, 19, f"=IF(P{r}=\"Y\",J{r}/R{r}*{P_AVAIL},0)", fmt="0.00")
        put(rm, r, 20, f"=IF(AND(P{r}=\"Y\",C{r}>={P_START},C{r}<={P_END}),\"Y\",\"N\")", align=CENTER)
        put(rm, r, 21, f"=IF({o}K{x['row']}=\"\",\"\",{o}K{x['row']})", font=F_LINK)
        put(rm, r, 22, f"=IF({o}L{x['row']}=\"\",\"\",{o}L{x['row']})", font=F_LINK)
        put(rm, r, 23, f"=IF({o}N{x['row']}=\"\",\"\",{o}N{x['row']})", font=F_LINK)
        put(rm, r, 24, AMBIGUOUS.get(x["name"], ""))
    rm.freeze_panes = "E4"
    rm.auto_filter.ref = f"A3:X{3 + n}"
    RM_LAST = 3 + n

    # ------------------------------------------------ Workload_Analysis tables
    section(wl, CAT0 - 2, "3. Tải theo 20 nhóm nghiệp vụ (kỳ phân tích)", 16)
    hdr(wl, CAT0 - 1, ["", "No", "業務群", "Số mục", "Confirmed", "To Verify", "Not Current",
                       "Log: giờ ghi\nh/tuần", "Log: chuẩn hóa\nh/tuần", "Tỷ trọng", "Hạng",
                       "FTE 2026上期", "FTE→h/tuần", "Chênh lệch\n(log − FTE)", "Routine\nh/tuần",
                       "Giờ giải phóng\n(đề xuất)", "Log chuẩn hóa\ntoàn kỳ W19–W43"], [3, 6, 24, 8, 9, 9, 9, 11, 11, 9, 7, 10, 10, 11, 10, 12, 13],
        height=42)
    for i, (cno, cname) in enumerate(CATEGORIES):
        r = CAT0 + i
        put(wl, r, 2, cno, align=CENTER)
        put(wl, r, 3, cname)
        put(wl, r, 4, f"=COUNTIF({MLC('I')},B{r})", fmt="0")
        for j, st in enumerate(STATUS):
            put(wl, r, 5 + j, f"=COUNTIFS({MLC('I')},B{r},{MLC('Z')},\"{st}\")", fmt="0")
        put(wl, r, 8, f"=SUMIFS({MLC('S')},{MLC('I')},B{r})", fmt="0.0")
        put(wl, r, 9, f"=SUMIFS({MLC('T')},{MLC('I')},B{r})", fmt="0.0")
        put(wl, r, 10, f"=IF($I${CAT0 + 20}=0,0,I{r}/$I${CAT0 + 20})", fmt="0.0%")
        put(wl, r, 11, f"=RANK(I{r},$I${CAT0}:$I${CAT0 + 19})", fmt="0")
        put(wl, r, 12, f"=SUMIFS({MLC('M')},{MLC('I')},B{r})", fmt="0.00")
        put(wl, r, 13, f"=L{r}*{P_AVAIL}", fmt="0.0")
        put(wl, r, 14, f"=I{r}-M{r}", fmt="0.0;[Red]-0.0")
        put(wl, r, 15, f"=SUMIFS({RA}!$F$37:$F$200,{RA}!$C$37:$C$200,B{r},{RA}!$H$37:$H$200,\"Routine\")", fmt="0.0")
        put(wl, r, 16, f"=SUMIFS({RA}!$O$37:$O$200,{RA}!$C$37:$C$200,B{r})", fmt="0.0")
        put(wl, r, 17, f"=SUMIFS({RM}!$S$4:${'S'}${RM_LAST},{RM}!$M$4:$M${RM_LAST},B{r})/COUNT($B${W0 + 1}:$B${W0 + len(weeks)})", fmt="0.0")
    r = CAT0 + 20
    put(wl, r, 3, "Tổng", font=F_BOLD)
    for col in range(4, 18):
        if col in (10, 11):
            continue
        put(wl, r, col, f"=SUM({CL(col)}{CAT0}:{CL(col)}{CAT0 + 19})", font=F_BOLD,
            fmt="0.00" if col == 12 else ("0" if col <= 7 else "0.0"))
    put(wl, r, 10, f"=SUM(J{CAT0}:J{CAT0 + 19})", font=F_BOLD, fmt="0%")
    put(wl, CAT0 + 21, 3, "Ghi chú: 'giờ ghi' = giờ ước tính Chanh ghi (có trùng giữa các việc). 'Chuẩn hóa' = chia lại "
        "theo tỷ trọng sao cho tổng mỗi tuần = giờ khả dụng (C7). FTE = phân bổ trong 業務分析表 2026上期.",
        font=F_SUB, border=False)
    wl.merge_cells(start_row=CAT0 + 21, start_column=3, end_row=CAT0 + 21, end_column=16)
    wl.row_dimensions[CAT0 + 21].height = 30

    section(wl, W0 - 1, "4. Tổng giờ ghi theo tuần (cho thấy giờ ước tính bị cộng trùng / vượt giờ làm)", 16)
    hdr(wl, W0, ["", "Tuần", "Số bản ghi", "Tổng h thấp", "Tổng h cao", "Tổng h giữa", "Giờ khả dụng",
                 "Tỷ lệ h giữa / khả dụng"])
    for i, w in enumerate(weeks):
        r = W0 + 1 + i
        put(wl, r, 2, w, fmt="0", align=CENTER)
        put(wl, r, 3, f"=COUNTIFS({RM}!$C$4:$C${RM_LAST},B{r},{RM}!$P$4:$P${RM_LAST},\"Y\")", fmt="0")
        put(wl, r, 4, f"=SUMIFS({RM}!$H$4:$H${RM_LAST},{RM}!$C$4:$C${RM_LAST},B{r},{RM}!$P$4:$P${RM_LAST},\"Y\")", fmt="0.0")
        put(wl, r, 5, f"=SUMIFS({RM}!$I$4:$I${RM_LAST},{RM}!$C$4:$C${RM_LAST},B{r},{RM}!$P$4:$P${RM_LAST},\"Y\")", fmt="0.0")
        put(wl, r, 6, f"=(D{r}+E{r})/2", fmt="0.0")
        put(wl, r, 7, f"={P_AVAIL}", fmt="0.0")
        put(wl, r, 8, f"=F{r}/G{r}", fmt="0%")
    WL_W1, WL_W2 = W0 + 1, W0 + len(weeks)
    r = WL_W2 + 1
    put(wl, r, 2, "TB kỳ PT", font=F_BOLD)
    for col in (3, 4, 5, 6):
        put(wl, r, col, f"=AVERAGEIFS({CL(col)}{WL_W1}:{CL(col)}{WL_W2},$B${WL_W1}:$B${WL_W2},\">=\"&{P_START},"
            f"$B${WL_W1}:$B${WL_W2},\"<=\"&{P_END})", font=F_BOLD, fmt="0.0")
    put(wl, r, 7, f"={P_AVAIL}", font=F_BOLD, fmt="0.0")
    put(wl, r, 8, f"=F{r}/G{r}", font=F_BOLD, fmt="0%")
    WL_AVG = r
    put(wl, r + 1, 2, "TB toàn bộ", font=F_BOLD)
    for col in (3, 4, 5, 6):
        put(wl, r + 1, col, f"=AVERAGE({CL(col)}{WL_W1}:{CL(col)}{WL_W2})", font=F_BOLD, fmt="0.0")
    put(wl, r + 1, 7, f"={P_AVAIL}", font=F_BOLD, fmt="0.0")
    put(wl, r + 1, 8, f"=F{r + 1}/G{r + 1}", font=F_BOLD, fmt="0%")

    X0 = WL_AVG + 4
    section(wl, X0, "5. Đối chiếu Plantable / kế hoạch tuần (No.4, 18, 20, 88, 89, 90, 85, 143)", 16)
    pt_items = [4, 18, 20, 88, 89, 90]
    meet_items = [85, 143]

    def sum_items(col, nos):
        return "+".join(f"SUMIFS({MLC(col)},{MLC('A')},{k})" for k in nos)

    rows_x = [
        ("Log: giờ chuẩn hóa h/tuần – kế hoạch (4,18,20,88–90)", f"={sum_items('T', pt_items)}"),
        ("Log: giờ chuẩn hóa h/tuần – họp kế hoạch (85,143)", f"={sum_items('T', meet_items)}"),
        ("Log: giờ ghi (chưa chuẩn hóa) h/tuần – kế hoạch + họp", f"={sum_items('S', pt_items + meet_items)}"),
        ("Buổi với Hori: ~2 ngày/tuần", f"={P_PT}*Workload_Analysis!$C$5"),
        ("Dòng nền 業務棚卸し R5 (週次生産計画会議・進捗調整・日本報告)", "=N('Original_業務棚卸し'!H5)"),
        ("Dòng nền 業務棚卸し R15 (生産計画調整・チュエン管理) – h giữa",
         f"=SUMIFS({RM}!$J$4:$J${RM_LAST},{RM}!$A$4:$A${RM_LAST},15)"),
        ("業務分析表 2026上期: FTE kế hoạch + họp → h/tuần", f"={sum_items('N', pt_items + meet_items)}"),
    ]
    for i, (k, f) in enumerate(rows_x):
        put(wl, X0 + 1 + i, 2, k)
        wl.merge_cells(start_row=X0 + 1 + i, start_column=2, end_row=X0 + 1 + i, end_column=6)
        put(wl, X0 + 1 + i, 7, f, fmt="0.0")
        put(wl, X0 + 1 + i, 8, f"=G{X0 + 1 + i}/{P_AVAIL}", fmt="0%")
    put(wl, X0 + len(rows_x) + 1, 2, "Tỷ lệ Chanh tự ước tính trong báo cáo tuần: xem Source_Evidence mục F "
        "(W40: 'Plantableに50%以上の時間を使う状態から脱却' ).", font=F_SUB, border=False)

    Y0 = X0 + len(rows_x) + 4
    section(wl, Y0, "6. Routine / Non-routine / Project / Seasonal (giờ chuẩn hóa h/tuần)", 16)
    for i, k in enumerate(["Routine", "Non-routine", "Project", "Seasonal"]):
        r = Y0 + 1 + i
        put(wl, r, 2, k)
        put(wl, r, 3, f"=SUMIFS({RA}!$F$37:$F$200,{RA}!$H$37:$H$200,B{r})", fmt="0.0")
        put(wl, r, 4, f"=IF(SUM($C${Y0 + 1}:$C${Y0 + 4})=0,0,C{r}/SUM($C${Y0 + 1}:$C${Y0 + 4}))", fmt="0%")

    Z0 = Y0 + 7
    section(wl, Z0, "7. 10 mục tốn nhiều giờ nhất (log chuẩn hóa, kỳ PT)", 16)
    hdr(wl, Z0 + 1, ["", "Hạng", "No.", "Mục", "h/tuần", "Tỷ trọng", "業務群", "Status"])
    for k in range(1, 11):
        r = Z0 + 1 + k
        put(wl, r, 2, k, align=CENTER)
        put(wl, r, 3, f"=INDEX({MLC('A')},MATCH(LARGE({MLC('AD')},B{r}),{MLC('AD')},0))")
        put(wl, r, 4, f"=INDEX({MLC('E')},MATCH(C{r},{MLC('A')},0))")
        put(wl, r, 5, f"=INDEX({MLC('T')},MATCH(C{r},{MLC('A')},0))", fmt="0.0")
        put(wl, r, 6, f"=E{r}/{P_AVAIL}", fmt="0%")
        put(wl, r, 7, f"=INDEX({MLC('J')},MATCH(C{r},{MLC('A')},0))")
        put(wl, r, 8, f"=INDEX({MLC('Z')},MATCH(C{r},{MLC('A')},0))")
    wl.freeze_panes = "A4"

    # ------------------------------------------------ Role_Analysis
    ra = S["Role_Analysis"]
    title(ra, "Role_Analysis – Phạm vi theo chức vụ (課長) & đề xuất giảm tải / giao việc",
          "DỰ THẢO – chưa phải quyết định. Cột I–L (chữ xanh) là đề xuất của Claude, Chanh/Hori sửa và duyệt. "
          "Giờ lấy từ 業務一覧_Chanh cột U.")
    # summaries
    section(ra, 4, "A. Theo mức phù hợp chức vụ", 17)
    hdr(ra, 5, ["", "Mức", "Số mục", "h/tuần hiện tại", "Tỷ trọng", "h giải phóng", "h sau đề xuất"])
    for i, (k, v) in enumerate(FIT.items()):
        r = 6 + i
        put(ra, r, 2, v)
        put(ra, r, 3, f"=COUNTIFS($J$37:$J$200,\"{k}*\",$D$37:$D$200,\"<>Not Current\")", fmt="0")
        put(ra, r, 4, f"=SUMIFS($F$37:$F$200,$J$37:$J$200,\"{k}*\")", fmt="0.0")
        put(ra, r, 5, f"=IF($D$11=0,0,D{r}/$D$11)", fmt="0%")
        put(ra, r, 6, f"=SUMIFS($O$37:$O$200,$J$37:$J$200,\"{k}*\")", fmt="0.0")
        put(ra, r, 7, f"=D{r}-F{r}", fmt="0.0")
    put(ra, 11, 2, "Tổng", font=F_BOLD)
    for col in (3, 4, 6, 7):
        put(ra, 11, col, f"=SUM({CL(col)}6:{CL(col)}10)", font=F_BOLD, fmt="0" if col == 3 else "0.0")
    put(ra, 11, 5, "=SUM(E6:E10)", font=F_BOLD, fmt="0%")

    section(ra, 13, "B. Theo vai trò (Operation → System Building)", 17)
    hdr(ra, 14, ["", "Vai trò", "Số mục", "h/tuần hiện tại", "Tỷ trọng", "h giải phóng", "h sau đề xuất", "Tỷ trọng sau"])
    for i, k in enumerate(ROLE_TYPES):
        r = 15 + i
        put(ra, r, 2, k)
        put(ra, r, 3, f"=COUNTIFS($I$37:$I$200,B{r},$D$37:$D$200,\"<>Not Current\")", fmt="0")
        put(ra, r, 4, f"=SUMIFS($F$37:$F$200,$I$37:$I$200,B{r})", fmt="0.0")
        put(ra, r, 5, f"=IF($D$20=0,0,D{r}/$D$20)", fmt="0%")
        put(ra, r, 6, f"=SUMIFS($O$37:$O$200,$I$37:$I$200,B{r})", fmt="0.0")
        put(ra, r, 7, f"=D{r}-F{r}", fmt="0.0")
        put(ra, r, 8, f"=IF($G$20=0,0,G{r}/$G$20)", fmt="0%")
    put(ra, 20, 2, "Tổng", font=F_BOLD)
    for col in (3, 4, 6, 7):
        put(ra, 20, col, f"=SUM({CL(col)}15:{CL(col)}19)", font=F_BOLD, fmt="0" if col == 3 else "0.0")
    put(ra, 20, 5, "=SUM(E15:E19)", font=F_BOLD, fmt="0%")
    put(ra, 20, 8, "=SUM(H15:H19)", font=F_BOLD, fmt="0%")

    section(ra, 22, "C. Tổng hợp tải trước / sau đề xuất", 17)
    rows_c = [
        ("Giờ khả dụng/tuần", f"={P_AVAIL}"),
        ("Giờ Chanh ghi TB/tuần trong 業務棚卸し (chưa chuẩn hóa, kỳ PT)", f"=Workload_Analysis!F{WL_AVG}"),
        ("Giờ của mục chưa ghi trong log (ước theo FTE 業務分析表)", f"=SUMIFS($F$37:$F$200,$G$37:$G$200,\"FTE\")"),
        ("Tải theo giờ ghi = (2) + (3)", "=C25+C26"),
        ("Tỷ lệ tải hiện tại / khả dụng", "=C27/C24"),
        ("Tỷ lệ giờ giải phóng nếu áp dụng đề xuất", "=IF(D11=0,0,F11/D11)"),
        ("Tải sau đề xuất h/tuần (theo giờ ghi)", "=C27*(1-C29)"),
        ("Tỷ lệ tải sau đề xuất / khả dụng", "=C30/C24"),
        ("Tỷ trọng Decision + System Building: hiện tại → sau đề xuất", "=TEXT(SUM(E18:E19),\"0%\")&\" → \"&TEXT(SUM(H18:H19),\"0%\")"),
    ]
    for i, (k, f) in enumerate(rows_c):
        r = 24 + i
        put(ra, r, 2, k)
        put(ra, r, 3, f, fmt="0%" if "Tỷ lệ" in k else "0.0")
        put(ra, r, 1, f"({i + 1})", align=CENTER)
    put(ra, 33, 2, "Lưu ý: giờ ghi là ước tính, có thể trùng giữa các việc hoặc phản ánh làm thêm giờ. Bảng A/B dùng giờ chuẩn hóa "
        "(tổng = giờ khả dụng) để so tỷ trọng. Cần đo thực tế (Action_Plan A05).", font=F_SUB, border=False)
    ra.merge_cells("B33:L33")
    ra.row_dimensions[33].height = 30

    RAH = ["No.", "Mục", "業務群", "Status", "Kiểu công việc", "h/tuần", "Căn cứ giờ", "Routine",
           "Vai trò (dự thảo)", "Phù hợp chức vụ", "Đề xuất", "Giao cho (ứng viên)", "Tỷ lệ giảm",
           "Lý do / cách làm", "h giải phóng", "h còn lại", "Tự đánh giá của Chanh trong log (bản ghi mới nhất)"]
    hdr(ra, 36, RAH, [7, 32, 7, 11, 16, 8, 8, 11, 15, 30, 12, 22, 8, 46, 9, 9, 50], height=36)
    RA0 = 37
    for i, (no, it) in enumerate(items.items()):
        r = RA0 + i
        put(ra, r, 1, no, font=F_BOLD, align=CENTER)
        put(ra, r, 2, f"=INDEX({MLC('E')},MATCH(A{r},{MLC('A')},0))")
        put(ra, r, 3, f"=INDEX({MLC('I')},MATCH(A{r},{MLC('A')},0))", align=CENTER)
        put(ra, r, 4, f"=INDEX({MLC('Z')},MATCH(A{r},{MLC('A')},0))", align=CENTER)
        put(ra, r, 5, it["act"])
        put(ra, r, 6, f"=IF(D{r}=\"Not Current\",0,INDEX({MLC('U')},MATCH(A{r},{MLC('A')},0)))", fmt="0.0")
        put(ra, r, 7, f"=INDEX({MLC('V')},MATCH(A{r},{MLC('A')},0))", align=CENTER)
        put(ra, r, 8, it["routine"])
        put(ra, r, 9, it["role"], font=F_INPUT)
        put(ra, r, 10, FIT[it["fit"]], font=F_INPUT)
        put(ra, r, 11, it["prop"], font=F_INPUT)
        put(ra, r, 12, it["dele"], font=F_INPUT)
        put(ra, r, 13, f"=IFERROR(VLOOKUP(K{r},{PROP_RNG},2,FALSE),0)", fmt="0%")
        put(ra, r, 14, it["note"])
        put(ra, r, 15, f"=F{r}*M{r}", fmt="0.0")
        put(ra, r, 16, f"=F{r}-O{r}", fmt="0.0")
        put(ra, r, 17, it["self"])
    RA_LAST = RA0 + len(items) - 1
    status_cf(ra, f"D{RA0}:D{RA_LAST}", f"D{RA0}")
    dv_list(ra, f"I{RA0}:I{RA_LAST}", ROLE_TYPES)
    dv_list(ra, f"K{RA0}:K{RA_LAST}", [p[0] for p in PROPOSALS])
    dv_list(ra, f"J{RA0}:J{RA_LAST}", list(FIT.values()))
    ra.freeze_panes = f"C{RA0}"
    ra.auto_filter.ref = f"A36:Q{RA_LAST}"
    assert RA_LAST <= 200

    # ------------------------------------------------ Missing_Tasks
    mt = S["Missing_Tasks"]
    title(mt, "Missing_Tasks – Việc thiếu / chưa rõ (Gap Analysis)",
          "Existing 業務 = mục đã có. Proposed Action: Add / Split / Clarify / Merge. "
          "Mục mới đã được thêm vào 業務一覧_Chanh với No. 103+.")
    MTH = ["Gap ID", "業務群", "Existing 業務", "Missing / Unclear Activity", "Evidence", "Proposed Action",
           "No. trong 業務一覧", "Status", "Ưu tiên", "Ghi chú"]
    hdr(mt, 4, MTH, [8, 7, 30, 46, 46, 11, 9, 11, 8, 40])
    gaps = []
    # (1) mục mới từ log
    for no, it in items.items():
        if no < 103:
            continue
        if it["names"]:
            gaps.append((it["cat"], "— (không có trong 業務分析表)",
                         f"No.{no} {it['D']}: {it['vi']}",
                         f"業務棚卸し {len(it['names'])} tên: " + " / ".join(it["names"][:4]) +
                         (" …" if len(it["names"]) > 4 else ""), "Add", no,
                         "High" if it["fit"] in ("A",) else "Medium", it["note"]))
        else:
            gaps.append((it["cat"], "— (không có trong 業務分析表 và 業務棚卸し)",
                         f"No.{no} {it['D']}: {it['vi']}", it["other"] or it.get("src_hint", ""),
                         "Add" if it["prop"] != "対象外" else "Clarify", no,
                         "Medium", it["note"]))
    # (2) mục có trong 業務分析表 nhưng không có trong log
    for no, it in items.items():
        if no >= 103 or it["names"]:
            continue
        cov = COVERED_BY.get(no)
        if cov:
            gaps.append((it["cat"], f"No.{no} {it['D']}",
                         f"Không ghi riêng trong 業務棚卸し; giờ đang tính trong No.{cov} → xác nhận còn làm và có cần ghi riêng",
                         f"業務分析表2026上期 Chanh {it['c26']:.2f}; " + it["other"], "Merge", no, "Low", it["note"]))
        else:
            gaps.append((it["cat"], f"No.{no} {it['D']}",
                         "Có FTE trong 業務分析表 nhưng không thấy trong 業務棚卸し (log tuần W19–W43) → giờ chưa được ghi",
                         f"業務分析表2026上期 Chanh {it['c26']:.2f}; " + it["other"], "Clarify", no,
                         "High" if no in (102, 8, 9) else "Medium", it["note"]))
    # (3) tách / làm rõ cụ thể
    manual = [
        ("06", 1, "Debit Note: chưa ghi các bước kiểm tra chi tiết – giờ công chuẩn vs thực tế; kỹ thuật viên, phí mẫu, thuê máy, vận chuyển nội địa; nghỉ việc/vắng/quay lại → điều chỉnh tiền; luồng JP→TIVN→APIC→TIVN→JP",
         "Brief §6A; log W41–W43 (QC14名, D/N請求人員)", "Split", "High", "Làm checklist để người chuẩn bị kiểm, Chanh duyệt."),
        ("01", 18, "Plantable: tách 4 bước theo 移管資料 – sắp xếp đơn hàng; đăng ký PT theo 7 quy tắc chọn chuyền; xem lại kế hoạch (NPL, yêu cầu KD, chuyền trống tạm); phát hành (Thứ 4 → nhà máy, Thứ 2 → KD)",
         "移管資料 H4–H28; 週次スケジュール O–R", "Split", "High", "Giúp xác định bước nào giao Planning Leader."),
        ("01", 18, "Giờ Plantable không thống nhất: dòng nền R5 16h/tuần, R15 15–20h, Hori ~2 ngày/tuần, log W34+ 6–8h",
         "Original_業務棚卸し R5, R15; Brief §6C; log", "Clarify", "High", "Đo thực tế 2 tuần."),
        ("16", 85, "Tên '生産進捗会議…' không rõ là họp với nhà máy (85) hay với kinh doanh 1課・2課 (143)",
         "Record_Map cột X", "Clarify", "Medium", ""),
        ("12", 77, "'TNF・Goldwin DN発行対応': DN = Debit Note hay Delivery Note?",
         "Original_業務棚卸し R143", "Clarify", "Low", ""),
        ("05", 110, "Manpower: tách – xác nhận nhân sự cuối tháng; điều chỉnh sau Tết/nghỉ dài; đàm phán bổ sung người; nâng ý thức quản lý của technical staff",
         "Brief §6B; log W21, W36, W38, W43", "Split", "High", ""),
        ("14", 8, "Audit: thêm theo dõi cải thiện sau audit (CAPA) và quản lý chứng nhận",
         "Brief §6D; 引継ぎ オーディット B2–B4", "Add", "Medium", ""),
        ("14", 10, "Khách: guesthouse, xe, lịch trình, welcome board… (danh sách bàn giao của Inoue)",
         "Brief §6D; 引継ぎ 来客 B6–B24", "Clarify", "Medium", "Đề xuất giao GA/Admin."),
        ("13", 7, "Chi phí gia công ngoài, đặc biệt data thêu", "Brief §6D; log W38–W40 (刺繍外注)", "Clarify", "Medium", ""),
        ("03", 107, "Chậm vận chuyển NPL, kiểm tra lead time logistics", "Brief §6E; Repo lead time 20260722", "Clarify", "Medium", ""),
        ("15", 121, "Email: dòng nền 5–7h/tuần nhưng từ W34 không còn dòng email riêng → có thể đang nằm trong No.122",
         "Original_業務棚卸し R22, R32; log W22–W33", "Clarify", "High", "Ghi riêng email trong log tuần."),
        ("20", 64, "Sự cố đột xuất ghi chung chung (5–10h/tuần) – không có danh sách case",
         "Original_業務棚卸し R23; log", "Clarify", "Medium", "Ghi số case R1/R2/R3 mỗi tuần."),
        ("18", 128, "Việc đã giao (Delivery Schedule → Cường W29; Sample → W33) cần ghi rõ để trừ khỏi tải của Chanh",
         "log W29–W36", "Clarify", "Medium", ""),
        ("17", 125, "Chanh review báo cáo tuần của Hori/Naito: vai trò là gì?", "log W30–W34", "Clarify", "Low", ""),
        ("18", 129, "Bàn giao Kushida: ghi lần cuối W29 → đã xong?", "log W21–W29", "Clarify", "Low", ""),
        ("18", 102, "Staff TIVN: duyệt time card, đơn nghỉ, ra ngoài; điều chỉnh phân công khi nghỉ dài",
         "引継ぎ スタッフ管理 B2–B8 ●; 業務分析表 No.102", "Add", "High", "Ghi vào log tuần."),
        ("07", 111, "期中変更対応 (No.3) và đàm phán hợp đồng (No.111) chồng nhau", "業務分析表 No.3; log", "Merge", "Low", ""),
        ("19", None, "Dữ liệu log: 分類 có 60 giá trị, 頻度 19 giá trị, 想定時間 ghi dạng chuỗi '2～3' / '5～8h'",
         "Audit", "Clarify", "High", "Thêm cột 業務No. (theo file này) vào mẫu 業務棚卸し."),
    ]
    for cat, no, txt, evd, act, pri, note in manual:
        ex = f"No.{no} {items[no]['D']}" if no else "—"
        gaps.append((cat, ex, txt, evd, act, no if no else "", pri, note))
    # (4) không còn hiện tại
    for h in handover:
        if h["sheet"] == "報告" and h["row"] in (4, 5, 13):
            gaps.append(("17", "—", f"{h['text']} (bàn giao 2023)", f"引継ぎ 報告 B{h['row']}{h['mark']}",
                         "Clarify", "", "Low", "Not Current – việc thời COVID."))
    for i, g in enumerate(gaps):
        r = 5 + i
        cat, ex, txt, evd, act, no, pri, note = g
        put(mt, r, 1, f"G{i + 1:02d}", align=CENTER)
        put(mt, r, 2, cat, align=CENTER)
        put(mt, r, 3, ex)
        put(mt, r, 4, txt)
        put(mt, r, 5, evd)
        put(mt, r, 6, act, align=CENTER)
        put(mt, r, 7, no, align=CENTER)
        if no != "":
            put(mt, r, 8, f"=IFERROR(INDEX({MLC('Z')},MATCH(G{r},{MLC('A')},0)),\"\")", align=CENTER)
        else:
            put(mt, r, 8, "Not Current" if "Not Current" in note else "To Verify", align=CENTER)
        put(mt, r, 9, pri, align=CENTER)
        put(mt, r, 10, note)
    MT_LAST = 4 + len(gaps)
    status_cf(mt, f"H5:H{MT_LAST}", "H5")
    mt.freeze_panes = "D5"
    mt.auto_filter.ref = f"A4:J{MT_LAST}"

    # ------------------------------------------------ Duplicate_Review
    dr = S["Duplicate_Review"]
    title(dr, "Duplicate_Review – Bản ghi trùng / chồng lấn",
          "Giờ chuẩn hóa (Record_Map cột S) đã chia lại theo tuần nên không cộng trùng; sheet này chỉ ra chỗ trùng để Chanh sửa log.")
    section(dr, 4, "A. Trùng hoàn toàn: dòng tuần W19 giống hệt dòng nền (không tuần)", 8)
    hdr(dr, 5, ["Dòng tuần", "Dòng nền", "業務名", "Xử lý"], [10, 10, 40, 60])
    r = 6
    for x in log:
        if x.get("dup_of"):
            put(dr, r, 1, f"R{x['row']}")
            put(dr, r, 2, f"R{x['dup_of']}")
            put(dr, r, 3, x["name"])
            put(dr, r, 4, "Dòng nền không tính vào tải tuần (Record_Map cột P = N). Giữ dòng tuần.")
            r += 1
    r += 1
    section(dr, r, "B. Giống hệt tuần trước (có thể chép lại chưa cập nhật)", 8)
    r += 1
    hdr(dr, r, ["Tuần", "Số bản ghi", "Giống tuần trước", "Tỷ lệ"])
    for w in weeks:
        r += 1
        tot_w = sum(1 for x in log if x["wk"] == w)
        same = sum(1 for x in log if x["wk"] == w and x.get("carry"))
        put(dr, r, 1, f"W{w}")
        put(dr, r, 2, tot_w)
        put(dr, r, 3, same)
        put(dr, r, 4, f"=IF(B{r}=0,0,C{r}/B{r})", fmt="0%")
    r += 1
    put(dr, r, 1, "Báo cáo W43 tự ghi: 'phân biệt việc còn tiếp diễn với việc thực sự làm trong tuần'. "
        "Tuần có tỷ lệ cao nên được kiểm lại.", font=F_SUB, border=False)
    r += 2
    section(dr, r, "C. Nhiều tên khác nhau được gộp vào một mục", 8)
    r += 1
    hdr(dr, r, ["No.", "Mục", "Số tên", "Các tên trong log"])
    for no, it in items.items():
        if len(it["names"]) > 1:
            r += 1
            put(dr, r, 1, no)
            put(dr, r, 2, it["D"])
            put(dr, r, 3, len(it["names"]))
            put(dr, r, 4, " / ".join(it["names"]))
    r += 2
    section(dr, r, "D. Nhóm chồng lấn thời gian (cùng tuần, cùng một công việc ghi ở nhiều dòng)", 8)
    r += 1
    hdr(dr, r, ["Nhóm", "Các dòng / mục", "Chồng lấn", "Đề xuất"])
    overlaps = [
        ("OV-01", "Dòng nền R5 週次生産計画会議 16h, R15 生産計画調整・チュエン管理 15–20h, R19 Plan table調整 4–6h, R7 資材遅延対応に伴う生産計画再編 5–10h, R18 材料遅延 5–8h",
         "Cùng một chu kỳ Plantable tuần; cộng lại 45–60h/tuần chỉ riêng kế hoạch", "Merge: dùng No.18 + 4 + 107; ghi giờ một lần"),
        ("OV-02", "W20: R46 生産計画・進捗管理 15–20h, R45 進捗管理・社内調整 8–12h, R51 生産計画／調整業務 10–15h",
         "Tên mức nhóm (lấy từ danh mục 'リスト'), trùng nhau", "Clarify: thay bằng thao tác cụ thể"),
        ("OV-03", "W34–W43: Plantable更新 (18) 6–8h + ライン稼働・空き (104) 4–5h + 進捗・差異 (63) 4–5h + Risk (105) 3–4h + 資材遅延 (107) 3–5h + Delivery (77) 3–4h",
         "Các việc kiểm soát hằng ngày cùng dữ liệu Plantable; ranh giới không rõ", "Giữ tách nhưng ghi rõ phạm vi; đo thực tế"),
        ("OV-04", "Email (121) / Điều phối (122) / Báo cáo Nhật (124)", "Email là phương tiện của điều phối và báo cáo",
         "Clarify: ghi email riêng hoặc gộp vào 122"),
        ("OV-05", "Họp (85, 143) và nội dung kế hoạch (4, 18)", "Giờ họp có thể đã nằm trong giờ Plantable", "Clarify"),
        ("OV-06", "Tổ chức (126) / Kiểm kê–ủy quyền (127) / Leader (128) / MD–Planning (117)",
         "Cùng dự án thay đổi tổ chức, ghi 3–4 dòng mỗi tuần", "Giữ tách theo kết quả đầu ra; giới hạn thời gian"),
        ("OV-07", "Báo cáo tuần (87) và cập nhật 業務棚卸し (127)", "Báo cáo tuần là đầu vào của kiểm kê", "Merge vào 87"),
        ("OV-08", "Sự cố (64) / NPL trễ (107) / Risk (105)", "Một sự cố NPL có thể ghi ở cả 3", "Quy tắc: ghi ở 1 mục"),
        ("OV-09", "D/N (1) / Manpower (110) / Hợp đồng (111) – log W41–W43 (QC14名, Members List, 直間比率)",
         "Cùng chủ đề nhân sự tính phí", "Clarify phạm vi từng mục"),
        ("OV-10", "業務分析表 No.4 生産計画調整, 18, 20, 85, 88–90", "Trong 業務分析表 cũng tách 7 dòng cho một chu kỳ", "Giữ để khớp 業務分析表"),
    ]
    for o in overlaps:
        r += 1
        for j, v in enumerate(o):
            put(dr, r, 1 + j, v)
    dr.column_dimensions["B"].width = 60
    dr.column_dimensions["C"].width = 45
    dr.column_dimensions["D"].width = 50

    # ------------------------------------------------ Audit
    au = S["Audit"]
    title(au, "Audit – Kiểm tra workbook 業務棚卸し (Task 1)", "Số liệu là công thức trên Original_業務棚卸し / Record_Map.")
    O = "'Original_業務棚卸し'!"
    E_ = f"{O}$E$2:$E${LOG_LAST}"
    A_ = f"{O}$A$2:$A${LOG_LAST}"
    audit_rows = [
        ("File nguồn", os.path.basename(f_log).split("-", 1)[-1] + " / sheet 業務棚卸し", ""),
        ("Tổng bản ghi (có 業務名)", f"=COUNTA({E_})", "Brief ghi ~218 → số cũ; hiện tại lớn hơn"),
        ("  Dòng nền (không có 報告週)", f"=SUMPRODUCT(({E_}<>\"\")*({A_}=\"\"))", "Brief ghi 23 dòng nền"),
        ("  Bản ghi theo tuần", f"=SUMPRODUCT(({E_}<>\"\")*({A_}<>\"\"))", ""),
        ("Số tuần có ghi", f"=SUMPRODUCT(({A_}<>\"\")/COUNTIF({A_},{A_}&\"\"))", ""),
        ("Tuần đầu / cuối", f"=\"W\"&MIN({RM}!$C$4:$C${RM_LAST})&\" – W\"&MAX({RM}!$C$4:$C${RM_LAST})", ""),
        ("Bản ghi/tuần: ít nhất – nhiều nhất",
         f"=MIN(Workload_Analysis!$C${WL_W1}:$C${WL_W2})&\" – \"&MAX(Workload_Analysis!$C${WL_W1}:$C${WL_W2})", ""),
        ("Số tên nghiệp vụ khác nhau (nguyên văn)", f"=SUMPRODUCT(({E_}<>\"\")/COUNTIF({E_},{E_}&\"\"))", "Brief ghi ~191 → số cũ"),
        ("Số mục sau chuẩn hóa (No.) có trong log", f"=COUNTIF({MLC('O')},\">0\")", "Xem 業務一覧_Chanh"),
        ("Số giá trị 分類 khác nhau", f"=SUMPRODUCT(({O}$F$2:$F${LOG_LAST}<>\"\")/COUNTIF({O}$F$2:$F${LOG_LAST},{O}$F$2:$F${LOG_LAST}&\"\"))", "Không thống nhất → dùng 20 nhóm"),
        ("Số giá trị 頻度 khác nhau", f"=SUMPRODUCT(({O}$G$2:$G${LOG_LAST}<>\"\")/COUNTIF({O}$G$2:$G${LOG_LAST},{O}$G$2:$G${LOG_LAST}&\"\"))", "vd 毎日／週次, 都度/週次"),
        ("想定時間 dạng số", f"=COUNT({O}$H$2:$H${LOG_LAST})", ""),
        ("想定時間 dạng chuỗi (vd '2～3', '5～8h')", f"=COUNTA({O}$H$2:$H${LOG_LAST})-COUNT({O}$H$2:$H${LOG_LAST})", "Đã tách thành h thấp/h cao trong Record_Map"),
        ("Bản ghi tên mức nhóm (quá rộng)", f"=COUNTIF({RM}!$O$4:$O${RM_LAST},\"*Too broad*\")", ""),
        ("Bản ghi trùng hoàn toàn với dòng nền", f"=COUNTIF({RM}!$O$4:$O${RM_LAST},\"*Exact dup*\")", "W19 = chép dòng nền"),
        ("Bản ghi giống hệt tuần trước", f"=COUNTIF({RM}!$O$4:$O${RM_LAST},\"*Same as prev week*\")", "Có thể chưa cập nhật"),
        ("Bản ghi gán chưa chắc", f"=COUNTIF({RM}!$O$4:$O${RM_LAST},\"*Mapping?*\")", "Xem Record_Map cột X"),
        ("Bản ghi có 期限 (cột O)", f"=COUNTA({O}$O$2:$O${LOG_LAST})", ""),
        ("Bản ghi có 期待効果 (cột P)", f"=COUNTA({O}$P$2:$P${LOG_LAST})", ""),
        ("Công thức trong sheet gốc", n_formulas, "0 = chỉ có giá trị"),
        ("Tổng h giữa/tuần TB kỳ PT", f"=Workload_Analysis!F{WL_AVG}", "So với giờ khả dụng → có trùng giờ"),
        ("Tỷ lệ so với giờ khả dụng", f"=Workload_Analysis!H{WL_AVG}", ""),
        ("業務分析表 2026上期: số mục", len(an26), ""),
        ("業務分析表 2026上期: mục có FTE của Chanh", f"=COUNTIF({MLC('M')},\">0\")", ""),
        ("業務分析表 2026上期: tổng FTE của Chanh", f"=SUM({MLC('M')})", "Phải = 1.00"),
        ("Số mục mới (No.103+)", f"=COUNTIF({MLC('A')},\">=103\")", ""),
    ]
    hdr(au, 4, ["Kiểm tra", "Kết quả", "Ghi chú"], [50, 22, 50])
    for i, (k, f, nte) in enumerate(audit_rows):
        r = 5 + i
        put(au, r, 1, k)
        c = put(au, r, 2, f, align=CENTER)
        if "Tỷ lệ" in k:
            c.number_format = "0%"
        elif "FTE của Chanh" in k and "tổng" in k:
            c.number_format = "0.00"
        elif "h giữa" in k:
            c.number_format = "0.0"
        put(au, r, 3, nte)
    r = 5 + len(audit_rows) + 1
    section(au, r, "Số bản ghi theo 分類 gốc (15 giá trị nhiều nhất)", 3)
    for j, (k, v) in enumerate(Counter(x["cat"] for x in log).most_common(15)):
        put(au, r + 1 + j, 1, k)
        put(au, r + 1 + j, 2, f"=COUNTIF({O}$F$2:$F${LOG_LAST},A{r + 1 + j})", align=CENTER)
    r = r + 17
    section(au, r, "Số bản ghi theo 頻度 gốc", 3)
    for j, (k, v) in enumerate(Counter(x["freq"] for x in log).most_common()):
        put(au, r + 1 + j, 1, k)
        put(au, r + 1 + j, 2, f"=COUNTIF({O}$G$2:$G${LOG_LAST},A{r + 1 + j})", align=CENTER)

    # ------------------------------------------------ Source_Evidence
    se = S["Source_Evidence"]
    title(se, "Source_Evidence – Nguồn và trích xuất", "Mọi mục trong 業務一覧_Chanh truy được về một trong các nguồn dưới đây.")
    section(se, 4, "A. Danh sách nguồn", 9)
    hdr(se, 5, ["ID", "File (tên khi upload)", "Sheet", "Cập nhật", "Nội dung", "Dùng để", "Có trong workbook", "", ""],
        [7, 40, 26, 14, 46, 40, 26, 10, 10])
    meta = lambda p: openpyxl.load_workbook(p, read_only=True).properties  # noqa: E731
    srcs = [
        ("S1", f_log, "業務棚卸し + 21 sheet báo cáo tuần", "Bảng kiểm kê 434 bản ghi (W19–W43) + báo cáo tuần", "Bằng chứng việc thực tế, giờ ước tính", "Original_業務棚卸し (copy đủ); Record_Map"),
        ("S2", f_26, "<更新>APIC(2026年度上期)", "業務分析表 2026上期 – FTE từng người cho 102 mục", "Xương sống danh sách (No.1–102)", "Original_業務分析表 (cột Chanh/総計)"),
        ("S3", f_2324, "参考_APIC(2023/2024年度); 議事録", "業務分析表 2023/2024; biên bản 評価加算", "So sánh FTE qua các năm; mục 142", "Cột K–L 業務一覧"),
        ("S4", f_role, "APIC BTNB (2)", "業務分担表 có cột Chanh (●/○), tạo 2025-12", "Bằng chứng phụ", "Mục C dưới đây"),
        ("S5", f_transfer, "移管資料 / 週次スケジュール", "Tài liệu chuyển giao Plantable (2026-05-22)", "Bước Plantable, lịch tuần, họp KD", "Mục E dưới đây"),
        ("S6", f_handover, "10 sheet", "Danh sách bàn giao của Inoue (2023-04)", "Việc quản lý thừa hưởng", "Mục D dưới đây"),
        ("S7", None, "—", "Brief Codex §6 (buổi 業務棚卸 với Hori)", "Mục cần xác nhận", "Ghi trong cột 'Bằng chứng khác'"),
        ("S8", None, "—", "Repo: docs/quy-trinh-lead-time.md, data/nguon/*20260722*", "Dự án lead time (No.141)", ""),
        ("S9", find(args.src, "*2018.10APIC.xlsx"), "APIC BTNB", "業務分担・業務フロー表 2018.10", "Tham khảo lịch sử (mẫu gốc của S4)", "Không dùng trực tiếp"),
        ("S10", find(args.src, "*20240910.xlsx"), "生産管理業務 / 役割分担表0910", "Bảng nghiệp vụ TI (khối Uniqlo) – không có Chanh/APIC", "Chỉ tham khảo cấu trúc", "Không dùng"),
    ]
    for i, (sid, p, sh, cont, use, inwb) in enumerate(srcs):
        r = 6 + i
        put(se, r, 1, sid)
        put(se, r, 2, os.path.basename(p).split("-", 1)[-1] if p else "—")
        put(se, r, 3, sh)
        if p:
            pr = meta(p)
            put(se, r, 4, f"{pr.modified:%Y-%m-%d}" if pr.modified else "")
        else:
            put(se, r, 4, "")
        put(se, r, 5, cont)
        put(se, r, 6, use)
        put(se, r, 7, inwb)
    r = 6 + len(srcs) + 1
    put(se, r, 1, "Tên file khi upload bị mất ký tự tiếng Nhật (thành '_'). File gốc không đưa lên repo vì repo công khai.",
        font=F_SUB, border=False)

    r += 2
    section(se, r, "B. 業務分析表 APIC – FTE của Chanh qua các năm (tất cả mục)", 9)
    r += 1
    hdr(se, r, ["No.2026", "分類", "工程", "作業", "作業内容", "頻度", "Chanh 2023", "Chanh 2024", "Chanh 2026上期",
                f"{meta26['jp_name']} 2026", "Tất cả 2026", "Ô nguồn 2026", "Trong 業務一覧"])
    for it in an26:
        r += 1
        vals = [it["no"], it["A"], it["B"], it["D"], it["E"], it["F"], it["c23"], it["c24"], it["chanh"], it["jp"],
                it["total"], f"'<更新>APIC(2026年度上期) '!{it['chanh_ref']}", "Y" if it["no"] in items else ""]
        for j, v in enumerate(vals):
            put(se, r, 1 + j, v, fmt="0.00" if 6 <= j <= 10 else None)
    r += 1
    put(se, r, 4, "Tổng", font=F_BOLD)
    for j in (7, 8, 9, 10, 11):
        put(se, r, j, f"=SUM({CL(j)}{r - len(an26)}:{CL(j)}{r - 1})", font=F_BOLD, fmt="0.00")

    r += 3
    section(se, r, "C. 業務分担表（file Chanh, 2025-12/2026-01）– dòng có ●/○ ở cột Chanh", 9)
    r += 1
    hdr(se, r, ["Dòng", "分類", "工程", "作業", "作業内容", "工数(h)", "頻度", "Chanh", "No. trong 業務一覧"])
    for x in role:
        if not x["N"]:
            continue
        r += 1
        for j, v in enumerate([f"R{x['row']}", x["A"], x["B"], x["C"], x["D"], x["L"], x["M"], x["N"], x["no"]]):
            put(se, r, 1 + j, v)

    r += 3
    section(se, r, "D. Bàn giao của Inoue (2023-04)", 9)
    r += 1
    hdr(se, r, ["Sheet", "Dòng", "Nội dung", "Dấu", "Ghi chú (VI)", "No. trong 業務一覧"])
    for h in handover:
        r += 1
        for j, v in enumerate([h["sheet"], f"B{h['row']}", h["text"], h["mark"], h["note"], h["no"] or "—"]):
            put(se, r, 1 + j, v)

    r += 3
    section(se, r, "E. 移管資料 / 週次スケジュール (2026-05-22) – mục liên quan Chanh", 9)
    r += 1
    hdr(se, r, ["Sheet", "Ô", "Nội dung", "No. trong 業務一覧"])
    for sh, ref, txt, no in TRANSFER_ITEMS:
        r += 1
        for j, v in enumerate([sh, ref, txt, no]):
            put(se, r, 1 + j, v)

    r += 3
    section(se, r, "F. Tỷ lệ tải Chanh tự ước tính trong báo cáo tuần (không phải đo thực tế)", 9)
    r += 1
    hdr(se, r, ["Sheet (tuần)", "Ô", "Nhóm việc", "Tỷ lệ"])
    for sh, ref, lab, val in self_est:
        r += 1
        for j, v in enumerate([sh, ref, lab, val]):
            put(se, r, 1 + j, v)
    se.freeze_panes = "A6"

    # ------------------------------------------------ Action_Plan
    apn = S["Action_Plan"]
    title(apn, "Action_Plan – Việc cần xác nhận và bước tiếp theo",
          "Giai đoạn 1 = hoàn thiện inventory. Đề xuất giao việc (Role_Analysis) chỉ triển khai sau khi Chanh/Hori duyệt.")
    APH = ["ID", "Việc cần làm", "Lý do / căn cứ", "Người phụ trách", "Thời hạn đề xuất", "Ưu tiên", "Trạng thái", "Liên kết"]
    hdr(apn, 4, APH, [7, 52, 52, 16, 14, 9, 12, 30])
    acts = [
        ("A01", "Rà các mục 'To Verify' trong 業務一覧_Chanh, điền cột Y (Confirmed / Not Current)",
         f"=COUNTIF({MLC('Z')},\"To Verify\")&\" mục đang To Verify\"", "Chanh", "1 tuần", "High", "業務一覧_Chanh cột Y"),
        ("A02", "Xác nhận các bản ghi gán chưa chắc (Record_Map cột X) và sửa No. nếu sai",
         f"=COUNTIF({RM}!$O$4:$O${RM_LAST},\"*Mapping?*\")&\" bản ghi\"", "Chanh", "1 tuần", "High", "Record_Map"),
        ("A03", "Xác nhận các mục có FTE trong 業務分析表 nhưng không có trong log (audit, khách, báo cáo tháng, họp sáng, quản lý staff…)",
         "Missing_Tasks: Clarify", "Chanh", "1 tuần", "High", "Missing_Tasks"),
        ("A04", "Thêm cột 業務No. (theo danh sách này) vào mẫu 業務棚卸し; thống nhất 頻度 và ghi giờ dạng số (thấp/cao)",
         "Audit: 60 giá trị 分類, 19 giá trị 頻度, giờ dạng chuỗi", "Chanh", "Từ W45", "High", "Audit"),
        ("A05", "Đo thời gian thực tế 2 tuần (time log theo 業務No.) cho các nhóm lớn: Plantable, điều phối/email, hợp đồng, họp, sự cố",
         f"=\"Tổng giờ ghi TB \"&TEXT(Workload_Analysis!F{WL_AVG},\"0\")&\"h/tuần so với \"&TEXT({P_AVAIL},\"0\")&\"h khả dụng\"",
         "Chanh", "2 tuần", "High", "Workload_Analysis mục 4"),
        ("A06", "Xác nhận giờ làm thực tế/tuần (có tăng ca?) và sửa tham số C5–C6", "Giả định 8h × 6 ngày", "Chanh", "1 tuần", "Medium", "Workload_Analysis C5:C6"),
        ("A07", "Báo cáo tuần: chỉ ghi việc thực sự làm trong tuần (như W43), không chép tuần trước",
         "Duplicate_Review mục B", "Chanh", "Từ W45", "Medium", "Duplicate_Review"),
        ("A08", "Thống nhất giờ Plantable (16h dòng nền, 2 ngày/tuần theo buổi với Hori, 6–8h log W34+)",
         "Workload_Analysis mục 5", "Chanh + Hori", "Sau A05", "High", "Workload_Analysis"),
        ("A09", "Review đề xuất trong Role_Analysis cùng Hori: chọn 5–8 mục giao trước (ưu tiên mức C và E)",
         "Role_Analysis mục A", "Chanh + Hori", "Sau A01–A03", "High", "Role_Analysis"),
        ("A10", "Lập RACI cho các mục giao: người làm / người kiểm / người duyệt / quy tắc escalation R1–R3",
         "Brief §3: giao quyền đi kèm trách nhiệm, quy tắc, escalation", "Chanh", "Sau A09", "High", "Role_Analysis"),
        ("A11", "Đề xuất cập nhật 業務分析表 2026下期 theo việc thực tế (mục No.103+) – qua Hori/TI",
         "業務分析表 gắn với phân bổ phí ủy thác và 評価加算; biên bản 20231127 A37 yêu cầu báo TI khi 100 mục thay đổi",
         "Chanh + Hori", "Trước 2026下期", "Medium", "Source_Evidence B"),
        ("A12", "Cập nhật brief: số bản ghi hiện tại là 434 (không phải 218), 323 tên (không phải 191); 23 dòng nền đúng",
         "Audit", "Chanh", "Ngay", "Low", "Audit"),
        ("A13", "Quyết định nơi lưu file nguồn (repo hiện công khai)", "File nguồn chứa phân bổ chi phí và thông tin hải quan", "Chanh", "Ngay", "Medium", "README"),
    ]
    for i, (aid, task, why, owner, due, pri, link) in enumerate(acts):
        r = 5 + i
        for j, v in enumerate([aid, task, why, owner, due, pri]):
            put(apn, r, 1 + j, v)
        put(apn, r, 7, "Chưa làm", font=F_INPUT, fill=FILL_INPUT, align=CENTER)
        put(apn, r, 8, link)
    dv_list(apn, f"G5:G{4 + len(acts)}", ["Chưa làm", "Đang làm", "Xong"])
    apn.freeze_panes = "C5"

    # ------------------------------------------------ Advice
    ad = S["Advice"]
    title(ad, "Advice – Tư vấn và gợi ý", "Ý kiến của Claude dựa trên dữ liệu trong file. Số liệu ở cột C là công thức, tự cập nhật.")
    hdr(ad, 4, ["#", "Gợi ý", "Số liệu liên quan"], [5, 110, 30])
    advice = [
        ("Dữ liệu giờ hiện chưa đủ để kết luận tải. Giờ trong 業務棚卸し là ước tính và cộng trùng; tổng mỗi tuần thường vượt giờ làm. "
         "Đo thực tế 2 tuần theo 業務No. trước khi chốt con số giao việc.",
         f"=\"Giờ ghi TB: \"&TEXT(Workload_Analysis!F{WL_AVG},\"0\")&\"h / \"&TEXT({P_AVAIL},\"0\")&\"h\""),
        ("Phần lớn giờ ghi trong log nằm ở các việc không có trong 業務分析表 (hợp đồng TI–APIC, tổ chức, cải tiến, rủi ro, điều phối, email). "
         "Nên đề xuất cập nhật 業務分析表 2026下期 – nhưng vì nó gắn với phí ủy thác và 評価加算, phải qua Hori và TI.",
         f"=TEXT(SUMIFS({MLC('T')},{MLC('A')},\">=103\")/{P_AVAIL},\"0%\")&\" giờ log ở mục No.103+\""),
        ("Plantable: mục tiêu từ ~2 ngày/tuần xuống ≤ 0,5 ngày. Planning Leader cập nhật theo 7 quy tắc chọn chuyền (移管資料) và lịch tuần cố định; "
         "Chanh chỉ review thay đổi lớn và case R1. Viết các quy tắc thành SOP trước khi giao.",
         f"=TEXT(Workload_Analysis!G{X0 + 1},\"0.0\")&\"h/tuần (log, kế hoạch)\""),
        ("Chỉ nhận lên mình case R1. Bảng R1/R2/R3 đã có (No.105) – thêm cột 'người quyết định' cho từng mức để Leader/MD tự xử lý R2/R3.",
         ""),
        ("Email và điều phối: đặt 2 khung giờ xử lý mail/ngày, quy tắc CC, mẫu mail Nhật–Việt, dùng AI dịch (No.135). "
         "Lập bảng 'ai trả lời câu hỏi gì' với TIJ/SW1/SW2 để câu hỏi thường không qua Chanh.",
         f"=TEXT(SUMIFS({MLC('T')},{MLC('A')},121)+SUMIFS({MLC('T')},{MLC('A')},122),\"0.0\")&\"h/tuần (121+122)\""),
        ("Họp: gửi danh sách thông tin còn thiếu trước họp (như 移管資料 đề xuất), giới hạn 60–90 phút, MD ghi biên bản và TODO có hạn.",
         ""),
        ("Báo cáo: gộp 業務日報, báo cáo tuần, Workload report, Management report, báo cáo Nhật thành 1 bộ tuần; báo cáo tháng nhỏ (tồn kho, thiết bị, mã ngoài 2 tuần) giao staff và tự động hóa.",
         f"=TEXT(SUMIFS(Workload_Analysis!I{CAT0}:I{CAT0 + 19},Workload_Analysis!B{CAT0}:B{CAT0 + 19},\"17\"),\"0.0\")&\"h/tuần nhóm 17\""),
        ("Hợp đồng TI–APIC là dự án có điểm kết thúc. Sau khi ký, tải nhóm 07 sẽ giảm – nên lên kế hoạch dùng giờ đó cho người và hệ thống, không để việc vận hành lấp vào.",
         f"=TEXT(SUMIFS(Workload_Analysis!I{CAT0}:I{CAT0 + 19},Workload_Analysis!B{CAT0}:B{CAT0 + 19},\"07\"),\"0.0\")&\"h/tuần nhóm 07\""),
        ("Việc hành chính (xe, khách sạn, guesthouse, thẻ hội viên, Gala, giám khảo) không phải việc của 課長 – chuyển GA/Admin.",
         f"=TEXT(SUMIFS({RA}!$F$37:$F$200,{RA}!$J$37:$J$200,\"D*\"),\"0.0\")&\"h/tuần mức D\""),
        ("Việc quản lý người (duyệt chuyên cần, đánh giá, tuyển dụng, 1-on-1 với Leader) hầu như không xuất hiện trong log. "
         "Đây là việc cốt lõi của 課長 – nên đặt lịch cố định (ví dụ 30 phút/tuần/Leader) và ghi vào log.",
         ""),
        ("Khi giao việc: giao cả quyền và tiêu chí (khi nào tự quyết, khi nào báo). Tránh 'giao ngược' – Leader hỏi lại mọi thứ. "
         "Chu kỳ 任せる→確認→評価 (log W36) là hướng đúng; đo bằng số case Leader tự đóng mỗi tuần.",
         ""),
        ("Mục tiêu tỷ trọng: tăng Decision + System Building + Control, giảm Operation. Theo dõi chỉ số này mỗi tháng trong Role_Analysis mục B.",
         f"={RA}!C32"),
        ("Tải cao kéo dài là rủi ro sức khỏe và rủi ro tổ chức (phụ thuộc 1 người). Dùng file này để thống nhất với Hori thứ tự ưu tiên, "
         "và những việc sẽ dừng hoặc làm chậm lại khi quá tải.",
         f"={RA}!C28"),
    ]
    for i, (t, f) in enumerate(advice):
        r = 5 + i
        put(ad, r, 1, i + 1, align=CENTER)
        put(ad, r, 2, t)
        put(ad, r, 3, f, fmt="0%")
        ad.row_dimensions[r].height = 48

    # ------------------------------------------------ README
    rd = S["README"]
    title(rd, "業務棚卸 – Chanh (TIVN–APIC) | Kiểm kê nghiệp vụ & phân tích tải",
          "Tạo bằng nghiep-vu/build_inventory.py. Không sửa dữ liệu nguồn; mọi tổng hợp là công thức.")
    rd.column_dimensions["A"].width = 26
    rd.column_dimensions["B"].width = 110
    lines = [
        ("Mục đích", "Liệt kê đầy đủ việc Chanh đang làm, dựa trên 業務分析表 2026年度上期 (No.1–102) và bổ sung việc thực tế ngoài bảng (No.103+); "
                     "rà tải, phạm vi theo chức vụ, đề xuất giảm tải/giao việc (dự thảo)."),
        ("Cách đọc", "1) 業務一覧_Chanh: danh sách nghiệp vụ (= Normalized_Inventory của brief). 2) Workload_Analysis: tải theo nhóm, theo tuần. "
                     "3) Role_Analysis: phù hợp chức vụ + đề xuất. 4) Missing_Tasks, Duplicate_Review, Action_Plan, Advice."),
        ("Sheet theo brief", "Original_Inventory → Original_業務棚卸し (+ Original_業務分析表); Normalized_Inventory → 業務一覧_Chanh (+ Record_Map nối từng bản ghi); "
                             "các sheet còn lại giữ tên như brief."),
        ("Status", "Confirmed = có trong log tuần trong 13 tuần gần nhất. To Verify = chỉ có trong 業務分析表/nguồn khác, hoặc lâu không xuất hiện. "
                   "Not Current = lịch sử/dự kiến. Chanh ghi đè ở cột Y."),
        ("Giờ", "Log: giờ Chanh ước tính (想定時間), không phải đo thực tế. 'Chuẩn hóa' chia lại để tổng tuần = giờ khả dụng (tránh cộng trùng). "
                "FTE: phân bổ trong 業務分析表 × giờ khả dụng."),
        ("Màu", "Chữ xanh dương / nền vàng = giá trị nhập, sửa được. Chữ xanh lá = liên kết tới sheet khác. Chữ đen = công thức."),
        ("Đề xuất", "Cột I–L của Role_Analysis là dự thảo của Claude. Không có việc giao nào được quyết định khi chưa có phê duyệt của Chanh/Hori."),
        ("Nguồn", "Xem Source_Evidence mục A. File gốc không đưa lên repo (repo công khai)."),
    ]
    for i, (k, v) in enumerate(lines):
        put(rd, 4 + i, 1, k, font=F_BOLD)
        put(rd, 4 + i, 2, v)
        rd.row_dimensions[4 + i].height = 32
    r = 4 + len(lines) + 1
    section(rd, r, "Số liệu chính (công thức)", 2)
    key = [
        ("Bản ghi trong 業務棚卸し", "=Audit!B6"),
        ("Tên nghiệp vụ khác nhau", "=Audit!B12"),
        ("Mục trong 業務一覧_Chanh", f"=COUNT({MLC('A')})"),
        ("  Confirmed / To Verify / Not Current",
         f"=COUNTIF({MLC('Z')},\"Confirmed\")&\" / \"&COUNTIF({MLC('Z')},\"To Verify\")&\" / \"&COUNTIF({MLC('Z')},\"Not Current\")"),
        ("Giờ ghi TB/tuần (kỳ PT) vs khả dụng", f"=TEXT(Workload_Analysis!F{WL_AVG},\"0\")&\"h vs \"&TEXT({P_AVAIL},\"0\")&\"h\""),
        ("Tải theo giờ ghi: hiện tại → sau đề xuất", f"=TEXT({RA}!C28,\"0%\")&\" → \"&TEXT({RA}!C31,\"0%\")&\" giờ khả dụng\""),
        ("Giờ ở mức C+D+E (nên giao/chuyển/giảm)", f"=TEXT({RA}!D8+{RA}!D9+{RA}!D10,\"0.0\")&\"h/tuần (\"&TEXT({RA}!E8+{RA}!E9+{RA}!E10,\"0%\")&\")\""),
    ]
    for i, (k, f) in enumerate(key):
        put(rd, r + 1 + i, 1, k)
        put(rd, r + 1 + i, 2, f)

    # tab colors
    for nme, col in [("README", "1F3864"), ("業務一覧_Chanh", "C00000"), ("Role_Analysis", "C00000"),
                     ("Workload_Analysis", "C00000"), ("Original_業務分析表", "808080"),
                     ("Original_業務棚卸し", "808080"), ("Source_Evidence", "808080")]:
        S[nme].sheet_properties.tabColor = col
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90

    wb.save(args.out)
    print(f"saved {args.out}: {len(items)} items, {len(log)} log records, {len(gaps)} gaps, weeks {weeks[0]}–{weeks[-1]}")
    print("self-estimates:", len(self_est))


if __name__ == "__main__":
    main()
