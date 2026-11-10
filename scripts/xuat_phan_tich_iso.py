"""Xuất file Excel phân tích công việc 3 nhân sự ISO kiêm kế hoạch (APIC, kỳ 03–10/11/2026).

Chạy: python3 scripts/xuat_phan_tich_iso.py [đường_dẫn_ra.xlsx]
Mặc định ghi ra docs/iso-ke-hoach/Phan-tich-cong-viec-ISO.xlsx.
Nguồn: ba file trong data/nguon/iso/, tách bằng scripts/iso_nhat_ky.py.
Mọi số thống kê là công thức Excel tính từ sheet NhatKy và LoiHienTruong.
"""
import re
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, "scripts")
from iso_nhat_ky import NGUON, VIEC, doc_tat_ca  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else "docs/iso-ke-hoach/Phan-tich-cong-viec-ISO.xlsx"

FONT = "Arial"
INK, MUTED, ACCENT = "17203B", "5B6482", "2B3F9E"
HEAD, INPUT_FILL = "E6EAFA", "FFF7CC"
HUONG_FILL = {"GIỮ": "E1F3E9", "CHUYỂN": "FFF1D6", "BỎ": "FBE6E4"}
thin = Side(style="thin", color="D8DDE8")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
NGUOI = list(NGUON)  # THAN, VUONG, DAO


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(h):
    return PatternFill("solid", start_color=h, end_color=h)


def put(ws, r, c, v, bold=False, bg=None, align=WRAP, fmt=None, color=INK, size=10, border=True, italic=False):
    cell = ws.cell(r, c, v)
    cell.font = f(size, bold, color, italic)
    cell.alignment = align
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = BOX
    if fmt:
        cell.number_format = fmt
    return cell


def header(ws, r, labels, widths=None):
    for i, t in enumerate(labels, 1):
        put(ws, r, i, t, bold=True, bg=HEAD, align=CENTER)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def title(ws, text, sub=None):
    put(ws, 1, 1, text, bold=True, size=14, color=ACCENT, border=False, align=Alignment(vertical="center"))
    if sub:
        put(ws, 2, 1, sub, color=MUTED, italic=True, border=False, align=Alignment(vertical="center"))


def huong(ma):
    q = VIEC[ma][3]
    return "GIỮ" if q.startswith("GIỮ") else ("BỎ" if q.startswith("BỎ") else "CHUYỂN")


def hm(m):
    return f"{m // 60:02d}:{m % 60:02d}"


# --------------------------------------------------------------------------- dữ liệu
rows = doc_tat_ca()
log = []  # mỗi dòng nhật ký (tách phần phút phụ của khung gộp 5S + VDKL)
for x in rows:
    log.append((x, x["ma"], x["phut"], x["ghi_chu"]))
    for m, p in x.get("phut_phu", {}).items():
        log.append((x, m, p, "Phần VDKL của khung gộp '5S và VDKL' (chia đôi khung giờ)"))

# Lỗi hiện trường: dòng của việc I1/I2 có nội dung cụ thể (bỏ dòng mô tả chung)
CHUNG = ["kiểm tra vệ sinh 5s,bảo hộ", "ghi nhận hình ảnh", "kiểm tra xác suất", "kiểm tra xác xuất",
         "kiểm tra bảng đánh dấu", "bảng kiểm tra vdkl hàng ngày có", "bảng kiểm tra vdkl hàng ngày của quản lý",
         "*kiểm tra công nhân", "bảng chéo vật dụng, ktra", "đầu ca vào làm việc thực hiện tốt"]
# (người, dòng) ghi nhận không phải lỗi: khu vực đạt, kiểm tra không phát hiện lỗi
KHONG_LOI = {("THAN", 10), ("THAN", 58), ("VUONG", 49), ("VUONG", 56), ("VUONG", 92), ("VUONG", 132),
             ("VUONG", 177), ("VUONG", 189), ("DAO", 8), ("DAO", 18), ("DAO", 113)}
# Dòng kiểm tra lại lỗi của ngày trước
KIEM_LAI = {("THAN", 32), ("THAN", 56), ("THAN", 57), ("THAN", 99), ("THAN", 129), ("VUONG", 91),
            ("VUONG", 108), ("VUONG", 111), ("VUONG", 129), ("VUONG", 130), ("VUONG", 131)}

CHU_DE = [
    ("VDKL / VDNH / kim", r"vdkl|vdnh|bảng chéo|chắn kim|kim loại"),
    ("BHLĐ: nón, bao tay", r"nón|bao tay"),
    ("Điện, cơ sở vật chất, an toàn", r"điện|đèn|gạch|kính|rèm|trụ chỉ|nút cao su|lối thoát|mút|cách điện|ổ thuyền"),
    ("Rập chưa trả / để sai chỗ", r"rập"),
    ("Bàn ủi, khu ủi", r"ủi"),
    ("BTP chạm đất, để lộn xộn", r"chạm đất|rơi|btp|bán thành phẩm|hàng hoá|hàng thành phẩm|hàng chạm|để hàng"),
    ("Vệ sinh máy, nền, khu vực", r"vệ sinh|dơ|bụi|dầu|keo|rác|qr"),
    ("Đồ dùng sai chỗ, treo trên máy", r"treo|đồ dùng|chai nước|bảng tên|túi|nhãn|kệ|chỉ|dây|giấy|sắp xếp|gọn"),
]


def chu_de(t):
    t = t.lower()
    for ten, pat in CHU_DE:
        if re.search(pat, t):
            return ten
    return "Khác"


def bien_phap(t):
    t = t.lower()
    if not t.strip():
        return "Không ghi biện pháp"
    if re.search(r"biên bản|abc|hạ loại", t):
        return "Kỷ luật / đe dọa xử lý ABC"
    if re.search(r"cơ điện|anh quang|anh đức|a kha|anh kha|kcs|kỹ thuật|ktt|tổ mẫu|chị liên|cải tiến", t):
        return "Chuyển bộ phận khác xử lý"
    if re.search(r"cung cấp|cho thay|đã treo|đã cho|cho công nhân thay|chuyển về kho", t):
        return "ISO tự xử lý / cấp vật tư"
    if re.search(r"ql đã|quản lý đã gửi|công nhân đã|quản lý kiểm tra và đã", t):
        return "Quản lý tự khắc phục, gửi ảnh"
    return "Nhắc nhở / yêu cầu chung"


def trang_thai(t):
    t = t.lower()
    if not t.strip():
        return "Không ghi kết quả"
    if "chưa" in t:
        return "Chưa đạt"
    if re.search(r"sẽ kiểm tra|kiểm tra lại|sẽ thực hiện", t):
        return "Hẹn kiểm tra lại"
    return "Báo đạt trong ngày"


def chuan_khu_vuc(t):
    t = t.replace("\n", " ").strip()
    m = re.findall(r"\b(\d{1,2})\s*([ab])\b", t.lower())
    if m and len(m) == 1:
        return f"{m[0][0]}{m[0][1].upper()}"
    return t[:30] if t else "(không ghi)"


loi = []
for x in rows:
    if x["ma"] not in ("I1", "I2"):
        continue
    e = x["noi_dung"].lower()
    if not e or any(e.startswith(g) for g in CHUNG):
        continue
    k = (x["nguoi"], x["dong"])
    loai = "Không lỗi" if k in KHONG_LOI else ("Kiểm tra lại lỗi cũ" if k in KIEM_LAI else "Lỗi")
    loi.append(dict(x=x, loai=loai, chu_de=chu_de(x["noi_dung"] + " " + x["khu_vuc"]),
                    bp=bien_phap(x["bien_phap"]), tt=trang_thai(x["ket_qua"]),
                    kv=chuan_khu_vuc(x["khu_vuc"])))

wb = Workbook()

# --------------------------------------------------------------------------- NhatKy
wsL = wb.active
wsL.title = "NhatKy"
title(wsL, "Nhật ký công việc tách từ 3 báo cáo (mỗi dòng = một dòng báo cáo gốc)",
      "Phút phân bổ: khung giờ tự ghi trong báo cáo, chia đều cho các mã việc có trong khung. Không phải giờ đo thực tế.")
LH = ["Mã người", "Người báo cáo", "Ngày", "Bắt đầu", "Kết thúc", "Phút phân bổ", "Mã việc", "Hướng xử lý",
      "Hạng mục (gốc)", "Khu vực", "Nội dung (gốc)", "Biện pháp (gốc)", "Kết quả (gốc)", "Ô nguồn", "Ghi chú"]
header(wsL, 4, LH, [9, 20, 10, 8, 8, 9, 8, 10, 34, 14, 48, 36, 18, 22, 30])
r = 5
for x, ma, phut, note in log:
    vals = [x["nguoi"], NGUON[x["nguoi"]][1], x["ngay"], hm(x["bat_dau"]), hm(x["ket_thuc"]), round(phut, 2),
            ma, huong(ma), x["hang_muc"], x["khu_vuc"], x["noi_dung"], x["bien_phap"], x["ket_qua"],
            f"{NGUON[x['nguoi']][0].split('/')[-1]}!E{x['dong']}", note]
    for c, v in enumerate(vals, 1):
        put(wsL, r, c, v, fmt="dd/mm/yyyy" if c == 3 else ("0.0" if c == 6 else None))
    r += 1
L_LAST = r - 1
wsL.freeze_panes = "A5"
wsL.auto_filter.ref = f"A4:O{L_LAST}"


def rng(col):
    return f"NhatKy!${col}$5:${col}${L_LAST}"


# --------------------------------------------------------------------------- ThoiGian
wsT = wb.create_sheet("ThoiGian", 0)
title(wsT, "Thống kê thời gian theo mã việc (phút/ngày bình quân, theo giờ tự ghi trong báo cáo)",
      "Số ngày có báo cáo, phút/ngày và tỷ lệ đều là công thức tính từ sheet NhatKy.")
put(wsT, 4, 1, "Người báo cáo", bold=True, bg=HEAD)
put(wsT, 5, 1, "Mã người", bold=True, bg=HEAD)
put(wsT, 6, 1, "Số ngày có báo cáo", bold=True, bg=HEAD)
put(wsT, 7, 1, "Giờ ghi / ngày", bold=True, bg=HEAD)
PC = {}  # cột phút/ngày của từng người
for i, p in enumerate(NGUOI):
    c = 5 + i * 3
    PC[p] = c
    wsT.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 2)
    put(wsT, 4, c, NGUON[p][1], bold=True, bg=HEAD, align=CENTER)
    wsT.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 2)
    put(wsT, 5, c, p, bg=HEAD, align=CENTER)
    wsT.merge_cells(start_row=6, start_column=c, end_row=6, end_column=c + 2)
    put(wsT, 6, c, f'=SUMPRODUCT(({rng("A")}="{p}")/COUNTIFS({rng("A")},{rng("A")},{rng("C")},{rng("C")}))',
        align=CENTER, fmt="0")
    wsT.merge_cells(start_row=7, start_column=c, end_row=7, end_column=c + 2)
    put(wsT, 7, c, f'=SUMIFS({rng("F")},{rng("A")},"{p}")/60/{get_column_letter(c)}6', align=CENTER, fmt="0.00")
for c in range(2, 5):
    for rr in range(4, 8):
        put(wsT, rr, c, None, bg=HEAD)
TH = ["Mã việc", "Nội dung công việc", "Nhóm hiện tại", "Hướng xử lý"]
r = 9
for i, t in enumerate(TH, 1):
    put(wsT, r, i, t, bold=True, bg=HEAD, align=CENTER)
for p in NGUOI:
    for j, t in enumerate(["Phút/ngày", "% thời gian", "Số ngày có"]):
        put(wsT, r, PC[p] + j, t, bold=True, bg=HEAD, align=CENTER)
put(wsT, r, 14, "Tổng 3 người (giờ/ngày)", bold=True, bg=HEAD, align=CENTER)
put(wsT, r, 15, "Chức danh / bộ phận nên làm chính", bold=True, bg=HEAD, align=CENTER)
for col, w in zip("ABCDEFGHIJKLMNO", [8, 46, 12, 11, 9, 9, 8, 9, 9, 8, 9, 9, 8, 11, 44]):
    wsT.column_dimensions[col].width = w
r = 10
T_FIRST = r
for ma, (ten, nhom, owner, _) in VIEC.items():
    put(wsT, r, 1, ma, bold=True, align=CENTER)
    put(wsT, r, 2, ten)
    put(wsT, r, 3, nhom, align=CENTER)
    put(wsT, r, 4, huong(ma), bold=True, bg=HUONG_FILL[huong(ma)], align=CENTER)
    for p in NGUOI:
        c = PC[p]
        L = get_column_letter
        put(wsT, r, c, f'=SUMIFS({rng("F")},{rng("A")},"{p}",{rng("G")},$A{r})/{L(c)}$6', fmt="0", align=CENTER)
        put(wsT, r, c + 1, f'=IFERROR({L(c)}{r}/({L(c)}$7*60),0)', fmt="0%", align=CENTER)
        put(wsT, r, c + 2,
            f'=SUMPRODUCT(({rng("A")}="{p}")*({rng("G")}=$A{r})/COUNTIFS({rng("A")},{rng("A")},{rng("G")},{rng("G")},{rng("C")},{rng("C")}))',
            fmt="0", align=CENTER)
    put(wsT, r, 14, "=(" + "+".join(f"{get_column_letter(PC[p])}{r}" for p in NGUOI) + ")/60", fmt="0.0", align=CENTER)
    put(wsT, r, 15, owner)
    r += 1
T_LAST = r - 1
r += 1
put(wsT, r, 1, "Tổng hợp theo hướng xử lý", bold=True, color=ACCENT, border=False)
r += 1
T_SUM = {}
for h, nhan in [("GIỮ", "Việc ISO giữ lại (phải đổi cách làm)"), ("CHUYỂN", "Việc nên chuyển về đúng bộ phận"),
                ("BỎ", "Việc nên bỏ (vai trò trung gian)")]:
    put(wsT, r, 1, h, bold=True, bg=HUONG_FILL[h], align=CENTER)
    wsT.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    put(wsT, r, 2, nhan)
    for p in NGUOI:
        c = PC[p]
        L = get_column_letter(c)
        put(wsT, r, c, f'=SUMIF($D${T_FIRST}:$D${T_LAST},$A{r},{L}${T_FIRST}:{L}${T_LAST})', fmt="0", align=CENTER)
        L2 = get_column_letter(c + 1)
        put(wsT, r, c + 1, f'=SUMIF($D${T_FIRST}:$D${T_LAST},$A{r},{L2}${T_FIRST}:{L2}${T_LAST})', fmt="0%", align=CENTER)
    put(wsT, r, 14, f'=SUMIF($D${T_FIRST}:$D${T_LAST},$A{r},$N${T_FIRST}:$N${T_LAST})', fmt="0.0", align=CENTER)
    T_SUM[h] = r
    r += 1
put(wsT, r, 1, "Riêng việc ISO giữ lại, phần tuần tra 5S + VDKL (I1 + I2) chiếm:", bold=True, border=False)
r += 1
put(wsT, r, 1, "I1+I2", bold=True, align=CENTER)
wsT.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
put(wsT, r, 2, "Tuần tra 5S / nội quy / BHLĐ và kiểm VDKL – phần lớn là việc kiểm tra hằng ngày của quản lý chuyền")
for p in NGUOI:
    c = PC[p]
    L = get_column_letter(c)
    put(wsT, r, c, f"={L}{T_FIRST}+{L}{T_FIRST + 1}", fmt="0", align=CENTER)
    L2 = get_column_letter(c + 1)
    put(wsT, r, c + 1, f"={L2}{T_FIRST}+{L2}{T_FIRST + 1}", fmt="0%", align=CENTER)
put(wsT, r, 14, f"=N{T_FIRST}+N{T_FIRST + 1}", fmt="0.0", align=CENTER)
T_I12 = r
r += 2
for t in [
    "Cách đọc: % thời gian = phút/ngày của mã việc ÷ tổng giờ tự ghi/ngày của người đó.",
    "Giả định 1: số phút trong một khung giờ chia đều cho các mã việc ghi trong khung (báo cáo không ghi riêng).",
    "Giả định 2: khung '15h' của Thân không có giờ kết thúc, tính 15h00–15h45. Khung '5S và VDKL' của Vương chia đôi cho I1 và I2.",
    "Giới hạn: khung giờ trong báo cáo gần như giống nhau mỗi ngày (xem sheet GhiChuDuLieu) nên đây là cơ cấu thời gian tự khai, "
    "cần đo lại trong 10 ngày thử nghiệm.",
]:
    put(wsT, r, 1, t, color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
    r += 1
wsT.freeze_panes = "E10"

# --------------------------------------------------------------------------- TrungLap
wsD = wb.create_sheet("TrungLap", 1)
title(wsD, "Ma trận trùng lặp: việc 3 nhân sự ISO đang làm so với chức năng của bộ phận khác",
      "Chủ trì = bộ phận có chức năng và quyền làm việc này. Phối hợp = cung cấp thông tin hoặc hỗ trợ. "
      "Cột ISO: 'Làm thay' = ISO đang làm việc của người khác.")
BP = ["Quản lý chuyền / chuyền trưởng", "Quản đốc xưởng", "Kế hoạch / điều độ", "Phòng cắt", "In / thêu / laze",
      "Kho phụ liệu", "Cơ điện", "Kỹ thuật chuyền / tổ mẫu", "Thống kê / văn thư", "Nhóm 5S xưởng", "ISO"]
MT = {
    "I1": ["Chủ trì", "Phối hợp", "", "", "", "", "Phối hợp", "", "", "Phối hợp", "Làm thay (tuần tra 2–4 lượt/ngày); "
           "đúng vai trò nếu chỉ đánh giá xác suất"],
    "I2": ["Chủ trì", "", "", "", "", "", "Phối hợp", "", "", "", "Làm thay phần kiểm hằng ngày; đúng vai trò khi xác minh xác suất"],
    "I3": ["", "", "", "", "", "", "", "Chủ trì", "Phối hợp", "", "Làm thay (việc nhập số liệu)"],
    "I4": ["", "", "Chủ trì", "", "", "", "", "Chủ trì", "", "", "Đúng chức năng với tài liệu hệ thống; làm thay khi phân phối tài liệu sản xuất theo mã"],
    "I5": ["", "", "", "", "", "", "", "", "", "", "Đúng chức năng"],
    "P1": ["Chủ trì", "", "Chủ trì", "Phối hợp", "", "", "", "", "", "", "Làm thay"],
    "P2": ["", "", "Chủ trì", "Chủ trì", "Chủ trì", "", "", "", "", "", "Làm thay, tạo kênh hối thúc thứ hai"],
    "P3": ["Chủ trì", "", "", "", "", "Chủ trì", "", "", "", "", "Làm thay (trung chuyển phiếu)"],
    "P4": ["Phối hợp", "Chủ trì", "Chủ trì", "", "", "", "", "", "", "", "Làm thay, vượt quyền (phân chia kế hoạch chuyền)"],
    "P5": ["Phối hợp", "", "Phối hợp", "", "", "", "", "", "Chủ trì", "", "Làm thay (đánh máy, in, cấp phát)"],
    "P6": ["Chủ trì", "", "", "Chủ trì", "", "", "", "Phối hợp", "", "", "Làm thay"],
    "P7": ["Chủ trì", "", "Phối hợp", "Phối hợp", "", "", "Phối hợp", "", "", "", "Trung gian chuyển tin – không tạo giá trị"],
    "M1": ["Phối hợp", "", "", "", "", "Chủ trì", "", "", "", "", "Làm thay (hậu cần)"],
}
TAC_DONG = {
    "I1": "Quản lý chuyền dựa vào ISO để phát hiện lỗi; phần lớn biện pháp là nhắc nhở và lỗi tái diễn theo cùng chủ đề (nón, bao tay, BTP chạm đất, rập) – xem ChatLuongKiemTra.",
    "I2": "Kết quả ghi 'Đạt' chung chung, không ghi số chuyền, số túi đã kiểm; không chứng minh được độ bao phủ. Vương còn kiểm bóp VDKL cuối ngày thay quản lý.",
    "I3": "Nhập 2 lần/ngày ở 3 xưởng; có thể dùng thiết bị ghi tự động hoặc giao người tại chỗ.",
    "I4": "Tài liệu ban hành hằng ngày là tài liệu sản xuất theo mã (list đổ hàng, TCC) – luồng Kỹ thuật / Kế hoạch. 3 người làm song song ở 3 xưởng.",
    "I5": "Báo cáo ngày chép mẫu 52–61% dòng; không có danh sách lỗi mở, hạn đóng, quá hạn.",
    "P1": "Chuyền trưởng không trực tiếp chịu trách nhiệm số đăng ký; phòng cắt nhận yêu cầu qua người trung gian; Đạo ghi ngày đăng ký '5-11' lặp sau ngày đó.",
    "P2": "Cắt, in, thêu nhận hối thúc từ ISO song song với Kế hoạch; không rõ ai quyết thứ tự ưu tiên.",
    "P3": "Thêm một bước trung chuyển phiếu giữa chuyền và kho.",
    "P4": "ISO 'kiểm tra phân chia kế hoạch, đẩy năng suất' – chồng lên quyền của Kế hoạch và quản đốc.",
    "P5": "Việc văn thư lặp lại theo mã hàng; nên in từ mẫu chuẩn.",
    "P6": "Chuyền chỉnh máy phải chờ ISO ứng vải thay vì nhận trực tiếp.",
    "P7": "Thông tin lỗi chuyền đi vòng qua ISO, chậm và mất dấu trách nhiệm.",
    "M1": "ISO cấp kệ chỉ, rổ, vải vệ sinh, cắt ra bàn ủi – việc hậu cần.",
}
header(wsD, 4, ["Mã", "Công việc"] + BP + ["Hệ quả / bằng chứng", "Phút/ngày (3 người)"],
       [6, 36] + [11] * 10 + [26, 50, 11])
r = 5
for ma in VIEC:
    put(wsD, r, 1, ma, bold=True, align=CENTER)
    put(wsD, r, 2, VIEC[ma][0])
    for j, v in enumerate(MT[ma]):
        bg = None
        if v == "Chủ trì":
            bg = "D5E8D4"
        elif v == "Phối hợp":
            bg = "F2F4F8"
        elif j == len(BP) - 1:
            bg = HUONG_FILL["BỎ"] if ("Làm thay" in v or "Trung gian" in v) else HUONG_FILL["GIỮ"]
        put(wsD, r, 3 + j, v, bg=bg, align=CENTER)
    put(wsD, r, 3 + len(BP), TAC_DONG[ma])
    tr = T_FIRST + list(VIEC).index(ma)
    put(wsD, r, 4 + len(BP), f"=ThoiGian!N{tr}*60", fmt="0", align=CENTER)
    r += 1
wsD.freeze_panes = "C5"

# --------------------------------------------------------------------------- GiuChuyenBo
wsG = wb.create_sheet("GiuChuyenBo", 2)
title(wsG, "Bảng việc GIỮ / CHUYỂN / BỎ – đề xuất để chốt với chị Miền",
      "Cột 'Quyết định chốt' và 'Ngày bàn giao' để điền khi họp. Phút/ngày lấy từ sheet ThoiGian.")
GH = ["Mã", "Công việc", "Đề xuất", "Phút/ngày hiện tại (3 người)", "Người / bộ phận nhận", "Quyền quyết định",
      "Đầu ra bắt buộc", "Tiêu chí nghiệm thu", "Cần xác minh trước khi chốt", "Quyết định chốt", "Ngày bàn giao"]
header(wsG, 4, GH, [6, 34, 16, 11, 30, 30, 32, 40, 34, 14, 12])
GCB = {
    "I1": ("Quản lý chuyền kiểm tra hằng ngày theo checklist; quản đốc chịu trách nhiệm kết quả. ISO đánh giá xác suất theo lịch xoay vòng 3 xưởng.",
           "Quản đốc xử lý vi phạm. ISO có quyền mở lỗi, đặt hạn khắc phục, báo cáo lỗi quá hạn lên ban giám đốc.",
           "Checklist 5S/BHLĐ có ký mỗi ca; phiếu đánh giá xác suất của ISO; danh sách lỗi mở (khu vực, chủ đề, hạn, người đóng).",
           "100% chuyền có checklist ký mỗi ca. ISO đánh giá mỗi chuyền ≥ 1 lần/tuần. 100% lỗi có hạn và bằng chứng đóng. Lỗi tái diễn cùng chuyền + chủ đề giảm sau 2 tuần.",
           "Yêu cầu tần suất kiểm tra của khách hàng (Toray) và bộ tiêu chuẩn hiện hành."),
    "I2": ("Quản lý chuyền cấp phát, thu hồi, ký bảng đánh chéo hằng ngày và kiểm bóp VDKL cuối ca. ISO xác minh xác suất; kiểm 100% khi kim gãy, mất VDKL, thay mã.",
           "ISO có quyền dừng chuyền khi thiếu VDKL chưa tìm thấy (theo quy định khách hàng). Quản đốc xử lý vi phạm.",
           "Bảng đánh chéo ký đủ; phiếu xác minh ghi số chuyền, số túi kiểm, số lệch, xử lý lệch.",
           "100% bảng ký đủ mỗi ngày. Mỗi lượt xác minh ghi rõ chuyền và số túi kiểm, không ghi 'Đạt' chung. Lệch được đóng trong ngày.",
           "Quy định VDKL của khách hàng: tần suất xác minh tối thiểu, ai được ký."),
    "I3": ("Tổ mẫu / thống kê chuyền nhập tại chỗ, hoặc lắp thiết bị ghi tự động.", "Quản đốc phân công người nhập.",
           "Số liệu nhiệt độ, độ ẩm đủ các lần đo/ngày.", "Không thiếu lần đo trong 10 ngày. ISO xem lại hằng tuần, xử lý khi vượt ngưỡng.",
           "Ngưỡng nhiệt độ, độ ẩm và tần suất đo theo yêu cầu."),
    "I4": ("Tài liệu sản xuất theo mã (list đổ hàng, TCC…): Kế hoạch / Kỹ thuật phân phối cho 3 xưởng. ISO chỉ kiểm soát tài liệu hệ thống và quy tắc phân phối.",
           "ISO duyệt ban hành, thu hồi, hủy tài liệu hệ thống. Kỹ thuật duyệt nội dung tài liệu sản xuất.",
           "Danh mục tài liệu hiệu lực theo mã hàng; sổ phân phối, thu hồi, hủy.",
           "Kiểm xác suất tại chuyền: 0 tài liệu hết hiệu lực. Tài liệu đến chuyền trước ngày vào chuyền.",
           "Số tài liệu ban hành/ngày ở mỗi xưởng (đếm trong 10 ngày thử)."),
    "I5": ("ISO.", "Trưởng ISO.", "Danh sách lỗi mở dùng chung; báo cáo tuần thay báo cáo ngày theo mẫu.",
           "Báo cáo tuần có: số lỗi mới, đã đóng, quá hạn, tái diễn theo xưởng và chủ đề.", ""),
    "P1": ("Chuyền trưởng đăng ký BTP và xác nhận nhận. Kế hoạch chốt ngày, lượng, thứ tự. Phòng cắt cập nhật thực tế trên cùng bảng.",
           "Kế hoạch duyệt ngoại lệ (cấp sớm, cấp trễ, đổi thứ tự).", "Một bảng đăng ký – cấp – nhận BTP theo mã, chuyền, ngày.",
           "100% lượt cấp có đăng ký của chuyền trưởng và xác nhận nhận. Không còn đăng ký qua ISO.",
           "Đối chiếu hồ sơ cấp BTP của một số mã (sheet DoiChieuBTP) với phản hồi của nhóm kỹ thuật."),
    "P2": ("Kế hoạch / điều độ theo dõi đồng bộ cắt – in – thêu – laze – keo theo thứ tự đã chốt.",
           "Kế hoạch quyết thứ tự ưu tiên; cắt, in, thêu chỉ nhận yêu cầu từ Kế hoạch.", "Bảng đồng bộ đầu vào theo chuyền, cập nhật trong ngày.",
           "Chuyền không thiếu vốn do thiếu đồng bộ; mọi yêu cầu hối thúc đi qua một đầu mối.",
           "Kế hoạch có đủ người để nhận việc này không (≈ phút/ngày ở cột D)."),
    "P3": ("Chuyền trưởng đăng ký trực tiếp với Kho phụ liệu.", "Kho phụ liệu cấp theo định mức; quản đốc duyệt vượt định mức.",
           "Phiếu đăng ký – cấp – nhận phụ liệu.", "Không còn phiếu đi qua ISO; không phát sinh thiếu phụ liệu do chuyển phiếu.", ""),
    "P4": ("Quản đốc / chuyền trưởng chịu trách nhiệm năng suất. Kế hoạch phân chuyền.", "Kế hoạch và quản đốc.",
           "Bảng năng suất giờ của chuyền; kế hoạch chuyền.", "ISO không còn ghi việc phân chia kế hoạch, đẩy năng suất.",
           "Quyền ISO hiện được giao: báo thiếu, yêu cầu cấp, đổi thứ tự (cần xác nhận với quản đốc A3)."),
    "P5": ("Thống kê chuyền / Kế hoạch in từ mẫu chuẩn theo mã hàng. ISO chỉ ban hành mẫu list đánh chéo một lần.",
           "Kế hoạch.", "List đánh chéo, list giấy số, bảng năng suất có sẵn trước ngày vào chuyền.",
           "Chuyền có đủ biểu mẫu trước khi vào chuyền; ISO không còn đánh máy, in, cấp phát.", ""),
    "P6": ("Kỹ thuật chuyền / chuyền trưởng nhận vải thử máy trực tiếp từ phòng cắt.", "Kế hoạch (thời điểm cấp).",
           "Phiếu ứng vải thử máy theo mã mới.", "Vải thử máy có trước ngày chỉnh máy; không qua ISO.", ""),
    "P7": ("BỎ vai trò trung gian. Chuyền trưởng báo thẳng bộ phận xử lý qua một kênh báo lỗi chung có người nhận và hạn phản hồi.",
           "Quản đốc quy định kênh và thời hạn phản hồi.", "Nhóm báo lỗi chung; mỗi loại lỗi có người nhận.",
           "Lỗi chuyền được bộ phận xử lý phản hồi trong thời hạn; ISO không còn chuyển tin.", ""),
    "M1": ("Kho / hành chính xưởng cấp theo đề nghị của quản lý chuyền.", "Quản đốc.", "Phiếu đề nghị cấp.",
           "ISO không còn cấp vật tư.", ""),
}
r = 5
for ma in VIEC:
    owner, quyen, dau_ra, tieu_chi, xac_minh = GCB[ma]
    tr = T_FIRST + list(VIEC).index(ma)
    vals = [ma, VIEC[ma][0], VIEC[ma][3], f"=ThoiGian!N{tr}*60", owner, quyen, dau_ra, tieu_chi, xac_minh, None, None]
    for c, v in enumerate(vals, 1):
        bg = HUONG_FILL[huong(ma)] if c == 3 else (INPUT_FILL if c in (10, 11) else None)
        put(wsG, r, c, v, bg=bg, bold=c in (1, 3), fmt="0" if c == 4 else ("dd/mm/yyyy" if c == 11 else None),
            align=CENTER if c in (1, 3, 4) else WRAP)
    r += 1
for extra in [
    ("N1", "Lập biên bản, hạ loại thi đua, đe dọa xử lý ABC đối với công nhân", "CHUYỂN",
     "ISO ghi nhận và đề xuất; quản đốc / nhân sự quyết định kỷ luật.", "Quản đốc / Nhân sự.",
     "Biên bản do quản đốc ký.", "ISO không tự lập biên bản kỷ luật.",
     "Quy chế thi đua: ai có quyền hạ loại."),
    ("N2", "Theo dõi sửa chữa cơ điện, cơ sở vật chất (ổ điện, đèn, kính, gạch, vòng chắn kim, vạch lối thoát hiểm)", "GIỮ – CHỈ KIỂM TRA ĐÓNG",
     "Cơ điện chịu trách nhiệm sửa theo phiếu. ISO chỉ ghi nhận, đặt hạn và kiểm tra đóng.", "Trưởng phòng cơ điện.",
     "Phiếu yêu cầu sửa chữa có hạn.", "Hạng mục an toàn (lối thoát hiểm, điện, vòng chắn kim) đóng trong hạn, không chờ lịch khách đến.",
     "Vạch lối thoát hiểm A3 ghi 'sẽ thực hiện sau khi khách gần tới' – cần đặt hạn ngay."),
]:
    ma, ten, de_xuat, owner, quyen, dau_ra, tieu_chi, xm = extra
    vals = [ma, ten, de_xuat, "chưa đo", owner, quyen, dau_ra, tieu_chi, xm, None, None]
    for c, v in enumerate(vals, 1):
        h = "GIỮ" if de_xuat.startswith("GIỮ") else "CHUYỂN"
        bg = HUONG_FILL[h] if c == 3 else (INPUT_FILL if c in (10, 11) else None)
        put(wsG, r, c, v, bg=bg, bold=c in (1, 3), align=CENTER if c in (1, 3, 4) else WRAP)
    r += 1
G_LAST = r - 1
dv = DataValidation(type="list", formula1='"Đồng ý,Sửa,Chưa chốt"', allow_blank=True)
wsG.add_data_validation(dv)
dv.add(f"J5:J{G_LAST}")
put(wsG, r + 1, 1, "Ô nền vàng: điền khi họp. Đề xuất dựa trên 3 báo cáo tự ghi; cột I là những điểm phải xác minh trước khi chốt.",
    color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
wsG.freeze_panes = "C5"

# --------------------------------------------------------------------------- LoiHienTruong
wsF = wb.create_sheet("LoiHienTruong")
title(wsF, "Lỗi hiện trường ghi trong báo cáo (5S, nội quy, BHLĐ, VDKL)",
      "Phân loại chủ đề, biện pháp, trạng thái theo từ khóa trên nội dung gốc; có thể sửa tay ở cột F–I.")
FH = ["Mã người", "Người", "Ngày", "Khu vực (gốc)", "Khu vực chuẩn hóa", "Loại dòng", "Chủ đề", "Kiểu biện pháp",
      "Trạng thái ghi", "Số lần cùng người + khu vực + chủ đề", "Nội dung", "Biện pháp", "Kết quả", "Ô nguồn"]
header(wsF, 4, FH, [9, 18, 10, 16, 12, 14, 22, 22, 16, 12, 48, 40, 18, 22])
r = 5
F_FIRST = 5
F_LAST = F_FIRST + len(loi) - 1
for z in loi:
    x = z["x"]
    vals = [x["nguoi"], NGUON[x["nguoi"]][1], x["ngay"], x["khu_vuc"], z["kv"], z["loai"], z["chu_de"], z["bp"],
            z["tt"], f'=IF(F{r}="Lỗi",COUNTIFS($A${F_FIRST}:$A${F_LAST},A{r},$E${F_FIRST}:$E${F_LAST},E{r},'
                     f'$G${F_FIRST}:$G${F_LAST},G{r},$F${F_FIRST}:$F${F_LAST},"Lỗi"),"")',
            x["noi_dung"], x["bien_phap"], x["ket_qua"], f"{NGUON[x['nguoi']][0].split('/')[-1]}!E{x['dong']}"]
    for c, v in enumerate(vals, 1):
        put(wsF, r, c, v, fmt="dd/mm/yyyy" if c == 3 else None, align=CENTER if c == 10 else WRAP)
    r += 1
wsF.freeze_panes = "A5"
wsF.auto_filter.ref = f"A4:N{F_LAST}"


def frng(col):
    return f"LoiHienTruong!${col}${F_FIRST}:${col}${F_LAST}"


# --------------------------------------------------------------------------- ChatLuong
wsQ = wb.create_sheet("ChatLuongKiemTra", 3)
title(wsQ, "Chất lượng kiểm tra: phát hiện xong thì xử lý thế nào?",
      "Đếm từ sheet LoiHienTruong (chỉ dòng 'Lỗi' trừ khi ghi khác).")
put(wsQ, 4, 1, "Chỉ số", bold=True, bg=HEAD)
for i, p in enumerate(NGUOI):
    put(wsQ, 4, 2 + i, NGUON[p][1], bold=True, bg=HEAD, align=CENTER)
put(wsQ, 4, 5, "Cả 3", bold=True, bg=HEAD, align=CENTER)
wsQ.column_dimensions["A"].width = 52
for col in "BCDE":
    wsQ.column_dimensions[col].width = 16
r = 5
Q = {}


def qrow(label, formula_for, fmt="0", key=None):
    global r
    put(wsQ, r, 1, label)
    for i, p in enumerate(NGUOI):
        put(wsQ, r, 2 + i, formula_for(p), fmt=fmt, align=CENTER)
    put(wsQ, r, 5, formula_for(None), fmt=fmt, align=CENTER)
    if key:
        Q[key] = r
    r += 1


def cond(p):
    return f',{frng("A")},"{p}"' if p else ""


qrow("Số lỗi ghi nhận", lambda p: f'=COUNTIFS({frng("F")},"Lỗi"{cond(p)})', key="n")
qrow("Số dòng kiểm tra lại lỗi của ngày trước", lambda p: f'=COUNTIFS({frng("F")},"Kiểm tra lại lỗi cũ"{cond(p)})', key="kl")
r += 1
put(wsQ, r, 1, "Kiểu biện pháp (tỷ lệ trên số lỗi)", bold=True, color=ACCENT, border=False)
r += 1
for bp in ["Nhắc nhở / yêu cầu chung", "Quản lý tự khắc phục, gửi ảnh", "Chuyển bộ phận khác xử lý",
           "ISO tự xử lý / cấp vật tư", "Kỷ luật / đe dọa xử lý ABC", "Không ghi biện pháp"]:
    col = None
    qrow(bp, lambda p, bp=bp: f'=IFERROR(COUNTIFS({frng("F")},"Lỗi",{frng("H")},"{bp}"{cond(p)})/'
                              f'COUNTIFS({frng("F")},"Lỗi"{cond(p)}),0)', fmt="0%", key="bp_" + bp)
r += 1
put(wsQ, r, 1, "Trạng thái ghi trong báo cáo", bold=True, color=ACCENT, border=False)
r += 1
for tt in ["Báo đạt trong ngày", "Hẹn kiểm tra lại", "Chưa đạt", "Không ghi kết quả"]:
    qrow(tt, lambda p, tt=tt: f'=IFERROR(COUNTIFS({frng("F")},"Lỗi",{frng("I")},"{tt}"{cond(p)})/'
                              f'COUNTIFS({frng("F")},"Lỗi"{cond(p)}),0)', fmt="0%", key="tt_" + tt)
r += 1
put(wsQ, r, 1, "Số lỗi theo chủ đề (lỗi lặp lại = hệ thống chưa sửa gốc)", bold=True, color=ACCENT, border=False)
r += 1
for ten, _ in CHU_DE + [("Khác", "")]:
    qrow(ten, lambda p, ten=ten: f'=COUNTIFS({frng("F")},"Lỗi",{frng("G")},"{ten}"{cond(p)})')
r += 1
qrow("Số lỗi thuộc cặp khu vực + chủ đề bị ghi ≥ 2 lần",
     lambda p: f'=COUNTIFS({frng("J")},">=2"{cond(p)})', key="lap")
qrow("Tỷ lệ lỗi tái diễn trên tổng số lỗi", lambda p: (
    f'=IFERROR(COUNTIFS({frng("J")},">=2"{cond(p)})/COUNTIFS({frng("F")},"Lỗi"{cond(p)}),0)'), fmt="0%", key="lap_pct")
r += 1
for t in [
    "Đọc kết quả: phần lớn lỗi được xử lý bằng nhắc nhở và báo 'đạt' ngay trong ngày; rất ít dòng ghi lại kết quả kiểm tra lần sau.",
    "Cùng một chủ đề (nón / bao tay, BTP chạm đất, rập, bàn ủi, vệ sinh) xuất hiện lặp lại ở nhiều chuyền và nhiều ngày:",
    "kiểm tra đang thay cho quản lý chuyền, không tạo thay đổi hệ thống (nguyên nhân, trách nhiệm, chỉ tiêu).",
]:
    put(wsQ, r, 1, t, color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
    r += 1

# --------------------------------------------------------------------------- NoiDungThucChat
L = get_column_letter
wsW = wb.create_sheet("NoiDungThucChat", 4)
title(wsW, "Khối lượng việc thực chất ước tính từ số lượt việc ghi trong báo cáo (không dùng giờ tự khai)",
      "Lượt/ngày đếm từ NhatKy và LoiHienTruong. Phút/lượt là giả định thấp – cao (ô vàng), sửa theo số đo thực tế.")
WH = ["Mã", "Đơn vị đếm", "Phút/lượt thấp", "Phút/lượt cao"] + \
     [f"Lượt/ngày {NGUON[p][1]}" for p in NGUOI] + \
     [f"Phút thấp {NGUON[p][1]}" for p in NGUOI] + [f"Phút cao {NGUON[p][1]}" for p in NGUOI]
header(wsW, 4, WH, [7, 44, 10, 10] + [12] * 9)
DON_VI = {
    "I1": ("Lượt tuần tra 1 xưởng (8–11 chuyền)", 20, 35), "I2": ("Lượt kiểm bảng / bóp VDKL", 15, 30),
    "I3": ("Lần nhập nhiệt độ, độ ẩm", 5, 10), "I4": ("Lượt ban hành / photo / đóng mộc tài liệu", 20, 40),
    "I5": ("Lần lập, gửi báo cáo", 10, 20), "P1": ("Lần kiểm vốn, đăng ký BTP", 15, 25),
    "P2": ("Lượt theo dõi cấp BTP / in / thêu / laze", 15, 30), "P3": ("Lần làm phiếu phụ liệu", 10, 20),
    "P4": ("Lần theo dõi năng suất, kế hoạch", 10, 20), "P5": ("Lượt làm list / bảng biểu", 15, 30),
    "P6": ("Lần ứng, nhận vải thử máy", 10, 20), "P7": ("Lượt đọc mail, chuyển tin", 15, 30),
    "M1": ("Lần cấp vật tư phụ trợ", 10, 20),
}
r = 5
W_FIRST = r
for ma in list(VIEC) + ["LOI"]:
    if ma == "LOI":
        dv_, lo, hi = "Ghi nhận 1 lỗi (xem, chụp ảnh, gửi nhóm, nhắc)", 3, 6
    else:
        dv_, lo, hi = DON_VI[ma]
    put(wsW, r, 1, ma, bold=True, align=CENTER)
    put(wsW, r, 2, dv_)
    put(wsW, r, 3, lo, bg=INPUT_FILL, align=CENTER, color="0000FF")
    put(wsW, r, 4, hi, bg=INPUT_FILL, align=CENTER, color="0000FF")
    for i, p in enumerate(NGUOI):
        days = f"ThoiGian!${L(PC[p])}$6"
        if ma == "LOI":
            fx = f'=COUNTIFS({frng("A")},"{p}",{frng("F")},"Lỗi")/{days}'
        else:
            fx = f'=COUNTIFS({rng("A")},"{p}",{rng("G")},$A{r},{rng("F")},">0")/{days}'
        put(wsW, r, 5 + i, fx, fmt="0.0", align=CENTER)
        put(wsW, r, 8 + i, f"={L(5 + i)}{r}*$C{r}", fmt="0", align=CENTER)
        put(wsW, r, 11 + i, f"={L(5 + i)}{r}*$D{r}", fmt="0", align=CENTER)
    r += 1
W_LAST = r - 1
W = {}
for key, label, fn in [
    ("lo", "Khối lượng thực chất – giả định thấp (phút/ngày)", lambda i: f"=SUM({L(8 + i)}{W_FIRST}:{L(8 + i)}{W_LAST})"),
    ("hi", "Khối lượng thực chất – giả định cao (phút/ngày)", lambda i: f"=SUM({L(11 + i)}{W_FIRST}:{L(11 + i)}{W_LAST})"),
    ("kh", "Giờ tự khai (phút/ngày)", lambda i: f"=ThoiGian!{L(PC[NGUOI[i]])}7*60"),
    ("pl", "Tỷ lệ lấp đầy giờ khai – thấp", None),
    ("ph", "Tỷ lệ lấp đầy giờ khai – cao", None),
]:
    wsW.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    put(wsW, r, 1, label, bold=True)
    for i in range(3):
        if key == "pl":
            fx, fmt = f"=IFERROR({L(5 + i)}{W['lo']}/{L(5 + i)}{W['kh']},0)", "0%"
        elif key == "ph":
            fx, fmt = f"=IFERROR({L(5 + i)}{W['hi']}/{L(5 + i)}{W['kh']},0)", "0%"
        else:
            fx, fmt = fn(i), "0"
        put(wsW, r, 5 + i, fx, fmt=fmt, align=CENTER, bold=True)
    W[key] = r
    r += 1
r += 1
for t in [
    "Đọc kết quả: với giả định cao, nội dung việc ghi được gần lấp đầy giờ khai; với giả định thấp chỉ khoảng một nửa.",
    "Vì vậy báo cáo KHÔNG chứng minh được nhân sự thiếu việc hay ngồi không. Dư người nằm ở chỗ làm thay việc của người khác "
    "và kiểm tra lặp lại nhiều lượt, không nằm ở thời gian trống.",
    "Muốn biết chính xác cần chụp ảnh làm việc (work sampling): ghi việc đang làm tại 20–30 thời điểm ngẫu nhiên mỗi ngày trong 10 ngày.",
]:
    put(wsW, r, 1, t, color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
    r += 1
wsW.freeze_panes = "C5"

# --------------------------------------------------------------------------- DinhBien
wsB = wb.create_sheet("DinhBien", 5)
title(wsB, "Định biên theo 3 kịch bản (ô vàng là giả định, sửa được)",
      "Tính số người cần cho đúng chức năng ISO và cho phần việc kế hoạch đang nằm trong 3 vị trí, rồi so với 3 người hiện có.")
for col, w in zip("ABCDEFGHIJK", [56, 10, 10, 10, 10, 10, 10, 12, 12, 12, 44]):
    wsB.column_dimensions[col].width = w
r = 4
put(wsB, r, 1, "Thông số chung", bold=True, bg=HEAD)
put(wsB, r, 2, "Giá trị", bold=True, bg=HEAD, align=CENTER)
r += 1
B = {}
for key, label, v, fmt in [("gio", "Giờ làm việc / người / ngày", 8, "0.0"),
                           ("hs", "Tỷ lệ thời gian dùng được cho công việc (trừ họp, phát sinh)", 0.85, "0%"),
                           ("hien_co", "Số người hiện có (ISO kiêm kế hoạch)", 3, "0"),
                           ("chuyen", "Số chuyền may", 30, "0")]:
    put(wsB, r, 1, label)
    put(wsB, r, 2, v, bg=INPUT_FILL, fmt=fmt, align=CENTER, color="0000FF")
    B[key] = f"$B${r}"
    r += 1
put(wsB, r, 1, "Phút hiệu dụng / người / ngày", bold=True)
put(wsB, r, 2, f"={B['gio']}*60*{B['hs']}", fmt="0", align=CENTER)
B["fte"] = f"$B${r}"
r += 2

SC = ["Thấp", "Cơ sở", "Cao"]
put(wsB, r, 1, "A. Việc ISO đúng chức năng (cả 3 xưởng)", bold=True, bg=HEAD)
for j, s in enumerate(SC):
    put(wsB, r, 2 + 2 * j, f"{s}: SL", bold=True, bg=HEAD, align=CENTER)
    put(wsB, r, 3 + 2 * j, f"{s}: phút", bold=True, bg=HEAD, align=CENTER)
for j, s in enumerate(SC):
    put(wsB, r, 8 + j, f"Phút/ngày {s.lower()}", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 11, "Căn cứ", bold=True, bg=HEAD)
r += 1
a_first = r
I4_ROW = T_FIRST + list(VIEC).index("I4")
for label, vals, can_cu in [
    ("Đánh giá xác suất 5S / BHLĐ / nội quy (chuyền/ngày × phút/chuyền)", [(6, 8), (10, 8), (15, 10)],
     "Quản lý chuyền tự kiểm hằng ngày; ISO xoay vòng mỗi chuyền 2–5 ngày/lần"),
    ("Xác minh VDKL xác suất (chuyền/ngày × phút/chuyền)", [(6, 6), (10, 6), (15, 8)],
     "Kiểm bảng ký + đếm một số túi; sự cố kim gãy kiểm 100% (đã gồm trong kịch bản cao)"),
    ("Di chuyển giữa 3 xưởng (lượt × phút)", [(2, 10), (2, 10), (3, 10)], ""),
    ("Theo dõi lỗi mở, kiểm tra tái diễn (lỗi/ngày × phút)", [(6, 5), (8, 5), (12, 6)], "Một danh sách lỗi mở chung"),
    ("Đánh giá nội bộ, đào tạo, làm việc khách hàng, báo cáo tuần (phút/ngày)", [(1, 45), (1, 60), (1, 90)], "Bình quân ngày"),
    ("Kiểm soát tài liệu hệ thống: phút I4 hiện nay × % ISO giữ lại",
     [(f"=ThoiGian!N{I4_ROW}*60", 0.2), (f"=ThoiGian!N{I4_ROW}*60", 0.3), (f"=ThoiGian!N{I4_ROW}*60", 0.5)],
     "Tài liệu sản xuất theo mã (list đổ hàng, TCC) do Kỹ thuật / Kế hoạch phân phối; ISO chỉ giữ tài liệu hệ thống"),
]:
    put(wsB, r, 1, label)
    for j, (q, m) in enumerate(vals):
        isf = isinstance(q, str)
        put(wsB, r, 2 + 2 * j, q, fmt="0", align=CENTER, bg=None if isf else INPUT_FILL, color="008000" if isf else "0000FF")
        put(wsB, r, 3 + 2 * j, m, fmt="0%" if isinstance(m, float) and m < 1 else "0", align=CENTER, bg=INPUT_FILL, color="0000FF")
        put(wsB, r, 8 + j, f"={L(2 + 2 * j)}{r}*{L(3 + 2 * j)}{r}", fmt="0", align=CENTER)
    put(wsB, r, 11, can_cu, color=MUTED)
    r += 1
B["share_row"] = r - 1
put(wsB, r, 1, "Tổng phút ISO / ngày", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"=SUM({L(8 + j)}{a_first}:{L(8 + j)}{r - 1})", fmt="0", align=CENTER, bold=True)
r += 1
put(wsB, r, 1, "Số người ISO cần (tương đương)", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"={L(8 + j)}{r - 1}/{B['fte']}", fmt="0.00", align=CENTER, bold=True, bg="E1F3E9")
B["iso_row"] = r
r += 2

put(wsB, r, 1, "B. Việc kế hoạch đang nằm trong 3 vị trí, cần người nhận tập trung", bold=True, bg=HEAD)
for j, s in enumerate(SC):
    put(wsB, r, 8 + j, f"Phút/ngày {s.lower()}", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 11, "Căn cứ", bold=True, bg=HEAD)
r += 1
b_first = r
for ma, nhan in [("P2", "Đồng bộ cắt – in – thêu – laze – keo"), ("P4", "Năng suất, phân chuyền"),
                 ("P5", "Biểu mẫu chuyền (list đánh chéo, giấy số, bảng năng suất)")]:
    put(wsB, r, 1, f"{ma} – {nhan}")
    for j in range(3):
        put(wsB, r, 8 + j, f"=ThoiGian!N{T_FIRST + list(VIEC).index(ma)}*60", fmt="0", align=CENTER, color="008000")
    put(wsB, r, 11, "Giờ tự khai hiện nay (giữ nguyên, có thể thừa)", color=MUTED)
    r += 1
put(wsB, r, 1, "Phân phối tài liệu sản xuất theo mã (phần I4 không còn ở ISO)")
for j in range(3):
    put(wsB, r, 8 + j, f"={L(2 + 2 * j)}{B['share_row']}*(1-{L(3 + 2 * j)}{B['share_row']})", fmt="0", align=CENTER)
r += 1
put(wsB, r, 1, "Tổng phút kế hoạch / ngày", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"=SUM({L(8 + j)}{b_first}:{L(8 + j)}{r - 1})", fmt="0", align=CENTER, bold=True)
r += 1
put(wsB, r, 1, "Số người kế hoạch cần thêm (tương đương)", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"={L(8 + j)}{r - 1}/{B['fte']}", fmt="0.00", align=CENTER, bold=True, bg="FFF1D6")
B["kh_row"] = r
r += 2

put(wsB, r, 1, "C. Việc trả về từng chuyền / bộ phận (chia nhỏ, không cần thêm người)", bold=True, bg=HEAD)
put(wsB, r, 8, "Phút/ngày", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 9, "Phút/chuyền/ngày", bold=True, bg=HEAD, align=CENTER)
r += 1
c_first = r
for ma, nhan in [("P1", "Chuyền trưởng: đăng ký BTP"), ("P3", "Chuyền trưởng ↔ kho: phụ liệu"),
                 ("P6", "Kỹ thuật chuyền: vải thử máy"), ("I3", "Tổ mẫu / thống kê: nhiệt độ, độ ẩm"),
                 ("M1", "Kho / hành chính: vật tư phụ trợ"), ("P7", "Bỏ trung gian: chuyền báo thẳng bộ phận xử lý")]:
    put(wsB, r, 1, f"{ma} – {nhan}")
    put(wsB, r, 8, f"=ThoiGian!N{T_FIRST + list(VIEC).index(ma)}*60", fmt="0", align=CENTER, color="008000")
    put(wsB, r, 9, f"=H{r}/{B['chuyen']}", fmt="0.0", align=CENTER)
    r += 1
put(wsB, r, 1, "Tổng", bold=True)
put(wsB, r, 8, f"=SUM(H{c_first}:H{r - 1})", fmt="0", align=CENTER, bold=True)
put(wsB, r, 9, f"=H{r}/{B['chuyen']}", fmt="0.0", align=CENTER, bold=True)
r += 2

put(wsB, r, 1, "D. Tổng nhu cầu so với hiện có", bold=True, bg=HEAD)
for j, s in enumerate(SC):
    put(wsB, r, 8 + j, s, bold=True, bg=HEAD, align=CENTER)
r += 1
put(wsB, r, 1, "Tổng người cần (ISO + kế hoạch)", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"={L(8 + j)}{B['iso_row']}+{L(8 + j)}{B['kh_row']}", fmt="0.00", align=CENTER, bold=True)
B["tong_row"] = r
r += 1
put(wsB, r, 1, "Dư so với số người hiện có", bold=True)
for j in range(3):
    put(wsB, r, 8 + j, f"={B['hien_co']}-{L(8 + j)}{B['tong_row']}", fmt="0.00", align=CENTER, bold=True, bg=HUONG_FILL["BỎ"])
B["du_row"] = r
r += 2

put(wsB, r, 1, "E. Phương án bố trí", bold=True, bg=HEAD)
put(wsB, r, 2, "Người ISO", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 3, "Người kế hoạch", bold=True, bg=HEAD, align=CENTER)
for j, s in enumerate(SC):
    put(wsB, r, 8 + j, f"Đủ? ({s.lower()})", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 11, "Ghi chú", bold=True, bg=HEAD)
r += 1
for nhan, n_iso, n_kh, note in [
    ("PA1 – Giữ 3 người ISO kiêm kế hoạch, mỗi xưởng 1 người", 3, 0, "Hiện trạng: dư người, ISO không độc lập, kênh chỉ đạo song song"),
    ("PA2 – 1 ISO chung 3 xưởng + 1 điều độ / tài liệu sản xuất thuộc Kế hoạch; điều chuyển 1 người", 1, 1,
     "Đề xuất: tách hai vai trò, báo cáo cho hai cấp khác nhau"),
    ("PA3 – 1 ISO chung; Kế hoạch hiện có tự nhận phần kế hoạch; điều chuyển 2 người", 1, 0,
     "Chỉ khi Kế hoạch xác nhận còn công suất ≈ số người ở mục B"),
]:
    put(wsB, r, 1, nhan)
    put(wsB, r, 2, n_iso, align=CENTER)
    put(wsB, r, 3, n_kh, align=CENTER)
    for j in range(3):
        iso_need = f"{L(8 + j)}${B['iso_row']}"
        kh_need = f"{L(8 + j)}${B['kh_row']}"
        if n_kh == 0 and n_iso == 3:
            fx = f'=IF($B{r}-{iso_need}-{kh_need}>0.05,"Dư "&TEXT($B{r}-{iso_need}-{kh_need},"0.0")&" người","Vừa đủ")'
        elif n_kh == 0:
            fx = f'=IF($B{r}>={iso_need},"ISO đủ; KH phải nhận "&TEXT({kh_need},"0.0"),"ISO thiếu")'
        else:
            fx = (f'=IF(AND($B{r}>={iso_need},$C{r}>={kh_need}),"Đủ",IF($B{r}+$C{r}>={iso_need}+{kh_need}-0.1,'
                  f'"Vừa đủ tổng: ISO thiếu "&TEXT({iso_need}-$B{r},"0.0")&", KH dư "&TEXT($C{r}-{kh_need},"0.0"),'
                  f'"Thiếu "&TEXT({iso_need}+{kh_need}-$B{r}-$C{r},"0.0")))')
        put(wsB, r, 8 + j, fx, align=CENTER)
    put(wsB, r, 11, note, color=MUTED)
    r += 1
r += 1
for t in [
    "Kịch bản cao đã cộng dồn mọi giả định bất lợi (đánh giá 15 chuyền/ngày, giữ 50% việc tài liệu ở ISO, giữ nguyên giờ tự khai của việc kế hoạch).",
    "Mục C: việc trả về chuyền chỉ vài phút mỗi chuyền mỗi ngày, không phải lý do để giữ người.",
]:
    put(wsB, r, 1, t, color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
    r += 1

# --------------------------------------------------------------------------- PhanBien
wsP = wb.create_sheet("PhanBien", 1)
title(wsP, "Phản biện bố trí hiện nay và đề xuất trong handoff",
      "Handoff chỉ dùng để tham khảo. Mỗi dòng: nhận định, phản biện, bằng chứng, ảnh hưởng tới quyết định.")
header(wsP, 4, ["#", "Nhận định / cách bố trí đang có", "Phản biện", "Bằng chứng", "Ảnh hưởng tới quyết định"],
       [5, 40, 56, 44, 40])
PB = [
    ("Bố trí hiện nay: mỗi xưởng một người 'ISO kiêm kế hoạch'.",
     "Chia người theo xưởng, không theo khối lượng: A3 chỉ 8 chuyền nhưng vẫn 1 người như xưởng 10–11 chuyền. Ghép hai chức năng "
     "ngược nhau – ISO phải kiểm tra độc lập, kế hoạch là việc vận hành – vào một người nằm trong xưởng.",
     "Chuyền nhắc trong báo cáo: Thân 1A–6A (11), Đạo 6B–11B (10), Vương 12A–15B (8).",
     "Bỏ mô hình 1 người/xưởng. Tách vai trò ISO và vai trò kế hoạch."),
    ("ISO tự làm việc vận hành rồi tự kiểm.",
     "Người kiểm đang tự làm việc mình kiểm: tự kiểm bảng đánh chéo, bóp VDKL thay quản lý chuyền, tự đăng ký BTP, tự cấp vật tư. "
     "Như vậy không còn kiểm tra độc lập (ISO 9001 mục 9.2.2 c yêu cầu đánh giá khách quan, vô tư).",
     "ThoiGian: I2, P1, P5, M1. Vương kiểm bóp VDKL 16h30–17h mỗi ngày.",
     "ISO phải trực thuộc cấp trên xưởng và không làm việc vận hành."),
    ("Kiểm tra 5S / nội quy 2–4 lượt mỗi ngày.",
     "Đây là giám sát thay quản lý chuyền. Kết quả: lỗi báo 'đạt' trong ngày nhưng lặp lại theo cùng chủ đề – kiểm nhiều không làm hệ thống tốt lên.",
     "ChatLuongKiemTra: nhắc nhở 60%, báo đạt trong ngày 84%, tái diễn 46%.",
     "Giảm còn đánh giá xác suất; quản lý chuyền chịu trách nhiệm hằng ngày có chỉ tiêu."),
    ("Kiểm VDKL xác suất mỗi ngày, kết quả 'Đạt'.",
     "Không ghi số chuyền, số túi đã kiểm, số lệch. 20 ngày công không phát hiện lệch số lượng nào: hoặc kiểm soát rất tốt, hoặc kiểm hình thức. "
     "Không có số liệu thì không chứng minh được với khách hàng.",
     "NhatKy, mã I2: kết quả 'ĐẠT' / 'Đạt' / để trống, không có số lượng.",
     "Bắt buộc ghi số lượng trong phiếu xác minh."),
    ("Nhập nhiệt độ, độ ẩm hằng ngày.",
     "Chuyền 14A ghi đúng 26°C / 73% cả 5 ngày. Cần kiểm tra thiết bị hoặc cách ghi. Việc này nên tự động hoặc giao người tại chỗ.",
     "Báo cáo Vương 05, 06, 07, 09, 10/11.", "Chuyển khỏi ISO."),
    ("Báo cáo công việc hằng ngày.",
     "Phần lớn là mẫu chép lại, khung giờ cố định, có lỗi chép (đăng ký BTP 'cho ngày 5-11' ghi cả ngày 06/11). Báo cáo không dùng để quản lý được.",
     "GhiChuDuLieu: 52–61% dòng chép nguyên văn.",
     "Thay bằng danh sách lỗi mở và báo cáo tuần có chỉ số."),
    ("Handoff: 'Chưa chốt giảm người; thử 10 ngày rồi xem giữ 3, 2 hay 1'.",
     "Trong cả 3 kịch bản (kể cả kịch bản cao), tổng nhu cầu ISO + kế hoạch vẫn dưới 3 người. Phương án 'giữ 3' có thể loại ngay; "
     "chạy thử chỉ để chọn giữa 1 và 2.",
     "DinhBien mục D.",
     "Quyết định cấu trúc ngay, chạy thử để hiệu chỉnh."),
    ("Handoff: 'Một đầu mối ISO chung và một người liên hệ tại mỗi xưởng từ nhân sự sẵn có'.",
     "Nếu 'người liên hệ' lại là chính các nhân sự này thì vị trí cũ vẫn còn nguyên, chỉ đổi tên. Đầu mối tại xưởng phải là quản đốc / phó quản đốc, "
     "như một trách nhiệm có chỉ tiêu, không phải một người làm thay.",
     "TrungLap: I1, I2 – chủ trì là quản lý chuyền.",
     "Ghi rõ người liên hệ = quản đốc / phó quản đốc."),
    ("Handoff: bước tiếp theo là đối chiếu hồ sơ cấp BTP với phản hồi thực địa.",
     "Việc này giúp thiết kế quy trình BTP, nhưng không ảnh hưởng câu hỏi 'ISO có dư không'. Đăng ký BTP không thuộc ISO dù hồ sơ ra sao. "
     "Không nên để quyết định tổ chức chờ bước này.",
     "TrungLap P1, P2.", "Chạy song song, không phải điều kiện."),
    ("Handoff giữ 'ban hành tài liệu' là việc ISO.",
     "Tài liệu đang ban hành hằng ngày là tài liệu sản xuất theo mã hàng (list đổ hàng, TCC) – thuộc luồng Kỹ thuật / Kế hoạch. "
     "ISO chỉ cần kiểm soát tài liệu hệ thống (quy trình, biểu mẫu) và quy tắc phân phối.",
     "Báo cáo Vương 03/11: 'list đổ hàng, tcc, tài liệu liên quan'.",
     "Phần lớn việc tài liệu chuyển ra khỏi ISO (DinhBien: ISO giữ 20–50%)."),
    ("Handoff đo bằng tần suất ghi nhận (số ngày có việc).",
     "Tần suất không cho biết khối lượng hay giá trị, nên không đủ để định biên. Cần đo theo thời gian, số lượt và đầu ra.",
     "ThoiGian, NoiDungThucChat, DinhBien.", "Dùng mô hình 3 kịch bản."),
    ("Handoff: 'Duy trì đủ các lượt kiểm tra bắt buộc'.",
     "Chưa xác định lượt nào là bắt buộc. 4 lượt tuần tra mỗi ngày hiện nay là thói quen, chưa thấy căn cứ yêu cầu khách hàng.",
     "Không có danh mục yêu cầu trong hồ sơ.",
     "Lập danh mục yêu cầu bắt buộc của khách hàng (Toray) trước khi chốt tần suất."),
    ("Handoff chưa kiểm tra công suất nơi nhận việc.",
     "Phần kế hoạch đang nằm trong 3 vị trí ≈ 0,7–0,8 người. Nếu Kế hoạch không có người nhận, việc sẽ quay lại ISO hoặc rơi rớt.",
     "DinhBien mục B.", "Xác nhận công suất Kế hoạch trước khi rút người (PA2 hay PA3)."),
    ("Handoff chưa nêu rủi ro an toàn bị trì hoãn.",
     "Vạch lối thoát hiểm ghi 'sẽ thực hiện sau khi khách gần tới'; vòng chắn kim gãy 12A, 15A, 15B; ổ điện rớt ở nhiều chuyền; "
     "ba hạng mục ngày 06/11 ghi 'kiểm tra lại' nhưng các ngày sau không thấy theo dõi.",
     "LoiHienTruong; báo cáo Đạo ngày 06/11; Vương 10/11.", "Giao cơ điện có hạn sửa, không phụ thuộc lịch khách."),
    ("Gốc vấn đề.",
     "Ba vị trí này tồn tại vì quản đốc, chuyền trưởng chưa chịu trách nhiệm trực tiếp về BTP, 5S, VDKL. Rút người mà không giao trách nhiệm "
     "và chỉ tiêu cho tuyến chuyền thì kỷ luật hiện trường sẽ tụt.",
     "TrungLap: chủ trì đúng là quản lý chuyền / Kế hoạch.",
     "Giao trách nhiệm + chỉ tiêu cho tuyến chuyền trước, rút người sau."),
]
r = 5
for i, (a, b, c, d) in enumerate(PB, 1):
    for col, v in enumerate([i, a, b, c, d], 1):
        put(wsP, r, col, v, align=CENTER if col == 1 else WRAP, bold=col == 2)
    r += 1
r += 1
put(wsP, r, 1, "Gợi ý bố trí người – chỉ dựa trên báo cáo, cần quản đốc và chị Miền xác nhận", bold=True, color=ACCENT, border=False)
r += 1
header(wsP, r, ["", "Người", "Điểm thấy được trong báo cáo", "Điểm cần lưu ý", "Vai trò có thể phù hợp"])
r += 1
for row in [
    ("Thân", "79% thời gian là việc kiểu ISO. Là người duy nhất ghi rõ lỗi gốc từ ngày nào và hẹn giờ kiểm tra lại "
             "(5 dòng kiểm tra lại); 15% lỗi có hẹn kiểm tra lại.",
     "72% biện pháp là nhắc nhở, hay dùng đe dọa xử lý ABC. Lỗi tủ VDNH 3B báo 'cải thiện' 3 ngày liền rồi vẫn lặp.",
     "ISO chung 3 xưởng, nếu được đào tạo đánh giá xác suất và xử lý nguyên nhân."),
    ("Vương (A3)", "40% thời gian là điều phối kế hoạch: in / thêu / laze, năng suất, biểu mẫu, phụ liệu. Có biện pháp xử lý gốc "
                   "(cài nhắc trả rập tự động, đổi ra bàn ủi màu tối).",
     "Ghi 'phân chia kế hoạch chuyền' – cần làm rõ quyền. Nhiệt độ 14A ghi giống hệt 5 ngày. Ngày 10/11 ghi sai thành 2026-10-11.",
     "Điều độ / kế hoạch xưởng hoặc phân phối tài liệu sản xuất (PA2)."),
    ("Đạo", "42% lỗi do quản lý tự khắc phục và gửi ảnh – đẩy trách nhiệm về quản lý (đúng hướng).",
     "61% dòng chép nguyên văn; lỗi chép ngày '5-11'; 4 hạng mục hẹn kiểm tra lại không thấy theo dõi.",
     "Cần quản đốc nhận xét thêm trước khi bố trí; ưu điểm là đẩy trách nhiệm về quản lý."),
]:
    put(wsP, r, 1, "")
    for col, v in enumerate(row, 2):
        put(wsP, r, col, v, bold=col == 2)
    r += 1
wsP.freeze_panes = "C5"

# --------------------------------------------------------------------------- KetLuan
wsK = wb.create_sheet("KetLuan", 0)
title(wsK, "APIC – 3 nhân sự ISO kiêm kế hoạch: bố trí có hợp lý không, có dư người không?",
      "Nguồn: 3 báo cáo công việc hằng ngày 03–10/11/2026. Số liệu là công thức từ các sheet chi tiết.")
wsK.column_dimensions["A"].width = 70
for col in "BCDE":
    wsK.column_dimensions[col].width = 15
r = 4
put(wsK, r, 1, "Trả lời ngắn", bold=True, bg=HEAD)
for j, s in enumerate(SC):
    put(wsK, r, 2 + j, f"Kịch bản {s.lower()}", bold=True, bg=HEAD, align=CENTER)
r += 1
for label, row_key, fmt, bg in [("Số người ISO cần (đúng chức năng, cả 3 xưởng)", "iso_row", "0.00", None),
                                ("Số người kế hoạch cần thêm (phần việc đang nằm trong 3 vị trí)", "kh_row", "0.00", None),
                                ("Tổng người cần", "tong_row", "0.00", None),
                                ("Dư so với 3 người hiện có", "du_row", "0.00", HUONG_FILL["BỎ"])]:
    put(wsK, r, 1, label, bold=row_key in ("tong_row", "du_row"))
    for j in range(3):
        put(wsK, r, 2 + j, f"=DinhBien!{L(8 + j)}{B[row_key]}", fmt=fmt, align=CENTER, color="008000", bold=True, bg=bg)
    r += 1
r += 1
put(wsK, r, 1, "Chỉ số chính", bold=True, bg=HEAD)
for i, p in enumerate(NGUOI):
    put(wsK, r, 2 + i, NGUON[p][1], bold=True, bg=HEAD, align=CENTER)
put(wsK, r, 5, "Cả 3", bold=True, bg=HEAD, align=CENTER)
r += 1


def krow(label, fn, fmt, tot=None):
    global r
    put(wsK, r, 1, label)
    for i, p in enumerate(NGUOI):
        put(wsK, r, 2 + i, fn(p), fmt=fmt, align=CENTER, color="008000")
    put(wsK, r, 5, tot if tot else "", fmt=fmt, align=CENTER, color="008000")
    r += 1


krow("Số ngày có báo cáo", lambda p: f"=ThoiGian!{L(PC[p])}6", "0",
     "=" + "+".join(f"ThoiGian!{L(PC[p])}6" for p in NGUOI))
krow("Giờ tự ghi / ngày", lambda p: f"=ThoiGian!{L(PC[p])}7", "0.0")
krow("% thời gian làm việc của bộ phận khác (nên CHUYỂN)", lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_SUM['CHUYỂN']}", "0%")
krow("% thời gian trung gian chuyển tin (nên BỎ)", lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_SUM['BỎ']}", "0%")
krow("% thời gian tuần tra 5S + kiểm VDKL (làm thay kiểm tra hằng ngày của quản lý chuyền)",
     lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_I12}", "0%")
krow("Nội dung việc ghi được lấp đầy giờ khai – giả định thấp", lambda p: f"=NoiDungThucChat!{L(5 + NGUOI.index(p))}{W['pl']}", "0%")
krow("Nội dung việc ghi được lấp đầy giờ khai – giả định cao", lambda p: f"=NoiDungThucChat!{L(5 + NGUOI.index(p))}{W['ph']}", "0%")
krow("Số lỗi hiện trường ghi nhận", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['n']}", "0",
     f"=ChatLuongKiemTra!E{Q['n']}")
krow("% lỗi xử lý bằng nhắc nhở chung", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['bp_Nhắc nhở / yêu cầu chung']}",
     "0%", f"=ChatLuongKiemTra!E{Q['bp_Nhắc nhở / yêu cầu chung']}")
krow("% lỗi báo 'đạt' ngay trong ngày", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['tt_Báo đạt trong ngày']}",
     "0%", f"=ChatLuongKiemTra!E{Q['tt_Báo đạt trong ngày']}")
krow("% lỗi tái diễn (cùng khu vực + chủ đề ≥ 2 lần)", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['lap_pct']}",
     "0%", f"=ChatLuongKiemTra!E{Q['lap_pct']}")
r += 1
put(wsK, r, 1, "Kết luận", bold=True, color=ACCENT, border=False, size=12)
r += 1
KL = [
    "1. CÓ DƯ NGƯỜI. Ngay cả kịch bản cao (mọi giả định bất lợi), tổng nhu cầu cho đúng chức năng ISO cộng phần việc kế hoạch đang nằm trong "
    "3 vị trí vẫn chỉ khoảng 2 người; kịch bản cơ sở khoảng 1,5 người. Ba người là dư khoảng 1 người (kịch bản cao) đến 1,5 người (cơ sở).",
    "2. Dư không phải vì ngồi không – báo cáo vẫn kín giờ và nội dung việc ghi được lấp khoảng 50–95% giờ khai tùy giả định. Dư vì LÀM THAY: "
    "tuần tra 2–4 lượt/ngày thay quản lý chuyền, đăng ký BTP thay chuyền trưởng, hối thúc cắt – in – thêu thay Kế hoạch, gom phiếu thay kho, "
    "chuyển tin thay kênh báo lỗi. Bỏ phần làm thay thì công việc ISO còn lại nhỏ.",
    "3. BỐ TRÍ CHƯA HỢP LÝ về nguyên tắc: (a) chia người theo xưởng chứ không theo khối lượng (A3 chỉ 8 chuyền); (b) ghép việc kiểm tra độc lập "
    "với việc vận hành vào một người nằm trong xưởng – tự làm rồi tự kiểm; (c) không có mục tiêu, chỉ tiêu, danh sách lỗi mở, hạn đóng.",
    "4. KIỂM TRA KÉM HIỆU QUẢ: 60% lỗi chỉ nhắc nhở, 84% báo đạt trong ngày nhưng 46% lặp lại; kiểm VDKL luôn 'Đạt' không có số lượng; "
    "nhiệt độ 14A ghi y hệt 5 ngày; 52–61% dòng báo cáo chép lại. Đúng là 'thấy gì làm đó, thấy đâu nói đó'.",
    "5. ĐỀ XUẤT (PA2): 1 ISO chung cho 3 xưởng, trực thuộc cấp trên xưởng, chỉ đánh giá xác suất, xác minh VDKL, kiểm soát tài liệu hệ thống, "
    "theo dõi khắc phục. 1 người sang Kế hoạch làm điều độ đầu vào và phân phối tài liệu sản xuất cho 3 xưởng. 1 người điều chuyển. "
    "Nếu Kế hoạch xác nhận tự nhận được phần kế hoạch thì PA3: chỉ giữ 1.",
    "6. ĐIỀU KIỆN: giao trách nhiệm và chỉ tiêu cho quản đốc, chuyền trưởng (5S, VDKL hằng ngày, đăng ký BTP) trước khi rút người. "
    "Rút người trước khi giao trách nhiệm thì kỷ luật hiện trường sẽ tụt.",
    "7. KHÁC HANDOFF: không cần chờ thử 10 ngày mới loại phương án giữ 3; đối chiếu hồ sơ BTP không phải điều kiện của quyết định định biên; "
    "'người liên hệ mỗi xưởng' phải là quản đốc, không phải các nhân sự này. Xem sheet PhanBien.",
    "Giới hạn: giờ là giờ tự khai; phút/lượt và quy mô đánh giá là giả định (ô vàng ở DinhBien, NoiDungThucChat). "
    "Gợi ý bố trí từng người chỉ dựa trên báo cáo, cần quản đốc xác nhận.",
]
for t in KL:
    wsK.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    put(wsK, r, 1, t, border=False, align=Alignment(wrap_text=True, vertical="top"))
    wsK.row_dimensions[r].height = 46
    r += 1


# --------------------------------------------------------------------------- DoiChieuBTP
wsC = wb.create_sheet("DoiChieuBTP")
title(wsC, "Mẫu đối chiếu cấp BTP theo mã hàng với phản hồi thực địa của nhóm kỹ thuật",
      "Điền ô nền vàng. Dòng 5 là VÍ DỤ minh họa cách điền – xóa khi nhập số liệu thật. Cột tính toán tự động.")
CH = ["Mã hàng", "Chuyền", "Xưởng", "Ngày chuyền đăng ký", "SL đăng ký", "Ngày Kế hoạch chốt", "Ngày cắt cấp", "SL cấp",
      "Ngày vào chuyền thực tế", "Người phản hồi (kỹ thuật)", "Phản hồi thực địa", "Cấp sớm (ngày)",
      "Lệch SL", "Kết luận tự động", "Nguồn hồ sơ"]
header(wsC, 4, CH, [12, 8, 8, 12, 10, 12, 12, 10, 12, 16, 30, 10, 10, 26, 22])
ex = ["VÍ DỤ-001", "12A", "A3", "=DATE(2026,11,4)", 500, "=DATE(2026,11,4)", "=DATE(2026,11,5)", 450, "=DATE(2026,11,9)",
      "(tên kỹ thuật)", "Cấp sớm, BTP để lâu ở đầu chuyền", None, None, None, "Sổ cắt / phiếu cấp"]
for rr in range(5, 31):
    for c in range(1, 16):
        v = ex[c - 1] if rr == 5 else None
        if c == 12:
            v = f'=IF(OR(G{rr}="",I{rr}=""),"",I{rr}-G{rr})'
        elif c == 13:
            v = f'=IF(OR(E{rr}="",H{rr}=""),"",H{rr}-E{rr})'
        elif c == 14:
            v = (f'=IF(L{rr}="","",IF(L{rr}>2,"Cấp sớm >2 ngày",IF(L{rr}<0,"Vào chuyền trước khi cấp đủ – kiểm tra",'
                 f'IF(AND(M{rr}<>"",M{rr}<0),"Cấp thiếu so với đăng ký","Đúng hạn"))))')
        inp = c not in (12, 13, 14)
        put(wsC, rr, c, v, bg=INPUT_FILL if inp else None,
            fmt="dd/mm/yyyy" if c in (4, 6, 7, 9) else ("+0;-0;0" if c in (12, 13) else None),
            italic=(rr == 5), color=MUTED if rr == 5 and inp else INK)
put(wsC, 32, 1, "Ngưỡng 'cấp sớm' = 2 ngày là giả định để lọc; chỉnh trong công thức cột N theo quy định của Kế hoạch.",
    color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
wsC.freeze_panes = "B5"

# --------------------------------------------------------------------------- GhiChuDuLieu
wsN = wb.create_sheet("GhiChuDuLieu")
title(wsN, "Ghi chú dữ liệu, giả định và điểm khác với handoff")
wsN.column_dimensions["A"].width = 130
notes = [
    "Kỳ dữ liệu: file nhận ngày 10/11 có thêm ngày 10/11 (Vương, Đạo). Thân: 03,04,05,06,07,09/11. Vương và Đạo: 03,04,05,06,07,09,10/11. "
    "Ngày 08/11 là Chủ nhật.",
    "Handoff ghi 17 ngày công trong kỳ 03–09/11; bản file hiện tại có 18 ngày trong kỳ đó (mỗi báo cáo 6 ngày) và 20 ngày nếu tính cả 10/11.",
    "Đạo ghi 'ứng/nhận vải thử máy' 4 ngày (04, 06, 07, 09/11) chứ không phải 3/5 như handoff.",
    "Đạo ghi 'kiểm tra vốn chuyền để đăng ký BTP cho ngày 5-11' ở các ngày 03, 04, 05 và 06/11 – ngày 06/11 vẫn ghi cho ngày 05/11, "
    "dấu hiệu chép mẫu. Từ 07/11 ghi 9-11, 10-11, 'hôm sau'.",
    "Vương ghi ngày 10/11 thành 2026-10-11 (đảo ngày tháng); đã hiểu là 10/11.",
    "Báo cáo không ghi tên người trong file. Báo cáo thứ ba (file NỘI DUNG KIỂM TRA) là của Đạo, theo xác nhận ngày 10/11; chuyền nhắc tới là 6B–11B và có nhắc 'cầu qua Xí nghiệp A1', "
    "nên không phải A1. Thân nhắc chuyền 1A–6A. Vương ghi Xí nghiệp A3, chuyền 12A–15B.",
    "Tỷ lệ dòng chép nguyên văn (cùng hạng mục + nội dung + biện pháp xuất hiện ở ≥ 2 ngày): Thân 54%, Vương 52%, Đạo 61%. "
    "Khung giờ: Thân và Đạo dùng một bộ khung duy nhất cho mọi ngày; Vương 3 bộ trong 7 ngày.",
    "Phân bổ phút: khung giờ chia đều cho các mã việc có trong khung. Ví dụ khung 15h45–16h30 của Thân có 4 việc → mỗi việc ≈ 11 phút.",
    "Phân loại mã việc theo từ khóa (scripts/iso_nhat_ky.py). Phân loại chủ đề, kiểu biện pháp, trạng thái lỗi theo từ khóa "
    "(scripts/xuat_phan_tich_iso.py); có thể sửa tay trong sheet LoiHienTruong, các bảng thống kê tự cập nhật.",
    "Tạo lại file: python3 scripts/xuat_phan_tich_iso.py",
]
for i, t in enumerate(notes, 3):
    put(wsN, i, 1, f"• {t}", border=False, align=Alignment(wrap_text=True, vertical="top"))
    wsN.row_dimensions[i].height = 30

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print("Đã ghi", OUT, "| dòng nhật ký:", len(log), "| dòng lỗi:", len(loi))
