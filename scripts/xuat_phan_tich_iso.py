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
NGUOI = list(NGUON)  # THAN, VUONG, BC3


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
             ("VUONG", 177), ("VUONG", 189), ("BC3", 8), ("BC3", 18), ("BC3", 113)}
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
    "I4": ["", "", "Phối hợp", "", "", "", "", "Phối hợp", "", "", "Đúng chức năng nhưng 3 người làm song song ở 3 xưởng"],
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
    "I4": "Thân và báo cáo 3 ghi 2 khung giờ/ngày cho ban hành tài liệu; 3 đầu mối làm cùng một việc, khó kiểm soát phiên bản.",
    "I5": "Báo cáo ngày chép mẫu 52–61% dòng; không có danh sách lỗi mở, hạn đóng, quá hạn.",
    "P1": "Chuyền trưởng không trực tiếp chịu trách nhiệm số đăng ký; phòng cắt nhận yêu cầu qua người trung gian; báo cáo 3 ghi ngày đăng ký '5-11' lặp sau ngày đó.",
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
    "I4": ("Một đầu mối kiểm soát tài liệu chung cho 3 xưởng (thuộc ISO). Kỹ thuật / kế hoạch chịu trách nhiệm nội dung tài liệu.",
           "Trưởng ISO duyệt ban hành, thu hồi, hủy. Kỹ thuật duyệt nội dung kỹ thuật.",
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

# --------------------------------------------------------------------------- DinhBien
wsB = wb.create_sheet("DinhBien", 4)
title(wsB, "Mô hình định biên ISO sau sắp xếp (ô vàng là giả định, sửa được)",
      "Khối lượng ISO cần làm khi đã chuyển việc kế hoạch / văn thư / trung gian ra và đổi tuần tra sang đánh giá xác suất.")
for col, w in zip("ABCDEF", [58, 14, 14, 16, 16, 40]):
    wsB.column_dimensions[col].width = w
r = 4
put(wsB, r, 1, "Thông số chung", bold=True, bg=HEAD)
put(wsB, r, 2, "Giá trị", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 3, "Ghi chú", bold=True, bg=HEAD)
r += 1
B = {}
for key, label, v, fmt, note in [
    ("gio", "Giờ làm việc / người / ngày", 8, "0.0", "Giờ hành chính"),
    ("hs", "Tỷ lệ thời gian dùng được cho công việc", 0.85, "0%", "Trừ họp, di chuyển, việc phát sinh"),
    ("chuyen", "Số chuyền may", 30, "0", "Theo handoff: 30 chuyền, 3 xưởng"),
    ("xuong", "Số xưởng", 3, "0", ""),
]:
    put(wsB, r, 1, label)
    put(wsB, r, 2, v, bg=INPUT_FILL, fmt=fmt, align=CENTER, color="0000FF")
    put(wsB, r, 3, note, color=MUTED)
    wsB.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    B[key] = f"$B${r}"
    r += 1
put(wsB, r, 1, "Phút làm việc hiệu dụng / người / ngày", bold=True)
put(wsB, r, 2, f"={B['gio']}*60*{B['hs']}", fmt="0", align=CENTER)
B["phut_fte"] = f"$B${r}"
r += 2

put(wsB, r, 1, "A. Hiện trạng (giờ tự ghi/ngày, 3 người)", bold=True, bg=HEAD)
for c, t in enumerate(["Giờ/ngày", "Tương đương người"], 2):
    put(wsB, r, c, t, bold=True, bg=HEAD, align=CENTER)
r += 1
for h, nhan in [("GIỮ", "Việc ISO giữ lại (theo cách làm hiện nay)"), ("CHUYỂN", "Việc nên chuyển về bộ phận khác"),
                ("BỎ", "Việc nên bỏ (trung gian chuyển tin)")]:
    put(wsB, r, 1, nhan, bg=HUONG_FILL[h])
    put(wsB, r, 2, f"=ThoiGian!N{T_SUM[h]}", fmt="0.0", align=CENTER, color="008000")
    put(wsB, r, 3, f"=B{r}*60/{B['phut_fte']}", fmt="0.00", align=CENTER)
    B["a_" + h] = r
    r += 1
put(wsB, r, 1, "Tổng", bold=True)
put(wsB, r, 2, f"=SUM(B{r - 3}:B{r - 1})", fmt="0.0", align=CENTER, bold=True)
put(wsB, r, 3, f"=SUM(C{r - 3}:C{r - 1})", fmt="0.00", align=CENTER, bold=True)
r += 2

put(wsB, r, 1, "B. Khối lượng ISO sau sắp xếp (cho cả 3 xưởng)", bold=True, bg=HEAD)
for c, t in enumerate(["Số lượng / ngày", "Phút / đơn vị", "Phút / ngày", "", "Căn cứ"], 2):
    put(wsB, r, c, t, bold=True, bg=HEAD, align=CENTER)
r += 1
b_first = r
for label, sl, ph, can_cu in [
    ("Đánh giá xác suất 5S / BHLĐ / nội quy (số chuyền/ngày)", 10, 8, "Mỗi chuyền ≥ 1 lần/3 ngày; quản lý chuyền tự kiểm hằng ngày"),
    ("Xác minh VDKL xác suất (số chuyền/ngày)", 10, 6, "Kiểm bảng ký + một số túi; kiểm 100% khi có sự cố"),
    ("Di chuyển giữa 3 xưởng (lượt/ngày)", 2, 10, ""),
    ("Theo dõi lỗi mở, kiểm tra tái diễn (lỗi/ngày)", 8, 5, "Danh sách lỗi mở dùng chung"),
    ("Đánh giá nội bộ, đào tạo, làm việc khách hàng, báo cáo tuần (phút/ngày)", 1, 60, "Bình quân ngày"),
]:
    put(wsB, r, 1, label)
    put(wsB, r, 2, sl, bg=INPUT_FILL, align=CENTER, color="0000FF")
    put(wsB, r, 3, ph, bg=INPUT_FILL, align=CENTER, color="0000FF")
    put(wsB, r, 4, f"=B{r}*C{r}", fmt="0", align=CENTER)
    put(wsB, r, 6, can_cu, color=MUTED)
    r += 1
put(wsB, r, 1, "Kiểm soát tài liệu tập trung: phút hiện tại × (1 − % giảm khi gộp)")
put(wsB, r, 2, f"=ThoiGian!N{T_FIRST + list(VIEC).index('I4')}*60", fmt="0", align=CENTER, color="008000")
put(wsB, r, 3, 0.3, bg=INPUT_FILL, fmt="0%", align=CENTER, color="0000FF")
put(wsB, r, 4, f"=B{r}*(1-C{r})", fmt="0", align=CENTER)
put(wsB, r, 6, "Cột B = phút/ngày I4 hiện nay của 3 người; cột C = % giảm do một đầu mối", color=MUTED)
r += 1
put(wsB, r, 1, "Tổng phút ISO cần / ngày", bold=True)
put(wsB, r, 4, f"=SUM(D{b_first}:D{r - 1})", fmt="0", align=CENTER, bold=True)
B["iso_phut"] = f"$D${r}"
r += 1
put(wsB, r, 1, "Số người ISO chuyên trách cần (tương đương)", bold=True)
put(wsB, r, 4, f"={B['iso_phut']}/{B['phut_fte']}", fmt="0.00", align=CENTER, bold=True, bg="E1F3E9")
B["iso_fte"] = f"$D${r}"
r += 2

put(wsB, r, 1, "C. Việc chuyển đi phải có nơi nhận", bold=True, bg=HEAD)
put(wsB, r, 2, "Phút/ngày", bold=True, bg=HEAD, align=CENTER)
put(wsB, r, 3, "Tương đương người", bold=True, bg=HEAD, align=CENTER)
r += 1
c_first = r
for ma, nhan in [("P2", "Kế hoạch / điều độ: đồng bộ cắt – in – thêu – laze"),
                 ("P4", "Kế hoạch / quản đốc: năng suất, phân chuyền"),
                 ("P5", "Thống kê / kế hoạch: biểu mẫu chuyền"),
                 ("P1", "Chuyền trưởng (30 người): đăng ký BTP"),
                 ("P3", "Chuyền trưởng ↔ kho: phụ liệu"),
                 ("P6", "Kỹ thuật chuyền: vải thử máy"),
                 ("I3", "Tổ mẫu / thống kê: nhiệt độ, độ ẩm"),
                 ("M1", "Kho / hành chính: vật tư phụ trợ")]:
    put(wsB, r, 1, f"{ma} – {nhan}")
    put(wsB, r, 2, f"=ThoiGian!N{T_FIRST + list(VIEC).index(ma)}*60", fmt="0", align=CENTER, color="008000")
    put(wsB, r, 3, f"=B{r}/{B['phut_fte']}", fmt="0.00", align=CENTER)
    r += 1
put(wsB, r, 1, "Trong đó dồn về Kế hoạch / thống kê (P2 + P4 + P5)", bold=True)
put(wsB, r, 2, f"=SUM(B{c_first}:B{c_first + 2})", fmt="0", align=CENTER, bold=True)
put(wsB, r, 3, f"=B{r}/{B['phut_fte']}", fmt="0.00", align=CENTER, bold=True, bg="FFF1D6")
B["kh_fte"] = f"$C${r}"
r += 1
put(wsB, r, 1, "Tổng việc chuyển đi")
put(wsB, r, 2, f"=SUM(B{c_first}:B{r - 2})", fmt="0", align=CENTER)
put(wsB, r, 3, f"=B{r}/{B['phut_fte']}", fmt="0.00", align=CENTER)
r += 2

put(wsB, r, 1, "D. So sánh phương án số người ISO chuyên trách", bold=True, bg=HEAD)
for c, t in enumerate(["Số người", "Công suất (phút/ngày)", "Dư (+) / thiếu (−) so với nhu cầu", "Đánh giá"], 2):
    put(wsB, r, c, t, bold=True, bg=HEAD, align=CENTER)
r += 1
for n, nhan in [(3, "Giữ 3 (hiện trạng, 1 người/xưởng)"), (2, "Giảm còn 2"), (1, "Giảm còn 1 đầu mối chung")]:
    put(wsB, r, 1, nhan)
    put(wsB, r, 2, n, align=CENTER)
    put(wsB, r, 3, f"=B{r}*{B['phut_fte']}", fmt="0", align=CENTER)
    put(wsB, r, 4, f"=C{r}-{B['iso_phut']}", fmt='+0;-0;0', align=CENTER)
    put(wsB, r, 5, f'=IF(D{r}<0,"Thiếu – không đủ phủ",IF(D{r}>{B["phut_fte"]}*0.5,"Dư > nửa người","Vừa đủ"))',
        align=CENTER)
    r += 1
r += 1
for t in [
    "Cách dùng: sửa ô vàng theo số đo thực tế của 10 ngày thử nghiệm (số chuyền đánh giá/ngày, phút/chuyền, số lỗi mở, phút kiểm soát tài liệu).",
    "Mục C cho biết việc chuyển đi không biến mất: phần dồn về Kế hoạch / thống kê phải có người nhận, nếu không sẽ quay lại ISO.",
]:
    put(wsB, r, 1, t, color=MUTED, italic=True, border=False, align=Alignment(vertical="top"))
    r += 1

# --------------------------------------------------------------------------- KetLuan
wsK = wb.create_sheet("KetLuan", 0)
title(wsK, "APIC – Phân tích công việc 3 nhân sự ISO kiêm kế hoạch (báo cáo 03–10/11/2026)",
      "Nguồn: 3 báo cáo công việc hằng ngày (Thân, Vương – A3, báo cáo thứ ba chưa rõ tên). Số liệu là công thức từ các sheet chi tiết.")
wsK.column_dimensions["A"].width = 70
for col in "BCDE":
    wsK.column_dimensions[col].width = 15
r = 4
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


L = get_column_letter
krow("Số ngày có báo cáo", lambda p: f"=ThoiGian!{L(PC[p])}6", "0",
     "=" + "+".join(f"ThoiGian!{L(PC[p])}6" for p in NGUOI))
krow("Giờ tự ghi / ngày", lambda p: f"=ThoiGian!{L(PC[p])}7", "0.0")
krow("% thời gian cho việc nên CHUYỂN về bộ phận khác", lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_SUM['CHUYỂN']}", "0%")
krow("% thời gian cho việc nên BỎ (trung gian chuyển tin)", lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_SUM['BỎ']}", "0%")
krow("% thời gian tuần tra 5S + kiểm VDKL (trùng việc hằng ngày của quản lý chuyền)",
     lambda p: f"=ThoiGian!{L(PC[p] + 1)}{T_I12}", "0%")
krow("Số lỗi hiện trường ghi nhận", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['n']}", "0",
     f"=ChatLuongKiemTra!E{Q['n']}")
krow("% lỗi xử lý bằng nhắc nhở chung", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['bp_Nhắc nhở / yêu cầu chung']}",
     "0%", f"=ChatLuongKiemTra!E{Q['bp_Nhắc nhở / yêu cầu chung']}")
krow("% lỗi báo 'đạt' ngay trong ngày", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['tt_Báo đạt trong ngày']}",
     "0%", f"=ChatLuongKiemTra!E{Q['tt_Báo đạt trong ngày']}")
krow("% lỗi tái diễn (cùng khu vực + chủ đề ≥ 2 lần)", lambda p: f"=ChatLuongKiemTra!{L(2 + NGUOI.index(p))}{Q['lap_pct']}",
     "0%", f"=ChatLuongKiemTra!E{Q['lap_pct']}")
r += 1
put(wsK, r, 1, "Định biên (từ sheet DinhBien)", bold=True, bg=HEAD)
put(wsK, r, 2, "Giá trị", bold=True, bg=HEAD, align=CENTER)
r += 1
put(wsK, r, 1, "Số người ISO chuyên trách cần sau sắp xếp (tương đương)")
put(wsK, r, 2, f"=DinhBien!{B['iso_fte']}", fmt="0.00", align=CENTER, color="008000", bold=True)
r += 1
put(wsK, r, 1, "Việc kế hoạch / thống kê phải có người nhận (tương đương người)")
put(wsK, r, 2, f"=DinhBien!{B['kh_fte']}", fmt="0.00", align=CENTER, color="008000", bold=True)
r += 2
put(wsK, r, 1, "Kết luận", bold=True, color=ACCENT, border=False, size=12)
r += 1
KL = [
    "1. Vị trí 'ISO kiêm kế hoạch, mỗi xưởng một người' đang làm ba loại việc trộn lẫn: (a) việc ISO, (b) việc kế hoạch – văn thư của "
    "chuyền trưởng, kế hoạch, phòng cắt, kho, (c) trung gian chuyển tin. Loại (b) và (c) không thuộc chức năng ISO và tạo kênh chỉ đạo song song.",
    "2. Ngay cả phần 'ISO' cũng chủ yếu là tuần tra 2–4 lượt/ngày và nhắc nhở – tức làm thay việc kiểm tra hằng ngày của quản lý chuyền. "
    "Lỗi được báo 'đạt' trong ngày nhưng cùng chủ đề lặp lại ở nhiều chuyền, nhiều ngày: kiểm tra không tạo thay đổi hệ thống.",
    "3. Báo cáo ngày phần lớn chép lại mẫu (52–61% dòng giống nguyên văn ngày khác, khung giờ cố định). Không có mục tiêu, chỉ tiêu, "
    "danh sách lỗi mở, hạn đóng – đúng với nhận định 'thấy gì làm đó, thấy đâu nói đó'.",
    "4. Về số người: nếu ISO chỉ làm đúng chức năng (đánh giá xác suất, xác minh VDKL, kiểm soát tài liệu tập trung, theo dõi khắc phục), "
    "khối lượng ước tính ở sheet DinhBien tương đương khoảng 1 người cho cả 3 xưởng. Ba vị trí hiện nay là dư so với chức năng ISO.",
    "5. Nhưng việc kế hoạch đang nằm trong 3 vị trí này không tự biến mất. Phần dồn về Kế hoạch / thống kê (đồng bộ cắt – in – thêu, "
    "năng suất, biểu mẫu) cần người nhận; nếu Kế hoạch không có công suất, có thể điều chuyển một người sang Kế hoạch / điều độ thay vì để lại trong ISO.",
    "6. Đề xuất: không giữ mô hình 3 ISO kiêm kế hoạch. Mục tiêu 1 đầu mối ISO chuyên trách (+ đầu mối kiểm soát tài liệu nếu chưa gộp được), "
    "chuyển việc theo sheet GiuChuyenBo, chạy thử 10 ngày làm việc để đo lại giờ thực tế rồi chốt 1 hay 2.",
    "Lưu ý: số giờ là giờ tự ghi theo khung cố định, không phải giờ đo. Báo cáo thứ ba chưa xác định người viết. Kết luận về năng lực từng người "
    "cần thêm phản hồi của quản đốc và đối chiếu hồ sơ cấp BTP (sheet DoiChieuBTP).",
]
for t in KL:
    wsK.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    put(wsK, r, 1, t, border=False, align=Alignment(wrap_text=True, vertical="top"))
    wsK.row_dimensions[r].height = 44
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
    "Kỳ dữ liệu: file nhận ngày 10/11 có thêm ngày 10/11 (Vương, báo cáo 3). Thân: 03,04,05,06,07,09/11. Vương và báo cáo 3: 03,04,05,06,07,09,10/11. "
    "Ngày 08/11 là Chủ nhật.",
    "Handoff ghi 17 ngày công trong kỳ 03–09/11; bản file hiện tại có 18 ngày trong kỳ đó (mỗi báo cáo 6 ngày) và 20 ngày nếu tính cả 10/11.",
    "Báo cáo 3 ghi 'ứng/nhận vải thử máy' 4 ngày (04, 06, 07, 09/11) chứ không phải 3/5 như handoff.",
    "Báo cáo 3 ghi 'kiểm tra vốn chuyền để đăng ký BTP cho ngày 5-11' ở các ngày 03, 04, 05 và 06/11 – ngày 06/11 vẫn ghi cho ngày 05/11, "
    "dấu hiệu chép mẫu. Từ 07/11 ghi 9-11, 10-11, 'hôm sau'.",
    "Vương ghi ngày 10/11 thành 2026-10-11 (đảo ngày tháng); đã hiểu là 10/11.",
    "Báo cáo Vương và Thân không ghi tên trong file, chỉ có tên file. Báo cáo 3 không có tên; chuyền nhắc tới là 6B–11B và có nhắc 'cầu qua Xí nghiệp A1', "
    "nên không phải A1. Thân nhắc chuyền 1A–6A. Vương ghi Xí nghiệp A3, chuyền 12A–15B.",
    "Tỷ lệ dòng chép nguyên văn (cùng hạng mục + nội dung + biện pháp xuất hiện ở ≥ 2 ngày): Thân 54%, Vương 52%, báo cáo 3 61%. "
    "Khung giờ: Thân và báo cáo 3 dùng một bộ khung duy nhất cho mọi ngày; Vương 3 bộ trong 7 ngày.",
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
