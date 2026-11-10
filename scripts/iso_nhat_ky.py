"""Tách ba báo cáo công việc hằng ngày của nhân sự ISO kiêm kế hoạch thành nhật ký từng dòng.

Mỗi dòng nhật ký: người báo cáo, ngày, khung giờ, số phút được phân bổ, mã việc, nội dung,
ô nguồn. Dùng chung cho scripts/xuat_phan_tich_iso.py.

Quy tắc phân bổ thời gian (giả định, ghi rõ trong file Excel):
- Thời gian lấy từ cột "Thời gian" do nhân sự tự ghi, không phải đo thực tế.
- Trong một khung giờ có nhiều mã việc, số phút chia đều cho các mã việc có mặt.
- Khung "15h" của Thân không có giờ kết thúc, tính 15h00–15h45 (khung kế tiếp bắt đầu 15h45).
"""
import datetime as dt
import re

from openpyxl import load_workbook

NGUON = {
    "THAN": ("data/nguon/iso/Bao-cao-hang-ngay-Than.xlsx", "Thân"),
    "VUONG": ("data/nguon/iso/Bao-cao-hang-ngay-Vuong.xlsx", "Vương (A3)"),
    "BC3": ("data/nguon/iso/NOI-DUNG-KIEM-TRA.xlsx", "Báo cáo 3 (chưa rõ tên)"),
}

# Mã việc: (tên, nhóm, chức danh/bộ phận nên làm chính, đề xuất)
VIEC = {
    "I1": ("Tuần tra 5S, vệ sinh công nghiệp, nội quy, BHLĐ (đầu ca, giữa ca, 9h, 15h)",
           "ISO", "Quản lý chuyền / quản lý khu vực (kiểm tra hằng ngày); Nhóm 5S xí nghiệp",
           "GIỮ – ĐỔI CÁCH LÀM"),
    "I2": ("Kiểm soát vật dụng kim loại / vật dụng nguy hiểm (bảng đánh chéo, xác suất, bóp VDKL)",
           "ISO", "Quản lý chuyền (kiểm hằng ngày, ký bảng); ISO xác minh xác suất",
           "GIỮ – GIẢM TẦN SUẤT"),
    "I3": ("Nhập nhiệt độ, độ ẩm", "ISO", "Tổ mẫu / thống kê chuyền, hoặc thiết bị ghi tự động",
           "CHUYỂN"),
    "I4": ("Kiểm tra, photo, đóng mộc, ban hành, lưu, hủy tài liệu sản xuất", "ISO",
           "Kiểm soát tài liệu tập trung (1 đầu mối cho 3 xưởng)", "GIỮ – GỘP 1 ĐẦU MỐI"),
    "I5": ("Tổng hợp vi phạm, nhập và gửi báo cáo", "ISO", "ISO", "GIỮ – RÚT GỌN"),
    "P1": ("Kiểm tra vốn chuyền, đăng ký bán thành phẩm (BTP) với phòng cắt", "KẾ HOẠCH",
           "Chuyền trưởng đăng ký; Kế hoạch chốt ngày, lượng, thứ tự", "CHUYỂN"),
    "P2": ("Theo dõi cắt cấp BTP, đồng bộ in / thêu / laze / keo, hối thúc đầu vào", "KẾ HOẠCH",
           "Kế hoạch / điều độ; Phòng cắt; bộ phận in, thêu, laze", "CHUYỂN"),
    "P3": ("Làm file, tổng hợp, chuyển phiếu đăng ký phụ liệu xuống kho", "KẾ HOẠCH",
           "Chuyền trưởng đăng ký trực tiếp với Kho phụ liệu", "CHUYỂN"),
    "P4": ("Theo dõi năng suất, phân chia kế hoạch từng chuyền", "KẾ HOẠCH",
           "Quản đốc / chuyền trưởng (năng suất); Kế hoạch (phân chuyền)", "CHUYỂN"),
    "P5": ("Đánh máy, in, cấp list đánh chéo, list giấy số, bảng ghi năng suất", "VĂN THƯ",
           "Thống kê chuyền / Kế hoạch (mẫu chuẩn in sẵn theo mã hàng)", "CHUYỂN"),
    "P6": ("Ứng, nhận vải thử máy cho mã hàng mới", "KẾ HOẠCH",
           "Kỹ thuật chuyền / chuyền trưởng nhận trực tiếp từ phòng cắt", "CHUYỂN"),
    "P7": ("Đọc mail, nhóm chat, chuyển thông tin lỗi chuyền cho bộ phận khác", "TRUNG GIAN",
           "Chuyền trưởng báo thẳng bộ phận liên quan theo kênh báo lỗi chuẩn", "BỎ"),
    "M1": ("Cấp vật tư phụ trợ (kệ chỉ, rổ, vải vệ sinh, ra bàn ủi, mặt bàn)", "HẬU CẦN",
           "Kho / hành chính xưởng, quản lý chuyền đề nghị", "CHUYỂN"),
}

# Từ khóa phân loại theo thứ tự ưu tiên (so khớp chữ thường)
QUY_TAC = [
    ("P6", ["vải thử máy"]),
    ("M1", ["cắt ra bàn ủi", "chuẩn bị vải  vệ sinh", "chuẩn bị vải vệ sinh", "bàn btp, bàn bị rách"]),
    ("P3", ["phụ liệu cho chuyền", "đăng ký phụ liệu"]),
    ("P1", ["đăng ký bán thành phẩm", "đăng ký btp", "vốn chuyền để đăng ký"]),
    ("P2", ["cung cấp bán thành phẩm", "bán thành phẩm phòng cắt", "theo dõi cấp bán thành phẩm",
            "kiểm tra in", "kiểm tra thêu", "kiểm tra  thêu", "laze"]),
    ("P5", ["list đánh chéo", "list giấy số", "bảng ghi năng suất"]),
    ("P4", ["năng suất từng chuyền", "kế hoạch từng chuyền"]),
    ("I3", ["nhiệt độ", "thiệt độ"]),
    ("I4", ["tài liệu", "huỷ tài liệu"]),
    ("P7", ["mail", "tiếp nhận thông tin", "liên hệ các bộ phận"]),
    ("I5", ["nhập báo cáo", "danh sách cán bộ", "tổng hợp và báo cáo"]),
    ("I2", ["vật dụng kim loại", "vdkl", "vdnh"]),
    ("I1", ["5s", "vệ sinh", "nội quy", "tác phong", "nhận dạng", "đầu giờ ca"]),
]

# Từ khóa đủ rõ để một dòng không có hạng mục riêng vẫn được tách khỏi tiêu đề khung giờ
MANH = ["vải thử máy", "cắt ra bàn ủi", "đăng ký phụ liệu", "phụ liệu cho chuyền",
        "đăng ký bán thành phẩm", "cung cấp bán thành phẩm", "bán thành phẩm phòng cắt",
        "tiếp nhận thông tin", "nhiệt độ", "ban hành tài liệu", "danh sách cán bộ"]

# Tiêu đề khung giờ gồm nhiều việc: phân loại theo nội dung từng dòng
TIEU_DE_GOP = ["báo cáo", "ban hành tài liệu,kiểm tra", "kiểm tra và xử lý thông tin",
               "kiểm tra +ban hành", "tổng hợp và báo cáo", "kiểm tra mail,trên nhóm"]


def phan_loai(text):
    t = (text or "").lower()
    for ma, kws in QUY_TAC:
        if any(k in t for k in kws):
            return ma
    return None


def doc_gio(s):
    """'7h30-8h00'' -> (450, 480). '15h' -> (900, None)."""
    s = str(s).replace("'", "").replace(" ", "").replace("\n", "")
    parts = re.findall(r"(\d{1,2})h(\d{0,2})", s)
    mins = [int(h) * 60 + (int(m) if m else 0) for h, m in parts]
    if not mins:
        return None, None
    return mins[0], (mins[1] if len(mins) > 1 else None)


def doc_ngay(v, nguoi):
    if isinstance(v, dt.datetime):
        d = v.date()
        # Vương ghi ngày 10/11 thành 2026-10-11 (đảo ngày tháng)
        if d.month == 10 and d.day == 11 and nguoi == "VUONG":
            return dt.date(2026, 11, 10), "Ô ghi 2026-10-11, hiểu là 10/11 (đảo ngày tháng)"
        return d, ""
    m = re.match(r"\s*(\d{1,2})\s*-\s*(\d{1,2})", str(v))
    return dt.date(2026, int(m.group(2)), int(m.group(1))), ""


def doc_bao_cao(ma_nguoi):
    path, _ = NGUON[ma_nguoi]
    ws = load_workbook(path, data_only=True).worksheets[0]
    head = 3 if ma_nguoi in ("THAN", "VUONG") else 2
    rows = []
    ngay = gio = tieu_de = None
    ghi_chu_ngay = ""
    for r in range(head + 1, ws.max_row + 1):
        a, b, c, d, e, f, g = (ws.cell(r, k).value for k in range(1, 8))
        if not any([b, c, d, e, f, g]):
            continue
        if a:
            ngay, ghi_chu_ngay = doc_ngay(a, ma_nguoi)
            tieu_de = None
        if b:
            gio = doc_gio(b)
            tieu_de = None
        if c:
            tieu_de = str(c)
        tdl = (tieu_de or "").lower()
        gop = (not tieu_de) or any(tdl.startswith(x) for x in TIEU_DE_GOP)
        el = (e or "").lower() if not c else ""
        if not c and any(k in el for k in MANH):
            ma = phan_loai(e)
        elif c and not gop:
            ma = phan_loai(c)
        elif gop:
            ma = phan_loai(e) or phan_loai(c) or phan_loai(tieu_de)
        else:
            ma = None
        if ma is None:
            ma = phan_loai(tieu_de) or phan_loai(e)
        ma_phu = []
        if ma in ("I1", "I2") and re.search(r"5s", tdl) and re.search(r"vdkl|kim loại", tdl):
            ma, ma_phu = "I1", ["I2"]
        rows.append(dict(ma_phu=ma_phu, nguoi=ma_nguoi, ngay=ngay, bat_dau=gio[0], ket_thuc=gio[1], dong=r,
                         hang_muc=str(c or tieu_de or "").strip(), khu_vuc=str(d or "").strip(),
                         noi_dung=str(e or "").strip(), bien_phap=str(f or "").strip(),
                         ket_qua=str(g or "").strip(), ma=ma, ghi_chu=ghi_chu_ngay))
    # Giờ kết thúc còn thiếu: lấy giờ bắt đầu của khung kế tiếp trong cùng ngày
    for i, x in enumerate(rows):
        if x["ket_thuc"] is None:
            nxt = [y["bat_dau"] for y in rows[i + 1:] if y["ngay"] == x["ngay"] and y["bat_dau"] > x["bat_dau"]]
            x["ket_thuc"] = nxt[0] if nxt else x["bat_dau"] + 30
    # Phân bổ phút: chia đều khung giờ cho các mã việc có mặt trong khung
    khung = {}
    for x in rows:
        khung.setdefault((x["ngay"], x["bat_dau"]), []).append(x)
    for xs in khung.values():
        cac_ma = sorted({m for x in xs for m in [x["ma"]] + x["ma_phu"]})
        phut = (xs[0]["ket_thuc"] - xs[0]["bat_dau"]) / len(cac_ma)
        dau = set()
        for x in xs:
            # Chỉ dòng đầu tiên của mỗi mã việc trong khung mang số phút
            x["phut"] = 0 if x["ma"] in dau else phut
            dau.add(x["ma"])
            x["phut_phu"] = {}
            for m in x["ma_phu"]:
                if m not in dau:
                    x["phut_phu"][m] = phut
                    dau.add(m)
    return rows


def doc_tat_ca():
    out = []
    for k in NGUON:
        out += doc_bao_cao(k)
    return out


if __name__ == "__main__":
    import collections
    rows = doc_tat_ca()
    for k, (_, ten) in NGUON.items():
        rs = [x for x in rows if x["nguoi"] == k]
        days = sorted({x["ngay"] for x in rs})
        tong = sum(x["phut"] + sum(x["phut_phu"].values()) for x in rs)
        print(f"\n== {ten}: {len(days)} ngày {[d.strftime('%d/%m') for d in days]}, "
              f"{tong/60:.1f} giờ ghi, {tong/len(days)/60:.2f} giờ/ngày")
        c = collections.Counter()
        dd = collections.defaultdict(set)
        for x in rs:
            c[x["ma"]] += x["phut"]
            dd[x["ma"]].add(x["ngay"])
            for m, p in x["phut_phu"].items():
                c[m] += p
                dd[m].add(x["ngay"])
        for ma, p in sorted(c.items(), key=lambda t: -t[1]):
            print(f"  {ma}: {p/len(days):6.1f} phút/ngày {p/tong*100:5.1f}%  {len(dd[ma])} ngày")
    none = [x for x in rows if x["ma"] is None]
    print("\nChưa phân loại:", len(none))
    for x in none[:30]:
        print(" ", x["nguoi"], x["dong"], x["hang_muc"][:50], "|", x["noi_dung"][:60])
