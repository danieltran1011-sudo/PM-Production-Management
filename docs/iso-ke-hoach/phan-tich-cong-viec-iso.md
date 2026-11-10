# APIC – Phân tích công việc 3 nhân sự ISO kiêm kế hoạch

Nguồn: 3 báo cáo công việc hằng ngày, kỳ 03–10/11/2026 (Thân 6 ngày; Vương – A3 7 ngày; báo cáo thứ ba chưa rõ tên 7 ngày).
Số liệu đầy đủ, công thức và nhật ký từng dòng: `Phan-tich-cong-viec-ISO.xlsx` (tạo lại bằng `python3 scripts/xuat_phan_tich_iso.py`).

> Giờ trong bảng là **giờ tự ghi** theo khung giờ trong báo cáo, không phải giờ đo. Khung giờ gần như giống nhau mỗi ngày
> và 52–61% dòng báo cáo chép nguyên văn từ ngày khác, nên đây là cơ cấu công việc tự khai. Cần đo lại trong 10 ngày thử.

## Kết luận

1. **Vị trí "ISO kiêm kế hoạch, mỗi xưởng một người" làm ba loại việc trộn lẫn**: việc ISO; việc kế hoạch, văn thư của chuyền trưởng, Kế hoạch, phòng cắt, kho; và vai trò trung gian chuyển tin. Hai loại sau không thuộc chức năng ISO và tạo kênh chỉ đạo song song với Kế hoạch.
2. **Phần "ISO" phần lớn cũng là làm thay quản lý chuyền.** Tuần tra 5S và kiểm vật dụng kim loại (VDKL) chiếm 50–64% thời gian, 2–4 lượt mỗi ngày. 60% lỗi chỉ xử lý bằng nhắc nhở, 84% lỗi được báo "đạt" ngay trong ngày, nhưng 46% lỗi lặp lại cùng khu vực và cùng chủ đề. Kiểm tra như vậy không làm thay đổi hệ thống.
3. **Không có mục tiêu, chỉ tiêu, danh sách lỗi mở hay hạn đóng.** Báo cáo ngày là mẫu chép lại. Điều này khớp với nhận định "thấy gì làm đó, thấy đâu nói đó".
4. **Về số người:** nếu ISO chỉ làm đúng chức năng (đánh giá xác suất, xác minh VDKL, kiểm soát tài liệu tập trung, theo dõi khắc phục), khối lượng ước tính khoảng **0,9 người cho cả 3 xưởng**. So với chức năng ISO, ba vị trí hiện nay là dư.
5. **Việc kế hoạch không tự biến mất.** Khoảng 3,7 giờ mỗi ngày (≈ 0,5 người) phải chuyển về Kế hoạch / thống kê: đồng bộ cắt – in – thêu – laze, năng suất, biểu mẫu. Nếu Kế hoạch không có công suất, có thể điều chuyển một người sang Kế hoạch / điều độ thay vì giữ trong ISO.
6. **Đề xuất:** bỏ mô hình 3 ISO kiêm kế hoạch. Mục tiêu là 1 đầu mối ISO chuyên trách cho cả 3 xưởng, thêm đầu mối kiểm soát tài liệu nếu chưa gộp được. Chuyển việc theo bảng giữ / chuyển / bỏ, chạy thử 10 ngày làm việc, đo lại rồi chốt 1 hay 2 người.

## Cơ cấu thời gian tự ghi (% giờ mỗi ngày)

| | Thân | Vương (A3) | Báo cáo 3 |
|---|---|---|---|
| Giờ ghi/ngày | 8,2 | 8,8 | 8,2 |
| Việc ISO giữ lại (phải đổi cách làm) | 79% | 55% | 65% |
| – trong đó tuần tra 5S + kiểm VDKL | 64% | 50% | 53% |
| Việc nên chuyển về bộ phận khác | 8% | 40% | 27% |
| Việc nên bỏ (trung gian chuyển tin) | 12% | 5% | 8% |

## Trùng lặp với bộ phận, chức danh khác

| Việc ISO đang làm | Đúng ra thuộc về | Hệ quả |
|---|---|---|
| Tuần tra 5S, nội quy, BHLĐ 2–4 lượt/ngày | Quản lý chuyền (kiểm hằng ngày), nhóm 5S xưởng | Quản lý chuyền dựa vào ISO để phát hiện lỗi; lỗi lặp lại |
| Kiểm bảng đánh chéo, bóp VDKL cuối ngày | Quản lý chuyền; ISO chỉ xác minh xác suất | Ghi "Đạt" chung chung, không ghi số chuyền và số túi đã kiểm |
| Đăng ký BTP, kiểm vốn chuyền | Chuyền trưởng đăng ký; Kế hoạch chốt | Phòng cắt nhận yêu cầu qua người trung gian |
| Hối thúc cắt, in, thêu, laze, keo | Kế hoạch / điều độ | Hai kênh hối thúc, không rõ ai quyết ưu tiên |
| Theo dõi năng suất, "phân chia kế hoạch chuyền" | Quản đốc, Kế hoạch | Vượt quyền |
| Đánh máy, in list đánh chéo, list giấy số, bảng năng suất | Thống kê / Kế hoạch (mẫu chuẩn) | Việc văn thư lặp theo mã hàng |
| Gom phiếu phụ liệu, ứng vải thử máy, cấp kệ, rổ, vải vệ sinh | Chuyền ↔ Kho, kỹ thuật chuyền, hành chính | Thêm một bước trung chuyển |
| Đọc mail, nhóm chat, chuyển lỗi chuyền cho bộ phận khác | Chuyền trưởng báo thẳng qua kênh chung | Thông tin đi vòng, chậm, mất dấu trách nhiệm |
| Ban hành, photo, đóng mộc tài liệu ở cả 3 xưởng | Một đầu mối kiểm soát tài liệu | Ba người cùng làm một việc |
| Lập biên bản, hạ loại thi đua | Quản đốc / Nhân sự | ISO tự kỷ luật |

## Điểm cần xử lý ngay

- Vạch lối thoát hiểm ở A3 bị bong tróc, báo cáo ghi "sẽ thực hiện sau khi khách gần tới". Đây là hạng mục an toàn, cần đặt hạn sửa ngay.
- Vòng chắn kim bị gãy ở chuyền 12A, 15A, 15B (10/11) và tủ vật dụng nguy hiểm chuyền 3B lặp lỗi 05–07/11. Cả hai đều chưa có bằng chứng đóng lỗi.

## File Excel

| Sheet | Nội dung |
|---|---|
| KetLuan | Chỉ số chính theo người, kết luận |
| ThoiGian | Phút/ngày, % thời gian, số ngày có cho 13 mã việc × 3 người |
| TrungLap | Ma trận việc × bộ phận (chủ trì, phối hợp, ISO làm thay) và hệ quả |
| GiuChuyenBo | Đề xuất giữ / chuyển / bỏ: người nhận, quyền quyết định, đầu ra, tiêu chí nghiệm thu, điểm cần xác minh, ô chốt khi họp |
| ChatLuongKiemTra | Kiểu biện pháp, trạng thái, chủ đề lỗi, tỷ lệ tái diễn |
| DinhBien | Mô hình số người ISO cần, giả định sửa được, so sánh phương án 3 / 2 / 1 |
| DoiChieuBTP | Mẫu đối chiếu hồ sơ cấp BTP theo mã với phản hồi kỹ thuật (bước tiếp theo) |
| LoiHienTruong, NhatKy | Dữ liệu từng dòng, có ô nguồn |
| GhiChuDuLieu | Giả định, sai lệch dữ liệu, khác biệt với handoff |

## Khác với handoff

- File hiện có thêm ngày 10/11. Trong kỳ 03–09/11 có 18 ngày công, không phải 17. Báo cáo 3 ghi "vải thử máy" 4 ngày, không phải 3/5.
- Báo cáo 3 ghi "đăng ký BTP cho ngày 5-11" ở các ngày 03–06/11, kể cả ngày 06/11.
- Vương ghi ngày 10/11 thành 2026-10-11.
