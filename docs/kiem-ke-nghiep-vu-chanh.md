# 業務棚卸 – Chanh (TIVN–APIC): kết quả giai đoạn 1

File: `nghiep-vu/Kiem-ke-nghiep-vu-Chanh-2026.xlsx` (tạo bằng `nghiep-vu/build_inventory.py`; file kết quả và file nguồn không đưa lên repo vì repo công khai).
Số liệu dưới đây là giá trị tại 2026-10-28 (log đến W43). Trong file, mọi tổng hợp là công thức, tự cập nhật khi sửa.

## Cách làm

1. **Xương sống:** 業務分析表 2026年度上期 (102 mục). Lấy 41 mục có FTE của Chanh (tổng = 1,00 FTE) và mục No.25 (FTE = 0 nhưng có trong log).
2. **Bằng chứng việc thực tế:** sheet `業務棚卸し` (434 bản ghi: 23 dòng nền + 411 bản ghi tuần W19–W43). Mỗi bản ghi được gán vào một No. (sheet `Record_Map`).
3. **Việc ngoài 業務分析表:** thêm mục mới **No.103–148** (từ log, bàn giao của Inoue, 移管資料, brief, repo lead time).
4. **Trạng thái:** Confirmed = có trong log 13 tuần gần nhất; To Verify = chỉ có trong 業務分析表/nguồn khác; Not Current = lịch sử/dự kiến.

## Kết quả chính

| Chỉ số | Giá trị |
|---|---|
| Bản ghi trong 業務棚卸し | 434 (brief ghi 218 – số cũ); 323 tên khác nhau (brief ghi 191); 23 dòng nền (đúng) |
| Mục trong danh sách nghiệp vụ | 88 = 42 mục 業務分析表 + 46 mục mới |
| Trạng thái | 43 Confirmed / 43 To Verify / 2 Not Current |
| Giờ Chanh tự ghi TB/tuần (W34–W43) | 68,5h so với 48h khả dụng (8h × 6 ngày) → 143% |
| Thêm mục có FTE nhưng không ghi trong log | +4,3h → tổng 152% |
| Giờ log nằm ở mục ngoài 業務分析表 (No.103+) | 57% |
| Giờ ở mức nên giao / chuyển / giảm (C+D+E) | ≈ 43% |
| Nếu áp dụng toàn bộ đề xuất (dự thảo) | 152% → ~91% giờ khả dụng; tỷ trọng Decision + System Building 29% → 40% |

Giờ trong 業務棚卸し là **ước tính**, có trùng giữa các việc (báo cáo tuần W43 tự ghi điều này). 143% có thể là làm thêm giờ, ghi trùng, hoặc cả hai → cần đo thực tế 2 tuần.

### Nhóm tốn giờ nhất (giờ chuẩn hóa, W34–W43)

1. 生産計画・Plantable – 15%
2. 生産リスク管理 (chuyền trống, R1/R2/R3) – 11%
3. TIJ・APIC連絡調整 (điều phối, email) – 9%
4. 資材・納期管理 – 7%
5. 組織・人材管理 – 7%

### Plantable: các con số chưa khớp

| Nguồn | h/tuần |
|---|---|
| Log W34–W43 (kế hoạch + họp, chưa chuẩn hóa) | ~12,8 |
| Dòng nền R5 週次生産計画会議 | 16 |
| Buổi với Hori: ~2 ngày/tuần | 16 |
| 業務分析表 2026上期 (No.4, 18, 20, 85, 88–90, 143) | ~17,3 |

Log cho thấy Plantable đang giảm dần từ khi giao việc cho Leader (W29+). Cần thống nhất con số với Hori sau khi đo.

### Phạm vi theo chức vụ (課長)

| Mức | Ý nghĩa | Tỷ trọng giờ |
|---|---|---|
| A | Cốt lõi 課長: quyết định, đàm phán, rủi ro, tổ chức, con người | 41% |
| B | 課長 review/duyệt, staff làm | 17% |
| C | Nên giao Leader/staff | 34% |
| D | Chuyển GA/総務/法務/cấp trên | 3% |
| E | Bỏ / đơn giản hóa | 5% |

Đề xuất chi tiết từng mục ở sheet `Role_Analysis` (dự thảo, chưa quyết định).

## Đề xuất giảm tải (dự thảo – cần Chanh/Hori duyệt)

- **Plantable (No.18, 20, 88–90):** Planning Leader cập nhật theo 7 quy tắc chọn chuyền của 移管資料; Chanh chỉ review thay đổi lớn và case R1. Viết quy tắc thành SOP trước.
- **Theo dõi hằng ngày (No.21, 63, 104):** giao Planning Leader / MD; Chanh nhận báo cáo ngoại lệ.
- **Báo cáo (No.80, 87, 92–98, 124):** gộp thành một bộ báo cáo tuần; báo cáo tháng nhỏ giao staff, tự động hóa.
- **Hành chính (No.10, 12, 119, 130, 140):** chuyển GA/Admin.
- **Hợp đồng khác, NDA, bảo hiểm (No.112):** chuyển 総務/法務.
- **Giữ và tăng thời gian:** rủi ro R1 (105), hợp đồng TI–APIC (111), Manpower (110), tổ chức và ủy quyền (126–128), con người (102, 137, 138).

## Tư vấn

1. Đo thời gian thực tế 2 tuần theo No. của file này trước khi chốt con số giao việc.
2. Thêm cột `業務No.` vào mẫu 業務棚卸し; thống nhất 頻度 (đang có 19 kiểu ghi) và ghi giờ dạng số.
3. Phần lớn việc thực tế không có trong 業務分析表. Đề xuất cập nhật 業務分析表 2026下期 – vì nó gắn với phí ủy thác và 評価加算, phải qua Hori và TI (biên bản 2023-11-27).
4. Chỉ nhận lên mình case R1; thêm "người quyết định" cho từng mức R1/R2/R3.
5. Email: 2 khung giờ/ngày, quy tắc CC, mẫu mail, AI dịch; bảng "ai trả lời câu hỏi gì" với TIJ/SW1/SW2.
6. Họp: gửi danh sách thông tin thiếu trước họp, giới hạn 60–90 phút, MD ghi biên bản và TODO.
7. Hợp đồng TI–APIC là dự án có điểm kết thúc: lên kế hoạch dùng giờ giải phóng sau khi ký cho con người và hệ thống.
8. Việc quản lý người (duyệt chuyên cần, đánh giá, tuyển dụng, 1-on-1 với Leader) gần như không có trong log – đặt lịch cố định.
9. Khi giao: giao cả quyền và tiêu chí; tránh "giao ngược"; đo số case Leader tự đóng mỗi tuần.
10. Dùng file này để thống nhất với Hori thứ tự ưu tiên và việc sẽ dừng khi quá tải.

## Việc cần Chanh làm tiếp

Xem sheet `Action_Plan`. Quan trọng nhất:

- Rà 43 mục To Verify và sửa cột Y trong `業務一覧_Chanh`.
- Xác nhận 11 bản ghi gán chưa chắc (`Record_Map` cột X).
- Đo thời gian thực tế 2 tuần, sửa tham số giờ làm (`Workload_Analysis` C5:C6).

## Chạy lại

```
python3 -I nghiep-vu/build_inventory.py --src <thư mục file nguồn> --out nghiep-vu/Kiem-ke-nghiep-vu-Chanh-2026.xlsx
```

Sau đó mở bằng Excel (hoặc tính lại bằng LibreOffice) để có giá trị công thức. File nguồn không có trong repo vì repo công khai.
