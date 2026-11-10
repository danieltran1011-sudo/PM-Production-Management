# PM Production Management

Workflow chuẩn bị mã hàng từ ETA đến START (ライン投入, vào chuyền), dùng để thảo luận mốc giao nguyên phụ liệu với TIJ.

**Chuẩn mới:** ETA (về cảng) trước START 28 ngày, tương đương 工場着 trước START 21 ngày.
**Chuẩn cũ:** ETA trước START 21 ngày, 工場着 trước START 14 ngày.
**START** = ライン投入 (chuẩn bị xong, vào chuyền). Triển khai chuyền đến 5 pcs đầu diễn ra sau START.

Số ngày chuẩn lấy từ file rà soát 20260722, cập nhật theo cuộc họp rà soát lead time (in gia công ngoài sau khi cắt, thêu làm trong nhà máy, rập + sơ đồ trên 12.000 giây tối thiểu 3 ngày).

## Nội dung

| Đường dẫn | Mô tả |
|---|---|
| `workflow/flowchart.html` | Flowchart lead time song ngữ Nhật–Việt (生産工程リードタイム): so sánh chuẩn cũ và mới, flowchart theo bộ phận, thời gian từng bước đã xác nhận trong họp, khả năng làm trong 10–12 ngày, lý do thay đổi, đề nghị gửi TIJ, cải tiến phía nhà máy. |
| `workflow/planner.html` | Công cụ lập lịch từng mã hàng: nhập ETA, SAH, mã mới/lặp lại, in (gia công ngoài tại Hạnh Đan / Vĩnh Phát), thêu (trong nhà máy) → ngày của từng bước, Gantt, tính ngược ETA muộn nhất, danh sách nhiều mã, sửa số ngày chuẩn. |
| `workflow/schedule.js` | Phần tính lịch dùng chung cho hai trang. |
| `workflow/schedule.test.js` | Test đối chiếu với file Excel rà soát và các quy tắc đã xác nhận trong họp. |
| `docs/Leadtime-TIJ.xlsx` | File Excel song ngữ Nhật–Việt: 概要 (tổng hợp, so sánh chuẩn), フロー (Gantt chọn được điều kiện), 理由・依頼 (lý do, đề nghị TIJ), 標準日数 (thời gian chuẩn, sửa được), 計算 (công thức tính). Tạo lại bằng `python3 scripts/xuat_excel.py`. |
| `docs/quy-trinh-lead-time.md` | Tài liệu quy trình, bảng thời gian chuẩn, bảng so sánh, việc cần làm sau họp, các điểm còn mở. |
| `data/nguon/` | File Excel gốc (lead time; `iso/`: 3 báo cáo công việc hằng ngày của nhân sự ISO kiêm kế hoạch). |
| `docs/iso-ke-hoach/phan-tich-cong-viec-iso.md` | Phân tích công việc 3 nhân sự ISO kiêm kế hoạch: trùng lặp với bộ phận khác, việc giữ / chuyển / bỏ, định biên. |
| `docs/iso-ke-hoach/Phan-tich-cong-viec-ISO.xlsx` | File Excel phân tích (thời gian, ma trận trùng lặp, giữ / chuyển / bỏ, chất lượng kiểm tra, định biên, mẫu đối chiếu cấp BTP). Tạo lại bằng `python3 scripts/xuat_phan_tich_iso.py`. |

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```
