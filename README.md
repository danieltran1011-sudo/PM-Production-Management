# PM Production Management

Workflow chuẩn bị mã hàng từ ETA đến START (ra chuyền), dùng để thảo luận mốc giao nguyên phụ liệu với TIJ.

**Chuẩn mới:** ETA (về cảng) trước START 28 ngày, tương đương 工場着 trước START 21 ngày.
**Chuẩn cũ:** ETA trước START 21 ngày, 工場着 trước START 14 ngày.

## Nội dung

| Đường dẫn | Mô tả |
|---|---|
| `workflow/flowchart.html` | Flowchart lead time song ngữ Nhật–Việt (生産工程リードタイム): so sánh chuẩn cũ và mới, flowchart theo bộ phận, lý do thay đổi, đề nghị gửi TIJ. |
| `workflow/planner.html` | Công cụ lập lịch từng mã hàng: nhập ETA, SAH, mã mới/lặp lại, in/thêu → ngày của từng bước, Gantt, tính ngược ETA muộn nhất, danh sách nhiều mã, sửa số ngày chuẩn. |
| `workflow/schedule.js` | Phần tính lịch dùng chung cho hai trang. |
| `workflow/schedule.test.js` | Test đối chiếu với các cột N1…N27 trong file Excel rà soát. |
| `docs/quy-trinh-lead-time.md` | Tài liệu quy trình, bảng thời gian chuẩn, bảng so sánh, các điểm còn mở. |
| `data/nguon/` | Hai file Excel gốc. |

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```
