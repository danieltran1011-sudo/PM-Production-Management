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
| `bao-cao/index.html` | Báo cáo sản xuất từ các file KH_UPDATE (sheet XN1AB): nạp nhiều file (nhiều xưởng, nhiều tháng, nhiều năm), tự lấy số chốt từng tháng; so sánh tháng trước, cùng kỳ năm trước, lũy kế năm tài chính; theo nhãn hàng, chuyền, cụm chuyền, chủng loại, xưởng; lao động chuyền, năng suất, hiệu suất; xuất Excel. |
| `bao-cao/report.js` | Phần đọc file và tính báo cáo (dùng chung cho trang và test). |
| `bao-cao/report.test.js` | Test đọc sheet, gộp nhiều file, lao động chuyền, năm tài chính, so sánh cùng kỳ. |
| `docs/quy-trinh-lead-time.md` | Tài liệu quy trình, bảng thời gian chuẩn, bảng so sánh, việc cần làm sau họp, các điểm còn mở. |
| `data/nguon/` | Hai file Excel gốc. |

## Báo cáo tháng từ file KH_UPDATE

Mở `bao-cao/index.html`, kéo thả các file `KH_UPDATE__APA1.YYYYMMDD.xlsm`, `KH_UPDATE__APA2.YYYYMMDD.xlsm`. Dữ liệu được lưu trong trình duyệt, lần sau chỉ cần thêm file mới.

- Ngày update lấy từ tên file (`.20261031.` → 31/10/2026); xưởng lấy từ cột 工場.
- Mỗi tháng lấy số từ **file chốt**: file update sớm nhất từ ngày cuối tháng trở đi. Vì vậy mỗi tháng cần ít nhất một file update vào ngày cuối tháng (hoặc đầu tháng sau).
- **KH đầu tháng** lấy từ file update trước ngày 1 của tháng.
- Năm tài chính mặc định bắt đầu tháng 4, đổi được ở bộ lọc.
- Cách tính chi tiết ở cuối trang.

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```
