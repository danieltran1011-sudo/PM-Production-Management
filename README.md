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
| `docs/quy-trinh-lead-time.md` | Tài liệu quy trình, bảng thời gian chuẩn, bảng so sánh, việc cần làm sau họp, các điểm còn mở. |
| `docs/phan-cong-cong-viec.md` | Dữ liệu phân công công việc TIVN tại APIC, BTNB: nhân sự và việc chính từng người, danh mục ~100 việc kèm người làm và % thời gian, các việc tính vào trị giá hải quan, mã hàng cố vấn 天方 phụ trách. |
| `docs/phan-tich-khoi-luong.md` | Phân tích khối lượng công việc hiện tại của nhóm merchandiser APIC (bulk, mẫu, giờ dịch, chỉ số tải), nhận xét, dữ liệu cần xác nhận, cách theo dõi theo sổ tay MD. |
| `data/khoi-luong/` | Số liệu thực tế BULK/mẫu đã gom và tải hiện tại theo người (JSON), script `build.py` tính lại. |
| `data/phan-cong/` | Dữ liệu phân công dạng JSON (APIC 2026 上期, APIC/BTNB 2023–2024, bảng 2018, phân công brand MD 22/7/2026) và script `build.py` tạo lại từ Excel. |
| `data/nguon/` | Các file Excel gốc. |

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```
