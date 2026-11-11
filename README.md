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
| `bao-cao/xlsx-charts.js` | Thêm biểu đồ Excel (sửa được trong Excel) vào file xuất ra: sheet "Bieu do" gồm bảng số và 13 biểu đồ. |
| `bao-cao/i18n-ja.js` | Từ điển Việt → Nhật: trang và file Excel hiển thị song ngữ (nút "Việt + 日本語 / Chỉ tiếng Việt" ở đầu trang). Sửa thuật ngữ tại đây. |
| `bao-cao/xlsx-style.js` | Định dạng file Excel xuất ra (màu tiêu đề, viền, sọc xen kẽ, định dạng % và hàng nghìn, cố định hàng tiêu đề, tô xanh/đỏ đạt/chưa đạt cam kết TI). |
| Sổ ライン空き原価配分表 (APIC) | Kéo vào trang cùng các file khác: mỗi sheet "N月実績" là một tháng. Trang đối chiếu chuyền trống / may Nhà Bè theo sổ với số trang tính và file APIC負担金額 của TI, thống kê nguyên nhân và chi phí, báo ô tổng trong sổ cộng thiếu dòng. |
| Báo cáo tuần · tháng (mẫu TIVN) | Mục mới trên trang, theo mẫu 週報 của TIVN APIC課: bảng tổng hợp tháng (năm trước, cùng kỳ, KH 3 tháng tới), sản lượng / 工程時間 / lot so cùng kỳ, cơ cấu 工程時間 (5 nhóm), sản phẩm (アイテム), lot (7 nhóm), số theo tuần (tuần từ thứ Hai, tuần 1 có ngày 1/4), kế hoạch từng chuyền, lời bình tự động tiếng Nhật + Việt; sheet Excel 'Bao cao TIVN'. Số theo tuần cần file nạp từ bản này. |
| `bao-cao/report.test.js` | Test đọc sheet, gộp nhiều file, lao động chuyền, năm tài chính, so sánh cùng kỳ. |
| `docs/quy-trinh-lead-time.md` | Tài liệu quy trình, bảng thời gian chuẩn, bảng so sánh, việc cần làm sau họp, các điểm còn mở. |
| `data/nguon/` | Hai file Excel gốc. |

## Báo cáo tháng từ file KH_UPDATE

Mở `bao-cao/index.html`, kéo thả các file `KH_UPDATE__APA1.YYYYMMDD.xlsm`, `KH_UPDATE__APA2.YYYYMMDD.xlsm`. Dữ liệu được lưu trong trình duyệt, lần sau chỉ cần thêm file mới.

- Ngày update lấy từ tên file (`.20261031.` → 31/10/2026); xưởng lấy từ cột 工場.
- Mỗi tháng lấy số từ **file chốt**: file update sớm nhất từ ngày cuối tháng trở đi. Vì vậy mỗi tháng cần ít nhất một file update vào ngày cuối tháng (hoặc đầu tháng sau).
- **KH đầu tháng** lấy từ file update trước ngày 1 của tháng.
- Năm tài chính mặc định bắt đầu tháng 4, đổi được ở bộ lọc.
- Mục **Tổng hợp toàn xưởng**: gộp mọi xưởng (hoặc theo bộ lọc), xem từng tháng trong năm tài chính (kèm cả năm, cùng kỳ năm trước) hoặc các năm tài chính: sản lượng, kế hoạch, mã hàng, chuyển mã, lot, thời gian công đoạn, chuyền, lao động, năng suất, hiệu suất, lên năng suất, lead time, ổn định kế hoạch.
- Mục **Cam kết với TI** (hợp đồng 4/2026–3/2029): năng suất +5%/năm so với năm chuẩn NTC 2025, hàng khó (工程時間 > 20.000 giây) ≤ 2% và lot nhỏ (≤ 500 sp trên chuyền) ≤ 10% tổng số lượng sản xuất năm; thực tế, dự báo cả năm, hạn mức còn lại, danh sách mã.
- **Hàng TI / hàng APIC tự bán**: theo danh sách nhãn hàng TI (sửa được trên trang); bộ lọc "Nhóm hàng". Cam kết với TI chỉ tính hàng TI, kèm cách tính của TI (sản lượng hàng TI / chuyền-ngày, năm chuẩn 647.091 sp).
- Mục **Sử dụng chuyền**: mỗi nửa chuyền A/B mỗi ngày làm việc may hàng TI, hàng APIC tự bán hay trống; giai đoạn chuyền trống / APIC自販分 theo tháng; so sánh các năm. Nạp file **APIC負担金額** của TI để đối chiếu từng chuyền và số tiền.
- Mục **Phân tích** (theo kỳ: lũy kế năm tài chính, 12 tháng, tháng, toàn bộ): (1) cỡ lot và hiệu suất, (2) lên năng suất so với 習熟日数, (3) lead time may so với 縫製リードタイム, (4) chuyển mã và giờ chuẩn mất, (5) độ ổn định kế hoạch (KH đầu tháng so với TH), (6) xếp hạng chuyền, (7) tải các tháng tới so với năng lực.
- Kiểm tra file: công thức tổng ở XN1AB!A1:R3 (sheet TORAY) và tháng còn thiếu file cuối tháng.
- Cách tính chi tiết ở cuối trang.

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```
