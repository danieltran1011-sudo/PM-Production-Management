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
| `data/nguon/` | Hai file Excel gốc. |

## Sử dụng

Mở `workflow/flowchart.html` hoặc `workflow/planner.html` bằng trình duyệt (không cần cài đặt). Số ngày chuẩn và ngày nghỉ sửa ở trang lập lịch được dùng chung cho trang flowchart trên cùng trình duyệt.

Chạy test:

```
npm test
```

## Báo cáo điều hành PLANTABLE (`bao-cao/`)

Báo cáo Excel song ngữ (tiếng Việt + tiếng Nhật) cho CEO, Giám đốc nhà máy, Quản lý sản xuất, dựng **chỉ từ file PLANTABLE** (APIC P-Plan …xlsm). Không cần file 受注進行表: sheet `4A. 受注進行表` được tự dựng từ PLANTABLE.

Chạy (từ thư mục gốc repo):

```
python3 bao-cao/scripts/tao_bao_cao.py "<đường dẫn file PLANTABLE .xlsm>"
```

Tuỳ chọn: `--ngay YYYY-MM-DD` (ngày tham chiếu, mặc định = ngày trong tên file + 1), `--ngay-file YYYY-MM-DD` (ngày của bản PLANTABLE).

Kết quả: `bao-cao/Bao_cao_dieu_hanh_PLANTABLE_<ngày>.xlsx`, `bao-cao/Keiei_Report_PLANTABLE_<ngày>_JP.xlsx`, và thêm 1 mốc vào `bao-cao/lich-su/do_phu_don_theo_ban_chup.csv` (xu hướng 受注進行表 theo tuần).

| File | Vai trò |
|---|---|
| `bao-cao/scripts/doc_plantable.py` | Đọc sheet PLANTABLE, làm sạch, nhận diện dòng chấm (QTY nền chấm = dòng thứ 2 của chuyền lớn AB, không cộng SL). |
| `bao-cao/scripts/build_report.py` | Dựng toàn bộ sheet (tóm tắt, CEO, GĐ nhà máy, QL sản xuất, mùa vụ, so sánh mùa, khách hàng, công suất, 受注進行表, chi phí & lô, rủi ro, chất lượng dữ liệu, Data, Phương pháp). |
| `bao-cao/scripts/juchu_plan.py` | Dựng 受注進行表 từ PLANTABLE: thanh line × thời gian, đã chốt / tạm tính (暫定). |
| `bao-cao/scripts/ja_tr.py` | Bản dịch tiếng Nhật (từ điển + quy tắc cho câu có số liệu). |

Quy ước chính: xưởng APIC (gộp AP1A/AP1B/AP1C), 30 line thường × 32 CN, line nhỏ 小ロット 16 CN; hàng SMS (mẫu chào hàng) tách khỏi hàng sản xuất; "đã chốt" trong 受注進行表 = đã bắt đầu may hoặc ghi chú NVL mới nhất có 生地 OK & 資材 OK (đã đối chiếu với 受注進行表 bản 09/11/2026).

Lưu ý: sheet 4A và lịch sử tự trượt theo ngày tham chiếu; một số mốc so sánh ở các sheet khác (27SS vs 26SS, T3–T6/2027, Tết 2027) đang cố định cho giai đoạn cuối 2026 và cần rà soát khi sang mùa mới.
