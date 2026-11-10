# APIC – 3 nhân sự ISO kiêm kế hoạch: bố trí có hợp lý không, có dư người không?

Nguồn: 3 báo cáo công việc hằng ngày 03–10/11/2026 (Thân 6 ngày; Vương – A3 7 ngày; báo cáo thứ ba chưa rõ tên 7 ngày).
Số liệu, công thức, giả định sửa được: `Phan-tich-cong-viec-ISO.xlsx` (tạo lại bằng `python3 scripts/xuat_phan_tich_iso.py`).
Handoff trước đó chỉ được dùng để tham khảo. Các điểm khác với handoff nằm ở mục "Phản biện handoff".

## Trả lời ngắn

**Có dư người, và bố trí hiện nay chưa hợp lý.**

| | Thấp | Cơ sở | Cao |
|---|---|---|---|
| Số người ISO cần (đúng chức năng, cả 3 xưởng) | 0,51 | 0,74 | 1,31 |
| Số người cho phần việc kế hoạch đang nằm trong 3 vị trí | 0,82 | 0,78 | 0,71 |
| **Tổng người cần** | **1,33** | **1,53** | **2,02** |
| **Dư so với 3 người hiện có** | **1,7** | **1,5** | **1,0** |

Kịch bản cao đã cộng dồn mọi giả định bất lợi: đánh giá 15 chuyền/ngày, ISO giữ 50% việc tài liệu, và giữ nguyên giờ tự khai của việc kế hoạch. Dù vậy vẫn dư khoảng 1 người.

## Vì sao dư

**Dư không phải vì ngồi không.** Báo cáo vẫn kín giờ. Nếu tính lại khối lượng theo số lượt việc thực tế, nội dung ghi được lấp khoảng 50–95% giờ khai, tùy giả định số phút cho mỗi lượt. Vì vậy báo cáo không chứng minh được ai thiếu việc.

**Dư vì làm thay việc của người khác:**

| Việc đang làm | Đúng ra thuộc về | Thân | Vương | BC3 |
|---|---|---|---|---|
| Tuần tra 5S, nội quy 2–4 lượt/ngày; kiểm bảng và bóp vật dụng kim loại (VDKL) | Quản lý chuyền kiểm hằng ngày; ISO chỉ đánh giá xác suất | 64% | 50% | 53% |
| Đăng ký BTP; hối thúc cắt – in – thêu – laze; năng suất; biểu mẫu; phụ liệu; vải thử máy | Chuyền trưởng, Kế hoạch, kho, kỹ thuật | 8% | 40% | 27% |
| Đọc mail, chuyển tin lỗi chuyền | Chuyền báo thẳng bộ phận xử lý | 12% | 5% | 8% |

Khi bỏ phần làm thay, việc ISO còn lại khoảng 0,5–1,3 người cho cả 3 xưởng. Việc trả về chuyền chỉ khoảng 10 phút mỗi chuyền mỗi ngày, không phải lý do để giữ người.

## Vì sao bố trí chưa hợp lý

1. **Chia người theo xưởng, không theo khối lượng.** A3 chỉ có 8 chuyền nhưng vẫn được bố trí 1 người như hai xưởng 10–11 chuyền.
2. **ISO tự làm rồi tự kiểm.** Người kiểm tra đang làm chính việc mình kiểm: kiểm bảng VDKL thay quản lý chuyền, tự đăng ký BTP, tự cấp vật tư. Như vậy không còn kiểm tra độc lập (ISO 9001 mục 9.2.2 c).
3. **Hai kênh chỉ đạo song song.** Cắt, in, thêu nhận hối thúc từ cả ISO lẫn Kế hoạch. Vương còn ghi việc "phân chia kế hoạch chuyền", tức là vượt quyền.
4. **Không có mục tiêu, chỉ tiêu, danh sách lỗi chưa đóng hay hạn đóng.**
   - 60% lỗi chỉ xử lý bằng nhắc nhở. 84% lỗi báo "đạt" ngay trong ngày, nhưng 46% lặp lại ở cùng khu vực và cùng chủ đề.
   - Kiểm VDKL xác suất luôn ghi "Đạt" mà không có số lượng.
   - Nhiệt độ / độ ẩm chuyền 14A ghi y hệt 26°C / 73% trong 5 ngày.
   - 52–61% dòng báo cáo chép nguyên văn từ ngày khác.

## Đề xuất

**Phương án 2 (khuyến nghị):**
- 1 người ISO chung cho 3 xưởng, trực thuộc cấp trên xưởng. Người này chỉ làm: đánh giá xác suất, xác minh VDKL có số lượng, kiểm soát tài liệu hệ thống, theo dõi lỗi chưa đóng và báo cáo tuần.
- 1 người chuyển sang Kế hoạch: điều độ đầu vào (cắt – in – thêu – laze) và phân phối tài liệu sản xuất theo mã hàng cho 3 xưởng.
- 1 người điều chuyển.

**Phương án 3:** chỉ giữ 1 người ISO, nếu Kế hoạch xác nhận tự nhận được khoảng 0,8 người việc.

**Điều kiện:** giao trách nhiệm và chỉ tiêu cho quản đốc, chuyền trưởng trước khi rút người (5S và VDKL hằng ngày, đăng ký BTP, báo lỗi trực tiếp). Rút người trước khi giao trách nhiệm thì kỷ luật hiện trường sẽ tụt.

**Việc an toàn cần sửa ngay**, không chờ lịch khách đến:
- Vạch lối thoát hiểm ở A3.
- Vòng chắn kim bị gãy ở chuyền 12A, 15A, 15B.
- Ổ điện rớt ở nhiều chuyền.

## Phản biện handoff

| Handoff | Phản biện |
|---|---|
| Chưa chốt giảm người; thử 10 ngày rồi xem giữ 3, 2 hay 1 | Cả 3 kịch bản đều dưới 3 người, nên loại ngay phương án giữ 3. Chạy thử chỉ để chọn giữa 1 và 2. |
| Một đầu mối ISO chung và một người liên hệ tại mỗi xưởng từ nhân sự sẵn có | Nếu "người liên hệ" lại chính là 3 người này thì vị trí cũ vẫn còn nguyên, chỉ đổi tên. Người liên hệ phải là quản đốc hoặc phó quản đốc. |
| Bước tiếp theo là đối chiếu hồ sơ cấp BTP | Việc này dùng để thiết kế quy trình BTP, không quyết định ISO có dư hay không. Không để quyết định tổ chức chờ bước này. |
| Ban hành tài liệu là việc ISO | Tài liệu ban hành hằng ngày là tài liệu sản xuất theo mã (list đổ hàng, TCC), thuộc luồng Kỹ thuật / Kế hoạch. ISO chỉ giữ tài liệu hệ thống. |
| Đo bằng tần suất ghi nhận (số ngày có việc) | Tần suất không phản ánh khối lượng. Cần đo theo thời gian, số lượt và đầu ra. |
| Duy trì đủ các lượt kiểm tra bắt buộc | Chưa xác định lượt nào bắt buộc. Cần lập danh mục yêu cầu của khách hàng (Toray). |
| Chuyển việc về Kế hoạch | Chưa kiểm tra Kế hoạch còn công suất không. Phần việc chuyển sang tương đương khoảng 0,7–0,8 người. |
| (không nêu) | Các việc an toàn đang bị trì hoãn. Gốc vấn đề là tuyến chuyền chưa chịu trách nhiệm, chứ không chỉ là thừa người. |

## Gợi ý bố trí từng người

Chỉ dựa trên báo cáo, cần quản đốc xác nhận.

- **Thân:** 79% thời gian là việc kiểu ISO. Là người duy nhất ghi rõ lỗi gốc từ ngày nào và hẹn giờ kiểm tra lại. Tuy vậy vẫn chủ yếu nhắc nhở và đe dọa xử lý ABC. Phù hợp với vị trí ISO chung nếu được đào tạo đánh giá xác suất.
- **Vương (A3):** 40% thời gian là điều phối kế hoạch. Có biện pháp xử lý gốc: nhắc trả rập tự động, đổi ra bàn ủi sang màu tối. Phù hợp với vị trí điều độ / kế hoạch. Cần làm rõ quyền đang có.
- **Báo cáo 3:** chưa rõ người viết. Có cách làm đúng hướng: để quản lý tự khắc phục và gửi ảnh. Nhưng 61% dòng chép lại và có 4 việc hẹn kiểm tra lại mà không thấy theo dõi.

## Giới hạn

- Giờ là giờ tự khai theo khung cố định.
- Số phút cho mỗi lượt việc và quy mô đánh giá là giả định (ô vàng trong sheet DinhBien và NoiDungThucChat). Nên đo bằng chụp ảnh làm việc (work sampling): 20–30 thời điểm ngẫu nhiên mỗi ngày trong 10 ngày.

## File Excel

| Sheet | Nội dung |
|---|---|
| KetLuan | Trả lời ngắn theo 3 kịch bản, chỉ số chính, kết luận |
| ThoiGian | Phút/ngày, % thời gian của 13 mã việc × 3 người |
| PhanBien | 15 điểm phản biện bố trí hiện nay và handoff; gợi ý bố trí từng người |
| TrungLap | Ma trận việc × bộ phận (chủ trì, phối hợp, ISO làm thay) |
| GiuChuyenBo | Giữ / chuyển / bỏ: người nhận, quyền, đầu ra, tiêu chí nghiệm thu, ô chốt |
| ChatLuongKiemTra | Biện pháp, trạng thái, chủ đề lỗi, tỷ lệ tái diễn |
| NoiDungThucChat | Khối lượng việc ước tính theo số lượt, so với giờ khai |
| DinhBien | 3 kịch bản, so sánh phương án PA1 / PA2 / PA3 |
| NhatKy, LoiHienTruong | Dữ liệu từng dòng, có ô nguồn |
| DoiChieuBTP | Mẫu đối chiếu hồ sơ cấp BTP |
| GhiChuDuLieu | Giả định, sai lệch dữ liệu |
