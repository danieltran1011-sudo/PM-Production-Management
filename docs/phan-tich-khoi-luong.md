# Phân tích khối lượng công việc hiện tại – nhóm merchandiser APIC

Cập nhật: 27/10/2026. Nguồn dữ liệu:
- `data/nguon/Thuc-te-BULK-19FW-27SS-20261012.xlsx`: hàng đại trà theo người, mùa, brand (số lượng, số mã, giây TB, ngày chuyền).
- `data/nguon/Thuc-te-mau-SPL-23FW-27SS-20261012.xlsx`: mẫu (1st/2nd/3rd, SMS) theo người, mùa, brand, kèm bảng giờ dịch tài liệu cho mỗi mã theo brand.
- `data/nguon/APIC-MD-Best-Practice-Handbook.docx`: định hướng MD (5 chức năng, KPI, chu kỳ tuần).
- Bảng phân tích nghiệp vụ APIC 2026 上期 (`docs/phan-cong-cong-viec.md`).

Tính lại: `python3 data/khoi-luong/build.py` → `data/khoi-luong/hien-tai.json`.

## Cách phân tích

Khối lượng của một merchandiser chủ yếu sinh ra từ **mã hàng** họ phụ trách. Mỗi mã kéo theo việc dịch tài liệu, kiểm rập/sơ đồ, theo dõi NPL, theo dõi chuyền. Vì vậy đo theo 4 yếu tố, tính cho từng người:

| Yếu tố | Đo bằng | Ý nghĩa |
|---|---|---|
| 1. Số mã | Mã bulk 27SS + mã mẫu 27SS | Số bộ tài liệu, số lần trao đổi với TIJ/nhà máy |
| 2. Ngày chuyền bulk | 稼働日数 bulk 26FW (đang chạy) + 27SS (đang chuẩn bị) | Thời gian phải theo dõi chuyền, NPL, sự cố, xuất hàng |
| 3. Ngày may mẫu | 稼働日数 mẫu 27SS | Theo dõi tổ mẫu, comment, gửi mẫu |
| 4. Giờ dịch | Số mã 27SS × giờ dịch/mã của brand | Việc dịch chiếm 30–40% thời gian theo bảng của công ty |

**Chỉ số tải** = trung bình 4 yếu tố, mỗi yếu tố tính theo phần của người đó so với trung bình nhóm. **1,0 = mức trung bình của nhóm**, 2,0 = gấp đôi trung bình.

Ngoài ra cần nhìn thêm:
- **Số brand**: nhiều brand thì nhiều đầu mối, nhiều quy cách tài liệu.
- **Việc cố định không theo mã** (báo cáo, họp, hành chính): lấy từ bảng phân tích nghiệp vụ 2026. Ví dụ Nhu còn lo xuất hàng, P/L, danh sách công nhân, nhận NPL từ HCM.
- **So với công suất**: một mùa khoảng 6 tháng ≈ 1.250 giờ làm việc. Nếu việc dịch chỉ nên chiếm 30–40% thì mỗi người có khoảng **375–500 giờ dịch/mùa**.

"Hiện tại" (cuối 10/2026) gồm: bulk 26FW đang sản xuất, bulk 27SS đang chuẩn bị, mẫu 27SS đang làm.

## Kết quả

| Người | Bulk 26FW mã / SL / ngày chuyền | Bulk 27SS mã / SL / ngày chuyền | Mẫu 27SS mã / SL / ngày | Giờ dịch 27SS (ước) | Số brand | **Chỉ số tải** |
|---|---|---|---|---|---|---|
| Cuong | 81 / 42.691 / 717 | 48 / 48.820 / 706 | 89 / 192 / 738 | **582** | 4 (HH, NW, TNF東京, YONEX) | **1,7** |
| Luan | 89 / 179.012 / 2.922 | 45 / 207.757 / 2.471 | 25 / 52 / 104 | 268 | 1 (TNF GTC) | **1,5** |
| Hien | 35 / 31.394 / 579 | 40 / 20.275 / 371 | 96 / 282 / 782 | 389 | 4 (ALLTERRAIN, TETON, MONTBELL, DESCENT) | **1,5** |
| Q.Huy | 77 / 50.760 / 935 | 43 / 26.439 / 451 | 80 / 146 / 492 | 407 | 4 (TNF別注東京, GOLDWIN, NANGA, CANTERBURY) | **1,3** |
| B.Trang | 15 / 2.244 / 300 | 0 | 104 / 340 / 680 | 347 | 4 (TNF mẫu, TNF別注, FOXFIRE, TATRAS) | **1,1** |
| Nhu | 79 / 56.164 / 770 | 35 / 32.056 / 448 | 34 / 49 / 98 | 230 | 1 (TNF別注富山) | **0,7** |
| H.Trang | 0 | 0 | 8 / 24 / 96 | 13 | 1 (141) | **0,1** |
| Khanh | 0 | 0 | 6 / 11 / 22 | 20 | 1 (TNF mẫu) | **0,1** |

Chi tiết từng yếu tố (1,0 = trung bình):

| Người | Số mã | Ngày chuyền bulk | Ngày may mẫu | Giờ dịch |
|---|---|---|---|---|
| Cuong | 1,7 | 1,1 | 2,0 | 2,1 |
| Luan | 0,9 | **4,0** | 0,3 | 0,9 |
| Hien | 1,7 | 0,7 | 2,1 | 1,4 |
| Q.Huy | 1,5 | 1,0 | 1,3 | 1,4 |
| B.Trang | 1,3 | 0,2 | 1,8 | 1,2 |
| Nhu | 0,8 | 0,9 | 0,3 | 0,8 |

### Xu hướng số mã theo mùa (mã bulk + mã mẫu)

| Người | 24FW | 25SS | 25FW | 26SS | 26FW | 27SS |
|---|---|---|---|---|---|---|
| Cuong | 65+65 | 53+93 | 74+131 | 49+56 | 81+113 | 48+89 |
| Q.Huy | 24+29 | 32+46 | 19+101 | 33+103 | 77+127 | 43+80 |
| Khanh | 47+108 | 15+77 | 36+183 | 23+112 | 0+145 | 0+6 |
| Nhu | 20+95 | 1+37 | 66+26 | 49+63 | 79+49 | 35+34 |
| Hien | 22+66 | 10+47 | 20+12 | 27+51 | 35+0 | 40+96 |
| Luan | 83+0 | 88+39 | 86+61 | 41+14 | 89+15 | 45+25 |
| B.Trang | – | – | – | 2+0 | 15+0 | 0+104 |
| H.Trang | – | 6+8 | 2+14 | 1+39 | 0+2 | 0+8 |

## Nhận xét

1. **Cuong đang tải cao nhất**: nhiều mã nhất ở cả bulk và mẫu, 4 brand, ước khoảng 580 giờ dịch cho 27SS, vượt mức 375–500 giờ. HH là brand nặng nhất (4,3 giờ dịch/mã, 14 trang/mã) và đang tăng (mẫu 27SS: 34→45 cái, bulk 26FW: 40→56 mã).
2. **Luan tải kiểu khác**: chỉ 1 brand, ít mã, nhưng số lượng bulk rất lớn (27SS: 207.757 cái, khoảng 62% số lượng bulk của nhóm), ngày chuyền gấp 4 lần trung bình. Việc nặng của Luan là theo dõi chuyền, NPL, xuất hàng, không phải dịch.
3. **Hien và Q.Huy trên trung bình**: Hien chủ yếu do mẫu 27SS (TETON SMS 157 cái); Q.Huy tăng mạnh từ 26FW vì nhận thêm bulk GOLDWIN (trước là Khanh).
4. **B.Trang (mới) đã nhận phần mẫu TNF** trước đây của Khanh (27SS: 340 cái, 104 mã). Mức tải gần trung bình.
5. **Nhu thấp hơn** về mã hàng, nhưng còn việc hành chính, xuất hàng (bảng nghiệp vụ 2026 ghi khoảng 30% thời gian), nên tổng thể có thể gần trung bình.
6. **Khanh và H.Trang gần như không có mã ở 27SS.** Khanh từng có tải mẫu lớn nhất nhóm (26FW: 566 cái, 145 mã), nay chuyển sang B.Trang và Q.Huy. Cần xác nhận hai bạn đang làm việc gì khác (xem câu hỏi bên dưới).
7. **Bảng phân tích nghiệp vụ 2026 chưa phản ánh thực tế**: bảng chia gần như đều % cho các merchandiser (dịch 15%, tài liệu mẫu 14%, NPL thiếu ~30%…), trong khi chỉ số tải thực tế chênh nhau từ 0,7 đến 1,7.

### Gợi ý cân bằng

- Giảm cho Cuong: chuyển 1 brand nhỏ (YONEX hoặc NW) hoặc một phần mẫu HH sang người đang tải thấp (Nhu, Khanh, H.Trang nếu còn trống).
- Hỗ trợ Luan theo việc, không theo brand: tách phần theo dõi NPL/xuất hàng một phần lô lớn cho người khác.
- Đưa bảng phân tích nghiệp vụ (業務分析表) kỳ sau theo số liệu thực tế. Lưu ý: các việc liên quan khai trị giá hải quan (ra rập, nhảy size, tài liệu may, phát triển SP) thay đổi người làm thì phải báo TI.

## Dữ liệu cần xác nhận

- **File SPL**: sheet "シーズンごと纏め" lệch nhiều so với sheet từng người (vd. Nhu 26SS: 346 cái/112 mã so với 146/63; Q.Huy 24FW: 414/110 so với 108/29). Phân tích dùng sheet từng người vì khớp với sheet "全体纏め" và có ghi chú cập nhật.
- **File BULK**: sheet "纏め" cũ hơn sheet từng người (thiếu YONEX của Cuong, FOXFIRE/TATRAS của B.Trang, Nhu 27SS 30.056 so với 32.056). Phân tích dùng sheet từng người.
- **Mẫu 27SS chưa có giây TB** cho nhiều brand, nên ngày may đang bằng số cái. Phân tích tạm lấy giây TB của brand mùa 26SS để ước (vd. HH 11.036 giây → 4 ngày/cái).
- **TNF別注 富山 (Nhu) và TNF別注 東京 (Q.Huy)**: ở các mùa 25FW, 26SS, số mẫu của hai bạn giống hệt nhau (vd. 146 cái/63 mã), có thể bị nhập trùng.
- **H.Trang bulk 26SS**: 384 cái, 1 mã nhưng 768 ngày chuyền (ghi chú "SPL Line"), số này chưa hợp lý nên không tính vào mùa hiện tại.
- **Giờ dịch/mã** trong bảng là ước tính chung theo brand. Giả định mỗi mã ở mỗi đợt (bulk, 1st–3rd, SMS) dịch một bộ tài liệu.

## Theo dõi định kỳ (theo sổ tay MD)

Sổ tay MD muốn MD chuyển từ "theo đuôi đơn hàng" sang "kiểm soát sản xuất" với 5 chức năng: kế hoạch (PPC), NPL (MRC), năng lực (CC), rủi ro (PRC), chuẩn hóa liên lạc (CS). Để làm được, cần đo khối lượng thường xuyên:

1. **Mỗi mùa**: cập nhật 2 file BULK/SPL (sheet từng người), chạy lại `build.py`, xem chỉ số tải trước khi giao brand mùa mới.
2. **Mỗi tuần** (thứ Sáu – Weekly MD Report): số mã đang mở, NPL chưa có thông tin, rủi ro đang mở theo từng người.
3. **Ghi giờ theo 5 chức năng của sổ tay** trong 2–4 tuần (mỗi người tự ghi, mức 30 phút). Hiện các bạn dành phần lớn thời gian cho dịch và NPL thiếu, còn ít thời gian cho kế hoạch, năng lực, rủi ro. Số liệu này cho biết cần giảm việc gì để chuyển dần lên Level 2–3.

## Câu hỏi mở

- Khanh và H.Trang hiện làm việc gì (CAD, việc khác, nghỉ)? H.Trang có phải là "Trang" (CAD) trong bảng nghiệp vụ 2026 không?
- Có tính Vy, nhóm kỹ thuật, nhóm CAD vào phân tích không? Hiện 2 file chỉ có nhóm merchandiser.
- Muốn trọng số 4 yếu tố khác nhau không (vd. tăng trọng số giờ dịch)?
