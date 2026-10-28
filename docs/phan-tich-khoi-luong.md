# Phân tích khối lượng công việc hiện tại – nhóm MD (merchandiser) APIC

Phạm vi: chỉ nhóm MD. Nhóm kỹ thuật và CAD sẽ phân tích riêng sau.

Cập nhật: 28/10/2026. Nguồn dữ liệu:
- `data/nguon/Thuc-te-BULK-19FW-27SS-20261012.xlsx`: hàng đại trà theo người, mùa, brand (số lượng, số mã, giây TB, ngày chuyền).
- `data/nguon/Thuc-te-mau-SPL-23FW-27FW-20261012.xlsx`: mẫu (1st/2nd/3rd, SMS) theo người, mùa, brand (bản cập nhật có thêm mùa 27FW), kèm bảng giờ dịch tài liệu cho mỗi mã theo brand.
- `data/nguon/Phan-cong-brand-MD-APIC-20260722.xlsx`: phân công brand của nhóm MD ngày 22/7/2026.
- `data/nguon/APIC-P-Plan-Plantable-2026-20261026.xlsm`: Plantable 26/10/2026 – kế hoạch sản xuất theo chuyền (ngày vào chuyền, kết thúc từng lô), sổ theo dõi đơn S27 (MD phụ trách từng đơn).
- `data/nguon/APIC-Sample-20261023.xlsx`: kế hoạch và thực tế may mẫu 2025–2026 (MD phụ trách, loại mẫu, ngày cần, trạng thái), công suất phòng mẫu.
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

**Chỉ số tải** = trung bình 4 yếu tố. Mỗi yếu tố tính bằng phần của người đó so với trung bình nhóm. **1,0 = mức trung bình của nhóm**, 1,3 = cao hơn trung bình 30%.

**Hệ số độ khó**: số liệu chỉ đếm số mã, số ngày, chưa thấy được mã nào khó. Quản lý đánh giá ALLTERRAIN và DESCENTE (của Hien) phức tạp hơn: tài liệu, rập không rõ, phải hỏi kỹ thuật nhiều; TIJ chuẩn bị yếu, hay thiếu NPL. Nên mỗi mã, mỗi ngày của hai brand này được tính **× 1,3**. Các brand khác × 1,0. Có thể chỉnh hệ số trong `HE_SO_KHO` ở `build.py` và `du-bao.py`.

**So với ai:** chỉ so trong 6 MD làm đủ kỳ: Cuong, Luan, Hien, Huy, B.Trang, Nhu.
- **Khanh** vừa đi làm lại được 2 tuần (giữa 10/2026) sau nghỉ thai sản, nên không đưa vào phép so (nếu đưa vào, mức trung bình bị kéo xuống và mọi người trông nặng hơn thực tế).
- **H.Trang** thuộc nhóm CAD, chỉ phụ thêm TNF HAT&CAP và 141, nên để sang phần phân tích CAD/kỹ thuật (làm sau).

Ngoài ra cần nhìn thêm:
- **Số brand**: nhiều brand thì nhiều đầu mối, nhiều quy cách tài liệu.
- **Việc cố định không theo mã** (báo cáo, họp, hành chính): lấy từ bảng phân tích nghiệp vụ 2026. Ví dụ Nhu còn lo xuất hàng, P/L, danh sách công nhân, nhận NPL từ HCM (khoảng 30% thời gian); Cuong kiểm kế hoạch sản xuất mẫu của tất cả brand.
- **So với công suất**: một mùa khoảng 6 tháng ≈ 1.250 giờ làm việc. Nếu việc dịch chỉ nên chiếm 30–40% thì mỗi người có khoảng **375–500 giờ dịch/mùa**.

**Kỳ tính ở phần "Kết quả"** là mùa mẫu 27SS vừa qua: bulk 26FW + 27SS, mẫu 27SS. Tình hình hiện tại (28/10/2026) khác: bulk 26FW đã xong trên 95%, sản xuất đã chuyển sang 27SS; phòng mẫu đang may mẫu triển lãm 27FW và mẫu PPS 27SS. Phần "Dự báo theo 2 giai đoạn" bên dưới tính theo tình hình hiện tại này.

## Kết quả

| Người | Bulk 26FW mã / SL / ngày chuyền | Bulk 27SS mã / SL / ngày chuyền | Mẫu 27SS mã / SL / ngày | Giờ dịch 27SS (ước) | Số brand | **Chỉ số tải** |
|---|---|---|---|---|---|---|
| Hien | 35 / 31.394 / 579 | 40 / 20.275 / 371 | 96 / 282 / 782 | 389 | 4 (ALLTERRAIN, TETON, MONTBELL, DESCENT) | **1,27** |
| Cuong | 81 / 42.691 / 717 | 48 / 48.820 / 706 | 89 / 192 / 738 | **582** | 4 (HH, NW, TNF東京, YONEX) | **1,25** |
| Luan | 89 / 179.012 / 2.922 | 45 / 207.757 / 2.471 | 25 / 52 / 104 | 268 | 1 (TNF 富山企画) | **1,13** |
| Huy | 77 / 50.760 / 935 | 43 / 26.439 / 451 | 80 / 146 / 492 | 407 | 4 (TNF別注 HOSOYA, GOLDWIN, NANGA, CANTERBURY) | **0,98** |
| B.Trang | 15 / 2.244 / 300 | 0 | 104 / 340 / 680 | 347 | 4 (TNF mẫu, Premium Snow 別注, FOXFIRE, TATRAS) | **0,84** |
| Nhu | 79 / 56.164 / 770 | 35 / 32.056 / 448 | 34 / 49 / 98 | 230 | 1 (TNF別注 SHIGEMOTO/NAKAO) | **0,52** (≈ 0,75 nếu tính 30% việc văn phòng) |
| *Khanh – vừa đi làm lại* | 0 | 0 | 6 / 11 / 22 | 20 | 1 | *không so* |

Chi tiết từng yếu tố (1,0 = trung bình):

| Người | Số mã | Ngày chuyền bulk | Ngày may mẫu | Giờ dịch |
|---|---|---|---|---|
| Hien | 1,4 | 0,6 | 1,8 | 1,2 |
| Cuong | 1,2 | 0,8 | 1,5 | 1,5 |
| Luan | 0,6 | **3,0** | 0,2 | 0,7 |
| Huy | 1,1 | 0,8 | 1,0 | 1,1 |
| B.Trang | 1,0 | 0,2 | 1,4 | 0,9 |
| Nhu | 0,6 | 0,7 | 0,2 | 0,6 |

### Trọng số là gì, có cần đổi không

Chỉ số tải đang coi 4 yếu tố **quan trọng ngang nhau** (mỗi yếu tố 25%). "Trọng số" là quyết định yếu tố nào được tính nặng hơn. Ví dụ nếu cho rằng dịch tài liệu là việc tốn sức nhất thì cho giờ dịch 40%, ba yếu tố kia mỗi cái 20%.

Thử 4 cách đặt trọng số:

| Người | Chia đều (25% mỗi yếu tố) | Ưu tiên giờ dịch (40%) | Ưu tiên ngày chuyền bulk (40%) | Ưu tiên ngày may mẫu (40%) |
|---|---|---|---|---|
| Hien | **1,27** | 1,26 | 1,14 | **1,38** |
| Cuong | 1,25 | **1,31** | 1,16 | 1,30 |
| Luan | 1,13 | 1,05 | **1,50** | 0,95 |
| Huy | 0,98 | 1,00 | 0,94 | 0,98 |
| B.Trang | 0,84 | 0,85 | 0,71 | 0,94 |
| Nhu | 0,52 | 0,54 | 0,55 | 0,46 |

Kết luận: **Hien và Cuong cao nhất, Nhu thấp nhất** dù đặt trọng số kiểu nào (bảng đã tính hệ số độ khó). Chỉ có vị trí của **Luan** thay đổi: Luan đứng đầu nếu coi theo dõi chuyền bulk là việc nặng nhất, đứng giữa nếu coi dịch hoặc mẫu nặng hơn. Lô của Luan rất lớn nhưng là 1 brand, mã lặp lại nhiều, nên mỗi ngày chuyền có lẽ tốn ít công hơn brand nhỏ nhiều mã. Vì vậy **giữ chia đều** cho đến khi có số liệu ghi giờ thực tế (mục "Theo dõi định kỳ") để đặt trọng số theo thực tế.

### Xu hướng số mã theo mùa (mã bulk + mã mẫu)

| Người | 24FW | 25SS | 25FW | 26SS | 26FW | 27SS | 27FW (chỉ mẫu) |
|---|---|---|---|---|---|---|---|
| Cuong | 65+65 | 53+93 | 74+131 | 49+56 | 81+113 | 48+89 | **170** |
| Huy | 24+29 | 32+46 | 19+101 | 33+103 | 77+127 | 43+80 | 62 |
| Khanh | 47+108 | 15+77 | 36+183 | 23+112 | 0+145 | 0+6 | 0 |
| Nhu | 20+95 | 1+37 | 66+26 | 49+63 | 79+49 | 35+34 | 9 |
| Hien | 22+66 | 10+47 | 20+12 | 27+51 | 35+0 | 40+96 | 51 |
| Luan | 83+0 | 88+39 | 86+61 | 41+14 | 89+15 | 45+25 | 12 |
| B.Trang | – | – | – | 2+0 | 15+0 | 0+104 | **155** |

### Mẫu mùa 27FW (thực tế, file SPL 12/10/2026) so với 26FW

| Người | 1st–3rd 26FW (cái / mã) | 1st–3rd 27FW | Triển lãm SMS 26FW (cái / mã / ngày may) | Triển lãm SMS 27FW | Ghi chú 27FW |
|---|---|---|---|---|---|
| Cuong | 70 / 63 | **96 / 86** | 159 / 50 / 678 | **298 / 84 / 1.130** | HELLY HANSEN 267 cái, NEUTRALWORKS 31; YONEX 1st 5 |
| B.Trang | – | 73 / 45 | – | **611 / 110 / 1.712** | TNF 367 cái (94 mã), **YONEX 244 cái** (16 mã) |
| Khanh | 86 / 49 | – | 480 / 96 / 2.400 | – | Nghỉ thai sản; mẫu TNF chuyển cho B.Trang |
| Huy | 169 / 91 | 41 / 32 | 205 / 36 / 831 | 209 / 30 / 1.015 | GOLDWIN 194 cái, CANTERBURY 15 |
| Hien | (sheet thiếu) | 29 / 27 | (sheet thiếu) | 158 / 24 / 714 | **SKI 82 cái, ~22.000 giây/cái** (6 ngày/cái), ALLTERRAIN 35, TETON 41 |
| Nhu | 89 / 49 | 9 / 9 | – | – | TNF 別注 |
| Luan | 18 / 15 | 12 / 12 | – | – | TNF 富山企画 |

Ngày may của mẫu chưa có giây TB được ước theo giây TB của brand mùa trước (YONEX chưa có nên tạm tính 1 ngày/cái, có thể thấp hơn thực tế).

Nhận xét về mẫu 27FW:
- **Cuong tăng mạnh**: SMS 27FW 298 cái (gần gấp đôi 26FW), cộng 96 mẫu 1st. Đây là lý do Cuong cao điểm từ tháng 9–11, nối tiếp cao điểm vào chuyền HELLY HANSEN tháng 12.
- **B.Trang (nhân viên mới) gánh đợt mẫu triển lãm lớn nhất nhóm**: 611 cái, trong đó YONEX 244 cái (brand mới nhận). Tổng mẫu TNF thấp hơn năm trước (367 so với 480 cái của Khanh mùa 26FW), nhưng có thêm YONEX.
- **Hien có mẫu SKI rất nặng** (82 cái, mỗi cái ~22.000 giây), nên đợt mẫu FW của Hien không nhẹ như số ước trước đây (36 cái).
- Các đơn mẫu triển lãm 27FW còn lại (chưa xong) đã nằm trong phần "Tải theo tháng" (dồn vào 11/2026), nên kết quả theo tháng không đổi.

## Nhận xét

1. **Hien và Cuong cao nhất trong mùa SS.** Hien (1,27) chủ yếu do mẫu 27SS (TETON SMS 157 cái, ALLTERRAIN SMS 81 cái) và độ khó của ALLTERRAIN/DESCENTE. Cuong (1,25): nhiều mã, nhiều mẫu, ước khoảng 580 giờ dịch cho 27SS, vượt mức 375–500 giờ. Cuong còn kiểm kế hoạch sản xuất mẫu của tất cả brand (không có trong số liệu), nên thực tế còn cao hơn. YONEX đã chuyển sang B.Trang từ S27. HH (HELLY HANSEN) là brand nặng nhất (4,3 giờ dịch/mã, 14 trang/mã) và đang tăng (mẫu 27SS: 34→45 cái, bulk 26FW: 40→56 mã).
2. **Luan tải kiểu khác** (1,2): chỉ 1 brand, ít mã, nhưng số lượng bulk rất lớn (27SS: 207.757 cái, khoảng 62% số lượng bulk của nhóm), ngày chuyền gấp 3 lần trung bình. Việc nặng của Luan là theo dõi chuyền, NPL, xuất hàng, không phải dịch.
3. **Huy ở mức trung bình** (0,98), nhưng tăng mạnh từ 26FW vì nhận thêm Goldwin đại trà trong lúc Khanh nghỉ.
4. **B.Trang (mới)** (0,9): nhận mẫu TNF (ALL) từ 27SS thay Khanh (340 cái, 104 mã) và 5 brand nhỏ (Premium Snow 別注, FOXFIRE, TATRAS, Kasco, YONEX từ S27). Chỉ số chưa tính hết YONEX/Kasco 27SS vì file chưa có số. Với người mới, khối lượng này là nặng.
5. **Nhu thấp nhất** về mã hàng (0,5), nhưng khoảng 30% thời gian dành cho việc hành chính, xuất hàng. Tính cả phần đó thì khoảng 0,8, vẫn thấp hơn nhóm.
6. **Khanh đi làm lại được 2 tuần.** Trước khi nghỉ, Khanh có tải mẫu lớn nhất nhóm (26FW: 566 cái, 145 mã, 2.916 ngày may) và làm Goldwin đại trà. Đây là nguồn lực có kinh nghiệm để cân bằng lại nhóm.
7. **Bảng phân tích nghiệp vụ 2026 chưa phản ánh thực tế**: bảng chia gần như đều % cho các merchandiser (dịch 15%, tài liệu mẫu 14%, NPL thiếu ~30%…), trong khi chỉ số tải thực tế chênh nhau từ 0,5 đến 1,3.

### Quyết định (10/2026)

- **Giờ làm**: mọi người 8 giờ + 1 giờ tăng ca/ngày (9 giờ). Khanh làm đến 16h, có thể tăng ca đến 17h (8 giờ, khoảng 8/9 công suất người khác).
- **Hien giữ nguyên** (Teton Bros, ALLTERRAIN, DESCENTE…), không giao thêm. Đơn của Hien chiếm khoảng 4/30 chuyền. Đơn ALLTERRAIN, DESCENTE phức tạp hơn (hệ số độ khó 1,3).
- **Huy giữ nguyên**: Goldwin xuyên suốt cả mẫu và đại trà.
- **Nhu giữ nguyên**, theo dõi thêm (ngoài brand còn việc văn phòng, báo cáo tháng, hóa đơn…).
- **Cuong**: chuyển NEUTRALWORKS sang Khanh (ngay), **TNF 東京企画 sang Khanh từ 11/2026** (quyết định 28/10/2026, đưa sớm hơn 1 tháng theo Plantable); giữ **HELLY HANSEN** (brand chủ lực). Cuong nhận thêm **điều phối chung**: theo dõi xuất hàng cho cả nhóm, quản lý chung việc may mẫu (đã có việc kiểm kế hoạch sản xuất mẫu của mọi brand).
- **Khanh**: nhận NEUTRALWORKS, **1/3 mẫu TNF của B.Trang từ mẫu 1st 28SS**, và **một phần việc của Luan**.
- **Mẫu triển lãm 27FW** đã vào giai đoạn may, chỉ còn theo dõi tiến độ và dịch góp ý gửi khách: **người đang làm làm tiếp đến hết**, không chia lại.
- **Luan**: nhận một phần việc quản lý; giảm bớt việc bằng cách chuyển một phần cho Khanh. Một brand khó tách cho hai người, nên tách theo chức năng (Luan giữ brand).
- **Hỗ trợ tạm theo tháng** (Huy, Nhu hỗ trợ Cuong tháng 12; B.Trang hỗ trợ Huy tháng 1): **chưa cần**, theo dõi thêm.
- **Snowpeak**: Hien theo dõi. 4 lô Snowpeak trên Plantable là dòng cũ (xem mục Dữ liệu cần xác nhận), đã loại khỏi phần tính.
- **Tình hình hiện tại (28/10/2026)**: sản xuất 26FW đã xong trên 95% (còn vài mã), đang sản xuất 27SS. Phòng mẫu đang may mẫu triển lãm 27FW và **mẫu PPS 27SS** (mỗi mã có đơn sản xuất may 1 cái, tái PPS/tái mockup khoảng 5%). Sau triển lãm là mẫu phát triển, mẫu 1st 28SS, cũng khá nhiều.

## Tải theo tháng từ kế hoạch thật (11/2026 – 2/2027)

Tính lại: `python3 data/khoi-luong/ke-hoach-thang.py` → `data/khoi-luong/ke-hoach-thang.json`.

**Cách tính** (thay cho phần ước theo 2 giai đoạn trước đây):
- **Sản xuất**: Plantable 26/10/2026, sheet PLANTABLE 2026 (lịch theo chuyền). MD của mỗi mã lấy theo người theo dõi mẫu PPS của mã đó trong file mẫu (257/371 lô khớp); không khớp thì theo brand. Mỗi tháng đếm:
  - **mã vào chuyền** (mỗi mã tính một lần dù chia nhiều lô/chuyền): việc chuẩn bị, duyệt PPS, kiểm NPL, họp trước khi vào chuyền, dịch tài liệu;
  - **ngày chuyền đang chạy** (theo từng lô): theo dõi tiến độ, NPL, sự cố, xuất hàng.
- **Mẫu**: đơn mẫu còn tồn trong file mẫu (chưa xong, hạn đến hết 10/2026) dồn vào tháng 11; cộng **số đơn mẫu cùng tháng năm trước** (11/2025 – 2/2026) của brand, gán cho người đang phụ trách brand đó (mẫu TNF của Khanh năm trước → B.Trang; NEUTRALWORKS → Khanh; YONEX từ S27 → B.Trang).
- **Trọng số theo loại mẫu**: 1st/0次/đề án tính đủ (theo dõi + dịch trọn bộ); 2nd ×0,8 (dịch bản sửa ×0,5); SMS ×0,5 và PPS/mockup ×0,6 (chủ yếu theo dõi, dịch góp ý ×0,2); ảnh, lưu mẫu, test ×0,3.
- Hệ số độ khó ALLTERRAIN/DESCENTE (DCC) × 1,3. Công suất 9 giờ/ngày, Khanh 8 giờ. Cuong, Luan mỗi người ~1 giờ/ngày cho việc giao thêm.
- **Mốc 1,0 = khối lượng trung bình của 1 người trong 1 tháng của cả giai đoạn 11/2026 – 2/2027**, dùng chung cho mọi tháng, nên so được tháng bận với tháng rảnh.

**Số liệu chính theo tháng** (chia theo người phụ trách trước khi chuyển TNF 東京企画; từ 11/2026 các lô TNF 東京企画 tính cho Khanh):

| Tháng | Mã vào chuyền | Đơn mẫu (tồn + ước) | Ghi chú |
|---|---|---|---|
| 11/2026 | Luan 9, Huy 8, Hien 8, Cuong 5, Nhu 3 | B.Trang 121, Cuong 100, Huy 65, Hien 41, Nhu 35 | **Cao điểm mẫu**: mẫu triển lãm 27FW + PPS 27SS + đơn tồn. Phòng mẫu SW1 tháng 10 đã vượt công suất (404 thực tế + 476 dự kiến so với 236/tháng) |
| 12/2026 | **Cuong 21** (HELLY HANSEN 27SS), Hien 12, Luan 10, B.Trang 5, Khanh 3 | B.Trang 27, Hien 26, Huy 25, Nhu 24 | **Cao điểm vào chuyền của Cuong** |
| 1/2027 | **Huy 25** (TNF 別注 HOSOYA), Nhu 11, Hien 11, Cuong 10, Luan 7 | Luan 34 (PPS), Nhu 21, Hien 14 | Cao điểm vào chuyền của Huy |
| 2/2027 | Cuong 9, Hien 8, Nhu 7 | Cuong 15, Huy 15, Luan 13 | Tết: ít việc |

Luan đang có **27–29 mã chạy chuyền cùng lúc** trong 11–12/2026 (TNF 富山企画, 67 lô), nhiều nhất nhóm.

**Chỉ số tải theo tháng** (ô in đậm: ≥ 1,4)

| Phương án | Tháng | Cuong | Luan | Khanh | B.Trang | Huy | Hien | Nhu |
|---|---|---|---|---|---|---|---|---|
| **Đã quyết** (NW và TNF 東京企画 → Khanh từ 11/2026; Khanh hỗ trợ Luan ~25% tháng 11–12, ~35% tháng 1–2; 1/3 mẫu TNF → Khanh từ 12/2026) | 11/2026 | **1,71** | 1,30 | 0,78 | **1,42** | **1,65** | **1,48** | 1,05 |
|  | 12/2026 | **2,21** | **1,46** | 1,25 | 0,89 | 0,45 | **1,50** | 0,35 |
|  | 01/2027 | 1,32 | 1,27 | 0,81 | 0,15 | **1,61** | 0,89 | 1,17 |
|  | 02/2027 | 1,31 | 0,46 | 0,65 | 0,11 | 0,48 | 0,54 | 0,81 |
| Để so sánh: chưa chia việc của Luan, TNF 東京企画 → Khanh từ 12/2026 | 11/2026 | **1,89** | **1,73** | 0,15 | **1,42** | **1,65** | **1,48** | 1,05 |
|  | 12/2026 | **2,21** | **1,94** | 0,76 | 0,89 | 0,45 | **1,50** | 0,35 |
|  | 01/2027 | 1,32 | **1,96** | 0,12 | 0,15 | **1,61** | 0,89 | 1,17 |
|  | 02/2027 | 1,31 | 0,71 | 0,40 | 0,11 | 0,48 | 0,54 | 0,81 |
| Tham khảo, **chưa áp dụng**: hỗ trợ tạm (12/2026 Huy 25% + Nhu 15% hỗ trợ Cuong; 1/2027 B.Trang hỗ trợ Huy 20%) | 12/2026 | **1,41** | **1,46** | 1,25 | 0,89 | 0,94 | **1,50** | 0,57 |
|  | 01/2027 | 1,32 | 1,27 | 0,81 | 0,48 | 1,29 | 0,89 | 1,17 |

Nhu: chỉ số chỉ tính phần brand; Nhu còn ~30% thời gian cho việc văn phòng, nên chỉ số thực ≈ chỉ số ÷ 0,7.

Nhận xét (theo các quyết định đã chốt):
1. **Tháng 11/2026 là cao điểm của cả nhóm** (mẫu triển lãm 27FW, PPS 27SS, đơn mẫu tồn): Cuong 1,71, Huy 1,65, Hien 1,48, B.Trang 1,42. Phòng mẫu cũng đang vượt công suất. Chia lại việc không đủ; cần **ưu tiên đơn mẫu theo hạn** và tăng ca.
2. **Cuong cao điểm tháng 12 (2,21)**: 21 mã HELLY HANSEN vào chuyền. Quản lý quyết **chưa cần hỗ trợ tạm**; nếu tháng 12 Cuong quá tải thì phương án dự phòng là Huy (0,45) và Nhu (0,35) hỗ trợ hậu cần các lô HELLY HANSEN (Cuong còn ~1,41).
3. **Hien tháng 11–12 khá cao (1,48–1,50)**: 12 mã vào chuyền tháng 12 (ALLTERRAIN/DESCENTE khó ×1,3) cùng mẫu triển lãm SKI. Không giao thêm việc cho Hien như đã quyết.
4. **Huy cao điểm tháng 1 (1,61)**: 25 mã TNF 別注 HOSOYA vào chuyền; chưa cần hỗ trợ, theo dõi thêm.
5. **Luan về ~1,27–1,46** nhờ Khanh hỗ trợ; vẫn là người có nhiều mã chạy chuyền nhất (27–29 mã cùng lúc). Nên giữ phần hỗ trợ suốt mùa đại trà 27SS.
6. **Khanh tháng 12 lên 1,25** (TNF 東京企画, NEUTRALWORKS, 1/3 mẫu TNF, hỗ trợ Luan). Ở mức tăng ca đến 17h; nên giữ hỗ trợ Luan khoảng 25% trong tháng 12.
7. **B.Trang và Huy thấp vào tháng 12–2** theo số liệu, nhưng có thể do thiếu dữ liệu (mẫu 1st 28SS chưa có kế hoạch, các brand mới của B.Trang gần như chưa có trong Plantable).

**Số đơn hàng (案件) S27 mỗi MD phụ trách** (sổ theo dõi đơn S27): Huy 12 đơn (46 mã, ~30.000 cái), Nhu 10 đơn (34 mã, ~26.500 cái), Cuong 7,5 (71 mã, ~149.000 cái), Hien 2 (39 mã), Luan 1,5 (26 mã, ~121.000 cái). **Huy và Nhu làm nhiều đơn 別注 nhỏ**: mỗi đơn đều phải làm bảng cấu tạo, chi tiết đặt hàng, kế hoạch MD, kiểm công suất, kế hoạch sản xuất, nên khối lượng thực tế cao hơn số cái cho thấy.

### Phân chia việc giữa Luan và Khanh

| | Luan | Khanh |
|---|---|---|
| Brand phụ trách | **TNF 富山企画** (PPS~đại trà), đầu mối duy nhất với TIJ | **NEUTRALWORKS** (nhận từ Cuong, ngay); **TNF 東京企画** (nhận từ Cuong từ 11/2026); **1/3 mẫu TNF** cùng B.Trang từ 1st 28SS |
| TNF 富山企画 – Luan giữ | Kiểm tài liệu; kiểm rập/mẫu gốc; báo NPL thiếu cho TI; xử lý sự cố; làm việc với TIJ và nhà máy | – |
| **11–12/2026** – Khanh hỗ trợ TNF 富山企画 (~25% việc Luan) | Duyệt kết quả, kiểm bản dịch | Báo cáo tháng mã ngoài quy tắc 2 tuần; tiến độ NPL, báo nhà máy; BL, P/L, số DHL (cc Luan); **dịch một phần tài liệu** |
| **1–2/2027** (Khanh đã quen việc, mẫu TNF ít hơn) | Như trên | Tăng phần hỗ trợ lên ~1/3 việc Luan (thêm dịch tài liệu, theo dõi NPL cho các lô đang chạy) |
| Việc khác | Một phần việc quản lý (~1 giờ/ngày) | – |

**Bàn giao TNF 東京企画 (Cuong → Khanh)**: Khanh theo cùng Cuong các lô bulk 27SS đang chạy và mẫu PPS để nắm tài liệu, NPL, liên lạc TIJ, rồi làm đầu mối từ 11/2026; Cuong hỗ trợ khi cần. Báo TIJ và nhà máy người phụ trách mới trước khi chuyển.

Quy tắc phối hợp (theo sổ tay MD, chương 7): mọi việc Khanh làm cho TNF 富山企画 đều ghi rõ nội dung, hạn, người phụ trách, ngày theo dõi tiếp; liên lạc với nhà máy luôn cc Luan; liên lạc với TIJ chỉ qua Luan.

### Phân công dự kiến sau thay đổi

| Người | Brand / việc |
|---|---|
| Cuong | HELLY HANSEN (chủ lực); TNF 東京企画 bàn giao cho Khanh từ 11/2026; điều phối xuất hàng chung cả nhóm; quản lý chung may mẫu (kiểm kế hoạch sản xuất mẫu mọi brand) |
| Luan | TNF 富山企画; một phần việc quản lý |
| Khanh | NEUTRALWORKS; TNF 東京企画 (PPS~đại trà); hỗ trợ TNF 富山企画 (hậu cần, dịch); từ 1st 28SS: 1/3 mẫu TNF |
| Hien | Giữ nguyên (Teton Bros, ALLTERRAIN, Mont-Bell, DESCENTE GOLF) |
| Huy | Giữ nguyên (Goldwin, TNF 別注 HOSOYA, 141 hỗ trợ, CANTERBURY, NANGA) |
| B.Trang | TNF (ALL) mẫu (làm tiếp mẫu triển lãm 27FW; từ 1st 28SS chia 1/3 cho Khanh), Premium Snow 別注, FOXFIRE, TATRAS, Kasco, YONEX |
| Nhu | Giữ nguyên (TNF 別注 SHIGEMOTO/NAKAO, UTMF, TOKYO MARATHON) + việc văn phòng |

**Hỗ trợ tạm theo tháng**: quản lý quyết **chưa cần** (28/10/2026). Phương án dự phòng nếu quá tải: 12/2026 Huy, Nhu hỗ trợ Cuong (hậu cần các lô HELLY HANSEN vào chuyền); 1/2027 B.Trang hỗ trợ Huy (các đơn TNF 別注 HOSOYA).

Lưu ý: các việc liên quan khai trị giá hải quan (ra rập, nhảy size, tài liệu may, phát triển SP) thay đổi người làm thì phải báo TI. Các việc trong bảng trên không thuộc nhóm này.

## Dữ liệu cần xác nhận

- **File SPL**: sheet "シーズンごと纏め" lệch nhiều so với sheet từng người (vd. Nhu 26SS: 346 cái/112 mã so với 146/63; Q.Huy 24FW: 414/110 so với 108/29). Phân tích dùng sheet từng người vì khớp với sheet "全体纏め" và có ghi chú cập nhật.
- **File BULK**: sheet "纏め" cũ hơn sheet từng người (thiếu YONEX của Cuong, FOXFIRE/TATRAS của B.Trang, Nhu 27SS 30.056 so với 32.056). Phân tích dùng sheet từng người.
- **Mẫu 27SS chưa có giây TB** cho nhiều brand, nên ngày may đang bằng số cái. Phân tích tạm lấy giây TB của brand mùa 26SS để ước (vd. HH 11.036 giây → 4 ngày/cái).
- **TNF別注 富山 (Nhu) và TNF別注 東京 (Q.Huy)**: ở các mùa 25FW, 26SS, số mẫu của hai bạn giống hệt nhau (vd. 146 cái/63 mã), có thể bị nhập trùng.
- **H.Trang bulk 26SS**: 384 cái, 1 mã nhưng 768 ngày chuyền (ghi chú "SPL Line"), số này chưa hợp lý. H.Trang thuộc nhóm CAD nên không tính vào phân tích MD.
- **Plantable có dòng cũ bị đẩy ngày**: sheet PLANTABLE 2026, chuyền AP1B-9, **dòng 1976–1996 (21 dòng)** là đơn cũ mùa 23SS/23FW (Snowpeak, Montbell, 1 dòng ALLTERRAIN; đơn năm 2022–2023). Ô START dòng 1976 bị gõ cứng 28/11/2026, các dòng sau nối tiếp bằng công thức (=AI dòng trên), nên cả chuỗi bị đẩy sang **28/11/2026 – 19/10/2027**. Từ dòng 1997 trở đi ngày cũ bình thường (2023–2024). Trong giai đoạn phân tích (11/2026 – 3/2027) có **8 dòng** rơi vào: 4 lô Snowpeak (dòng 1976–1979) và 4 lô Montbell (dòng 1980–1983); đã loại khỏi phần tính. Nên xóa hoặc sửa cả 21 dòng trong Plantable để không chiếm chỗ chuyền 9B.
- **Mẫu 12/2026 – 2/2027** ước theo cùng kỳ năm trước; khi có kế hoạch mẫu 1st 28SS thật thì thay vào.
- **Giờ dịch/mã** trong bảng là ước tính chung theo brand. Giả định mỗi mã ở mỗi đợt (bulk, 1st–3rd, SMS) dịch một bộ tài liệu.

## Theo dõi định kỳ (theo sổ tay MD)

Sổ tay MD muốn MD chuyển từ "theo đuôi đơn hàng" sang "kiểm soát sản xuất" với 5 chức năng: kế hoạch (PPC), NPL (MRC), năng lực (CC), rủi ro (PRC), chuẩn hóa liên lạc (CS). Để làm được, cần đo khối lượng thường xuyên:

1. **Mỗi mùa**: cập nhật 2 file BULK/SPL (sheet từng người), chạy lại `build.py`, xem chỉ số tải trước khi giao brand mùa mới.
2. **Mỗi tuần** (thứ Sáu – Weekly MD Report): số mã đang mở, NPL chưa có thông tin, rủi ro đang mở theo từng người.
3. **Ghi giờ theo 5 chức năng của sổ tay** trong 2–4 tuần (mỗi người tự ghi, mức 30 phút). Hiện các bạn dành phần lớn thời gian cho dịch và NPL thiếu, còn ít thời gian cho kế hoạch, năng lực, rủi ro. Số liệu này cho biết cần giảm việc gì để chuyển dần lên Level 2–3.

## Câu hỏi mở

- Kế hoạch mẫu 1st 28SS (số mã theo brand) khi có, để thay số ước.
- Việc điều phối của Cuong và việc quản lý giao cho Luan chiếm bao nhiêu giờ/ngày? (đang tạm tính ~1 giờ/ngày mỗi người)
