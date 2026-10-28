# Phân tích khối lượng công việc hiện tại – nhóm MD (merchandiser) APIC

Phạm vi: chỉ nhóm MD. Nhóm kỹ thuật và CAD sẽ phân tích riêng sau.

Cập nhật: 28/10/2026. Nguồn dữ liệu:
- `data/nguon/Thuc-te-BULK-19FW-27SS-20261012.xlsx`: hàng đại trà theo người, mùa, brand (số lượng, số mã, giây TB, ngày chuyền).
- `data/nguon/Thuc-te-mau-SPL-23FW-27FW-20261012.xlsx`: mẫu (1st/2nd/3rd, SMS) theo người, mùa, brand (bản cập nhật có thêm mùa 27FW), kèm bảng giờ dịch tài liệu cho mỗi mã theo brand.
- `data/nguon/Phan-cong-brand-MD-APIC-20260722.xlsx`: phân công brand của nhóm MD ngày 22/7/2026.
- `data/nguon/APIC-P-Plan-Plantable-2026-20261026.xlsm`: Plantable 26/10/2026 – kế hoạch sản xuất theo chuyền (ngày vào chuyền, kết thúc từng lô), sổ theo dõi đơn S27 (MD phụ trách từng đơn).
- `data/nguon/APIC-Sample-20261023.xlsx`: kế hoạch và thực tế may mẫu 2025–2026 (MD phụ trách, loại mẫu, ngày cần, trạng thái), công suất phòng mẫu.
- `data/nguon/Lich-nghi-thu-7-TIVN-APIC-2026-20260811.xlsx`: lịch nghỉ thứ 7 theo nhóm A/B (8/2026 – 3/2027).
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

## Tải theo tháng cả vòng 12 tháng (11/2026 – 10/2027)

Tính lại: `python3 data/khoi-luong/ke-hoach-nam.py` → `data/khoi-luong/ke-hoach-nam.json`.

**Vòng mùa** (theo quản lý): mùa FW may mẫu phát triển (1st–3rd) từ tháng 7, mẫu triển lãm (SMS) tháng 10–11, sản xuất tháng 4–9 năm sau. Mùa SS may mẫu phát triển từ tháng 1, SMS tháng 4–5, sản xuất tháng 10 – 3 năm sau. Vì vậy mỗi tháng luôn có **mẫu của một mùa chồng lên sản xuất của mùa khác**. Lịch thật năm trước trong Plantable cho thấy hàng FW bắt đầu vào chuyền từ tháng 2–3 (PPS, chuẩn bị từ trước đó), sớm hơn mốc tháng 4.

**Cách tính:**
- **Sản xuất**: mùa đã có kế hoạch (đến 27SS) dùng lịch thật trong Plantable 26/10/2026; mùa chưa có kế hoạch (27FW, 28SS) dùng lịch thật của mùa tương ứng năm trước (26FW, 27SS) dời 52 tuần. Người phụ trách theo các quyết định 10/2026 (NW, TNF 東京企画 → Khanh; Goldwin → Huy; YONEX từ 27SS → B.Trang).
- **Mẫu**: đơn còn tồn dồn vào 11/2026 + số đơn cùng tháng năm trước theo brand (file mẫu 2025–2026), trọng số theo loại mẫu như trước; Khanh nhận 1/3 mẫu TNF của B.Trang từ 12/2026.
- **Công suất theo ngày làm việc thực**: thứ 2 – thứ 7, nghỉ thứ 7 theo nhóm (file lịch nghỉ thứ 7 2026, đến 3/2027; sau đó giả định mỗi nhóm nghỉ 2 thứ 7/tháng), trừ ngày lễ. **Tết 2027 chưa có lịch công ty: tạm tính nghỉ 4–12/2/2027**; Giỗ Tổ 16/4/2027 cũng là giả định. Giờ/ngày: 9 (8 + 1 tăng ca), Khanh 8.
- Khanh hỗ trợ Luan ~25% (11–12/2026), ~35% từ 1/2027; Cuong, Luan mỗi người ~1 giờ/ngày cho việc giao thêm. Hệ số độ khó ALLTERRAIN/DESCENTE ×1,3.
- **1,0 = khối lượng trung bình cho 1 ngày công trong cả 12 tháng.** Tháng ít ngày làm (tháng 2 có Tết) thì cùng khối lượng sẽ ra chỉ số cao hơn.

**Lịch giai đoạn và ngày làm việc**

| Tháng | Giai đoạn theo vòng mùa | Ngày làm việc (nhóm A / nhóm B) |
|---|---|---|
| 11/26 | Mẫu 1st–3rd 27FW, SMS 27FW, Sản xuất 27SS | 23 / 23 |
| 12/26 | Sản xuất 27SS | 25 / 25 |
| 1/27 | Mẫu 1st–3rd 28SS, Sản xuất 27SS | 22 / 22 |
| 2/27 | Mẫu 1st–3rd 28SS, Sản xuất 27SS, PPS, chuẩn bị 27FW | 14 / 15 |
| 3/27 | Mẫu 1st–3rd 28SS, Sản xuất 27SS, PPS, chuẩn bị 27FW | 25 / 25 |
| 4/27 | Mẫu 1st–3rd 28SS, SMS 28SS, Sản xuất 27FW | 22 / 22 |
| 5/27 | Mẫu 1st–3rd 28SS, SMS 28SS, Sản xuất 27FW | 23 / 24 |
| 6/27 | Sản xuất 27FW | 24 / 24 |
| 7/27 | Mẫu 1st–3rd 28FW, Sản xuất 27FW | 25 / 25 |
| 8/27 | Mẫu 1st–3rd 28FW, Sản xuất 27FW, PPS, chuẩn bị 28SS | 24 / 24 |
| 9/27 | Mẫu 1st–3rd 28FW, Sản xuất 27FW, PPS, chuẩn bị 28SS | 22 / 22 |
| 10/27 | Mẫu 1st–3rd 28FW, SMS 28FW, Sản xuất 28SS | 24 / 24 |

**Chỉ số tải theo tháng** (ô in đậm: ≥ 1,4, vượt mức tăng ca 1 giờ/ngày)

| Tháng | Cuong | Luan | Khanh | B.Trang | Huy | Hien | Nhu |
|---|---|---|---|---|---|---|---|
| 11/26 | **1,66** | **1,81** | 0,94 | 1,36 | **1,63** | **1,41** | 1,09 |
| 12/26 | **1,95** | **1,54** | 1,22 | 0,55 | 0,41 | 1,32 | 0,31 |
| 1/27 | 1,38 | 1,30 | 0,88 | 0,16 | **1,57** | 0,87 | 1,21 |
| 2/27 | **1,93** | 0,91 | 0,99 | 0,17 | 0,70 | 0,88 | 1,22 |
| 3/27 | 0,81 | 1,30 | 0,93 | 0,10 | 0,43 | 0,60 | 1,13 |
| 4/27 | 0,96 | **1,69** | 1,12 | 0,42 | 1,03 | 0,55 | **2,57** |
| 5/27 | 0,86 | **1,53** | 1,04 | 0,40 | **1,70** | 0,89 | 0,87 |
| 6/27 | 0,97 | 1,12 | 1,22 | 0,78 | 1,31 | 1,35 | 0,27 |
| 7/27 | **2,89** | 0,71 | 0,75 | 0,74 | **1,51** | 0,61 | 0,86 |
| 8/27 | 0,90 | 0,75 | 0,60 | 0,42 | 1,34 | 0,56 | 0,82 |
| 9/27 | 1,02 | 0,65 | 0,46 | 0,41 | 0,89 | 0,48 | 0,55 |
| 10/27 | **1,57** | 0,88 | 1,01 | 0,65 | 1,14 | 1,27 | 0,84 |

Như: chỉ số chỉ tính phần brand; Như còn ~30% thời gian cho việc văn phòng, nên chỉ số thực ≈ chỉ số ÷ 0,7.

**Số mã vào chuyền / số đơn mẫu theo tháng**

| Tháng | Cuong | Luan | Khanh | B.Trang | Huy | Hien | Nhu |
|---|---|---|---|---|---|---|---|
| 11/26 | 4 / 100 | 19 / 7 | 1 / 12 | 1 / 121 | 9 / 65 | 8 / 41 | 4 / 35 |
| 12/26 | 21 / 15 | 14 / 13 | 3 / 17 | 0 / 27 | 1 / 25 | 12 / 26 | 2 / 24 |
| 1/27 | 11 / 12 | 7 / 34 | 1 / 2 | 1 / 5 | 25 / 8 | 10 / 14 | 12 / 21 |
| 2/27 | 9 / 15 | 2 / 13 | 0 / 11 | 0 / 5 | 2 / 15 | 8 / 4 | 8 / 4 |
| 3/27 | 6 / 19 | 12 / 27 | 0 / 14 | 0 / 7 | 3 / 20 | 2 / 21 | 10 / 11 |
| 4/27 | 2 / 44 | 22 / 13 | 0 / 15 | 0 / 33 | 0 / 46 | 0 / 30 | 40 / 10 |
| 5/27 | 9 / 19 | 23 / 5 | 0 / 15 | 0 / 33 | 9 / 62 | 5 / 33 | 11 / 8 |
| 6/27 | 11 / 15 | 16 / 2 | 0 / 21 | 3 / 24 | 17 / 20 | 11 / 50 | 2 / 10 |
| 7/27 | 11 / 71 | 9 / 8 | 0 / 17 | 5 / 28 | 18 / 24 | 7 / 7 | 9 / 10 |
| 8/27 | 10 / 6 | 7 / 22 | 2 / 4 | 6 / 6 | 16 / 19 | 2 / 13 | 5 / 16 |
| 9/27 | 10 / 8 | 4 / 20 | 0 / 6 | 4 / 17 | 8 / 19 | 0 / 21 | 5 / 9 |
| 10/27 | 2 / 112 | 8 / 7 | 1 / 38 | 1 / 59 | 7 / 44 | 4 / 48 | 7 / 14 |

**Tháng cao điểm của từng người** (3 tháng cao nhất)

| Người | Tháng cao điểm |
|---|---|
| Cuong | 12/26 (1,95), 2/27 (1,93), 7/27 (2,89) |
| Luan | 11/26 (1,81), 12/26 (1,54), 4/27 (1,69) |
| Khanh | 12/26 (1,22), 4/27 (1,12), 6/27 (1,22) |
| B.Trang | 11/26 (1,36), 6/27 (0,78), 7/27 (0,74) |
| Huy | 11/26 (1,63), 1/27 (1,57), 5/27 (1,70) |
| Hien | 11/26 (1,41), 12/26 (1,32), 6/27 (1,35) |
| Nhu | 1/27 (1,21), 2/27 (1,22), 4/27 (2,57) |

Nhận xét:
1. **Mỗi người có chu kỳ cao điểm riêng, lặp lại theo mùa**:
   - **Cuong**: 11–12 (mẫu triển lãm FW, 21 mã HELLY HANSEN vào chuyền tháng 12), tháng 2 (Tết, ít ngày làm), **tháng 7 (mẫu 1st FW của HELLY HANSEN, ~71 đơn mẫu)** và tháng 10 (mẫu triển lãm FW).
   - **Luan**: 11–12 và **3–5** (hàng FW TNF 富山企画 vào chuyền, 22–23 mã/tháng), dù đã có Khanh hỗ trợ.
   - **Huy**: 11, 1 (25 mã TNF 別注 vào chuyền), **5 (mẫu Goldwin, ~62 đơn)**, 7.
   - **Nhu**: **tháng 4 rất cao** (40 mã TNF 別注 富山 mùa FW vào chuyền cùng tháng), cộng thêm việc văn phòng.
   - **Hien**: 11–12, 6 (mẫu ~50 đơn), 10.
2. **Tháng 11 vẫn là tháng nặng nhất của cả nhóm** (mẫu triển lãm FW + đại trà SS bắt đầu + đơn mẫu tồn). Tháng 3 và tháng 8–9 nhẹ hơn.
3. **Tháng 2/2027 chỉ có 14–15 ngày làm** (tạm tính Tết 4–12/2): ai có mã vào chuyền hoặc mẫu trong tháng 2 (Cuong 1,93, Nhu 1,22) sẽ bị dồn. Nên kéo bớt việc lên tháng 1 hoặc lùi sang tháng 3.
4. **Cần lên kế hoạch trước cho tháng 4/2027 (Nhu)** và **tháng 7/2027 (Cuong)**: hai đỉnh này đến từ lịch năm trước nên gần như chắc sẽ lặp lại. Khanh và B.Trang đang thấp trong các tháng này (theo số liệu), là người có thể hỗ trợ tạm.
5. **Khanh** dao động 0,9–1,2 sau khi nhận NW, TNF 東京企画, 1/3 mẫu TNF và hỗ trợ Luan – mức hợp lý cho người làm 8 giờ.
6. **B.Trang thấp gần như cả năm** theo số liệu, nhưng **có thể do thiếu dữ liệu**: các brand mới (FOXFIRE, TATRAS, Kasco, YONEX) gần như chưa có lịch sản xuất năm trước, và mẫu TNF năm trước một phần do Khanh làm. Cần số liệu thật trước khi kết luận B.Trang còn dư sức.
7. Số liệu 4–10/2027 là **ước theo năm trước**; khi có kế hoạch 27FW/28SS thật thì chạy lại.

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
- **Mẫu 12/2026 – 10/2027 và sản xuất 27FW, 28SS** ước theo cùng kỳ năm trước; khi có kế hoạch thật thì thay vào.
- **Lịch làm việc**: lịch nghỉ thứ 7 chỉ có đến 3/2027 (sau đó giả định mỗi nhóm nghỉ 2 thứ 7/tháng); **ngày nghỉ Tết 2027 (tạm 4–12/2) và Giỗ Tổ (tạm 16/4/2027) là giả định**.
- **Giờ dịch/mã** trong bảng là ước tính chung theo brand. Giả định mỗi mã ở mỗi đợt (bulk, 1st–3rd, SMS) dịch một bộ tài liệu.

## Theo dõi định kỳ (theo sổ tay MD)

Sổ tay MD muốn MD chuyển từ "theo đuôi đơn hàng" sang "kiểm soát sản xuất" với 5 chức năng: kế hoạch (PPC), NPL (MRC), năng lực (CC), rủi ro (PRC), chuẩn hóa liên lạc (CS). Để làm được, cần đo khối lượng thường xuyên:

1. **Mỗi mùa**: cập nhật 2 file BULK/SPL (sheet từng người), chạy lại `build.py`, xem chỉ số tải trước khi giao brand mùa mới.
2. **Mỗi tuần** (thứ Sáu – Weekly MD Report): số mã đang mở, NPL chưa có thông tin, rủi ro đang mở theo từng người.
3. **Ghi giờ theo 5 chức năng của sổ tay** trong 2–4 tuần (mỗi người tự ghi, mức 30 phút). Hiện các bạn dành phần lớn thời gian cho dịch và NPL thiếu, còn ít thời gian cho kế hoạch, năng lực, rủi ro. Số liệu này cho biết cần giảm việc gì để chuyển dần lên Level 2–3.

## Câu hỏi mở

- Lịch nghỉ Tết 2027 và các ngày lễ của công ty (đang tạm tính Tết 4–12/2/2027).

- Kế hoạch mẫu 1st 28SS (số mã theo brand) khi có, để thay số ước.
- Việc điều phối của Cuong và việc quản lý giao cho Luan chiếm bao nhiêu giờ/ngày? (đang tạm tính ~1 giờ/ngày mỗi người)
