# Phân tích khối lượng công việc hiện tại – nhóm MD (merchandiser) APIC

Phạm vi: chỉ nhóm MD. Nhóm kỹ thuật và CAD sẽ phân tích riêng sau.

Cập nhật: 27/10/2026. Nguồn dữ liệu:
- `data/nguon/Thuc-te-BULK-19FW-27SS-20261012.xlsx`: hàng đại trà theo người, mùa, brand (số lượng, số mã, giây TB, ngày chuyền).
- `data/nguon/Thuc-te-mau-SPL-23FW-27SS-20261012.xlsx`: mẫu (1st/2nd/3rd, SMS) theo người, mùa, brand, kèm bảng giờ dịch tài liệu cho mỗi mã theo brand.
- `data/nguon/Phan-cong-brand-MD-APIC-20260722.xlsx`: phân công brand của nhóm MD ngày 22/7/2026.
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

"Hiện tại" (cuối 10/2026) gồm: bulk 26FW đang sản xuất, bulk 27SS đang chuẩn bị, mẫu 27SS đang làm.

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

| Người | 24FW | 25SS | 25FW | 26SS | 26FW | 27SS |
|---|---|---|---|---|---|---|
| Cuong | 65+65 | 53+93 | 74+131 | 49+56 | 81+113 | 48+89 |
| Huy | 24+29 | 32+46 | 19+101 | 33+103 | 77+127 | 43+80 |
| Khanh | 47+108 | 15+77 | 36+183 | 23+112 | 0+145 | 0+6 |
| Nhu | 20+95 | 1+37 | 66+26 | 49+63 | 79+49 | 35+34 |
| Hien | 22+66 | 10+47 | 20+12 | 27+51 | 35+0 | 40+96 |
| Luan | 83+0 | 88+39 | 86+61 | 41+14 | 89+15 | 45+25 |
| B.Trang | – | – | – | 2+0 | 15+0 | 0+104 |

## Nhận xét

1. **Hien và Cuong cao nhất trong mùa SS.** Hien (1,27) chủ yếu do mẫu 27SS (TETON SMS 157 cái, ALLTERRAIN SMS 81 cái) và độ khó của ALLTERRAIN/DESCENTE. Cuong (1,25): nhiều mã, nhiều mẫu, ước khoảng 580 giờ dịch cho 27SS, vượt mức 375–500 giờ. Cuong còn kiểm kế hoạch sản xuất mẫu của tất cả brand (không có trong số liệu), nên thực tế còn cao hơn. YONEX đã chuyển sang B.Trang từ S27. HH (HELLY HANSEN) là brand nặng nhất (4,3 giờ dịch/mã, 14 trang/mã) và đang tăng (mẫu 27SS: 34→45 cái, bulk 26FW: 40→56 mã).
2. **Luan tải kiểu khác** (1,2): chỉ 1 brand, ít mã, nhưng số lượng bulk rất lớn (27SS: 207.757 cái, khoảng 62% số lượng bulk của nhóm), ngày chuyền gấp 3 lần trung bình. Việc nặng của Luan là theo dõi chuyền, NPL, xuất hàng, không phải dịch.
3. **Huy ở mức trung bình** (0,98), nhưng tăng mạnh từ 26FW vì nhận thêm Goldwin đại trà trong lúc Khanh nghỉ.
4. **B.Trang (mới)** (0,9): nhận mẫu TNF (ALL) từ 27SS thay Khanh (340 cái, 104 mã) và 5 brand nhỏ (Premium Snow 別注, FOXFIRE, TATRAS, Kasco, YONEX từ S27). Chỉ số chưa tính hết YONEX/Kasco 27SS vì file chưa có số. Với người mới, khối lượng này là nặng.
5. **Nhu thấp nhất** về mã hàng (0,5), nhưng khoảng 30% thời gian dành cho việc hành chính, xuất hàng. Tính cả phần đó thì khoảng 0,8, vẫn thấp hơn nhóm.
6. **Khanh đi làm lại được 2 tuần.** Trước khi nghỉ, Khanh có tải mẫu lớn nhất nhóm (26FW: 566 cái, 145 mã, 2.916 ngày may) và làm Goldwin đại trà. Đây là nguồn lực có kinh nghiệm để cân bằng lại nhóm.
7. **Bảng phân tích nghiệp vụ 2026 chưa phản ánh thực tế**: bảng chia gần như đều % cho các merchandiser (dịch 15%, tài liệu mẫu 14%, NPL thiếu ~30%…), trong khi chỉ số tải thực tế chênh nhau từ 0,5 đến 1,3.

### Quyết định (27/10/2026)

- **Giờ làm**: mọi người 8 giờ + 1 giờ tăng ca/ngày là hợp lý (9 giờ). Khanh làm đến 16h, có thể tăng ca đến 17h (8 giờ, tức khoảng 8/9 công suất của người khác).
- **NEUTRALWORKS chuyển từ Cuong sang Khanh.**
- **Huy giữ nguyên**: Huy làm Goldwin xuyên suốt (mẫu và đại trà) để nắm brand kỹ hơn. Khanh không nhận lại Goldwin.
- **Khanh nhận 1/3 mẫu triển lãm TNF của B.Trang.**
- **Quản lý sẽ giao một phần việc của mình cho Luan và Cuong**, nên cần giảm tải cho Luan (và Cuong).
- **Hien**: từ 26FW theo DESCENTE, ALLTERRAIN; đơn hàng phức tạp hơn (xem hệ số độ khó).
- **Nhu giữ nguyên**, theo dõi thêm: ngoài brand, Nhu làm việc văn phòng, báo cáo tháng, hóa đơn…
- **Mẫu theo mùa vụ**: mẫu triển lãm (SMS) lúc rất bận, lúc rảnh. Hiện đang làm mẫu triển lãm 27FW; xong sẽ chuyển sang phát triển mẫu 1st 28SS, cũng khá nhiều.

## Dự báo giai đoạn tới (khoảng 11/2026 – 2/2027)

File SPL mới có số liệu đến 27SS, nên phần đang làm (mẫu triển lãm 27FW) và sắp làm (1st 28SS) phải ước:
- **Mẫu triển lãm 27FW** ≈ SMS 26FW của cùng brand (cùng mùa FW). Sheet Hien thiếu mẫu 26FW, nên bù bằng số ở sheet tổng hợp (36 cái, 8 mã, tạm coi là ALLTERRAIN).
- **Mẫu 1st 28SS** ≈ 1st–3rd 27SS của cùng brand.
- **Bulk 27SS**: số thực tế.
- Brand gán cho người đang phụ trách (bảng 22/7/2026 và các quyết định ở trên). Hệ số độ khó ALLTERRAIN/DESCENTE × 1,3.
- Công suất: 9 giờ/ngày (8 + 1 tăng ca); Khanh 8 giờ. Mức trung bình tính trên cả 7 người theo công suất, nên chuyển việc giữa hai người không làm đổi chỉ số của người khác.
- **Việc quản lý giao cho Luan, Cuong**: chưa biết cụ thể, tạm tính mỗi người mất **10% thời gian (~0,9 giờ/ngày)**. Phần thời gian còn lại cho brand ít đi nên chỉ số tăng (chia cho 0,9).
- **Nhu**: chỉ số chỉ tính phần brand. Nhu dành ~30% thời gian cho văn phòng, báo cáo, hóa đơn, nên **chỉ số thực ≈ chỉ số ÷ 0,7** (ghi trong ngoặc).

Tính lại: `python3 data/khoi-luong/du-bao.py` → `data/khoi-luong/du-bao.json`.

| Phương án | Cuong | Luan | Hien | Huy | B.Trang | Nhu (thực) | Khanh |
|---|---|---|---|---|---|---|---|
| 1. Đã quyết (NW → Khanh, Khanh 1/3 mẫu TNF) | 1,21 | 1,36 | 0,81 | 1,28 | 1,10 | 0,56 (0,80) | 0,64 |
| 2. + giao việc quản lý cho Luan, Cuong (10% mỗi người) | **1,35** | **1,52** | 0,81 | 1,28 | 1,10 | 0,56 (0,80) | 0,64 |
| 3. + 25% bulk TNF 富山企画 27SS của Luan → Khanh | 1,35 | 1,18 | 0,81 | 1,28 | 1,10 | 0,56 (0,80) | 0,98 |
| 4. + 25% bulk TNF 富山企画 27SS của Luan → Nhu | 1,35 | 1,18 | 0,81 | 1,28 | 1,10 | 0,86 (**1,23**) | 0,64 |
| 5. + 15% → Khanh, 15% → Nhu | 1,35 | 1,15 | 0,81 | 1,28 | 1,10 | 0,71 (1,01) | 0,84 |
| 6. Phương án 5 + Hien hỗ trợ 25% mẫu HELLY HANSEN của Cuong | **1,15** | **1,15** | 0,98 | 1,28 | 1,10 | 0,71 (1,01) | 0,84 |

Nhận xét:
1. **Giao thêm việc quản lý làm Luan lên 1,52 và Cuong lên 1,35**, vượt mức tăng ca 1 giờ/ngày. Phải chuyển bớt việc brand của cả hai, không chỉ Luan.
2. **Giảm cho Luan: chia phần theo dõi bulk TNF 富山企画** (NPL, BL/P/L, xuất hàng) của một số lô/hợp đồng, Luan vẫn giữ brand và làm đầu mối với TIJ. Chia khoảng 30% là đủ đưa Luan về ~1,15.
   - Nhận về **Khanh** (còn công suất, đã quen TNF). Không nên đưa hết cho Khanh vì Khanh vừa nhận NEUTRALWORKS và 1/3 mẫu TNF, đang lấy lại nhịp.
   - Nhận về **Nhu** (đang làm TNF別注 富山, cùng nhóm TNF 富山). Nhu nhận 25% thì chỉ số thực lên 1,23, quá nhiều khi Nhu còn việc văn phòng; 15% thì vừa (~1,0).
   - Vì vậy **phương án 5 (15% Khanh + 15% Nhu)** cân bằng nhất. Cách chia gọn: theo hợp đồng (vd. TN26-026 và TN26-026-2 tách riêng như trong bảng thực tế xuất hàng).
3. **Giảm cho Cuong: Hien hỗ trợ khoảng 25% mẫu HELLY HANSEN trong đợt mẫu triển lãm FW.** Brand của Hien vào mùa thấp điểm (0,81 kể cả hệ số độ khó), nên còn chỗ. Cuong về ~1,15.
4. **Hien chỉ rảnh trong đợt mẫu FW.** Mùa SS (1st 28SS rồi SMS 28SS) Hien lại cao nhất nhóm (1,27 ở 27SS), nên phần hỗ trợ HELLY HANSEN chỉ nên làm đến hết mẫu triển lãm 27FW, không giao lâu dài.
5. **Huy** (1,28) giữ nguyên như đã quyết, trong phạm vi tăng ca 1 giờ/ngày.

### Gợi ý cân bằng

1. **Đã quyết**: NEUTRALWORKS → Khanh; Khanh nhận 1/3 mẫu triển lãm TNF của B.Trang; Huy giữ Goldwin xuyên suốt; Nhu giữ nguyên brand.
2. **Trước khi giao việc quản lý cho Luan**: chia khoảng 30% phần theo dõi bulk TNF 富山企画 27SS (theo hợp đồng) cho Khanh và Nhu, mỗi người ~15%. Luan vẫn làm đầu mối.
3. **Trước khi giao việc quản lý cho Cuong**: Hien hỗ trợ ~25% mẫu HELLY HANSEN đến hết đợt mẫu triển lãm 27FW.
4. **Tăng ca 1 giờ/ngày là đủ** cho người ở mức 1,0–1,3. Người từ 1,4 trở lên phải chia việc.
5. **Việc giao cho Luan, Cuong nên thuộc nhóm kế hoạch** (họp SX tuần, plan table, điều chỉnh ngày xuất…). Đây đúng là bước "Order Follow-up MD → Planning Coordinator" (Level 1 → 2) trong sổ tay MD, nên coi là đào tạo, không chỉ là chia việc.
6. Làm lại bảng phân tích nghiệp vụ (業務分析表) kỳ sau theo số liệu thực tế. Lưu ý: các việc liên quan khai trị giá hải quan thay đổi người làm thì phải báo TI.

## Dữ liệu cần xác nhận

- **File SPL**: sheet "シーズンごと纏め" lệch nhiều so với sheet từng người (vd. Nhu 26SS: 346 cái/112 mã so với 146/63; Q.Huy 24FW: 414/110 so với 108/29). Phân tích dùng sheet từng người vì khớp với sheet "全体纏め" và có ghi chú cập nhật.
- **File BULK**: sheet "纏め" cũ hơn sheet từng người (thiếu YONEX của Cuong, FOXFIRE/TATRAS của B.Trang, Nhu 27SS 30.056 so với 32.056). Phân tích dùng sheet từng người.
- **Mẫu 27SS chưa có giây TB** cho nhiều brand, nên ngày may đang bằng số cái. Phân tích tạm lấy giây TB của brand mùa 26SS để ước (vd. HH 11.036 giây → 4 ngày/cái).
- **TNF別注 富山 (Nhu) và TNF別注 東京 (Q.Huy)**: ở các mùa 25FW, 26SS, số mẫu của hai bạn giống hệt nhau (vd. 146 cái/63 mã), có thể bị nhập trùng.
- **H.Trang bulk 26SS**: 384 cái, 1 mã nhưng 768 ngày chuyền (ghi chú "SPL Line"), số này chưa hợp lý. H.Trang thuộc nhóm CAD nên không tính vào phân tích MD.
- **Giờ dịch/mã** trong bảng là ước tính chung theo brand. Giả định mỗi mã ở mỗi đợt (bulk, 1st–3rd, SMS) dịch một bộ tài liệu.

## Theo dõi định kỳ (theo sổ tay MD)

Sổ tay MD muốn MD chuyển từ "theo đuôi đơn hàng" sang "kiểm soát sản xuất" với 5 chức năng: kế hoạch (PPC), NPL (MRC), năng lực (CC), rủi ro (PRC), chuẩn hóa liên lạc (CS). Để làm được, cần đo khối lượng thường xuyên:

1. **Mỗi mùa**: cập nhật 2 file BULK/SPL (sheet từng người), chạy lại `build.py`, xem chỉ số tải trước khi giao brand mùa mới.
2. **Mỗi tuần** (thứ Sáu – Weekly MD Report): số mã đang mở, NPL chưa có thông tin, rủi ro đang mở theo từng người.
3. **Ghi giờ theo 5 chức năng của sổ tay** trong 2–4 tuần (mỗi người tự ghi, mức 30 phút). Hiện các bạn dành phần lớn thời gian cho dịch và NPL thiếu, còn ít thời gian cho kế hoạch, năng lực, rủi ro. Số liệu này cho biết cần giảm việc gì để chuyển dần lên Level 2–3.

## Câu hỏi mở

- Việc quản lý giao cho Luan và Cuong gồm những việc nào, mỗi người khoảng bao nhiêu giờ/ngày? (đang tạm tính 10%, ~0,9 giờ/ngày)
- Hệ số độ khó 1,3 cho ALLTERRAIN/DESCENTE có hợp lý không? Còn brand nào khó tương tự (vd. HELLY HANSEN nhiều trang tài liệu)?
- Số mẫu 26FW thực tế của Hien theo brand (sheet Hien trong file SPL bị thiếu).
- Số mẫu triển lãm 27FW thực tế (khi có) để thay số ước.
- Phân tích nhóm kỹ thuật và CAD (gồm H.Trang) sẽ làm riêng sau khi xong nhóm MD.
