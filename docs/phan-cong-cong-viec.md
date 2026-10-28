# Dữ liệu phân công công việc TIVN (APIC, BTNB)

Tài liệu tóm tắt các file Excel về phân công nghiệp vụ của nhân viên TIVN (Toray International Vietnam) tại hai nhà máy **APIC** và **BTNB**, dùng làm dữ liệu gốc khi phân công công việc.

## Nguồn

| File gốc (`data/nguon/`) | Nội dung | Dữ liệu đã trích (`data/phan-cong/`) |
|---|---|---|
| `Phan-tich-nghiep-vu-TIVN-APIC-2026-thuong-ky.xlsx` (2026年度上期, cập nhật 1/4/2026) | Bảng phân tích nghiệp vụ APIC: 102 việc × 18 người, tỷ lệ thời gian từng người. Danh sách mã hàng cố vấn 天方 phụ trách mùa 26SS. | `apic-2026-thuong-ky.json`, `ma-hang-co-van-amakata-26ss.json` |
| `Hop-TIVN-phi-uy-thac-APIC-BTNB-20240130.xlsx` | Biên bản họp 15/9/2023, 31/10/2023, 27/11/2023, 30/1/2024 về phí ủy thác TIVN và khai trị giá hải quan (評価加算). Bảng phân tích nghiệp vụ APIC và BTNB năm 2023, 2024. Ngân sách phí ủy thác 2023, 2024. | `apic-2024.json`, `btnb-2024.json`, `apic-2023.json`, `btnb-2023.json` |
| `Phan-cong-nghiep-vu-APIC-BTNB-2018-10.xlsx` | Bảng phân công cũ (10/2018): 110 việc, đánh dấu ●/○ theo người, kèm số giờ mỗi lần, tần suất, input/output. | `phan-cong-2018-10.json` |
| `Phan-cong-brand-MD-APIC-20260722.xlsx` | Bảng phân công brand của nhóm MD APIC ngày 22/7/2026 (担当業務表). Các sheet ẩn: bảng brand 2017–2018 và lịch sử chuyển brand. | `brand-md-apic-2026-07-22.json` |

Tạo lại JSON từ Excel: `python3 data/phan-cong/build.py . data/phan-cong`

**Cách đọc số liệu:** mỗi người có tổng thời gian = 1,0 (100%). Ví dụ "Cuong 16%" nghĩa là Cuong dành khoảng 16% thời gian làm việc cho việc đó. Riêng file 2018 chỉ đánh dấu ● (phụ trách chính), ○ (phía Nhật hỗ trợ), (●) (phụ trách khi cần).

## Phân công brand nhóm MD APIC (22/7/2026)

(1) = brand do phòng 1 (1課) phía TI Nhật quản lý, (2) = phòng 2 (2課).

| Người | Brand / phạm vi |
|---|---|
| Hien | Teton Bros (2), Mont-Bell (2), ALLTERRAIN (2), DESCENTE GOLF (2) |
| Nhu | TNF 別注 SHIGEMOTO (từ 26SS), TNF 別注 NAKAO (từ 27SS), UTMF フィニッシャーズ (1), TOKYO MARATHON (2) |
| Luan | TNF 富山企画 (1), từ PPS đến sản xuất đại trà |
| Cuong | TNF 東京企画 (1) từ PPS đến đại trà, HELLY HANSEN (1), NEUTRALWORKS (1), YONEX (2) đến 26FW; **kiểm kế hoạch sản xuất mẫu của tất cả brand** |
| Khanh | TNF (ALL) mẫu 1st–SMS từ 24FW, Goldwin (1) từ PPS đến đại trà từ 24FW, TORAIN bền vững (2) SMS. *Nghỉ thai sản, đi làm lại từ giữa 10/2026.* |
| Huy (Q.Huy) | Goldwin (1), TNF 別注 HOSOYA (từ 26SS), 141 (hỗ trợ), CANTERBURY (1) từ 26FW, NANGA (2) |
| Huyen Trang (H.Trang) | TNF HAT&CAP, 141. *Thuộc nhóm CAD (là "Trang" trong bảng nghiệp vụ 2026), làm thêm 2 brand nhỏ này.* |
| Bich Trang (B.Trang) | TNF (ALL) mẫu 1st–SMS từ 27SS, TNF Premium Snow 別注 26FW (1), FOXFIRE (2), TATRAS (2), Kasco (2), YONEX (2) từ S27 |

Tên viết tắt trong file BULK/SPL: HH = HELLY HANSEN, NW = NEUTRALWORKS, TNF(GTC企画) = TNF 富山企画, TNF別注 富山 = 別注 SHIGEMOTO/NAKAO (Nhu), TNF別注 東京 = 別注 HOSOYA (Huy).

Trong thời gian Khanh nghỉ: mẫu TNF từ 27SS do Bich Trang làm, Goldwin đại trà từ 26FW do Huy làm.

Quyết định 10/2026 (sau khi Khanh đi làm lại): Hien, Huy, Nhu giữ nguyên. Cuong chuyển NEUTRALWORKS sang Khanh (ngay) và TNF 東京企画 sang Khanh (từ 11/2026), giữ HELLY HANSEN (chủ lực), nhận thêm điều phối chung (theo dõi xuất hàng cả nhóm, quản lý chung may mẫu). Khanh nhận NEUTRALWORKS, 1/3 mẫu TNF của Bich Trang từ mẫu 1st 28SS (mẫu triển lãm 27FW đã vào may, người đang làm làm tiếp) và một phần việc của Luan (tách theo chức năng, Luan giữ TNF 富山企画). Luan nhận một phần việc quản lý. Phân chia chi tiết Luan – Khanh: xem `docs/phan-tich-khoi-luong.md`.

Lịch sử (sheet ẩn, 2017–2018): trước đây có Phong, Ngan, Thuan, Nhung làm merchandiser, phụ trách các brand như MARMOT, MILLET, SNOW PEAK, UMBRO, WOOLRICH, SCHOFFEL. Nhóm kỹ thuật khi đó: Cau dịch/phiên dịch tiếng Nhật, Boken, đặt xe; Quyen quản lý xưởng 1-A/B; Vy, Hong nghiệp vụ mẫu; Phung (email trucphuong, có lẽ là Phuong hiện nay) xưởng 1-A, hỗ trợ mẫu.

## Nhân sự APIC – 2026 上期 (dữ liệu mới nhất)

18 vị trí: 1 quản lý người Nhật, 11 nhóm quản lý, 5 nhóm kỹ thuật, 1 bộ phận TIVN trụ sở (đặt xe, tiếp khách).

| Người | Nhóm | Số việc | Việc chiếm nhiều thời gian nhất |
|---|---|---|---|
| Hori | Quản lý người Nhật | 35 | Họp SX tuần 15%, theo dõi tiến độ 10%, điều chỉnh KHSX 7%, sửa plan table / Update / báo cáo tuần |
| Chanh | Quản lý (trưởng nhóm) | 41 | Họp SX tuần 12%, theo dõi tiến độ 10%, plan table, KH trung hạn, điều chỉnh ngày xuất, kiểm thứ 3, các báo cáo tháng |
| Cuong | Quản lý – merchandiser | 37 | NPL thiếu (tổng ~31%), dịch tài liệu 15%, nhận tài liệu mẫu 14% |
| Hien | Quản lý – merchandiser | 37 | Giống Cuong |
| Huy | Quản lý – merchandiser | 40 | Giống Cuong, thêm họp cải tiến chất lượng, báo nhà máy điểm cần lưu ý |
| Khanh | Quản lý – merchandiser | 37 | NPL thiếu (~35%), dịch 15%, tài liệu mẫu 14% |
| Luan | Quản lý – merchandiser | 37 | NPL thiếu, dịch, tài liệu mẫu (15% mỗi việc), báo cáo mã ngoài quy tắc 2 tuần 10%, tiến độ NPL 5% |
| Nhu | Quản lý – hành chính, xuất hàng | 44 | Kiểm đột xuất danh sách công nhân 10%, NPL thiếu, dịch, nhận NPL từ HCM 6%, xuất hàng + P/L 10%, Debit Note, AUDIT |
| Bich Trang | Quản lý – merchandiser (mới) | 37 | NPL thiếu (~35%), dịch 15%, tài liệu mẫu 14% |
| Cau | Kỹ thuật/QC (cột nhóm quản lý) | 31 | Đi xưởng 30%, họp kỹ thuật 15%, cải tiến chất lượng, kiểm của khách, sự cố; phát triển SP, gửi mẫu |
| Trang | CAD | 9 | Ra rập, nhảy size, tài liệu may (25% mỗi việc), hướng dẫn may, xác nhận tài liệu |
| Ngan | CAD | 9 | Giống Trang |
| Vy | Kỹ thuật – mẫu | 13 | Comment mẫu + bảng đo 22%, sửa rập 20%, kiểm rập 10%, hỏi đáp tổ mẫu 10% |
| Quyen | Kỹ thuật | 22 | Sự cố khác màu/co rút 18%, cảnh báo & đề xuất 9%, hỏi đáp tổ mẫu 8%, các bước kiểm hàng |
| Hong | Kỹ thuật | 24 | Sự cố 18%, đi xưởng 15%, xác nhận tài liệu, các bước kiểm hàng |
| Phuong | Kỹ thuật | 16 | Xử lý sự cố 18%, hướng dẫn may 10%, xác nhận tài liệu 10%, kiểm may 9% |
| Luu | Kỹ thuật | 16 | Xử lý sự cố 17%, kiểm may 11%, họp trước khi vào chuyền 10%, hướng dẫn may 8% |
| TIVN trụ sở | Hành chính | 2 | Đặt xe/khách sạn 50%, tiếp khách 50% |

Ghi chú:
- Nhóm merchandiser (Cuong, Hien, Huy, Khanh, Luan, Nhu, Bich Trang) làm chung một bộ việc theo mã hàng: nhận order mẫu, dịch tài liệu, kiểm rập/sơ đồ, theo dõi và báo NPL thiếu, thêu/in/chần, kiểm nhãn. Phân việc giữa họ nên chia theo **mã hàng / khách hàng**.
- Chanh và Hori lo kế hoạch: plan table, KH trung hạn, họp SX tuần, điều chỉnh ngày xuất. Các báo cáo tháng (tồn kho, KH–thực tế, thiết bị) do Chanh làm.
- CAD (Trang, Ngan) là nhóm mới từ 2026, làm rập, nhảy size, tài liệu may.
- So với 2024: quản lý người Nhật đổi từ Inoue sang Hori; thêm Bich Trang, Trang, Ngan.

## Nhân sự BTNB – 2024 (dữ liệu BTNB mới nhất)

14 vị trí: quản lý người Nhật Naito, nhóm quản lý Thong (trưởng nhóm), Tuong, Nhi, T_Van, Tuan, Trinh; nhóm kỹ thuật H_Van, Chau, Minh, Thu, Hoa, 1 NEW STAFF; TIVN trụ sở.

| Người | Việc chính |
|---|---|
| Naito | Điều chỉnh KHSX 40%, điều chỉnh ngày xuất 10%, họp SX, báo cáo KH/thực tế, KH trung hạn |
| Thong | 61 việc, mỗi việc nhỏ: tài liệu mẫu 7%, báo cáo tháng, sửa KH, họp SX, đi xưởng, chấm công, kế toán |
| Tuong, Nhi, Tuan | Nhận và dịch tài liệu mẫu (~25%), dịch 5–6%, lỗi vải, NPL thiếu |
| T_Van, Trinh | Tài liệu mẫu (~20%), NPL thiếu ~8%, dịch 5%; Trinh thêm họp cải tiến chất lượng, kế toán |
| H_Van, Chau, Minh, Thu, Hoa | Hướng dẫn may (17–20%), comment mẫu (10–17%), cảnh báo & đề xuất (7–12%), kiểm theo rập/tài liệu |
| NEW STAFF | Xác nhận tài liệu 10%, sự cố, kiểm rập, kỹ thuật quản lý |

## Việc tính vào trị giá hải quan (評価加算)

Từ 12/2023, một phần phí ủy thác TIVN phải cộng vào trị giá tính thuế khi nhập hàng APIC/BTNB về Nhật (đã được Hải quan Kobe chấp nhận qua 包括評価申告). Thay đổi nội dung công việc của các mục này ảnh hưởng tới số tiền khai thuế, nên **khi phân công lại phải báo TI** (việc ③ của TIVN trong biên bản 27/11/2023).

| Năm | Các việc tính thuế (đánh dấu **[thuế]** trong bảng dưới) | Số tiền cộng thêm |
|---|---|---|
| 2023 | 6/100 việc: 商品開発, 加工技術, 生産技術, 評価技術, サンプル発送, 資材不足確認 (giai đoạn SX) | APIC $22.915/năm ÷ 500.000 cái ≈ **$0,05/cái**; BTNB $6.244/năm ÷ 450.000 cái ≈ **$0,02/cái** |
| 2024 (5/2024–4/2025) | Như 2023 | APIC $15.929/năm ÷ 600.000 cái ≈ **$0,03/cái**; BTNB $8.549/năm ÷ 500.000 cái ≈ **$0,02/cái** |
| 2026 上期 (APIC) | 7 việc: 商品開発, 加工技術, 生産技術, 評価技術, パターンメイク, グレーディング, 縫製仕様書. Gửi mẫu và NPL thiếu không còn tính. | Chưa có trong file |

Các điểm khác trong biên bản họp:
- Chứng từ (evidence) kèm invoice ghi "Purpose for customs valuation only", gộp tên "TIVN 業務委託費（開発費用・サンプル発送費用・資材不足確認費用）", không phải thanh toán.
- Cách tính: lấy ngân sách phí ủy thác → chia theo tỷ lệ thời gian từng người → cộng tiền của các việc tính thuế → chia cho số cái sản xuất một năm.
- Phí thuê máy may BTNB $35.000 do BTNB chịu nên không phải cộng.
- Phí ủy thác 3 người Nhật ở nhà máy QN chưa có danh sách công việc, cần làm.
- Tiền điện APIC (2020–8/2023 khoảng $226.000) dự định đàm phán để nhà máy chịu.

## Ngân sách phí ủy thác 2024 (nghìn USD)

| | APIC | BTNB |
|---|---|---|
| Nhân viên Việt Nam | 13 người, 295,83 → $22,76/người | 12 người, 335,87 → $30,53/người |
| Người Nhật (từ TIVN HCM, mỗi bên 1 người) | 180,75 | 180,75 |
| TIVN trụ sở | 91,5 | 91,5 |
| Tổng phân bổ | 568,08 (≈ 47,3/tháng) | 608,12 (≈ 50,7/tháng) |

## Mã hàng cố vấn 天方 phụ trách – 26SS

TIVN làm rập, nhảy size, tài liệu may cho các mã này. Tất cả đã xuất đủ số kế hoạch (cập nhật 1/4/2026).

| TIVN phụ trách | Brand | Mã hàng | Số lượng | ETD | Hợp đồng |
|---|---|---|---|---|---|
| Nhu | TNF | NB42660R | 460 | 29/1/2026 | S26別注 (富山設計) |
| Nhu | TNF | NP12661R | 507 | 10/2/2026 | S26別注 (富山設計) |
| Nhu | TNF | NB32662R | 686 | 4/2/2026 | S26別注 (富山設計) |
| Khanh | GOLDWIN | GM76151 | 1.436 | 5–11/3/2026 | GWS26-001 |
| Luan | TNF | NP72583 | 1.226 | 20/11/2025 | TN26-010, -1 |
| Luan | TNF | NP72581 | 689 | 21/11/2025 | TN26-010, -1 |
| Luan | TNF | NP72580 | 1.391 | 3/12/2025 | TN26-010, -1 |
| Luan | TNF | NPW72580 | 1.098 | 20/12/2025 | TN26-010-1 |
| Luan | TNF | NPW22532 | 10.540 | 14/1/2026 | TN26-026, -2 |
| Luan | TNF | NBW32531 | 2.940 | 14/1/2026 | TN26-026-1 |
| Luan | TNF | NP22532 | 13.600 | 29/1/2026 | TN26-026, -2 |
| Luan | TNF | NB32531 | 4.840 | 29/1/2026 | TN26-026, -2 |

Sheet ẩn cũ (25SS/25FW) liệt kê thêm GL25142, GM75159, HH12521, NB32531, NP22530, NP22532, NPW22532, NBW32531 (25SS, tổng 41.419 cái) và NP72580/81/83, NPW72580, NP62502, NPW62502, NS62515, GL25142, GLW25142 (25FW).

## Bảng phân công 10/2018 (tham khảo)

Tổ chức khi đó: văn phòng HCM (山口所長, 植井部長, HonNgoc, Huyen), phòng QLSX Nhật (石坂, 西山, 櫛田, 徐, 陸, Nga, 大谷), bộ phận kỹ thuật gia công (木戸, 中家).
- **APIC:** người Nhật thường trú 井上, 小山; TIVN Chanh, Phong, Hien, Cuong, Nhu, Luan, Nhung, Thuan (quản lý), Cau, Hong, Quyen, Phuong, Vy (kỹ thuật).
- **BTNB:** người Nhật thường trú 稲葉, 井上, 鈴木; TIVN Binh, Duc, Tuyen, Khang (quản lý), Thong, Thu, Phung, Minh, Van (kỹ thuật).

File 2018 có thêm số giờ mỗi lần làm (vd. họp SX tuần 5 giờ, theo dõi NPL 3 giờ/ngày, comment mẫu 6 giờ/ngày) và input/output của từng việc. Danh mục việc gần giống bảng hiện tại, nên dùng được để ước lượng khối lượng khi giao việc.

## Danh mục công việc và người làm

Cột APIC 2026 và BTNB 2024: người làm và % thời gian của họ. **[thuế]** = việc tính vào trị giá hải quan năm 2026 (APIC). n/a = BTNB 2024 không có việc này.

| # | Công đoạn | Công việc / 作業 | Tần suất | APIC 2026 上期 | BTNB 2024 |
|---|---|---|---|---|---|
| 1 | 総括 / 契約管理・交渉 | Debit Note – Kiểm tra Debit Note, đàm phán cập nhật theo thực tế | 月次 | Chanh 1%, Nhu 1% | Thong 2%, Naito 1% |
| 2 | 総括 / 契約管理・交渉 | 出勤管理 – Tổng hợp chấm công hàng tháng | 月次 | — | Thong 2%, Naito 1% |
| 3 | 総括 / 契約管理・交渉 | 期中変更対応 – Xử lý yêu cầu thay đổi nghiệp vụ giữa kỳ từ nhà máy | 不定期（頻度低） | Hori 1%, Chanh 1% | Thong 2%, Naito 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 4 | 総括 / 契約管理・交渉 | 生産計画調整 – Điều chỉnh, chốt kế hoạch sản xuất với nhà máy | 週次 | Hori 7%, Chanh 1% | Naito 40%, Thong 2% |
| 5 | 総括 / 契約管理・交渉 | 残業要請 – Yêu cầu tăng ca | 不定期（頻度低） | Hori 2%, Chanh 1% | n/a |
| 6 | 総括 / 契約管理・交渉 | Air代負担交渉 – Đàm phán chia phí hàng Air | 不定期（頻度低） | Hori 2%, Chanh 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1% |
| 7 | 総括 / 契約管理・交渉 | 外注工賃交渉 – Đàm phán đơn giá gia công ngoài | 不定期（頻度低） | Hori 2%, Chanh 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1% |
| 8 | 総括 / AUDIT | 事前準備 – Chuẩn bị AUDIT: hồ sơ, hướng dẫn nhà máy | 不定期（頻度低） | Chanh 1%, Nhu 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 9 | 総括 / AUDIT | 当日対応 – Tiếp đoàn AUDIT trong ngày | 不定期（頻度低） | Hori 1%, Chanh 1%, Nhu 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 10 | 総括 / 来客・出張対応 | 移動 – Đặt xe, khách sạn cho khách/công tác | 不定期（頻度高） | TIVN本社 50%, Hori 2%, Chanh 1%, Cau 1% | TIVN本社 50%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 11 | 総括 / 来客・出張対応 | 工場アテンド – Đưa khách đi nhà máy, ăn uống | 不定期（頻度低） | TIVN本社 50%, Hori 5%, Quyen 5%, Phuong 3%, Hong 2%, Luu 2%, Chanh 1% | TIVN本社 50%, Thong 2%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 12 | 総括 / 来客・出張対応 | 安全管理 – Nắm lịch trình, xử lý khẩn cấp | 不定期（頻度高） | Hori 5%, Chanh 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 13 | 企画 / 開発 | 商品開発 – Phát triển SP: vật liệu mới, kỹ thuật gia công, may thử **[thuế]** | 不定期（頻度中） | Cau 2% | NEW STAFF 1% |
| 14 | 企画 / 開発 | 加工技術 – Kỹ thuật gia công (dán/ép seam, máy rập tự động, máy nhồi lông) **[thuế]** | 不定期（頻度中） | Cau 1% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, NEW STAFF 1% |
| 15 | 企画 / 開発 | 生産技術 – Kỹ thuật sản xuất: IE, thiết kế chuyền **[thuế]** | 不定期（頻度低） | Cau 1% | Thong 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, NEW STAFF 1% |
| 16 | 企画 / 開発 | 管理技術 – Kỹ thuật quản lý: phân tích công đoạn, QC | 不定期（頻度中） | Cau 2% | NEW STAFF 7%, Thong 1% |
| 17 | 企画 / 開発 | 評価技術 – Đánh giá chức năng, sản phẩm, thử cơ lý **[thuế]** | 不定期（頻度中） | Cau 1% | Thong 1%, NEW STAFF 1% |
| 18 | 前段取り / 生産計画 | プランテーブル – Plan table – kế hoạch dài hạn | 週次 | Chanh 5%, Hori 2% | Naito 5%, Thong 2% |
| 19 | 前段取り / 生産計画 | 明細連絡 – Gửi chi tiết SKU cho nhà máy | 日次 | Chanh 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 20 | 前段取り / 生産計画 | 工場生産計画 – Kế hoạch SX trung hạn của nhà máy | 週次 | Chanh 4%, Hori 2% | Naito 5%, Thong 1%, Tuong 1% |
| 21 | 前段取り / 生産計画 | 資材進捗 – Theo dõi tiến độ NPL, báo nhà máy | 日次 | Chanh 5%, Luan 5%, Cuong 2%, Hien 2%, Huy 2%, Khanh 2%, Bich Trang 2%, Nhu 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 22 | 前段取り / 生産計画 | 注意連絡 – Báo nhà máy các điểm cần lưu ý đặc biệt | 週次 | Hori 2%, Chanh 1%, Huy 1%, Cau 1%, Quyen 1%, Hong 1% | Chau 4%, H_Van 2%, Minh 2%, Thu 2%, Hoa 2%, Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 23 | 前段取り / サンプル | 受注 – Nhận order mẫu, ghi sổ quản lý mẫu | 日次 | Vy 4%, Cuong 2%, Hien 2%, Huy 2%, Khanh 2%, Bich Trang 2%, Luan 1%, Nhu 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 24 | 前段取り / サンプル | 工場指示 – Chuyển yêu cầu mẫu từ Nhật cho nhà máy | 日次 | Hori 2%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 25 | 前段取り / サンプル | 納期回答 – Xếp lịch, trả lời ngày giao mẫu | 日次 | Vy 4%, Hori 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 26 | 前段取り / サンプル | 仕様 – Nhận và dịch tài liệu kỹ thuật (mẫu) | 日次 | Luan 15%, Cuong 14%, Hien 14%, Huy 14%, Khanh 14%, Bich Trang 14%, Nhu 5% | Tuong 26%, Nhi 25%, Tuan 25%, T_Van 20%, Trinh 19%, Thong 7% |
| 27 | 前段取り / サンプル | パターンチェック – Nhận rập, kiểm thông số | 日次 | Vy 10%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | NEW STAFF 3%, Thu 2%, H_Van 1%, Chau 1%, Minh 1%, Hoa 1% |
| 28 | 前段取り / サンプル | 用尺計算 – Kiểm sơ đồ nhà máy làm, yêu cầu sửa | 日次 | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thu 2%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Hoa 1%, NEW STAFF 1% |
| 29 | 前段取り / サンプル | 資材不足確認 – Kiểm NPL thiếu, báo TI | 日次 | Cuong 5%, Hien 5%, Huy 5%, Khanh 5%, Bich Trang 5%, Luan 1%, Nhu 1% | T_Van 8%, Trinh 8%, Tuan 4%, Thong 1%, Tuong 1%, Nhi 1% |
| 30 | 前段取り / サンプル | 進捗フォロー – Theo dõi tiến độ chuyền mẫu | 週3回程度 | Vy 5%, Hong 1% | Nhi 3%, Thong 2%, T_Van 2%, Tuan 2%, Trinh 2%, Tuong 1%, NEW STAFF 1% |
| 31 | 前段取り / サンプル | 諸確認 – Trả lời câu hỏi của tổ mẫu | 週3回程度 | Vy 10%, Quyen 8%, Hong 2%, Hori 1% | NEW STAFF 3%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 32 | 前段取り / サンプル | 問題抽出・改善提案 – Cảnh báo lỗi tài liệu/vật liệu, đề xuất cải tiến | 不定期（頻度中） | Quyen 9%, Vy 3%, Hong 2% | Chau 12%, H_Van 10%, Hoa 10%, Minh 8%, Thu 7%, NEW STAFF 4%, Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 33 | 前段取り / サンプル | コメント連絡 – Ghi, gửi comment mẫu và bảng đo | 日次 | Vy 22% | H_Van 17%, Thu 15%, Minh 14%, Chau 10%, Hoa 10%, NEW STAFF 3%, Tuong 2%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 34 | 前段取り / サンプル | 二次加工 – Kiểm vấn đề bulk của mẫu làm ở xưởng ngoài | 不定期（頻度低） | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 35 | 前段取り / サンプル | 出荷連絡 – Báo TI chi tiết xuất, số vận đơn | 日次 | Cau 2% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 36 | 前段取り / サンプル | 発送 – Gửi mẫu | 日次 | Cau 2% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 37 | 前段取り / サンプル | 生産条件設定 – Thiết lập điều kiện ép/seam, kiểm chức năng | 日次 | Vy 5%, Quyen 1%, Hong 1% | Thong 1%, Tuong 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 38 | 前段取り / サンプル | 工程分析 – Phân tích công đoạn, tính năng lực & chi phí | 不定期（頻度低） | Hori 1%, Chanh 1% | Thong 2%, NEW STAFF 2%, Tuong 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 39 | 前段取り / サンプル | 元見本管理 – Quản lý sổ mẫu gốc (nhập/xuất) | 不定期（頻度中） | Vy 2%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Cau 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 40 | 前段取り / 仕様書 | ファイル作成 – In, lưu file tài liệu theo brand/mùa | 不定期（頻度中） | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 41 | 前段取り / 仕様書 | 翻訳 – Dịch tài liệu sang tiếng Việt | 日次 | Cuong 15%, Hien 15%, Huy 15%, Khanh 15%, Luan 15%, Bich Trang 15%, Nhu 10% | Nhi 6%, Tuong 5%, T_Van 5%, Tuan 5%, Trinh 5%, Thong 3% |
| 42 | 前段取り / 仕様書 | チェック – Kiểm, báo TI khi lệch rập/mẫu gốc/tài liệu | 日次 | Nhu 5%, Luan 4%, Cuong 3%, Hien 3%, Huy 3%, Khanh 2%, Bich Trang 2% | NEW STAFF 8%, Minh 7%, H_Van 5%, Chau 5%, Thu 5%, Hoa 5%, Nhi 3%, Tuong 2%, T_Van 2%, Tuan 2%, Trinh 2% |
| 43 | 前段取り / 仕様書 | 修正連絡 – Dịch tài liệu sửa đổi, báo nhà máy | 不定期（頻度高） | Luan 4%, Nhu 4%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Bich Trang 1% | T_Van 2%, Tuan 2%, Trinh 2%, Tuong 1%, Nhi 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 44 | 前段取り / パターン | パターン確認と連絡 – Chuyển dữ liệu rập cho nhà máy, báo TI khi có vấn đề | 不定期（頻度高） | Cuong 2%, Hien 2%, Huy 2%, Khanh 2%, Luan 2%, Nhu 2%, Bich Trang 2% | n/a |
| 45 | 前段取り / パターン | パターン修正 – Chỉ đạo nhà máy sửa rập | 不定期（頻度中） | Vy 20%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Quyen 1%, Hong 1% | n/a |
| 46 | 前段取り / CAD | パターンメイク – Ra rập, đào tạo **[thuế]** | 日次 | Trang 25%, Ngan 25% | n/a |
| 47 | 前段取り / CAD | グレーデイング – Nhảy size, đào tạo **[thuế]** | 日次 | Trang 25%, Ngan 25% | n/a |
| 48 | 前段取り / CAD | 縫製仕様書 – Làm tài liệu kỹ thuật may, đào tạo **[thuế]** | 日次 | Trang 25%, Ngan 25% | n/a |
| 49 | 前段取り / 資材管理 | 資材着荷確認 – Báo nhà máy BL, P/L, số DHL | 日次 | Cuong 2%, Hien 2%, Huy 2%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 50 | 前段取り / 資材管理 | 資材不足確認 – Kiểm NPL thiếu, báo TI | 日次 | Khanh 20%, Bich Trang 20%, Cuong 16%, Hien 16%, Huy 13%, Nhu 10%, Luan 1%, Cau 1% | Tuong 3%, Nhi 2%, T_Van 2%, Tuan 2%, Trinh 2%, Thong 1% |
| 51 | 前段取り / 資材管理 | 生地不良確認 – Nhận báo cáo kiểm vải, xác nhận lỗi với TI | 日次 | Cuong 2%, Hien 2%, Huy 2%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Cau 1% | NEW STAFF 5%, Tuong 3%, Nhi 3%, Thong 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 52 | 本生産 / 裁断 | 明細確認 – Kiểm chi tiết cắt nhà máy lập | 日次 | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 53 | 本生産 / 裁断 | 要尺確認 – Kiểm định mức mẫu, báo TI | 日次 | Cuong 2%, Hien 2%, Huy 2%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 54 | 本生産 / 裁断 | トラブル対応 – Xử lý sự cố khác màu, co rút… | 不定期（頻度高） | Quyen 18%, Hong 18%, Phuong 6%, Luu 6%, Cau 5%, Hori 2%, Chanh 2%, Luan 2%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Nhu 1%, Bich Trang 1% | NEW STAFF 8%, Nhi 2%, T_Van 2%, Tuan 2%, Trinh 2%, H_Van 2%, Chau 2%, Minh 2%, Thu 2%, Hoa 2%, Thong 1%, Tuong 1% |
| 55 | 本生産 / 二次加工 | 刺繍 – Nhận dữ liệu thêu, liên hệ xưởng thêu | 不定期（頻度中） | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 56 | 本生産 / 二次加工 | プリント – Nhận dữ liệu in, liên hệ xưởng in | 不定期（頻度中） | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 57 | 本生産 / 二次加工 | キルト – Đánh giá khả năng triển khai bulk chần bông | 不定期（頻度低） | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 58 | 本生産 / 二次加工 | 品位確認・評価 – Kiểm hàng thực tế, test riêng của khách | 不定期（頻度低） | Chanh 4%, Cau 2%, Hori 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 59 | 本生産 / 生産 | 仕様確認 – Xác nhận tài liệu hàng đại trà với TI | 日次 | Trang 10%, Ngan 10%, Phuong 10%, Quyen 7%, Hong 7%, Luu 7%, Cau 2% | NEW STAFF 10%, T_Van 2%, Tuan 2%, Trinh 2%, Thong 1%, Tuong 1%, Nhi 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 60 | 本生産 / 生産 | 資材不足確認 – Kiểm NPL thiếu, báo TI | 日次 | Luan 15%, Cuong 10%, Hien 10%, Huy 10%, Khanh 10%, Bich Trang 10%, Nhu 5%, Cau 2% | T_Van 2%, Tuan 2%, Trinh 2%, Thong 1%, Tuong 1%, Nhi 1% |
| 61 | 本生産 / 生産 | 縫製指導 – Hướng dẫn may tại xưởng | 日次 | Trang 10%, Ngan 10%, Phuong 10%, Luu 8%, Quyen 5%, Hong 5%, Cau 2% | Chau 20%, H_Van 19%, Minh 19%, Thu 19%, Hoa 17%, NEW STAFF 1% |
| 62 | 本生産 / 生産 | シーム・溶着指導 – Hướng dẫn seam/ép tại xưởng | 日次 | Quyen 3%, Hong 3% | NEW STAFF 4%, Thong 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 63 | 本生産 / 生産 | 納期管理 – Theo dõi báo cáo SX ngày, tiến độ | 日次 | Hori 10%, Chanh 10% | Naito 3%, Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 64 | 本生産 / 生産 | トラブルシュート – Xử lý sự cố (chuẩn bị, KH, công đoạn, chất lượng) | 不定期（頻度高） | Phuong 18%, Luu 17%, Quyen 5%, Hong 5%, Hori 2%, Chanh 2% | NEW STAFF 5%, Chau 4%, Minh 4%, Thu 4%, Hoa 4%, H_Van 3%, Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 65 | 本生産 / 生産 | 生産結果記録 – Ghi nhận vấn đề–đối sách, áp dụng cho mã tương tự | 不定期（頻度中） | Luu 6%, Phuong 5%, Quyen 3%, Hong 3%, Cau 2%, Hori 1%, Chanh 1% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 66 | 本生産 / 生産 | 技術者MTG – Dự họp trước khi vào chuyền | 日次 | Luu 10%, Phuong 5%, Quyen 3%, Hong 3% | Thong 3%, NEW STAFF 3%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 67 | 本生産 / 検査 | 検査目合わせ – Thống nhất tiêu chuẩn với SP đầu chuyền | 不定期（頻度高） | Phuong 7%, Quyen 5%, Hong 5%, Luu 5%, Cau 2% | NEW STAFF 5%, H_Van 3%, Chau 3%, Minh 3%, Thu 3%, Hoa 3% |
| 68 | 本生産 / 検査 | 一次検査 – Kiểm may hàng ra chuyền | 日次 | Luu 11%, Phuong 9%, Quyen 5%, Hong 5%, Cau 2% | Hoa 3%, Thu 2%, H_Van 1%, Chau 1%, Minh 1%, NEW STAFF 1% |
| 69 | 本生産 / 検査 | 二次検査 – Kiểm ngoại quan hàng hoàn thiện | 日次 | Quyen 5%, Hong 5%, Phuong 5%, Luu 5%, Cau 2% | Hoa 3%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, NEW STAFF 1% |
| 70 | 本生産 / 検査 | 出荷前検査 – Kiểm xác suất sau QC nhà máy | 日次 | Quyen 5%, Hong 5%, Phuong 3%, Luu 3%, Cau 2% | Minh 2%, Hoa 2%, H_Van 1%, Chau 1%, Thu 1%, NEW STAFF 1% |
| 71 | 本生産 / 検査 | 下札、ﾈｰﾑ確認 – Kiểm thẻ bài, nhãn | 日次 | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 72 | 本生産 / 検査 | 客先検査 – Báo kết quả kiểm của khách cho nhà máy | 不定期（頻度高） | Cau 5%, Phuong 4%, Luu 4%, Quyen 3%, Hong 3% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 73 | 本生産 / 検査 | 機能検査 – Kiểm xác suất chống thấm, seam | 不定期（頻度中） | Phuong 4%, Luu 4%, Quyen 3%, Hong 3%, Cau 1% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 74 | 本生産 / 検査 | 第三者検品 – Làm việc với bên kiểm thứ 3 (tiêu chuẩn, lịch) | 不定期（頻度中） | Chanh 5%, Hori 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Minh 2%, Thu 2%, H_Van 1%, Chau 1%, Hoa 1% |
| 75 | 出荷 / 出荷 | P/L管理 – Kiểm P/L sau xuất, xác nhận bù nếu hụt | 日次 | Nhu 5% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 76 | 出荷 / 出荷 | 出荷連絡 – Báo TI chi tiết xuất, số vận đơn | 日次 | Nhu 5% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 77 | 出荷 / 出荷 | 日程調整 – Điều chỉnh ngày/số lượng xuất với nhà máy | 週次 | Chanh 5%, Hori 1% | Naito 10%, Tuong 2%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 78 | 工場管理 / 備品管理 | 消耗品管理 – Đặt vật tư tiêu hao, lập phiếu giao | 不定期（頻度低） | — | Thong 2% |
| 79 | 工場管理 / 労務管理 | 出勤簿管理 – Tổng hợp chấm công tháng | 月次 | — | Thong 3%, Naito 1% |
| 80 | その他 / 日次作業 | 業務日報 – Báo cáo ngày cho TI Tokyo | 日次 | Chanh 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Cau 1%, Trang 1%, Ngan 1% | H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 81 | その他 / 日次作業 | 進捗報告会 – Họp tiến độ SX & chuyên cần với quản đốc | 日次 | Phuong 6%, Chanh 5%, Vy 5%, Luu 5%, Hori 2%, Quyen 2%, Hong 2%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Cau 1%, Trang 1%, Ngan 1% | Thong 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 82 | その他 / 日次作業 | 技術MTG – Họp kỹ thuật/QC buổi sáng | 日次 | Cau 15%, Luu 6%, Vy 5%, Phuong 4%, Trang 2%, Ngan 2%, Quyen 2%, Hong 2%, Hori 1% | Thong 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 83 | その他 / 日次作業 | 工場ラウンド – Đi xưởng kiểm nội dung báo cáo | 日次 | Cau 30%, Hong 15%, Hori 1%, Chanh 1% | Hoa 6%, Thong 3%, Naito 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1% |
| 84 | その他 / 日次作業 | HCMより材料引き取り – Xác nhận cách nhận NPL từ HCM | 日次 | Nhu 6% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 85 | その他 / 週次作業 | 生産MTG – Họp KHSX tuần tới với nhà máy | 週次 | Hori 15%, Chanh 12%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Naito 7%, Thong 3%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 86 | その他 / 週次作業 | サンプルMTG – Họp KH tổ mẫu | 週次 | Vy 5%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Chau 2%, Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Minh 1%, Thu 1%, Hoa 1% |
| 87 | その他 / 週次作業 | 業務週報 – Báo cáo tuần cho TI Tokyo |  | Hori 5%, Chanh 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1%, Cau 1%, Trang 1%, Ngan 1% | Naito 5% |
| 88 | その他 / 週次作業 | 計画修正連絡 – Báo phòng QLSX phương án sửa KH, chuyền trống | 週次 | Hori 5%, Chanh 5% | Thong 3%, Naito 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 89 | その他 / 週次作業 | プランテーブル修正 – Gửi plan table đã sửa | 週次 | Hori 5%, Chanh 5% | Thong 2%, Naito 1% |
| 90 | その他 / 週次作業 | Update修正内容確認 – Đối chiếu Update với plan table | 週次 | Hori 5%, Chanh 4% | Naito 5%, Thong 3%, Tuong 2%, Nhi 2%, T_Van 2%, Tuan 2%, Trinh 2% |
| 91 | その他 / 週次作業 | 出勤名簿の抜き打ち確認 – Kiểm đột xuất danh sách công nhân với chuyền | 週次 | Nhu 10% | Thong 1%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 92 | その他 / 月次作業 | 業務月報 – Báo cáo tháng | 月次 | Hori 2%, Chanh 1% | Thong 3%, Naito 2%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| 93 | その他 / 月次作業 | 材料在庫管理 – Báo cáo tồn kho NPL | 月次 | Chanh 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 94 | その他 / 月次作業 | 製品在庫管理 – Báo cáo tồn thành phẩm | 月次 | Chanh 1% | Naito 3%, Thong 2% |
| 95 | その他 / 月次作業 | 生産計画・実績報告 – Báo cáo KH/thực tế SX | 月次 | Chanh 1% | Naito 5%, Thong 2% |
| 96 | その他 / 月次作業 | 機材管理報告 – Báo cáo thiết bị cho nhà máy mượn | 月次 | Chanh 1% | Thong 1% |
| 97 | その他 / 月次作業 | 計画・実績情報報告 – Báo cáo KH–thực tế theo mã, so với bảng tra nhanh | 月次 | Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Luan 1%, Nhu 1%, Bich Trang 1% | Naito 2% |
| 98 | その他 / 月次作業 | 2週間ルール外品番報告 – Báo cáo mã ngoài quy tắc 2 tuần | 月次 | Luan 10%, Chanh 1%, Cuong 1%, Hien 1%, Huy 1%, Khanh 1%, Nhu 1%, Bich Trang 1% | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 99 | その他 / 月次作業 | 品質改善報告会議 – Họp/báo cáo cải tiến chất lượng | 月次 | Cau 5%, Huy 1% | Thong 2%, Trinh 1% |
| 100 | その他 / 月次作業 | 品質改善報告会議 – Họp/báo cáo cải tiến chất lượng | 月次 | Cau 2%, Hori 1%, Chanh 1%, Huy 1%, Quyen 1%, Hong 1%, Phuong 1%, Luu 1% | Thong 2%, Trinh 1%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1%, NEW STAFF 1% |
| 101 | その他 / 月次作業 | 出勤名簿・D/N照合 – Đối chiếu số công nhân với D/N | 月次 | Hori 1%, Chanh 1% | Thong 2%, Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| 102 | その他 / 月次作業 | TIVNスタッフ労務管理 – Kiểm chấm công, tăng ca nhân viên TIVN | 月次 | Hori 1%, Chanh 1% | Thong 2%, Naito 1% |
| B43 | 前段取り / パターン | パターン連絡 – Chuyển dữ liệu rập cho nhà máy *(chỉ BTNB)* | 不定期（頻度高） | — | Tuong 1%, Nhi 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| B44 | 前段取り / パターン | チェック – Kiểm, báo TI khi lệch rập/mẫu gốc/tài liệu *(chỉ BTNB)* | 不定期（頻度中） | — | NEW STAFF 8%, H_Van 5%, Chau 5%, Minh 5%, Thu 5%, Hoa 5%, Nhi 3%, Tuong 1%, T_Van 1%, Tuan 1%, Trinh 1% |
| B45 | 前段取り / パターン | 修正 – Chỉ đạo nhà máy sửa lỗi rập *(chỉ BTNB)* | 不定期（頻度中） | — | NEW STAFF 4%, H_Van 1%, Chau 1%, Minh 1%, Thu 1%, Hoa 1% |
| B97 | その他 / 月次作業 | 経理業務 – Kế toán: chứng từ, thanh toán, cước *(chỉ BTNB)* | 月次 | — | Thong 2%, T_Van 1%, Trinh 1% |
