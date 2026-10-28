# Bàn giao: phân tích khối lượng và phân công nhóm MD APIC

Cập nhật: 28/10/2026. Nhánh: `claude/vigilant-fermi-3p4gcx` (đã push, chưa mở PR).

## 1. Tóm tắt

- **Mục tiêu**: dựng dữ liệu phân công công việc TIVN (APIC, BTNB) và phân tích khối lượng hiện tại của **nhóm MD (merchandiser) APIC** để quản lý cân bằng việc. Nhóm kỹ thuật và CAD sẽ phân tích riêng sau (chưa làm).
- **Trạng thái**: đã có phân tích mùa vừa qua, **tải theo tháng cả vòng 12 tháng 11/2026 – 10/2027** (mẫu và sản xuất của hai mùa chồng nhau, theo ngày làm việc thực), các quyết định phân công của quản lý, bản Word/Excel để tải về. Không còn việc dở dang; đang chờ dữ liệu mới (kế hoạch mẫu 1st 28SS, Plantable cập nhật).
- **Tài liệu chính**: `docs/phan-tich-khoi-luong.md` (phân tích, quyết định, dự báo), `docs/phan-cong-cong-viec.md` (dữ liệu phân công, phân công brand), bản tải về trong `docs/tai-ve/`.

## 2. Nhân sự nhóm MD (tên dùng trong dữ liệu)

| Tên trong file | Người | Ghi chú |
|---|---|---|
| Cuong | Cường | HELLY HANSEN (chủ lực); điều phối xuất hàng chung, quản lý chung may mẫu |
| Luan | Luân | TNF 富山企画 (TNF(GTC企画)), PPS~đại trà; nhận thêm một phần việc quản lý |
| Khanh | Khánh | Đi làm lại từ giữa 10/2026 sau nghỉ thai sản; làm đến 16h, tăng ca được đến 17h |
| B.Trang / Bich Trang / "B Trang" | Bích Trang | Nhân viên mới; TNF (ALL) mẫu 1st–SMS từ 27SS, Premium Snow 別注, FOXFIRE, TATRAS, Kasco, YONEX từ S27 |
| Q.Huy / Huy | Huy | Goldwin (xuyên suốt), TNF 別注 HOSOYA, 141 (hỗ trợ), CANTERBURY, NANGA |
| Hien | Hiền | Teton Bros, ALLTERRAIN (+81), Mont-Bell, DESCENTE (GOLF, SKI), Snowpeak |
| Nhu / Như | Như | TNF 別注 SHIGEMOTO/NAKAO, UTMF, TOKYO MARATHON + việc văn phòng (~30%: báo cáo tháng, hóa đơn…) |
| H.Trang / "Trang" (trong file mẫu) | Huyền Trang | **Nhóm CAD**, chỉ phụ TNF HAT&CAP, 141 → không tính trong phân tích MD |

Viết tắt brand: HH = HELLY HANSEN, NW = NEUTRALWORKS, DCC = DESCENTE (TI: WTNB). Viết tắt người phía TI: SGM = Shigemoto, HSY = Hosoya, NKO = Nakao, WTNB/WTB = Watanabe/Watabe, KBR = Kobori.

## 3. Quyết định của quản lý (đã chốt)

| Ngày | Quyết định |
|---|---|
| 10/2026 | Giờ làm 8 giờ + 1 giờ tăng ca/ngày; Khánh 8 giờ |
| 10/2026 | Hiền giữ nguyên, **không giao thêm việc** |
| 10/2026 | Huy giữ nguyên, Goldwin xuyên suốt cả mẫu và đại trà |
| 10/2026 | Như giữ nguyên, theo dõi thêm |
| 10/2026 | NEUTRALWORKS: Cường → Khánh |
| 10/2026 | Cường giữ HELLY HANSEN, nhận điều phối xuất hàng chung và quản lý chung may mẫu |
| 10/2026 | Luân giữ TNF 富山企画 (đầu mối duy nhất với TIJ). Khánh hỗ trợ **theo chức năng, không chia brand**: báo cáo mã ngoài quy tắc 2 tuần, tiến độ NPL, BL/P/L (cc Luân), dịch một phần tài liệu (~25% việc Luân tháng 11–12, ~35% tháng 1–2) |
| 28/10/2026 | Mẫu triển lãm 27FW: người đang làm làm tiếp (đã vào may). Khánh nhận 1/3 mẫu TNF của Bích Trang từ mẫu 1st 28SS |
| 28/10/2026 | **TNF 東京企画: Cường → Khánh từ 11/2026** |
| 28/10/2026 | Hỗ trợ tạm theo tháng (Huy, Như hỗ trợ Cường tháng 12; Bích Trang hỗ trợ Huy tháng 1): **chưa cần**, giữ làm phương án dự phòng |
| 28/10/2026 | Snowpeak do Hiền theo dõi |

## 4. Kết quả chính (tải 12 tháng, theo các quyết định đã chốt)

1,0 = khối lượng trung bình cho 1 ngày công trong cả 12 tháng; ≥ 1,4 = vượt mức tăng ca 1 giờ/ngày. Từ 4/2027 là số ước theo cùng kỳ năm trước.

| Tháng | Cường | Luân | Khánh | Bích Trang | Huy | Hiền | Như (chỉ phần brand) |
|---|---|---|---|---|---|---|---|
| 11/2026 | 1,72 | 1,85 | 0,97 | 1,42 | 1,69 | 1,46 | 1,13 |
| 12/2026 | 1,94 | 1,54 | 1,21 | 0,55 | 0,41 | 1,31 | 0,31 |
| 1/2027 | 1,37 | 1,30 | 0,88 | 0,16 | 1,57 | 0,87 | 1,21 |
| 2/2027 | 2,04 | 0,87 | 0,96 | 0,17 | 0,68 | 0,93 | 1,20 |
| 3/2027 | 0,80 | 1,30 | 0,93 | 0,10 | 0,43 | 0,60 | 1,12 |
| 4/2027 | 0,96 | 1,69 | 1,12 | 0,41 | 1,03 | 0,55 | 2,56 |
| 5/2027 | 0,85 | 1,52 | 1,04 | 0,40 | 1,69 | 0,89 | 0,86 |
| 6/2027 | 0,97 | 1,11 | 1,22 | 0,77 | 1,30 | 1,34 | 0,27 |
| 7/2027 | 2,88 | 0,70 | 0,74 | 0,73 | 1,51 | 0,61 | 0,86 |
| 8/2027 | 0,90 | 0,75 | 0,60 | 0,42 | 1,33 | 0,56 | 0,82 |
| 9/2027 | 1,01 | 0,65 | 0,46 | 0,41 | 0,89 | 0,48 | 0,55 |
| 10/2027 | 1,56 | 0,88 | 1,01 | 0,65 | 1,14 | 1,27 | 0,84 |

Điểm cần theo dõi: tháng 11 nặng nhất cả nhóm; Cường cao điểm 12, 2 (Tết ít ngày làm, 2,04), **7** (mẫu 1st FW HELLY HANSEN), 10; Luân 11–12 và **3–5** (hàng FW vào chuyền); Huy 11, 1, **5**; **Như tháng 4** (40 mã TNF 別注 FW vào chuyền); Hiền 11–12, 6. Bích Trang thấp có thể do thiếu dữ liệu các brand mới.

## 5. Dữ liệu nguồn và các điểm dễ sai

Tất cả file gốc ở `data/nguon/`.

| File | Dùng cho | Lưu ý |
|---|---|---|
| `Thuc-te-BULK-19FW-27SS-20261012.xlsx` | Đại trà theo người/mùa/brand | **Dùng sheet từng người**; sheet "纏め" cũ hơn |
| `Thuc-te-mau-SPL-23FW-27FW-20261012.xlsx` | Mẫu theo người/mùa/brand, giờ dịch/mã theo brand | **Dùng sheet từng người** (khớp "全体纏め"); "シーズンごと纏め" lệch nhiều. Sheet Hien thiếu mẫu 26FW |
| `APIC-P-Plan-Plantable-2026-20261026.xlsm` | Lịch sản xuất theo chuyền (sheet PLANTABLE (原価) 2026年), sổ theo dõi đơn S27 (S27 受注・進捗管理台帳) | **Sheet S27本発注 là bảng mô phỏng** (ngày kéo đến 2034), không dùng làm lịch. **Dòng 1976–1996 (chuyền AP1B-9) là đơn cũ 23SS/23FW bị đẩy ngày** sang 28/11/2026 – 19/10/2027 (ô START dòng 1976 gõ cứng, các dòng sau nối công thức); đã loại 8 dòng rơi vào giai đoạn phân tích, đã báo quản lý nên sửa trong Plantable |
| `APIC-Sample-20261023.xlsx` | Đơn mẫu 2025–2026 (sheet サンプル表2025/2026: MD, loại mẫu, ngày cần, trạng thái ●/×), công suất phòng mẫu (サンプル状況2026) | Cột khác nhau giữa sheet 2025 và 2026; tên người viết nhiều kiểu (Như/Nhu, Luân/Luan, "B Trang", "Trang" = H.Trang) |
| `Phan-cong-brand-MD-APIC-20260722.xlsx` | Phân công brand 22/7/2026 (sheet hiển thị) + lịch sử 2017–2018 (sheet ẩn) | (1)/(2) = phòng 1課/2課 phía TI |
| `Phan-tich-nghiep-vu-TIVN-APIC-2026-thuong-ky.xlsx` | % thời gian theo ~100 việc của 18 người (2026 上期) | Số tự khai, chia gần đều; không phản ánh khối lượng thực |
| `Hop-TIVN-phi-uy-thac-APIC-BTNB-20240130.xlsx` | Biên bản họp phí ủy thác, khai trị giá hải quan (評価加算), bảng nghiệp vụ 2023–2024 | Đổi người làm các việc tính thuế (phát triển SP, rập, nhảy size, tài liệu may…) phải báo TI |
| `Lich-nghi-thu-7-TIVN-APIC-2026-20260811.xlsx` | Lịch nghỉ thứ 7 theo nhóm A (Luân, Khánh, Như, Huy, Bích Trang…) / B (Cường, Hiền…), kỳ lương 26–25 | Chỉ đến 3/2027; ngày lễ, Tết lấy theo lịch nghỉ công ty quản lý báo 28/10/2026 (ghi trong `lich_lam_viec.py`) |
| `APIC-MD-Best-Practice-Handbook.docx` | Định hướng MD (5 chức năng, KPI, chu kỳ tuần) | |

**Cách gán MD cho mã sản xuất** (Plantable không có cột MD): YONEX từ 27SS luôn gán Bích Trang; còn lại lấy người theo dõi mẫu PPS/RE-PPS/MOCKUP của mã đó trong sheet サンプル表2026 (khớp 251/363 lô); không khớp thì theo brand (TNF → Luân; TNF 別注: HSY → Huy, còn lại → Như; HH, NW → Cường; Goldwin, NANGA, CANTERBURY → Huy; ALLTERRAIN, TETON, DESCENTE/DCC, Mont-Bell, Snowpeak → Hiền; FOXFIRE, TATRAS, Kasco, YONEX (từ 27SS) → Bích Trang). Không dùng mẫu 1st/SMS để gán, vì mẫu TNF do Bích Trang làm nhưng đại trà do Luân/Cường theo dõi.

## 6. Mô hình và giả định

- **Chỉ số tải** = trung bình 4 yếu tố (mỗi yếu tố so với mức trung bình): số mã, ngày chuyền, ngày may mẫu / đơn mẫu, giờ dịch. Trọng số đều 25% (đã thử 4 cách đặt trọng số: Cường cao nhất, Như thấp nhất không đổi; chỉ vị trí Luân phụ thuộc trọng số).
- **Hệ số độ khó**: ALLTERRAIN, DESCENTE (DCC) × 1,3 (quản lý đánh giá: tài liệu, rập không rõ, TIJ hay thiếu NPL).
- **Công suất**: 9 giờ/ngày, Khánh 8 giờ. Việc giao thêm: Cường, Luân mỗi người ~1 giờ/ngày (10%) – **chưa có số thật**.
- **Tải 12 tháng** (`ke-hoach-nam.py`, dùng chính): sản xuất mùa đến 27SS lấy lịch thật trong Plantable; 27FW, 28SS lấy lịch của 26FW, 27SS năm trước dời 52 tuần; mẫu = đơn tồn (dồn vào 11/2026) + cùng tháng năm trước; công suất = ngày làm việc thực từng người (`lich_lam_viec.py`: thứ 2–7, nghỉ thứ 7 theo nhóm, trừ lễ; lễ theo lịch công ty: 24/11/2026, 1/1/2027, **Tết dự kiến 4–14/2/2027**, 30/4–1/5/2027; Giỗ Tổ 16/4 và Quốc khánh 2–3/9/2027 tạm theo luật) × giờ/ngày; mốc 1,0 = trung bình 1 ngày công cả 12 tháng.
- **Tải 4 tháng** (`ke-hoach-thang.py`, bản trước, không còn trong tài liệu): mã vào chuyền (mỗi mã tính 1 lần dù nhiều lô), ngày chuyền đang chạy (theo lô), đơn mẫu = đơn tồn (hạn ≤ 10/2026, chưa xong, dồn vào tháng 11) + số đơn cùng tháng năm trước theo brand; trọng số loại mẫu: 1st/0次 1,0; 2nd 0,8; PPS/mockup 0,6; SMS 0,5; khác 0,3. Mốc 1,0 chung cho cả 4 tháng.
- **Mùa vừa qua** (`build.py`): bulk 26FW + 27SS, mẫu 27SS; Khánh (nghỉ) và H.Trang (CAD) không tính vào mức trung bình.
- `du-bao.py` là bản ước cũ theo 2 giai đoạn (trước khi có Plantable), **không còn dùng** trong tài liệu.

## 7. Cách chạy lại

Chạy từ thư mục gốc repo (cần `openpyxl`, `pandoc`, `python-docx`, LibreOffice):

```
python3 data/phan-cong/build.py . data/phan-cong     # dữ liệu phân công (nếu đổi file nghiệp vụ)
python3 data/khoi-luong/build.py                     # mùa vừa qua → hien-tai.json, thuc-te.json
python3 data/khoi-luong/trich-ke-hoach.py data/nguon/APIC-P-Plan-Plantable-2026-20261026.xlsm data/nguon/APIC-Sample-20261023.xlsx 2025-11-01 2027-04-30 data/khoi-luong/plantable-11-2025-4-2027.json
                                                      # trích Plantable (11/2025 – 4/2027) + file mẫu → mau-2025-2026.json
python3 data/khoi-luong/ke-hoach-nam.py              # tải 12 tháng → ke-hoach-nam.json
python3 data/khoi-luong/lich_lam_viec.py             # xem số ngày làm việc từng tháng
python3 data/khoi-luong/xuat-excel.py                # → docs/tai-ve/Phan-tich-khoi-luong-MD-APIC.xlsx
# tính lại công thức Excel bằng LibreOffice (vd. script recalc của skill xlsx) trước khi gửi
python3 data/khoi-luong/xuat-word.py docs/phan-tich-khoi-luong.md docs/tai-ve/Phan-tich-khoi-luong-MD-APIC.docx x
python3 data/khoi-luong/xuat-word.py docs/phan-cong-cong-viec.md docs/tai-ve/Phan-cong-cong-viec-TIVN.docx x
npm test                                              # test workflow lead time (13 test)
```

**Khi có Plantable/file mẫu mới**: lưu vào `data/nguon/`, chạy `python3 data/khoi-luong/trich-ke-hoach.py <plantable.xlsm> <sample.xlsx> [từ ngày] [đến ngày]` (lấy từ 12 tháng trước giai đoạn tính để có lịch năm trước; tự loại dòng mùa cũ hơn 26SS), rồi sửa `THANG` và `MUA_CO_KE_HOACH` trong `ke-hoach-nam.py` khi chuyển sang giai đoạn mới. Khi có lịch nghỉ thứ 7 / lịch Tết mới thì sửa `lich_lam_viec.py`. Bảng số trong `docs/phan-tich-khoi-luong.md` (mục "Tải theo tháng cả vòng 12 tháng") được viết theo `ke-hoach-nam.json` → chạy lại thì phải cập nhật bảng.

## 8. Việc còn mở

1. **Kế hoạch mẫu 1st 28SS** (số mã theo brand): hiện ước theo cùng kỳ năm trước, nên Bích Trang, Huy tháng 12–2 có thể thấp hơn thực tế.
2. **Đại trà 27SS của FOXFIRE, TATRAS, Kasco, YONEX** (Bích Trang) gần như chưa có trong Plantable/BULK.
3. **Giờ thực tế** của việc điều phối (Cường) và việc quản lý giao cho Luân (đang tạm 1 giờ/ngày).
4. **Mẫu 26FW của Hiền** (sheet Hiền trong file SPL thiếu); YONEX chưa có giây TB (tạm 1 ngày/cái).
5. **Plantable dòng 1976–1996**: đã báo quản lý nên sửa; khi Plantable mới đã sửa thì không cần lọc thủ công.
6. Mẫu TNF别注 富山 (Như) và 東京 (Huy) mùa 25FW, 26SS có số giống hệt nhau trong file SPL – có thể nhập trùng.
7. **Phân tích nhóm kỹ thuật và CAD** (gồm H.Trang): quản lý muốn làm sau, tách riêng khỏi MD.
8. **Lịch Tết 2027 chính thức** (đang tính theo dự kiến 4–14/2), ngày nghỉ Giỗ Tổ và Quốc khánh 2027, lịch nghỉ thứ 7 sau 3/2027 (đang giả định). **Kế hoạch 27FW/28SS thật** để thay số ước từ 4/2027.
9. Đề xuất theo dõi định kỳ theo sổ tay MD: mỗi mùa chạy lại; báo cáo tuần (thứ Sáu); thử ghi giờ theo 5 chức năng MD trong 2–4 tuần để đặt trọng số theo thực tế.

## 9. Lưu ý khi làm việc tiếp

- Người dùng là quản lý nhóm MD, trao đổi bằng **tiếng Việt**, quen thuật ngữ Nhật trong sản xuất. Quyết định phân công do quản lý chốt; mô hình chỉ đưa phương án và số liệu.
- **Không ghi nhận xét cá nhân/tính cách** của nhân viên vào tài liệu (tài liệu có thể được chia sẻ); chỉ ghi quyết định ("giữ nguyên, không giao thêm").
- Khi người dùng gửi file mới: đọc, đối chiếu với bản cũ, chỉ ra chỗ số liệu lệch/không hợp lý trước khi tính.
- Người dùng muốn bản **Word/Excel** để tải về (gửi lại sau mỗi lần cập nhật).
- Các việc liên quan khai trị giá hải quan (phát triển SP, ra rập, nhảy size, tài liệu may) đổi người làm thì phải báo TI.
