# Bàn giao: phân tích khối lượng và phân công nhóm MD APIC

Cập nhật: 28/10/2026. Nhánh: `claude/vigilant-fermi-3p4gcx` (đã push, chưa mở PR).

## 1. Tóm tắt

- **Mục tiêu**: dựng dữ liệu phân công công việc TIVN (APIC, BTNB) và phân tích khối lượng hiện tại của **nhóm MD (merchandiser) APIC** để quản lý cân bằng việc. Nhóm kỹ thuật và CAD sẽ phân tích riêng sau (chưa làm).
- **Trạng thái**: đã có phân tích mùa vừa qua, tải theo tháng 11/2026 – 2/2027 từ Plantable và file mẫu thật, các quyết định phân công của quản lý, bản Word/Excel để tải về. Không còn việc dở dang; đang chờ dữ liệu mới (kế hoạch mẫu 1st 28SS, Plantable cập nhật).
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

## 4. Kết quả chính (tải theo tháng, theo các quyết định đã chốt)

1,0 = khối lượng trung bình 1 người/1 tháng trong giai đoạn 11/2026 – 2/2027. ≥ 1,4 = vượt mức tăng ca 1 giờ/ngày.

| Tháng | Cường | Luân | Khánh | Bích Trang | Huy | Hiền | Như (chỉ phần brand) |
|---|---|---|---|---|---|---|---|
| 11/2026 | 1,71 | 1,30 | 0,78 | 1,42 | 1,65 | 1,48 | 1,05 |
| 12/2026 | 2,21 | 1,46 | 1,25 | 0,89 | 0,45 | 1,50 | 0,35 |
| 1/2027 | 1,32 | 1,27 | 0,81 | 0,15 | 1,61 | 0,89 | 1,17 |
| 2/2027 | 1,31 | 0,46 | 0,65 | 0,11 | 0,48 | 0,54 | 0,81 |

Điểm cần theo dõi: tháng 11 cả nhóm cao điểm (mẫu triển lãm 27FW + PPS 27SS + đơn mẫu tồn; phòng mẫu vượt công suất); Cường cao điểm tháng 12 (21 mã HELLY HANSEN vào chuyền); Huy cao điểm tháng 1 (25 mã TNF 別注 HOSOYA); Khánh tháng 12 ~1,25 (nên giữ hỗ trợ Luân ~25%).

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
| `APIC-MD-Best-Practice-Handbook.docx` | Định hướng MD (5 chức năng, KPI, chu kỳ tuần) | |

**Cách gán MD cho mã sản xuất** (Plantable không có cột MD): YONEX từ 27SS luôn gán Bích Trang; còn lại lấy người theo dõi mẫu PPS/RE-PPS/MOCKUP của mã đó trong sheet サンプル表2026 (khớp 251/363 lô); không khớp thì theo brand (TNF → Luân; TNF 別注: HSY → Huy, còn lại → Như; HH, NW → Cường; Goldwin, NANGA, CANTERBURY → Huy; ALLTERRAIN, TETON, DESCENTE/DCC, Mont-Bell, Snowpeak → Hiền; FOXFIRE, TATRAS, Kasco, YONEX (từ 27SS) → Bích Trang). Không dùng mẫu 1st/SMS để gán, vì mẫu TNF do Bích Trang làm nhưng đại trà do Luân/Cường theo dõi.

## 6. Mô hình và giả định

- **Chỉ số tải** = trung bình 4 yếu tố (mỗi yếu tố so với mức trung bình): số mã, ngày chuyền, ngày may mẫu / đơn mẫu, giờ dịch. Trọng số đều 25% (đã thử 4 cách đặt trọng số: Cường cao nhất, Như thấp nhất không đổi; chỉ vị trí Luân phụ thuộc trọng số).
- **Hệ số độ khó**: ALLTERRAIN, DESCENTE (DCC) × 1,3 (quản lý đánh giá: tài liệu, rập không rõ, TIJ hay thiếu NPL).
- **Công suất**: 9 giờ/ngày, Khánh 8 giờ. Việc giao thêm: Cường, Luân mỗi người ~1 giờ/ngày (10%) – **chưa có số thật**.
- **Tải theo tháng** (`ke-hoach-thang.py`): mã vào chuyền (mỗi mã tính 1 lần dù nhiều lô), ngày chuyền đang chạy (theo lô), đơn mẫu = đơn tồn (hạn ≤ 10/2026, chưa xong, dồn vào tháng 11) + số đơn cùng tháng năm trước theo brand; trọng số loại mẫu: 1st/0次 1,0; 2nd 0,8; PPS/mockup 0,6; SMS 0,5; khác 0,3. Mốc 1,0 chung cho cả 4 tháng.
- **Mùa vừa qua** (`build.py`): bulk 26FW + 27SS, mẫu 27SS; Khánh (nghỉ) và H.Trang (CAD) không tính vào mức trung bình.
- `du-bao.py` là bản ước cũ theo 2 giai đoạn (trước khi có Plantable), **không còn dùng** trong tài liệu.

## 7. Cách chạy lại

Chạy từ thư mục gốc repo (cần `openpyxl`, `pandoc`, `python-docx`, LibreOffice):

```
python3 data/phan-cong/build.py . data/phan-cong     # dữ liệu phân công (nếu đổi file nghiệp vụ)
python3 data/khoi-luong/build.py                     # mùa vừa qua → hien-tai.json, thuc-te.json
python3 data/khoi-luong/trich-ke-hoach.py            # trích Plantable + file mẫu → plantable-10-2026-3-2027.json, mau-2025-2026.json
python3 data/khoi-luong/ke-hoach-thang.py            # tải theo tháng → ke-hoach-thang.json
python3 data/khoi-luong/xuat-excel.py                # → docs/tai-ve/Phan-tich-khoi-luong-MD-APIC.xlsx
# tính lại công thức Excel bằng LibreOffice (vd. script recalc của skill xlsx) trước khi gửi
python3 data/khoi-luong/xuat-word.py docs/phan-tich-khoi-luong.md docs/tai-ve/Phan-tich-khoi-luong-MD-APIC.docx x
python3 data/khoi-luong/xuat-word.py docs/phan-cong-cong-viec.md docs/tai-ve/Phan-cong-cong-viec-TIVN.docx x
npm test                                              # test workflow lead time (13 test)
```

**Khi có Plantable/file mẫu mới**: lưu vào `data/nguon/`, chạy `python3 data/khoi-luong/trich-ke-hoach.py <plantable.xlsm> <sample.xlsx> [từ ngày] [đến ngày]` (mặc định 1/10/2026 – 31/3/2027; tự loại dòng mùa cũ hơn 26FW), rồi sửa danh sách `THANG` trong `ke-hoach-thang.py` nếu đổi giai đoạn. Script đã kiểm tra: chạy lại cho đúng các số trong tài liệu. Bảng số trong `docs/phan-tich-khoi-luong.md` (mục "Tải theo tháng") đang viết tay theo `ke-hoach-thang.json` → chạy lại thì phải cập nhật bảng.

## 8. Việc còn mở

1. **Kế hoạch mẫu 1st 28SS** (số mã theo brand): hiện ước theo cùng kỳ năm trước, nên Bích Trang, Huy tháng 12–2 có thể thấp hơn thực tế.
2. **Đại trà 27SS của FOXFIRE, TATRAS, Kasco, YONEX** (Bích Trang) gần như chưa có trong Plantable/BULK.
3. **Giờ thực tế** của việc điều phối (Cường) và việc quản lý giao cho Luân (đang tạm 1 giờ/ngày).
4. **Mẫu 26FW của Hiền** (sheet Hiền trong file SPL thiếu); YONEX chưa có giây TB (tạm 1 ngày/cái).
5. **Plantable dòng 1976–1996**: đã báo quản lý nên sửa; khi Plantable mới đã sửa thì không cần lọc thủ công.
6. Mẫu TNF别注 富山 (Như) và 東京 (Huy) mùa 25FW, 26SS có số giống hệt nhau trong file SPL – có thể nhập trùng.
7. **Phân tích nhóm kỹ thuật và CAD** (gồm H.Trang): quản lý muốn làm sau, tách riêng khỏi MD.
8. Đề xuất theo dõi định kỳ theo sổ tay MD: mỗi mùa chạy lại; báo cáo tuần (thứ Sáu); thử ghi giờ theo 5 chức năng MD trong 2–4 tuần để đặt trọng số theo thực tế.

## 9. Lưu ý khi làm việc tiếp

- Người dùng là quản lý nhóm MD, trao đổi bằng **tiếng Việt**, quen thuật ngữ Nhật trong sản xuất. Quyết định phân công do quản lý chốt; mô hình chỉ đưa phương án và số liệu.
- **Không ghi nhận xét cá nhân/tính cách** của nhân viên vào tài liệu (tài liệu có thể được chia sẻ); chỉ ghi quyết định ("giữ nguyên, không giao thêm").
- Khi người dùng gửi file mới: đọc, đối chiếu với bản cũ, chỉ ra chỗ số liệu lệch/không hợp lý trước khi tính.
- Người dùng muốn bản **Word/Excel** để tải về (gửi lại sau mỗi lần cập nhật).
- Các việc liên quan khai trị giá hải quan (phát triển SP, ra rập, nhảy size, tài liệu may) đổi người làm thì phải báo TI.
