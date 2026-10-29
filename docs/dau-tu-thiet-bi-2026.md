# Hồ sơ giải trình đầu tư thiết bị APIC 2026 / APIC 2026年度設備投資 説明資料

File chính: `dau-tu-thiet-bi-2026/APIC2026年度設備投資_説明資料_費用対効果.xlsx` (tiếng Nhật, dùng để trình TIJ).
Tạo lại file: `python3 dau-tu-thiet-bi-2026/build_workbook.py` rồi mở bằng Excel (Excel tự tính lại công thức).

Nguồn: `data/nguon/APIC2026年度設備投資予算案_20260403.xlsx` (danh sách đầu tư, biên bản họp 2/25 và 3/25) và các số liệu đã thống nhất trong trao đổi kỹ thuật.

## Logic trình bày

現状 → 課題 → 数値根拠 → 導入効果 → 費用対効果 → 投資妥当性 → 優先順位

Không chứng minh máy nào cũng phải mua. Chia thành:

- **必須投資**: an toàn (Spot Press), yêu cầu khách hàng (Nawon, GW 26FW), duy trì sản xuất (máy vắt sổ, DUMA, ép nóng lạnh 60×12).
- **改善投資**: Pattern Seamer, Spec-D, ép phẳng 125×83 → đánh giá bằng số, đưa vào theo từng giai đoạn.
- **戦略／検証投資**: Spec-IJ, US-512, máy 縮絨 (thử nghiệm xong mới quyết định).

## Số chính (tỷ giá 160 円/USD)

| | USD | 円 |
|---|---|---|
| Bản gốc (優先①) | 267,343 | 42.77 triệu (vượt 4,000万円) |
| **Phương án điều chỉnh** | **246,923** | **39.51 triệu (dưới mức trần 0.49 triệu)** |
| APIC chịu | 50,000 | 8.00 triệu |
| TI chịu | 196,923 | 31.51 triệu |

Điều chỉnh so với bản gốc:

| No. | Thiết bị | Gốc → Điều chỉnh | Lý do |
|---|---|---|---|
| 11 | Spot Press | 10 → 30 (+15,280 USD) | An toàn, ưu tiên số 1. Tạm tính: cần 45 − 15 máy mới đang có. Tự cập nhật khi nhập số liệu ở sheet 09 |
| 5 | Pattern Seamer lớn | 6 → 4 (−13,000) | Hiệu quả đo được bằng số mới có ~6,000 USD/năm (rental 3,016.54 + gia công quilting ~3,000). Mua tiếp sau khi đo Before/After |
| 7 | Seam Spec-D | 3 → 2 (−9,700) | Đo hiệu quả 2 người → 1 người trước, rồi mua thêm |
| 10 | Ép phẳng 125×83 | 2 → 1 (−13,000) | Xác lập cách kiểm chứng độ chính xác (⑤) trước |

Phương án cắt thêm nếu cần: No.14 HTM-8709 → H&H CSL-207 (−9,480 USD, cần sales xác nhận vì GW yêu cầu "Nawon 4 loại"), No.9 6→4 máy, đẩy US-512 / Spec-IJ sang ②.

## Các sheet trả lời từng mục 要確認事項

| Mục | Sheet | Tình trạng |
|---|---|---|
| ① Rental, 省人 | `12_レンタル費`, `07_省人効果`, `13_費用対効果PL` | Có khung + công thức, chờ số liệu |
| ② Vận hành Pattern Seamer | `08_PS運用設計` | Có phương án, chờ đo thực tế |
| ③ Spot Press | `09_スポットプレス` | Có công thức, chờ hiện trạng máy cũ |
| ⑤ Ép phẳng | `10_平プレス検証` | Có bảng 検証項目｜方法｜判定基準｜頻度｜担当, chờ quy cách của hãng |
| ⑦ 縮絨 | `11_縮絨試験` | Không đưa vào ngân sách 2026, thử nghiệm rồi quyết định |

Định nghĩa 省人化 (sheet 07): 工程秒削減 → 必要工数削減 → 必要人員減 → 余力人員再配置 → 工場全体の生産性向上. Bảng tính hiện cả 2 cơ sở: theo cycle time và theo 作業者拘束時間 (vì có case PS không nhanh hơn).

## Lỗi trong file gốc nên sửa trước khi gửi TIJ (sheet 14)

1. **R9 (DUMA)**: ghi tiết kiệm rental 13,440 USD/năm, nhưng tổng rental cả năm 2025 chỉ 4,890.75 USD (bỏ PS lớn còn 1,874.21). TIJ dễ bắt lỗi này.
2. R4 ghi "bình quân 4 máy/line" → logic mới là 3.2 máy/line (100 máy thường / 112 máy peak).
3. Công thức tổng I25, J25, O25 bỏ sót dòng 4–11 → ưu tiên ② hiện 10 máy (thực tế 38), 258,730 USD (thực tế 355,230).
4. M27 lỗi tham chiếu, No.18 bị trùng, R5 (67 máy) lệch S5 (68 máy), tỷ giá lẫn 150 và 160.
5. Sheet biên bản họp trong file gốc có tên người phát biểu (gồm 上田課長) → bỏ khỏi bản gửi chính thức. File giải trình mới không nhắc tên này.

## Số liệu còn thiếu (ô vàng trong file)

- Chi tiết rental 2024, 2026 (máy, số lượng, thời gian, D/N, style/line); số tiền rental máy vắt sổ 10–12/2024.
- Đo Before/After Pattern Seamer và Spec-D (工程秒, 自動縫製秒, thời gian làm song song) + sản lượng năm theo mã.
- Spot Press cũ: số lượng, năm, sửa được không, số máy dùng theo line, nhu cầu peak, ảnh, near miss.
- Tiêu chuẩn của hãng cho ép phẳng (độ phẳng, nhiệt độ, áp lực) và chu kỳ hiệu chuẩn.
- APIC 50,000 USD: máy nào, ai sở hữu tài sản, khấu hao thế nào (P/L đang để 2 trường hợp A/B).
- Chi phí nhân công/người/năm (để quy đổi 省人 ra tiền).
- Thống nhất tỷ giá (150 hay 160).
