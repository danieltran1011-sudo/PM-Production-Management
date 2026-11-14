# -*- coding: utf-8 -*-
"""Vietnamese -> Japanese translation pass for the PLANTABLE executive workbook (in-memory)."""
import re
from copy import copy

SHEETS = {
    '0. Mục lục': '0. 目次', '1. Tóm tắt điều hành': '1. エグゼクティブサマリー', '1A. CEO': '1A. CEO',
    '1B. GĐ Nhà máy': '1B. 工場長', '1C. QL Sản xuất': '1C. 生産管理', '2. Mùa vụ': '2. シーズン別',
    '3. Khách hàng': '3. 顧客別', '4. Công suất': '4. 生産能力', '5. Chi phí & Lô': '5. 加工費とロット',
    '6. Rủi ro giao hàng': '6. 納期リスク', '7. Chất lượng DL': '7. データ品質', 'Data': 'データ', 'Phương pháp': '前提・定義', '2A. So sánh mùa': '2A. シーズン比較', '4A. 受注進行表': '4A. 受注進行表',
}

EX = {
 # --- generic labels / statuses ---
 'Có': 'あり', 'Mùa': 'シーズン', 'Đầy': 'フル稼働', 'Mẫu': 'サンプル', 'Mới': '新規', 'Năm': '年', 'Cột': '列', 'Trễ': '遅れ',
 'Thấp': '低', 'Việc': '項目', 'Tổng': '合計', 'Khác': 'その他', 'Số CN': '人数', 'trống': '空き', 'Tháng': '月', 'Mẫu; ': 'サンプル; ',
 'Xưởng': '工場', 'Số đơn': 'オーダー数', 'Chỉ số': '指標', 'số đơn': 'オーダー数', 'Thứ Tư': '水曜', 'Vấn đề': '問題', 'Tất cả': '全員',
 'Thường': '通常', 'Quy mô': '規模', 'Mức ETD': 'ETDレベル', 'Mẫu gửi': 'サンプル送付', 'Thứ Sáu': '金曜', 'So sánh': '比較',
 'Nhóm lô': 'ロット区分', 'Giá trị': '値', 'Từ 27FW': '27FWから', '(trống)': '(空欄)', 'Thứ Hai': '月曜', 'Ghi chú': '備考',
 'Nội dung': '内容', 'Thời hạn': '期限', 'Sẵn sàng': '準備完了', 'Kết luận': '判定', 'Kiểm tra': '要確認', 'Đúng hạn': '納期内',
 'Tỷ trọng': '構成比', 'Cụm line': 'ライン番号', 'Đánh giá': '評価', 'Liên tục': '継続', 'Tác động': '影響', 'Theo dõi': '注意',
 'Số style': '品番数', 'Dòng gốc': '元の行', 'Hành động': 'アクション', 'Brand gốc': 'ブランド(元)', 'Đọc nhanh': '一言評価',
 'Hằng ngày': '毎日', 'Người đọc': '読者', 'Tăng/giảm': '増減', 'Tổng / TB': '合計 / 平均', 'Đọc trước': '最初に読む',
 'Giờ chuẩn': '標準時間(h)', 'Thiếu đơn': '受注不足', 'Hàng tuần': '毎週', 'Thiếu SAM': 'SAM未入力', 'Còn trống': '余力あり',
 'Tổng quan': '概況', '% số dòng': '行数比', 'Quyết định': '意思決定事項', 'Chưa duyệt': '未承認', 'Hàng tháng': '毎月',
 'Trạng thái': '状態', 'Trung bình': '中', 'NVL về HCM': '資材HCM着', 'Thiếu 工場ETD': '工場ETD未入力', 'Mã xưởng cũ': '旧工場コード',
 'Đơn còn lại': '残オーダー', 'Số dòng đơn': 'オーダー行数', 'Đang/sắp SX': '生産中/予定', 'Lô SX (GV2)': '生産ロット(原価2)',
 'so với 2025': '2025年比', 'so với 26SS': '26SS比', 'Thiếu START': 'START未入力', 'Thiếu Brand': 'ブランド未入力',
 '% sản lượng': '数量比', 'Tên cột gốc': '元の列名', 'Lô TB (pcs)': '平均ロット(pcs)', 'Brand trống': 'ブランド空欄',
 '(không tên)': '(列名なし)', 'Số lượng đơn': '受注数量', 'Nghiêm trọng': '重大', 'Lô đặt (GV1)': '発注ロット(原価1)',
 'Mẫu phản hồi': 'サンプル回答', 'Đơn sẵn sàng': '準備完了オーダー', 'Hiệu suất lô': 'ロット効率', 'Chưa có ngày': '日付なし',
 'Lọc/đếm sai.': '抽出・集計ミスの原因。', 'Tỷ lệ sử dụng': '稼働率', 'Đề xuất xử lý': '対策案', 'Ban điều hành': '経営陣',
 'Giai đoạn mẫu': 'サンプル段階', 'Tổng SL (pcs)': '合計数量(pcs)', 'Kế hoạch 2027': '2027年計画', 'Mức ảnh hưởng': '影響度',
 'Chưa xác định': '未定', 'riêng mùa 27SS': '27SSのみ', 'lô 2.001–5.000': 'ロット2,001–5,000', 'Vượt định biên': '定員超過',
 'Định biên (CN)': '定員(人)', 'Áp dụng từ 27FW': '27FWから適用', 'Số dòng lô ≤300': '≤300ロット行数', 'Công-ngày trống': '空き人日',
 'Hiệu suất lô TB': '平均ロット効率', 'Danh sách sheet': 'シート一覧', 'Ô lỗi công thức': '数式エラーセル', 'Chỉ số vận hành': '操業指標',
 'Đơn thiếu 工場ETD': '工場ETD未入力オーダー', 'Lệch ETD (ngày)': 'ETD差(日)', 'Ý nghĩa / xử lý': '意味 / 処理', 'Sản lượng (pcs)': '数量(pcs)',
 'Tuần 47–48/2026': '2026年第47–48週', 'Chi phí chia lô': 'ロット分割コスト', 'SAM TB (phút/sp)': '平均SAM(分/枚)', 'Tác động kỳ vọng': '期待効果',
 'trên tổng số đơn': '全オーダーに対して', 'Đọc thêm khi cần': '必要に応じて', 'Tỷ trọng hàng 別注': '別注比率', 'Trước 15/12/2026': '2026/12/15まで',
 'Nội dung đề xuất': '提案内容', 'Đơn đang/ sắp SX': '生産中/予定オーダー', 'Quản lý sản xuất': '生産管理', 'Cờ chất lượng DL': 'データ品質フラグ',
 'Nhóm lô SX (pcs)': '生産ロット区分(pcs)', 'Trước 30/11/2026': '2026/11/30まで', 'Công-ngày có đơn': '受注あり人日', 'Ký tự full-width': '全角文字',
 'Mức tải mục tiêu': '目標稼働率', 'Năng suất lô nhỏ': '小ロットの生産性', 'Giám đốc nhà máy': '工場長', 'Kế hoạch SX / IT': '生産計画 / IT',
 'SL còn lại (pcs)': '残数量(pcs)', 'CP gia công 1 / 2': '加工費1 / 2', 'Phương pháp; Data': '前提・定義; データ', 'Loại trừ (mùa cũ)': '除外(旧シーズン)',
 'CP gia công (USD)': '加工費(USD)', 'Ổn định doanh thu': '売上の安定', 'So sánh / ghi chú': '比較 / 備考', 'USD/pcs bình quân': '平均USD/pcs',
 'Có thể đưa vào SX': '投入可能日', 'Định nghĩa chỉ số': '指標の定義', 'Kế hoạch Tết 2027': '2027年テト計画', 'Mua hàng / QA mẫu': '購買 / サンプルQA',
 'Bức tranh 30 giây': '30秒サマリー', '% chi phí chia lô': 'ロット分割コスト率', 'Đơn chưa duyệt mẫu': 'サンプル未承認オーダー',
 'Giảm số lần đổi mã': '品番切替回数の削減', 'yêu cầu Sales nhập': '営業に入力依頼', 'Sản lượng mùa 27SS': '27SS数量',
 'Line trống từ ngày': 'ライン空き開始日', 'Tổng hợp công suất': '生産能力まとめ', 'Mã hàng / tên hàng': '品番 / 品名',
 'B. PHÁT HIỆN CHÍNH': 'B. 主な発見', 'Theo năm (SS + FW)': '年別 (SS + FW)', 'SL xưởng khác (pcs)': '他工場数量(pcs)',
 'END+38 / chênh lệch': 'END+38 / 差', 'Kế hoạch SX + Sales': '生産計画 + 営業', 'Khó đo tiến độ mẫu.': 'サンプル進捗を測れない。',
 'riêng T2/2027 (Tết)': '2027年2月(テト)', 'Đơn ETD trễ ≥8 ngày': 'ETD 8日以上遅れのオーダー', 'Ngưỡng lô nhỏ (pcs)': '小ロット閾値(pcs)',
 'Ghi chú / phương án': '備考 / 代替案', 'Chuẩn hoá danh mục.': 'マスタを統一する。', 'Đẩy NVL & duyệt mẫu': '資材・サンプル承認の促進',
 'Số đơn đang/ sắp SX': '生産中/予定オーダー数', 'Phụ trách (đề xuất)': '担当(案)', 'Ngày NVL vào SX được': '資材投入可能日',
 'Tỷ trọng Top 3 brand': '上位3ブランド比率', 'Lô SX trung vị (pcs)': '生産ロット中央値(pcs)', 'Tổng hợp theo mùa vụ': 'シーズン別集計',
 'Đơn giá GC (USD/pcs)': '加工単価(USD/pcs)', 'Câu hỏi được trả lời': '答えられる問い', 'Tách cột "別注 (Y/N)".': '「別注(Y/N)」列を分ける。',
 'Cấu hình đơn sắp tới': '今後の編成', 'Ngày làm việc (T2–T7)': '稼働日(月–土)', 'Công suất (công-ngày)': '能力(人日)',
 'Bằng QTY × SAM / 3600': 'QTY × SAM / 3600', 'Đơn giá GC2 (USD/pcs)': '加工単価2(USD/pcs)', 'Ghi chú gốc (rút gọn)': '元コメント(抜粋)',
 'Chi phí chia lô (USD)': 'ロット分割コスト(USD)', 'Tỷ trọng TNF (gồm 別注)': 'TNF比率(別注含む)', 'Brand ghi lẫn ghi chú': 'ブランド欄に備考が混在',
 '4. Xử lý đơn 27SS trễ': '4. 27SS遅延オーダーの対応', 'công-ngày cần bổ sung': '追加が必要な人日', 'Đơn giá GC bất thường': '加工単価の異常値',
 'Đơn giá GC1 (USD/pcs)': '加工単価1(USD/pcs)', 'NVL về sau ngày START': '資材着がSTARTより後', 'Xu hướng 27SS vs 26SS': '27SS vs 26SS',
 'bình quân T1–T10/2026': '2026年1–10月平均', 'Đặt tên cột hoặc xoá.': '列名を付けるか削除する。', 'Ban điều hành + Sales': '経営陣 + 営業',
 'NVL chưa xác định (未定)': '資材未定', 'Chi phí gia công (USD)': '加工費(USD)', 'ETD xưởng được yêu cầu': '要求された工場ETD',
 'Giai đoạn mẫu nhiều mã': 'サンプル段階コストが多様', 'Giá trị gia công (USD)': '加工金額(USD)', 'ｺﾒﾝﾄ (công thức =U−AJ)': 'ｺﾒﾝﾄ (数式 =U−AJ)',
 'SL kế hoạch APIC (pcs)': 'APIC計画数量(pcs)', 'Ngày có thể đưa vào SX': '投入可能日', 'Tăng/giảm 27SS vs 26SS': '27SS vs 26SS増減',
 'Chính sách lô tối thiểu': '最小ロット方針', '% trên chi phí gia công': '加工費に対する比率', 'Cột BL không có tiêu đề': 'BL列に見出しなし',
 'Tỷ lệ sử dụng T3–T6/2027': '稼働率 2027年3–6月', 'Lấp công suất T3–T6/2027': '2027年3–6月の能力を埋める', 'Bọc IFERROR / sửa nguồn.': 'IFERRORで囲む/元データ修正。',
 'Báo cáo tự động, tin cậy': 'レポートの自動化・信頼性向上', 'C. KHUYẾN NGHỊ HÀNH ĐỘNG': 'C. 推奨アクション', 'Hiệu suất lô TB (平均/MAX)': '平均ロット効率(平均/MAX)',
 'Sản lượng năm 2026 (pcs)': '2026年数量(pcs)', 'Giảm dừng line, đổi lịch': 'ライン停止・計画変更の削減', 'Xu hướng lô nhỏ theo mùa': 'シーズン別小ロット推移',
 'trên tổng SL đang/sắp SX': '生産中/予定数量に対して', 'Sản lượng mùa 27SS (pcs)': '27SS数量(pcs)', 'Line và số công nhân/line': 'ラインとライン人数',
 'Tết Nguyên đán 06/02/2027': '旧正月 2027/02/06', 'Cơ cấu khách hàng / brand': '顧客 / ブランド構成', 'Mức "Cao" (trễ 8–29 ngày)': '「高」(8–29日遅れ)',
 'CP GC2 – theo lô SX (USD)': '加工費2 – 生産ロット基準(USD)', 'Danh mục khách hàng (pcs)': '顧客ポートフォリオ(pcs)', 'Lệch ETD (cột S) vô nghĩa': 'ETD差(S列)が無意味な値',
 'Dòng mùa cũ còn treo lịch': '旧シーズン行が計画に残存', 'Mùa cũ nhưng lịch SX muộn': '旧シーズンだが生産日程が遅い', '▲ Lô nhỏ đắt gấp ~2,7 lần': '▲ 小ロットは約2.7倍の単価',
 'lần đổi mã / line / tháng': '回 / ライン / 月', '3. Đa dạng hoá khách hàng': '3. 顧客の分散', 'Mã mùa vụ không thống nhất': 'シーズンコードの表記ゆれ',
 '2. Chính sách lô tối thiểu': '2. 最小ロット方針', 'Mẫu chưa có phản hồi/duyệt': 'サンプル未回答/未承認', 'Chỉ số sức khoẻ kinh doanh': '事業健全性指標',
 'pcs cần bổ sung để đạt 90%': '90%達成に必要な追加数量', 'tỷ lệ mùa 27SS (đang giảm)': '27SSの比率(低下中)', 'Mã xưởng cũ AP1A/AP1B/AP1C': '旧工場コード AP1A/AP1B/AP1C',
 'Tỷ trọng TNF riêng mùa 27SS': '27SSのTNF比率', 'Định biên CN may xưởng APIC': 'APIC縫製定員', 'Thiếu 工場ETD (đơn tương lai)': '工場ETD未入力(今後のオーダー)',
 'Số mã hàng bắt đầu (đổi mã)': '着手品番数(切替)', 'Số đơn bắt đầu trong 6 tuần': '6週間以内に着手するオーダー数', '1. Lấp công suất T3–T6/2027': '1. 2027年3–6月の能力を埋める',
 'Ngưỡng "Cao" (ngày trễ ETD)': '「高」の閾値(ETD遅れ日数)', 'Chỉ số tập trung khách hàng': '顧客集中度', 'Tỷ lệ sử dụng T2/2027 (Tết)': '稼働率 2027年2月(テト)',
 'Chuẩn hoá dữ liệu PLANTABLE': 'PLANTABLEデータの標準化', 'Không có header – không dùng': '見出しなし – 使用せず', 'nếu không lấp được công suất': '能力を埋められない場合',
 'Kiểm soát đơn chạy song song': '並行オーダーの管理', 'CP GC1 – theo đặt hàng (USD)': '加工費1 – 発注条件基準(USD)', 'Quyết định cần CEO phê duyệt': 'CEO承認が必要な事項',
 'tỷ trọng số dòng đơn lô ≤300': '≤300ロットの行数比', 'Tỷ trọng TNF trong sản lượng': '数量に占めるTNF比率', 'Chi phí gia công & quy mô lô': '加工費とロット規模',
 '≈ Sản lượng cần bổ sung (pcs)': '≈ 追加必要数量(pcs)', 'giá trị gia công bị ảnh hưởng': '影響を受ける加工金額', 'Đơn giá so với lô 2.001–5.000': 'ロット2,001–5,000との単価比',
 '5. Chuẩn hoá dữ liệu kế hoạch': '5. 計画データの標準化', '4. Công suất; 5. Chi phí & Lô': '4. 生産能力; 5. 加工費とロット', 'Việc Giám đốc nhà máy cần làm': '工場長のアクション',
 'Ngày tham chiếu (cắt dữ liệu)': '基準日(データ締め)', 'Mức "Theo dõi" (trễ 1–7 ngày)': '「注意」(1–7日遅れ)', 'chi phí nhân công chưa có đơn': '受注のない人件費',
 'Cột BO/BP (END+38, chênh lệch)': 'BO/BP列 (END+38, 差)', 'Cột mẫu gửi/phản hồi trộn kiểu': 'サンプル送付/回答列の形式混在', 'Tỷ lệ sử dụng T2/2027 (có Tết)': '稼働率 2027年2月(テト含む)',
 'Tỷ lệ sử dụng ngày công (APIC)': '人日稼働率(APIC)', 'Không rõ ý nghĩa → không dùng.': '意味不明 → 使用せず。', 'Ngày, "●", "継続", ghi chú tự do': '日付、「●」、「継続」、自由記述',
 'Không kiểm soát được hạn giao.': '納期を管理できない。', 'Chi phí chia lô mùa 27SS (USD)': '27SSロット分割コスト(USD)', 'Lỗi dữ liệu cần sửa ở file gốc': '元ファイルで修正すべきデータ不備',
 'nhiều đơn cùng lúc trên 1 line': '1ラインで複数オーダー同時', 'Lịch may; ETD kế hoạch = END+10': '縫製日程; 計画ETD = END+10', 'sample sent / reply / giai đoạn': 'sample sent / reply / 段階',
 'Tăng/giảm vs cùng mùa năm trước': '前年同シーズン比', 'Trạng thái mẫu (PPS, SMS, 1st…)': 'サンプル状況 (PPS, SMS, 1st…)', 'Đơn giá gia công theo quy mô lô': 'ロット規模別の加工単価',
 'Tỷ trọng hàng 別注 (bespoke) 26FW': '26FWの別注比率', 'Lệch ETD vô nghĩa (thiếu 工場ETD)': 'ETD差が無意味(工場ETD未入力)', 'Tách cột ngày & cột trạng thái.': '日付列と状態列を分ける。',
 'Tách đơn mẫu khỏi đơn sản xuất.': 'サンプルと量産オーダーを分ける。', 'Chi phí gia công năm 2026 (USD)': '2026年加工費(USD)', 'Tỷ lệ sử dụng nhà máy T3–T6/2027': '工場稼働率 2027年3–6月',
 'Lập kế hoạch nhân lực T3–T6/2027': '2027年3–6月の要員計画', 'Giảm rủi ro tập trung khách hàng': '顧客集中リスクの低減', '4. Chia nhỏ lô làm tăng chi phí.': '4. ロット分割がコストを押し上げている。',
 'Sản lượng năm 2026 (26SS + 26FW)': '2026年数量 (26SS + 26FW)', 'Hiệu suất lô ≤300 pcs (平均日産/MAX)': '≤300ロットの効率(平均日産/MAX)', 'Đơn sắp SX trễ ETD ≥ 8 ngày (pcs)': 'ETD 8日以上遅れの予定数量(pcs)',
 'Chỉ 234 dòng (2024–25) có dữ liệu': 'データがあるのは234行(2024–25)のみ', 'Mức "Nghiêm trọng" (trễ ≥30 ngày)': '「重大」(30日以上遅れ)', 'Brand × mùa, tập trung khách hàng': 'ブランド × シーズン、顧客集中度',
 'Tỷ trọng TNF trong tổng sản lượng': '総数量に占めるTNF比率', '1. Quy mô đi ngang, mùa 27SS giảm.': '1. 規模は横ばい、27SSは減少。', 'Chuẩn hoá: 1st → SMS → PPS → BULK.': '統一: 1st → SMS → PPS → BULK。',
 'Phương pháp, giả định & ánh xạ cột': '前提・定義・列対応', 'Không đo được lead time duyệt mẫu.': 'サンプル承認リードタイムを測れない。', 'Bỏ hoặc kéo công thức cho toàn bộ.': '削除するか全行に数式を展開する。',
 'Tỷ lệ sử dụng bình quân T3–T6/2027': '平均稼働率 2027年3–6月', 'Đơn có NVL về sau ngày START (pcs)': '資材着がSTARTより後の数量(pcs)', 'Khó tổng hợp; 別注 nên là cột riêng.': '集計しにくい; 別注は別列にすべき。',
 'Xưởng (chuẩn hoá ký tự full-width)': '工場(全角文字を統一)', 'Chi phí nhân công / công-ngày (USD)': '人件費 / 人日(USD)', 'Tỷ lệ sử dụng bình quân T1–T10/2026': '平均稼働率 2026年1–10月',
 'Đơn vướng NVL (trễ / chưa xác định)': '資材問題のオーダー(遅れ/未定)', 'Dùng để tính công-ngày cần bổ sung.': '追加必要人日の計算に使用。', 'Công suất xưởng APIC theo ngày công': 'APIC工場の生産能力(人日)',
 'Ngưỡng "Nghiêm trọng" (ngày trễ ETD)': '「重大」の閾値(ETD遅れ日数)', '2. Phụ thuộc lớn vào một khách hàng.': '2. 特定顧客への依存度が高い。', 'Sửa công thức S: trả "" khi U trống.': 'S列の数式を修正: Uが空なら""を返す。',
 'SAM – tổng thời gian chuẩn (giây/sp)': 'SAM – 標準時間合計(秒/枚)', '6. Rủi ro giao hàng; 7. Chất lượng DL': '6. 納期リスク; 7. データ品質', 'Toàn bộ đơn đang/sắp SX với cờ rủi ro': '生産中/予定の全オーダーとリスクフラグ',
 '5. Rủi ro giao hàng tập trung ở 27SS.': '5. 納期リスクは27SSに集中。', 'Tỷ lệ sử dụng nhà máy T11/2026–T1/2027': '工場稼働率 2026年11月–2027年1月', 'giờ chuẩn theo SAM / giờ công kế hoạch': 'SAM標準時間 / 計画工数',
 'Sản lượng sắp SX trễ ETD ≥ 8 ngày (pcs)': 'ETD 8日以上遅れの予定数量(pcs)', 'Chi phí thấp, giúp ra quyết định nhanh.': '低コストで迅速な意思決定に役立つ。', 'Do 工場ETD trống nhưng công thức vẫn trừ.': '工場ETDが空でも数式が計算しているため。',
 'Hiệu suất yêu cầu (giờ chuẩn / giờ công)': '必要効率(標準時間 / 工数)', 'Sales cập nhật 工場ETD trước khi xếp line.': 'ライン割付前に営業が工場ETDを入力する。', 'Năng suất 2026 (pcs / công-ngày), T1–T10': '2026年生産性(pcs/人日)、1–10月',
 'Tỷ lệ sử dụng ngày công T11/2026–T1/2027': '人日稼働率 2026年11月–2027年1月', 'Chi phí chia lô tích luỹ 23SS–27FW (USD)': 'ロット分割コスト累計 23SS–27FW(USD)', 'Tỷ lệ sử dụng bình quân T11/2026–T1/2027': '平均稼働率 2026年11月–2027年1月',
 'Rủi ro giao hàng – đơn đang/ sắp sản xuất': '納期リスク – 生産中/予定オーダー', 'Số style, lô trung vị, ngày công kế hoạch': '品番数、ロット中央値、計画人日', 'Số mã hàng bắt đầu bình quân / tháng 2026': '2026年 月平均着手品番数',
 'Giảm đơn giá GC lô nhỏ (đang gấp 2,7 lần)': '小ロット加工単価の削減(現在2.7倍)', 'Bắt buộc nhập Brand; dùng Data Validation.': 'ブランドを必須入力にし、入力規則を使う。', 'Tổng/trung bình trong file gốc có thể sai.': '元ファイルの合計/平均が誤る可能性。',
 'Đổi cột Factory trong file gốc thành APIC.': '元ファイルのFactory列をAPICに統一する。', 'Ngày công kế hoạch vs định biên theo tháng': '月別 計画人日 vs 定員', 'Mùa vụ (chuẩn hoá: bỏ dấu ", khoảng trắng)': 'シーズン(統一: " と空白を除去)',
 'KPI tổng hợp, phát hiện chính, khuyến nghị': '総合KPI、主な発見、推奨アクション', 'Số đơn chạy song song / line / ngày (2026)': '並行オーダー数 / ライン / 日 (2026)', 'Nhịp làm việc đề xuất cho Quản lý sản xuất': '生産管理の推奨業務サイクル',
 'Lịch line theo tuần (16 tuần từ 16/11/2026)': '週別ライン計画 (2026/11/16から16週)', 'Danh sách thả xuống cố định cho cột Status.': 'Status列をドロップダウン固定にする。',
 'Đơn giá GC lô ≤300 pcs so với lô 2.001–5.000': '≤300ロットの加工単価(2,001–5,000比)', 'Sản lượng, giờ chuẩn, chi phí theo mùa & năm': 'シーズン・年別の数量、標準時間、費用',
 'Tải theo line, nhân lực, hiệu suất line, Tết': 'ライン別負荷、要員、ライン効率、テト', 'MỤC LỤC – BÁO CÁO ĐIỀU HÀNH KẾ HOẠCH SẢN XUẤT': '目次 – 生産計画 経営レポート',
 'Dữ liệu đã làm sạch (nguồn cho mọi công thức)': 'クレンジング済みデータ(全数式の元)', 'Pivot/lọc trong file gốc bị tách đôi một mùa.': '元ファイルのピボット/フィルタで1シーズンが2つに分かれる。',
 'Chi phí nhân công chưa có đơn T3–T6/2027 (USD)': '受注のない人件費 2027年3–6月(USD)', 'SL cần bổ sung T3–T6/2027 để đạt tải 90% (pcs)': '稼働率90%に必要な追加数量 2027年3–6月(pcs)',
 'Không theo dõi được kho nhận hàng cho mùa mới.': '新シーズンの倉庫着を追跡できない。', '1. Tóm tắt điều hành; 2. Mùa vụ; 3. Khách hàng': '1. エグゼクティブサマリー; 2. シーズン別; 3. 顧客別',
 'Tỷ lệ sử dụng ngày công APIC T11/2026 – T1/2027': 'APIC人日稼働率 2026年11月 – 2027年1月', 'Chỉ có dữ liệu cho 234 dòng 2024–25 – không dùng': '2024–25の234行のみデータあり – 使用せず',
 'Mục tiêu: chi phí chia lô ≤10% chi phí gia công.': '目標: ロット分割コスト ≤ 加工費の10%。', '≈ Chi phí nhân công chưa có đơn để phân bổ (USD)': '≈ 受注で配賦できない人件費(USD)',
 'Làm lệch trung bình nếu không gia quyền theo SL.': '数量加重しないと平均が歪む。', 'Lịch line 16 tuần, checklist sẵn sàng 6 tuần tới': '16週ライン計画、6週間の準備チェックリスト',
 'Ngày NVL về HCM / ngày có thể đưa vào SX (=W+21)': '資材HCM着日 / 投入可能日 (=W+21)', 'BÁO CÁO ĐIỀU HÀNH – KẾ HOẠCH SẢN XUẤT (PLANTABLE)': '経営レポート – 生産計画 (PLANTABLE)',
 'Rủi ro: một khách hàng chiếm 64% (80% riêng 27SS).': 'リスク: 1顧客が64%(27SSでは80%)。', '6. Dữ liệu chưa đủ tin cậy để ra quyết định nhanh.': '6. データの信頼性が迅速な意思決定には不十分。',
 'Xoá hoặc đánh dấu huỷ các dòng không còn sản xuất.': '生産しない行は削除またはキャンセル表示する。', 'Số đơn chạy song song bình quân / line / ngày (2026)': '平均並行オーダー数 / ライン / 日 (2026)',
 'So sánh chi phí tăng ca với rủi ro phạt hoặc mất đơn.': '残業コストとペナルティ・失注リスクを比較。', 'Lệch ETD vô nghĩa (thiếu 工場ETD); Đơn giá GC bất thường': 'ETD差が無意味(工場ETD未入力); 加工単価の異常値',
 'Giả định có thể chỉnh (ô xanh), ánh xạ cột, định nghĩa': '変更可能な前提(青字セル)、列対応、定義', 'Hiệu suất yêu cầu bình quân 2026 (giờ chuẩn / giờ công)': '2026年 平均必要効率(標準時間 / 工数)',
 'Chất lượng dữ liệu – các điểm cần làm sạch tại file gốc': 'データ品質 – 元ファイルで修正すべき点', 'Bằng QTY × đơn giá gia công (USD/pcs) của giá vốn 1 / 2': 'QTY × 原価1 / 2の加工単価(USD/pcs)',
 'TRANG GIÁM ĐỐC NHÀ MÁY – CÔNG SUẤT, NHÂN LỰC, HIỆU SUẤT': '工場長ページ – 生産能力・要員・効率', 'Công-ngày cần bổ sung T3–T6/2027 để đạt mức tải mục tiêu': '目標稼働率に必要な追加人日 2027年3–6月',
 'TRANG CEO – BỨC TRANH KINH DOANH & QUYẾT ĐỊNH CẦN ĐƯA RA': 'CEOページ – 事業の全体像と意思決定事項', 'Brand (gộp "TNF\\n別注" → TNF + cờ 別注; "HH" → HELLY HANSEN)': 'ブランド(「TNF\\n別注」→ TNF + 別注フラグ; 「HH」→ HELLY HANSEN)',
 'Bản đồ tải theo line (tỷ lệ sử dụng ngày công theo tháng)': 'ライン別負荷マップ(月別人日稼働率)', 'Dễ hiểu nhầm là 2 xưởng khác nhau khi tổng hợp công suất.': '能力集計時に別工場と誤解されやすい。',
 'TRANG QUẢN LÝ SẢN XUẤT – LỊCH LINE & ĐỘ SẴN SÀNG ĐƠN HÀNG': '生産管理ページ – ライン計画とオーダー準備状況', 'Mùa cũ nhưng lịch SX muộn; Lệch ETD vô nghĩa (thiếu 工場ETD)': '旧シーズンだが生産日程が遅い; ETD差が無意味(工場ETD未入力)',
 'Factory "AP２", giai đoạn mẫu "ＰＰＳ", "F23 " có khoảng trắng': 'Factory「AP２」、サンプル段階「ＰＰＳ」、「F23 」に空白', 'Bức tranh kinh doanh, khách hàng, quyết định cần phê duyệt': '事業の全体像、顧客、承認が必要な事項',
 'Danh sách đơn bắt đầu trong 6 tuần tới – kiểm tra sẵn sàng': '今後6週間に着手するオーダー – 準備状況チェック', 'Hiệu quả theo cụm line (A+B) – các đơn bắt đầu T1–T10/2026': 'ライン番号別(A+B)の効率 – 2026年1–10月着手分',
 'Giá vốn gia công theo lô sản xuất thực tế (gộp theo 集約CODE)': '実際の生産ロット(集約CODE単位)による加工原価', '3. Xưởng APIC kín lịch đến T1/2027, sau đó là khoảng trống lớn.': '3. APICは2027年1月まで満稼働、その後大きな空き。',
 'END ≥ ngày tham chiếu, loại các dòng "mùa cũ nhưng lịch SX muộn"': 'END ≥ 基準日、「旧シーズンだが生産日程が遅い」行を除外', 'Tuần này/6 tuần tới chạy gì, đơn nào chưa sẵn sàng, đơn nào trễ?': '今週/今後6週間に何を流すか、どれが未準備か、どれが遅れか?',
 'Đã gộp thành APIC trong báo cáo (cột "Mã xưởng cũ" ở sheet Data)': 'レポートではAPICに統合(データシートの「旧工場コード」列)', 'Giá vốn gia công theo điều kiện đặt hàng (lô = tổng SL của style)': '発注条件による加工原価(ロット = 品番の総数量)',
 'Kinh doanh đang đi về đâu? Rủi ro tài chính nào? Cần quyết định gì?': '事業はどこへ向かっているか? 財務リスクは? 何を決めるべきか?',
 'Lấy từ file gốc AT5: 工賃単価 × 総人員 / 縫直 / 25 ngày = (760×1442)/947/25.': '元ファイルAT5より: 工賃単価 × 総人員 / 縫直 / 25日 = (760×1442)/947/25。',
 '7 phát hiện chính, 6 khuyến nghị; định nghĩa chỉ số và dữ liệu sạch.': '主な発見7点、推奨アクション6点; 指標の定義とクレンジング済みデータ。',
 'File gốc lưu ngày 2026-11-13; phân tích lấy 2026-11-14 làm "hôm nay".': '元ファイル保存日 2026-11-13; 分析では2026-11-14を「本日」とする。',
 'Bằng 平均日産 / 日産MAX (giá vốn 2) – lô càng nhỏ càng không đạt tốc độ tối đa': '平均日産 / 日産MAX (原価2) – ロットが小さいほど最大速度に届かない',
 'Nhà máy đủ hay thiếu đơn? Line nào trống, khi nào? Nhân lực, Tết, hiệu suất?': '受注は足りているか? どのラインがいつ空くか? 要員、テト、効率は?',
 'Ánh xạ cột (sheet gốc "PLANTABLE (原価) 2026年", header dòng 9, dữ liệu từ dòng 10)': '列対応(元シート「PLANTABLE (原価) 2026年」、見出し9行目、データ10行目から)',
 'Không phân bổ được doanh thu/chi phí theo khách hàng (gồm đơn HH26-103, mẫu SMS).': '顧客別に売上/費用を配分できない(HH26-103、SMSサンプル等)。',
 'Nguồn: sheet Data (SUMIFS/COUNTIFS). Số style & lô trung vị tính sẵn từ dữ liệu sạch.': '出典: データシート(SUMIFS/COUNTIFS)。品番数とロット中央値はクレンジング済みデータから算出。',
 'Gửi Sales danh sách đơn ETD trễ ≥8 ngày và đơn thiếu 工場ETD để thống nhất ngày giao mới.': 'ETD 8日以上遅れと工場ETD未入力のリストを営業に送り、新納期を合意する。',
 'Ảnh hưởng trực tiếp đến độ tin cậy của báo cáo. Đề xuất giao người phụ trách chuẩn hoá.': 'レポートの信頼性に直結。標準化の担当者を決めることを提案。',
 'Ô chữ xanh dương = giá trị đầu vào có thể chỉnh; mọi bảng khác tự tính lại theo các ô này.': '青字セル = 変更可能な入力値。他の表はこれらのセルに基づき自動再計算。',
 'Cột "Line trống từ ngày" tô đỏ = line hết đơn trước 15/03/2027, cần giao đơn 27FW sớm nhất.': '「ライン空き開始日」が赤 = 2027/03/15より前に受注が切れるライン。27FWを最優先で割付。',
 'Phương án dự phòng: tạm điều chuyển nhân lực, cho nghỉ phép bù, hoặc nhận gia công ngắn hạn.': '代替案: 一時的な配置転換、代休消化、短期の受託加工。',
 'Sản lượng (pcs) theo brand × mùa. "Khác" = phần còn lại. Brand trống = 108k pcs (xem sheet 7).': 'ブランド × シーズン別数量(pcs)。「その他」= 残り。ブランド空欄 = 108k pcs(シート7参照)。',
 'START = END dòng trước + 1 nên dòng thừa cuối khối line bị đẩy vào tương lai → làm sai tải line.': 'START = 前行END + 1のため、ライン末尾の不要行が将来に押し出され、ライン負荷を歪める。',
 'Đọc trong 3 phút. Số liệu liên kết trực tiếp với các sheet phân tích; ngày tham chiếu 14/11/2026.': '3分で読める内容。数値は分析シートと直接リンク; 基準日 2026/11/14。',
 'Họp line: rà lịch tuần trên Gantt và xử lý các ô đỏ/cam. Cập nhật START/END thực tế vào PLANTABLE.': 'ライン会議: ガントで週計画を確認し、赤/橙のセルを処理。実績START/ENDをPLANTABLEに反映。',
 'Mục tiêu brand ngoài TNF đạt ≥ 30% sản lượng. Tìm hiểu lý do MONTBELL giảm và có kế hoạch giữ khách.': 'TNF以外のブランドを数量の30%以上に。MONTBELL減少の理由を調べ、顧客維持策を立てる。',
 'Bằng CP2 − CP1: phần chi phí tăng thêm vì lô SX thực tế nhỏ hơn lô đặt hàng (kém hiệu suất học việc)': '加工費2 − 加工費1: 実際の生産ロットが発注ロットより小さいことによる追加コスト(習熟効率の低下)',
 'Bằng Σ(ﾗｲﾝ数 × ngày làm việc T2–T7 từ START→END) / (định biên × ngày làm việc). Chưa trừ nghỉ lễ/Tết.': '受注あり人日 / (定員 × 稼働日)。祝日・テトは未控除。',
 'Theo dõi sản lượng ngày so với 予定日産数. Line đạt dưới 80% kế hoạch 2 ngày liên tiếp thì báo GĐ nhà máy.': '日産を予定日産数と比較。2日連続で計画の80%未満のラインは工場長に報告。',
 'Xưởng APIC: 30 line thường × 32 CN = 960 CN may (line nhỏ 小ロットライン = 16 CN). Ngày tham chiếu 14/11/2026.': 'APIC工場: 通常ライン30 × 32人 = 縫製960人(小ロットライン = 16人)。基準日 2026/11/14。',
 'Ghi chú: 27FW mới bắt đầu nhận đơn (chưa đầy đủ). "Cùng mùa năm trước" = so 27SS với 26SS, 26FW với 25FW…': '注: 27FWは受注開始直後(未確定)。「前年同シーズン」= 27SS対26SS、26FW対25FW…',
 'Dùng cho họp sản xuất hằng tuần. Lịch theo tuần (bắt đầu thứ Hai) và danh sách đơn bắt đầu trong 6 tuần tới.': '週次生産会議用。週別計画(月曜始まり)と今後6週間に着手するオーダー一覧。',
 'Lô ≤300 chỉ đạt ~34% tốc độ tối đa. Áp dụng mẫu chuẩn hoá, cử tổ trưởng kèm line và giảm thời gian học việc.': '≤300ロットは最大速度の約34%しか出ない。標準化サンプル、班長の張り付き、習熟期間短縮を実施。',
 'Đơn có END ≥ ngày tham chiếu. Lệch ETD = 工場ETD yêu cầu − ETD kế hoạch (âm = trễ). Ngưỡng ở sheet Phương pháp.': 'END ≥ 基準日のオーダー。ETD差 = 要求工場ETD − 計画ETD(マイナス = 遅れ)。閾値は前提・定義シート。',
 'Giao Kế hoạch SX chuẩn hoá file PLANTABLE (danh mục, công thức, xoá dòng treo) để báo cáo này cập nhật hàng tuần.': '生産計画がPLANTABLEを標準化(マスタ、数式、残存行削除)し、本レポートを毎週更新できるようにする。',
 'Lệch ETD = 工場ETD yêu cầu − ETD kế hoạch. Âm = kế hoạch trễ hơn yêu cầu. Giá trị |x|>365 bị loại (do 工場ETD trống).': 'ETD差 = 要求工場ETD − 計画ETD。マイナス = 計画が要求より遅い。|x|>365は除外(工場ETD空欄のため)。',
 'Tính bằng Python từ dữ liệu sạch (giá trị cố định); các bảng còn lại dùng công thức SUMIFS/COUNTIFS trên sheet Data': 'クレンジング済みデータからPythonで算出(固定値); 他の表はデータシート上のSUMIFS/COUNTIFS数式',
 'CUT = START − 10 ngày. NVL "OK" khi ngày có thể đưa vào SX ≤ START. Kết luận tự cập nhật khi sửa ngày hoặc trạng thái.': 'CUT = START − 10日。投入可能日 ≤ STARTなら資材「OK」。日付や状態を修正すると判定は自動更新。',
 'Nguồn: APIC P-Plan 2026 (2026-11-13), sheet PLANTABLE (原価) 2026年. Ngày tham chiếu 14/11/2026. Bấm vào tên sheet để mở.': '出典: APIC P-Plan 2026 (2026-11-13)、シート PLANTABLE (原価) 2026年。基準日 2026/11/14。シート名をクリックで移動。',
 'Chốt đơn 27FW sớm, ưu tiên HELLY HANSEN, GOLDWIN và brand mới. Cân nhắc đưa đơn 27SS đang trễ sang T3 để giảm tải T1–T2.': '27FWを早期確定。HELLY HANSEN、GOLDWIN、新規ブランドを優先。遅れている27SSを3月に回し1–2月の負荷を下げることも検討。',
 'Dùng danh mục thả xuống cho mùa, brand, factory. Tách cột 別注. Sửa công thức cột S. Xoá các dòng treo. Bắt buộc nhập 工場ETD.': 'シーズン・ブランド・工場はドロップダウン。別注列を分離。S列の数式修正。残存行を削除。工場ETDを必須入力に。',
 'Đơn giá GC2 = giá vốn gia công theo lô SX thực tế. Lô càng nhỏ → chưa qua giai đoạn học việc → năng suất thấp, đơn giá cao.': '加工単価2 = 実際の生産ロットによる加工原価。ロットが小さい → 習熟前に終わる → 生産性が低く単価が高い。',
 'Đặt mục tiêu brand ngoài TNF ≥30% sản lượng năm 2027. Giao Sales tìm hiểu lý do MONTBELL giảm về 0 và lập kế hoạch giữ khách.': '2027年はTNF以外のブランドを数量の30%以上に。MONTBELLがゼロになった理由を営業が調査し、顧客維持策を立てる。',
 'Đặt MOQ hoặc phụ phí cho lô ≤300 pcs. Gộp lô cùng style hoặc cùng 集約CODE. Theo dõi % chi phí chia lô theo mùa (mục tiêu ≤10%).': '≤300ロットにMOQまたは割増料金を設定。同一品番・同一集約CODEはロットをまとめる。ロット分割コスト率をシーズン別に管理(目標 ≤10%)。',
 'Nhiều line chạy 2 đơn cùng lúc (song song AB, A với A1/A2). Theo dõi năng suất thực tế từng đơn và cân đối nhân lực giữa các đơn.': '多くのラインで2オーダーを同時に流している(AB並行、AとA1/A2)。オーダー別の実績生産性を追い、要員配分を調整。',
 'Nguồn: APIC P-Plan 2026 (bản 2026-11-13), sheet "PLANTABLE (原価) 2026年" – 3.494 dòng đơn, 23SS → 27FW. Ngày tham chiếu: 14/11/2026.': '出典: APIC P-Plan 2026 (2026-11-13版)、シート「PLANTABLE (原価) 2026年」– 3,494オーダー行、23SS → 27FW。基準日: 2026/11/14。',
 'Kiểm tra "sẵn sàng" cho các đơn có CUT trong 10 ngày tới (NVL, mẫu PPS, rập, phụ liệu). Đơn nào chưa sẵn sàng thì báo Mua hàng/QA.': '10日以内にCUTするオーダーの準備状況(資材、PPS、型紙、副資材)を確認。未準備は購買/QAへ連絡。',
 'Lưu ý: rủi ro giao hàng đánh giá theo 工場ETD trong file PLANTABLE (Lệch ETD = 工場ETD − ETD kế hoạch). Tỷ lệ sử dụng chưa trừ ngày lễ.': '注: 納期リスクはPLANTABLEの工場ETDで評価(ETD差 = 工場ETD − 計画ETD)。稼働率は祝日未控除。',
 'Áp dụng MOQ hoặc phụ phí cho lô ≤300 pcs từ 27FW. Lô này chiếm 25% số đơn nhưng chỉ 4% sản lượng, với đơn giá gia công gấp 2,7 lần.': '27FWから≤300ロットにMOQまたは割増料金を適用。オーダー数の25%だが数量は4%、加工単価は2.7倍。',
 'Họp xử lý theo danh sách ở sheet 6, ưu tiên mức "Nghiêm trọng". Các lựa chọn: tăng line hoặc tăng ca, tách lô, hoặc đàm phán lùi 工場ETD.': 'シート6のリストで対策会議、「重大」を優先。選択肢: ライン増・残業、ロット分割、工場ETDの延期交渉。',
 'Khoảng 2 lần đổi mã / line / tháng (mỗi line 32 CN). Gom lô cùng style hoặc cùng 集約CODE vào một line. Dành line nhỏ 小ロットライン (14A1, 14A2…) cho lô ≤300.': '1ライン(32人)あたり月約2回の品番切替。同一品番・集約CODEを同じラインに集約。≤300ロットは小ロットライン(14A1, 14A2…)に回す。',
 'Đơn giá GC cao thường do SAM cao (hàng phức tạp) hoặc lô nhỏ. Cụm 14 chạy toàn line nhỏ (14A1, 14A2, 14B1, 14B2); các cụm 3, 9, 13, 15 cũng thường chia line nhỏ.': '加工単価が高いのはSAMが高い(複雑な商品)か小ロットのため。14番は全て小ロットライン(14A1, 14A2, 14B1, 14B2); 3、9、13、15番もよく小ロットに分割。',
 'Cách cập nhật: dán dữ liệu PLANTABLE mới → chạy lại script phân tích (hoặc nhờ Claude) → các sheet tự tính lại theo sheet Data và các ô xanh ở sheet Phương pháp.': '更新方法: 新しいPLANTABLEデータを貼付 → 分析スクリプトを再実行(またはClaudeに依頼) → データシートと前提・定義の青字セルに基づき自動再計算。',
 'Lưu ý: ﾗｲﾝ数 ở mỗi dòng là số CN của cả line, nên không cộng dồn khi nhiều đơn chạy song song. Tháng 2/2027 có Tết Nguyên đán (06/02/2027) nên ngày làm việc thực tế ít hơn.': '注: 各行のﾗｲﾝ数はライン全体の人数のため、並行オーダーでは合算しない。2027年2月は旧正月(2027/02/06)があり実稼働日は少ない。',
 '30 line thường (1A, 1B … 15A, 15B) × 32 CN. Line nhỏ 小ロットライン (vd 14A1, 14A2) = 16 CN, 2 line nhỏ = 1 line thường. AP1A/AP1B/AP1C cũ đã gộp thành APIC. Tham số 縫直 trong file gốc = 947.': '通常ライン30本(1A, 1B … 15A, 15B) × 32人。小ロットライン(例 14A1, 14A2) = 16人、小ロット2本 = 通常1本。旧AP1A/AP1B/AP1CはAPICに統合。元ファイルの縫直 = 947。',
 'Line thường = 32 CN (sửa được ở cột B). Line nhỏ 小ロットライン (vd 14A1, 14A2) = 16 CN, được tính vào line mẹ (14A). Đơn chạy chung AB chia đều cho A và B. Tỷ lệ = công-ngày có đơn / (32 × ngày làm việc). Line nhỏ có đơn tính 16 CN.': '通常ライン = 32人(B列で変更可)。小ロットライン(例 14A1, 14A2) = 16人、親ライン(14A)に計上。ABオーダーはAとBに按分。稼働率 = 受注あり人日 / (32 × 稼働日)。',
 'Mỗi ô hiển thị style có SL lớn nhất đang chạy trong tuần trên line đó ([A1] = line nhỏ 小ロットライン; +n = thêm n đơn khác). Đỏ = có đơn trễ ETD ≥30 ngày hoặc NVL về sau START. Cam = trễ 8–29 ngày hoặc NVL 未定. Xanh nhạt = đang chạy bình thường. Xám = line trống.': '各セルはその週・そのラインで数量最大の品番([A1] = 小ロットライン; +n = 他n件)。赤 = ETD 30日以上遅れ or 資材着がSTART後。橙 = 8–29日遅れ or 資材未定。水色 = 通常稼働。灰 = 空きライン。',
 'TNF chiếm 64% tổng sản lượng và 80% mùa 27SS. Top 3 brand chiếm 84%. MONTBELL giảm mạnh: 106k pcs (24FW) → 77k (25FW) → 15k (26FW) → 0 ở 27SS. LL.BEAN đã dừng từ 25SS. Ngược lại HELLY HANSEN (+83% 27SS so với 26SS), GOLDWIN và ALLTERRAIN 81 đang tăng. Hàng 別注 tăng từ 0% lên 24% sản lượng (26FW).': 'TNFが総数量の64%、27SSでは80%。上位3ブランドで84%。MONTBELLは急減: 106k pcs (24FW) → 77k (25FW) → 15k (26FW) → 27SSは0。LL.BEANは25SSで終了。一方HELLY HANSEN(27SSは26SS比+83%)、GOLDWIN、ALLTERRAIN 81は増加。別注は0%から数量の24%へ(26FW)。',
 'Sản lượng năm 2024 đạt 811k pcs, 2025 còn 735k pcs (−9%), 2026 là 773k pcs (+5%). Mùa 27SS hiện có 336k pcs, thấp hơn 26SS 11%. Tổng giờ chuẩn mỗi năm giữ quanh 2,0–2,1 triệu giờ vì hàng FW phức tạp hơn nhiều (SAM ~190–200 phút/sp so với ~110–130 phút/sp của hàng SS). Vì vậy nên đo năng lực bằng giờ chuẩn thay vì số pcs.': '2024年は811k pcs、2025年は735k pcs(−9%)、2026年は773k pcs(+5%)。27SSは現時点336k pcsで26SS比−11%。FWはSSよりはるかに複雑(SAM 約190–200分/枚 vs SS 約110–130分/枚)なため、年間標準時間は約200–210万時間で横ばい。能力はpcsではなく標準時間で測るべき。',
 'Lô ≤300 pcs chiếm 25% số dòng đơn nhưng chỉ 4% sản lượng. Đơn giá gia công của lô này là 39,4 USD/pcs, gấp 2,7 lần lô 2.001–5.000 (14,8 USD/pcs), vì chưa qua đường cong học việc (chỉ đạt ~34% tốc độ tối đa). Phần "chi phí chia lô" (giá vốn theo lô sản xuất thực tế so với lô đặt hàng) cộng dồn khoảng 8,7 triệu USD qua 10 mùa. Tỷ lệ này đã giảm từ ~20% (25SS) xuống ~11% (27SS), là hướng tốt cần duy trì. Riêng 26FW tỷ lệ lô nhỏ lại tăng lên 33% số dòng.': '≤300ロットはオーダー行の25%だが数量は4%。加工単価は39.4 USD/pcsで、2,001–5,000ロット(14.8 USD/pcs)の2.7倍。習熟曲線を越えられない(最大速度の約34%)ため。「ロット分割コスト」(実生産ロット基準と発注ロット基準の原価差)は10シーズン累計で約870万USD。比率は約20%(25SS)から約11%(27SS)へ改善しており維持すべき。ただし26FWは小ロット行が33%に増加。',
 # --- non-diacritic labels ---
 'Cao': '高', 'ETD KH': '計画ETD', 'NVL': '資材', 'NVL; ': '資材; ', 'ETD; ': 'ETD; ', 'CP GC1 (USD)': '加工費1 (USD)', 'CP GC2 (USD)': '加工費2 (USD)',
 'SAM (s/sp)': 'SAM (秒/枚)', 'Sales / Merchandising': '営業 / MD', 'Top 3 brand': '上位3ブランド', 'Line': 'ライン', 'Brand': 'ブランド',
 'Style': '品番', 'Item': 'アイテム', 'Part': 'パート', 'Contract': '契約#', 'n/a': '－', '2025 (SS+FW)': '2025年 (SS+FW)', '2026 (SS+FW)': '2026年 (SS+FW)',
 '3. 501–1.000': '3. 501–1,000', '4. 1.001–2.000': '4. 1,001–2,000', '5. 2.001–5.000': '5. 2,001–5,000', '6. >5.000': '6. >5,000',
 'Data': 'データ', 'Phương pháp': '前提・定義',
 # --- new strings ---
 'Mã mùa không khớp lịch SX': 'シーズンコードと生産日程の不一致', 'Kiểm tra mã mùa (lịch SX muộn)': 'シーズンコード要確認(生産日程が遅い)',
 'Xác nhận mã mùa đúng với Sales.': '営業とシーズンコードを確認する。', 'Có thể là đơn bổ sung hoặc ghi sai mùa; vẫn được tính trong báo cáo.': '追加オーダーか表記ミスの可能性; レポートでは集計に含む。',
 # --- charts ---
 'Sản lượng theo mùa (pcs)': 'シーズン別数量 (pcs)', 'Chi phí chia lô theo mùa (USD)': 'シーズン別ロット分割コスト (USD)',
 'Sản lượng tích luỹ 23SS–27FW theo brand (pcs)': 'ブランド別累計数量 23SS–27FW (pcs)', 'APIC – tỷ lệ sử dụng ngày công theo tháng': 'APIC – 月別人日稼働率',
 'Đơn giá gia công theo quy mô lô (USD/pcs)': 'ロット規模別 加工単価 (USD/pcs)', 'Sản lượng theo năm (pcs)': '年別数量 (pcs)',
 # --- status chips (prefix symbol handled separately) ---
 'Giảm': '減少', 'Ổn định': '安定', 'Kín lịch': 'フル稼働', 'Cải thiện': '改善', 'Khối lượng': '物量', 'Rủi ro cao': '高リスク',
 'Lỗ hổng lớn': '大きな空き', 'Đổi mã nhiều': '切替が多い', 'Phụ thuộc cao': '依存度が高い', 'Tập trung cao': '集中度が高い',
 'Giảm – cần bù': '減少 – 補填が必要', 'Rủi ro uy tín': '信用リスク', 'Mất năng suất': '生産性ロス', 'Đang cải thiện': '改善中',
 'Ổn định, tăng nhẹ': '安定・微増', 'Kín đến T1/2027': '2027年1月まで満稼働', 'Lập kế hoạch Tết': 'テト計画を作成', 'Cân đối lại line': 'ライン再配分',
 'Đôn đốc mua hàng': '購買を督促', 'Đôn đốc QA / khách': 'QA・顧客を督促', 'Bổ sung dữ liệu': 'データ補完', 'Mục tiêu ≥90%': '目標 ≥90%',
 'Cần xử lý NVL': '資材対応が必要', 'Thiếu đơn ': '受注不足', 'Rủi ro tài chính': '財務リスク', 'Theo dõi ': '注意',
 'Lô nhỏ đắt gấp ~2,7 lần': '小ロットは約2.7倍の単価',
}


EX.update({
 'A. CHỈ SỐ CHÍNH': 'A. 主要指標', 'Xu hướng 4 năm': '4年間の推移', 'Giờ chuẩn APIC': 'APIC標準時間(h)', '% dòng lô ≤300': '≤300ロット行の比率',
 'Công-ngày có đơn: mỗi ngày làm việc (T2–T7), line thường (vd 1A) có đơn tính 32 CN, mỗi line nhỏ (vd 14A1) có đơn tính 16 CN; đơn AB chiếm cả 1A và 1B. Không cộng trùng khi nhiều đơn chạy song song trên cùng line. Công suất = định biên APIC (30 line × 32 CN) × ngày làm việc. AP1A/AP1B/AP1C cũ được gộp thành APIC. Đã loại các dòng "mùa cũ" còn treo lịch.':
 '受注あり人日: 各稼働日(月–土)に、受注のある通常ライン(例 1A)は32人、小ロットライン(例 14A1)は16人として計上。ABオーダーは1Aと1Bの両方を占有。同一ラインの並行オーダーは重複計上しない。能力 = APIC定員(30ライン × 32人) × 稼働日。旧AP1A/AP1B/AP1CはAPICに統合。残存する「旧シーズン」行は除外。',
})

EX.update({
 'Nhóm': '区分', 'Dòng chấm': '網掛け行', 'Quy mô lô': 'ロット規模', 'Mùa trước': '前シーズン', 'SL cùng kỳ': '前年同期 数量', 'SL mùa trước': '前シーズン 数量',
 'SL mùa phân tích': '分析シーズン 数量', '% SL hàng 別注': '別注の数量比', '% SL lô ≤300': '≤300ロットの数量比', '% SL TNF': 'TNFの数量比',
 'Δ vs cùng kỳ': 'Δ 前年同期比', 'Δ vs mùa trước': 'Δ 前シーズン比', 'SAM (phút/sp)': 'SAM (分/枚)', 'Mùa phân tích': '分析シーズン',
 'Cùng kỳ năm trước': '前年同期', 'Số ngày may TB': '平均縫製日数', 'Số ngày may': '縫製日数', 'Nhóm ngày may': '縫製日数区分',
 'Nhóm SAM (thời gian chế tạo)': 'SAM区分(製造時間)', '% SL (cùng kỳ)': '数量比(前年同期)', '% SL (phân tích)': '数量比(分析)', '% SL (mùa trước)': '数量比(前シーズン)',
 'Δ SL vs cùng kỳ': 'Δ 数量 前年同期比', 'Δ SL vs mùa trước': 'Δ 数量 前シーズン比', 'Lô TB – cùng kỳ': '平均ロット – 前年同期', 'Lô TB – phân tích': '平均ロット – 分析',
 'Số đơn (cùng kỳ)': 'オーダー数(前年同期)', 'Số đơn (phân tích)': 'オーダー数(分析)', 'Đơn giá GC – cùng kỳ': '加工単価 – 前年同期', 'Đơn giá GC – phân tích': '加工単価 – 分析',
 'Số ngày may TB / đơn': '平均縫製日数 / オーダー', 'Gõ mã mùa: 23SS … 27FW': 'シーズンコードを入力: 23SS … 27FW', 'Lô TB (pcs / dòng đơn)': '平均ロット(pcs / オーダー行)',
 'SAM TB (phút) – cùng kỳ': '平均SAM(分) – 前年同期', 'SAM TB (phút) – phân tích': '平均SAM(分) – 分析', 'Tỷ trọng (mùa phân tích)': '構成比(分析シーズン)',
 'B. Theo nhãn hàng (brand)': 'B. ブランド別', 'A. So sánh tổng quan 3 mùa': 'A. 3シーズン総合比較', 'C. Theo quy mô lô sản xuất': 'C. 生産ロット規模別',
 'SAM TB – thời gian chế tạo (phút/sp)': '平均SAM – 製造時間(分/枚)', 'Lô = lô sản xuất thực tế (集約, giá vốn 2).': 'ロット = 実際の生産ロット(集約、原価2)。',
 'F. Xu hướng tất cả các mùa (tỷ trọng sản lượng)': 'F. 全シーズンの推移(数量構成比)', '27FW mới bắt đầu nhận đơn nên chưa so sánh được.': '27FWは受注開始直後のため比較不可。',
 'D. Theo thời gian chế tạo sản phẩm (SAM, phút/sp)': 'D. 製品の製造時間別(SAM、分/枚)', 'E. Theo thời gian may của đơn (START → END, ngày)': 'E. オーダーの縫製期間別(START → END、日)',
 'Thống kê & so sánh theo mùa, nhãn hàng, quy mô lô, thời gian chế tạo': 'シーズン・ブランド・ロット規模・製造時間別の統計と比較',
 'Dòng chấm (chuyền lớn ghi 2 lần)': '網掛け行(大ラインの2重記載)',
 'SAM = tổng thời gian chuẩn may 1 sản phẩm (cột TIME). Hàng phức tạp (jacket kỹ thuật) thường trên 180 phút.': 'SAM = 製品1枚の標準縫製時間合計(TIME列)。複雑な商品(機能ジャケット)は通常180分超。',
 'Δ của các dòng % là chênh lệch điểm phần trăm. Đỏ = giảm >10%, xanh = tăng >10% (với chi phí/đơn giá, tăng là bất lợi).': '%行のΔはポイント差。赤 = 10%超の減少、緑 = 10%超の増加(費用・単価は増加が不利)。',
 'Mùa trước = mùa liền kề trước (27SS so với 26FW). Cùng kỳ = cùng mùa năm trước (27SS so với 26SS). Đã bỏ dòng chấm (chuyền lớn ghi 2 lần). Đổi 3 ô chữ xanh để so sánh mùa khác.': '前シーズン = 直前のシーズン(27SS対26FW)。前年同期 = 前年の同シーズン(27SS対26SS)。網掛け行(大ラインの2重記載)は除外。青字の3セルを変えると他シーズンも比較可能。',
 'Ô QTY nền chấm = dòng thứ 2 của chuyền lớn (vd 4AB, 5AB). Không cộng SL, giờ chuẩn, chi phí; chỉ dùng để tính chiếm line. Lô 原価1/原価2 được tính lại sau khi loại dòng chấm (điều chỉnh theo cùng công thức của file gốc).': 'QTYセルが網掛け = 大ライン(例 4AB, 5AB)の2行目。数量・標準時間・費用には加算せず、ライン占有の計算のみに使用。原価1/原価2のロットは網掛け行を除いて再計算(元ファイルと同じ式で補正)。',
 'Nếu cộng sẽ gấp đôi SL. Công thức gốc SUMIF theo style (標準原価1) vẫn cộng cả dòng chấm nên lô bị phóng đại.': '加算すると数量が2倍になる。元ファイルの品番別SUMIF(標準原価1)は網掛け行も合算するためロットが過大。',
 '1. ≤5 ngày': '1. ≤5日', '2. 6–10 ngày': '2. 6–10日', '3. 11–20 ngày': '3. 11–20日', '4. 21–40 ngày': '4. 21–40日', '5. >40 ngày': '5. >40日',
 '1. <60 phút': '1. <60分', '2. 60–119 phút': '2. 60–119分', '3. 120–179 phút': '3. 120–179分', '4. 180–239 phút': '4. 180–239分', '5. ≥240 phút': '5. ≥240分',
 'Chỉ số': '指標', 'SL (pcs)': '数量(pcs)', '1. Tóm tắt điều hành; 2A. So sánh mùa; 3. Khách hàng': '1. エグゼクティブサマリー; 2A. シーズン比較; 3. 顧客別', '4. Công suất; 5. Chi phí & Lô; 2A. So sánh mùa': '4. 生産能力; 5. 加工費とロット; 2A. シーズン比較', 'So với mùa trước & cùng kỳ: tổng quan, brand, quy mô lô, SAM, ngày may': '前シーズン・前年同期との比較: 総合、ブランド、ロット規模、SAM、縫製日数', 'Brand (top 12 của 25SS–27FW)': 'ブランド(25SS–27FW上位12)',
})


EX.update({
 'Hàng SMS': 'SMS品', 'Số dòng SMS': 'SMS行数', 'Brand chính': '主なブランド', 'Mã SMS / style': 'SMSコード / 品番', 'SL SMS (pcs)': 'SMS数量(pcs)',
 'Công-ngày dùng cho SMS': 'SMSに使用した人日', 'Nội dung (ghi chú gốc)': '内容(元コメント)',
 'Tổng hợp theo mùa vụ – hàng sản xuất': 'シーズン別集計 – 量産品',
 'Hàng SMS hỗ trợ chuyền mẫu (tách riêng, không tính trong các bảng trên)': 'サンプルライン支援のSMS品(別集計、上表には含まない)',
 'Hàng SMS hỗ trợ chuyền mẫu – đang/sắp chạy (tách riêng, không tính là hàng sản xuất)': 'サンプルライン支援のSMS品 – 生産中/予定(量産品とは別扱い)',
 'Nhận diện SMS: style/hợp đồng/item chứa "SMS", hoặc ghi chú có SMS / 展示会サンプル / 展示会投入 (trừ "SMSサンプル対応"), SL ≤500 pcs.': 'SMSの判定: 品番/契約/アイテムに「SMS」を含む、またはコメントにSMS / 展示会サンプル / 展示会投入を含む(「SMSサンプル対応」を除く)、数量 ≤500 pcs。',
 'Nguồn: sheet Data (SUMIFS/COUNTIFS). Chỉ hàng sản xuất: đã bỏ dòng chấm và tách hàng SMS (bảng riêng bên dưới). Số style & lô trung vị tính sẵn từ dữ liệu sạch.': '出典: データシート(SUMIFS/COUNTIFS)。量産品のみ: 網掛け行を除外し、SMS品は別集計(下表)。品番数とロット中央値はクレンジング済みデータから算出。',
 'SMS = mẫu chào hàng / mẫu triển lãm (展示会サンプル) chạy trên line nhỏ để hỗ trợ chuyền mẫu. Không cộng vào sản lượng, chi phí, rủi ro giao hàng của hàng sản xuất; vẫn tính chiếm line.': 'SMS = セールスマンサンプル / 展示会サンプル。小ロットラインでサンプルラインを支援。量産品の数量・費用・納期リスクには含めず、ライン占有のみ計上。',
 'Mùa trước = mùa liền kề trước (27SS so với 26FW). Cùng kỳ = cùng mùa năm trước (27SS so với 26SS). Chỉ hàng sản xuất: đã bỏ dòng chấm (chuyền lớn ghi 2 lần) và hàng SMS. Đổi 3 ô chữ xanh để so sánh mùa khác.': '前シーズン = 直前のシーズン(27SS対26FW)。前年同期 = 前年の同シーズン(27SS対26SS)。量産品のみ: 網掛け行(大ラインの2重記載)とSMS品を除外。青字の3セルを変えると他シーズンも比較可能。',
 'Mẫu chào hàng / mẫu triển lãm hỗ trợ chuyền mẫu (style/hợp đồng/item chứa SMS, hoặc ghi chú SMS / 展示会サンプル / 展示会投入; trừ "SMSサンプル対応"; SL ≤500). Tách khỏi hàng sản xuất: không tính vào sản lượng, chi phí, lô, rủi ro; vẫn tính chiếm line. Xem bảng riêng ở sheet 2 và 1C.': 'サンプルラインを支援するセールスマンサンプル / 展示会サンプル(品番/契約/アイテムにSMSを含む、またはコメントにSMS / 展示会サンプル / 展示会投入; 「SMSサンプル対応」を除く; 数量 ≤500)。量産品とは分け、数量・費用・ロット・リスクに含めずライン占有のみ計上。シート2と1Cに別表。',
 'Mỗi ô hiển thị style có SL lớn nhất đang chạy trong tuần trên line đó ([A1] = line nhỏ 小ロットライン; +n = thêm n đơn khác). Đỏ = có đơn trễ ETD ≥30 ngày hoặc NVL về sau START. Cam = trễ 8–29 ngày hoặc NVL 未定. Xanh nhạt = đang chạy bình thường. Tím = chỉ chạy hàng SMS (mẫu chào hàng, hỗ trợ chuyền mẫu). Xám = line trống.': '各セルはその週・そのラインで数量最大の品番([A1] = 小ロットライン; +n = 他n件)。赤 = ETD 30日以上遅れ or 資材着がSTART後。橙 = 8–29日遅れ or 資材未定。水色 = 通常稼働。紫 = SMS品(サンプルライン支援)のみ。灰 = 空きライン。',
})


EX.update({
 'Độ phủ đơn hàng theo tháng (line-ngày)': '月別 受注充足(ライン日)', 'Độ phủ đơn (gồm tạm tính) theo bản chụp tuần': '週次スナップショット別 受注充足率(暫定込み)', 'Ngày': '日付', 'Đã chốt': '確定', 'Bản chụp': 'スナップショット', '% đã chốt': '確定率', '% tạm tính': '暫定比率', 'Mùa / brand': 'シーズン / ブランド',
 'Đã chốt đến': '確定分の終了日', 'trên 30 line': '30ライン中', 'Đơn cuối cùng': '最後のオーダー', 'Tạm tính (暫定)': '暫定', 'Ngày làm việc': '稼働日',
 'Line trống từ': 'ライン空き開始日', '% gồm tạm tính': '暫定込み充足率', 'Cần chốt đơn': '受注確定が必要', 'Cần đơn 27FW': '27FW受注が必要', '4. Khuyến nghị.': '4. 推奨アクション。',
 'Line nhỏ (小ロット)': '小ロットライン', 'Trong đó tạm tính': 'うち暫定', 'Ưu tiên giao đơn': 'オーダー割付を優先', 'riêng phần đã chốt': '確定分のみ',
 'T3–T6/2027: đã chốt': '2027年3–6月: 確定', 'Công suất (line-ngày)': '能力(ライン日)', 'đơn chưa chốt (chữ đỏ)': '未確定オーダー(赤字)', 'Tạm tính đến (hết đơn)': '暫定分の終了日(受注切れ)',
 'Thiếu đơn, đang chững': '受注不足・伸び悩み', 'T3–T6/2027: gồm tạm tính': '2027年3–6月: 暫定込み', 'T12/2026–T2/2027: đã chốt': '2026年12月–2027年2月: 確定',
 'F. Nhận định & khuyến nghị': 'F. 所見と推奨アクション', '≈ pcs (theo năng suất 2026)': '≈ pcs(2026年生産性ベース)', 'Kín nhưng chủ yếu tạm tính': '満杯だが大半が暫定',
 'PLANTABLE (line-ngày có lịch)': 'PLANTABLE(計画あり ライン日)', 'A. Độ phủ đơn hàng theo tháng': 'A. 月別の受注充足率', 'Line-ngày còn trống T3–T8/2027': '空きライン日 2027年3–8月',
 'T12/2026–T2/2027: gồm tạm tính': '2026年12月–2027年2月: 暫定込み', 'Đang chạy (quanh ngày bản chụp)': '稼働中(スナップショット時点)', 'Độ phủ T3–T6/2027 (gồm tạm tính)': '充足率 2027年3–6月(暫定込み)',
 'Độ phủ T12/2026–T2/2027 (gồm tạm tính)': '充足率 2026年12月–2027年2月(暫定込み)', 'Tỷ trọng tạm tính trong T12/2026–T2/2027': '2026年12月–2027年2月の暫定比率',
 '2. T3–T6/2027 thiếu đơn và đang chững lại.': '2. 2027年3–6月は受注不足で伸び悩み。', '3. Hầu hết line hết đơn từ giữa T2 đến cuối T3/2027.': '3. ほとんどのラインが2027年2月中旬–3月末に受注切れ。',
 '1. 3 tháng tới kín lịch nhưng phần lớn còn tạm tính.': '1. 今後3か月は満杯だが大半が暫定。', 'Số line (32 CN) hết đơn trước 15/03/2027 (kể cả tạm tính)': '2027/03/15前に受注が切れるライン数(暫定込み)',
 'C. Từng line (32 CN): đang chạy, đã chốt đến, tạm tính đến': 'C. ライン別(32人): 稼働中、確定分の終了、暫定分の終了', 'B. Đơn theo mùa × brand – các tháng tới (line-ngày, gồm tạm tính)': 'B. シーズン × ブランド別 – 今後の月(ライン日、暫定込み)',
 'D. Xu hướng qua các bản chụp tuần (độ phủ gồm tạm tính / đã chốt)': 'D. 週次スナップショットの推移(暫定込み / 確定の充足率)', '4. Công suất; 4A. Nhận đơn 受注進行表; 2A. So sánh mùa': '4. 生産能力; 4A. 受注進行表分析; 2A. シーズン比較',
 'Độ phủ đơn (đã chốt / tạm tính) từ 受注進行表, xu hướng 13 tuần, so với PLANTABLE': '受注進行表による受注充足率(確定 / 暫定)、13週の推移、PLANTABLEとの比較',
 'PLANTABLE = công-ngày có lịch trong PLANTABLE (line thường 32 CN, line nhỏ 16 CN) quy về line-ngày. Cột cuối dương = đơn đã có trong 受注進行表 nhưng chưa xếp lịch chi tiết trong PLANTABLE.': 'PLANTABLE = PLANTABLEで計画済みの人日(通常ライン32人、小ロット16人)をライン日に換算。最終列がプラス = 受注進行表にはあるがPLANTABLEで詳細計画されていないオーダー。',
 'Nguồn: file ALL APIC 受注進行表 2026 (đọc các thanh mũi tên trong bảng Gantt). Chữ trắng = đơn đã chốt; chữ đỏ = tạm tính (暫定). Đơn vị: line-ngày = 1 line thường (32 CN) × 1 ngày làm việc (T2–T7); line nhỏ = 0,5. Công suất = 30 line.': '出典: ALL APIC 受注進行表 2026(ガント図の矢印を読み取り)。白字 = 確定; 赤字 = 暫定。単位: ライン日 = 通常ライン1本(32人) × 稼働日1日(月–土); 小ロットライン = 0.5。能力 = 30ライン。',
 '(a) Sales chốt các đơn S27 đang chữ đỏ trong tháng 11–12; (b) đặt mục tiêu đưa độ phủ T3–T6/2027 lên ≥60% trước 15/12/2026 và ≥90% trước 31/01/2027; (c) theo dõi bảng xu hướng D hằng tuần; nếu không tăng sau 2 tuần thì kích hoạt phương án nhân lực (sheet 1B).': '(a) 営業は赤字のS27オーダーを11–12月中に確定; (b) 2027年3–6月の充足率を2026/12/15までに60%以上、2027/01/31までに90%以上にする目標を設定; (c) 推移表Dを毎週確認し、2週間伸びなければ要員計画(シート1B)を発動。',
})


EX.update({
 'A. Độ phủ đơn hàng theo tháng': 'A. 月別の受注充足率', 'Số đơn bắt đầu': '着手オーダー数', 'SL bắt đầu (pcs)': '着手数量(pcs)',
 'Độ phủ đơn (gồm tạm tính) theo bản chụp': 'スナップショット別 受注充足率(暫定込み)', 'Ngày bản chụp': 'スナップショット日', 'Nguồn': '出典',
 'Đang chạy (tuần tham chiếu)': '稼働中(基準週)', 'Cần đơn mới': '新規受注が必要', 'Thiếu đơn': '受注不足', 'D. Từng line (32 CN): đang chạy, đã chốt đến, hết đơn': 'D. ライン別(32人): 稼働中、確定分の終了、受注切れ',
 'C. Đơn theo mùa × brand – các tháng tới (line-ngày, gồm tạm tính)': 'C. シーズン × ブランド別 – 今後の月(ライン日、暫定込み)',
 'F. Cơ cấu brand theo năm (tỷ trọng line-ngày, từ PLANTABLE)': 'F. 年別ブランド構成(ライン日比率、PLANTABLEより)', 'G. Nhận định & khuyến nghị': 'G. 所見と推奨アクション',
 '1. 3 tháng tới kín lịch nhưng phần lớn còn tạm tính.': '1. 今後3か月は満杯だが大半が暫定。', '2. Các tháng sau còn thiếu đơn.': '2. その後の月は受注不足。', '3. Line hết đơn.': '3. 受注切れのライン。',
 '4. Khuyến nghị.': '4. 推奨アクション。', '受注進行表 (file)': '受注進行表(ファイル)',
 'B. Bảng 受注進行表 (line × nửa tháng) – nhãn = mùa_brand chiếm nhiều nhất; chữ đỏ = tạm tính; * = line nhỏ 小ロット': 'B. 受注進行表(ライン × 半月) – ラベル = 最も多いシーズン_ブランド; 赤字 = 暫定; * = 小ロットライン',
 'Nền xanh dương = hàng SS, xanh lá = hàng FW, tím = SMS (mẫu chào hàng), trắng = trống. Ô gộp = cùng một nhóm đơn chạy liên tục.': '青 = SS品、緑 = FW品、紫 = SMS(セールスマンサンプル)、白 = 空き。結合セル = 同じオーダー群が連続稼働。',
 'Lịch sử lưu ở bao-cao/lich-su/do_phu_don_theo_ban_chup.csv. Các mốc trước 13/11/2026 lấy từ file 受注進行表; từ nay mỗi lần chạy báo cáo với PLANTABLE mới sẽ thêm 1 mốc (nguồn PLANTABLE).': '履歴は bao-cao/lich-su/do_phu_don_theo_ban_chup.csv に保存。2026/11/13以前は受注進行表ファイルから取得; 今後は新しいPLANTABLEでレポートを作成するたびに1点追加(出典 PLANTABLE)。',
 '(a) Sales chốt các đơn đang tạm tính trong 3 tháng tới; (b) mục tiêu độ phủ các tháng 4–7 ≥60% sau 4 tuần và ≥90% sau 8 tuần; (c) nếu độ phủ không tăng sau 2 bản chụp liên tiếp thì kích hoạt phương án nhân lực (sheet 1B).': '(a) 営業は今後3か月の暫定オーダーを確定; (b) 4–7か月目の充足率を4週間後に60%以上、8週間後に90%以上に; (c) 2回連続で充足率が伸びなければ要員計画(シート1B)を発動。',
 '4. Công suất; 4A. 受注進行表; 2A. So sánh mùa': '4. 生産能力; 4A. 受注進行表; 2A. シーズン比較',
 '受注進行表 tự dựng từ PLANTABLE: bảng line × nửa tháng, độ phủ đã chốt / tạm tính, xu hướng theo bản chụp': 'PLANTABLEから自動作成した受注進行表: ライン × 半月の表、確定 / 暫定の充足率、スナップショット推移',
})

EX.update(SHEETS)

def _re_rules(ctx):
    N, RK = ctx['N_'], ctx['RK']
    man = lambda x: f'{x/1e4:,.0f}万'
    K = ctx['K']; V1=lambda x: f'{x:.1f}'
    return [
        (r'^bản (\S+) \(4 tuần trước\)$', lambda m: f'{m.group(1)}版(4週間前)'),
        (r'^E\. Lịch sử cơ cấu brand theo năm \(tỷ trọng line-ngày, nguồn sheet (.*)\)$', lambda m: f'E. 年別ブランド構成の推移(ライン日比率、出典シート {m.group(1)})'),
        (r'^7\. 受注進行表 xác nhận', lambda m: '7. 受注進行表でも確認: 2027年2月までは満杯だが大半が暫定; 3–6月は空き。'),
        (r'^受注進行表 – tiến độ nhận đơn \(tự dựng từ PLANTABLE, ngày tham chiếu (\S+)\)', lambda m: f'受注進行表 – 受注状況(PLANTABLEから自動作成、基準日 {m.group(1)})'),
        (r'^Không cần file 受注進行表 riêng', lambda m: '受注進行表ファイルは不要。PLANTABLEの各オーダー = ライン上の1本のバー(START→END)。確定(推定) = 縫製開始済み、または最新の資材コメントで生地OK・資材OK; それ以外 = 暫定(赤字)。単位: ライン日 = 通常ライン1本(32人) × 稼働日1日; 小ロット = 0.5; 能力30ライン。'),
        (r'^Độ phủ (T\d+/\d+–T\d+/\d+) \(gồm tạm tính\)$', lambda m: f'充足率 {m.group(1)}(暫定込み)'),
        (r'^Line-ngày còn trống (T\d+/\d+–T\d+/\d+)$', lambda m: f'空きライン日 {m.group(1)}'),
        (r'^Số line \(32 CN\) hết đơn trước (\S+)$', lambda m: f'{m.group(1)}前に受注が切れるライン数(32人ライン)'),
        (r'^Tỷ trọng tạm tính trong (T\d+/\d+–T\d+/\d+)$', lambda m: f'{m.group(1)}の暫定比率'),
        (r'^bản chụp (\S+)$', lambda m: f'{m.group(1)}時点'),
        (r'^(T\d+/\d+–T\d+/\d+): (gồm tạm tính|đã chốt)$', lambda m: f'{m.group(1)}: '+('暫定込み' if m.group(2)=='gồm tạm tính' else '確定')),
        (r'^E\. Xu hướng qua các bản chụp \(độ phủ (\S+) và (\S+)\)$', lambda m: f'E. スナップショット別推移(充足率 {m.group(1)} と {m.group(2)})'),
        (r'^(T\d+/\d+–T\d+/\d+) đã kín khoảng', lambda m: f'{ctx["JN"]["w1"]}は能力の約{ctx["JN"]["t1"]:.0%}が埋まっているが、確定(稼働済み or 資材OK)は{ctx["JN"]["c1"]:.0%}のみ。残りは暫定で、資材手配と日程確保のため早期確定が必要。'),
        (r'^(T\d+/\d+–T\d+/\d+) kể cả tạm tính mới phủ', lambda m: f'{ctx["JN"]["w2"]}は暫定込みでも充足率{ctx["JN"]["t2"]:.0%}(ピーク{ctx["JN"]["pk"]:.0%}、{ctx["JN"]["pkd"]:%m/%d})。表Eを毎週確認。'),
        (r'^\d+/30 line hết cả đơn tạm tính trước', lambda m: f'{ctx["JN"]["nfree"]}/30ラインが暫定分も含め{ctx["JN"]["lim"]:%Y/%m/%d}前に受注切れ(表D、赤セル) – これらのラインに新規オーダーを優先割付。'),
        (r'^7\. 受注進行表 \(tự dựng từ PLANTABLE\)', lambda m: '7. 受注進行表(PLANTABLEから自動作成): 今後3か月は満杯だが大半が暫定; その後は空き。'),
        (r'^T\d+/\d+–T\d+/\d+ phủ \d+% công suất nhưng chỉ', lambda m: f'{K["j_w1"]}は能力の{K["j_t1"]:.0%}が埋まっているが確定は{K["j_c1"]:.0%}のみ(稼働済み or 資材OK)。{K["j_w2"]}は暫定込みでも{K["j_t2"]:.0%}(ピーク{K["j_pk"]:.0%}、{K["j_pkd"]:%m/%d})。詳細と受注進行表はシート4A。'),
        (r'^▲ Lô nhỏ đắt gấp', lambda m: f'▲ 小ロットは約{V1(K["lot_x"])}倍の単価'),
        (r'^Giảm đơn giá GC lô nhỏ', lambda m: f'小ロット加工単価の削減(現在{V1(K["lot_x"])}倍)'),
        (r'^Rủi ro: một khách hàng chiếm', lambda m: f'リスク: 1顧客が{K["tnf"]:.0%}(27SSでは{K["tnf27"]:.0%})。'),
        (r'^Áp dụng MOQ hoặc phụ phí cho lô ≤300 pcs từ 27FW', lambda m: f'27FWから≤300ロットにMOQまたは割増料金を適用。オーダー数の{K["lot_rows"]:.0%}だが数量は{K["lot_qty"]:.0%}、加工単価は{V1(K["lot_x"])}倍。'),
        (r'^Lô ≤300 chỉ đạt', lambda m: f'≤300ロットは最大速度の約{K["eff_s"]:.0%}しか出ない。標準化サンプル、班長の張り付き、習熟期間短縮を実施。'),
        (r'^Sản lượng \(pcs\) theo brand × mùa', lambda m: f'ブランド × シーズン別数量(pcs)。「その他」= 残り。ブランド空欄 = {K["blank"]:.0f}k pcs(シート7参照)。'),
        (r'^(\d+) dòng / ([\d,]+) pcs có ô QTY nền chấm', lambda m: f'{m.group(1)}行 / {m.group(2)} pcsのQTYセルが網掛け(主に大ラインABの2行目)'),
        (r'^Báo cáo: không cộng SL dòng chấm', lambda m: f'レポート: 網掛け行の数量は加算せず、ライン占有の計算のみに使用; {ctx["N_ADJ1"]}行の原価1ロットを再計算。元ファイル: 2行目のQTYを0にするかSUMIFから除外。'),
        (r'^Sản lượng năm 2024 đạt', lambda m: f'2024年は{K["y24"]/1000:.0f}k pcs、2025年は{K["y25"]/1000:.0f}k pcs({K["g25"]:+.0%})、2026年は{K["y26"]/1000:.0f}k pcs({K["g26"]:+.0%})。27SSは現時点{K["s27"]/1000:.0f}k pcs(26SS比{K["g27"]:+.0%})。FWはSSよりはるかに複雑(SAM 約{K["samFW_lo"]:.0f}–{K["samFW_hi"]:.0f}分/枚 vs SS 約{K["samSS_lo"]:.0f}–{K["samSS_hi"]:.0f}分/枚)なため、年間標準時間は約{V1(K["h_lo"])}–{V1(K["h_hi"])}百万時間。能力はpcsではなく標準時間で測るべき。大ラインの2重計上を避けるため網掛け{K["ndot"]}行({K["qdot"]:.0f}k pcs)を除外し、サンプルライン支援のSMS品{K["nsms"]}行({K["qsms"]:,.0f} pcs)は別集計。'),
        (r'^TNF chiếm', lambda m: f'TNFが総数量の{K["tnf"]:.0%}、27SSでは{K["tnf27"]:.0%}。上位3ブランドで{K["top3"]:.0%}。MONTBELLは急減: {K["mb"][0]:.0f}k pcs (24FW) → {K["mb"][1]:.0f}k (25FW) → {K["mb"][2]:.0f}k (26FW) → 27SSは{K["mb"][3]:.0f}k。LL.BEANは25SSで終了。27SSの前年同期(26SS)比: HELLY HANSEN {K["hh"]:+.0%}、ALLTERRAIN 81 {K["at"]:+.0%}、一方GOLDWINは{K["gw"]:+.0%}。別注は0%から数量の{K["bes26"]:.0%}へ(26FW)。'),
        (r'^Lô ≤300 pcs chiếm', lambda m: f'≤300ロットはオーダー行の{K["lot_rows"]:.0%}だが数量は{K["lot_qty"]:.0%}。加工単価は{V1(K["upc_s"])} USD/pcsで、2,001–5,000ロット({V1(K["upc_m"])} USD/pcs)の{V1(K["lot_x"])}倍。習熟曲線を越えられない(最大速度の約{K["eff_s"]:.0%})ため。「ロット分割コスト」(実生産ロット基準と発注ロット基準の原価差)は10シーズン累計で約{K["split"]*100:,.0f}万USD。比率は25SS {K["sp25"]:.0%}、26FW {K["sp26fw"]:.0%}、27SS {K["sp27"]:.0%}。ただし26FWは小ロット行が{K["small26fw"]:.0%}に増加。'),
        (r'^\d+k pcs không có brand', lambda m: f'{K["blank"]:.0f}k pcsにブランドなし。大ライン(AB)は2行記載され2行目は網掛け: {K["ndot"]}行 / {K["qdot"]:.0f}k pcs、2重計上の恐れ。シーズンコードの表記ゆれ。元ファイルに数式エラー139セル。23SSの9行が2027年2月まで残存しライン7の負荷を歪める(分析から除外)。工場ETD空欄時にETD差列が無意味な値になる。詳細はシート7。'),
        (r'^Thường \+ line nhỏ: (.*)$', lambda m: '通常＋小ロット: ' + m.group(1)),
        (r'^Line nhỏ: (.*)$', lambda m: '小ロット: ' + m.group(1)),
        (r'^(\d+) dòng / ([\d,]+) pcs$', lambda m: f'{m.group(1)}行 / {m.group(2)} pcs'),
        (r'^(\d+) cách ghi cho 10 mùa (.*)$', lambda m: f'10シーズンに{m.group(1)}通りの表記 ' + m.group(2).replace('vd', '例').replace(' và ', ' と ')),
        (r'^(\d+) giá trị Brand gốc \(vd (.*)\)$', lambda m: f'元のブランド値が{m.group(1)}種類(例 {m.group(2)})'),
        (r'^(\d+) ô lỗi \(VALUE, N/A…\)$', lambda m: f'エラーセル{m.group(1)}個 (VALUE, N/A…)'),
        (r'^(\d+) dòng 23SS ở line 7, lịch đến 02/2027 \(dòng gốc (.*)\); đã loại khỏi phân tích$', lambda m: f'23SSの{m.group(1)}行がライン7で2027年2月まで残存(元の行 {m.group(2)}); 分析から除外'),
        (r'^(\d+) dòng 26SS/26FW có START muộn \(vd 26FW chạy 02–03/2027; dòng gốc (.*)\)$', lambda m: f'26SS/26FWの{m.group(1)}行のSTARTが遅い(例 26FWが2027年2–3月に生産; 元の行 {m.group(2)})'),
        (r'^(\d+) dòng có giá trị ±40\.000\+ ngày$', lambda m: f'{m.group(1)}行が±40,000日超の値'),
        (r'^(\d+) đơn đang/sắp SX$', lambda m: f'生産中/予定の{m.group(1)}オーダー'),
        (r'^(\d+) dòng có số liệu$', lambda m: f'{m.group(1)}行にデータあり'),
        (r'^(\d+) mã khác nhau \((.*)\)$', lambda m: f'{m.group(1)}種類のコード ({m.group(2)})'),
        (r'^(\d+) dòng > 200 USD/pcs \(SL rất nhỏ / mẫu\)$', lambda m: f'{m.group(1)}行が200 USD/pcs超(極小数量 / サンプル)'),
        (r'^Xử lý (\d+) đơn trễ ETD ≥ 8 ngày$', lambda m: f'ETD 8日以上遅れの{m.group(1)}オーダーへの対応'),
        (r'^Bảo vệ ~(\d+)k pcs giao đúng hạn$', lambda m: f'約{m.group(1)}k pcsの納期を守る'),
        (r'^Tránh khoảng (\d+)k pcs', lambda m: f'約{m.group(1)}k pcs分の能力の空き(人件費 約{man(N["usd_gap"])}USD)を回避'),
        (r'^Có (\d+) đơn NVL về sau ngày START', lambda m: f'資材着がSTARTより後のオーダー{RK["nml"]}件、資材未定{RK["nmu"]}件、サンプル未承認{RK["nsm"]}件。START 4週間前に「生産準備」チェックリストを運用。'),
        (r'^Có (\d+) đơn \(', lambda m: f'生産中/予定のオーダーは{RK["n"]}件({RK["q"]:.0f}k pcs)。うち{RK["n8"]}件({RK["q8"]:.0f}k pcs、約{RK["p8"]:.0%})は計画ETDが工場ETDより8日以上遅く、{RK["n30"]}件は30日以上遅い。{RK["nml"]}件({RK["qml"]:.0f}k pcs)は資材着が縫製開始より後。{RK["nsm"]}件はサンプル未回答/未承認。{RK["nfe"]}件は工場ETD未入力。27SSの計画ETD遅れ率は{K["late27"]:.0%}で、24FW–26FWの{K["late_lo"]:.0%}–{K["late_hi"]:.0%}を上回る。'),
        (r'^Kế hoạch dùng khoảng', lambda m: f'2026年11月–2027年1月の計画は人日の約{N["u_now"]:.0%}(2026年平均{N["u26"]:.0%})。2027年2月はテトのため{N["u_feb"]:.0%}。3月は{N["u_mar"]:.0%}、4–6月はほぼ受注なし。3–6月に稼働率90%を確保するには約{N["pcs_gap"]/1000:,.0f}k pcsの追加が必要。埋められない場合、約{man(N["usd_gap"])}USDの人件費が配賦先を失う。27FWは現在{K["s27fw"]/1000:.0f}k pcsのみ(26FWは{K["s26fw"]/1000:.0f}k pcs)。'),
        (r'^Sản lượng đi ngang quanh', lambda m: f'数量は年{min(K["y24"],K["y25"],K["y26"])/1e6:.2f}–{max(K["y24"],K["y25"],K["y26"])/1e6:.2f}百万pcsで横ばい; 27SSは26SS比{K["g27"]:+.0%}。工場は2027年1月まで満稼働だが、3月以降はほぼ受注がない: 約{N["pcs_gap"]/1000:,.0f}k pcsの追加が必要で、埋められなければ約{man(N["usd_gap"])}USDの人件費が配賦先を失う。TNFが数量の{K["tnf"]:.0%}。小ロットは単価{V1(K["lot_x"])}倍。生産予定数量の約{RK["p8"]:.0%}が計画ETD 8日以上遅れ。'),
        (r'^Duyệt mục tiêu chốt khoảng', lambda m: f'2026/12/15までに27FW 約{N["pcs_gap"]/1000:,.0f}k pcsの確定を目標として承認。閑散月を埋めるオーダーには柔軟な価格(加工賃の値引き)を認め、伸びているブランド(HELLY HANSEN、ALLTERRAIN 81)を優先し、GOLDWINの数量回復を図る。'),
        (r'^Phê duyệt ngân sách tăng ca', lambda m: f'ETD 8日以上遅れのオーダー(約{ctx["q8k"]:.0f}k pcs、加工金額 約{man(ctx["risk_usd"])}USD)を納期内に戻すための残業・ライン増の予算を承認。並行して顧客と工場ETDの延期を交渉。'),
        (r'^Với tải hiện tại chỉ', lambda m: f'3月の負荷は{N["u_mar"]:.0%}、4月以降はほぼ0。シナリオを準備: (a) 27FW受注がある場合、(b) 不足する場合は有給消化、多能工訓練、ライン移動。11/30までにCEOへ報告。'),
        (r'^T11/2026–T1/2027 tải ~', lambda m: f'2026年11月–2027年1月の負荷は約{N["u_now"]:.0%}、2月はテトで稼働日が減る。休暇日程を確定し、テト前の残業を計画、テト前の工場ETDのオーダーを優先。'),
        (r'^108k pcs không có brand', lambda m: '108k pcsにブランドなし。シーズンコードの表記ゆれ。元ファイルに数式エラー139セル。23SSの9行が2027年2月まで残存しライン7の負荷を歪める(分析から除外)。工場ETD空欄時にETD差列が無意味な値になる。詳細はシート7。'),
    ]

def tr(s, ctx):
    if not isinstance(s, str): return s
    if s in EX: return EX[s]
    for sym in ('✔ ', '▲ ', '● '):
        if s.startswith(sym) and s[2:] in EX: return sym + EX[s[2:]]
    for pat, fn in _re_rules(ctx):
        m = re.match(pat, s, flags=re.S)
        if m: return fn(m)
    if ', ' in s and '(trống)' in s and len(s)<80:
        return ', '.join(EX.get(p,p) for p in s.split(', '))
    if '; ' in s:
        parts = s.split('; ')
        if all(p in EX for p in parts): return '; '.join(EX[p] for p in parts)
    return s

def _fix_formula(f, ctx):
    for vi, ja in sorted(SHEETS.items(), key=lambda x: -len(x[0])):
        f = f.replace(f"'{vi}'!", f"'{ja}'!")
    f = re.sub(r"(?<![\w'])Data!", "'データ'!", f)
    return re.sub(r'"([^"]*)"', lambda m: '"' + (tr(m.group(1), ctx) if m.group(1) else '') + '"', f)

def _title_text(t):
    try: return ''.join(r.t for p in t.tx.rich.p for r in (p.r or []))
    except Exception: return None

def translate_wb(wb, ctx, font_name='Meiryo UI'):
    from openpyxl.styles import Font
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str):
                    c.value = _fix_formula(v, ctx) if v.startswith('=') else tr(v, ctx)
                if c.hyperlink is not None and c.hyperlink.location is None and isinstance(c.hyperlink.target, str) and c.hyperlink.target.startswith('#'):
                    tgt = c.hyperlink.target
                    for vi, ja in SHEETS.items(): tgt = tgt.replace(f"'{vi}'!", f"'{ja}'!")
                    c.hyperlink = tgt
                if c.has_style and c.font is not None and c.font.name == 'Arial':
                    f = copy(c.font); f.name = font_name; c.font = f
        for rng in ws.conditional_formatting:
            for rule in rng.rules:
                if rule.formula: rule.formula = [_fix_formula(x, ctx) for x in rule.formula]
        for ch in ws._charts:
            for attr in ('title',):
                t = getattr(ch, attr)
                if t is not None:
                    txt = _title_text(t)
                    if txt: ch.title = tr(txt, ctx)
            for ax in (getattr(ch, 'x_axis', None), getattr(ch, 'y_axis', None)):
                if ax is not None and ax.title is not None:
                    txt = _title_text(ax.title)
                    if txt: ax.title = tr(txt, ctx)
            for s in ch.series:
                for ref in (getattr(s.val, 'numRef', None) if s.val else None,
                            getattr(s.cat, 'numRef', None) if s.cat else None,
                            getattr(s.cat, 'strRef', None) if s.cat else None,
                            getattr(s.tx, 'strRef', None) if s.tx else None):
                    if ref is not None and ref.f: ref.f = _fix_formula(ref.f, ctx)
            ch.roundedCorners = False
            if ch.title is not None: ch.title.overlay = False
            for ax in (getattr(ch, 'x_axis', None), getattr(ch, 'y_axis', None)):
                if ax is not None and ax.title is not None: ax.title.overlay = False
        if ws.auto_filter and ws.auto_filter.ref: pass
    for ws in wb.worksheets:
        if ws.title in SHEETS: ws.title = SHEETS[ws.title]
