# Hướng dẫn cho Claude

- Khi người dùng gửi file PLANTABLE mới (APIC P-Plan …xlsm) và muốn cập nhật báo cáo: chạy
  `python3 bao-cao/scripts/tao_bao_cao.py "<file>"` từ thư mục gốc repo, sau đó tính lại công thức (skill xlsx: `recalc.py`) cho cả 2 file kết quả, kiểm tra 0 lỗi, rồi gửi file cho người dùng.
- Người dùng KHÔNG gửi file 受注進行表 nữa: sheet `4A. 受注進行表` phải tự dựng từ PLANTABLE (đã có sẵn trong pipeline).
- Mọi chuỗi tiếng Việt mới thêm vào báo cáo phải có bản dịch trong `bao-cao/scripts/ja_tr.py`; sau khi build, quét file `_JP` để chắc không còn chữ tiếng Việt.
- Giữ quy ước: dòng chấm (QTY nền chấm) không cộng SL; hàng SMS tách riêng; line thường 32 CN, line nhỏ 16 CN; xưởng = APIC.
- Trả lời người dùng bằng tiếng Việt.
