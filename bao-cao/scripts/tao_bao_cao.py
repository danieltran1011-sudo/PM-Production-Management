# -*- coding: utf-8 -*-
"""Tạo báo cáo điều hành từ file PLANTABLE (không cần file 受注進行表).

Cách dùng (từ thư mục gốc repo):
    python3 bao-cao/scripts/tao_bao_cao.py "<file PLANTABLE .xlsm>" [--ngay YYYY-MM-DD] [--ngay-file YYYY-MM-DD]

--ngay       : ngày tham chiếu ("hôm nay" của báo cáo). Mặc định = ngày trong tên file + 1, nếu không có thì hôm nay.
--ngay-file  : ngày của file PLANTABLE (bản chụp 受注進行表). Mặc định = ngày trong tên file, nếu không có thì --ngay.

Kết quả (trong bao-cao/):
    Bao_cao_dieu_hanh_PLANTABLE_<ngay>.xlsx   (tiếng Việt)
    Keiei_Report_PLANTABLE_<ngay>_JP.xlsx     (tiếng Nhật)
và cập nhật bao-cao/lich-su/do_phu_don_theo_ban_chup.csv (thêm 1 mốc 受注進行表).
"""
import argparse, datetime as dt, os, re, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
BAO_CAO = os.path.dirname(HERE)

ap = argparse.ArgumentParser()
ap.add_argument('plantable')
ap.add_argument('--ngay')
ap.add_argument('--ngay-file')
a = ap.parse_args()

m = re.search(r'(20\d\d)[-_.]?(\d\d)[-_.]?(\d\d)', os.path.basename(a.plantable))
fdate = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None
ref = dt.date.fromisoformat(a.ngay) if a.ngay else ((fdate + dt.timedelta(days=1)) if fdate else dt.date.today())
pld = dt.date.fromisoformat(a.ngay_file) if a.ngay_file else (fdate or ref)

wd = tempfile.mkdtemp(prefix='plantable_')
print(f'Ngày tham chiếu: {ref}  |  Ngày file: {pld}  |  Thư mục tạm: {wd}')
subprocess.run([sys.executable, '-I', os.path.join(HERE, 'doc_plantable.py'), a.plantable, wd], check=True)
env = dict(os.environ, WORK_DIR=wd, REF_DATE=ref.isoformat(), PL_DATE=pld.isoformat(), SNAP_DATE=pld.isoformat(),
           OUT_FILE=os.path.join(BAO_CAO, f'Bao_cao_dieu_hanh_PLANTABLE_{ref.isoformat()}.xlsx'),
           HIST_CSV=os.path.join(BAO_CAO, 'lich-su', 'do_phu_don_theo_ban_chup.csv'))
subprocess.run([sys.executable, '-I', os.path.join(HERE, 'build_report.py')], check=True, env=env, cwd=wd)
print('Xong. Mở file bằng Excel (công thức tự tính lại khi mở).')
