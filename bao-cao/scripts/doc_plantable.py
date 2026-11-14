# -*- coding: utf-8 -*-
"""Bước 1: đọc file PLANTABLE (.xlsm/.xlsx) -> raw.pkl, clean.pkl, dotted.json trong thư mục làm việc.

Cách dùng: python3 doc_plantable.py <file PLANTABLE> <thư mục làm việc>
"""
import sys, os, json, unicodedata, datetime as dt
import numpy as np, pandas as pd, openpyxl
from openpyxl.utils import get_column_letter as L


def find_sheet(wb):
    for n in wb.sheetnames:
        if n.replace(' ', '').upper().startswith('PLANTABLE'):
            return n
    raise SystemExit('Không tìm thấy sheet PLANTABLE trong file')


def read_raw(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[find_sheet(wb)]
    rows = [(i,) + tuple(r) for i, r in enumerate(ws.iter_rows(min_row=10, values_only=True), start=10)]
    return pd.DataFrame(rows, columns=['row'] + [L(i + 1) for i in range(len(rows[0]) - 1)])


def read_dotted(path):
    """Ô QTY (cột Q) tô nền chấm = dòng thứ 2 của chuyền lớn (vd 4AB), không cộng số lượng."""
    wb = openpyxl.load_workbook(path)
    ws = wb[find_sheet(wb)]
    out = {}
    for r in range(10, ws.max_row + 1):
        c = ws.cell(r, 17)
        pt = c.fill.patternType
        if c.value is not None and pt not in (None, 'solid'):
            out[r] = pt
    return out


def clean(df):
    def s(x):
        if x is None or (isinstance(x, float) and np.isnan(x)): return None
        return unicodedata.normalize('NFKC', str(x)).replace('\n', ' ').replace('"', '').strip() or None
    def num(x):
        if isinstance(x, (int, float)) and not isinstance(x, bool) and not (isinstance(x, float) and np.isnan(x)): return float(x)
        return np.nan
    def date(x):
        if isinstance(x, (dt.datetime, pd.Timestamp)) and not pd.isna(x) and x.year >= 2015: return pd.Timestamp(x)
        return pd.NaT
    d = df[df.Q.notna() | df.I.notna()].copy()
    o = pd.DataFrame({'row': d.row})
    o['season'] = d.A.map(s); o['factory'] = d.B.map(s); o['line'] = d.C.map(num); o['part'] = d.D.map(s)
    o['workers'] = d.E.map(num); o['brand_raw'] = d.G.map(s); o['contract'] = d.H.map(s); o['style'] = d.I.map(s)
    o['item'] = d.M.map(s); o['qty'] = d.Q.map(num); o['cust_date'] = d.R.map(date)
    o['etd_gap'] = d.S.map(lambda x: float(x) if isinstance(x, int) and not isinstance(x, bool) else np.nan)
    o['fac_etd'] = d.U.map(date); o['mat_eta'] = d.W.map(date); o['input_ok'] = d.X.map(date)
    o['input_txt'] = d.X.map(lambda x: x if isinstance(x, str) else None)
    o['smp_sent'] = d.Z; o['smp_reply'] = d.AA; o['smp_stage'] = d.AB.map(s)
    o['sam'] = d.AC.map(num); o['plan_daily'] = d.AD.map(num); o['sew_days'] = d.AF.map(num)
    o['start'] = d.AH.map(date); o['end'] = d.AI.map(date); o['etd_plan'] = d.AJ.map(date)
    for k, c in {'c1_qty': 'AP', 'c1_eff': 'AR', 'c1_learn': 'AS', 'c1_max': 'AT', 'c1_lt': 'AU', 'c1_avg': 'AV', 'c1_rate': 'AW', 'c1_cost': 'AX', 'c1_upc': 'AY',
                 'c2_qty': 'BB', 'c2_eff': 'BD', 'c2_learn': 'BE', 'c2_max': 'BF', 'c2_lt': 'BG', 'c2_avg': 'BH', 'c2_rate': 'BI', 'c2_cost': 'BJ', 'c2_upc': 'BK', 'BL': 'BL'}.items():
        o[k] = d[c].map(num)
    o['wh_date'] = d.BO.map(date); o['wh_gap'] = np.where(o.wh_date.notna(), d.BP.map(num), np.nan)
    o['comment'] = d.L.map(s)
    def br(b):
        if not isinstance(b, str) or not b: return '(trống)'
        base = b.upper().replace('別注', '').strip()
        return {'HH': 'HELLY HANSEN'}.get(base, base)
    o['brand'] = o.brand_raw.map(br)
    o['bespoke'] = o.brand_raw.fillna('').str.contains('別注')
    return o


if __name__ == '__main__':
    src, wd = sys.argv[1], sys.argv[2]
    os.makedirs(wd, exist_ok=True)
    raw = read_raw(src); raw.to_pickle(os.path.join(wd, 'raw.pkl'))
    clean(raw).to_pickle(os.path.join(wd, 'clean.pkl'))
    json.dump(read_dotted(src), open(os.path.join(wd, 'dotted.json'), 'w'))
    print('OK', len(raw), 'dòng')
