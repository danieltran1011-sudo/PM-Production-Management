# -*- coding: utf-8 -*-
"""Dựng 受注進行表 (tiến độ nhận đơn) trực tiếp từ dữ liệu PLANTABLE đã làm sạch.

Quy ước:
- Mỗi line (cụm) gồm 2 line thường A/B (32 CN). Line nhỏ 小ロット (A1, A2, B1, B2…) = 0,5 line.
- Đơn chuyền lớn AB chiếm cả A và B. Dòng chấm (AB ghi lần 2) chỉ dùng để chiếm line.
- Đã chốt (ước tính) = đã bắt đầu may (START ≤ ngày tham chiếu) HOẶC ghi chú NVL mới nhất có 生地 OK và 資材 OK.
  Còn lại = tạm tính (暫定), tương đương chữ đỏ trong 受注進行表.
"""
import re
import numpy as np
import pandas as pd

SEASON_CODE = lambda s: (('S' if str(s).endswith('SS') else 'F') + str(s)[:2]) if isinstance(s, str) and len(s) >= 4 else '?'
BRAND_SHORT = {'HELLY HANSEN': 'HH', 'ALLTERRAIN 81': 'ALLTERRAIN', '(trống)': '?', 'TETON BROS': 'TETON'}


def mat_ok_from_note(note):
    first = str(note or '').split('\n\n')[0]
    return bool(re.search(r'生地[:：]\s*OK', first)) and bool(re.search(r'資材[:：]\s*OK', first))


def wdays(a, b):
    return [d for d in pd.date_range(a, b) if d.dayofweek < 6]


def expand(rows, REF):
    """rows: DataFrame có line, part, start, end, season, brand, kind ('SX'|'SMS'|'DOT'), conf"""
    rec = []
    for r in rows.itertuples():
        if pd.isna(r.start) or pd.isna(r.end) or pd.isna(r.line) or r.end < r.start or (r.end - r.start).days > 200:
            continue
        p = r.part if isinstance(r.part, str) else ''
        if p.startswith('AB') or p == '':
            hs = {'A': 1.0, 'B': 1.0} if p.startswith('AB') else {'A': 0.5, 'B': 0.5}
        elif re.match(r'^[AB]\d$', p):
            hs = {p[0]: 0.5}
        elif p in ('A', 'B'):
            hs = {p: 1.0}
        else:
            continue
        lab = SEASON_CODE(r.season) + ('_SMS' if r.kind == 'SMS' else '') + '_' + BRAND_SHORT.get(r.brand, str(r.brand))
        small = bool(re.match(r'^[AB]\d$', p))
        for d in wdays(r.start, r.end):
            for h, w in hs.items():
                rec.append((int(r.line), h, d, w, bool(r.conf), lab, r.season, r.brand, r.kind, small, p))
    return pd.DataFrame(rec, columns=['line', 'half', 'd', 'w', 'conf', 'label', 'season', 'brand', 'kind', 'small', 'part'])


def coverage(E):
    g = E.groupby(['line', 'half', 'd'])
    tot = g.w.sum().clip(upper=1)
    conf = E[E.conf].groupby(['line', 'half', 'd']).w.sum().reindex(tot.index).fillna(0)
    df = pd.DataFrame({'tot': tot, 'conf': np.minimum(conf, tot)}).reset_index()
    df['m'] = df.d.dt.to_period('M')
    return df


def allocate(E):
    """Phân bổ độ phủ (≤1) của mỗi line-ngày cho từng nhãn theo tỷ trọng."""
    E = E.copy()
    s = E.groupby(['line', 'half', 'd']).w.transform('sum')
    E['alloc'] = E.w / np.maximum(s, 1e-9) * np.minimum(s, 1)
    return E


def monthly(C, months):
    out = []
    for m in months:
        c = C[C.m == m]
        wd = len(wdays(m.start_time, m.end_time))
        out.append(dict(m=m, wd=wd, cap=30 * wd, conf=float(c.conf.sum()), tot=float(c.tot.sum())))
    return pd.DataFrame(out)


def grid(E, periods):
    """periods: list of (start,end,label). -> dict[(line,half,k)] = (label, conf_share, small)"""
    A = allocate(E)
    G = {}
    for k, (a, b, _) in enumerate(periods):
        x = A[(A.d >= a) & (A.d <= b)]
        if x.empty:
            continue
        nd = len(wdays(a, b))
        for (ln, h), y in x.groupby(['line', 'half']):
            by = y.groupby('label').alloc.sum().sort_values(ascending=False)
            lab = by.index[0]
            yl = y[y.label == lab]
            cs = yl[yl.conf].alloc.sum() / max(yl.alloc.sum(), 1e-9)
            cover = y.alloc.sum() / max(nd, 1)
            G[(ln, h, k)] = (lab, cs, bool(yl.small.mean() > 0.5), cover, len(by))
    return G


def per_line(E, REF):
    rows = []
    for ln in range(1, 16):
        for h in 'AB':
            e = E[(E.line == ln) & (E.half == h)]
            fw = e[e.d >= REF]
            cur = e[(e.d >= REF) & (e.d <= REF + pd.Timedelta(days=6))]
            ce = e[e.conf].d.max() if len(e[e.conf]) else pd.NaT
            te = e.d.max() if len(e) else pd.NaT
            last = e[e.d == te].label.unique() if len(e) else []
            sm = sorted(set(f'{ln}{p}' for p in fw.part.unique() if re.match(r'^[AB]\d$', str(p))))
            rows.append(dict(line=f'{ln}{h}', cur=', '.join(sorted(cur.label.unique()))[:40], conf_end=ce, tent_end=te,
                             last=', '.join(sorted(last))[:40], small=', '.join(sm)))
    return pd.DataFrame(rows)
