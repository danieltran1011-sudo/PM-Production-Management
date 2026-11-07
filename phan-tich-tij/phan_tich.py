"""Phân tích file chi phí gia công TIJ gửi hàng tháng (APIC1A / APIC1B).

Cách dùng:
    python3 phan-tich-tij/phan_tich.py <file.xlsm> [file-bao-cao.html]

Đọc sheet NEW (từng mã hàng trên từng chuyền + sản lượng từng ngày) và sheet HCM
(nhân sự, tiền công, tổng tiền tháng), tính các chỉ số rồi xuất một trang HTML.
Cần pandas và openpyxl.
"""
import datetime
import json
import os
import sys
import unicodedata

import openpyxl
import pandas as pd

FACTORIES = ("APIC1A", "APIC1B")
WORK_SECONDS = 8 * 3600  # hiệu suất quy đổi theo ngày 8 giờ

# Vị trí cột sheet NEW (0 = cột A), theo dòng tiêu đề 7
COLS = {
    0: "factory", 1: "line", 2: "part", 3: "season", 4: "sales", 5: "dept", 6: "brand",
    7: "style", 12: "ppl", 13: "sec", 14: "qty_prod", 15: "qty_done", 16: "remain",
    17: "avg_day", 18: "days",
    22: "qty1", 30: "unit1",
    31: "qty2", 38: "unit2",
    39: "rate3", 40: "cost3", 41: "unit3",
    42: "mqty", 50: "mcost5", 51: "munit5",
    60: "r31", 61: "d31", 62: "r32", 63: "d32", 64: "r52", 65: "d52",
}
NUMERIC = ["ppl", "sec", "qty_prod", "qty_done", "remain", "days", "qty1", "unit1", "qty2",
           "unit2", "rate3", "cost3", "unit3", "mqty", "mcost5", "munit5",
           "r31", "d31", "r32", "d32", "r52", "d52"]
HCM_COLS = ["year", "month", "factory", "wage", "total_ppl", "sew_ppl", "line_unit", "std_days",
            "sew_rate", "learnA", "learnB", "effA", "effB", "maxeff", "contract_ppl",
            "sew_ppl_all", "sample_ppl", "total_usd", "mass_usd", "sample_usd"]


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["NEW"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[6]
    month_info = (rows[5][0], rows[5][1])  # Year / Month ở A6:B6
    dcols = [(i, h) for i, h in enumerate(header) if isinstance(h, datetime.datetime)]
    recs, daily = [], []
    for ri, r in enumerate(rows[7:], start=8):
        if r[0] not in FACTORIES:
            continue
        d = {name: r[i] if i < len(r) else None for i, name in COLS.items()}
        d["row"] = ri
        recs.append(d)
        for i, h in dcols:
            v = r[i] if i < len(r) else None
            if isinstance(v, (int, float)) and v:
                daily.append((ri, h, v))
    df = pd.DataFrame(recs)
    dd = pd.DataFrame(daily, columns=["row", "date", "pcs"])

    hrows = []
    for r in wb["HCM"].iter_rows(min_row=4, max_col=21, values_only=True):
        if isinstance(r[1], int) and r[1] > 2000:
            hrows.append(r[1:21])
    hcm = pd.DataFrame(hrows, columns=HCM_COLS)
    wb.close()

    for c in NUMERIC:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["line", "part", "season", "sales", "style"]:
        df[c] = df[c].map(lambda v: "" if v is None else str(v).strip())
    df["brand_n"] = df.brand.map(lambda v: "" if v is None else str(v).strip().upper())
    df["dept_n"] = df.dept.map(lambda v: "" if v is None else
                               unicodedata.normalize("NFKC", str(v)).upper())
    dd = dd.merge(df[["row", "factory", "line", "part", "ppl", "sec"]], on="row")
    dd["es"] = dd.pcs * dd.sec  # giây chuẩn tạo ra
    dd["ym"] = dd.date.dt.to_period("M")
    hcm["ym"] = [pd.Period(year=y, month=m, freq="M") for y, m in zip(hcm.year, hcm.month)]
    span = dd.groupby("row").date.agg(["min", "max", "nunique"])
    span.columns = ["start", "end", "proddays"]
    df = df.merge(span, left_on="row", right_index=True, how="left")
    return df, dd, hcm, month_info


def r(v, n=1):
    if v is None or pd.isna(v):
        return None
    return round(float(v), n)


def variance_table(c, key, min_qty=0):
    g = c.groupby(key).agg(n=("row", "count"), qty=("qty_done", "sum"), act=("act", "sum"),
                           std2=("std2", "sum"), std1=("std1", "sum"), sec=("sec", "mean"))
    g = g[g.qty >= min_qty]
    out = []
    for k, v in g.iterrows():
        name = " · ".join(str(x) for x in k) if isinstance(k, tuple) else str(k)
        out.append({"name": name, "n": int(v.n), "qty": int(v.qty), "act": r(v.act, 0),
                    "std2": r(v.std2, 0), "gap2": r(v.act - v.std2, 0),
                    "vs2": r((v.act / v.std2 - 1) * 100, 1), "vs1": r((v.act / v.std1 - 1) * 100, 1),
                    "sec": r(v.sec, 0)})
    return out


def analyse(df, dd, hcm, month_info):
    year, month = int(month_info[0]), int(month_info[1])
    cur = pd.Period(year=year, month=month, freq="M")
    prev = cur - 1
    last_year = cur - 12
    win_start = (cur - 11).start_time  # 12 tháng gần nhất

    # ---------- theo tháng ----------
    line_part_day = dd.groupby(["factory", "line", "part", "date", "ym"]).agg(
        es=("es", "sum"), ppl=("ppl", "max"), pcs=("pcs", "sum")).reset_index()
    mm = dd.groupby(["ym", "factory"]).agg(pcs=("pcs", "sum"), es=("es", "sum"),
                                           days=("date", "nunique")).reset_index()
    cap = line_part_day.groupby(["ym", "factory"]).ppl.sum().rename("ppl_days").reset_index()
    mm = mm.merge(cap, on=["ym", "factory"])
    mm = mm.merge(hcm[["ym", "factory", "sew_ppl", "total_usd", "mass_usd", "sample_usd",
                       "wage", "total_ppl", "sew_rate"]], on=["ym", "factory"], how="left")
    mm["usd_pc"] = mm.mass_usd / mm.pcs
    mm["usd_min"] = mm.mass_usd / (mm.es / 60)
    mm["eff"] = mm.es / (mm.ppl_days * WORK_SECONDS) * 100
    mm["avg_sec"] = mm.es / mm.pcs

    monthly = []
    for ym, g in mm.groupby("ym"):
        item = {"ym": str(ym)}
        for _, v in g.iterrows():
            item[v.factory] = {"pcs": int(v.pcs), "hrs": r(v.es / 3600, 0), "paid": r(v.mass_usd, 0),
                               "usd_pc": r(v.usd_pc, 2), "usd_min": r(v.usd_min, 3),
                               "eff": r(v.eff, 1), "days": int(v.days), "avg_sec": r(v.avg_sec, 0)}
        monthly.append(item)

    def kpi(f, p):
        v = mm[(mm.ym == p) & (mm.factory == f)]
        if v.empty:
            return None
        v = v.iloc[0]
        return {"pcs": int(v.pcs), "hrs": r(v.es / 3600, 0), "paid": r(v.mass_usd, 0),
                "sample": r(v.sample_usd, 0), "usd_pc": r(v.usd_pc, 2), "usd_min": r(v.usd_min, 3),
                "eff": r(v.eff, 1), "sewers": r(v.sew_ppl, 0), "avg_sec": r(v.avg_sec, 0),
                "days": int(v.days)}

    last12 = mm[(mm.ym > cur - 12) & (mm.ym <= cur)]
    kpis = {}
    for f in FACTORIES:
        l12 = last12[last12.factory == f]
        kpis[f] = {"cur": kpi(f, cur), "prev": kpi(f, prev), "ly": kpi(f, last_year),
                   "avg12_usd_min": r(l12.mass_usd.sum() / (l12.es.sum() / 60), 3),
                   "avg12_pcs": r(l12.pcs.mean(), 0), "avg12_eff": r(
                       l12.es.sum() / (l12.ppl_days.sum() * WORK_SECONDS) * 100, 1),
                   "styles": int(dd[(dd.ym == cur) & (dd.factory == f)].row.nunique()),
                   "starts": int(((df.start.dt.to_period("M") == cur) & (df.factory == f)).sum())}

    # ---------- chuyền tháng hiện tại ----------
    oct_lines = []
    lpd = line_part_day[line_part_day.ym == cur]
    for (f, ln), g in lpd.groupby(["factory", "line"]):
        rows_cur = dd[(dd.ym == cur) & (dd.factory == f) & (dd.line == ln)]
        sub = df[df.row.isin(rows_cur.row.unique())]
        oct_lines.append({
            "factory": f, "line": ln, "pcs": int(g.pcs.sum()), "hrs": r(g.es.sum() / 3600, 0),
            "eff": r(g.es.sum() / (g.ppl.sum() * WORK_SECONDS) * 100, 1),
            "ppl": r(g.groupby("date").ppl.sum().median(), 0),
            "styles": int(rows_cur.row.nunique()),
            "starts": int((sub.start.dt.to_period("M") == cur).sum()),
            "avg_sec": r(g.es.sum() / g.pcs.sum(), 0),
            "brands": ", ".join(sub.brand_n.value_counts().index[:3]),
        })
    oct_lines.sort(key=lambda x: (x["factory"], int("".join(ch for ch in x["line"] if ch.isdigit()) or 0)))

    # ---------- chi phí thực tế so với dự toán (mã đã xong, 12 tháng) ----------
    c = df[(df.unit3 > 0) & df.unit2.notna() & (df.qty_done > 0) & (df.end >= win_start)].copy()
    c["std1"] = c.unit1 * c.qty_done
    c["std2"] = c.unit2 * c.qty_done
    c["act"] = c.unit3 * c.qty_done
    c["lot"] = pd.cut(c.qty_done, [0, 300, 600, 1000, 2000, 1e9],
                      labels=["Dưới 300", "300–600", "600–1.000", "1.000–2.000", "Trên 2.000"])
    c["cx"] = pd.cut(c.sec, [0, 3000, 5000, 8000, 12000, 1e9],
                     labels=["Dưới 3.000 s", "3.000–5.000 s", "5.000–8.000 s", "8.000–12.000 s",
                             "Trên 12.000 s"])
    tot = {"n": int(len(c)), "qty": int(c.qty_done.sum()), "act": r(c.act.sum(), 0),
           "std2": r(c.std2.sum(), 0), "std1": r(c.std1.sum(), 0),
           "gap2": r(c.act.sum() - c.std2.sum(), 0), "gap1": r(c.act.sum() - c.std1.sum(), 0),
           "vs2": r((c.act.sum() / c.std2.sum() - 1) * 100, 1),
           "vs1": r((c.act.sum() / c.std1.sum() - 1) * 100, 1)}
    variance = {
        "total": tot,
        "factory": variance_table(c, "factory"),
        "line": variance_table(c, ["factory", "line"]),
        "brand": sorted(variance_table(c, "brand_n", 3000), key=lambda x: -x["qty"]),
        "sales": sorted(variance_table(c, "sales", 3000), key=lambda x: -x["qty"]),
        "season": variance_table(c, "season", 1000),
        "lot": variance_table(c.astype({"lot": str}), "lot"),
        "cx": variance_table(c.astype({"cx": str}), "cx"),
    }
    order_lot = ["Dưới 300", "300–600", "600–1.000", "1.000–2.000", "Trên 2.000"]
    order_cx = ["Dưới 3.000 s", "3.000–5.000 s", "5.000–8.000 s", "8.000–12.000 s", "Trên 12.000 s"]
    variance["lot"].sort(key=lambda x: order_lot.index(x["name"]))
    variance["cx"].sort(key=lambda x: order_cx.index(x["name"]))

    cols = ["factory", "line", "part", "season", "brand_n", "sales", "style", "sec", "ppl",
            "qty_done", "proddays", "unit2", "unit3", "d32", "end"]
    big = c[c.qty_done >= 100]

    def style_rows(s):
        out = []
        for _, v in s[cols].iterrows():
            out.append({"factory": v.factory, "line": f"{v.line}-{v.part}", "season": v.season,
                        "brand": v.brand_n, "sales": v.sales, "style": v.style, "sec": r(v.sec, 0),
                        "qty": int(v.qty_done), "days": int(v.proddays), "unit2": r(v.unit2, 2),
                        "unit3": r(v.unit3, 2), "gap": r(v.d32, 0),
                        "pct": r((v.unit3 / v.unit2 - 1) * 100, 1), "end": v.end.strftime("%d/%m/%Y")})
        return out
    top_over = style_rows(big.sort_values("d32", ascending=False).head(10))
    top_under = style_rows(big.sort_values("d32").head(10))

    # ---------- đường cong học nghề & chi phí đổi mã ----------
    x = dd[dd.ppl > 0].sort_values(["row", "date"]).copy()
    x["k"] = x.groupby("row").cumcount() + 1
    x["eff"] = x.es / (x.ppl * WORK_SECONDS)
    nmax = x.groupby("row").k.transform("max")
    x = x[nmax >= 10]
    plateau = x[x.k >= 8].groupby("row").eff.median().rename("pl")
    x = x.merge(plateau, on="row")
    x = x[x.pl > 0]
    x["rel"] = x.eff / x.pl
    recent = x[x.row.isin(df[df.start >= win_start].row)]
    curve = [{"day": int(k), "pct": r(v * 100, 1)}
             for k, v in recent[recent.k <= 15].groupby("k").rel.median().items()]
    ramp = recent[recent.k <= 7].copy()
    ramp["lost"] = (ramp.pl - ramp.eff).clip(lower=0)
    lost = ramp.groupby("row").lost.sum()
    sew_rate = hcm[(hcm.ym > cur - 12) & (hcm.ym <= cur) & hcm.factory.isin(FACTORIES)].sew_rate.mean()
    ppl = df.set_index("row").loc[lost.index].ppl
    paid12 = hcm[(hcm.ym > cur - 12) & (hcm.ym <= cur) & hcm.factory.isin(FACTORIES)].mass_usd.sum()
    changeover = {"starts": int(len(lost)),
                  "all_starts": int((df.start >= win_start).sum()),
                  "lost_days": r(lost.mean(), 1),
                  "usd": r((lost * ppl * sew_rate).sum(), 0),
                  "paid12": r(paid12, 0),
                  "sew_rate": r(sew_rate, 2)}

    # ---------- lịch sử dài hạn (HCM) ----------
    h = hcm[hcm.factory.isin(FACTORIES)]
    yearly = []
    for y, g in h.groupby("year"):
        a = g[g.factory == "APIC1A"]
        yearly.append({"year": int(y), "months": int(a.month.nunique()), "wage": r(a.wage.mean(), 0),
                       "ppl": r(a.total_ppl.mean(), 0), "sewers": r(a.sew_ppl_all.mean(), 0),
                       "paid": r(g.total_usd.sum(), 0)})
    wage_m = [{"ym": str(p), "wage": r(w, 0)} for p, w in
              h[h.factory == "APIC1A"].set_index("ym").wage.items()]

    # ---------- đang sản xuất dở ----------
    wip = df[df.start.notna() & df.unit3.isna() & (df.end.dt.to_period("M") == cur)]
    wip_list = []
    for _, v in wip.sort_values(["factory", "line"]).iterrows():
        wip_list.append({"factory": v.factory, "line": f"{v.line}-{v.part}", "brand": v.brand_n,
                         "style": v.style, "season": v.season, "qty": r(v.qty_prod, 0),
                         "done": r(v.qty_done, 0), "remain": r(-(v.remain or 0), 0),
                         "start": v.start.strftime("%d/%m"), "days": int(v.proddays)})

    # ---------- chất lượng dữ liệu ----------
    dq = []
    n0 = int((df.unit3 == 0).sum())
    if n0:
        dq.append(f"{n0} dòng có đơn giá thực tế ③ = 0 (thường là mã chỉ may vài chiếc, 1 ngày) — "
                  "các dòng này làm chênh lệch trông 'tốt' giả; báo cáo đã loại ra.")
    nfw = int(df.dept.map(lambda v: isinstance(v, str) and v != unicodedata.normalize("NFKC", v)).sum())
    if nfw:
        dq.append(f"{nfw} dòng ghi mã phòng ban bằng chữ full-width (ví dụ ＡＤ５ thay vì AD5) — pivot của TIJ "
                  "sẽ tách thành 2 nhóm khác nhau.")
    nbr = df.groupby("brand_n").brand.nunique()
    dup = [b for b, k in nbr.items() if k > 1 and b]
    if dup:
        dq.append("Tên brand viết hoa/thường không thống nhất: " + ", ".join(dup[:6]) + ".")
    nseason = int(((df.season == "") | (df.season == "None")).sum())
    ndummy = int((df.season == "ﾀﾞﾐｰ").sum())
    dq.append(f"{nseason} dòng không có mùa (Status) và {ndummy} dòng 'ﾀﾞﾐｰ' (dữ liệu giả để giữ chỗ).")
    na = int((df.avg_day.astype(str) == "#N/A").sum())
    if na:
        dq.append(f"{na} dòng có lỗi #N/A ở cột số lượng bình quân/ngày (chủ yếu mã cũ hoặc chưa chạy).")
    dq.append("Sheet HCM ghi BTNB cố định 697 USD/người, 1.141 người cho mọi tháng — có vẻ là số tham chiếu, "
              "không phải số thực tế từng tháng.")
    dq.append("Đơn giá công/người tăng mạnh vào một số tháng (6, 7, 11, 12) — nhiều khả năng là tháng có "
              "thưởng/phụ cấp; chi phí/phút các tháng này cao hơn dù năng suất không đổi.")

    return {
        "month": f"{month:02d}/{year}", "ym": str(cur), "prev": str(prev), "ly": str(last_year),
        "generated": datetime.date.today().strftime("%d/%m/%Y"),
        "rows": int(len(df)), "first_day": dd.date.min().strftime("%m/%Y"),
        "last_day": dd.date.max().strftime("%d/%m/%Y"),
        "kpis": kpis, "monthly": monthly, "lines": oct_lines, "variance": variance,
        "top_over": top_over, "top_under": top_under, "curve": curve, "changeover": changeover,
        "yearly": yearly, "wage_m": wage_m, "wip": wip_list, "dq": dq,
    }


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    df, dd, hcm, info = load(src)
    data = analyse(df, dd, hcm, info)
    here = os.path.dirname(os.path.abspath(__file__))
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, f"bao-cao-{data['ym']}.html")
    template = open(os.path.join(here, "template.html"), encoding="utf-8").read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = template.replace("__DATA__", payload).replace("__MONTH__", data["month"])
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print("Đã ghi", out)


if __name__ == "__main__":
    main()
