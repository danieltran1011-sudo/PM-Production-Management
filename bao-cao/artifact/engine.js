/* PLANTABLE engine – đọc file APIC P-Plan (.xlsm/.xlsx) và tính toàn bộ phân tích điều hành.
 * Chạy được trong trình duyệt (window.PlanEngine) và Node (module.exports). Cần JSZip.
 * Quy ước giống pipeline Python (bao-cao/scripts): xưởng APIC, line thường 32 CN / line nhỏ 16 CN,
 * dòng chấm (QTY nền chấm) không cộng SL, hàng SMS tách riêng, 受注進行表 tự dựng từ PLANTABLE.
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- ngày (số ngày kể từ 1970-01-01, UTC)
  const DAY = 86400000;
  const dFromSerial = s => Math.round(s) - 25569;
  const dFromYMD = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / DAY);
  const toDate = d => new Date(d * DAY);
  const ymd = d => { const t = toDate(d); return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()]; };
  const iso = d => d == null ? '' : toDate(d).toISOString().slice(0, 10);
  const dmy = d => { if (d == null) return ''; const [y, m, dd] = ymd(d); return `${String(dd).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`; };
  const dow = d => ((d % 7) + 11) % 7; // 0 = Chủ nhật
  const isWork = d => dow(d) !== 0;     // T2–T7
  const mIdx = d => { const [y, m] = ymd(d); return y * 12 + m - 1; };
  const mStart = mi => dFromYMD(Math.floor(mi / 12), (mi % 12) + 1, 1);
  const mEnd = mi => mStart(mi + 1) - 1;
  const mLabel = mi => `${String((mi % 12) + 1).padStart(2, '0')}/${Math.floor(mi / 12)}`;
  const wdCount = (a, b) => { let n = 0; for (let d = a; d <= b; d++) if (isWork(d)) n++; return n; };
  const mWd = mi => wdCount(mStart(mi), mEnd(mi));

  // ---------------------------------------------------------------- đọc xlsx
  function colNum(ref) { let n = 0; for (const ch of ref) { const c = ch.charCodeAt(0); if (c < 65 || c > 90) break; n = n * 26 + (c - 64); } return n; }
  function colName(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function unesc(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16))).replace(/&amp;/g, '&'); }

  async function readWorkbook(buf, JSZip, onProgress) {
    const zip = await JSZip.loadAsync(buf);
    const txt = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
    onProgress && onProgress('workbook');
    const wbx = await txt('xl/workbook.xml');
    const rels = await txt('xl/_rels/workbook.xml.rels');
    const sheets = [...wbx.matchAll(/<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g)].map(m => ({ name: unesc(m[1]), rid: m[2] }));
    const relMap = {}; for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) { const id = /Id="([^"]+)"/.exec(m[1]); const t = /Target="([^"]+)"/.exec(m[1]); if (id && t) relMap[id[1]] = t[1]; }
    const pt = sheets.find(s => s.name.replace(/\s/g, '').toUpperCase().startsWith('PLANTABLE'));
    if (!pt) throw new Error('NO_PLANTABLE');
    let target = relMap[pt.rid]; target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    onProgress && onProgress('strings');
    const sst = [];
    const ssx = await txt('xl/sharedStrings.xml');
    if (ssx) for (const m of ssx.matchAll(/<si>([\s\S]*?)<\/si>/g)) { let t = ''; for (const r of m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) t += r[1]; sst.push(unesc(t)); }
    onProgress && onProgress('styles');
    const stx = await txt('xl/styles.xml') || '';
    const fillsX = (/<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(stx) || [, ''])[1];
    const fills = [...fillsX.matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)].map(m => { const p = /patternType="(\w+)"/.exec(m[1] || ''); return p ? p[1] : 'none'; });
    const xfsX = (/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stx) || [, ''])[1];
    const xfFill = [...xfsX.matchAll(/<xf\b([^>]*?)(?:\/>|>)/g)].map(m => { const f = /fillId="(\d+)"/.exec(m[1]); return f ? +f[1] : 0; });
    onProgress && onProgress('sheet');
    const sx = await txt(target);
    const rows = new Map();
    const rowRe = /<row\b[^>]*?\br="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
    const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = rowRe.exec(sx))) {
      const r = +m[1]; if (r < 9 || !m[2]) continue;
      const row = {};
      let c; cRe.lastIndex = 0;
      while ((c = cRe.exec(m[2]))) {
        const a = c[1]; const ref = /r="([A-Z]+)\d+"/.exec(a); if (!ref) continue;
        const col = ref[1]; const t = (/\bt="(\w+)"/.exec(a) || [])[1]; const s = +((/\bs="(\d+)"/.exec(a) || [, 0])[1]);
        const inner = c[2] || '';
        const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
        let val = null;
        if (t === 's') val = v ? sst[+v[1]] : null;
        else if (t === 'inlineStr') { const is = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(inner); val = is ? unesc(is[1]) : null; }
        else if (t === 'str') val = v ? unesc(v[1]) : null;
        else if (t === 'e') val = v ? { err: v[1] } : null;
        else if (t === 'b') val = v ? v[1] === '1' : null;
        else val = v ? +v[1] : null;
        row[col] = val;
        if (col === 'Q') row._qfill = fills[xfFill[s]] || 'none';
      }
      rows.set(r, row);
    }
    return { sheetName: pt.name, rows };
  }

  // ---------------------------------------------------------------- làm sạch
  const nfkc = s => (s == null ? null : String(s).normalize('NFKC'));
  const S = x => { if (x == null || typeof x === 'object') return null; const t = nfkc(x).replace(/\n/g, ' ').replace(/"/g, '').trim(); return t || null; };
  const N = x => (typeof x === 'number' && isFinite(x) ? x : NaN);
  const D = x => (typeof x === 'number' && x > 42000 && x < 80000 ? dFromSerial(x) : null); // ≥ 2015
  const ORDER_ALL = s => { const m = /^(\d\d)(SS|FW)$/.exec(s || ''); return m ? (+m[1]) * 2 + (m[2] === 'FW' ? 1 : 0) : -1; };

  function clean(raw) {
    const out = []; let errCells = 0; const seasonRaw = new Set();
    for (const [r, x] of raw.rows) {
      if (r < 10) continue;
      for (const k in x) if (x[k] && typeof x[k] === 'object' && x[k].err) errCells++;
      if (x.Q == null && x.I == null) continue;
      if (typeof x.A === 'string') seasonRaw.add(x.A.trim());
      const brandRaw = S(x.G);
      const base = brandRaw ? brandRaw.toUpperCase().replace('別注', '').trim() : null;
      const o = {
        row: r, season: S(x.A), factoryRaw: S(x.B), line: N(x.C), part: S(x.D), workers: N(x.E), brandRaw,
        brand: base ? ({ HH: 'HELLY HANSEN' }[base] || base) : '(trống)', bespoke: !!(brandRaw && brandRaw.includes('別注')),
        contract: S(x.H), style: S(x.I), item: S(x.M), qty: N(x.Q), qtyRaw: N(x.Q),
        etdGap: typeof x.S === 'number' && Number.isInteger(x.S) ? x.S : NaN,
        facEtd: D(x.U), matEta: D(x.W), inputOk: D(x.X), inputTxt: typeof x.X === 'string' ? x.X : null,
        smpReply: x.AA, smpStage: S(x.AB), sam: N(x.AC), start: D(x.AH), end: D(x.AI), etdPlan: D(x.AJ),
        c1Qty: N(x.AP), c1Upc: N(x.AY), c2Qty: N(x.BB), c2Upc: N(x.BK), c2Avg: N(x.BH), c2Max: N(x.BF),
        comment: S(x.L), note: typeof x.V === 'string' ? x.V : (x.V == null ? '' : String(x.V)),
        Icode: x.I == null ? null : String(x.I), Pcode: x.P == null ? null : String(x.P),
        dotted: !['none', 'solid', undefined].includes(x._qfill),
      };
      o.factory = ['AP1A', 'AP1B', 'AP1C'].includes(o.factoryRaw) ? 'APIC' : o.factoryRaw;
      out.push(o);
    }
    return { rows: out, errCells, seasonVariants: seasonRaw.size };
  }

  // ---------------------------------------------------------------- tiện ích
  const sum = (a, f = x => x) => a.reduce((s, x) => { const v = f(x); return s + (isFinite(v) ? v : 0); }, 0);
  const groupBy = (a, f) => { const m = new Map(); for (const x of a) { const k = f(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; };
  const median = a => { const b = a.filter(isFinite).sort((x, y) => x - y); if (!b.length) return NaN; const h = b.length >> 1; return b.length % 2 ? b[h] : (b[h - 1] + b[h]) / 2; };
  const seasonList = rows => [...new Set(rows.map(r => r.season).filter(s => ORDER_ALL(s) >= 0))].sort((a, b) => ORDER_ALL(a) - ORDER_ALL(b));
  const prevSeason = s => { const k = ORDER_ALL(s) - 1; return `${String(Math.floor(k / 2)).padStart(2, '0')}${k % 2 ? 'FW' : 'SS'}`; };
  const yoySeason = s => `${String(+s.slice(0, 2) - 1).padStart(2, '0')}${s.slice(2)}`;
  const hw = (p, h) => { p = p || ''; if (p.startsWith('AB') || p === '' || p === 'SPL') return 0.5; return p.startsWith(h) ? 1 : 0; };
  const smpOk = x => (typeof x === 'number' && x > 42000) || (typeof x === 'string' && ['●', '継続', '並行進行でOK', 'パターンの通り進行OK'].includes(x.trim()));
  const matOkNote = note => { const f = String(note || '').split(/\n\s*\n/)[0]; return /生地[:：]\s*OK/.test(f) && /資材[:：]\s*OK/.test(f); };
  const lotBand = q => !isFinite(q) ? null : q <= 300 ? 0 : q <= 500 ? 1 : q <= 1000 ? 2 : q <= 2000 ? 3 : q <= 5000 ? 4 : 5;
  const samBand = s => !isFinite(s) ? null : s / 60 < 60 ? 0 : s / 60 < 120 ? 1 : s / 60 < 180 ? 2 : s / 60 < 240 ? 3 : 4;
  const dayBand = n => !isFinite(n) ? null : n <= 5 ? 0 : n <= 10 ? 1 : n <= 20 ? 2 : n <= 40 ? 3 : 4;

  // c1/c2 lô: tính lại sau khi loại dòng chấm (cùng công thức file gốc, điều chỉnh theo tỷ lệ)
  function ucalc(AP, SAM, W) {
    if (!(AP > 0) || !(SAM > 0) || !(W > 0)) return NaN;
    const eff = (3.51 * Math.log(AP) + 76.97) / 100, learn = 2.23 * Math.pow(SAM, 0.11), mx = Math.floor(28800 / SAM * W * eff);
    return mx > 0 ? (AP / mx + learn / 2) / AP : NaN;
  }

  root.__PE_BASE = { DAY, dFromYMD, ymd, iso, dmy, dow, isWork, mIdx, mStart, mEnd, mLabel, wdCount, mWd, readWorkbook, clean, sum, groupBy, median, seasonList, prevSeason, yoySeason, hw, smpOk, matOkNote, lotBand, samBand, dayBand, ucalc, ORDER_ALL, colName };
})(typeof window !== 'undefined' ? window : globalThis);

(function (root) {
  'use strict';
  const B = root.__PE_BASE;
  const { dFromYMD, ymd, iso, isWork, mIdx, mStart, mEnd, mLabel, wdCount, mWd, sum, groupBy, median, seasonList, prevSeason, yoySeason, hw, smpOk, matOkNote, lotBand, samBand, dayBand, ucalc, ORDER_ALL } = B;
  const isNum = Number.isFinite;
  const wdays = (a, b) => { const r = []; for (let d = a; d <= b; d++) if (isWork(d)) r.push(d); return r; };

  function analyze(cl, opt) {
    const REF = opt.ref;                 // ngày tham chiếu (số ngày)
    const SNAP = opt.snap ?? REF - 1;    // ngày của file PLANTABLE
    const CAPW = opt.capWorkers ?? 960;  // định biên
    const RATE = opt.rate ?? 46.29017951425554; // USD / công-ngày
    const TARGET = opt.target ?? 0.9;
    const all = cl.rows.map(r => Object.assign({}, r));
    // ---- cờ
    for (const o of all) {
      const k = ORDER_ALL(o.season);
      let late = false;
      if (k >= 0 && o.start != null) { const y = 2000 + +o.season.slice(0, 2); const lim = o.season.endsWith('SS') ? dFromYMD(y, 6, 30) : dFromYMD(y, 12, 31); late = o.start > lim; }
      o.stale = late && k <= ORDER_ALL('25FW') && k >= 0 && (ORDER_ALL(o.season) <= ORDER_ALL(seasonCut(REF)));
      o.chkSeason = late && !o.stale;
      o.gap = isNum(o.etdGap) && Math.abs(o.etdGap) < 365 ? o.etdGap : NaN;
      o.gapBad = isNum(o.etdGap) && !isNum(o.gap);
    }
    // ---- dòng chấm + tính lại lô 原価1 / 原価2
    const nd = all.filter(o => !o.dotted);
    const sumBy = (key) => { const m = new Map(); for (const o of nd) { const k = o[key]; if (k == null) continue; m.set(k, (m.get(k) || 0) + (isNum(o.qty) ? o.qty : 0)); } return m; };
    const byI = sumBy('Icode'), byP = sumBy('Pcode');
    let nAdj1 = 0;
    for (const o of all) {
      for (const [qk, uk, map, key] of [['c1Qty', 'c1Upc', byI, 'Icode'], ['c2Qty', 'c2Upc', byP, 'Pcode']]) {
        const nw = map.get(o[key]);
        if (nw == null || !isNum(o[qk]) || Math.abs(nw - o[qk]) < 1) continue;
        if (qk === 'c1Qty' && !o.dotted) nAdj1++;
        const r = ucalc(nw, o.sam, o.workers) / ucalc(o[qk], o.sam, o.workers);
        if (isNum(r)) { o[uk] = o[uk] * r; o[qk] = nw; }
      }
    }
    const DOT = all.filter(o => o.dotted);
    let rest = all.filter(o => !o.dotted);
    // ---- hàng SMS
    for (const o of rest) {
      const key = ((o.style || '') + ' ' + (o.contract || '') + ' ' + (o.item || '')).toUpperCase(); const com = o.comment || '';
      o.sms = (key.includes('SMS') || (/SMS/i.test(com) && !com.includes('SMSサンプル対応')) || /展示会サンプル|展示会投入/.test(com)) && o.qty <= 500;
    }
    const SMS = rest.filter(o => o.sms);
    const P = rest.filter(o => !o.sms);  // hàng sản xuất
    for (const o of P) {
      o.matLate = o.inputOk != null && o.start != null && o.inputOk > o.start;
      o.matUnk = !!(o.inputTxt && o.inputTxt.includes('未定'));
      o.smpOk = smpOk(o.smpReply);
      o.stdH = isNum(o.qty) && isNum(o.sam) ? o.qty * o.sam / 3600 : 0;
      o.cost1 = isNum(o.qty) && isNum(o.c1Upc) ? o.qty * o.c1Upc : NaN;
      o.cost2 = isNum(o.qty) && isNum(o.c2Upc) ? o.qty * o.c2Upc : NaN;
      o.lot = lotBand(o.c2Qty); o.samB = samBand(o.sam);
      o.sewDays = o.start != null && o.end != null ? o.end - o.start + 1 : NaN; o.dayB = dayBand(o.sewDays);
    }
    const seasons = seasonList(P);
    const R = { ref: REF, snap: SNAP, seasons, counts: { rows: all.length, dotted: DOT.length, dottedQty: sum(DOT, o => o.qtyRaw), sms: SMS.length, smsQty: sum(SMS, o => o.qty), prod: P.length, nAdj1, errCells: cl.errCells, seasonVariants: cl.seasonVariants } };

    // ---- theo mùa
    const seasonStat = (rows) => {
      const qty = sum(rows, o => o.qty), h = sum(rows, o => o.stdH), c1 = sum(rows, o => o.cost1), c2 = sum(rows, o => o.cost2);
      const qc2 = sum(rows.filter(o => isNum(o.cost2)), o => o.qty);
      return { n: rows.length, styles: new Set(rows.map(o => o.style)).size, qty, h, sam: qty ? h * 60 / sum(rows.filter(o => isNum(o.sam)), o => o.qty) : NaN, c1, c2, split: c2 - c1, splitPct: c1 ? c2 / c1 - 1 : NaN,
        upc: qc2 ? c2 / qc2 : NaN, lotMed: median(rows.map(o => o.c2Qty)), lotAvg: rows.length ? qty / rows.length : NaN, bespoke: qty ? sum(rows.filter(o => o.bespoke), o => o.qty) / qty : NaN,
        small: rows.length ? rows.filter(o => o.c2Qty <= 300).length / rows.length : NaN, smallQty: qty ? sum(rows.filter(o => o.lot === 0), o => o.qty) / qty : NaN,
        sewDays: (a => a.length ? sum(a) / a.length : NaN)(rows.map(o => o.sewDays).filter(isNum)), tnf: qty ? sum(rows.filter(o => o.brand === 'TNF'), o => o.qty) / qty : NaN,
        brands: new Set(rows.map(o => o.brand)).size };
    };
    const bySeason = groupBy(P, o => o.season);
    R.season = seasons.map(s => Object.assign({ season: s }, seasonStat(bySeason.get(s) || [])));
    for (const s of R.season) { const p = R.season.find(x => x.season === prevSeason(s.season)), y = R.season.find(x => x.season === yoySeason(s.season)); s.dPrev = p ? s.qty / p.qty - 1 : NaN; s.dYoy = y ? s.qty / y.qty - 1 : NaN; }
    const years = [...new Set(seasons.map(s => s.slice(0, 2)))];
    R.year = years.map(y => { const ss = R.season.filter(s => s.season.startsWith(y)); return { year: 2000 + +y, qty: sum(ss, s => s.qty), h: sum(ss, s => s.h), c2: sum(ss, s => s.c2), split: sum(ss, s => s.split), full: ss.filter(s => s.n >= 100).length === 2 }; });
    R.year.forEach((y, i) => { y.d = i ? y.qty / R.year[i - 1].qty - 1 : NaN; });
    R.seasonStat = seasonStat; R.P = P; R.SMS = SMS; R.DOT = DOT;

    // ---- brand
    const bq = new Map(); for (const o of P) bq.set(o.brand, (bq.get(o.brand) || 0) + (isNum(o.qty) ? o.qty : 0));
    const totQ = sum(P, o => o.qty);
    R.brandTotal = [...bq.entries()].sort((a, b) => b[1] - a[1]).map(([b, q]) => ({ brand: b, qty: q, share: q / totQ }));
    R.brandSeason = {}; for (const o of P) { const k = o.brand; R.brandSeason[k] = R.brandSeason[k] || {}; R.brandSeason[k][o.season] = (R.brandSeason[k][o.season] || 0) + (isNum(o.qty) ? o.qty : 0); }

    // ---- lô
    R.lot = [0, 1, 2, 3, 4, 5].map(b => { const rows = P.filter(o => o.lot === b); const q = sum(rows, o => o.qty), c = sum(rows, o => o.cost2), qc = sum(rows.filter(o => isNum(o.cost2)), o => o.qty);
      const eff = rows.map(o => o.c2Max ? o.c2Avg / o.c2Max : NaN).filter(isNum); return { band: b, n: rows.length, qty: q, cost: c, upc: qc ? c / qc : NaN, eff: eff.length ? sum(eff) / eff.length : NaN }; });
    const nl = sum(R.lot, x => x.n), ql = sum(R.lot, x => x.qty); R.lot.forEach(x => { x.nShare = x.n / nl; x.qShare = x.qty / ql; x.vsMid = x.upc / R.lot[4].upc; });

    // ---- công suất (chiếm line theo nửa line A/B, 32/16 CN, không cộng trùng)
    const occRows = P.concat(DOT.map(o => Object.assign({}, o, { qty: 0 })), SMS.map(o => Object.assign({}, o, { qty: 0 }))).filter(o => !o.stale && o.start != null && o.end != null && o.end >= o.start && o.end - o.start < 200 && isNum(o.line));
    const occ = new Map();     // key line|h|d -> Set(parts)
    const dayQty = new Map(), dayH = new Map(), lineDayRows = new Map();
    for (const o of occRows) {
      const days = wdays(o.start, o.end); if (!days.length) continue;
      for (const d of days) {
        if (o.factory === 'APIC') {
          for (const h of ['A', 'B']) if (hw(o.part, h) > 0) { const k = o.line + '|' + h + '|' + d; if (!occ.has(k)) occ.set(k, new Set()); occ.get(k).add(o.part || ''); }
          const mi = mIdx(d); dayQty.set(mi, (dayQty.get(mi) || 0) + (isNum(o.qty) ? o.qty / days.length : 0)); dayH.set(mi, (dayH.get(mi) || 0) + (isNum(o.qty) && isNum(o.sam) ? o.qty * o.sam / 3600 / days.length : 0));
          const kk = o.line + '|' + d; lineDayRows.set(kk, (lineDayRows.get(kk) || 0) + 1);
        }
      }
    }
    const occVal = ps => { for (const p of ps) if (p === '' || p === 'SPL' || p.startsWith('AB') || p === 'A' || p === 'B') return 32; return Math.min(32, 16 * [...ps].filter(p => /\d$/.test(p)).length); };
    const occM = new Map(), occLH = new Map();
    for (const [k, ps] of occ) { const [ln, h, d] = k.split('|'); const mi = mIdx(+d); const v = occVal(ps); occM.set(mi, (occM.get(mi) || 0) + v); const k2 = ln + '|' + h + '|' + mi; occLH.set(k2, (occLH.get(k2) || 0) + v); }
    const m0 = mIdx(REF);
    R.m0 = m0;
    R.cap = []; for (let mi = m0 - 10; mi <= m0 + 7; mi++) { const wd = mWd(mi), c = CAPW * wd, o = occM.get(mi) || 0;
      const starts = P.filter(x => x.factory === 'APIC' && x.start != null && x.start >= mStart(mi) && x.start <= mEnd(mi) && !x.stale).length;
      R.cap.push({ mi, wd, cap: c, occ: o, util: o / c, free: Math.max(c - o, 0), qty: dayQty.get(mi) || 0, h: dayH.get(mi) || 0, eff: o ? (dayH.get(mi) || 0) / (o * 8) : NaN, starts }); }
    const capW = (a, b) => { const x = R.cap.filter(c => c.mi >= a && c.mi <= b); return { occ: sum(x, c => c.occ), cap: sum(x, c => c.cap), qty: sum(x, c => c.qty) }; };
    const past = capW(m0 - 10, m0 - 1), now = capW(m0, m0 + 2), nx3 = capW(m0 + 3, m0 + 3), fut = capW(m0 + 4, m0 + 7);
    const prod = past.occ ? past.qty / past.occ : NaN;
    const wdGap = Math.max(fut.cap * TARGET - fut.occ, 0);
    R.capK = { uPast: past.occ / past.cap, uNow: now.occ / now.cap, uNext: nx3.occ / nx3.cap, uFut: fut.occ / fut.cap, uFut1: (R.cap.find(c => c.mi === m0 + 4) || {}).util, prod, wdGap, pcsGap: wdGap * prod, usdGap: wdGap * RATE,
      parallel: (() => { const v = [...lineDayRows.entries()].filter(([k]) => { const d = +k.split('|')[1]; return mIdx(d) >= m0 - 10 && mIdx(d) < m0; }).map(([, n]) => n); return v.length ? sum(v) / v.length / 2 : NaN; })(),
      effPast: (() => { const x = R.cap.filter(c => c.mi >= m0 - 10 && c.mi < m0); return sum(x, c => c.h) / (sum(x, c => c.occ) * 8); })(),
      startsAvg: (() => { const x = R.cap.filter(c => c.mi >= m0 - 10 && c.mi < m0); return sum(x, c => c.starts) / x.length; })() };
    // heatmap theo nửa line
    R.heat = []; const fwdAll = P.filter(o => !o.stale && o.end != null && o.end >= REF);
    for (let ln = 1; ln <= 15; ln++) for (const h of ['A', 'B']) {
      const cells = []; for (let mi = m0; mi <= m0 + 7; mi++) { const v = occLH.get(ln + '|' + h + '|' + mi) || 0; cells.push(v ? v / (32 * mWd(mi)) : null); }
      const ff = fwdAll.filter(o => o.factory === 'APIC' && o.line === ln && hw(o.part, h) > 0);
      const small = [...new Set(ff.map(o => o.part).filter(p => /^[AB]\d$/.test(p || '')))].sort();
      const last = ff.length ? Math.max(...ff.map(o => o.end)) : null;
      R.heat.push({ line: ln + h, cells, n: ff.length, qty: sum(ff, o => o.qty), freeFrom: last != null ? last + 1 : null, small: small.map(p => ln + p) });
    }
    // hiệu quả theo cụm line (10 tháng trước)
    R.linePerf = []; for (let ln = 1; ln <= 15; ln++) { const rows = P.filter(o => o.factory === 'APIC' && o.line === ln && !o.stale && o.start != null && o.start >= mStart(m0 - 10) && o.start < mStart(m0));
      const q = sum(rows, o => o.qty), h = sum(rows, o => o.stdH), c = sum(rows, o => o.cost2), e = rows.map(o => o.c2Max ? o.c2Avg / o.c2Max : NaN).filter(isNum);
      R.linePerf.push({ line: ln, n: rows.length, qty: q, h, sam: q ? h * 60 / q : NaN, lot: rows.length ? q / rows.length : NaN, cost: c, upc: q ? c / q : NaN, eff: e.length ? sum(e) / e.length : NaN }); }

    // ---- rủi ro giao hàng
    const fw = P.filter(o => !o.stale && o.end != null && o.end >= REF);
    const lvl = o => o.facEtd == null ? 'nofac' : !isNum(o.gap) ? 'check' : o.gap <= -30 ? 'crit' : o.gap <= -8 ? 'high' : o.gap < 0 ? 'watch' : 'ok';
    for (const o of fw) { o.level = lvl(o); o.mat = o.matUnk ? 'unk' : (o.inputOk != null && o.start != null ? (o.inputOk > o.start ? 'late' : 'ok') : ''); o.score = ({ crit: 30, high: 20, watch: 10 }[o.level] || 0) + (o.mat === 'late' ? 5 : 0) + (o.mat === 'unk' ? 5 : 0) + (o.smpOk ? 0 : 2) + (o.facEtd == null ? 3 : 0); o.value = isNum(o.qty) && isNum(o.c2Upc) ? o.qty * o.c2Upc : NaN; }
    fw.sort((a, b) => b.score - a.score || (a.start ?? 0) - (b.start ?? 0));
    R.risk = fw;
    const cnt = f => ({ n: fw.filter(f).length, q: sum(fw.filter(f), o => o.qty), v: sum(fw.filter(f), o => o.value) });
    R.riskK = { all: cnt(() => true), crit: cnt(o => o.level === 'crit'), high: cnt(o => o.level === 'high'), watch: cnt(o => o.level === 'watch'), nofac: cnt(o => o.level === 'nofac'), matLate: cnt(o => o.mat === 'late'), matUnk: cnt(o => o.mat === 'unk'), smp: cnt(o => !o.smpOk) };
    R.riskK.ge8 = cnt(o => o.level === 'crit' || o.level === 'high');
    const sh = s => { const x = P.filter(o => o.season === s && isNum(o.gap)); return x.length ? x.filter(o => o.gap < 0).length / x.length : NaN; };
    R.lateRate = Object.fromEntries(seasons.map(s => [s, sh(s)]));

    // ---- sẵn sàng 6 tuần, Gantt 16 tuần
    R.ready = fw.filter(o => o.start != null && o.start >= REF && o.start < REF + 42).sort((a, b) => a.start - b.start || a.line - b.line).map(o => {
      const nvl = o.matUnk ? 'unk' : (o.inputTxt && /OK/i.test(o.inputTxt)) ? 'ok' : o.inputOk != null ? (o.inputOk <= o.start ? 'ok' : 'late') : 'nodate';
      const etd = o.facEtd == null ? 'nofac' : o.level;
      const block = []; if (nvl === 'late' || nvl === 'unk') block.push('nvl'); if (!o.smpOk) block.push('smp'); if (etd === 'crit' || etd === 'high') block.push('etd'); if (etd === 'nofac') block.push('nofac');
      return { o, nvl, etd, ready: !block.length, block };
    });
    let w0 = REF; while (B.dow(w0) !== 1) w0++;
    R.weeks = Array.from({ length: 16 }, (_, i) => w0 + 7 * i);
    const gRows = fw.map(o => ({ o, crit: o.level === 'crit' || o.mat === 'late', high: o.level === 'high' || o.mat === 'unk', sms: false }))
      .concat(DOT.filter(o => !o.stale && o.end != null && o.end >= REF).map(o => ({ o: Object.assign({}, o, { qty: o.qtyRaw }), crit: false, high: false, sms: false })))
      .concat(SMS.filter(o => !o.stale && o.end != null && o.end >= REF).map(o => ({ o, crit: false, high: false, sms: true })));
    R.gantt = [];
    for (let ln = 1; ln <= 15; ln++) for (const h of ['A', 'B']) {
      const sub = gRows.filter(g => g.o.factory === 'APIC' && g.o.line === ln && hw(g.o.part, h) > 0);
      R.gantt.push({ line: ln + h, cells: R.weeks.map(ws => { const act = sub.filter(g => g.o.start <= ws + 6 && g.o.end >= ws); if (!act.length) return null; const tp = act.slice().sort((a, b) => (b.o.qty || 0) - (a.o.qty || 0))[0];
        const pp = /^[AB]\d$/.test(tp.o.part || '') ? tp.o.part : ''; const allSms = act.every(g => g.sms);
        return { text: (allSms ? '[SMS] ' : (pp ? `[${pp}] ` : '')) + String(tp.o.style || '').slice(0, 12) + (act.length > 1 ? ` +${act.length - 1}` : ''), st: act.some(g => g.crit) ? 'crit' : act.some(g => g.high) ? 'high' : allSms ? 'sms' : 'ok' }; }) });
    }
    R.smsFwd = SMS.filter(o => !o.stale && o.end != null && o.end >= REF).sort((a, b) => a.start - b.start);
    R.smsSeason = seasonList(SMS.concat(P)).map(s => { const x = SMS.filter(o => o.season === s); const cd = sum(x, o => (isNum(o.workers) ? o.workers : 0) * (o.start != null && o.end != null && o.end >= o.start && o.end - o.start < 200 ? wdCount(o.start, o.end) : 0));
      const bb = [...groupBy(x, o => o.brand).entries()].map(([b, a]) => [b, sum(a, o => o.qty)]).sort((p, q) => q[1] - p[1]).slice(0, 3).map(p => p[0]); return { season: s, n: x.length, qty: sum(x, o => o.qty), cd, brands: bb }; });

    // ---- 受注進行表 tự dựng
    R.juchu = juchu(P, DOT, SMS, REF, m0);

    // ---- chất lượng dữ liệu
    const blank = P.filter(o => o.brand === '(trống)');
    R.dq = { blankN: blank.length, blankQ: sum(blank, o => o.qty), brandVariants: new Set(all.map(o => o.brandRaw)).size, stale: all.filter(o => o.stale).map(o => o.row), chk: all.filter(o => o.chkSeason).map(o => o.row),
      gapBad: all.filter(o => o.gapBad).length, noFac: fw.filter(o => o.facEtd == null).length, stages: new Set(all.map(o => o.smpStage).filter(Boolean)).size, upcBad: P.filter(o => o.c2Upc > 200).length };

    // ---- số liệu cho văn bản
    R.K = textNumbers(R, P);
    return R;
  }
  function seasonCut(REF) { const [y, m] = ymd(REF); return `${String((y - 2000) - 1).padStart(2, '0')}FW`; }

  function textNumbers(R, P) {
    const K = {}; const ss = R.season; const last = s => ss.find(x => x.season === s);
    const fulls = R.year.filter(y => y.full);
    K.years = fulls.slice(-3);
    const cur = ss.filter(s => s.n >= 100).slice(-1)[0] || ss[ss.length - 1]; // mùa đang sản xuất chính
    K.cur = cur.season; K.curQty = cur.qty; K.curYoy = cur.dYoy; K.curPrev = cur.dPrev;
    const fwS = ss.filter(s => s.season.endsWith('FW') && s.n >= 100), ssS = ss.filter(s => s.season.endsWith('SS') && s.n >= 100);
    K.samFW = [Math.min(...fwS.map(s => s.sam)), Math.max(...fwS.map(s => s.sam))]; K.samSS = [Math.min(...ssS.map(s => s.sam)), Math.max(...ssS.map(s => s.sam))];
    K.hours = [Math.min(...fulls.map(y => y.h)) / 1e6, Math.max(...fulls.map(y => y.h)) / 1e6];
    K.tnf = (R.brandTotal.find(b => b.brand === 'TNF') || { share: 0 }).share; K.top3 = sum(R.brandTotal.slice(0, 3), b => b.share);
    K.tnfCur = cur.tnf;
    const bs = (b, s) => (R.brandSeason[b] || {})[s] || 0;
    const yo = yoySeason(K.cur);
    K.brandGrowth = ['TNF', 'HELLY HANSEN', 'GOLDWIN', 'ALLTERRAIN 81', 'MONTBELL', 'TETON BROS'].map(b => ({ brand: b, cur: bs(b, K.cur), yoy: bs(b, yo), g: bs(b, yo) ? bs(b, K.cur) / bs(b, yo) - 1 : NaN }));
    const lot = R.lot; K.lotRows = lot[0].nShare; K.lotQty = lot[0].qShare; K.upcS = lot[0].upc; K.upcM = lot[4].upc; K.lotX = lot[0].upc / lot[4].upc; K.effS = lot[0].eff;
    K.split = sum(ss, s => s.split); K.splitCur = cur.splitPct; K.splitList = ss.filter(s => s.n >= 100).slice(-4).map(s => [s.season, s.splitPct]);
    K.bespokeMax = Math.max(...ss.filter(s => s.n >= 100).map(s => s.bespoke));
    return K;
  }

  // ---------------------------------------------------------------- 受注進行表
  const SEASON_CODE = s => (/^\d\d(SS|FW)$/.test(s || '') ? (s.endsWith('SS') ? 'S' : 'F') + s.slice(0, 2) : '?');
  const BSHORT = { 'HELLY HANSEN': 'HH', 'ALLTERRAIN 81': 'ALLTERRAIN', '(trống)': '?', 'TETON BROS': 'TETON' };
  function juchu(P, DOT, SMS, REF, m0) {
    const rows = P.map(o => [o, 'SX']).concat(DOT.map(o => [o, 'DOT']), SMS.map(o => [o, 'SMS'])).filter(([o]) => !o.stale);
    const E = []; // {line,h,d,w,conf,label,season,brand,kind,small,part}
    for (const [o, kind] of rows) {
      if (o.start == null || o.end == null || !isNum(o.line) || o.end < o.start || o.end - o.start > 200) continue;
      const p = o.part || ''; let hs;
      if (p.startsWith('AB')) hs = { A: 1, B: 1 }; else if (p === '') hs = { A: .5, B: .5 }; else if (/^[AB]\d$/.test(p)) hs = { [p[0]]: .5 }; else if (p === 'A' || p === 'B') hs = { [p]: 1 }; else continue;
      const conf = o.start <= REF || matOkNote(o.note);
      const label = SEASON_CODE(o.season) + (kind === 'SMS' ? '_SMS' : '') + '_' + (BSHORT[o.brand] || o.brand);
      const small = /^[AB]\d$/.test(p);
      for (const d of wdays(o.start, o.end)) for (const h in hs) E.push({ line: o.line, h, d, w: hs[h], conf, label, season: o.season, brand: o.brand, kind, small, part: p });
    }
    // phủ theo line-ngày
    const g = new Map(); for (const e of E) { const k = e.line + '|' + e.h + '|' + e.d; let x = g.get(k); if (!x) g.set(k, x = { w: 0, c: 0, list: [] }); x.w += e.w; if (e.conf) x.c += e.w; x.list.push(e); }
    const cov = new Map(); for (const [k, x] of g) { const tot = Math.min(1, x.w), conf = Math.min(tot, x.c); cov.set(k, { tot, conf }); for (const e of x.list) e.alloc = e.w / Math.max(x.w, 1e-9) * tot; }
    const mon = []; for (let mi = m0 - 2; mi <= m0 + 9; mi++) { let t = 0, c = 0; for (const [k, v] of cov) { const d = +k.split('|')[2]; if (mIdx(d) === mi) { t += v.tot; c += v.conf; } } const wd = mWd(mi); mon.push({ mi, wd, cap: 30 * wd, conf: c, tot: t }); }
    // lưới nửa tháng
    const per = []; for (let mi = m0 - 1; mi <= m0 + 9; mi++) { per.push([mStart(mi), mStart(mi) + 14, mLabel(mi).slice(0, 3) + mLabel(mi).slice(5) + ' 前']); per.push([mStart(mi) + 15, mEnd(mi), mLabel(mi).slice(0, 3) + mLabel(mi).slice(5) + ' 後']); }
    const grid = {};
    const byLH = groupBy(E, e => e.line + '|' + e.h);
    for (const [k, es] of byLH) {
      per.forEach(([a, b], j) => {
        const x = es.filter(e => e.d >= a && e.d <= b); if (!x.length) return;
        const by = new Map(); for (const e of x) by.set(e.label, (by.get(e.label) || 0) + e.alloc);
        const labs = [...by.entries()].sort((p, q) => q[1] - p[1]); const lab = labs[0][0];
        const yl = x.filter(e => e.label === lab); const cs = sum(yl.filter(e => e.conf), e => e.alloc) / Math.max(sum(yl, e => e.alloc), 1e-9);
        grid[k + '|' + j] = { lab, tent: cs < 0.5, small: yl.filter(e => e.small).length > yl.length / 2, more: labs.length > 1 };
      });
    }
    // mùa × brand các tháng tới
    const sb = {}; for (const e of E) { const mi = mIdx(e.d); if (mi < m0 + 1 || mi > m0 + 9) continue; const key = e.kind === 'SMS' ? e.season + ' SMS' : e.season + ' ' + e.brand; const r = sb[key] = sb[key] || { key, m: {}, tot: 0, tent: 0 }; r.m[mi] = (r.m[mi] || 0) + e.alloc; r.tot += e.alloc; if (!e.conf) r.tent += e.alloc; }
    const sbList = Object.values(sb).sort((a, b) => b.tot - a.tot);
    // từng line
    const lines = []; for (let ln = 1; ln <= 15; ln++) for (const h of ['A', 'B']) {
      const e = byLH.get(ln + '|' + h) || []; const fwd = e.filter(x => x.d >= REF); const cur = [...new Set(e.filter(x => x.d >= REF && x.d <= REF + 6).map(x => x.label))].sort();
      const ce = e.filter(x => x.conf); const confEnd = ce.length ? Math.max(...ce.map(x => x.d)) : null; const end = e.length ? Math.max(...e.map(x => x.d)) : null;
      lines.push({ line: ln + h, cur: cur.join(', '), confEnd, end, last: end != null ? [...new Set(e.filter(x => x.d === end).map(x => x.label))].sort().join(', ') : '', small: [...new Set(fwd.map(x => x.part).filter(p => /^[AB]\d$/.test(p)))].sort().map(p => ln + p).join(', ') });
    }
    // cơ cấu brand theo năm
    const [yR] = ymd(REF); const brandYear = {}; for (const e of E) { const y = ymd(e.d)[0]; if (y < yR - 3 || y > yR) continue; brandYear[e.brand] = brandYear[e.brand] || {}; brandYear[e.brand][y] = (brandYear[e.brand][y] || 0) + e.alloc; }
    const w1 = [m0 + 1, m0 + 3], w2 = [m0 + 4, m0 + 7];
    const win = w => { const x = mon.filter(r => r.mi >= w[0] && r.mi <= w[1]); return { tot: sum(x, r => r.tot) / sum(x, r => r.cap), conf: sum(x, r => r.conf) / sum(x, r => r.cap), free: sum(x, r => Math.max(r.cap - r.tot, 0)), tentShare: sum(x, r => r.tot - r.conf) / Math.max(sum(x, r => r.tot), 1e-9) }; };
    const limD = mStart(w2[0]) + 14;
    return { mon, per, grid, sb: sbList, lines, brandYear, years: [yR - 3, yR - 2, yR - 1, yR], w1, w2, k1: win(w1), k2: win(w2), lim: limD, nFree: lines.filter(l => l.end != null && l.end < limD).length };
  }

  root.PlanEngine = Object.assign({}, B, { analyze, async run(buf, JSZip, opt, onProgress) { const raw = await B.readWorkbook(buf, JSZip, onProgress); onProgress && onProgress('analyze'); const cl = B.clean(raw); const R = analyze(cl, opt); R.sheetName = raw.sheetName; return R; } });
  if (typeof module !== 'undefined' && module.exports) module.exports = root.PlanEngine;
})(typeof window !== 'undefined' ? window : globalThis);

(function (root) {
  'use strict';
  const PE = root.PlanEngine; const isNum = Number.isFinite;
  const r2 = x => (isNum(x) ? Math.round(x * 1000) / 1000 : null);
  // Bản tóm gọn cho giao diện + lưu db (không chứa dữ liệu thô)
  function view(R, meta) {
    const o2 = o => ({ row: o.row, s: o.season, l: o.line, p: o.part, b: o.brand, st: o.style, it: o.item, q: o.qty, a: o.start, e: o.end, ep: o.etdPlan, fe: o.facEtd, g: isNum(o.gap) ? o.gap : null,
      lv: o.level, mt: o.mat, sm: o.smpOk ? 1 : 0, io: o.inputOk, itx: o.inputTxt, v: r2(o.value), stg: o.smpStage, c: (o.comment || '').slice(0, 120) });
    // khối dữ liệu cho phần so sánh mùa (chọn mùa trên trang)
    const cube = {};
    for (const o of R.P) {
      const s = o.season; if (!s) continue; const c = cube[s] = cube[s] || { n: 0, q: 0, h: 0, c2: 0, qc: 0, split: 0, qs: 0, sd: 0, sdn: 0, brand: {}, lot: {}, sam: {}, day: {}, bes: 0, tnf: 0 };
      const q = isNum(o.qty) ? o.qty : 0; c.n++; c.q += q; c.h += o.stdH; if (isNum(o.cost2)) { c.c2 += o.cost2; c.qc += q; } if (isNum(o.cost2) && isNum(o.cost1)) c.split += o.cost2 - o.cost1; if (isNum(o.sam)) c.qs += q;
      if (isNum(o.sewDays)) { c.sd += o.sewDays; c.sdn++; } if (o.bespoke) c.bes += q; if (o.brand === 'TNF') c.tnf += q;
      const add = (m, k) => { if (k == null) return; const x = m[k] = m[k] || [0, 0, 0, 0, 0]; x[0]++; x[1] += q; x[2] += o.stdH; if (isNum(o.cost2)) { x[3] += o.cost2; x[4] += q; } };
      add(c.brand, o.brand); add(c.lot, o.lot); add(c.sam, o.samB); add(c.day, o.dayB);
    }
    const J = R.juchu;
    return {
      v: 1, meta, ref: R.ref, snap: R.snap, m0: R.m0, seasons: R.seasons, counts: R.counts,
      season: R.season.map(s => { const x = Object.assign({}, s); for (const k in x) if (typeof x[k] === 'number') x[k] = r2(x[k]); return x; }),
      year: R.year, brandTotal: R.brandTotal.slice(0, 25), brandSeason: Object.fromEntries(R.brandTotal.slice(0, 25).map(b => [b.brand, R.brandSeason[b.brand]])),
      lot: R.lot, cap: R.cap, capK: R.capK, heat: R.heat, linePerf: R.linePerf, riskK: R.riskK, lateRate: R.lateRate, K: R.K, dq: R.dq, cube,
      risk: R.risk.map(o2), ready: R.ready.map(x => Object.assign(o2(x.o), { nvl: x.nvl, etd: x.etd, ok: x.ready ? 1 : 0, blk: x.block })),
      weeks: R.weeks, gantt: R.gantt, smsFwd: R.smsFwd.map(o2), smsSeason: R.smsSeason,
      juchu: { mon: J.mon, per: J.per, grid: J.grid, sb: J.sb.slice(0, 20), lines: J.lines, brandYear: J.brandYear, years: J.years, w1: J.w1, w2: J.w2, k1: J.k1, k2: J.k2, lim: J.lim, nFree: J.nFree },
    };
  }
  // tách view thành các phần ≤ 200 KB để lưu db
  function split(vw) {
    const parts = {}; const big = ['risk', 'ready', 'gantt', 'juchu', 'cube', 'heat'];
    const core = {}; for (const k in vw) if (!big.includes(k)) core[k] = vw[k];
    parts.core = core; for (const k of big) parts[k] = { [k]: vw[k] };
    return Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, JSON.stringify(v)]));
  }
  PE.view = view; PE.splitView = split;
})(typeof window !== 'undefined' ? window : globalThis);
