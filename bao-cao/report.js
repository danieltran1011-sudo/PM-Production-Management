/*
 * Tổng hợp báo cáo từ các file KH_UPDATE (sheet XN1AB): kế hoạch (KH) và thực hiện (TH)
 * theo ngày của từng mã hàng trên từng chuyền.
 *
 * - parseRows(rows, opts): đọc 1 sheet (mảng 2 chiều, ô ngày là số serial Excel) → dữ liệu theo mã × tháng.
 * - buildDataset(files): gộp nhiều file; mỗi tháng của mỗi xưởng lấy số từ file chốt tháng đó.
 * - summarize(ds, opts): so sánh tháng / tháng trước / cùng kỳ / lũy kế năm tài chính.
 *
 * Dùng được cả trong trình duyệt (window.PMReport) và Node (require).
 * Ngày luôn ở dạng chuỗi 'YYYY-MM-DD', tháng ở dạng 'YYYY-MM'.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PMReport = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var HOURS_PER_DAY = 8; // giờ làm việc/ngày, giống công thức năng lực trong file (3600*8)

  // ---------- Ngày tháng ----------
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function isoFromUTC(d) { return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function serialToISO(v, date1904) {
    var ms = Date.UTC(1899, 11, 30) + Math.round((v + (date1904 ? 1462 : 0)) * 864e5);
    return isoFromUTC(new Date(ms));
  }
  function toISO(v, date1904) {
    if (typeof v === 'number') return v > 20000 && v < 80000 ? serialToISO(Math.floor(v), date1904) : null;
    if (v instanceof Date && !isNaN(v)) {
      // SheetJS (cellDates) tạo ngày theo giờ máy; lấy ngày gần nhất để tránh lệch múi giờ
      var t = new Date(v.getTime() + 12 * 36e5);
      return t.getFullYear() + '-' + pad(t.getMonth() + 1) + '-' + pad(t.getDate());
    }
    if (typeof v === 'string') {
      var m = v.trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
      if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    }
    return null;
  }
  function addDays(iso, n) {
    var p = iso.split('-');
    return isoFromUTC(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)));
  }
  function weekday(iso) { var p = iso.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); }
  function addMonths(ym, n) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n;
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    return y + '-' + pad(m + 1);
  }
  function daysInMonth(ym) { return new Date(Date.UTC(+ym.slice(0, 4), +ym.slice(5, 7), 0)).getUTCDate(); }
  function monthStart(ym) { return ym + '-01'; }
  function monthEnd(ym) { return ym + '-' + pad(daysInMonth(ym)); }
  function monthLabel(ym) { return (+ym.slice(5, 7)) + '/' + ym.slice(0, 4); }

  // Năm tài chính đặt tên theo năm bắt đầu: bắt đầu tháng 4 → NTC 2026 = 4/2026–3/2027.
  function fiscalYear(ym, startMonth) {
    var s = startMonth || 1, y = +ym.slice(0, 4), m = +ym.slice(5, 7);
    return m >= s ? y : y - 1;
  }
  function fiscalMonths(fy, startMonth) {
    var first = fy + '-' + pad(startMonth || 1), out = [];
    for (var i = 0; i < 12; i++) out.push(addMonths(first, i));
    return out;
  }
  function fiscalLabel(fy, startMonth) {
    var s = startMonth || 1;
    if (s === 1) return 'NTC ' + fy;
    return 'NTC ' + fy + ' (' + s + '/' + fy + '–' + (s - 1) + '/' + (fy + 1) + ')';
  }

  // Ngày update lấy từ tên file, ví dụ "KH_UPDATE__APA1.20261031.xlsm" → 2026-10-31
  function updateDateFromName(name) {
    var re = /(20\d{2})(\d{2})(\d{2})/g, m;
    while ((m = re.exec(String(name || '')))) {
      var mo = +m[2], d = +m[3];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return m[1] + '-' + m[2] + '-' + m[3];
    }
    return null;
  }

  // ---------- Chuẩn hóa ----------
  function norm(v) { return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); }
  function num(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
  function numOrNull(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string' && v.trim() !== '' && isFinite(+v)) return +v;
    return null;
  }

  // Tên nhãn hàng viết sai đã gặp trong file
  var BRAND_ALIASES = { 'ALLTERAIN': 'ALLTERRAIN', 'DESCSENTE': 'DESCENTE' };

  // "TÊN SAI = TÊN ĐÚNG", mỗi dòng một cặp
  function parseAliasText(text) {
    var out = {};
    String(text || '').split(/\r?\n/).forEach(function (line) {
      var p = line.split('=');
      if (p.length < 2) return;
      var from = norm(p[0]).toUpperCase(), to = norm(p.slice(1).join('=')).toUpperCase();
      if (from && to) out[from] = to;
    });
    return out;
  }
  function normBrand(raw, aliases) {
    var b = norm(raw).toUpperCase();
    if (!b) return '(CHƯA GHI)';
    if (aliases && aliases[b]) return aliases[b];
    return BRAND_ALIASES[b] || b;
  }
  function normItem(raw) { return norm(raw) || '(chưa ghi)'; }
  // Cụm chuyền theo số: 4A, 4B, 4AB, 4A1 → "4"
  function lineGroup(line) { var m = norm(line).match(/^(\d+)/); return m ? m[1] : norm(line) || '(chưa ghi)'; }
  // Chuyền gốc: chuyền tách gộp về chuyền chính (4A1, 4A2 → 4A; 4B1 → 4B; 4AB2 → 4AB)
  function baseLine(line) {
    var t = norm(line), m = t.match(/^(\d+)([A-Za-z]+)\d*$/);
    return m ? m[1] + m[2].toUpperCase() : t || '(chưa ghi)';
  }

  // ---------- Đọc sheet ----------
  var HEADERS = {
    factory: ['工場'], line: ['ライン'], brand: ['ブランド'], style: ['品番'], item: ['アイテム'],
    po: ['PO'], lot: ['集約ロット(pcs)'], qty: ['生産数(pcs)'], sam: ['工程分析時間(sec)'],
    hc: ['ライン編成人数(人/line)']
  };
  // Vị trí cột của mẫu file hiện tại (T..AC), dùng khi không tìm thấy tiêu đề
  var DEFAULT_COLS = { factory: 19, line: 20, brand: 21, style: 22, item: 23, po: 24, lot: 25, qty: 26, sam: 27, hc: 28 };

  function findColumns(rows) {
    var limit = Math.min(rows.length, 40);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var fc = -1, hasStyle = false;
      for (var c = 0; c < row.length; c++) {
        var t = norm(row[c]);
        if (t === '工場' && fc < 0) fc = c;
        if (t === '品番') hasStyle = true;
      }
      if (fc < 0 || !hasStyle) continue;
      var cols = {};
      Object.keys(HEADERS).forEach(function (k) {
        for (var c = fc; c < row.length; c++) {
          if (HEADERS[k].indexOf(norm(row[c])) >= 0) { cols[k] = c; return; }
        }
        cols[k] = DEFAULT_COLS[k];
      });
      return { headerRow: r, cols: cols, found: true };
    }
    return { headerRow: 5, cols: Object.assign({}, DEFAULT_COLS), found: false };
  }

  // Cột ngày: dãy ô ngày liên tiếp (mỗi ô +1 ngày) dài nhất trong các hàng tiêu đề
  function findDayColumns(rows, date1904) {
    var best = null, limit = Math.min(rows.length, 40);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [], start = -1, prev = null, dates = [];
      for (var c = 0; c <= row.length; c++) {
        var iso = c < row.length ? toISO(row[c], date1904) : null;
        if (iso && prev && iso === addDays(prev, 1)) { dates.push(iso); }
        else {
          if (start >= 0 && dates.length > 1 && (!best || dates.length > best.dates.length)) best = { row: r, start: start, dates: dates };
          start = iso ? c : -1; dates = iso ? [iso] : [];
        }
        prev = iso;
      }
    }
    return best;
  }

  function rowType(v) {
    var t = norm(v).toUpperCase();
    if (t === 'KH' || t === '計画') return 'kh';
    if (t === 'TH' || t === '実績') return 'th';
    return '';
  }
  function findTypeColumn(rows, from) {
    var best = 1, bestCount = -1;
    for (var c = 0; c < 8; c++) {
      var n = 0;
      for (var r = from; r < rows.length; r++) if (rows[r] && rowType(rows[r][c]) === 'kh') n++;
      if (n > bestCount) { bestCount = n; best = c; }
    }
    return best;
  }

  /*
   * rows: mảng 2 chiều của sheet XN1AB (ô ngày là số serial hoặc Date).
   * opts: { name, updateDate, date1904, aliases }
   */
  function parseRows(rows, opts) {
    opts = opts || {};
    var warnings = [];
    var hdr = findColumns(rows);
    if (!hdr.found) warnings.push('Không thấy hàng tiêu đề (工場, 品番): dùng vị trí cột mặc định T–AC.');
    var dayInfo = findDayColumns(rows, opts.date1904);
    if (!dayInfo) throw new Error('Không tìm thấy dãy cột ngày trong sheet.');
    var cols = hdr.cols, d0 = dayInfo.start, dates = dayInfo.dates, nd = dates.length;
    var firstData = Math.max(hdr.headerRow, dayInfo.row) + 1;
    var typeCol = findTypeColumn(rows, firstData);

    // Ngày nghỉ: hàng phía trên có nhiều ô "x" nhất trong các cột ngày; không có thì nghỉ Chủ nhật
    var off = new Array(nd), offRow = -1, offMax = 0;
    for (var r = 0; r < firstData; r++) {
      if (r === dayInfo.row || !rows[r]) continue;
      var cnt = 0;
      for (var j = 0; j < nd; j++) if (norm(rows[r][d0 + j]).toLowerCase() === 'x') cnt++;
      if (cnt > offMax) { offMax = cnt; offRow = r; }
    }
    for (var j2 = 0; j2 < nd; j2++) {
      off[j2] = offRow >= 0 ? norm(rows[offRow][d0 + j2]).toLowerCase() === 'x' : weekday(dates[j2]) === 0;
    }
    if (offRow < 0) warnings.push('Không thấy hàng đánh dấu ngày nghỉ ("x"): coi Chủ nhật là ngày nghỉ.');

    // Mã hàng: mỗi mã là cặp hàng KH + TH liền nhau (hàng trên KH là "dự định" lý thuyết)
    var styles = [], factoryCount = {}, missingHc = 0;
    for (var i = firstData; i < rows.length - 1; i++) {
      var kr = rows[i], tr = rows[i + 1];
      if (!kr || !tr || rowType(kr[typeCol]) !== 'kh' || rowType(tr[typeCol]) !== 'th') continue;
      var pr = rows[i - 1] || [];
      var get = function (k) { var v = kr[cols[k]]; return v == null || v === '' ? pr[cols[k]] : v; };
      var style = norm(get('style')), brandRaw = norm(get('brand'));
      if (!style || /^master$/i.test(style) || /^master$/i.test(brandRaw)) continue;
      var s = {
        factory: norm(get('factory')) || opts.factory || '', lineRaw: norm(get('line')) || '(chưa ghi)',
        brandRaw: brandRaw, brand: normBrand(brandRaw, opts.aliases), item: normItem(get('item')), style: style,
        po: norm(get('po')), sam: numOrNull(get('sam')), hc: numOrNull(get('hc')), qty: numOrNull(get('qty')),
        kh: new Array(nd), th: new Array(nd), endTH: -1
      };
      for (var j3 = 0; j3 < nd; j3++) {
        s.kh[j3] = Math.max(0, num(kr[d0 + j3]));
        var tv = tr[d0 + j3];
        s.th[j3] = Math.max(0, num(tv));
        if (s.endTH < 0 && norm(tv).toUpperCase() === 'END') s.endTH = j3;
      }
      if (!(s.hc > 0)) missingHc++;
      factoryCount[s.factory] = (factoryCount[s.factory] || 0) + 1;
      styles.push(s);
      i++;
    }
    if (missingHc) warnings.push(missingHc + ' mã không có số lao động chuyền: không tính vào số lao động.');

    var factory = opts.factory || Object.keys(factoryCount).sort(function (a, b) { return factoryCount[b] - factoryCount[a]; })[0] || '';
    if (!factory) {
      var fm = String(opts.name || '').match(/KH_UPDATE_+([A-Za-z0-9]+)/);
      factory = fm ? fm[1] : '(chưa rõ xưởng)';
      warnings.push('Không có tên xưởng trong cột 工場: dùng "' + factory + '".');
    }
    styles.forEach(function (st) { if (!st.factory) st.factory = factory; st.line = baseLine(st.lineRaw); });

    // Ngày update: theo tên file; không có thì lấy ngày cuối cùng có số TH
    var updateDate = opts.updateDate || updateDateFromName(opts.name), updateFrom = 'name';
    if (opts.updateDate) updateFrom = 'manual';
    if (!updateDate) {
      updateFrom = 'data';
      for (var j4 = nd - 1; j4 >= 0 && !updateDate; j4--) {
        for (var k = 0; k < styles.length; k++) if (styles[k].th[j4] > 0) { updateDate = dates[j4]; break; }
      }
      updateDate = updateDate || dates[nd - 1];
      warnings.push('Tên file không có ngày update: lấy ngày cuối có số TH (' + updateDate + ').');
    }

    // Lao động: mỗi chuyền gốc, mỗi ngày có sản lượng tính số người 1 lần rồi chia cho các mã theo
    // tỷ lệ sản lượng ngày đó. Số người của chuyền tách (A1, A2) lấy số lớn nhất trong ngày của
    // từng chuyền tách; chuyền gốc = max(số người chuyền chính, tổng các chuyền tách) để ngày
    // chuyển mã không bị đếm trùng.
    function allocate(key) {
      var pd = styles.map(function () { return new Array(nd).fill(0); });
      var ld = styles.map(function () { return new Array(nd).fill(0); });
      for (var j = 0; j < nd; j++) {
        var byBase = {};
        styles.forEach(function (st, idx) {
          if (!(st[key][j] > 0 && st.hc > 0)) return;
          var bk = st.factory + '|' + st.line;
          var b = byBase[bk] || (byBase[bk] = { ids: [], tot: 0, main: 0, subs: {} });
          b.ids.push(idx); b.tot += st[key][j];
          if (st.lineRaw.toUpperCase() === st.line) b.main = Math.max(b.main, st.hc);
          else b.subs[st.lineRaw] = Math.max(b.subs[st.lineRaw] || 0, st.hc);
        });
        Object.keys(byBase).forEach(function (bk) {
          var b = byBase[bk], subSum = 0;
          Object.keys(b.subs).forEach(function (k) { subSum += b.subs[k]; });
          var h = Math.max(b.main, subSum);
          b.ids.forEach(function (idx) {
            var share = styles[idx][key][j] / b.tot;
            pd[idx][j] = h * share; ld[idx][j] = share;
          });
        });
      }
      return { pd: pd, ld: ld };
    }
    var aK = allocate('kh'), aT = allocate('th');

    // Tháng có trong file và số ngày làm việc
    var months = {};
    for (var j5 = 0; j5 < nd; j5++) {
      var ym = dates[j5].slice(0, 7);
      var mo = months[ym] || (months[ym] = { days: 0, off: 0, lastWork: null });
      mo.days++; if (off[j5]) mo.off++; else mo.lastWork = dates[j5];
    }
    Object.keys(months).forEach(function (ym) {
      var mo = months[ym], total = daysInMonth(ym), work = mo.days - mo.off;
      // ngày của tháng nằm ngoài file: coi là ngày làm việc trừ Chủ nhật
      if (mo.days < total) {
        for (var d = 1; d <= total; d++) {
          var iso = ym + '-' + pad(d);
          if ((iso < dates[0] || iso > dates[nd - 1]) && weekday(iso) !== 0) work++;
        }
      }
      mo.work = work; mo.full = mo.days === total;
    });

    // Dữ liệu theo mã × tháng
    var facts = [];
    styles.forEach(function (st, idx) {
      var first = -1;
      for (var j = 0; j < nd; j++) if (st.kh[j] > 0 || st.th[j] > 0) { first = j; break; }
      var endISO = st.endTH > 0 ? addDays(dates[st.endTH], -1) : null;
      var byM = {};
      for (var j6 = 0; j6 < nd; j6++) {
        if (!(st.kh[j6] > 0 || st.th[j6] > 0)) continue;
        var m = dates[j6].slice(0, 7);
        var f = byM[m] || (byM[m] = { kh: 0, th: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, thDays: 0 });
        f.kh += st.kh[j6]; f.th += st.th[j6];
        f.pdK += aK.pd[idx][j6]; f.pdT += aT.pd[idx][j6];
        f.ldK += aK.ld[idx][j6]; f.ldT += aT.ld[idx][j6];
        if (st.th[j6] > 0) f.thDays++;
      }
      Object.keys(byM).forEach(function (m) {
        var f = byM[m], sam = st.sam || 0;
        facts.push({
          f: st.factory, m: m, line: st.line, lineRaw: st.lineRaw, group: lineGroup(st.line), brand: st.brand, brandRaw: st.brandRaw,
          item: st.item, style: st.style, po: st.po, sam: st.sam, hc: st.hc, qty: st.qty,
          kh: f.kh, th: f.th, pdK: f.pdK, pdT: f.pdT, ldK: f.ldK, ldT: f.ldT, thDays: f.thDays,
          sah: f.th * sam / 3600, sahK: f.kh * sam / 3600,
          start: first >= 0 && dates[first].slice(0, 7) === m ? 1 : 0,
          end: endISO && endISO <= updateDate && endISO.slice(0, 7) === m ? 1 : 0
        });
      });
    });

    return {
      name: opts.name || '', factory: factory, updateDate: updateDate, updateFrom: updateFrom,
      firstDate: dates[0], lastDate: dates[nd - 1], months: months,
      styleCount: styles.length, facts: facts, warnings: warnings
    };
  }

  // ---------- Gộp nhiều file ----------
  /*
   * Mỗi tháng của mỗi xưởng lấy số từ 1 file:
   *  - "Đã chốt": file update sớm nhất từ ngày làm việc cuối cùng của tháng trở đi (theo lịch nghỉ
   *    trong file; ví dụ file 28/4 chốt tháng 4 khi 29/4–1/5 nghỉ lễ). Số cuối tháng, 月末実績.
   *  - "Đang chạy" / "Kế hoạch": chưa có file chốt → file update mới nhất.
   * KH đầu tháng: file update mới nhất trước ngày 1 của tháng.
   */
  function buildDataset(files, opts) {
    var aliases = (opts && opts.aliases) || null;
    var byF = {};
    files.forEach(function (fl) { (byF[fl.factory] = byF[fl.factory] || []).push(fl); });
    var facts = [], meta = {}, monthsAll = {};
    Object.keys(byF).forEach(function (fac) {
      var list = byF[fac].slice().sort(function (a, b) { return a.updateDate < b.updateDate ? -1 : a.updateDate > b.updateDate ? 1 : 0; });
      var monthSet = {};
      list.forEach(function (fl) { fl.facts.forEach(function (x) { monthSet[x.m] = 1; }); });
      Object.keys(monthSet).sort().forEach(function (ym) {
        var start = monthStart(ym);
        // Ngày chốt: ngày làm việc cuối tháng theo lịch của file (tháng đủ ngày), không thì ngày cuối tháng
        var closeDate = function (fl) { var mo = fl.months[ym]; return mo && mo.full && mo.lastWork ? mo.lastWork : monthEnd(ym); };
        var has = function (fl) { return fl.months[ym] && fl.facts.some(function (x) { return x.m === ym; }); };
        var src = null, status;
        for (var i = 0; i < list.length; i++) if (list[i].updateDate >= closeDate(list[i]) && has(list[i])) { src = list[i]; break; }
        if (src) status = 'chot';
        else {
          for (var i2 = list.length - 1; i2 >= 0; i2--) if (has(list[i2])) { src = list[i2]; break; }
          status = src.updateDate >= start ? 'dang' : 'kh';
        }
        var open = null;
        for (var i3 = list.length - 1; i3 >= 0; i3--) if (list[i3].updateDate < start && list[i3] !== src && has(list[i3])) { open = list[i3]; break; }
        var wd = (src.months[ym] && src.months[ym].work) || 1;
        meta[fac + '|' + ym] = { f: fac, m: ym, status: status, src: src.name, srcDate: src.updateDate, open: open ? open.name : null, openDate: open ? open.updateDate : null, wd: wd };
        monthsAll[ym] = 1;
        src.facts.forEach(function (x) {
          if (x.m !== ym) return;
          facts.push(Object.assign({}, x, { brand: normBrand(x.brandRaw, aliases), wd: wd, status: status, khOpen: 0, open: false }));
        });
        if (open) open.facts.forEach(function (x) {
          if (x.m !== ym) return;
          facts.push(Object.assign({}, x, {
            brand: normBrand(x.brandRaw, aliases), wd: wd, status: status, khOpen: x.kh, open: true,
            kh: 0, th: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, thDays: 0, sah: 0, sahK: 0, start: 0, end: 0
          }));
        });
      });
    });
    return {
      facts: facts, meta: meta, months: Object.keys(monthsAll).sort(),
      factories: Object.keys(byF).sort()
    };
  }

  // ---------- Tổng hợp ----------
  var SUM_FIELDS = ['kh', 'th', 'khOpen', 'pdK', 'pdT', 'ldK', 'ldT', 'sah', 'sahK', 'start', 'end'];

  function aggregate(facts) {
    var a = { kh: 0, th: 0, khOpen: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, sah: 0, sahK: 0, start: 0, end: 0, laborSum: 0, laborSumK: 0, n: 0 };
    var styles = {}, months = {}, lines = {};
    facts.forEach(function (x) {
      SUM_FIELDS.forEach(function (k) { a[k] += x[k] || 0; });
      a.laborSum += (x.pdT || 0) / (x.wd || 1);
      a.laborSumK += (x.pdK || 0) / (x.wd || 1);
      if (!x.open && (x.kh > 0 || x.th > 0)) { styles[x.f + '|' + x.style] = 1; months[x.m] = 1; }
      if (!x.open && x.th > 0) lines[x.f + '|' + x.line] = 1;
      a.n++;
    });
    a.styles = Object.keys(styles).length;
    a.lines = Object.keys(lines).length;
    a.monthCount = Object.keys(months).length;
    return a;
  }

  // Chỉ tiêu suy ra; nMonths = số tháng của kỳ (để tính lao động bình quân/ngày)
  function derive(a, nMonths) {
    var n = nMonths || a.monthCount || 1;
    return Object.assign({}, a, {
      rate: a.kh > 0 ? a.th / a.kh : null,
      rateOpen: a.khOpen > 0 ? a.th / a.khOpen : null,
      labor: a.pdT > 0 ? a.laborSum / n : null,          // lao động bình quân/ngày (người)
      laborK: a.pdK > 0 ? a.laborSumK / n : null,
      perLine: a.ldT > 0 ? a.pdT / a.ldT : null,         // người/chuyền bình quân
      productivity: a.pdT > 0 ? a.th / a.pdT : null,     // sp/người/ngày
      efficiency: a.pdT > 0 ? a.sah / (a.pdT * HOURS_PER_DAY) : null, // giờ chuẩn / giờ công
      avgSam: a.th > 0 ? a.sah * 3600 / a.th : null
    });
  }

  function filterFacts(facts, opts) {
    return facts.filter(function (x) {
      if (opts.factory && x.f !== opts.factory) return false;
      if (opts.brand && x.brand !== opts.brand) return false;
      if (opts.months && opts.months.indexOf(x.m) < 0) return false;
      return true;
    });
  }

  function groupBy(facts, keyFn) {
    var map = {};
    facts.forEach(function (x) { var k = keyFn(x); (map[k] = map[k] || []).push(x); });
    return map;
  }

  var DIMS = {
    brand: { label: 'Nhãn hàng', key: function (x) { return x.brand; } },
    line: { label: 'Chuyền', key: function (x) { return x.f + ' · ' + x.line; } },
    group: { label: 'Cụm chuyền (theo số)', key: function (x) { return x.f + ' · ' + x.group; } },
    item: { label: 'Chủng loại', key: function (x) { return x.item; } },
    factory: { label: 'Xưởng', key: function (x) { return x.f; } }
  };

  function pct(a, b) { return a != null && b ? a / b - 1 : null; }

  // Các kỳ so sánh cho tháng ym
  function periods(ym, startMonth, available) {
    var fy = fiscalYear(ym, startMonth), fyM = fiscalMonths(fy, startMonth);
    var avail = function (list) { return list.filter(function (m) { return available.indexOf(m) >= 0; }); };
    // Lũy kế cùng kỳ: đúng các tháng đã có của năm này, lùi 12 tháng
    var ytd = avail(fyM.filter(function (m) { return m <= ym; }));
    var py = ytd.map(function (m) { return addMonths(m, -12); });
    return {
      fy: fy, cur: [ym], prev: avail([addMonths(ym, -1)]), yoy: avail([addMonths(ym, -12)]),
      ytd: avail(ytd), ytdPrev: avail(py),
      fyAll: avail(fyM), fyPrevAll: avail(fiscalMonths(fy - 1, startMonth))
    };
  }

  // Số tháng có số thực hiện trong danh sách (mẫu số của lao động bình quân/ngày)
  function actualMonths(facts, months) {
    var set = {};
    facts.forEach(function (x) { if (!x.open && x.th > 0 && months.indexOf(x.m) >= 0) set[x.m] = 1; });
    return Object.keys(set).length || months.length || 1;
  }

  /*
   * opts: { month, factory, brand, startMonth, dim }
   * Trả về chỉ tiêu tổng của từng kỳ và bảng theo chiều dim.
   */
  function summarize(ds, opts) {
    var startMonth = opts.startMonth || 4;
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand });
    var P = periods(opts.month, startMonth, ds.months);
    var calc = function (months, facts) {
      var fs = filterFacts(facts || base, { months: months });
      return derive(aggregate(fs), actualMonths(base, months));
    };
    var totals = {};
    Object.keys(P).forEach(function (k) { if (Array.isArray(P[k])) totals[k] = calc(P[k]); });

    var rows = [];
    if (opts.dim && DIMS[opts.dim]) {
      var keyFn = DIMS[opts.dim].key;
      var allMonths = [].concat(P.cur, P.prev, P.yoy, P.ytd, P.ytdPrev);
      var groups = groupBy(filterFacts(base, { months: allMonths }), keyFn);
      Object.keys(groups).forEach(function (k) {
        var g = groups[k], row = { key: k };
        ['cur', 'prev', 'yoy', 'ytd', 'ytdPrev'].forEach(function (p) {
          row[p] = derive(aggregate(filterFacts(g, { months: P[p] })), actualMonths(base, P[p]));
        });
        row.share = totals.cur.th > 0 ? row.cur.th / totals.cur.th : null;
        if (row.cur.th || row.cur.kh || row.prev.th || row.yoy.th || row.ytd.th || row.ytd.kh || row.ytdPrev.th) rows.push(row);
      });
      rows.sort(function (a, b) { return (b.cur.th - a.cur.th) || (b.ytd.th - a.ytd.th) || (a.key < b.key ? -1 : 1); });
    }
    return { periods: P, totals: totals, rows: rows };
  }

  // Bảng theo tháng (mọi tháng có dữ liệu)
  function monthly(ds, opts) {
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand });
    var g = groupBy(base, function (x) { return x.m; });
    return ds.months.map(function (m) {
      var st = ds.factories.filter(function (f) { return !opts.factory || f === opts.factory; })
        .map(function (f) { return ds.meta[f + '|' + m]; }).filter(Boolean);
      var status = st.length ? (st.every(function (s) { return s.status === 'chot'; }) ? 'chot' : st.some(function (s) { return s.status === 'dang'; }) ? 'dang' : st.every(function (s) { return s.status === 'kh'; }) ? 'kh' : 'dang') : 'kh';
      return { month: m, fy: fiscalYear(m, opts.startMonth || 4), status: status, v: derive(aggregate(g[m] || []), 1) };
    });
  }

  // Bảng theo năm tài chính; dim (tùy chọn) → ma trận chiều × năm
  /*
   * Bảng theo năm tài chính. Mỗi năm có thêm "cùng kỳ năm trước": đúng các tháng đã có TH
   * của năm đó, lùi 12 tháng (để năm đang chạy so được với năm trước).
   * dim (tùy chọn) → ma trận chiều × năm, so năm gần nhất có TH với cùng kỳ năm trước.
   */
  function fiscal(ds, opts) {
    var startMonth = opts.startMonth || 4;
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand });
    var thMonths = {};
    base.forEach(function (x) { if (!x.open && x.th > 0) thMonths[x.m] = 1; });
    var years = {};
    ds.months.forEach(function (m) { years[fiscalYear(m, startMonth)] = 1; });
    var list = Object.keys(years).map(Number).sort();
    var calc = function (facts, months) { return derive(aggregate(filterFacts(facts, { months: months })), actualMonths(base, months)); };
    var out = list.map(function (fy) {
      var months = fiscalMonths(fy, startMonth).filter(function (m) { return ds.months.indexOf(m) >= 0; });
      var withTh = months.filter(function (m) { return thMonths[m]; });
      var same = withTh.map(function (m) { return addMonths(m, -12); }).filter(function (m) { return thMonths[m]; });
      return {
        fy: fy, label: fiscalLabel(fy, startMonth), months: months, thMonths: withTh, sameMonths: same,
        v: calc(base, months), done: calc(base, withTh), same: same.length ? calc(base, same) : null
      };
    });
    var lastTh = -1;
    out.forEach(function (y, i) { if (y.v.th > 0) lastTh = i; });
    var matrix = null;
    if (opts.dim && DIMS[opts.dim]) {
      var groups = groupBy(base, DIMS[opts.dim].key);
      var ly = lastTh >= 0 ? out[lastTh] : null;
      matrix = Object.keys(groups).map(function (k) {
        var cells = out.map(function (y) { return calc(groups[k], y.months); });
        var same = ly && ly.same ? calc(groups[k], ly.sameMonths) : null;
        return { key: k, cells: cells, last: ly ? cells[lastTh] : null, same: same };
      }).filter(function (r) { return r.cells.some(function (c) { return c.th > 0; }); });
      matrix.sort(function (a, b) {
        var x = a.last ? a.last.th : 0, y = b.last ? b.last.th : 0;
        return (y - x) || (a.key < b.key ? -1 : 1);
      });
    }
    return { years: out, lastTh: lastTh, matrix: matrix };
  }


  return {
    HOURS_PER_DAY: HOURS_PER_DAY, BRAND_ALIASES: BRAND_ALIASES, DIMS: DIMS,
    toISO: toISO, serialToISO: serialToISO, addDays: addDays, addMonths: addMonths, monthEnd: monthEnd, monthLabel: monthLabel,
    fiscalYear: fiscalYear, fiscalMonths: fiscalMonths, fiscalLabel: fiscalLabel, updateDateFromName: updateDateFromName,
    parseAliasText: parseAliasText, normBrand: normBrand, lineGroup: lineGroup, baseLine: baseLine,
    parseRows: parseRows, buildDataset: buildDataset, aggregate: aggregate, derive: derive,
    filterFacts: filterFacts, groupBy: groupBy, periods: periods, summarize: summarize, monthly: monthly, fiscal: fiscal, pct: pct
  };
});
