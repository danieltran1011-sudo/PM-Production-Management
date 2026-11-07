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
  var BRAND_ALIASES = { 'ALLTERAIN': 'ALLTERRAIN', 'DESCSENTE': 'DESCENTE', 'MONBELL': 'MONTBELL', 'MONTBEL': 'MONTBELL', 'MONTBEEL': 'MONTBELL', 'MONBEL': 'MONTBELL', 'MONT-BELL': 'MONTBELL', 'MONT BELL': 'MONTBELL' };

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

  // ---------- Hàng của TI hay hàng APIC tự bán ----------
  // Danh sách nhãn hàng TI theo hợp đồng; nhãn ngoài danh sách là hàng APIC tự bán.
  var TI_BRANDS_DEFAULT = ['TNF(GTC企画)', 'TNF(東京企画)', 'TNF(東京企画、GTC企画共通)サンプル', 'TNF別注 富山', 'TNF別注 東京', 'NW', 'UTMF', 'BOYS/DTN',
    'Gymphlex', 'NANGA', 'DANTON', 'TETON BROS', 'ALLTERRAIN', 'AIGLE', 'ORCIVAL', 'HH', 'WOOLRICH', 'CHAMPION', '241', 'SNOWPEAK', 'LECOQ', 'SRIXON',
    'MACKINTOSH', 'MARMOT', 'DESCENTE', 'C3FIT', 'MONTBELL', 'MILLET', 'GOLDWINサンプル', 'GOLDWIN', 'GSM', 'RESURRECTION', 'CANTERBURY', 'HYKE', '141',
    'HAT +', 'LL.BEAN', 'FOXFIRE', 'TATRAS', 'KASCO', 'YONEX', 'KOTI', 'UMBRO', 'CCC', 'Inov-8', "B'S", 'GTO', 'Finisher'].join('\n');
  // Nhãn TI bổ sung sau: tự thêm vào danh sách người dùng đã lưu trước đó
  var TI_BRANDS_ADDED = ['Finisher'];
  // Cách viết khác nhau của cùng nhãn hàng (đã bỏ dấu cách, ký tự đặc biệt)
  var BRAND_KEY_ALIASES = { HELLYHANSEN: 'HH', NEUTRALWORK: 'NW', NEUTRALWORKS: 'NW', THENORTHFACE: 'TNF', DESCSENTE: 'DESCENTE', DESCSENT: 'DESCENTE',
    WOORICH: 'WOOLRICH', ALLTERAIN: 'ALLTERRAIN', LECOQSPORTIF: 'LECOQ', SNOWPEAKS: 'SNOWPEAK',
    MONBELL: 'MONTBELL', MONTBEL: 'MONTBELL', MONTBEEL: 'MONTBELL', MONBEL: 'MONTBELL' };
  function brandKey(name) {
    var t = norm(name).toUpperCase();
    t = t.split('(')[0].split('（')[0].replace(/別注.*$/, '').replace(/サンプル.*$/, '').replace(/企画.*$/, '');
    var k = t.replace(/[^A-Z0-9]/g, '');
    return BRAND_KEY_ALIASES[k] || k;
  }
  function tiMatcher(text) {
    var keys = {};
    String(text == null ? TI_BRANDS_DEFAULT : text).split(/\r?\n/).forEach(function (l) { var k = brandKey(l); if (k) keys[k] = 1; });
    var list = Object.keys(keys), cache = {};
    return function (brandRaw) {
      if (brandRaw in cache) return cache[brandRaw];
      var k = brandKey(brandRaw), ok = !!keys[k];
      // ALLTERRAIN81, YONEX WC… → bắt đầu bằng tên trong danh sách (tên từ 3 ký tự)
      if (!ok && k) ok = list.some(function (x) { return x.length >= 3 && k.indexOf(x) === 0; });
      return (cache[brandRaw] = ok ? 'TI' : 'APIC');
    };
  }

  // ---------- Đọc sheet ----------
  var HEADERS = {
    factory: ['工場'], line: ['ライン'], brand: ['ブランド'], style: ['品番'], item: ['アイテム'],
    po: ['PO'], lot: ['集約ロット(pcs)'], aggLot: ['集約ロット'], qty: ['生産数(pcs)'], sam: ['工程分析時間(sec)'],
    hc: ['ライン編成人数(人/line)']
  };
  // Vị trí cột của mẫu file hiện tại (T..AC), dùng khi không tìm thấy tiêu đề
  var DEFAULT_COLS = { factory: 19, line: 20, brand: 21, style: 22, item: 23, po: 24, lot: 25, qty: 26, sam: 27, hc: 28, aggLot: 37, learnTheo: 41, ltTheo: 45 };
  var LATE_HEADERS = { learnTheo: ['習熟日数(day)'], ltTheo: ['縫製リードタイム(day)'] };

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
      // Cột lý thuyết nằm sau cột 集約ロット (AL): 習熟日数 (AP), 縫製リードタイム (AT, hàng "dự định")
      Object.keys(LATE_HEADERS).forEach(function (k) {
        for (var c = (cols.aggLot || 0) + 1; c < row.length; c++) {
          if (LATE_HEADERS[k].indexOf(norm(row[c])) >= 0) { cols[k] = c; return; }
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
        aggLot: numOrNull(get('aggLot')),
        learnTheo: numOrNull(pr[cols.learnTheo]), ltTheo: numOrNull(pr[cols.ltTheo]),
        kh: new Array(nd), th: new Array(nd), theo: new Array(nd), endTH: -1
      };
      // Hàng lý thuyết ("dự định") ngay trên hàng KH: lịch may theo năng lực lý thuyết
      var hasTheo = rowType(pr[typeCol]) === '' && norm(pr[typeCol]) !== '';
      for (var j3 = 0; j3 < nd; j3++) {
        s.kh[j3] = Math.max(0, num(kr[d0 + j3]));
        s.theo[j3] = hasTheo ? Math.max(0, num(pr[d0 + j3])) : 0;
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
        if (!(st.kh[j6] > 0 || st.th[j6] > 0 || st.theo[j6] > 0)) continue;
        var m = dates[j6].slice(0, 7);
        var f = byM[m] || (byM[m] = { kh: 0, th: 0, theo: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, thDays: 0 });
        f.kh += st.kh[j6]; f.th += st.th[j6]; f.theo += st.theo[j6];
        f.pdK += aK.pd[idx][j6]; f.pdT += aT.pd[idx][j6];
        f.ldK += aK.ld[idx][j6]; f.ldT += aT.ld[idx][j6];
        if (st.th[j6] > 0) f.thDays++;
      }
      Object.keys(byM).forEach(function (m) {
        var f = byM[m], sam = st.sam || 0;
        facts.push({
          f: st.factory, m: m, line: st.line, lineRaw: st.lineRaw, group: lineGroup(st.line), brand: st.brand, brandRaw: st.brandRaw,
          item: st.item, style: st.style, po: st.po, sam: st.sam, hc: st.hc, qty: st.qty, aggLot: st.aggLot,
          theo: f.theo > 0 ? 1 : 0,
          kh: f.kh, th: f.th, pdK: f.pdK, pdT: f.pdT, ldK: f.ldK, ldT: f.ldT, thDays: f.thDays,
          sah: f.th * sam / 3600, sahK: f.kh * sam / 3600,
          start: first >= 0 && dates[first].slice(0, 7) === m ? 1 : 0,
          end: endISO && endISO <= updateDate && endISO.slice(0, 7) === m ? 1 : 0
        });
      });
    });

    // Chuyền vật lý theo ngày (nửa chuyền A/B, như cách TI tính 29 chuyền): nhãn hàng nào chiếm chuyền
    var lineDays = buildLineDays(styles, nd);

    // Từng lượt may (1 mã trên 1 chuyền): lên năng suất, lead time thực tế
    var runs = [];
    styles.forEach(function (st) {
      var r = runStats(st, dates, off, updateDate);
      if (r) runs.push(r);
    });

    return {
      v: DATA_VERSION, name: opts.name || '', factory: factory, updateDate: updateDate, updateFrom: updateFrom,
      firstDate: dates[0], lastDate: dates[nd - 1], months: months,
      styleCount: styles.length, facts: facts, runs: runs, warnings: warnings,
      lineDays: { off: off.map(function (o) { return o ? '1' : '0'; }).join(''), brands: lineDays.brands, lines: lineDays.lines }
    };
  }

  // ---------- Chuyền vật lý (nửa chuyền A/B) theo ngày ----------
  /*
   * Chuyền trong file: 4A, 4B (nửa chuyền, ~30 người), 4AB (ghép cả 2 nửa), 4A1/4A2 (nửa của 4A, ~16 người).
   * Mỗi ngày, mỗi nửa chuyền A/B có mức chiếm từ 0 (trống) đến 1: có mã chạy ở chuyền chính (4A, 4AB) → 1;
   * chỉ có chuyền tách → 0,5 mỗi chuyền tách (tối đa 1). Phần chiếm được chia cho nhãn hàng theo sản lượng ngày đó.
   */
  function physicalLines(line) {
    var m = norm(line).toUpperCase().match(/^(\d+)(AB|A|B)(\d*)$/);
    if (!m) return [];
    var sub = !!m[3], halves = m[2] === 'AB' ? ['A', 'B'] : [m[2]];
    return halves.map(function (h) { return { line: m[1] + h, sub: sub, raw: norm(line).toUpperCase() }; });
  }
  function buildLineDays(styles, nd) {
    var brands = [], bIdx = {}, lines = {};
    var bi = function (b) { if (!(b in bIdx)) { bIdx[b] = brands.length; brands.push(b); } return bIdx[b]; };
    for (var j = 0; j < nd; j++) {
      var day = {};
      styles.forEach(function (st) {
        if (!(st.th[j] > 0)) return;
        physicalLines(st.lineRaw).forEach(function (pl) {
          var d = day[pl.line] || (day[pl.line] = { main: false, subs: {}, by: {}, tot: 0 });
          if (pl.sub) d.subs[pl.raw] = 1; else d.main = true;
          var k = bi(st.brandRaw || '');
          d.by[k] = (d.by[k] || 0) + st.th[j]; d.tot += st.th[j];
        });
      });
      Object.keys(day).forEach(function (pl) {
        var d = day[pl], occ = d.main ? 1 : Math.min(1, 0.5 * Object.keys(d.subs).length);
        var parts = Object.keys(d.by).map(function (k) { return k + ':' + (Math.round(occ * d.by[k] / d.tot * 100) / 100); });
        (lines[pl] = lines[pl] || {})[j] = parts.join(',');
      });
    }
    return { brands: brands, lines: lines };
  }

  // ---------- Lượt may: lên năng suất (習熟), lead time ----------
  var DATA_VERSION = 4; // tăng khi file đã nạp cần nạp lại để có đủ chỉ tiêu
  var CURVE_DAYS = [1, 2, 3, 5, 7, 10];
  var RAMP_LEVEL = 0.9; // đạt 90% sản lượng ổn định là xong giai đoạn lên năng suất
  var RAMP_MAX_DAYS = 30;

  /*
   * Sản lượng ổn định = phân vị 75% của sản lượng các ngày (bỏ ngày cuối, thường làm dở),
   * để vài ngày đột biến không kéo mức ổn định lên.
   * Ngày lên năng suất = thứ tự ngày sản xuất đầu tiên đạt 90% mức ổn định, xét trong 30 ngày đầu
   * (không đạt thì để trống).
   * Giờ chuẩn mất = tổng (ổn định − thực tế) × thời gian may của các ngày trước đó.
   * Lead time thực tế = số ngày làm việc từ ngày may đầu tiên đến ngày cuối cùng.
   */
  function runStats(st, dates, off, updateDate) {
    var idx = [];
    for (var j = 0; j < st.th.length; j++) if (st.th[j] > 0) idx.push(j);
    if (!idx.length) return null;
    var first = idx[0], last = idx[idx.length - 1];
    var vals = idx.map(function (k) { return st.th[k]; });
    var body = vals.length >= 3 ? vals.slice(0, -1) : vals.slice();
    var sorted = body.slice().sort(function (a, b) { return a - b; });
    var stable = sorted[Math.min(sorted.length - 1, Math.floor(0.75 * (sorted.length - 1) + 0.5))];
    var ramp = null;
    for (var k = 0; k < Math.min(vals.length, RAMP_MAX_DAYS); k++) if (vals[k] >= RAMP_LEVEL * stable) { ramp = k + 1; break; }
    var loss = 0, sam = st.sam || 0;
    if (ramp) for (var k2 = 0; k2 < ramp - 1; k2++) loss += Math.max(0, stable - vals[k2]) * sam / 3600;
    var span = 0;
    for (var j2 = first; j2 <= last; j2++) if (!off[j2]) span++;
    var endISO = st.endTH > 0 ? addDays(dates[st.endTH], -1) : null;
    var total = vals.reduce(function (a, b) { return a + b; }, 0);
    return {
      f: st.factory, line: st.line, lineRaw: st.lineRaw, group: lineGroup(st.line), brand: st.brand, brandRaw: st.brandRaw,
      item: st.item, style: st.style, sam: st.sam, hc: st.hc, qty: st.qty, aggLot: st.aggLot,
      learnTheo: st.learnTheo, ltTheo: st.ltTheo,
      start: dates[first], end: dates[last], days: vals.length, span: span, th: total,
      stable: stable, ramp: ramp, loss: loss,
      curve: CURVE_DAYS.map(function (d) { return vals.length >= d ? Math.min(1.5, vals[d - 1] / stable) : null; }),
      complete: !!endISO && endISO <= updateDate,
      truncated: first < 3 // bắt đầu ngay đầu dữ liệu của file: có thể đã may từ trước
    };
  }

  // ---------- Kiểm tra công thức tổng của file (sheet TORAY lấy từ XN1AB hàng 1–3) ----------
  // Ô tổng và chỉ tiêu tương ứng ở sheet TORAY
  var SUMMARY_CELLS = [
    { cell: 'C2', key: 'kh', label: 'KH (計画)' },
    { cell: 'C3', key: 'th', label: 'TH (実績)' },
    { cell: 'M1', key: 'changes', label: 'Số lần chuyển mã (切替数)' },
    { cell: 'F1', key: 'lot', label: 'Lot trung bình (平均ロット)' },
    { cell: 'H1', key: 'people', label: 'LĐ bình quân/chuyền (平均ライン人数)' },
    { cell: 'G1', key: 'sam', label: 'Thời gian may bình quân (平均工程分析時間)' }
  ];
  /*
   * rows: mảng 2 chiều sheet XN1AB (giá trị đã lưu); formulas: { 'C2': 'SUMIF(...)', ... } (mọi ô A1:R3);
   * errors: { 'K2': '#VALUE!' } các ô đang ra lỗi.
   * - Công thức có vùng dữ liệu (bắt đầu từ hàng 7–13) kết thúc trước hàng dữ liệu cuối → tính thiếu.
   * - SUMPRODUCT có các vùng khác số hàng → #VALUE!.
   * Tính lại các chỉ tiêu chính với đủ mọi hàng (cùng cách tính của công thức trong file) để so sánh.
   */
  function checkSummary(rows, formulas, errors) {
    formulas = formulas || {}; errors = errors || {};
    var t = function (r) { return norm(r && r[1]); };
    var last = 0;
    rows.forEach(function (r, i) { var x = t(r); if (x === 'KH' || x === 'TH' || x === 'dự định') last = i + 1; });
    var issues = [], others = [], byCell = {};
    SUMMARY_CELLS.forEach(function (sc) { byCell[sc.cell] = sc; });
    var cellOrder = function (a, b) {
      var ra = +a.replace(/[A-Z]+/, ''), rb = +b.replace(/[A-Z]+/, '');
      return ra - rb || a.replace(/\d+/, '').localeCompare(b.replace(/\d+/, ''));
    };
    Object.keys(formulas).sort(cellOrder).forEach(function (cell) {
      var f = formulas[cell], ends = [], all = [], re = /\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)/g, m;
      while ((m = re.exec(f))) {
        if (!(+m[2] >= 7 && +m[2] <= 13)) continue;
        all.push(m[0].replace(/\$/g, ''));
        // vùng tiêu chí cột B quyết định số hàng (Excel tự kéo dài vùng cộng của SUMIF)
        if (m[1] === 'B' || !/SUMIF|AVERAGEIF/.test(f)) ends.push(+m[4]);
      }
      var sc = byCell[cell], problems = [], item = { cell: cell, label: sc ? sc.label : '', end: null, mismatch: null, error: errors[cell] || null };
      var end = ends.length ? Math.min.apply(null, ends) : null;
      // các số hàng cuối ngắn hơn dữ liệu (để thay bằng Ctrl+H)
      var shortEnds = {};
      all.forEach(function (r) { var e = +r.split(':')[1].replace(/[A-Z]+/, ''); if (e < last) shortEnds[e] = 1; });
      item.shortEnds = Object.keys(shortEnds).map(Number).sort(function (a, b) { return a - b; });
      if (end != null && end < last) { item.end = end; problems.push('dừng ở hàng ' + end); }
      if (/SUMPRODUCT/.test(f)) {
        var sizes = {};
        all.forEach(function (r) { var p = r.split(':'); sizes[+p[1].replace(/[A-Z]+/, '') - +p[0].replace(/[A-Z]+/, '')] = 1; });
        if (Object.keys(sizes).length > 1) { item.mismatch = all.join(', '); problems.push('vùng không cùng số hàng (' + item.mismatch + ') → #VALUE!'); }
      }
      if (item.error) problems.push('đang ra lỗi ' + item.error);
      if (!problems.length) return;
      if (sc && end != null && end < last) issues.push({ cell: cell, key: sc.key, label: sc.label, end: end });
      item.text = problems.join('; ');
      others.push(item);
    });
    Object.keys(errors).forEach(function (cell) {
      if (!formulas[cell]) others.push({ cell: cell, label: '', end: null, mismatch: null, error: errors[cell], shortEnds: [], text: 'đang ra lỗi ' + errors[cell] });
    });
    // Tính lại theo cột tóm tắt từng hàng: C sản lượng tháng chọn, D cờ TH, E SL, F lot, G giây, H người, M chuyển mã
    var r1 = rows[0] || [], d6 = norm((rows[5] || [])[3]) || 'TH', b1 = norm(r1[1]) || 'dự định';
    var calc = function (endRow) {
      var a = { kh: 0, th: 0, changes: 0, fs: 0, fn: 0, gn: 0, gd: 0, hn: 0, hd: 0 };
      for (var i = 12; i < Math.min(endRow, rows.length); i++) {
        var r = rows[i]; if (!r) continue;
        var b = t(r), c = num(r[2]), d = num(r[3]), e = num(r[4]), fv = numOrNull(r[5]), g = num(r[6]), h = num(r[7]);
        if (b === b1 && fv != null) { a.fs += fv; a.fn++; }
        a.gn += e * g * d; a.hn += (fv || 0) * h * d;
        if (b === d6) { a.gd += e; a.hd += fv || 0; a.changes += num(r[12]); }
        if (b === 'KH') a.kh += c;
        if (b === 'TH') a.th += c;
      }
      return { kh: a.kh, th: a.th, changes: a.changes, lot: a.fn ? a.fs / a.fn : null, sam: a.gd ? a.gn / a.gd : null, people: a.hd ? a.hn / a.hd : null };
    };
    var y = (rows[4] || [])[1], mo = (rows[5] || [])[1];
    return {
      lastRow: last, issues: issues, cells: others,
      month: typeof y === 'number' && typeof mo === 'number' ? y + '-' + pad(mo) : null,
      file: { kh: num((rows[1] || [])[2]), th: num((rows[2] || [])[2]), changes: num(r1[12]), lot: numOrNull(r1[5]), sam: numOrNull(r1[6]), people: numOrNull(r1[7]) },
      full: calc(last)
    };
  }

  // ---------- Gộp nhiều file ----------
  /*
   * Mỗi tháng của mỗi xưởng lấy số từ 1 file:
   *  - "Đã chốt": file update sớm nhất từ ngày làm việc cuối cùng của tháng trở đi (theo lịch nghỉ
   *    trong file; ví dụ file 28/4 chốt tháng 4 khi 29/4–1/5 nghỉ lễ). Số cuối tháng, 月末実績.
   *  - "Đang chạy" / "Kế hoạch": chưa có file chốt → file update mới nhất.
   * KH đầu tháng: file update mới nhất trước ngày 1 của tháng, trong vòng 10 ngày (file cuối tháng trước).
   */
  var OPEN_MAX_DAYS = 10;
  function buildDataset(files, opts) {
    var aliases = (opts && opts.aliases) || null;
    var ownerOf = tiMatcher(opts && opts.tiBrands != null ? opts.tiBrands : null);
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
        // File phải có dữ liệu từ đầu tháng (tháng ở đầu dữ liệu của file chỉ có vài ngày thì không dùng)
        var has = function (fl) { return fl.months[ym] && fl.firstDate <= start && fl.facts.some(function (x) { return x.m === ym; }); };
        if (!list.some(has)) return; // không có file nào đủ tháng này
        var src = null, status;
        for (var i = 0; i < list.length; i++) if (list[i].updateDate >= closeDate(list[i]) && has(list[i])) { src = list[i]; break; }
        if (src) status = 'chot';
        else {
          for (var i2 = list.length - 1; i2 >= 0; i2--) if (has(list[i2])) { src = list[i2]; break; }
          status = src.updateDate >= start ? 'dang' : 'kh';
        }
        var open = null;
        // KH đầu tháng: file update trong 10 ngày trước ngày 1 (file cuối tháng trước); file cũ hơn thì không dùng
        var openFrom = addDays(start, -OPEN_MAX_DAYS);
        for (var i3 = list.length - 1; i3 >= 0; i3--) {
          var cand = list[i3];
          if (cand.updateDate < start && cand.updateDate >= openFrom && cand !== src && has(cand)) { open = cand; break; }
        }
        var wd = (src.months[ym] && src.months[ym].work) || 1;
        meta[fac + '|' + ym] = { f: fac, m: ym, status: status, src: src.name, srcDate: src.updateDate, open: open ? open.name : null, openDate: open ? open.updateDate : null, wd: wd };
        monthsAll[ym] = 1;
        src.facts.forEach(function (x) {
          if (x.m !== ym) return;
          facts.push(Object.assign({}, x, { brand: normBrand(x.brandRaw, aliases), owner: ownerOf(x.brandRaw), wd: wd, status: status, khOpen: 0, open: false }));
        });
        if (open) open.facts.forEach(function (x) {
          if (x.m !== ym) return;
          facts.push(Object.assign({}, x, {
            brand: normBrand(x.brandRaw, aliases), owner: ownerOf(x.brandRaw), wd: wd, status: status, khOpen: x.kh, open: true,
            kh: 0, th: 0, theo: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, thDays: 0, sah: 0, sahK: 0, start: 0, end: 0
          }));
        });
      });
    });
    // Lượt may: cùng xưởng, mã, chuyền, ngày bắt đầu → lấy bản trong file update mới nhất
    var runMap = {};
    files.slice().sort(function (a, b) { return a.updateDate < b.updateDate ? -1 : 1; }).forEach(function (fl) {
      (fl.runs || []).forEach(function (r) {
        if (r.truncated) return;
        runMap[r.f + '|' + r.style + '|' + r.lineRaw + '|' + r.start] = Object.assign({}, r, { brand: normBrand(r.brandRaw, aliases), owner: ownerOf(r.brandRaw) });
      });
    });
    var runs = Object.keys(runMap).map(function (k) { return runMap[k]; });
    // Mã lặp lại: đã có lượt may của cùng mã kết thúc trước ngày bắt đầu ít nhất 14 ngày
    var byStyle = {};
    runs.forEach(function (r) { (byStyle[r.style] = byStyle[r.style] || []).push(r); });
    runs.forEach(function (r) {
      var before = addDays(r.start, -14);
      r.repeat = byStyle[r.style].some(function (o) { return o !== r && o.end < before; });
    });
    // Chuyền-ngày: gộp các xưởng (số chuyền là chung), mỗi tháng lấy từ file chốt tháng đó của từng xưởng
    var lineDay = {}, dayOff = {}, lineMonthOk = {};
    // Tháng dùng được cho thống kê chuyền: mọi xưởng đều đã chốt (chuyền là chung giữa các xưởng)
    var facAll = Object.keys(byF);
    Object.keys(monthsAll).forEach(function (m) {
      lineMonthOk[m] = facAll.every(function (f) { return meta[f + '|' + m] && meta[f + '|' + m].status === 'chot'; });
    });
    var lineMonthAny = {};
    Object.keys(meta).forEach(function (k) {
      var me = meta[k];
      if (me.status !== 'chot') return;
      lineMonthAny[me.m] = true;
      var src = files.filter(function (fl) { return fl.factory === me.f && fl.updateDate === me.srcDate; })[0];
      if (!src || !src.lineDays) return;
      var L = src.lineDays;
      for (var j = 0; j < L.off.length; j++) {
        var iso = addDays(src.firstDate, j);
        if (iso.slice(0, 7) !== me.m) continue;
        var isOff = L.off.charAt(j) === '1';
        dayOff[iso] = iso in dayOff ? dayOff[iso] && isOff : isOff;
        Object.keys(L.lines).forEach(function (pl) {
          var e = L.lines[pl][j];
          if (!e) return;
          var d = lineDay[iso] || (lineDay[iso] = {}), slot = d[pl] || (d[pl] = {});
          e.split(',').forEach(function (part) {
            var q = part.split(':'), b = L.brands[+q[0]];
            slot[b] = (slot[b] || 0) + (+q[1]);
          });
        });
      }
    });
    var lineLife = computeLineLife(lineDay, dayOff, opts && opts.lineOverrides);
    return {
      facts: facts, meta: meta, months: Object.keys(monthsAll).sort(), lineLife: lineLife,
      factories: Object.keys(byF).sort(), runs: runs, lineDay: lineDay, dayOff: dayOff, lineMonthOk: lineMonthOk, lineMonthAny: lineMonthAny, ownerOf: ownerOf,
      oldFiles: files.filter(function (fl) { return (fl.v || 0) < DATA_VERSION; }).length,
      oldFileList: files.filter(function (fl) { return (fl.v || 0) < DATA_VERSION; }).map(function (fl) { return { factory: fl.factory, updateDate: fl.updateDate, name: fl.name }; })
    };
  }

  // ---------- Kiểm tra thiếu file cuối tháng ----------
  // Ngày chốt của tháng theo lịch của file: ngày làm việc cuối tháng (tháng đủ ngày), không thì ngày cuối tháng
  function closeDateOf(fl, ym) { var mo = fl.months && fl.months[ym]; return mo && mo.full && mo.lastWork ? mo.lastWork : monthEnd(ym); }
  var CLOSE_GRACE_DAYS = 10; // file update đầu tháng sau (đến ngày 10) vẫn là file cuối tháng trước

  // Tháng mà file này là "file cuối tháng", hoặc null nếu là file giữa tháng.
  // Lịch nghỉ trong các file không phải lúc nào cũng đánh dấu đủ ngày lễ (30/4, 31/8…), nên file
  // update trong 3 ngày cuối tháng cũng được tính là file cuối tháng.
  var CLOSE_EARLY_DAYS = 3;
  function closesMonth(fl) {
    var u = fl.updateDate, ym = u.slice(0, 7);
    if (u >= closeDateOf(fl, ym) || u >= addDays(monthEnd(ym), -CLOSE_EARLY_DAYS)) return ym;
    var prev = addMonths(ym, -1);
    if (u <= addDays(monthEnd(prev), CLOSE_GRACE_DAYS)) return prev;
    return null;
  }

  /*
   * Mỗi xưởng: các tháng từ tháng chốt đầu tiên đến tháng chốt gần nhất (của mọi xưởng)
   * mà không có file cuối tháng. ds (tùy chọn) để biết tháng thiếu đã có số chốt từ file khác chưa.
   */
  function coverage(files, ds) {
    var byF = {}, due = null;
    files.forEach(function (fl) {
      var c = closesMonth(fl);
      var g = byF[fl.factory] || (byF[fl.factory] = { factory: fl.factory, closed: {}, mid: [] });
      if (c) { g.closed[c] = true; if (!due || c > due) due = c; }
      else g.mid.push(fl.updateDate);
    });
    return Object.keys(byF).sort().map(function (fac) {
      var g = byF[fac], months = Object.keys(g.closed).sort(), missing = [];
      if (!months.length) return { factory: fac, first: null, last: null, expected: 0, have: 0, missing: [], mid: g.mid.sort() };
      for (var m = months[0]; m <= due; m = addMonths(m, 1)) {
        if (g.closed[m]) continue;
        var meta = ds && ds.meta[fac + '|' + m];
        missing.push({ m: m, status: meta ? meta.status : null, srcDate: meta ? meta.srcDate : null });
      }
      var expected = 0;
      for (var m2 = months[0]; m2 <= due; m2 = addMonths(m2, 1)) expected++;
      return { factory: fac, first: months[0], last: due, expected: expected, have: expected - missing.length, missing: missing, mid: g.mid.sort() };
    });
  }

  // ---------- Tổng hợp ----------
  var SUM_FIELDS = ['kh', 'th', 'khOpen', 'pdK', 'pdT', 'ldK', 'ldT', 'sah', 'sahK', 'start', 'end'];

  function aggregate(facts) {
    var a = { kh: 0, th: 0, khOpen: 0, pdK: 0, pdT: 0, ldK: 0, ldT: 0, sah: 0, sahK: 0, start: 0, end: 0, laborSum: 0, laborSumK: 0, n: 0,
      lotSum: 0, lotN: 0, samQty: 0, qtyW: 0 };
    var styles = {}, months = {}, lines = {}, openKeys = {};
    // Tháng có KH đầu tháng (để tỷ lệ đạt so KH đầu tháng chỉ tính trên các tháng đó)
    facts.forEach(function (x) { if (x.open && x.khOpen > 0) openKeys[x.f + '|' + x.m] = 1; });
    a.thOpen = 0;
    facts.forEach(function (x) {
      SUM_FIELDS.forEach(function (k) { a[k] += x[k] || 0; });
      if (!x.open && openKeys[x.f + '|' + x.m]) a.thOpen += x.th || 0;
      a.laborSum += (x.pdT || 0) / (x.wd || 1);
      a.laborSumK += (x.pdK || 0) / (x.wd || 1);
      if (!x.open && (x.kh > 0 || x.th > 0)) { styles[x.f + '|' + x.style] = 1; months[x.m] = 1; }
      if (!x.open && x.th > 0) lines[x.f + '|' + x.line] = 1;
      // 生産ロット bình quân (như TORAY 平均ロット): lot sau gộp đơn (集約ロット) của các mã có lịch lý thuyết trong tháng
      if (!x.open && x.theo && x.aggLot != null) { a.lotSum += x.aggLot; a.lotN++; }
      // 工程分析時間 bình quân (như TORAY): giây/sp, trọng số theo số lượng đơn (生産数) của các mã có TH trong tháng
      if (!x.open && x.th > 0 && x.qty > 0 && x.sam != null) { a.samQty += x.qty * x.sam; a.qtyW += x.qty; }
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
      rateOpen: a.khOpen > 0 ? a.thOpen / a.khOpen : null,
      labor: a.pdT > 0 ? a.laborSum / n : null,          // lao động bình quân/ngày (người)
      laborK: a.pdK > 0 ? a.laborSumK / n : null,
      perLine: a.ldT > 0 ? a.pdT / a.ldT : null,         // người/chuyền bình quân
      productivity: a.pdT > 0 ? a.th / a.pdT : null,     // sp/người/ngày
      efficiency: a.pdT > 0 ? a.sah / (a.pdT * HOURS_PER_DAY) : null, // giờ chuẩn / giờ công
      avgSam: a.qtyW > 0 ? a.samQty / a.qtyW : null,       // 工程分析時間 bình quân (giây)
      avgLot: a.lotN > 0 ? a.lotSum / a.lotN : null          // 生産ロット bình quân (枚)
    });
  }

  function filterFacts(facts, opts) {
    return facts.filter(function (x) {
      if (opts.factory && x.f !== opts.factory) return false;
      if (opts.brand && x.brand !== opts.brand) return false;
      if (opts.owner && x.owner !== opts.owner) return false;
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
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner });
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
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner });
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
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner });
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


  // ---------- Phân tích ----------
  var SAM_BANDS = [
    { id: 0, label: 'Dưới 8.000 giây' }, { id: 1, label: '8.000–12.000 giây' }, { id: 2, label: 'Trên 12.000 giây' }
  ];
  function samBand(sam) { if (!(sam > 0)) return null; return sam < 8000 ? 0 : sam <= 12000 ? 1 : 2; }
  var LOT_BANDS = [
    { id: 0, label: 'Dưới 500 sp', max: 500 }, { id: 1, label: '500–999 sp', max: 1000 },
    { id: 2, label: '1.000–1.999 sp', max: 2000 }, { id: 3, label: 'Từ 2.000 sp', max: Infinity }
  ];
  function lotBand(lot) {
    if (!(lot > 0)) return null;
    for (var i = 0; i < LOT_BANDS.length; i++) if (lot < LOT_BANDS[i].max) return i;
    return LOT_BANDS.length - 1;
  }
  function mean(list) { return list.length ? list.reduce(function (a, b) { return a + b; }, 0) / list.length : null; }
  function median(list) {
    if (!list.length) return null;
    var a = list.slice().sort(function (x, y) { return x - y; }), m = Math.floor(a.length / 2);
    return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
  }
  function stdev(list) {
    if (list.length < 2) return null;
    var m = mean(list);
    return Math.sqrt(list.reduce(function (a, x) { return a + (x - m) * (x - m); }, 0) / (list.length - 1));
  }

  // Các tháng của kỳ phân tích: 'month' tháng chọn, 'ytd' lũy kế năm tài chính, '12m' 12 tháng gần nhất, 'all' mọi tháng đến tháng chọn
  var ANALYSIS_PERIODS = [
    { id: 'ytd', label: 'Lũy kế năm tài chính' }, { id: '12m', label: '12 tháng gần nhất' },
    { id: 'month', label: 'Tháng báo cáo' }, { id: 'all', label: 'Toàn bộ dữ liệu' }
  ];
  function periodMonths(ds, month, kind, startMonth) {
    var upto = ds.months.filter(function (m) { return m <= month; });
    if (kind === 'month') return [month];
    if (kind === 'ytd') return periods(month, startMonth || 4, ds.months).ytd;
    if (kind === '12m') { var from = addMonths(month, -11); return upto.filter(function (m) { return m >= from; }); }
    return upto;
  }
  function analysisBase(ds, opts) {
    var months = opts.months || periodMonths(ds, opts.month, opts.period || 'ytd', opts.startMonth);
    var facts = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner, months: months })
      .filter(function (x) { return !x.open; });
    var runs = (ds.runs || []).filter(function (r) {
      return (!opts.factory || r.f === opts.factory) && (!opts.brand || r.brand === opts.brand) && (!opts.owner || r.owner === opts.owner);
    });
    return { months: months, facts: facts, runs: runs };
  }

  // 1. Cỡ lot và hiệu suất: các mã có TH trong kỳ, chia theo lot sau gộp (集約ロット) và mức thời gian may
  function lotEfficiency(ds, opts) {
    var b = analysisBase(ds, opts), withTh = b.facts.filter(function (x) { return x.th > 0; });
    var n = b.months.length || 1;
    var rows = LOT_BANDS.concat([{ id: null, label: 'Không rõ lot' }]).map(function (band) {
      var g = withTh.filter(function (x) { return lotBand(x.aggLot) === band.id; });
      return { band: band, v: derive(aggregate(g), n) };
    }).filter(function (r) { return r.v.th > 0; });
    var matrix = LOT_BANDS.map(function (band) {
      return {
        band: band,
        cells: SAM_BANDS.map(function (sb) {
          var g = withTh.filter(function (x) { return lotBand(x.aggLot) === band.id && samBand(x.sam) === sb.id; });
          return derive(aggregate(g), n);
        })
      };
    });
    return { months: b.months, rows: rows, matrix: matrix, total: derive(aggregate(withTh), n) };
  }

  // Lượt may đã xong trong kỳ (ngày cuối trong kỳ), đủ ngày để đo
  function doneRuns(b, minDays) {
    var set = {};
    b.months.forEach(function (m) { set[m] = 1; });
    return b.runs.filter(function (r) { return r.complete && set[r.end.slice(0, 7)] && r.days >= (minDays || 1); });
  }

  // 2. Lên năng suất (習熟): số ngày đến 90% sản lượng ổn định, so với 習熟日数 lý thuyết
  function learning(ds, opts) {
    var b = analysisBase(ds, opts), all5 = doneRuns(b, 5), runs = all5.filter(function (r) { return r.ramp != null; });
    var notReached = all5.length - runs.length;
    var summarize1 = function (list) {
      var withTheo = list.filter(function (r) { return r.learnTheo > 0; });
      return {
        n: list.length,
        ramp: mean(list.map(function (r) { return r.ramp; })),
        rampMedian: median(list.map(function (r) { return r.ramp; })),
        theo: mean(withTheo.map(function (r) { return r.learnTheo; })),
        slower: withTheo.length ? withTheo.filter(function (r) { return r.ramp > Math.ceil(r.learnTheo); }).length / withTheo.length : null,
        loss: list.reduce(function (a, r) { return a + r.loss; }, 0),
        curve: CURVE_DAYS.map(function (d, i) {
          var v = list.map(function (r) { return r.curve[i]; }).filter(function (x) { return x != null; });
          return v.length ? mean(v) : null;
        }),
        newRamp: mean(list.filter(function (r) { return !r.repeat; }).map(function (r) { return r.ramp; })),
        repeatRamp: mean(list.filter(function (r) { return r.repeat; }).map(function (r) { return r.ramp; })),
        nNew: list.filter(function (r) { return !r.repeat; }).length,
        nRepeat: list.filter(function (r) { return r.repeat; }).length
      };
    };
    var rows = SAM_BANDS.map(function (sb) {
      return { band: sb, v: summarize1(runs.filter(function (r) { return samBand(r.sam) === sb.id; })) };
    });
    var slowest = runs.filter(function (r) { return r.learnTheo > 0; })
      .map(function (r) { return Object.assign({ over: r.ramp - r.learnTheo }, r); })
      .sort(function (a, c) { return c.over - a.over; }).slice(0, 15);
    return { months: b.months, curveDays: CURVE_DAYS, rows: rows, total: summarize1(runs), slowest: slowest, notReached: notReached, maxDays: RAMP_MAX_DAYS };
  }

  // 3. Lead time may: số ngày làm việc từ ngày may đầu đến ngày cuối, so với 縫製リードタイム lý thuyết
  var LT_TOLERANCE = 2; // vượt quá lý thuyết hơn 2 ngày mới tính là chậm
  function leadTime(ds, opts) {
    var b = analysisBase(ds, opts), runs = doneRuns(b, 1);
    var summarize1 = function (list) {
      var withTheo = list.filter(function (r) { return r.ltTheo > 0; });
      var over = withTheo.filter(function (r) { return r.span > r.ltTheo + LT_TOLERANCE; });
      return {
        n: list.length,
        theo: mean(withTheo.map(function (r) { return r.ltTheo; })),
        span: mean(list.map(function (r) { return r.span; })),
        days: mean(list.map(function (r) { return r.days; })),
        overRate: withTheo.length ? over.length / withTheo.length : null,
        overDays: mean(over.map(function (r) { return r.span - r.ltTheo; }))
      };
    };
    var rows = SAM_BANDS.map(function (sb) {
      return { band: sb, v: summarize1(runs.filter(function (r) { return samBand(r.sam) === sb.id; })) };
    });
    var brands = groupBy(runs, function (r) { return r.brand; });
    var byBrand = Object.keys(brands).map(function (k) { return { key: k, v: summarize1(brands[k]) }; })
      .sort(function (a, c) { return c.v.n - a.v.n; });
    var worst = runs.filter(function (r) { return r.ltTheo > 0; })
      .map(function (r) { return Object.assign({ over: r.span - r.ltTheo }, r); })
      .sort(function (a, c) { return c.over - a.over; }).slice(0, 15);
    return { months: b.months, tolerance: LT_TOLERANCE, rows: rows, byBrand: byBrand, total: summarize1(runs), worst: worst };
  }

  // 4. Chuyển mã: số mã mới vào chuyền và giờ chuẩn mất trong giai đoạn lên năng suất, theo chuyền và theo tháng
  function changeover(ds, opts) {
    var b = analysisBase(ds, opts), set = {};
    b.months.forEach(function (m) { set[m] = 1; });
    var started = b.runs.filter(function (r) { return set[r.start.slice(0, 7)]; });
    var withTh = b.facts.filter(function (x) { return x.th > 0; });
    var mk = function (runsG, factsG) {
      var sah = factsG.reduce(function (a, x) { return a + x.sah; }, 0);
      var loss = runsG.reduce(function (a, r) { return a + r.loss; }, 0);
      return {
        starts: runsG.length, loss: loss, sah: sah, lossRate: sah + loss > 0 ? loss / (sah + loss) : null,
        running: factsG.length, lots: mean(runsG.filter(function (r) { return r.aggLot > 0; }).map(function (r) { return r.aggLot; }))
      };
    };
    var lk = function (x) { return x.f + ' · ' + x.line; };
    var rl = groupBy(started, lk), fl = groupBy(withTh, lk);
    var lines = Object.keys(fl).map(function (k) { return { key: k, v: mk(rl[k] || [], fl[k]) }; })
      .sort(function (a, c) { return c.v.starts - a.v.starts || (c.v.lossRate || 0) - (a.v.lossRate || 0); });
    var rm = groupBy(started, function (r) { return r.start.slice(0, 7); }), fm = groupBy(withTh, function (x) { return x.m; });
    var months = b.months.filter(function (m) { return fm[m]; }).map(function (m) { return { month: m, v: mk(rm[m] || [], fm[m]) }; });
    return { months: b.months, lines: lines, byMonth: months, total: mk(started, withTh) };
  }

  // 5. Độ ổn định kế hoạch: KH đầu tháng so với TH; mã thêm vào / bỏ ra so với KH đầu tháng
  function planStability(ds, opts) {
    var months = opts.months || periodMonths(ds, opts.month, opts.period || 'ytd', opts.startMonth);
    var all = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner, months: months });
    var openMonths = {};
    all.forEach(function (x) { if (x.open && x.khOpen > 0 && x.status === 'chot') openMonths[x.f + '|' + x.m] = 1; });
    var fs = all.filter(function (x) { return openMonths[x.f + '|' + x.m]; });
    var mk = function (g) {
      var openKeys = {}, doneKeys = {}, a = { khOpen: 0, kh: 0, th: 0 };
      g.forEach(function (x) {
        var k = x.f + '|' + x.m + '|' + x.style;
        if (x.open) { a.khOpen += x.khOpen; if (x.khOpen > 0) openKeys[k] = 1; }
        else { a.kh += x.kh; a.th += x.th; if (x.th > 0) doneKeys[k] = 1; }
      });
      var added = Object.keys(doneKeys).filter(function (k) { return !openKeys[k]; }).length;
      var dropped = Object.keys(openKeys).filter(function (k) { return !doneKeys[k]; }).length;
      return {
        khOpen: a.khOpen, kh: a.kh, th: a.th, rate: a.khOpen > 0 ? a.th / a.khOpen : null,
        diff: a.th - a.khOpen, planned: Object.keys(openKeys).length, done: Object.keys(doneKeys).length,
        added: added, dropped: dropped,
        changeRate: Object.keys(openKeys).length ? (added + dropped) / Object.keys(openKeys).length : null
      };
    };
    var brands = groupBy(fs, function (x) { return x.brand; });
    var byBrand = Object.keys(brands).map(function (k) { return { key: k, v: mk(brands[k]) }; })
      .filter(function (r) { return r.v.khOpen > 0 || r.v.th > 0; })
      .sort(function (a, c) { return c.v.khOpen - a.v.khOpen; });
    var mg = groupBy(fs, function (x) { return x.m; });
    var byMonth = Object.keys(mg).sort().map(function (m) { return { month: m, v: mk(mg[m]) }; });
    return { months: months, byBrand: byBrand, byMonth: byMonth, total: mk(fs) };
  }

  // 6. Xếp hạng chuyền: hiệu suất, năng suất trong kỳ; độ ổn định = độ lệch chuẩn hiệu suất các tháng
  var RANK_MIN_PD = 20; // bỏ qua chuyền có dưới 20 ngày công trong kỳ
  function lineRanking(ds, opts) {
    var b = analysisBase(ds, opts), withTh = b.facts.filter(function (x) { return x.th > 0; });
    var g = groupBy(withTh, function (x) { return x.f + ' · ' + x.line; });
    var rows = Object.keys(g).map(function (k) {
      var list = g[k], v = derive(aggregate(list), actualMonths(list, b.months));
      var bm = groupBy(list, function (x) { return x.m; });
      var effs = Object.keys(bm).sort().map(function (m) { return { month: m, eff: derive(aggregate(bm[m]), 1).efficiency }; })
        .filter(function (e) { return e.eff != null; });
      var vals = effs.map(function (e) { return e.eff; });
      return { key: k, v: v, monthly: effs, sd: stdev(vals), min: vals.length ? Math.min.apply(null, vals) : null, max: vals.length ? Math.max.apply(null, vals) : null };
    }).filter(function (r) { return r.v.pdT >= RANK_MIN_PD && r.v.efficiency != null; })
      .sort(function (a, c) { return c.v.efficiency - a.v.efficiency; });
    rows.forEach(function (r, i) { r.rank = i + 1; });
    return { months: b.months, rows: rows, minPd: RANK_MIN_PD };
  }

  // 7. Tải các tháng tới: giờ chuẩn của KH so với năng lực (LĐ bình quân × ngày làm việc × 8 giờ × hiệu suất)
  // Năng lực lấy theo 3 tháng đã chốt gần nhất của từng xưởng. Không lọc theo nhãn hàng.
  var LOAD_BASE_MONTHS = 3, LOAD_AHEAD = 6;
  function capacityLoad(ds, opts) {
    var facs = ds.factories.filter(function (f) { return !opts.factory || f === opts.factory; });
    var rows = [];
    facs.forEach(function (f) {
      var closed = ds.months.filter(function (m) { var me = ds.meta[f + '|' + m]; return me && me.status === 'chot'; });
      var withTh = closed.filter(function (m) { return ds.facts.some(function (x) { return x.f === f && x.m === m && !x.open && x.th > 0; }); });
      var base = withTh.slice(-LOAD_BASE_MONTHS);
      if (!base.length) return;
      var baseFacts = filterFacts(ds.facts, { factory: f, months: base }).filter(function (x) { return !x.open; });
      var bv = derive(aggregate(baseFacts), base.length);
      var lastClosed = closed[closed.length - 1];
      var ahead = ds.months.filter(function (m) { return m > lastClosed && ds.meta[f + '|' + m]; }).slice(0, LOAD_AHEAD);
      ahead.forEach(function (m) {
        var fs = filterFacts(ds.facts, { factory: f, months: [m] }).filter(function (x) { return !x.open; });
        var a = aggregate(fs), wd = ds.meta[f + '|' + m].wd;
        var hours = (bv.labor || 0) * wd * HOURS_PER_DAY;
        var capSah = hours * (bv.efficiency || 0);
        rows.push({
          f: f, month: m, status: ds.meta[f + '|' + m].status, wd: wd, kh: a.kh, sahK: a.sahK,
          labor: bv.labor, efficiency: bv.efficiency, base: base, capSah: capSah,
          load: capSah > 0 ? a.sahK / capSah : null,
          laborNeed: bv.efficiency > 0 && wd > 0 ? a.sahK / (bv.efficiency * HOURS_PER_DAY * wd) : null
        });
      });
    });
    // Tổng các xưởng theo tháng
    var totals = [];
    if (facs.length > 1) {
      var bm = groupBy(rows, function (r) { return r.month; });
      Object.keys(bm).sort().forEach(function (m) {
        var g = bm[m], capSah = 0, sahK = 0, kh = 0, need = 0;
        g.forEach(function (r) { capSah += r.capSah; sahK += r.sahK; kh += r.kh; need += r.laborNeed || 0; });
        totals.push({ f: 'Tất cả', month: m, kh: kh, sahK: sahK, capSah: capSah, load: capSah > 0 ? sahK / capSah : null, laborNeed: need,
          labor: g.reduce(function (a, r) { return a + (r.labor || 0); }, 0), factories: g.length });
      });
    }
    return { rows: rows, totals: totals, baseMonths: LOAD_BASE_MONTHS };
  }
  function loadStatus(load) { return load == null ? null : load > 1.05 ? 'qua-tai' : load < 0.85 ? 'thieu-hang' : 'vua'; }

  // ---------- Tổng hợp toàn xưởng: mọi chỉ tiêu cho một nhóm tháng ----------
  /*
   * opts: { factory, brand } (để trống = toàn xưởng, mọi nhãn hàng); months: danh sách tháng.
   * Chỉ tiêu cộng dồn (sản lượng, ngày công…) cộng qua các tháng; chỉ tiêu tỷ lệ tính lại trên cả nhóm.
   */
  function overviewMetrics(ds, opts, months) {
    var o = { factory: opts.factory, brand: opts.brand, owner: opts.owner, months: months };
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner });
    var main = filterFacts(base, { months: months }).filter(function (x) { return !x.open; });
    var withTh = main.filter(function (x) { return x.th > 0; });
    var nAct = actualMonths(base, months);
    var v = derive(aggregate(filterFacts(base, { months: months })), nAct);
    var facs = ds.factories.filter(function (f) { return !opts.factory || f === opts.factory; });
    // ngày làm việc: bình quân các xưởng mỗi tháng, cộng qua các tháng có TH
    var wd = 0, statuses = {};
    months.forEach(function (m) {
      var ws = facs.map(function (f) { return ds.meta[f + '|' + m]; }).filter(Boolean);
      ws.forEach(function (me) { statuses[me.status] = 1; });
      if (ws.length && withTh.some(function (x) { return x.m === m; })) wd += mean(ws.map(function (me) { return me.wd; }));
    });
    var brands = {}, lotSmall = 0, samHigh = 0;
    withTh.forEach(function (x) {
      brands[x.brand] = 1;
      if (lotBand(x.aggLot) === 0) lotSmall += x.th;
      if (samBand(x.sam) === 2) samHigh += x.th;
    });
    var co = changeover(ds, o).total, learn = learning(ds, o).total, lt = leadTime(ds, o).total, plan = planStability(ds, o).total;
    // % đạt KH đầu tháng: chỉ trên các tháng có KH đầu tháng (có file cuối tháng trước)
    var openM = {};
    filterFacts(base, { months: months }).forEach(function (x) { if (x.open && x.khOpen > 0) openM[x.f + '|' + x.m] = 1; });
    var thOpen = main.reduce(function (a, x) { return a + (openM[x.f + '|' + x.m] ? x.th : 0); }, 0);
    var status = statuses.chot && !statuses.dang && !statuses.kh ? 'chot' : statuses.dang ? 'dang' : statuses.kh && !statuses.chot ? 'kh' : 'dang';
    return {
      months: months, monthsWithTh: nAct, status: status, hasTh: v.th > 0,
      khOpen: v.khOpen || null, kh: v.kh, th: v.th, rateOpen: v.khOpen > 0 ? thOpen / v.khOpen : null, rate: v.rate,
      openMonths: Object.keys(openM).length,
      wd: wd || null, thPerDay: wd > 0 ? v.th / wd : null, sah: v.sah,
      styles: v.styles, brands: Object.keys(brands).length, rows: withTh.length, starts: co.starts, ends: v.end,
      avgSam: v.avgSam, avgLot: v.avgLot, newLot: co.lots,
      lotSmallShare: v.th > 0 ? lotSmall / v.th : null, samHighShare: v.th > 0 ? samHigh / v.th : null,
      lines: v.lines, lineDays: v.ldT, labor: v.labor, perLine: v.perLine, pdT: v.pdT,
      productivity: v.productivity, efficiency: v.efficiency, outPerLineDay: v.ldT > 0 ? v.th / v.ldT : null,
      rampLoss: co.loss, rampLossRate: co.lossRate,
      rampDays: learn.ramp, rampTheo: learn.theo, rampSlower: learn.slower, runsDone: lt.n,
      ltSpan: lt.span, ltTheo: lt.theo, ltOver: lt.overRate,
      planned: plan.planned || null, added: plan.planned ? plan.added : null, dropped: plan.planned ? plan.dropped : null,
      planChange: plan.changeRate
    };
  }

  // Theo tháng trong một năm tài chính: 12 tháng + cả năm + cùng kỳ năm trước
  function overviewYear(ds, opts, fy) {
    var startMonth = opts.startMonth || 4;
    var months = fiscalMonths(fy, startMonth);
    var cols = months.map(function (m) { return { month: m, v: ds.months.indexOf(m) >= 0 ? overviewMetrics(ds, opts, [m]) : null }; });
    var avail = months.filter(function (m) { return ds.months.indexOf(m) >= 0; });
    var base = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner });
    var withTh = avail.filter(function (m) { return base.some(function (x) { return !x.open && x.m === m && x.th > 0; }); });
    var same = withTh.map(function (m) { return addMonths(m, -12); }).filter(function (m) { return ds.months.indexOf(m) >= 0; });
    return {
      fy: fy, label: fiscalLabel(fy, startMonth), cols: cols,
      total: overviewMetrics(ds, opts, avail),
      done: overviewMetrics(ds, opts, withTh), doneMonths: withTh,
      same: same.length ? overviewMetrics(ds, opts, same) : null, sameMonths: same
    };
  }

  // Theo năm: mỗi năm tài chính có dữ liệu; năm gần nhất có thêm cùng kỳ năm trước
  function overviewYears(ds, opts) {
    var startMonth = opts.startMonth || 4, years = {};
    ds.months.forEach(function (m) { years[fiscalYear(m, startMonth)] = 1; });
    return Object.keys(years).map(Number).sort().map(function (fy) {
      var y = overviewYear(ds, opts, fy);
      return { fy: fy, label: y.label, v: y.done.hasTh ? y.done : y.total, doneMonths: y.doneMonths, same: y.same, sameMonths: y.sameMonths };
    });
  }

  // ---------- Thời gian hoạt động của từng nửa chuyền (chuyền mở mới / ngừng) ----------
  /*
   * Mỗi nửa chuyền có mặt từ ngày có sản lượng đầu tiên đến ngày có sản lượng cuối cùng.
   * Sản lượng đầu tiên trong 31 ngày đầu của dữ liệu → coi như có từ đầu (không biết trước đó);
   * sản lượng cuối cùng trong 31 ngày cuối → coi như vẫn còn đến cuối dữ liệu.
   * Khai báo tay (mỗi dòng "15B 1/10/2026" hoặc "15B 1/10/2026 31/3/2027", ngày dạng d/m/yyyy hoặc yyyy-mm-dd)
   * thay cho ngày tự nhận.
   */
  var LINE_EDGE_DAYS = 31;
  function parseLineOverrides(text) {
    var out = {};
    String(text || '').split(/\r?\n/).forEach(function (l) {
      var m = norm(l).toUpperCase().match(/^(\d+[AB])\s+(\S+)(?:\s+(\S+))?/);
      if (!m) return;
      var d = function (t) {
        if (!t) return null;
        var a = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        return a ? a[3] + '-' + pad(+a[2]) + '-' + pad(+a[1]) : toISO(t);
      };
      out[m[1]] = { from: d(m[2]), to: d(m[3]) };
    });
    return out;
  }
  function computeLineLife(lineDay, dayOff, overridesText) {
    var days = Object.keys(dayOff || {}).sort();
    if (!days.length) return {};
    var d0 = days[0], d1 = days[days.length - 1], first = {}, last = {};
    Object.keys(lineDay).sort().forEach(function (iso) {
      Object.keys(lineDay[iso]).forEach(function (pl) { if (!first[pl]) first[pl] = iso; last[pl] = iso; });
    });
    var ov = parseLineOverrides(overridesText), life = {};
    Object.keys(first).concat(Object.keys(ov)).forEach(function (pl) {
      if (life[pl]) return;
      var f = first[pl] || null, l = last[pl] || null;
      var from = !f || f <= addDays(d0, LINE_EDGE_DAYS) ? d0 : f;
      var to = !l || l >= addDays(d1, -LINE_EDGE_DAYS) ? d1 : l;
      var o = ov[pl];
      life[pl] = { line: pl, from: o && o.from ? o.from : from, to: o && o.to ? o.to : (o && o.from && !first[pl] ? d1 : to),
        firstProd: f, lastProd: l, manual: !!o, opened: (o && o.from ? o.from : from) > d0, closed: (o && o.to ? o.to : to) < d1 };
    });
    return life;
  }
  function lineAlive(ds, pl, iso) { var L = (ds.lineLife || {})[pl]; return !L || (iso >= L.from && iso <= L.to); }

  // ---------- Sử dụng chuyền: hàng TI, hàng APIC tự bán, chuyền trống ----------
  /*
   * Đơn vị: nửa chuyền A/B (như 29 chuyền trong cách tính của TI), mỗi ngày làm việc là 1 chuyền-ngày.
   * Chuyền tính trong tháng: mọi nửa chuyền có sản lượng trong năm tài chính chứa tháng đó.
   * Mỗi chuyền-ngày chia thành TI / APIC theo nhãn hàng của phần chiếm chuyền; phần còn lại là trống.
   * Ngày nghỉ (Chủ nhật, lễ) có may: tính riêng, không cộng vào chuyền-ngày.
   */
  var TI_METHOD = { benchQty: 647091, benchLineDays: 8613, overtime: 0.059, lines: 29,
    doc2025: { qty: 654487, lineDays: 8613, apicDays: 330, idleDays: 84, tiDaysPerLine: 283 } };
  function lineUniverse(ds, fy, startMonth) {
    var months = fiscalMonths(fy, startMonth || 4), set = {};
    Object.keys(ds.lineDay || {}).forEach(function (iso) {
      if (months.indexOf(iso.slice(0, 7)) < 0) return;
      Object.keys(ds.lineDay[iso]).forEach(function (pl) { set[pl] = 1; });
    });
    // chuyền khai báo tay / còn hoạt động trong năm dù không có sản lượng
    var fs = months[0] + '-01', fe = monthEnd(months[11]);
    Object.keys(ds.lineLife || {}).forEach(function (pl) { var L = ds.lineLife[pl]; if (L.from <= fe && L.to >= fs) set[pl] = 1; });
    return Object.keys(set).sort(function (a, b) { return a.localeCompare(b, 'en', { numeric: true }); });
  }
  function dayClass(slot, ownerOf) {
    var ti = 0, apic = 0, brands = {};
    Object.keys(slot || {}).forEach(function (b) {
      var w = slot[b];
      if (ownerOf(b) === 'TI') ti += w; else { apic += w; brands[b] = 1; }
    });
    var tot = ti + apic;
    if (tot > 1) { ti /= tot; apic /= tot; tot = 1; }
    return { ti: ti, apic: apic, idle: 1 - tot, apicBrands: Object.keys(brands) };
  }
  function lineUsage(ds, opts, months) {
    var startMonth = opts.startMonth || 4, ownerOf = ds.ownerOf || tiMatcher(null);
    var byMonth = [], lineTot = {}, periods = [];
    var uniCache = {};
    var okMap = opts.relaxed ? (ds.lineMonthAny || {}) : (ds.lineMonthOk || {});
    // Chế độ "relaxed" (đối chiếu khi thiếu file một xưởng): chỉ xét chuyền có sản lượng trong tháng
    var seenIn = function (m) {
      var set = {};
      Object.keys(ds.lineDay || {}).forEach(function (iso) { if (iso.slice(0, 7) === m) Object.keys(ds.lineDay[iso]).forEach(function (pl) { set[pl] = 1; }); });
      return Object.keys(set).sort(function (a, b) { return a.localeCompare(b, 'en', { numeric: true }); });
    };
    months.forEach(function (m) {
      var fy = fiscalYear(m, startMonth);
      var uni = opts.relaxed ? (uniCache['m' + m] = seenIn(m)) : (uniCache[fy] || (uniCache[fy] = lineUniverse(ds, fy, startMonth)));
      var a = { month: m, lines: uni.length, wd: 0, total: 0, ti: 0, apic: 0, idle: 0, offWork: 0 };
      var days = [], dim = daysInMonth(m);
      for (var d = 1; d <= dim; d++) days.push(m + '-' + pad(d));
      if (!okMap[m]) { byMonth.push(Object.assign(a, { noData: true })); return; }
      var alive = {};
      days.forEach(function (iso) {
        var off = ds.dayOff[iso] !== false, slots = (ds.lineDay || {})[iso] || {};
        if (!off) a.wd++;
        uni.forEach(function (pl) {
          if (!opts.relaxed && !lineAlive(ds, pl, iso)) return;
          alive[pl] = 1;
          var c = dayClass(slots[pl], ownerOf), t = lineTot[pl] || (lineTot[pl] = { line: pl, total: 0, ti: 0, apic: 0, idle: 0 });
          if (off) { a.offWork += c.ti + c.apic; return; }
          a.total++; a.ti += c.ti; a.apic += c.apic; a.idle += c.idle;
          t.total++; t.ti += c.ti; t.apic += c.apic; t.idle += c.idle;
        });
      });
      a.lines = Object.keys(alive).length;
      var endIso = days[days.length - 1];
      a.linesEnd = opts.relaxed ? a.lines : uni.filter(function (pl) { return lineAlive(ds, pl, endIso); }).length;
      byMonth.push(a);
    });
    // Giai đoạn liên tục may hàng APIC / chuyền trống của từng nửa chuyền (ngày làm việc)
    var allDays = [];
    months.forEach(function (m) { if (okMap[m]) for (var d = 1; d <= daysInMonth(m); d++) allDays.push(m + '-' + pad(d)); });
    var uniAll = {};
    months.forEach(function (m) { (uniCache[opts.relaxed ? 'm' + m : fiscalYear(m, startMonth)] || []).forEach(function (pl) { uniAll[pl] = 1; }); });
    Object.keys(uniAll).sort(function (a, b) { return a.localeCompare(b, 'en', { numeric: true }); }).forEach(function (pl) {
      var cur = null;
      var close = function () { if (cur) { periods.push(cur); cur = null; } };
      allDays.forEach(function (iso) {
        if (!(iso in (ds.dayOff || {})) || ds.dayOff[iso] !== false) return; // chỉ ngày làm việc
        if (!opts.relaxed && !lineAlive(ds, pl, iso)) { close(); return; }
        var c = dayClass(((ds.lineDay || {})[iso] || {})[pl], ownerOf);
        var kind = c.apic >= 0.25 ? 'apic' : c.idle >= 0.5 ? 'idle' : null;
        if (!kind) { close(); return; }
        if (cur && cur.kind !== kind) close();
        if (!cur) cur = { line: pl, kind: kind, from: iso, to: iso, days: 0, workDays: 0, brands: {} };
        cur.to = iso; cur.workDays++; cur.days += kind === 'apic' ? c.apic : c.idle;
        c.apicBrands.forEach(function (b) { cur.brands[b] = 1; });
      });
      close();
    });
    periods.forEach(function (p) { p.brands = Object.keys(p.brands).join(', '); });
    var total = byMonth.reduce(function (s, a) {
      ['wd', 'total', 'ti', 'apic', 'idle', 'offWork'].forEach(function (k) { s[k] += a[k] || 0; });
      return s;
    }, { wd: 0, total: 0, ti: 0, apic: 0, idle: 0, offWork: 0 });
    var lines = Object.keys(lineTot).map(function (k) { return lineTot[k]; }).sort(function (a, b) { return a.line.localeCompare(b.line, 'en', { numeric: true }); });
    var okMonths = months.filter(function (m) { return okMap[m]; });
    return { months: months, okMonths: okMonths, byMonth: byMonth, total: total, lines: lines, periods: periods,
      universe: Object.keys(uniAll).sort(function (a, b) { return a.localeCompare(b, 'en', { numeric: true }); }) };
  }

  // ---------- File APIC負担金額 của TI (chuyền trống / may hàng APIC tự bán, TI tính phí cho APIC) ----------
  // Nhận diện: ô có chữ "APIC負担金額" ở vài hàng đầu.
  function isBillingSheet(rows) {
    for (var r = 0; r < Math.min(rows.length, 5); r++) for (var c = 0; c < Math.min((rows[r] || []).length, 6); c++) if (/APIC負担金額/.test(norm(rows[r][c]))) return true;
    return false;
  }
  /*
   * Cột: ライン, Part, 人数, 部課, 縫製開始, 縫製エンド, ライン稼働単価, 日数, 金額, APIC負担比率, 請求額, (ghi chú).
   * File có thể có nhiều khối (総括 và N月度): lấy các khối "N月度"; không có thì lấy hết.
   */
  function parseBilling(rows, opts) {
    opts = opts || {};
    var blocks = [], cur = null, cols = null;
    rows.forEach(function (row) {
      row = row || [];
      var title = null;
      row.forEach(function (v) { if (/APIC負担金額/.test(norm(v))) title = norm(v); });
      if (title) { cur = { title: title, rows: [] }; blocks.push(cur); cols = null; return; }
      var hi = row.map(norm).indexOf('ﾗｲﾝ');
      if (hi < 0) hi = row.map(norm).indexOf('ライン');
      if (hi >= 0) {
        var h = row.map(norm);
        var f = function (re) { for (var i = hi; i < h.length; i++) if (re.test(h[i])) return i; return -1; };
        cols = { line: hi, part: f(/^Part$/i), people: f(/人数/), dept: f(/部課/), from: f(/開始/), to: f(/エンド|終了/), price: f(/単価/), days: f(/日数/),
          amount: f(/^金額/), ratio: f(/比率/), charge: f(/請求/) };
        cols.note = Math.max(cols.charge, cols.ratio, cols.amount) + 1;
        return;
      }
      if (!cur || !cols) return;
      var line = row[cols.line];
      if (typeof line !== 'number') return;
      var from = toISO(row[cols.from], opts.date1904), to = toISO(row[cols.to], opts.date1904);
      if (!from || !to) return;
      cur.rows.push({
        line: String(line), part: norm(row[cols.part]).toUpperCase(), people: numOrNull(row[cols.people]), dept: norm(row[cols.dept]),
        from: from, to: to, price: numOrNull(row[cols.price]), days: numOrNull(row[cols.days]),
        amount: numOrNull(row[cols.amount]), ratio: numOrNull(row[cols.ratio]), charge: numOrNull(row[cols.charge]), note: norm(row[cols.note])
      });
    });
    var monthly = blocks.filter(function (b) { return /月度/.test(b.title) && b.rows.length; });
    var use = monthly.length ? monthly : blocks.filter(function (b) { return b.rows.length; });
    var list = [];
    use.forEach(function (b) { list = list.concat(b.rows); });
    if (!list.length) throw new Error('không thấy dòng chuyền nào trong file APIC負担金額');
    // Tháng: theo "N月度" và năm của ngày may; không có thì tháng xuất hiện nhiều nhất
    var cnt = {};
    list.forEach(function (x) { var m = x.from.slice(0, 7); cnt[m] = (cnt[m] || 0) + 1; });
    var month = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; })[0];
    var tm = use.map(function (b) { var m = b.title.match(/(\d{1,2})月度/); return m ? +m[1] : null; }).filter(Boolean)[0];
    if (tm) {
      var cands = Object.keys(cnt).filter(function (m) { return +m.slice(5) === tm; });
      if (cands.length) month = cands[0]; else month = month.slice(0, 4) + '-' + pad(tm);
    }
    var sum = function (k) { return list.reduce(function (a, x) { return a + (x[k] || 0); }, 0); };
    return { kind: 'billing', name: opts.name || '', month: month, rows: list, days: sum('days'), amount: sum('amount'), charge: sum('charge') };
  }
  // Đối chiếu file APIC負担金額 của tháng với chuyền trống + may hàng APIC trang tính được (theo nửa chuyền)
  // File có chữ "NHABE" trong tên: mọi dòng là may hàng APIC tự bán (Nhà Bè)
  function billingSelf(bill, r) { return /NHABE/i.test(bill.name || '') || /自販|NHABE/i.test(r.note || ''); }
  function billingCompare(ds, bill, startMonth) {
    var full = (ds.lineMonthOk || {})[bill.month];
    var u = lineUsage(ds, { startMonth: startMonth || 4, relaxed: !full }, [bill.month]);
    var ok = !u.byMonth[0].noData, inUni = {};
    u.universe.forEach(function (pl) { inUni[pl] = 1; });
    var calc = {};
    u.periods.forEach(function (p) {
      var c = calc[p.line] || (calc[p.line] = { idle: 0, apic: 0, periods: [] });
      c[p.kind] += p.days; c.periods.push(p);
    });
    var ti = {};
    bill.rows.forEach(function (r) {
      var half = (r.part.match(/^[AB]/) || [''])[0], sub = r.part.length > 1;
      var keys = r.part === 'AB' ? [r.line + 'A', r.line + 'B'] : [r.line + half];
      keys.forEach(function (k) {
        var t = ti[k] || (ti[k] = { days: 0, rows: [], self: 0 });
        var d = (r.days || 0) * (sub ? 0.5 : 1);
        t.days += d; t.rows.push(r);
        if (billingSelf(bill, r)) t.self += d;
      });
    });
    var keys = {};
    Object.keys(ti).forEach(function (k) { keys[k] = 1; });
    Object.keys(calc).forEach(function (k) { keys[k] = 1; });
    var rows = Object.keys(keys).sort(function (a, b) { return a.localeCompare(b, 'en', { numeric: true }); }).map(function (k) {
      var t = ti[k] || { days: 0, rows: [], self: 0 }, c = calc[k] || { idle: 0, apic: 0, periods: [] };
      var mine = c.idle + c.apic;
      var has = ok && !!inUni[k];
      return { line: k, hasData: has, tiDays: t.days, tiSelf: t.self, tiRows: t.rows, idle: c.idle, apic: c.apic, mine: has ? mine : null, diff: has ? mine - t.days : null, periods: c.periods };
    });
    return { month: bill.month, dataOk: ok, full: !!full, rows: rows, tiDays: bill.days,
      mine: rows.reduce(function (a, r) { return a + (r.mine || 0); }, 0), missing: rows.filter(function (r) { return !r.hasData; }).length };
  }

  // Năng suất theo cách của TI: sản lượng hàng TI ÷ chuyền-ngày may hàng TI, quy về giờ hành chính (trừ tăng ca)
  function tiProductivity(ds, opts, months) {
    months = months.filter(function (m) { return (ds.lineMonthOk || {})[m]; });
    var base = filterFacts(ds.facts, { factory: opts.factory, months: months, owner: 'TI' }).filter(function (x) { return !x.open; });
    var th = base.reduce(function (a, x) { return a + x.th; }, 0);
    var u = lineUsage(ds, opts, months).total;
    var perLineDay = u.ti > 0 ? th / u.ti : null;
    var per8h = perLineDay != null ? perLineDay * (1 - TI_METHOD.overtime) : null;
    return { months: months, th: th, tiLineDays: u.ti, apicLineDays: u.apic, idleLineDays: u.idle, lineDays: u.total, perLineDay: perLineDay, per8h: per8h,
      annualEq: per8h != null ? per8h * TI_METHOD.benchLineDays : null };
  }

  // ---------- Cam kết hợp đồng với TI (4/2026–3/2029) ----------
  /*
   * - Năng suất: tăng 5%/năm so với năm chuẩn (25年度生産実績, NTC 2025).
   *   Chỉ tiêu: 'efficiency' (giờ chuẩn/giờ công) hoặc 'productivity' (sp/người/ngày).
   *   Mục tiêu năm thứ k: chuẩn × 1,05^k (lũy tiến) hoặc chuẩn × (1 + 5% × k) (tuyến tính).
   * - Hàng khó: 工程時間 > 20.000 giây, tối đa 2% tổng số lượng sản xuất trong năm.
   * - Lot nhỏ: số lượng sản xuất của mã trên từng chuyền (集約ロット) ≤ 500 sp, tối đa 10%.
   * Thực tế: tháng đã chốt (TH). Dự báo cả năm: tháng đã chốt lấy TH, tháng còn lại lấy KH.
   */
  var CONTRACT = { baseFy: 2025, firstFy: 2026, years: 3, growth: 0.05, compound: true, metric: 'efficiency',
    hardSam: 20000, hardMax: 0.02, smallLot: 500, smallMax: 0.10 };
  function isHard(x, c) { return x.sam > c.hardSam; }
  function isSmall(x, c) { var lot = x.aggLot > 0 ? x.aggLot : x.qty; return lot > 0 && lot <= c.smallLot; }
  function contractTarget(base, fy, c) {
    if (base == null) return null;
    var k = fy - c.baseFy;
    return c.compound ? base * Math.pow(1 + c.growth, k) : base * (1 + c.growth * k);
  }
  function contractShares(list, qtyKey, c) {
    var tot = 0, hard = 0, small = 0;
    list.forEach(function (x) {
      var q = x[qtyKey] || 0;
      tot += q; if (isHard(x, c)) hard += q; if (isSmall(x, c)) small += q;
    });
    return { qty: tot, hard: hard, small: small, hardShare: tot > 0 ? hard / tot : null, smallShare: tot > 0 ? small / tot : null };
  }
  function contractYear(ds, opts, fy, c, base) {
    var startMonth = opts.startMonth || 4;
    var months = fiscalMonths(fy, startMonth);
    var facs = ds.factories.filter(function (f) { return !opts.factory || f === opts.factory; });
    var all = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner, months: months }).filter(function (x) { return !x.open; });
    var closed = function (m) {
      var ms = facs.map(function (f) { return ds.meta[f + '|' + m]; }).filter(Boolean);
      return ms.length > 0 && ms.every(function (me) { return me.status === 'chot'; });
    };
    var doneMonths = months.filter(function (m) { return closed(m) && all.some(function (x) { return x.m === m && x.th > 0; }); });
    var planMonths = months.filter(function (m) { return doneMonths.indexOf(m) < 0; });
    var done = all.filter(function (x) { return doneMonths.indexOf(x.m) >= 0; });
    var v = derive(aggregate(done), doneMonths.length || 1);
    var actual = doneMonths.length ? v[c.metric] : null;
    var target = contractTarget(base, fy, c);
    var act = contractShares(done.filter(function (x) { return x.th > 0; }), 'th', c);
    // Dự báo: TH của tháng đã chốt + KH của tháng còn lại
    var fc = contractShares(done.filter(function (x) { return x.th > 0; }), 'th', c);
    var planFacts = all.filter(function (x) { return planMonths.indexOf(x.m) >= 0 && x.kh > 0; });
    var pl = contractShares(planFacts, 'kh', c);
    var fq = fc.qty + pl.qty, fh = fc.hard + pl.hard, fs = fc.small + pl.small;
    var monthly = months.map(function (m) {
      var g = all.filter(function (x) { return x.m === m; }), isDone = doneMonths.indexOf(m) >= 0;
      var mv = derive(aggregate(g), 1);
      var sh = isDone ? contractShares(g.filter(function (x) { return x.th > 0; }), 'th', c) : contractShares(g.filter(function (x) { return x.kh > 0; }), 'kh', c);
      return { month: m, done: isDone, value: isDone ? mv[c.metric] : null, th: isDone ? mv.th : null, kh: mv.kh,
        hardShare: sh.hardShare, smallShare: sh.smallShare, basis: isDone ? 'TH' : (sh.qty > 0 ? 'KH' : null) };
    });
    return {
      fy: fy, months: months, doneMonths: doneMonths, planMonths: planMonths, target: target, actual: actual,
      vsBase: actual != null && base ? actual / base - 1 : null, vsTarget: actual != null && target ? actual / target - 1 : null,
      metrics: v, th: act.qty, hardTh: act.hard, smallTh: act.small, hardShare: act.hardShare, smallShare: act.smallShare,
      // Dự báo chỉ khi năm đã có tháng chốt. Hạn mức còn lại = số sp hàng khó / lot nhỏ các tháng còn lại
      // được phép làm để cả năm không vượt giới hạn (âm = đã vượt dù các tháng sau không làm thêm).
      forecast: doneMonths.length ? {
        qty: fq, hardShare: fq > 0 ? fh / fq : null, smallShare: fq > 0 ? fs / fq : null, planQty: pl.qty,
        hardRoom: c.hardMax * fq - fc.hard, smallRoom: c.smallMax * fq - fc.small,
        planHard: pl.hard, planSmall: pl.small
      } : { qty: null, hardShare: null, smallShare: null, planQty: pl.qty, hardRoom: null, smallRoom: null },
      monthly: monthly, facts: done
    };
  }
  function contractTracking(ds, opts, cfg) {
    var c = Object.assign({}, CONTRACT, cfg || {});
    var startMonth = opts.startMonth || 4;
    var baseMonths = fiscalMonths(c.baseFy, startMonth);
    var bAll = filterFacts(ds.facts, { factory: opts.factory, brand: opts.brand, owner: opts.owner, months: baseMonths }).filter(function (x) { return !x.open; });
    var bMonths = baseMonths.filter(function (m) { return bAll.some(function (x) { return x.m === m && x.th > 0; }); });
    var bv = derive(aggregate(bAll), bMonths.length || 1);
    var base = bMonths.length ? bv[c.metric] : null;
    var bSh = contractShares(bAll.filter(function (x) { return x.th > 0; }), 'th', c);
    var benchmark = { fy: c.baseFy, months: bMonths, full: bMonths.length === 12, value: base, metrics: bv, hardShare: bSh.hardShare, smallShare: bSh.smallShare,
      monthly: baseMonths.map(function (m) { var g = bAll.filter(function (x) { return x.m === m; }); var mv = derive(aggregate(g), 1); return { month: m, value: g.some(function (x) { return x.th > 0; }) ? mv[c.metric] : null }; }) };
    var years = [];
    for (var k = 0; k < c.years; k++) years.push(contractYear(ds, opts, c.firstFy + k, c, base));
    // Cách tính của TI: sản lượng hàng TI / chuyền-ngày TI (giờ hành chính); năm chuẩn 647.091 sp = 29 chuyền × 297 ngày
    var tiBase = TI_METHOD.benchQty / TI_METHOD.benchLineDays;
    benchmark.ti = tiProductivity(ds, opts, bMonths);
    years.forEach(function (y) {
      y.ti = tiProductivity(ds, opts, y.doneMonths);
      y.tiTarget = contractTarget(tiBase, y.fy, c);
      y.tiTargetQty = contractTarget(TI_METHOD.benchQty, y.fy, c);
    });
    return { cfg: c, benchmark: benchmark, years: years, tiBase: tiBase };
  }
  // Danh sách mã gây tỷ lệ hàng khó / lot nhỏ trong năm (tháng đã chốt)
  function contractLists(year, cfg) {
    var c = Object.assign({}, CONTRACT, cfg || {});
    var hard = {}, small = {};
    year.facts.forEach(function (x) {
      if (!(x.th > 0)) return;
      if (isHard(x, c)) { var k = x.f + '|' + x.style; var h = hard[k] || (hard[k] = { f: x.f, style: x.style, brand: x.brand, item: x.item, sam: x.sam, th: 0, lines: {} }); h.th += x.th; h.lines[x.line] = 1; }
      if (isSmall(x, c)) { var k2 = x.f + '|' + x.style + '|' + x.lineRaw; var s2 = small[k2] || (small[k2] = { f: x.f, style: x.style, line: x.lineRaw, brand: x.brand, item: x.item, lot: x.aggLot > 0 ? x.aggLot : x.qty, sam: x.sam, th: 0 }); s2.th += x.th; }
    });
    var sort = function (o) { return Object.keys(o).map(function (k) { return o[k]; }).sort(function (a, b) { return b.th - a.th; }); };
    return { hard: sort(hard).map(function (h) { h.lines = Object.keys(h.lines).join(', '); return h; }), small: sort(small) };
  }

  return {
    HOURS_PER_DAY: HOURS_PER_DAY, BRAND_ALIASES: BRAND_ALIASES, DIMS: DIMS,
    toISO: toISO, serialToISO: serialToISO, addDays: addDays, addMonths: addMonths, monthEnd: monthEnd, monthLabel: monthLabel,
    fiscalYear: fiscalYear, fiscalMonths: fiscalMonths, fiscalLabel: fiscalLabel, updateDateFromName: updateDateFromName,
    parseAliasText: parseAliasText, normBrand: normBrand, lineGroup: lineGroup, baseLine: baseLine,
    SUMMARY_CELLS: SUMMARY_CELLS, checkSummary: checkSummary, closesMonth: closesMonth, coverage: coverage,
    parseRows: parseRows, buildDataset: buildDataset, aggregate: aggregate, derive: derive,
    filterFacts: filterFacts, groupBy: groupBy,
    SAM_BANDS: SAM_BANDS, LOT_BANDS: LOT_BANDS, ANALYSIS_PERIODS: ANALYSIS_PERIODS, CURVE_DAYS: CURVE_DAYS, DATA_VERSION: DATA_VERSION,
    physicalLines: physicalLines,
    samBand: samBand, lotBand: lotBand, periodMonths: periodMonths, runStats: runStats,
    lotEfficiency: lotEfficiency, learning: learning, leadTime: leadTime, changeover: changeover,
    planStability: planStability, lineRanking: lineRanking, capacityLoad: capacityLoad, loadStatus: loadStatus,
    overviewMetrics: overviewMetrics, overviewYear: overviewYear, overviewYears: overviewYears,
    TI_BRANDS_DEFAULT: TI_BRANDS_DEFAULT, TI_BRANDS_ADDED: TI_BRANDS_ADDED, TI_METHOD: TI_METHOD, brandKey: brandKey, tiMatcher: tiMatcher,
    lineUniverse: lineUniverse, lineUsage: lineUsage, computeLineLife: computeLineLife, parseLineOverrides: parseLineOverrides, tiProductivity: tiProductivity,
    findDayColumns: findDayColumns, isBillingSheet: isBillingSheet, billingSelf: billingSelf, parseBilling: parseBilling, billingCompare: billingCompare,
    CONTRACT: CONTRACT, contractTracking: contractTracking, contractLists: contractLists, contractTarget: contractTarget, periods: periods, summarize: summarize, monthly: monthly, fiscal: fiscal, pct: pct
  };
});
