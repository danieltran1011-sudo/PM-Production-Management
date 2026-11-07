/*
 * Định dạng ô (màu, viền, chữ đậm, định dạng số, cố định hàng tiêu đề) cho file .xlsx do SheetJS tạo.
 * SheetJS bản miễn phí không ghi được định dạng, nên mở gói .xlsx bằng JSZip, viết lại xl/styles.xml
 * và gắn chỉ số định dạng (s="…") vào từng ô trong sheet.
 *
 * Dùng được trong trình duyệt (window.PMXlsxStyle) và Node (require).
 *
 * applyStyles(zip, sheets) — zip: đối tượng JSZip đã nạp file .xlsx.
 * sheet = {
 *   name,                                   tên sheet
 *   cells: { 'A1': style, ... },            định dạng từng ô (ô chưa có sẽ được tạo trống để có viền / nền)
 *   rowHeights: { 1: 24, ... },             chiều cao hàng (số hàng Excel, 1-based)
 *   freeze: { row: 3, col: 1 },             cố định 3 hàng đầu và 1 cột đầu
 *   gridlines: false                        ẩn đường lưới
 * }
 * style = { b, i, sz, color: 'RRGGBB', fill: 'RRGGBB', border: 'all' | 'bottom', borderColor, h, v, wrap, fmt: '#,##0' }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PMXlsxStyle = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'm/d/yy', 49: '@' };
  var FONT_NAME = 'Calibri';

  function x(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function unx(s) { return String(s).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'); }

  function colIndex(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; }
  function colLetters(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function splitRef(ref) { var m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { c: colIndex(m[1]), r: +m[2] } : null; }

  // Bảng định dạng dùng chung cho cả file: mỗi kiểu ô khác nhau → một chỉ số xf
  function Registry() {
    this.fonts = ['<font><sz val="11"/><color theme="1"/><name val="' + FONT_NAME + '"/><family val="2"/><scheme val="minor"/></font>'];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = []; // [{id, code}]
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.keys = {};
  }
  Registry.prototype.part = function (list, xml) {
    var idx = list.indexOf(xml);
    if (idx >= 0) return idx;
    list.push(xml);
    return list.length - 1;
  };
  Registry.prototype.fmtId = function (code) {
    if (!code || code === 'General') return 0;
    for (var id in BUILTIN_FMT) if (BUILTIN_FMT[id] === code) return +id;
    for (var i = 0; i < this.numFmts.length; i++) if (this.numFmts[i].code === code) return this.numFmts[i].id;
    var nid = 164 + this.numFmts.length;
    this.numFmts.push({ id: nid, code: code });
    return nid;
  };
  Registry.prototype.xf = function (st) {
    if (!st) return 0;
    var key = JSON.stringify([st.b, st.i, st.sz, st.color, st.fill, st.border, st.borderColor, st.h, st.v, st.wrap, st.fmt, st.indent]);
    if (this.keys[key] != null) return this.keys[key];
    var font = 0, fill = 0, border = 0;
    if (st.b || st.i || st.sz || st.color) {
      font = this.part(this.fonts, '<font>' + (st.b ? '<b/>' : '') + (st.i ? '<i/>' : '') + '<sz val="' + (st.sz || 11) + '"/>' +
        (st.color ? '<color rgb="FF' + st.color + '"/>' : '<color theme="1"/>') + '<name val="' + FONT_NAME + '"/><family val="2"/><scheme val="minor"/></font>');
    }
    if (st.fill) fill = this.part(this.fills, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + st.fill + '"/><bgColor indexed="64"/></patternFill></fill>');
    if (st.border) {
      var bc = '<color rgb="FF' + (st.borderColor || 'D5DBE5') + '"/>';
      var side = function (n, on, w) { return on ? '<' + n + ' style="' + (w || 'thin') + '">' + bc + '</' + n + '>' : '<' + n + '/>'; };
      var all = st.border === 'all';
      border = this.part(this.borders, '<border>' + side('left', all) + side('right', all) + side('top', all || st.border === 'top', st.border === 'top' ? 'thin' : null) +
        side('bottom', all || st.border === 'bottom', st.border === 'bottom' ? 'medium' : null) + '<diagonal/></border>');
    }
    var numFmt = this.fmtId(st.fmt);
    var align = (st.h || st.v || st.wrap || st.indent) ? '<alignment' + (st.h ? ' horizontal="' + st.h + '"' : '') + (st.v ? ' vertical="' + st.v + '"' : '') +
      (st.wrap ? ' wrapText="1"' : '') + (st.indent ? ' indent="' + st.indent + '"' : '') + '/>' : '';
    var xml = '<xf numFmtId="' + numFmt + '" fontId="' + font + '" fillId="' + fill + '" borderId="' + border + '" xfId="0"' +
      (numFmt ? ' applyNumberFormat="1"' : '') + (font ? ' applyFont="1"' : '') + (fill ? ' applyFill="1"' : '') + (border ? ' applyBorder="1"' : '') +
      (align ? ' applyAlignment="1">' + align + '</xf>' : '/>');
    this.xfs.push(xml);
    this.keys[key] = this.xfs.length - 1;
    return this.keys[key];
  };
  Registry.prototype.xml = function () {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (this.numFmts.length ? '<numFmts count="' + this.numFmts.length + '">' + this.numFmts.map(function (f) { return '<numFmt numFmtId="' + f.id + '" formatCode="' + x(f.code) + '"/>'; }).join('') + '</numFmts>' : '') +
      '<fonts count="' + this.fonts.length + '">' + this.fonts.join('') + '</fonts>' +
      '<fills count="' + this.fills.length + '">' + this.fills.join('') + '</fills>' +
      '<borders count="' + this.borders.length + '">' + this.borders.join('') + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/></styleSheet>';
  };

  // Định dạng số cũ của SheetJS (ô có .z) → giữ lại khi viết styles.xml mới
  function oldFormats(stylesXml) {
    var codes = {}, out = [];
    var re = /<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g, m;
    while ((m = re.exec(stylesXml))) codes[m[1]] = unx(m[2]);
    var cx = (stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
    var rx = /<xf\b([^>]*)/g;
    while ((m = rx.exec(cx))) {
      var id = (m[1].match(/numFmtId="(\d+)"/) || [])[1] || '0';
      out.push(codes[id] || BUILTIN_FMT[id] || 'General');
    }
    return out;
  }

  async function sheetPath(zip, sheetName) {
    var wb = await zip.file('xl/workbook.xml').async('string');
    var rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    var re = /<sheet\b[^>]*>/g, m, rid = null;
    while ((m = re.exec(wb))) {
      var name = (m[0].match(/name="([^"]*)"/) || [])[1];
      if (name != null && unx(name.replace(/&apos;/g, "'")) === sheetName) rid = (m[0].match(/r:id="([^"]*)"/) || [])[1];
    }
    if (!rid) throw new Error('không thấy sheet ' + sheetName);
    var rel = rels.match(new RegExp('<Relationship\\b[^>]*Id="' + rid + '"[^>]*>'));
    var target = rel && (rel[0].match(/Target="([^"]*)"/) || [])[1];
    if (!target) throw new Error('không thấy đường dẫn sheet ' + sheetName);
    return target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  }

  var CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

  // Viết lại một sheet: gắn s="…", thêm ô trống có định dạng, chiều cao hàng, cố định tiêu đề
  function styleSheetXml(sx, spec, reg, oldFmt) {
    var cells = spec.cells || {}, heights = spec.rowHeights || {};
    var byRow = {};
    Object.keys(cells).forEach(function (ref) {
      var p = splitRef(ref); if (!p) return;
      (byRow[p.r] = byRow[p.r] || {})[p.c] = cells[ref];
    });
    Object.keys(heights).forEach(function (r) { byRow[r] = byRow[r] || {}; });
    var pending = Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; });
    var seen = {}, maxR = 0, maxC = 0;
    var oldXf = function (attrs) {
      var s = (attrs.match(/\bs="(\d+)"/) || [])[1];
      return s != null && oldFmt[+s] && oldFmt[+s] !== 'General' ? reg.xf({ fmt: oldFmt[+s] }) : 0;
    };
    var rowXml = function (r, attrs, inner) {
      var want = byRow[r] || {}, list = [], have = {};
      if (inner) {
        CELL_RE.lastIndex = 0;
        var m;
        while ((m = CELL_RE.exec(inner))) {
          var ref = (m[1].match(/\br="([A-Z]+\d+)"/) || [])[1];
          var p = ref && splitRef(ref);
          if (!p) continue;
          var a = m[1].replace(/\s+s="\d+"/, '');
          var s = want[p.c] ? reg.xf(want[p.c]) : oldXf(m[1]);
          list.push({ c: p.c, xml: '<c' + a + (s ? ' s="' + s + '"' : '') + (m[2] != null ? '>' + m[2] + '</c>' : '/>') });
          have[p.c] = 1;
          if (p.c > maxC) maxC = p.c;
        }
      }
      Object.keys(want).forEach(function (c) {
        c = +c;
        if (have[c]) return;
        list.push({ c: c, xml: '<c r="' + colLetters(c) + r + '" s="' + reg.xf(want[c]) + '"/>' });
        if (c > maxC) maxC = c;
      });
      list.sort(function (a, b) { return a.c - b.c; });
      attrs = (attrs || ' r="' + r + '"').replace(/\s+(ht|customHeight|spans)="[^"]*"/g, '');
      if (heights[r]) attrs += ' ht="' + heights[r] + '" customHeight="1"';
      if (r > maxR) maxR = r;
      return '<row' + attrs + (list.length ? '>' + list.map(function (q) { return q.xml; }).join('') + '</row>' : '/>');
    };
    var flushBefore = function (r) {
      var out = '';
      while (pending.length && pending[0] < r) { var pr = pending.shift(); if (!seen[pr]) out += rowXml(pr, null, ''); }
      return out;
    };
    var data = (sx.match(/<sheetData>([\s\S]*?)<\/sheetData>/) || [])[1];
    if (data == null) data = '';
    var newData = data.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, function (all, attrs, inner) {
      var r = +((attrs.match(/\br="(\d+)"/) || [])[1]);
      var pre = flushBefore(r);
      seen[r] = 1;
      if (r > maxR) maxR = r;
      if (!byRow[r] && !/\bs="\d+"/.test(inner || '')) {
        // hàng không đổi: giữ nguyên, chỉ cập nhật cột lớn nhất
        var refs = (inner || '').match(/<c r="[A-Z]+/g);
        if (refs) { var lc = colIndex(refs[refs.length - 1].slice(6)); if (lc > maxC) maxC = lc; }
        return pre + all;
      }
      if (pending.length && pending[0] === r) pending.shift();
      return pre + rowXml(r, attrs, inner || '');
    });
    newData += flushBefore(Infinity);
    sx = sx.replace(/<sheetData>[\s\S]*?<\/sheetData>|<sheetData\/>/, '<sheetData>' + newData + '</sheetData>');
    if (maxR > 0) {
      var dim = '<dimension ref="A1:' + colLetters(Math.max(maxC, 0)) + maxR + '"/>';
      sx = /<dimension\b[^>]*\/>/.test(sx) ? sx.replace(/<dimension\b[^>]*\/>/, dim) : sx;
    }
    // sheetView: cố định hàng / cột, ẩn lưới
    var fz = spec.freeze, pane = '';
    if (fz && (fz.row || fz.col)) {
      var tl = colLetters(fz.col || 0) + ((fz.row || 0) + 1);
      var ap = fz.row && fz.col ? 'bottomRight' : fz.row ? 'bottomLeft' : 'topRight';
      pane = '<pane' + (fz.col ? ' xSplit="' + fz.col + '"' : '') + (fz.row ? ' ySplit="' + fz.row + '"' : '') + ' topLeftCell="' + tl + '" activePane="' + ap + '" state="frozen"/>' +
        '<selection pane="' + ap + '" activeCell="' + tl + '" sqref="' + tl + '"/>';
    }
    var view = '<sheetViews><sheetView workbookViewId="0"' + (spec.gridlines === false ? ' showGridLines="0"' : '') + (spec.selected ? ' tabSelected="1"' : '') +
      (spec.zoom ? ' zoomScale="' + spec.zoom + '"' : '') + (pane ? '>' + pane + '</sheetView>' : '/>') + '</sheetViews>';
    if (/<sheetViews>[\s\S]*?<\/sheetViews>/.test(sx)) sx = sx.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, view);
    else sx = sx.replace(/(<dimension\b[^>]*\/>)/, '$1' + view);
    if (spec.tabColor) {
      var tab = '<sheetPr><tabColor rgb="FF' + spec.tabColor + '"/></sheetPr>';
      sx = /<sheetPr\b/.test(sx) ? sx : sx.replace(/(<worksheet\b[^>]*>)/, '$1' + tab);
    }
    return sx;
  }

  async function applyStyles(zip, sheets) {
    var reg = new Registry();
    var stylesXml = zip.file('xl/styles.xml') ? await zip.file('xl/styles.xml').async('string') : '';
    var oldFmt = oldFormats(stylesXml);
    var specs = {};
    (sheets || []).forEach(function (s) { specs[s.name] = s; });
    // mọi sheet trong file (cả sheet không có định dạng riêng) để giữ định dạng số cũ
    var wb = await zip.file('xl/workbook.xml').async('string');
    var names = [], re = /<sheet\b[^>]*name="([^"]*)"/g, m;
    while ((m = re.exec(wb))) names.push(unx(m[1].replace(/&apos;/g, "'")));
    for (var i = 0; i < names.length; i++) {
      var sp = await sheetPath(zip, names[i]);
      var sx = await zip.file(sp).async('string');
      zip.file(sp, styleSheetXml(sx, specs[names[i]] || {}, reg, oldFmt));
    }
    zip.file('xl/styles.xml', reg.xml());
    return zip;
  }

  return { applyStyles: applyStyles, styleSheetXml: styleSheetXml, Registry: Registry, colLetters: colLetters, colIndex: colIndex };
});
