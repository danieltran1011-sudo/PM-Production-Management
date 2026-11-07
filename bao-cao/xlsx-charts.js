/*
 * Thêm biểu đồ Excel (DrawingML chart, sửa được trong Excel) vào file .xlsx do SheetJS tạo.
 * SheetJS bản miễn phí không ghi được biểu đồ, nên mở gói .xlsx bằng JSZip rồi thêm:
 *   xl/charts/chartN.xml, xl/drawings/drawingN.xml (+ rels), thẻ <drawing> trong sheet, [Content_Types].xml.
 *
 * Dùng được trong trình duyệt (window.PMXlsxCharts) và Node (require).
 *
 * addCharts(zip, sheetName, charts) — zip: đối tượng JSZip đã nạp file .xlsx.
 * chart = {
 *   title, anchor: { col, row, cols, rows }   (ô bắt đầu, 0-based; kích thước theo số cột/hàng)
 *   cats: { ref: "'Bieu do'!$A$3:$A$14", values: ['T4', ...] },
 *   series: [{ name, ref, values: [số hoặc null], color: '2A78D6', type: 'bar' | 'line', dash: false }],
 *   stacked: false, yFmt: '#,##0' | '0%' | '#,##0.0'
 * }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PMXlsxCharts = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  var REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
  var REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
  var CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
  var CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

  function x(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

  function strCache(values) {
    return '<c:strCache><c:ptCount val="' + values.length + '"/>' + values.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + x(v) + '</c:v></c:pt>'; }).join('') + '</c:strCache>';
  }
  function numCache(values, fmt) {
    return '<c:numCache><c:formatCode>' + x(fmt || 'General') + '</c:formatCode><c:ptCount val="' + values.length + '"/>' +
      values.map(function (v, i) { return v == null || !isFinite(v) ? '' : '<c:pt idx="' + i + '"><c:v>' + v + '</c:v></c:pt>'; }).join('') + '</c:numCache>';
  }
  function serXml(s, i, cats, type, fmt) {
    var color = (s.color || '2A78D6').replace('#', '');
    var sp = type === 'line'
      ? '<c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' + (s.dash ? '<a:prstDash val="dash"/>' : '') + '<a:round/></a:ln></c:spPr>' +
        (s.dash ? '<c:marker><c:symbol val="none"/></c:marker>' : '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>')
      : '<c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>';
    return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:v>' + x(s.name) + '</c:v></c:tx>' + sp +
      '<c:cat><c:strRef><c:f>' + x(cats.ref) + '</c:f>' + strCache(cats.values) + '</c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>' + x(s.ref) + '</c:f>' + numCache(s.values, fmt) + '</c:numRef></c:val>' +
      (type === 'line' ? '<c:smooth val="0"/>' : '') + '</c:ser>';
  }
  function chartXml(ch) {
    var fmt = ch.yFmt || '#,##0';
    var bars = [], lines = [];
    ch.series.forEach(function (s, i) { (s.type === 'line' ? lines : bars).push({ s: s, i: i }); });
    var ax = '<c:axId val="50010"/><c:axId val="50020"/>';
    // biểu đồ cột luôn bắt đầu từ 0
    var yMin = ch.yMin != null ? ch.yMin : bars.length ? 0 : null;
    var plot = '';
    if (bars.length) {
      plot += '<c:barChart><c:barDir val="col"/><c:grouping val="' + (ch.stacked ? 'stacked' : 'clustered') + '"/><c:varyColors val="0"/>' +
        bars.map(function (b) { return serXml(b.s, b.i, ch.cats, 'bar', fmt); }).join('') +
        '<c:gapWidth val="' + (ch.stacked ? 60 : 80) + '"/>' + (ch.stacked ? '<c:overlap val="100"/>' : '') + ax + '</c:barChart>';
    }
    if (lines.length) {
      plot += '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
        lines.map(function (l) { return serXml(l.s, l.i, ch.cats, 'line', fmt); }).join('') + '<c:marker val="1"/>' + ax + '</c:lineChart>';
    }
    var txt = function (sz, color) { return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + sz + '"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="vi-VN"/></a:p></c:txPr>'; };
    var axLine = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFC5D2"/></a:solidFill></a:ln></c:spPr>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="' + NS_C + '" xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '"><c:roundedCorners val="0"/>' +
      '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="vi-VN" sz="1200" b="1"/><a:t>' + x(ch.title) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' + plot +
      '<c:catAx><c:axId val="50010"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>' +
      '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>' + axLine + txt(900, '5B6482') + '<c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/>' + (yMin != null ? '<c:min val="' + yMin + '"/>' : '') + '</c:scaling><c:delete val="0"/><c:axPos val="l"/>' +
      '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E3E7EF"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' +
      '<c:numFmt formatCode="' + x(fmt) + '" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>' + txt(900, '5B6482') +
      '<c:crossAx val="50010"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea>' +
      '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + txt(900, '17203B') + '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
      '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D8DDE8"/></a:solidFill></a:ln></c:spPr></c:chartSpace>';
  }
  function anchorXml(ch, i, rid) {
    var a = ch.anchor;
    return '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + a.col + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + a.row + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + (a.col + a.cols) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + (a.row + a.rows) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="' + x(ch.title) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="' + NS_C + '">' +
      '<c:chart xmlns:c="' + NS_C + '" xmlns:r="' + NS_R + '" r:id="' + rid + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
  }

  // Tìm file XML của sheet theo tên (workbook.xml → r:id → workbook.xml.rels)
  async function sheetPath(zip, sheetName) {
    var wb = await zip.file('xl/workbook.xml').async('string');
    var rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    var re = /<sheet\b[^>]*>/g, m, rid = null;
    while ((m = re.exec(wb))) {
      var name = (m[0].match(/name="([^"]*)"/) || [])[1];
      if (name != null && name.replace(/&amp;/g, '&').replace(/&apos;/g, "'") === sheetName) rid = (m[0].match(/r:id="([^"]*)"/) || [])[1];
    }
    if (!rid) throw new Error('không thấy sheet ' + sheetName);
    var rr = new RegExp('<Relationship\\b[^>]*Id="' + rid + '"[^>]*>'), rel = rels.match(rr);
    var target = rel && (rel[0].match(/Target="([^"]*)"/) || [])[1];
    if (!target) throw new Error('không thấy đường dẫn sheet ' + sheetName);
    return target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  }

  async function addCharts(zip, sheetName, charts) {
    if (!charts || !charts.length) return zip;
    var sp = await sheetPath(zip, sheetName);
    var dir = sp.slice(0, sp.lastIndexOf('/')), base = sp.slice(sp.lastIndexOf('/') + 1);
    // số thứ tự mới cho drawing / chart
    var n = 1;
    while (zip.file('xl/drawings/drawing' + n + '.xml')) n++;
    var c0 = 1;
    while (zip.file('xl/charts/chart' + c0 + '.xml')) c0++;
    var drawingRels = [], anchors = [];
    charts.forEach(function (ch, i) {
      var cn = c0 + i;
      zip.file('xl/charts/chart' + cn + '.xml', chartXml(ch));
      drawingRels.push('<Relationship Id="rId' + (i + 1) + '" Type="' + REL_CHART + '" Target="../charts/chart' + cn + '.xml"/>');
      anchors.push(anchorXml(ch, i, 'rId' + (i + 1)));
    });
    zip.file('xl/drawings/drawing' + n + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="' + NS_XDR + '" xmlns:a="' + NS_A + '">' + anchors.join('') + '</xdr:wsDr>');
    zip.file('xl/drawings/_rels/drawing' + n + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + drawingRels.join('') + '</Relationships>');
    // quan hệ sheet → drawing
    var relPath = dir + '/_rels/' + base + '.rels';
    var relXml = zip.file(relPath) ? await zip.file(relPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    var rid = 'rIdDrw' + n;
    relXml = relXml.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="' + REL_DRAWING + '" Target="../drawings/drawing' + n + '.xml"/></Relationships>');
    zip.file(relPath, relXml);
    // thẻ <drawing> trong sheet (trước legacyDrawing/tableParts/extLst nếu có, không thì trước </worksheet>)
    var sx = await zip.file(sp).async('string');
    if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="' + NS_R + '"');
    var tag = '<drawing r:id="' + rid + '"/>';
    var mm = sx.match(/<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/);
    sx = mm ? sx.replace(mm[0], tag + mm[0]) : sx.replace('</worksheet>', tag + '</worksheet>');
    zip.file(sp, sx);
    // [Content_Types].xml
    var ct = await zip.file('[Content_Types].xml').async('string');
    var add = '<Override PartName="/xl/drawings/drawing' + n + '.xml" ContentType="' + CT_DRAWING + '"/>';
    charts.forEach(function (_, i) { add += '<Override PartName="/xl/charts/chart' + (c0 + i) + '.xml" ContentType="' + CT_CHART + '"/>'; });
    zip.file('[Content_Types].xml', ct.replace('</Types>', add + '</Types>'));
    return zip;
  }

  return { addCharts: addCharts, chartXml: chartXml };
});
