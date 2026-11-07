const test = require('node:test');
const assert = require('node:assert');
const XS = require('./xlsx-style.js');

const SHEET = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
  '<sheetData><row r="1"><c r="A1" t="str"><v>Tiêu đề</v></c></row><row r="3"><c r="A3"><v>1</v></c><c r="B3" s="1"><v>0.5</v></c></row></sheetData></worksheet>';

test('Định dạng ô: gắn s, tạo ô trống có viền, giữ định dạng số cũ, cố định hàng', () => {
  const reg = new XS.Registry();
  const out = XS.styleSheetXml(SHEET, {
    cells: { A1: { b: 1, sz: 14 }, A3: { border: 'all', fmt: '#,##0' }, C3: { border: 'all' } },
    rowHeights: { 1: 24 }, freeze: { row: 2, col: 1 }, gridlines: false
  }, reg, ['General', '0.0%']);
  assert.match(out, /<row r="1" ht="24" customHeight="1"><c r="A1" t="str" s="1">/);
  assert.match(out, /<c r="A3" s="2"><v>1<\/v><\/c><c r="B3" s="3"><v>0.5<\/v><\/c><c r="C3" s="4"\/>/);
  assert.match(out, /<dimension ref="A1:C3"\/>/);
  assert.match(out, /showGridLines="0"/);
  assert.match(out, /<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"\/>/);
  const styles = reg.xml();
  assert.match(styles, /<numFmt numFmtId="164" formatCode="0.0%"\/>/);
  assert.match(styles, /<cellXfs count="5">/);
});

test('Định dạng ô: hàng chưa có trong sheet được thêm đúng thứ tự', () => {
  const reg = new XS.Registry();
  const out = XS.styleSheetXml(SHEET, { cells: { B2: { fill: 'EAEFF5' }, A5: { b: 1 } } }, reg, []);
  const rows = [...out.matchAll(/<row r="(\d+)"/g)].map(m => +m[1]);
  assert.deepStrictEqual(rows, [1, 2, 3, 5]);
  assert.match(out, /<c r="B3"><v>0.5<\/v><\/c>/); // định dạng cũ 'General' bị bỏ
});
