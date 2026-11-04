// Chạy: npm test
// Kiểm tra đọc sheet XN1AB (dữ liệu giả lập cùng bố cục file KH_UPDATE), gộp nhiều file,
// chọn file chốt tháng, lao động chuyền, năm tài chính và các kỳ so sánh.
const test = require('node:test');
const assert = require('node:assert');
const R = require('./report.js');

// Số serial Excel của ngày ISO
function serial(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5;
}

// Tạo sheet giống bố cục XN1AB: tiêu đề ở hàng 6, cột ngày từ CE (chỉ số 82), cặp hàng KH/TH từ hàng 13.
function makeSheet({ from, to, off = [], styles }) {
  const D0 = 82, dates = [];
  for (let d = from; d <= to; d = R.addDays(d, 1)) dates.push(d);
  const width = D0 + dates.length + 2;
  const blank = () => new Array(width).fill(null);
  const rows = [];
  for (let i = 0; i < 12; i++) rows.push(blank());
  const H = rows[5];
  H[1] = 10; H[19] = '工場'; H[20] = 'ライン'; H[21] = 'ブランド'; H[22] = '品番'; H[23] = 'アイテム';
  H[24] = 'PO'; H[25] = '集約ロット(pcs)'; H[26] = '生産数(pcs)'; H[27] = '工程分析時間(sec)'; H[28] = 'ライン編成人数(人/line)';
  H[6] = '工程分析時間(sec)'; H[7] = 'ライン編成人数(人/line)'; // cột tóm tắt bên trái, phải bỏ qua
  dates.forEach((d, j) => {
    H[D0 + j] = serial(d);
    if (off.includes(d)) rows[2][D0 + j] = 'x';
  });
  for (const s of styles) {
    const p = blank(), k = blank(), t = blank();
    p[1] = 'dự định '; k[1] = 'KH'; t[1] = 'TH';
    for (const r of [p, k, t]) {
      r[19] = s.factory; r[20] = s.line; r[21] = s.brand; r[22] = s.style; r[23] = s.item || 'JKT';
      r[26] = s.qty || 1000; r[27] = s.sam; r[28] = s.hc;
    }
    for (const [d, v] of Object.entries(s.kh || {})) k[D0 + dates.indexOf(d)] = v;
    for (const [d, v] of Object.entries(s.th || {})) t[D0 + dates.indexOf(d)] = v;
    if (s.endTH) t[D0 + dates.indexOf(s.endTH)] = 'END';
    p[D0 + 1] = 99; // hàng lý thuyết không được tính
    rows.push(p, k, t);
  }
  return rows;
}

test('Ngày update lấy từ tên file', () => {
  assert.strictEqual(R.updateDateFromName('KH_UPDATE__APA1.20261031.xlsm'), '2026-10-31');
  assert.strictEqual(R.updateDateFromName('60f485ea-KH_UPDATE__APA2.20250930.xlsm'), '2025-09-30');
  assert.strictEqual(R.updateDateFromName('KH update.xlsm'), null);
});

test('Năm tài chính bắt đầu tháng 4: NTC 2026 = 4/2026–3/2027', () => {
  assert.strictEqual(R.fiscalYear('2026-04', 4), 2026);
  assert.strictEqual(R.fiscalYear('2027-03', 4), 2026);
  assert.strictEqual(R.fiscalYear('2026-03', 4), 2025);
  assert.strictEqual(R.fiscalYear('2026-12', 1), 2026);
  assert.deepStrictEqual(R.fiscalMonths(2026, 4).slice(0, 2), ['2026-04', '2026-05']);
  assert.strictEqual(R.fiscalMonths(2026, 4)[11], '2027-03');
  assert.strictEqual(R.fiscalLabel(2026, 4), 'NTC 2026 (4/2026–3/2027)');
});

test('Chuẩn hóa nhãn hàng và cụm chuyền', () => {
  assert.strictEqual(R.normBrand('hELLY  HANSEN'), 'HELLY HANSEN');
  assert.strictEqual(R.normBrand('Helly Hansen'), 'HELLY HANSEN');
  assert.strictEqual(R.normBrand('ALLTERAIN'), 'ALLTERRAIN');
  assert.strictEqual(R.normBrand('TNF', R.parseAliasText('tnf = THE NORTH FACE')), 'THE NORTH FACE');
  assert.strictEqual(R.normBrand(''), '(CHƯA GHI)');
  assert.strictEqual(R.lineGroup('4AB2'), '4');
  assert.strictEqual(R.lineGroup('12A'), '12');
});

test('Đọc sheet: KH/TH theo tháng, ngày làm việc, bỏ hàng lý thuyết và mã master', () => {
  const rows = makeSheet({
    from: '2026-09-28', to: '2026-10-04', off: ['2026-10-04'],
    styles: [
      { factory: 'APIC1A', line: '1A', brand: 'master', style: 'master', sam: 9000, hc: 40, kh: { '2026-09-28': 50 } },
      { factory: 'APIC1A', line: '1A', brand: 'TNF', style: 'NP1', sam: 7200, hc: 30,
        kh: { '2026-09-29': 10, '2026-09-30': 20, '2026-10-01': 30 }, th: { '2026-09-29': 8, '2026-09-30': 22, '2026-10-01': 25 }, endTH: '2026-10-02' }
    ]
  });
  const p = R.parseRows(rows, { name: 'KH_UPDATE__APA1.20261003.xlsm' });
  assert.strictEqual(p.factory, 'APIC1A');
  assert.strictEqual(p.updateDate, '2026-10-03');
  assert.strictEqual(p.styleCount, 1);
  const sep = p.facts.find((f) => f.m === '2026-09'), oct = p.facts.find((f) => f.m === '2026-10');
  assert.deepStrictEqual([sep.kh, sep.th, oct.kh, oct.th], [30, 30, 30, 25]);
  assert.strictEqual(sep.sah, 30 * 2); // 7200 giây = 2 giờ chuẩn/sp
  assert.strictEqual(sep.start, 1);
  assert.strictEqual(oct.start, 0);
  assert.strictEqual(oct.end, 1); // END ở 2/10 → xong ngày 1/10
  assert.strictEqual(sep.pdT, 60); // 2 ngày × 30 người
  // Tháng 10 có 31 ngày; 1 ngày nghỉ "x" trong file (4/10), các ngày ngoài file trừ Chủ nhật
  assert.strictEqual(p.months['2026-10'].off, 1);
  assert.strictEqual(p.months['2026-10'].full, false);
});

test('Chuyền tách A1, A2 gộp về chuyền gốc', () => {
  assert.strictEqual(R.baseLine('4A1'), '4A');
  assert.strictEqual(R.baseLine('4a2'), '4A');
  assert.strictEqual(R.baseLine('12B3'), '12B');
  assert.strictEqual(R.baseLine('4AB2'), '4AB');
  assert.strictEqual(R.baseLine('15A'), '15A');
});

test('Lao động: chuyền gốc tính 1 lần mỗi ngày; chuyền tách cộng lại; ngày chuyển mã không đếm trùng', () => {
  const rows = makeSheet({
    from: '2026-10-01', to: '2026-10-03',
    styles: [
      { factory: 'APIC1A', line: '2A', brand: 'TNF', style: 'A', sam: 3600, hc: 32,
        kh: { '2026-10-01': 30 }, th: { '2026-10-01': 30, '2026-10-02': 10 } },
      { factory: 'APIC1A', line: '2A', brand: 'GOLDWIN', style: 'B', sam: 3600, hc: 30, kh: { '2026-10-01': 10 } },
      { factory: 'APIC1A', line: '2A1', brand: 'TNF', style: 'C', sam: 3600, hc: 16, th: { '2026-10-02': 20, '2026-10-03': 20 } },
      { factory: 'APIC1A', line: '2A2', brand: 'TNF', style: 'D', sam: 3600, hc: 16, th: { '2026-10-03': 20 } }
    ]
  });
  const p = R.parseRows(rows, { name: 'x.20261031.xlsm' });
  const by = (s) => p.facts.find((f) => f.style === s);
  assert.deepStrictEqual([...new Set(p.facts.map((f) => f.line))], ['2A']);
  assert.strictEqual(by('C').lineRaw, '2A1');
  // Ngày 1: chỉ A → 32 người. Ngày 2: A (2A) và C (2A1) cùng chạy → vẫn 32 người, chia 10:20.
  // Ngày 3: C (2A1) và D (2A2) → 16 + 16 = 32 người.
  assert.ok(Math.abs(by('A').pdT - (32 + 32 * 10 / 30)) < 1e-9);
  assert.ok(Math.abs(by('C').pdT - (32 * 20 / 30 + 16)) < 1e-9);
  assert.strictEqual(by('D').pdT, 16);
  const total = p.facts.reduce((t, f) => t + f.pdT, 0), lineDays = p.facts.reduce((t, f) => t + f.ldT, 0);
  assert.ok(Math.abs(total - 96) < 1e-9);
  assert.ok(Math.abs(lineDays - 3) < 1e-9); // 3 ngày × 1 chuyền gốc
  // KH ngày 1: A 30 sp và B 10 sp cùng trên 2A → 32 người chia 3:1
  assert.strictEqual(by('A').pdK, 24);
  assert.strictEqual(by('B').pdK, 8);
});

// Ba file của 1 xưởng: chốt tháng 9 (30/9), giữa tháng 10 (15/10), chốt tháng 10 (31/10)
function threeFiles() {
  const st = (kh, th) => ({ factory: 'APIC1A', line: '1A', brand: 'TNF', style: 'NP1', sam: 3600, hc: 30, kh, th });
  const f1 = R.parseRows(makeSheet({ from: '2026-09-01', to: '2026-11-30', styles: [st({ '2026-09-10': 100, '2026-10-10': 200, '2026-11-10': 300 }, { '2026-09-10': 90 })] }), { name: 'KH_UPDATE__APA1.20260930.xlsm' });
  const f2 = R.parseRows(makeSheet({ from: '2026-09-01', to: '2026-11-30', styles: [st({ '2026-09-10': 100, '2026-10-10': 210, '2026-11-10': 300 }, { '2026-09-10': 95, '2026-10-10': 50 })] }), { name: 'KH_UPDATE__APA1.20261015.xlsm' });
  const f3 = R.parseRows(makeSheet({ from: '2026-09-01', to: '2026-11-30', styles: [st({ '2026-09-10': 100, '2026-10-10': 220, '2026-11-10': 350 }, { '2026-09-10': 95, '2026-10-10': 230 })] }), { name: 'KH_UPDATE__APA1.20261031.xlsm' });
  return [f3, f1, f2]; // thứ tự tải lên không quan trọng
}

test('Gộp file: mỗi tháng lấy file chốt sớm nhất sau cuối tháng; KH đầu tháng từ file trước ngày 1', () => {
  const ds = R.buildDataset(threeFiles());
  assert.strictEqual(ds.meta['APIC1A|2026-09'].srcDate, '2026-09-30');
  assert.strictEqual(ds.meta['APIC1A|2026-09'].status, 'chot');
  assert.strictEqual(ds.meta['APIC1A|2026-10'].srcDate, '2026-10-31');
  assert.strictEqual(ds.meta['APIC1A|2026-10'].openDate, '2026-09-30');
  assert.strictEqual(ds.meta['APIC1A|2026-11'].status, 'kh');
  const s = R.summarize(ds, { month: '2026-10', startMonth: 4 });
  assert.strictEqual(s.totals.cur.th, 230);
  assert.strictEqual(s.totals.cur.kh, 220);
  assert.strictEqual(s.totals.cur.khOpen, 200);
  assert.strictEqual(s.totals.prev.th, 90); // tháng 9 lấy theo file 30/9
  assert.strictEqual(s.totals.ytd.th, 320);
  const nov = R.summarize(ds, { month: '2026-11', startMonth: 4 });
  assert.strictEqual(nov.totals.cur.kh, 350);
  assert.strictEqual(nov.totals.cur.khOpen, 300);
});

test('Tháng chưa có file chốt: lấy file mới nhất, trạng thái đang chạy', () => {
  const files = threeFiles().filter((f) => f.updateDate !== '2026-10-31');
  const ds = R.buildDataset(files);
  assert.strictEqual(ds.meta['APIC1A|2026-10'].status, 'dang');
  assert.strictEqual(ds.meta['APIC1A|2026-10'].srcDate, '2026-10-15');
});

test('So sánh cùng kỳ năm trước, lũy kế năm tài chính và bảng theo nhãn hàng', () => {
  const mk = (name, from, to, styles) => R.parseRows(makeSheet({ from, to, styles }), { name });
  const s = (style, brand, kh, th) => ({ factory: 'APIC1B', line: '7A', brand, style, sam: 3600, hc: 30, kh, th });
  const files = [
    mk('KH_UPDATE__APA2.20251031.xlsm', '2025-04-01', '2025-10-31', [
      s('O1', 'TNF', { '2025-04-05': 100, '2025-10-05': 400 }, { '2025-04-05': 100, '2025-10-05': 400 }),
      s('O2', 'GOLDWIN', { '2025-10-06': 100 }, { '2025-10-06': 100 })
    ]),
    mk('KH_UPDATE__APA2.20261031.xlsm', '2026-04-01', '2026-10-31', [
      s('N1', 'TNF', { '2026-04-05': 150, '2026-09-05': 100, '2026-10-05': 500 }, { '2026-04-05': 150, '2026-09-05': 100, '2026-10-05': 450 }),
      s('N2', 'ALLTERAIN', { '2026-10-06': 50 }, { '2026-10-06': 50 })
    ])
  ];
  const ds = R.buildDataset(files);
  const r = R.summarize(ds, { month: '2026-10', startMonth: 4, dim: 'brand' });
  assert.strictEqual(r.totals.cur.th, 500);
  assert.strictEqual(r.totals.yoy.th, 500);
  assert.strictEqual(r.totals.prev.th, 100);
  assert.deepStrictEqual(r.periods.ytd, ['2026-04', '2026-09', '2026-10']);
  assert.deepStrictEqual(r.periods.ytdPrev, ['2025-04', '2025-10']);
  assert.strictEqual(r.totals.ytd.th, 750);
  assert.strictEqual(r.totals.ytdPrev.th, 600);
  const tnf = r.rows.find((x) => x.key === 'TNF');
  assert.strictEqual(tnf.cur.th, 450);
  assert.strictEqual(tnf.yoy.th, 400);
  assert.ok(Math.abs(R.pct(tnf.cur.th, tnf.yoy.th) - 0.125) < 1e-9);
  assert.ok(r.rows.find((x) => x.key === 'ALLTERRAIN')); // tên viết sai đã được gộp
  assert.ok(r.rows.find((x) => x.key === 'GOLDWIN')); // có số cùng kỳ dù tháng này không có
  // Năm tài chính
  const fy = R.fiscal(ds, { startMonth: 4, dim: 'brand' });
  assert.deepStrictEqual(fy.years.map((y) => [y.fy, y.v.th]), [[2025, 600], [2026, 750]]);
  const tnfRow = fy.matrix.find((x) => x.key === 'TNF');
  assert.deepStrictEqual(tnfRow.cells.map((c) => c.th), [500, 700]);
  // Cùng kỳ của NTC 2026: các tháng có TH (4, 9, 10/2026) lùi 12 tháng → 4/2025 và 10/2025 có số
  assert.deepStrictEqual(fy.years[1].sameMonths, ['2025-04', '2025-10']);
  assert.strictEqual(fy.years[1].same.th, 600);
  assert.strictEqual(fy.years[0].same, null);
  assert.strictEqual(fy.years[1].done.kh, 800); // KH của các tháng đã có TH
  assert.strictEqual(tnfRow.last.th, 700);
  assert.strictEqual(tnfRow.same.th, 500);
  // Lao động bình quân/ngày: tháng 10/2026 có 2 ngày chạy × 30 người, chia cho số ngày làm việc
  const wd = ds.meta['APIC1B|2026-10'].wd;
  assert.ok(Math.abs(r.totals.cur.labor - 60 / wd) < 1e-9);
  assert.strictEqual(r.totals.cur.productivity, 500 / 60);
  assert.strictEqual(r.totals.cur.efficiency, 500 / (60 * 8));
});

test('Gộp hai xưởng: lao động bình quân/ngày cộng theo xưởng', () => {
  const mk = (name, factory, hc) => R.parseRows(makeSheet({
    from: '2026-10-01', to: '2026-10-31',
    styles: [{ factory, line: '1A', brand: 'TNF', style: factory + 'S', sam: 3600, hc, kh: { '2026-10-01': 10 }, th: { '2026-10-01': 10 } }]
  }), { name });
  const ds = R.buildDataset([mk('KH_UPDATE__APA1.20261031.xlsm', 'APIC1A', 30), mk('KH_UPDATE__APA2.20261031.xlsm', 'APIC1B', 20)]);
  const all = R.summarize(ds, { month: '2026-10', startMonth: 4 });
  const wd = ds.meta['APIC1A|2026-10'].wd;
  assert.ok(Math.abs(all.totals.cur.labor - 50 / wd) < 1e-9);
  assert.strictEqual(all.totals.cur.styles, 2);
  const one = R.summarize(ds, { month: '2026-10', startMonth: 4, factory: 'APIC1B' });
  assert.strictEqual(one.totals.cur.th, 10);
});

test('File update vào ngày làm việc cuối tháng (cuối tháng nghỉ lễ / Chủ nhật) vẫn chốt được tháng đó', () => {
  const st = (kh, th) => ({ factory: 'APIC1A', line: '1A', brand: 'TNF', style: 'S1', sam: 3600, hc: 30, kh, th });
  // 29/4–30/4/2023 nghỉ (đánh dấu x) → ngày làm việc cuối tháng 4 là 28/4
  const off = ['2023-04-29', '2023-04-30', '2023-05-01'];
  const fApr = R.parseRows(makeSheet({ from: '2023-04-01', to: '2023-06-30', off, styles: [st({ '2023-04-10': 100, '2023-05-10': 200 }, { '2023-04-10': 90 })] }), { name: 'KH UPDATE _APA1.20230428.xlsm' });
  const fMay = R.parseRows(makeSheet({ from: '2023-04-01', to: '2023-06-30', off, styles: [st({ '2023-04-10': 90, '2023-05-10': 200 }, { '2023-04-10': 90, '2023-05-10': 210 })] }), { name: 'KH UPDATE _APA1.20230531.xlsm' });
  assert.strictEqual(fApr.months['2023-04'].lastWork, '2023-04-28');
  const ds = R.buildDataset([fMay, fApr]);
  assert.strictEqual(ds.meta['APIC1A|2023-04'].srcDate, '2023-04-28');
  assert.strictEqual(ds.meta['APIC1A|2023-04'].status, 'chot');
  assert.strictEqual(ds.meta['APIC1A|2023-05'].srcDate, '2023-05-31');
  assert.strictEqual(ds.meta['APIC1A|2023-05'].openDate, '2023-04-28');
  const s = R.summarize(ds, { month: '2023-05', startMonth: 4 });
  assert.strictEqual(s.totals.cur.khOpen, 200);
  assert.strictEqual(s.totals.cur.th, 210);
});

test('Kiểm tra công thức tổng: phát hiện vùng dừng trước hàng dữ liệu cuối và tính lại đủ hàng', () => {
  const W = 20, rows = [];
  for (let i = 0; i < 12; i++) rows.push(new Array(W).fill(null));
  rows[0][1] = 'dự định '; rows[1][1] = 'KH'; rows[2][1] = 'TH';
  rows[4][1] = 2024; rows[5][1] = 3; rows[5][3] = 'TH';
  // 3 mã: KH/TH tháng chọn ở cột C, cờ chuyển mã ở cột M (chỉ số 12)
  [[100, 90], [200, 210], [50, 50]].forEach(([kh, th]) => {
    const p = new Array(W).fill(null), k = new Array(W).fill(null), t = new Array(W).fill(null);
    p[1] = 'dự định '; k[1] = 'KH'; t[1] = 'TH';
    p[2] = 0; k[2] = kh; t[2] = th; k[12] = 1; t[12] = 1;
    rows.push(p, k, t);
  });
  // giá trị đã lưu trong file: chỉ tính 2 mã đầu (hàng 13–18)
  rows[1][2] = 300; rows[2][2] = 300; rows[0][12] = 2;
  const formulas = {
    C2: 'SUMIF($B$13:$B$18,$B2,$C$13:$C$15)', C3: 'SUMIF($B$13:$B$21,$B3,$C$13:$C$21)',
    M1: 'SUMIF($B$13:$B$18,$D$6,M$13:M$18)'
  };
  const c = R.checkSummary(rows, formulas);
  assert.strictEqual(c.lastRow, 21);
  assert.strictEqual(c.month, '2024-03');
  assert.deepStrictEqual(c.issues.map((i) => i.cell + '@' + i.end), ['M1@18', 'C2@18']);
  assert.strictEqual(c.full.kh, 350);
  assert.strictEqual(c.full.th, 350);
  assert.strictEqual(c.full.changes, 3);
  assert.strictEqual(c.file.kh, 300);
  const ok = R.checkSummary(rows, { C2: 'SUMIF($B$13:$B$14871,$B2,$C$13:$C$14871)' });
  assert.strictEqual(ok.issues.length, 0);
});

test('Kiểm tra công thức tổng: SUMPRODUCT có vùng khác số hàng và ô đang lỗi', () => {
  const rows = [];
  for (let i = 0; i < 20; i++) rows.push(new Array(14).fill(null));
  rows[15][1] = 'KH'; rows[16][1] = 'TH';
  const c = R.checkSummary(rows, {
    I2: 'SUMPRODUCT($E13:$E15000,I13:I15000)/SUM($E13:$E15000)',
    K2: 'SUMPRODUCT($E13:$E15000,K13:K867)/SUM($E13:$E15000)'
  }, { K2: '#VALUE!', K3: '#VALUE!' });
  assert.deepStrictEqual(c.cells.map((x) => x.cell), ['K2', 'K3']);
  assert.match(c.cells[0].text, /không cùng số hàng/);
  assert.match(c.cells[0].text, /#VALUE!/);
  assert.strictEqual(c.issues.length, 0); // không phải ô chỉ tiêu chính của TORAY
});

test('Báo thiếu file cuối tháng của từng xưởng', () => {
  const st = (f) => ({ factory: f, line: '1A', brand: 'TNF', style: 'S', sam: 3600, hc: 30, kh: { '2023-05-10': 10 }, th: { '2023-05-10': 10 } });
  const off = ['2023-04-29', '2023-04-30'];
  const mk = (name, f) => R.parseRows(makeSheet({ from: '2023-04-01', to: '2023-10-31', off, styles: [st(f)] }), { name });
  const a = ['20230428', '20230531', '20230731', '20230815', '20231003'].map((d) => mk('KH UPDATE _APA1.' + d + '.xlsm', 'APIC1A'));
  const b = ['20230428', '20230531'].map((d) => mk('KH UPDATE _APA2.' + d + '.xlsm', 'APIC1B'));
  assert.strictEqual(R.closesMonth(a[0]), '2023-04'); // 28/4: ngày làm việc cuối tháng 4
  assert.strictEqual(R.closesMonth(a[3]), null); // 15/8: file giữa tháng
  assert.strictEqual(R.closesMonth(a[4]), '2023-09'); // 3/10: file cuối tháng 9
  // 29/8 khi lịch trong file coi 31/8 là ngày làm việc: vẫn là file cuối tháng 8 (trong 3 ngày cuối tháng)
  assert.strictEqual(R.closesMonth({ updateDate: '2026-08-29', months: { '2026-08': { full: true, lastWork: '2026-08-31' } } }), '2026-08');
  assert.strictEqual(R.closesMonth({ updateDate: '2026-08-27', months: { '2026-08': { full: true, lastWork: '2026-08-31' } } }), null);
  const cov = R.coverage(a.concat(b));
  const A = cov.find((x) => x.factory === 'APIC1A'), B = cov.find((x) => x.factory === 'APIC1B');
  assert.deepStrictEqual(A.missing.map((x) => x.m), ['2023-06', '2023-08']);
  assert.deepStrictEqual(A.mid, ['2023-08-15']);
  assert.deepStrictEqual([A.first, A.last, A.expected, A.have], ['2023-04', '2023-09', 6, 4]);
  // APIC1B dừng ở tháng 5 → thiếu 6, 7, 8, 9 (so với tháng gần nhất của mọi xưởng)
  assert.deepStrictEqual(B.missing.map((x) => x.m), ['2023-06', '2023-07', '2023-08', '2023-09']);
});
