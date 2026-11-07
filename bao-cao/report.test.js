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
    if (s.aggLot != null) p[37] = s.aggLot; // 集約ロット (cột AL, hàng lý thuyết)
    if (s.learnTheo != null) p[41] = s.learnTheo; // 習熟日数 lý thuyết (AP)
    if (s.ltTheo != null) p[45] = s.ltTheo; // 縫製リードタイム lý thuyết (AT)
    for (const [d, v] of Object.entries(s.theo || {})) p[D0 + dates.indexOf(d)] = v;
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
  assert.strictEqual(nov.totals.cur.khOpen, 0); // file 15/10 cách ngày 1/11 quá 10 ngày: không dùng làm KH đầu tháng
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

test('工程分析時間 và 生産ロット bình quân tính như sheet TORAY', () => {
  const st = (style, qty, sam, aggLot, th, theo) => ({ factory: 'APIC1A', line: '1A', brand: 'TNF', style, qty, sam, hc: 30, aggLot, th, theo });
  const rows = makeSheet({
    from: '2026-10-01', to: '2026-10-31',
    styles: [
      st('A', 1000, 10000, 1200, { '2026-10-05': 50 }, { '2026-10-05': 40 }),
      st('B', 3000, 6000, 3000, { '2026-10-06': 80 }, { '2026-10-06': 70 }),
      // C: có lịch lý thuyết nhưng chưa có TH trong tháng → tính vào lot, không tính vào 工程分析時間
      st('C', 500, 20000, 600, {}, { '2026-10-20': 10 })
    ]
  });
  const p = R.parseRows(rows, { name: 'x.20261031.xlsm' });
  const v = R.derive(R.aggregate(p.facts.filter((f) => f.m === '2026-10')), 1);
  assert.strictEqual(v.avgSam, (1000 * 10000 + 3000 * 6000) / 4000);
  assert.strictEqual(v.avgLot, (1200 + 3000 + 600) / 3);
  assert.strictEqual(v.styles, 2); // mã chỉ có lịch lý thuyết không tính là mã sản xuất
});

// ---------- Phân tích ----------
const days = (from, vals) => Object.fromEntries(vals.map((v, i) => [R.addDays(from, i), v]));

test('Lượt may: mức ổn định, ngày lên năng suất, giờ chuẩn mất, lead time thực tế', () => {
  const th = days('2026-10-05', [10, 20, 40, 50, 50, 50, 48, 30]);
  const rows = makeSheet({
    from: '2026-10-01', to: '2026-10-31', off: ['2026-10-11'],
    styles: [{ factory: 'APIC1A', line: '1A', brand: 'TNF', style: 'S', sam: 3600, hc: 30, kh: th, th, endTH: '2026-10-13', learnTheo: 3.5, ltTheo: 7 }]
  });
  const p = R.parseRows(rows, { name: 'x.20261031.xlsm' });
  const r = p.runs[0];
  assert.strictEqual(r.stable, 50); // phân vị 75% của [10,20,40,50,50,50,48] (bỏ ngày cuối)
  assert.strictEqual(r.ramp, 4); // ngày thứ 4 đạt ≥ 45
  assert.strictEqual(r.loss, 40 + 30 + 10); // (50−10)+(50−20)+(50−40), 1 giờ chuẩn/sp
  assert.strictEqual(r.days, 8);
  assert.strictEqual(r.span, 7); // 5/10–12/10: 8 ngày lịch, trừ 11/10 là ngày nghỉ (dù có may) → 7 ngày làm việc
  assert.strictEqual(r.complete, true);
  assert.strictEqual(r.learnTheo, 3.5);
  assert.strictEqual(r.ltTheo, 7);
  assert.deepStrictEqual(r.curve.slice(0, 3), [0.2, 0.4, 0.8]);
});

test('Phân tích 1–3: cỡ lot, lên năng suất, lead time', () => {
  const st = (style, lot, sam, th, endTH, extra) => Object.assign({ factory: 'APIC1A', line: '1A', brand: 'TNF', style, sam, hc: 30, aggLot: lot, kh: th, th, endTH, theo: { '2026-10-02': 1 } }, extra);
  const rows = makeSheet({
    from: '2026-10-01', to: '2026-10-31',
    styles: [
      st('SMALL', 300, 3600, days('2026-10-05', [5, 10, 10, 10, 10, 10]), '2026-10-11', { learnTheo: 2, ltTheo: 3 }),
      st('BIG', 2500, 9000, days('2026-10-05', [20, 40, 60, 60, 60, 60, 60]), '2026-10-12', { line: '2A', learnTheo: 3, ltTheo: 10 })
    ]
  });
  const ds = R.buildDataset([R.parseRows(rows, { name: 'x.20261031.xlsm' })]);
  const o = { month: '2026-10', startMonth: 4, period: 'month' };
  const lot = R.lotEfficiency(ds, o);
  assert.deepStrictEqual(lot.rows.map((r) => r.band.label), ['Dưới 500 sp', 'Từ 2.000 sp']);
  assert.ok(lot.rows[1].v.efficiency > lot.rows[0].v.efficiency);
  const learn = R.learning(ds, o);
  assert.strictEqual(learn.total.n, 2);
  assert.strictEqual(learn.rows[0].v.ramp, 2); // SMALL (< 8.000 giây): ngày 2 đạt 10
  assert.strictEqual(learn.rows[1].v.ramp, 3); // BIG (8.000–12.000 giây): ngày 3 đạt 60
  assert.strictEqual(learn.total.slower, 0);
  const lt = R.leadTime(ds, o);
  // SMALL: 6 ngày may (5–10/10) so với lý thuyết 3 → chậm hơn 2 ngày cho phép
  assert.strictEqual(lt.rows[0].v.overRate, 1);
  assert.strictEqual(lt.rows[1].v.overRate, 0);
  assert.strictEqual(lt.worst[0].style, 'SMALL');
});

test('Phân tích 4–6: chuyển mã, độ ổn định kế hoạch, xếp hạng chuyền', () => {
  const mk = (name, styles) => R.parseRows(makeSheet({ from: '2026-09-01', to: '2026-11-30', styles }), { name });
  const st = (style, line, hc, kh, th, extra) => Object.assign({ factory: 'APIC1A', line, brand: 'TNF', style, sam: 3600, hc, kh, th }, extra);
  // File cuối tháng 9: KH tháng 10 cho A, B
  const f9 = mk('KH UPDATE _APA1.20260930.xlsm', [
    st('A', '1A', 30, { '2026-10-05': 100 }, {}), st('B', '1A', 30, { '2026-10-12': 100 }, {})
  ]);
  // File cuối tháng 10: A làm đủ, B không làm, C thêm vào; chuyền 2A chạy chậm hơn
  const f10 = mk('KH UPDATE _APA1.20261031.xlsm', [
    st('A', '1A', 30, { '2026-10-05': 100 }, days('2026-10-05', [60, 120, 120, 120])),
    st('C', '2A', 30, { '2026-10-06': 100 }, days('2026-10-06', [30, 60, 60]))
  ]);
  const ds = R.buildDataset([f9, f10]);
  const o = { month: '2026-10', startMonth: 4, period: 'month' };
  const plan = R.planStability(ds, o);
  assert.strictEqual(plan.total.khOpen, 200);
  assert.strictEqual(plan.total.th, 570);
  assert.strictEqual(plan.total.added, 1); // C
  assert.strictEqual(plan.total.dropped, 1); // B
  const co = R.changeover(ds, o);
  assert.strictEqual(co.total.starts, 2);
  assert.ok(co.total.loss > 0);
  const rank = R.lineRanking(ds, Object.assign({}, o));
  assert.deepStrictEqual(rank.rows.map((r) => r.key), ['APIC1A · 1A', 'APIC1A · 2A']);
  assert.strictEqual(rank.rows[0].rank, 1);
});

test('Phân tích 7: tải tháng tới so với năng lực 3 tháng đã chốt gần nhất', () => {
  const st = (kh, th) => ({ factory: 'APIC1A', line: '1A', brand: 'TNF', style: 'S', sam: 3600, hc: 10, kh, th });
  // mọi ngày trừ Chủ nhật (file thử không có hàng ngày nghỉ "x" nên Chủ nhật là ngày nghỉ)
  const isSun = (d) => new Date(d + 'T00:00:00Z').getUTCDay() === 0;
  const all = (from, to, v) => { const o = {}; for (let d = from; d <= to; d = R.addDays(d, 1)) if (!isSun(d)) o[d] = v; return o; };
  // Tháng 10: may mỗi ngày 40 sp × 1 giờ = 40 giờ chuẩn, 10 người × 8 giờ = 80 giờ công → hiệu suất 50%
  const th10 = all('2026-10-01', '2026-10-31', 40);
  const kh11 = all('2026-11-01', '2026-11-30', 60);
  const fl = R.parseRows(makeSheet({ from: '2026-10-01', to: '2026-11-30', styles: [st(Object.assign({}, th10, kh11), th10)] }), { name: 'KH UPDATE _APA1.20261031.xlsm' });
  const ds = R.buildDataset([fl]);
  const cap = R.capacityLoad(ds, {});
  const nov = cap.rows.find((r) => r.month === '2026-11');
  assert.ok(Math.abs(nov.efficiency - 0.5) < 1e-9);
  assert.ok(Math.abs(nov.labor - 10) < 1e-9); // 10 người mỗi ngày làm việc
  // cần 60 giờ chuẩn/ngày, năng lực 10 × 8 × 50% = 40 → tải 150%
  assert.ok(Math.abs(nov.load - 1.5) < 1e-9);
  assert.strictEqual(R.loadStatus(nov.load), 'qua-tai');
});

test('Tổng hợp toàn xưởng: chỉ tiêu theo tháng, cả năm và cùng kỳ năm trước', () => {
  const st = (factory, style, kh, th) => ({ factory, line: '1A', brand: 'TNF', style, sam: 3600, hc: 10, aggLot: 1000, kh, th });
  const mk = (name, factory, from, to, styles) => R.parseRows(makeSheet({ from, to, styles }), { name });
  const files = [
    mk('KH UPDATE _APA1.20251031.xlsm', 'APIC1A', '2025-10-01', '2025-10-31', [st('APIC1A', 'O', { '2025-10-06': 100 }, { '2025-10-06': 100 })]),
    mk('KH UPDATE _APA1.20261031.xlsm', 'APIC1A', '2026-10-01', '2026-10-31', [st('APIC1A', 'A', { '2026-10-06': 200 }, { '2026-10-06': 150 })]),
    mk('KH UPDATE _APA2.20261031.xlsm', 'APIC1B', '2026-10-01', '2026-10-31', [st('APIC1B', 'B', { '2026-10-07': 100 }, { '2026-10-07': 90 })])
  ];
  const ds = R.buildDataset(files);
  const y = R.overviewYear(ds, { startMonth: 4 }, 2026);
  const oct = y.cols.find((c) => c.month === '2026-10').v;
  assert.strictEqual(oct.th, 240); // gộp 2 xưởng
  assert.strictEqual(oct.kh, 300);
  assert.strictEqual(oct.styles, 2);
  assert.strictEqual(oct.lines, 2); // 1A của mỗi xưởng
  assert.strictEqual(oct.rows, 2);
  assert.strictEqual(y.cols.find((c) => c.month === '2026-04').v, null); // tháng chưa có dữ liệu
  assert.deepStrictEqual(y.doneMonths, ['2026-10']);
  assert.deepStrictEqual(y.sameMonths, ['2025-10']);
  assert.strictEqual(y.same.th, 100);
  const one = R.overviewYear(ds, { startMonth: 4, factory: 'APIC1B' }, 2026);
  assert.strictEqual(one.done.th, 90);
  const ys = R.overviewYears(ds, { startMonth: 4 });
  assert.deepStrictEqual(ys.map((x) => [x.fy, x.v.th]), [[2025, 100], [2026, 240]]);
});

test('Cam kết TI: mục tiêu năng suất +5%/năm so với NTC 2025, tỷ lệ hàng khó và lot nhỏ', () => {
  const isSun = (d) => new Date(d + 'T00:00:00Z').getUTCDay() === 0;
  const all = (from, to, v) => { const o = {}; for (let d = from; d <= to; d = R.addDays(d, 1)) if (!isSun(d)) o[d] = v; return o; };
  const st = (style, sam, aggLot, line, th, kh) => ({ factory: 'APIC1A', line, brand: 'TNF', style, sam, hc: 10, aggLot, qty: aggLot, th, kh: kh || th });
  // NTC 2025 (tháng 10/2025): 1 mã 3.600 giây, 40 sp/ngày, 10 người → hiệu suất 50%
  const f25 = R.parseRows(makeSheet({ from: '2025-10-01', to: '2025-10-31', styles: [st('B', 3600, 1000, '1A', all('2025-10-01', '2025-10-31', 40))] }), { name: 'KH UPDATE _APA1.20251031.xlsm' });
  // NTC 2026 (tháng 10/2026): mã thường 900 sp (lot 2.000), mã khó 25.000 giây 50 sp, mã lot nhỏ 50 sp (lot 400 trên chuyền 2A)
  const f26 = R.parseRows(makeSheet({ from: '2026-10-01', to: '2026-11-30', styles: [
    st('N', 3600, 2000, '1A', { '2026-10-05': 900 }, { '2026-10-05': 900, '2026-11-05': 1000 }),
    st('H', 25000, 1000, '3A', { '2026-10-06': 50 }),
    st('S', 3600, 400, '2A', { '2026-10-07': 50 })
  ] }), { name: 'KH UPDATE _APA1.20261031.xlsm' });
  const ds = R.buildDataset([f25, f26]);
  const t = R.contractTracking(ds, { startMonth: 4 }, {});
  assert.ok(Math.abs(t.benchmark.value - 0.5) < 1e-9);
  assert.deepStrictEqual(t.years.map((y) => y.fy), [2026, 2027, 2028]);
  assert.ok(Math.abs(t.years[0].target - 0.525) < 1e-9); // 0,5 × 1,05
  assert.ok(Math.abs(t.years[2].target - 0.5 * 1.05 ** 3) < 1e-9);
  const lin = R.contractTracking(ds, { startMonth: 4 }, { compound: false });
  assert.ok(Math.abs(lin.years[2].target - 0.5 * 1.15) < 1e-9);
  const y = t.years[0];
  assert.deepStrictEqual(y.doneMonths, ['2026-10']);
  assert.strictEqual(y.th, 1000);
  assert.strictEqual(y.hardShare, 0.05); // 50 / 1.000
  assert.strictEqual(y.smallShare, 0.05);
  // Dự báo: 1.000 TH tháng 10 + 1.000 KH tháng 11 (mã thường) → hàng khó 50 / 2.000
  assert.strictEqual(y.forecast.hardShare, 0.025);
  assert.strictEqual(y.forecast.hardRoom, 0.02 * 2000 - 50); // −10: đã vượt dù tháng 11 không làm hàng khó
  assert.strictEqual(t.years[1].forecast.hardShare, null); // NTC 2027 chưa bắt đầu
  const L = R.contractLists(y);
  assert.deepStrictEqual(L.hard.map((h) => h.style), ['H']);
  assert.deepStrictEqual(L.small.map((h) => h.style + '@' + h.line), ['S@2A']);
});

test('Hàng TI / APIC tự bán: so khớp danh sách nhãn hàng TI', () => {
  const m = R.tiMatcher(null);
  for (const b of ['TNF', 'HELLY HANSEN', 'Helly Hansen', 'N/W', 'Neutral Works', 'DESCSENTE', 'Woorich', 'ALLTERRAIN81', 'YONEX WC', 'LL.BEAN', '241', 'Goldwin', 'UMBRO', 'CCC', 'Inov-8', "B'S", 'GTO'])
    assert.strictEqual(m(b), 'TI', b);
  for (const b of ['RIACHUELO', 'RICHLO']) assert.strictEqual(m(b), 'APIC', b);
  assert.strictEqual(R.tiMatcher('TNF\nGOLDWIN')('MONTBELL'), 'APIC');
});

test('Sử dụng chuyền: nửa chuyền A/B, chuyền ghép AB, chuyền tách; giai đoạn trống và may hàng APIC', () => {
  // 5–10/10/2026 (Thứ 2–Thứ 7). 1A: TI 5–7/10, APIC 8–9/10, trống 10/10. 1B: chuyền ghép 1AB may TI 5–6/10, 1B1 (nửa) TI 7/10.
  const st = (style, line, brand, th) => ({ factory: 'APIC1A', line, brand, style, sam: 3600, hc: 30, aggLot: 1000, th, kh: th });
  const rows = makeSheet({ from: '2026-10-01', to: '2026-10-31', styles: [
    st('T1', '1A', 'TNF', { '2026-10-05': 50, '2026-10-06': 50, '2026-10-07': 50 }),
    st('A1', '1A', 'RIACHUELO', { '2026-10-08': 40, '2026-10-09': 40 }),
    st('T2', '1AB', 'GOLDWIN', { '2026-10-05': 100, '2026-10-06': 100 }),
    st('T3', '1B1', 'TNF', { '2026-10-07': 20 })
  ] });
  const p = R.parseRows(rows, { name: 'KH UPDATE _APA1.20261031.xlsm' });
  assert.deepStrictEqual(Object.keys(p.lineDays.lines).sort(), ['1A', '1B']);
  const ds = R.buildDataset([p]);
  const u = R.lineUsage(ds, { startMonth: 4 }, ['2026-10']);
  const a = u.byMonth[0];
  // tháng 10/2026: 31 ngày, Chủ nhật nghỉ (file thử không có hàng "x") → 27 ngày làm việc × 2 nửa chuyền
  assert.strictEqual(a.wd, 27);
  assert.strictEqual(a.total, 54);
  // 1A: TI 5,6 (chia với 1AB) … ngày 5–6: 1A có cả T1 và 1AB → TI cả ngày; 7/10 TI; 8–9/10 APIC → TI 3, APIC 2
  // 1B: 5–6/10 TI (1AB), 7/10 nửa chuyền 1B1 → 0,5 → TI 2,5
  assert.ok(Math.abs(a.ti - 5.5) < 1e-9);
  assert.ok(Math.abs(a.apic - 2) < 1e-9);
  assert.ok(Math.abs(a.idle - (54 - 7.5)) < 1e-9);
  const apic = u.periods.filter((x) => x.kind === 'apic');
  assert.deepStrictEqual(apic.map((x) => [x.line, x.from, x.to, x.days, x.brands]), [['1A', '2026-10-08', '2026-10-09', 2, 'RIACHUELO']]);
  const idle1B = u.periods.filter((x) => x.kind === 'idle' && x.line === '1B');
  assert.deepStrictEqual(idle1B.map((x) => x.from), ['2026-10-01', '2026-10-07']); // 1–3/10 trống; 7/10 nửa chuyền 1B2 trống
  assert.strictEqual(idle1B[1].days, 21.5); // 7/10: 0,5 + 21 ngày làm việc trống từ 8/10 đến 31/10

  // Cách tính của TI: SL hàng TI ÷ chuyền-ngày TI
  const tp = R.tiProductivity(ds, { startMonth: 4 }, ['2026-10']);
  assert.strictEqual(tp.th, 150 + 200 + 20);
  assert.ok(Math.abs(tp.perLineDay - 370 / 5.5) < 1e-9);
  assert.ok(Math.abs(tp.per8h - 370 / 5.5 * (1 - 0.059)) < 1e-9);
});

test('File APIC負担金額 của TI: đọc khối tháng, đối chiếu với chuyền trống / may hàng APIC', () => {
  const S = (iso) => serial(iso);
  const rows = [
    [null, 'APIC負担金額_総括'],
    [null, 'ﾗｲﾝ', 'Part', '人数', '部課', '縫製開始', '縫製エンド', 'ライン稼働単価', '日数', '金額', 'APIC負担比率', '請求額'],
    [null, 1, 'A', 32, '1課', S('2026-09-20'), S('2026-10-09'), 1000, 99, 99000, 0.5, 49500],
    [],
    [null, 'APIC負担金額 10月度'],
    [null, 'ﾗｲﾝ', 'Part', '人数', '部課', '縫製開始', '縫製エンド', 'ライン稼働単価', '日数', '金額', 'APIC負担比率', '請求額'],
    [null, 1, 'A', 32, '1課', S('2026-10-08'), S('2026-10-09'), 1000, 2, 2000, 0.4, 800, 'APIC自販分　生産期間'],
    [null, 1, 'B1', 16, '1課', S('2026-10-07'), S('2026-10-07'), 500, 1, 500, 0.4, 200],
    [null, 'TTL', null, null, null, null, null, null, 3, 2500, null, 1000]
  ];
  assert.ok(R.isBillingSheet(rows));
  const b = R.parseBilling(rows, { name: 'x.xlsx' });
  assert.strictEqual(b.month, '2026-10'); // lấy khối "10月度", bỏ khối 総括
  assert.strictEqual(b.rows.length, 2);
  assert.strictEqual(b.charge, 1000);
  // Dữ liệu chuyền: 1A may TI 1–7/10, APIC 8–9/10; 1B chỉ nửa 1B2 may ngày 7/10 (1B1 trống)
  const st = (style, line, brand, th) => ({ factory: 'APIC1A', line, brand, style, sam: 3600, hc: 30, aggLot: 1000, th, kh: th });
  const all = (from, to) => { const o = {}; for (let d = from; d <= to; d = R.addDays(d, 1)) if (new Date(d + 'T00:00:00Z').getUTCDay() !== 0) o[d] = 30; return o; };
  const sheet = makeSheet({ from: '2026-10-01', to: '2026-10-31', styles: [
    st('T1', '1A', 'TNF', Object.assign(all('2026-10-01', '2026-10-07'), all('2026-10-10', '2026-10-31'))),
    st('A1', '1A', 'RIACHUELO', { '2026-10-08': 40, '2026-10-09': 40 }),
    st('T2', '1B', 'TNF', Object.assign(all('2026-10-01', '2026-10-06'), all('2026-10-08', '2026-10-31'))),
    st('T3', '1B2', 'TNF', { '2026-10-07': 20 })
  ] });
  const ds = R.buildDataset([R.parseRows(sheet, { name: 'KH UPDATE _APA1.20261031.xlsm' })]);
  const c = R.billingCompare(ds, b, 4);
  assert.ok(c.full && c.dataOk);
  const by = Object.fromEntries(c.rows.map((r) => [r.line, r]));
  assert.strictEqual(by['1A'].tiDays, 2);
  assert.strictEqual(by['1A'].apic, 2);
  assert.strictEqual(by['1A'].diff, 0);
  assert.strictEqual(by['1B'].tiDays, 0.5); // B1 = nửa chuyền
  assert.strictEqual(by['1B'].idle, 0.5);
  assert.strictEqual(by['1B'].diff, 0);
  assert.strictEqual(by['1A'].tiSelf, 2); // dòng ghi APIC自販分
  assert.strictEqual(by['1B'].tiSelf, 0);
  // File NHABE: mọi dòng là hàng tự bán
  const c2 = R.billingCompare(ds, Object.assign({}, b, { name: '2025_4 APIC_NHABE 250428.xlsx' }), 4);
  assert.strictEqual(c2.rows.find((r) => r.line === '1B').tiSelf, 0.5);
});

test('Chuyền mở mới: chỉ tính từ ngày có sản lượng đầu tiên; khai báo tay thay ngày tự nhận', () => {
  const isSun = (d) => new Date(d + 'T00:00:00Z').getUTCDay() === 0;
  const all = (from, to) => { const o = {}; for (let d = from; d <= to; d = R.addDays(d, 1)) if (!isSun(d)) o[d] = 30; return o; };
  const st = (style, line, th) => ({ factory: 'APIC1A', line, brand: 'TNF', style, sam: 3600, hc: 30, aggLot: 1000, th, kh: th });
  const fl = R.parseRows(makeSheet({ from: '2026-08-01', to: '2026-10-31', styles: [
    st('A', '1A', all('2026-08-01', '2026-10-31')), st('B', '1B', all('2026-08-01', '2026-10-31')), st('N', '15B', all('2026-10-01', '2026-10-31'))
  ] }), { name: 'KH UPDATE _APA1.20261031.xlsm' });
  const ds = R.buildDataset([fl]);
  assert.strictEqual(ds.lineLife['15B'].from, '2026-10-01');
  assert.ok(ds.lineLife['15B'].opened);
  assert.strictEqual(ds.lineLife['1A'].from, '2026-08-01');
  const u = R.lineUsage(ds, { startMonth: 4 }, ['2026-09', '2026-10']);
  assert.deepStrictEqual(u.byMonth.map((a) => a.lines), [2, 3]);
  assert.strictEqual(u.byMonth[0].idle, 0); // 15B chưa mở: không tính trống tháng 9
  const ds2 = R.buildDataset([fl], { lineOverrides: '15B 15/9/2026' });
  assert.strictEqual(ds2.lineLife['15B'].from, '2026-09-15');
  const u2 = R.lineUsage(ds2, { startMonth: 4 }, ['2026-09']);
  assert.strictEqual(u2.byMonth[0].idle, 14); // 15–30/9: 16 ngày trừ 2 Chủ nhật = 14 ngày làm việc trống
});

test('Biểu đồ Excel: XML có chuỗi số, tham chiếu ô, cột chồng, trục cột từ 0', () => {
  const C = require('./xlsx-charts.js');
  const xml = C.chartXml({
    title: 'Chuyền-ngày', stacked: true, yFmt: '#,##0',
    cats: { ref: "'Bieu do'!$A$5:$A$6", values: ['5/2026', '6/2026'] },
    series: [{ name: 'Hàng TI', ref: "'Bieu do'!$B$5:$B$6", values: [686, null], color: '2A78D6' },
      { name: 'Giới hạn', ref: "'Bieu do'!$C$5:$C$6", values: [1, 1], color: '5B6482', type: 'line', dash: true }]
  });
  assert.match(xml, /<c:grouping val="stacked"\/>/);
  assert.match(xml, /<c:f>'Bieu do'!\$B\$5:\$B\$6<\/c:f>/);
  assert.match(xml, /<c:pt idx="0"><c:v>686<\/c:v><\/c:pt>/);
  assert.doesNotMatch(xml, /<c:pt idx="1"><c:v>null/); // ô trống không ghi
  assert.match(xml, /<c:lineChart>/);
  assert.match(xml, /<a:prstDash val="dash"\/>/);
  assert.match(xml, /<c:min val="0"\/>/);
});

test('Tỷ lệ đạt so KH đầu tháng nhiều tháng: chỉ tính TH của các tháng có KH đầu tháng', () => {
  const v = R.derive(R.aggregate([
    { f: 'A', m: '2026-09', th: 100, kh: 100, open: false },
    { f: 'A', m: '2026-10', th: 200, kh: 200, open: false },
    { f: 'A', m: '2026-10', khOpen: 180, open: true }
  ]), 2);
  assert.strictEqual(v.rateOpen, 200 / 180); // tháng 9 không có KH đầu tháng nên không tính
});
