// Chạy: node --test workflow/
// Đối chiếu lịch tính ra với các cột N1..N27 trong file
// "Rà soát lại thời gian cần để vào chuyền - 20260722.xlsx" (tính liên tục, không bỏ Chủ nhật).
const test = require('node:test');
const assert = require('node:assert');
const S = require('./schedule.js');

const ETA = '2026-11-01'; // = N1
const N = (date) => S.diffDays(ETA, date) + 1;

function plan(seconds, isNew, extra) {
  return S.computeSchedule(Object.assign({
    eta: ETA, base: 'port', seconds, isNew, inan: true, theu: true, skipSunday: false
  }, extra));
}

function spans(s) {
  const out = {};
  s.rows.forEach((r) => { out[r.id] = [N(r.start), N(r.end)]; });
  return out;
}

test('Dưới 8000 giây, mã mới: khớp sheet "Dưới 8000"', () => {
  const s = plan(7500, true);
  assert.deepStrictEqual(spans(s), {
    xnk: [1, 7], khui: [8, 8], qc: [9, 9], test: [9, 9], xa: [10, 10], rap: [10, 10],
    sodo: [11, 11], cat: [12, 13], caitien: [12, 14], theu: [12, 12], inan: [12, 15],
    hop: [16, 17], trienkhai: [18, 22]
  });
  assert.strictEqual(N(s.milestones.btpReady), 16); // "←Có tại nhà máy"
  assert.strictEqual(N(s.milestones.lineStart), 23); // "←Ra chuyền"
});

test('Dưới 8000 giây, hàng lặp lại: ra chuyền N21', () => {
  assert.strictEqual(N(plan(7500, false).milestones.lineStart), 21);
});

test('8000-12000 giây: khớp sheet "8000-12000"', () => {
  const s = plan(10000, true);
  const sp = spans(s);
  assert.deepStrictEqual(sp.rap, [10, 11]);
  assert.deepStrictEqual(sp.sodo, [11, 12]);
  assert.deepStrictEqual(sp.cat, [13, 14]);
  assert.deepStrictEqual(sp.caitien, [13, 15]);
  assert.deepStrictEqual(sp.inan, [13, 16]);
  assert.strictEqual(N(s.milestones.btpReady), 17);
  assert.deepStrictEqual(sp.hop, [17, 18]);
  assert.deepStrictEqual(sp.trienkhai, [19, 25]);
  assert.strictEqual(N(s.milestones.lineStart), 26);
  assert.strictEqual(N(plan(10000, false).milestones.lineStart), 24);
});

test('Trên 12000 giây: khớp sheet "Trên 12000"', () => {
  assert.deepStrictEqual(spans(plan(13000, true)).trienkhai, [18, 25]);
  assert.strictEqual(N(plan(13000, true).milestones.lineStart), 26);
  assert.strictEqual(N(plan(13000, false).milestones.lineStart), 25);
});

test('Ranh giới mức SAH', () => {
  assert.strictEqual(S.bandIndex(7999), 0);
  assert.strictEqual(S.bandIndex(8000), 1);
  assert.strictEqual(S.bandIndex(12000), 1);
  assert.strictEqual(S.bandIndex(12001), 2);
});

test('Không in, không thêu: BTP có ngay sau cắt', () => {
  const s = plan(7500, true, { inan: false, theu: false });
  assert.strictEqual(N(s.milestones.btpReady), 14);
  assert.strictEqual(N(s.milestones.lineStart), 21);
  assert.ok(!s.rows.some((r) => r.id === 'inan' || r.id === 'theu'));
});

test('Vải đã về nhà máy: bỏ bước XNK, khui thùng ngay ngày ETA', () => {
  const s = plan(7500, true, { base: 'factory' });
  assert.strictEqual(s.rows[0].id, 'khui');
  assert.strictEqual(s.rows[0].start, ETA);
  assert.strictEqual(N(s.milestones.lineStart), 16);
});

test('Bỏ Chủ nhật: bước tại nhà máy không rơi vào Chủ nhật', () => {
  const s = plan(7500, true, { skipSunday: true, holidays: ['2026-11-10'] });
  s.rows.filter((r) => !r.calendar).forEach((r) => {
    assert.notStrictEqual(S.weekday(r.start), 0, r.id + ' bắt đầu Chủ nhật');
    assert.notStrictEqual(r.start, '2026-11-10');
  });
  assert.notStrictEqual(S.weekday(s.milestones.lineStart), 0);
});

test('ETA muộn nhất: tính ngược từ ngày ra chuyền mong muốn', () => {
  const input = { base: 'port', seconds: 7500, isNew: true, inan: true, theu: false, skipSunday: true };
  const target = '2026-12-15';
  const eta = S.latestEta(input, target);
  assert.ok(S.computeSchedule(Object.assign({}, input, { eta })).milestones.lineStart <= target);
  const later = S.computeSchedule(Object.assign({}, input, { eta: S.addDays(eta, 1) }));
  assert.ok(later.milestones.lineStart > target);
});

test('Chuẩn cũ 工場着 D−14 thiếu ngày, chuẩn mới D−21 đáp ứng gần hết các trường hợp', () => {
  const needs = [];
  for (const band of [0, 1, 2]) for (const isNew of [true, false]) for (const pt of [false, true]) {
    let mx = 0;
    for (let w = 0; w < 7; w++) {
      const start = S.addDays('2026-11-02', w);
      const inp = { base: 'factory', band, isNew, inan: pt, theu: pt, skipSunday: true };
      mx = Math.max(mx, S.diffDays(S.latestEta(inp, start), start));
    }
    needs.push(mx);
  }
  assert.strictEqual(Math.max(...needs), 22);
  assert.ok(needs.filter((n) => n <= 14).length <= 1);
  assert.ok(needs.filter((n) => n <= 21).length >= 10);
});
