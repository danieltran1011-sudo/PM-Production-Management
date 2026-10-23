// Chạy: npm test
// Đối chiếu lịch với file "Rà soát lại thời gian cần để vào chuyền - 20260722.xlsx"
// (các cột N1..N27, tính liên tục, không bỏ Chủ nhật) và các quy tắc trong biên bản họp rà soát lead time.
const test = require('node:test');
const assert = require('node:assert');
const S = require('./schedule.js');

const ETA = '2026-11-01'; // = N1 (hàng về cảng)
const N = (date) => S.diffDays(ETA, date) + 1;

function plan(band, extra) {
  return S.computeSchedule(Object.assign({
    eta: ETA, base: 'port', band, isNew: true, inan: false, theu: false, sub: 'hanhdan', skipSunday: false
  }, extra));
}
function spans(s) {
  const out = {};
  s.rows.forEach((r) => { out[r.id] = [N(r.start), N(r.end)]; });
  return out;
}

test('Trước khi cắt: khớp sheet "Dưới 8000" (N1..N13)', () => {
  const sp = spans(plan(0));
  assert.deepStrictEqual(sp.xnk, [1, 7]);
  assert.deepStrictEqual(sp.khui, [8, 8]);
  assert.deepStrictEqual(sp.qc, [9, 9]);
  assert.deepStrictEqual(sp.test, [9, 9]);
  assert.deepStrictEqual(sp.xa, [10, 10]);
  assert.deepStrictEqual(sp.rap, [10, 10]);
  assert.deepStrictEqual(sp.sodo, [11, 11]);
  assert.deepStrictEqual(sp.cat, [12, 13]);
});

test('8.000–12.000s và > 12.000s: rập + sơ đồ tối thiểu 3 ngày (họp), cắt từ N13', () => {
  for (const band of [1, 2]) {
    const sp = spans(plan(band));
    assert.deepStrictEqual(sp.rap, [10, 11]);
    assert.deepStrictEqual(sp.sodo, [11, 12]);
    assert.strictEqual(sp.sodo[1] - sp.rap[0] + 1, 3);
    assert.deepStrictEqual(sp.cat, [13, 14]);
  }
});

test('Nguyên liệu sẵn sàng ít nhất 4 ngày trước khi cắt', () => {
  for (const band of [0, 1, 2]) {
    const s = plan(band, { skipSunday: true });
    const khui = s.rows.find((r) => r.id === 'khui');
    assert.ok(s.calendar.countWork(khui.start, s.milestones.cutStart) >= 4);
  }
});

test('Không in/thêu: cắt xong trước START 2 ngày', () => {
  const s = plan(0);
  assert.strictEqual(N(s.milestones.cutDone), 14);
  assert.strictEqual(N(s.milestones.lineInput), 16);
  assert.ok(!s.rows.some((r) => r.id === 'gc'));
});

test('Có in: Hạnh Đan cắt xong trước START 4 ngày, Vĩnh Phát 6 ngày', () => {
  const hd = plan(0, { inan: true, sub: 'hanhdan' });
  assert.strictEqual(S.diffDays(hd.milestones.cutDone, hd.milestones.lineInput), 4);
  const bp = plan(0, { inan: true, sub: 'vinhphat' });
  assert.strictEqual(S.diffDays(bp.milestones.cutDone, bp.milestones.lineInput), 6);
  assert.match(bp.rows.find((r) => r.id === 'gc').name, /Vĩnh Phát/);
});

test('Triển khai chuyền sau START: mã mới 5 / 7 / 8 ngày, lặp lại 3 / 5 / 7', () => {
  [[0, 5, 3], [1, 7, 5], [2, 8, 7]].forEach(([band, moi, lap]) => {
    const a = plan(band);
    assert.strictEqual(S.diffDays(a.milestones.lineInput, a.milestones.firstPieces), moi);
    const b = plan(band, { isNew: false });
    assert.strictEqual(S.diffDays(b.milestones.lineInput, b.milestones.firstPieces), lap);
  });
});

test('Ranh giới mức SAH', () => {
  assert.strictEqual(S.bandIndex(7999), 0);
  assert.strictEqual(S.bandIndex(8000), 1);
  assert.strictEqual(S.bandIndex(12000), 1);
  assert.strictEqual(S.bandIndex(12001), 2);
});

test('Vải đã về nhà máy: bỏ bước XNK, khui thùng ngay ngày ETA', () => {
  const s = plan(0, { base: 'factory' });
  assert.strictEqual(s.rows[0].id, 'khui');
  assert.strictEqual(s.rows[0].start, ETA);
  assert.strictEqual(N(s.milestones.lineInput), 9);
});

test('Bỏ Chủ nhật: bước tại nhà máy không rơi vào Chủ nhật hay ngày lễ', () => {
  const s = plan(0, { skipSunday: true, inan: true, holidays: ['2026-11-10'] });
  s.rows.filter((r) => !r.calendar).forEach((r) => {
    assert.notStrictEqual(S.weekday(r.start), 0, r.id + ' bắt đầu Chủ nhật');
    assert.notStrictEqual(r.start, '2026-11-10');
  });
  assert.notStrictEqual(S.weekday(s.milestones.lineInput), 0);
});

test('ETA muộn nhất: tính ngược từ ngày START', () => {
  const input = { base: 'port', seconds: 7500, isNew: true, inan: true, sub: 'vinhphat', skipSunday: true };
  const target = '2026-12-15';
  const eta = S.latestEta(input, target);
  assert.ok(S.computeSchedule(Object.assign({}, input, { eta })).milestones.lineInput <= target);
  const later = S.computeSchedule(Object.assign({}, input, { eta: S.addDays(eta, 1) }));
  assert.ok(later.milestones.lineInput > target);
});

test('工場着 → START: 9–16 ngày; mốc 14 ngày chỉ thiếu với gia công Vĩnh Phát (khớp kết luận họp)', () => {
  const need = (band, gc, sub) => {
    let mx = 0;
    for (let w = 0; w < 7; w++) {
      const start = S.addDays('2026-11-02', w);
      const inp = { base: 'factory', band, isNew: true, inan: gc, sub, skipSunday: true };
      mx = Math.max(mx, S.diffDays(S.latestEta(inp, start), start));
    }
    return mx;
  };
  for (const band of [0, 1, 2]) {
    assert.ok(need(band, false, 'hanhdan') <= 14);
    assert.ok(need(band, true, 'hanhdan') <= 14);
    assert.ok(need(band, true, 'vinhphat') <= 21);
  }
  assert.ok(need(2, true, 'vinhphat') > 14);
  assert.strictEqual(need(2, true, 'vinhphat'), 16);
});

test('Thêu làm trong nhà máy: 1 ngày sau cắt, song song với chuẩn bị triển khai, không làm chậm START', () => {
  const base = plan(0);
  const t = plan(0, { theu: true });
  const r = t.rows.find((x) => x.id === 'theu');
  assert.ok(r && r.dept === 'cat' && r.days === 1);
  assert.ok(!t.rows.some((x) => x.id === 'gc'));
  assert.strictEqual(t.milestones.lineInput, base.milestones.lineInput);
});

test('Thêu gửi ngoài khi cao điểm: tính như gia công ngoài', () => {
  const t = plan(0, { theu: true, theuOut: true, sub: 'vinhphat' });
  assert.ok(!t.rows.some((x) => x.id === 'theu'));
  const gc = t.rows.find((x) => x.id === 'gc');
  assert.match(gc.name, /^Thêu – Vĩnh Phát/);
  assert.strictEqual(S.diffDays(t.milestones.cutDone, t.milestones.lineInput), 6);
});
