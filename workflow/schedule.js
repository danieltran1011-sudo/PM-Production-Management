/*
 * Lịch chuẩn bị 1 mã hàng: từ ngày ETA đến ngày ra chuyền.
 * Nguồn số liệu: "Rà soát lại thời gian cần để vào chuyền - 20260722.xlsx"
 * (3 sheet: Dưới 8000 / 8000-12000 / Trên 12000 giây).
 *
 * Dùng được cả trong trình duyệt (window.PMSchedule) và Node (require).
 * Ngày luôn ở dạng chuỗi 'YYYY-MM-DD'.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PMSchedule = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Mức thời gian chế tạo sản phẩm (SAH, giây/sản phẩm)
  var BANDS = [
    { id: 'lt8000', label: 'Dưới 8.000 giây', short: '< 8.000s' },
    { id: '8000-12000', label: 'Từ 8.000 đến 12.000 giây', short: '8.000–12.000s' },
    { id: 'gt12000', label: 'Trên 12.000 giây', short: '> 12.000s' }
  ];

  function bandIndex(seconds) {
    var s = Number(seconds) || 0;
    if (s < 8000) return 0;
    if (s <= 12000) return 1;
    return 2;
  }

  var DEPTS_JA = {
    xnk: '輸出入・輸送',
    bocvac: '荷役',
    qckho: '倉庫品質管理',
    cat: '裁断課',
    kythuat: '技術課',
    giacong: '外注加工',
    chuyen: '縫製ライン'
  };

  var DEPTS = {
    xnk: 'Xuất nhập khẩu & Vận chuyển',
    bocvac: 'Bốc vác',
    qckho: 'Chất lượng kho',
    cat: 'Phòng cắt',
    kythuat: 'Kỹ thuật',
    giacong: 'Gia công',
    chuyen: 'Chuyền'
  };

  // Công đoạn lớn (theo file "Thời gian tiêu chuẩn chuẩn bị 1 mã hàng từ ngày ETA")
  var PHASES = [
    { id: 'p0', label: 'Nhận hàng', labelJa: '入荷', rule: 'Khai nhận, vận chuyển về nhà máy. Tối đa 7 ngày kể cả luồng đỏ.' },
    { id: 'p1', label: 'Vải, phụ liệu', labelJa: '生地・副資材', rule: 'Bảng chuẩn: xong trước ngày cắt 5 ngày.' },
    { id: 'p2', label: 'Test vải', labelJa: '生地テスト', rule: 'Bảng chuẩn: xong trước ngày cắt 4 ngày.' },
    { id: 'p3', label: 'Điều chỉnh rập, đi sơ đồ', labelJa: 'パターン修正・マーキング', rule: 'Bảng chuẩn: xong trước ngày cắt 3 ngày.' },
    { id: 'p4', label: 'Cắt, in, thêu', labelJa: '裁断・プリント・刺繍', rule: 'Bảng chuẩn: trước ngày cấp chuyền 2 ngày, 4–6 ngày nếu có in/thêu.' },
    { id: 'p5', label: 'Chuyền', labelJa: '縫製ライン展開', rule: 'Bảng chuẩn: triển khai trước ngày ra chuyền 5 / 7 / 8 ngày theo SAH.' }
  ];

  // Số ngày chuẩn. Mảng 3 phần tử = [< 8000, 8000-12000, > 12000].
  var DEFAULT_STANDARDS = {
    xnk: [7, 7, 7],
    khui: [1, 1, 1],
    qc: [1, 1, 1],
    test: [1, 1, 1],
    xa: [1, 1, 1],
    rap: [1, 2, 1],
    sodo: [1, 2, 1],
    caitien: [3, 3, 3],
    cat: [2, 2, 2],
    theu: [1, 1, 1],
    inan: [4, 4, 4],
    hop: [2, 2, 2],
    tkMoi: [5, 7, 8],
    tkLap: [3, 5, 7]
  };

  /*
   * Các bước công việc và quan hệ phụ thuộc.
   *  FS = bắt đầu sau khi bước trước xong
   *  SS = bắt đầu cùng ngày bước trước bắt đầu (+ lag ngày làm việc)
   */
  var STEPS = [
    { id: 'xnk', nameJa: '通関・工場への輸送', key: 'xnk', dept: 'xnk', phase: 'p0', people: '?', calendar: true,
      name: 'Nhận hàng đến ETA nhà máy', detail: 'Khai nhận, chuyển hàng đến APIC sau khi hàng đến cảng',
      note: 'Cho dù luồng đỏ, tối đa 7 ngày (tính cả ngày nghỉ)', deps: [] },
    { id: 'khui', nameJa: '荷降ろし・開梱・数量確認', key: 'khui', dept: 'bocvac', phase: 'p1', people: '?',
      name: 'Xuống hàng, khui thùng kiểm đếm', detail: 'Xuống hàng, khui thùng kiểm đếm',
      deps: [{ id: 'xnk', type: 'FS' }] },
    { id: 'qc', nameJa: '抜き取り品質検査・幅確認', key: 'qc', dept: 'qckho', phase: 'p1', people: '1',
      name: 'Kiểm xác suất chất lượng, báo khổ', detail: 'Kiểm xác suất chất lượng, báo khổ (KC kho)',
      deps: [{ id: 'khui', type: 'FS' }] },
    { id: 'test', nameJa: '生地テスト（収縮率など）', key: 'test', dept: 'kythuat', phase: 'p2', people: '1',
      name: 'Test vải', detail: 'Test độ co rút, chất lượng vải', deps: [{ id: 'khui', type: 'FS' }] },
    { id: 'xa', nameJa: '放反', key: 'xa', dept: 'cat', phase: 'p1', people: '1',
      name: 'Xả vải', detail: 'Xả vải', deps: [{ id: 'qc', type: 'FS' }] },
    { id: 'rap', nameJa: 'パターン修正（収縮率・合わせ・ノッチ）', key: 'rap', dept: 'kythuat', phase: 'p3', people: '1',
      name: 'Điều chỉnh rập theo kết quả test vải', detail: 'Điều chỉnh rập theo độ co rút, khớp rập, thêm dấu bấm',
      deps: [{ id: 'test', type: 'FS' }] },
    { id: 'sodo', nameJa: 'マーキング・硬型紙出力', key: 'sodo', dept: 'kythuat', phase: 'p3', people: '2',
      name: 'Đi sơ đồ + in rập cứng', detail: 'Làm rập tay & in rập, đi sơ đồ & in rập cứng (2 người)',
      note: 'Bắt đầu từ ngày thứ 2 của điều chỉnh rập', deps: [{ id: 'rap', type: 'SS', lag: 1 }] },
    { id: 'cat', nameJa: '裁断・番号付け・パーツ交換', key: 'cat', dept: 'cat', phase: 'p4', people: '1',
      name: 'Cắt (bao gồm laser)', detail: 'Cắt, đánh số, thay thân',
      deps: [{ id: 'sodo', type: 'FS' }, { id: 'xa', type: 'FS' }] },
    { id: 'caitien', nameJa: '改良パターン作成', key: 'caitien', dept: 'kythuat', phase: 'p4', people: '1',
      name: 'Rập cải tiến', detail: 'Làm rập cải tiến song song với cắt', deps: [{ id: 'cat', type: 'SS' }] },
    { id: 'theu', nameJa: '刺繍', key: 'theu', dept: 'giacong', phase: 'p4', people: '?', when: 'theu',
      name: 'Thêu', detail: 'Thêu bán thành phẩm', deps: [{ id: 'cat', type: 'SS' }] },
    { id: 'inan', nameJa: 'プリント', key: 'inan', dept: 'giacong', phase: 'p4', people: '?', when: 'inan',
      name: 'In', detail: 'In bán thành phẩm', deps: [{ id: 'cat', type: 'SS' }] },
    { id: 'hop', nameJa: '展開用パターン作成・展開会議', key: 'hop', dept: 'chuyen', phase: 'p5', people: '?',
      name: 'Làm rập + họp triển khai', detail: 'Làm rập triển khai các size, họp triển khai',
      deps: [{ id: 'cat', type: 'FS' }, { id: 'theu', type: 'FS' }, { id: 'inan', type: 'FS' }] },
    { id: 'trienkhai', nameJa: 'ライン展開指導（先頭5枚まで）', key: null, dept: 'chuyen', phase: 'p5', people: '1',
      name: 'Kỹ thuật triển khai', detail: 'Hướng dẫn tới khi công nhân ra được 5 pcs đầu chuyền',
      deps: [{ id: 'hop', type: 'FS' }] }
  ];

  // ---------- Ngày tháng ----------
  function parse(s) {
    var p = String(s).split('-');
    return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  }
  function fmt(d) {
    return d.toISOString().slice(0, 10);
  }
  function addDays(s, n) {
    var d = parse(s);
    d.setUTCDate(d.getUTCDate() + n);
    return fmt(d);
  }
  function weekday(s) { return parse(s).getUTCDay(); } // 0 = Chủ nhật
  function diffDays(a, b) { return Math.round((parse(b) - parse(a)) / 86400000); }

  function makeCalendar(opts) {
    opts = opts || {};
    var holidays = {};
    (opts.holidays || []).forEach(function (h) { holidays[h] = true; });
    var skipSunday = opts.skipSunday !== false;
    function isWork(s) {
      if (holidays[s]) return false;
      if (skipSunday && weekday(s) === 0) return false;
      return true;
    }
    function nextWork(s) {
      var guard = 0;
      while (!isWork(s) && guard++ < 400) s = addDays(s, 1);
      return s;
    }
    // Ngày làm việc thứ n sau s (s phải là ngày làm việc)
    function addWork(s, n) {
      var d = s;
      for (var i = 0; i < n; i++) d = nextWork(addDays(d, 1));
      return d;
    }
    function countWork(a, b) { // số ngày làm việc trong [a, b)
      var c = 0, d = a, guard = 0;
      while (d < b && guard++ < 2000) { if (isWork(d)) c++; d = addDays(d, 1); }
      return c;
    }
    return { isWork: isWork, nextWork: nextWork, addWork: addWork, countWork: countWork };
  }

  function maxDate(list) {
    return list.reduce(function (m, d) { return d > m ? d : m; });
  }

  /*
   * input: {
   *   eta: 'YYYY-MM-DD',
   *   base: 'port' (ETA hàng về cảng, có bước XNK) | 'factory' (vải đã về nhà máy),
   *   seconds: SAH giây, (hoặc band: 0|1|2)
   *   isNew: true = mã mới, false = hàng lặp lại,
   *   inan: có in, theu: có thêu,
   *   skipSunday: true, holidays: ['YYYY-MM-DD', ...],
   *   standards: ghi đè DEFAULT_STANDARDS
   * }
   */
  function computeSchedule(input) {
    var band = input.band != null ? input.band : bandIndex(input.seconds);
    var std = Object.assign({}, DEFAULT_STANDARDS, input.standards || {});
    var cal = makeCalendar({ skipSunday: input.skipSunday, holidays: input.holidays });
    var base = input.base === 'factory' ? 'factory' : 'port';
    var flags = { inan: !!input.inan, theu: !!input.theu };
    var done = {};
    var rows = [];

    STEPS.forEach(function (st) {
      var key = st.key || (input.isNew === false ? 'tkLap' : 'tkMoi');
      var dur = Math.max(0, Math.ceil(Number((std[key] || [])[band]) || 0));
      var skipped = (st.when && !flags[st.when]) || (st.id === 'xnk' && base === 'factory');
      if (skipped) { done[st.id] = null; return; }

      var start;
      if (st.id === 'xnk') {
        start = input.eta;
      } else {
        var cands = [];
        st.deps.forEach(function (dp) {
          var p = done[dp.id];
          if (!p) return;
          if (dp.type === 'SS') cands.push(cal.addWork(p.start, dp.lag || 0));
          else cands.push(p.finish);
        });
        start = cands.length ? maxDate(cands) : cal.nextWork(input.eta);
      }

      var end, finish;
      if (st.calendar) {
        end = addDays(start, dur - 1);
        finish = cal.nextWork(addDays(start, dur));
      } else if (dur === 0) {
        end = null;
        finish = start;
      } else {
        end = cal.addWork(start, dur - 1);
        finish = cal.nextWork(addDays(end, 1));
      }
      var r = {
        id: st.id, dept: st.dept, deptLabel: DEPTS[st.dept], phase: st.phase,
        name: st.name, nameJa: st.nameJa, deptLabelJa: DEPTS_JA[st.dept], detail: st.detail, note: st.note || '', people: st.people,
        days: dur, start: start, end: end || start, finish: finish, calendar: !!st.calendar,
        deps: st.deps.map(function (d) { return d.id; }).filter(function (id) { return done[id]; })
      };
      done[st.id] = r;
      if (dur > 0) rows.push(r);
    });

    var cut = done.cat;
    var btp = maxDate([done.cat, done.theu, done.inan].filter(Boolean).map(function (r) { return r.finish; }));
    var fabricAtFactory = done.khui ? done.khui.start : cal.nextWork(input.eta);
    var lineStart = done.trienkhai.finish;
    return {
      band: band,
      bandLabel: BANDS[band].label,
      rows: rows,
      milestones: {
        eta: input.eta,
        fabricAtFactory: fabricAtFactory,
        cutStart: cut.start,
        btpReady: btp,
        deployStart: done.trienkhai.start,
        lineStart: lineStart
      },
      totalDays: diffDays(input.eta, lineStart) + 1, // đánh số như Excel: ngày ETA = N1
      workDays: cal.countWork(input.eta, lineStart),
      calendar: cal
    };
  }

  // Ngày ETA muộn nhất để mốc (mặc định: ra chuyền) rơi đúng hoặc trước ngày target.
  // milestone: 'lineStart' (ra chuyền) | 'deployStart' (bắt đầu triển khai chuyền)
  function latestEta(input, target, milestone) {
    var key = milestone || 'lineStart';
    var guess = addDays(target, -120);
    var best = null;
    for (var i = 0; i < 160; i++) {
      var d = addDays(guess, i);
      var s = computeSchedule(Object.assign({}, input, { eta: d }));
      if (s.milestones[key] <= target) best = d;
      else if (best) break;
    }
    return best;
  }

  return {
    BANDS: BANDS, DEPTS: DEPTS, DEPTS_JA: DEPTS_JA, PHASES: PHASES, STEPS: STEPS, DEFAULT_STANDARDS: DEFAULT_STANDARDS,
    bandIndex: bandIndex, computeSchedule: computeSchedule, latestEta: latestEta,
    makeCalendar: makeCalendar, addDays: addDays, diffDays: diffDays, weekday: weekday
  };
});
