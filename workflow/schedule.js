/*
 * Lịch chuẩn bị 1 mã hàng: từ ngày ETA đến START (ライン投入 = chuẩn bị xong, vào chuyền).
 * Nguồn số liệu: "Rà soát lại thời gian cần để vào chuyền - 20260722.xlsx"
 * (3 sheet: Dưới 8000 / 8000-12000 / Trên 12000 giây), cập nhật theo biên bản
 * họp rà soát lead time (gia công in/thêu sau cắt, rập + sơ đồ > 12.000s tối thiểu 3 ngày).
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

  // Nơi gia công in / thêu (gia công ngoài)
  var SUBCONTRACTORS = [
    { id: 'hanhdan', key: 'gcHD', label: 'Hạnh Đan', note: 'Gia công 2–3 ngày + vận chuyển (chiều gửi, trưa hôm sau tới; chiều về tương tự)' },
    { id: 'vinhphat', key: 'gcVP', label: 'Vĩnh Phát', note: 'Gia công 4–5 ngày + vận chuyển (chiều gửi, trưa hôm sau tới; chiều về tương tự)' }
  ];

  // Công đoạn lớn. Quy tắc lấy từ file "Thời gian tiêu chuẩn..." và họp rà soát lead time.
  var PHASES = [
    { id: 'p0', label: 'Nhận hàng', labelJa: '入荷',
      rule: 'Khai nhận, vận chuyển từ cảng về nhà máy. Tối đa 7 ngày kể cả luồng đỏ.',
      ruleJa: '港着から工場着まで（通関・輸送）最大7日（赤線含む）。' },
    { id: 'p1', label: 'Vải, phụ liệu', labelJa: '生地・副資材',
      rule: 'Khui thùng 1 ngày, kiểm chất lượng + báo khổ 1 ngày, xả vải 1 ngày. Nguyên liệu phải sẵn sàng ít nhất 4 ngày trước khi cắt.',
      ruleJa: '開梱検品1日、品質検査・有効幅確認1日、放反1日。原材料は裁断開始の4日前までに準備完了。' },
    { id: 'p2', label: 'Test vải', labelJa: '生地テスト',
      rule: 'Test co rút và chỉnh rập theo kết quả: 1–2 ngày.',
      ruleJa: '収縮率テストと型紙微調整：1〜2日。' },
    { id: 'p3', label: 'Điều chỉnh rập, đi sơ đồ', labelJa: 'パターン修正・マーキング',
      rule: 'Đi sơ đồ + in rập 1 ngày. Mã từ 8.000 giây cần tối thiểu 3 ngày cho chỉnh rập + sơ đồ.',
      ruleJa: 'マーカー作成・型紙出力1日。難易度の高い品番は型紙調整＋マーカーで最低3日。' },
    { id: 'p4', label: 'Cắt, in, thêu', labelJa: '裁断・プリント・刺繍',
      rule: 'Cắt, đánh số, bó 2 ngày. Có in: cắt xong và gửi gia công trước ライン投入 4 ngày (Hạnh Đan 2–3 ngày, Vĩnh Phát 4–5 ngày, cộng vận chuyển). Thêu làm trong nhà máy; chỉ gửi ngoài khi cao điểm quá tải.',
      ruleJa: '裁断・ナンバリング・束ね2日。プリントありはライン投入4日前に裁断完了・外注発送（Hạnh Đan 2〜3日、Vĩnh Phát 4〜5日＋輸送）。刺繍は社内で実施（繁忙期の能力超過時のみ外注）。' },
    { id: 'p5', label: 'Chuyền', labelJa: '縫製ライン',
      rule: 'Đọc tài liệu kỹ thuật, rập triển khai, họp triển khai xong trước ライン投入. Sau START: triển khai 5 ngày (8.000s), 7–8 ngày (12.000–15.000s); 1 kỹ thuật ≈ 1.000–1.200 giây/ngày.',
      ruleJa: '技術資料の読み込み・展開準備はライン投入前に完了。START後のライン展開：8,000秒で約5日、12,000〜15,000秒で7〜8日（技術者1人 1,000〜1,200秒/日）。' }
  ];

  // Số ngày chuẩn. Mảng 3 phần tử = [< 8000, 8000-12000, > 12000].
  var DEFAULT_STANDARDS = {
    xnk: [7, 7, 7],
    khui: [1, 1, 1],
    qc: [1, 1, 1],
    test: [1, 1, 1],
    xa: [1, 1, 1],
    nlTruocCat: [4, 4, 4],  // nguyên liệu sẵn sàng trước ngày cắt (tính từ ngày khui thùng)
    rap: [1, 2, 2],         // > 12.000s: họp yêu cầu rập + sơ đồ tối thiểu 3 ngày
    sodo: [1, 2, 2],
    caitien: [3, 3, 3],
    cat: [2, 2, 2],
    gcHD: [4, 4, 4],        // Hạnh Đan 2–3 ngày + vận chuyển = cắt xong trước ライン投入 4 ngày
    gcVP: [6, 6, 6],        // Vĩnh Phát 4–5 ngày + vận chuyển
    theuNM: [1, 1, 1],      // thêu trong nhà máy (không qua vận chuyển)
    hop: [2, 2, 2],
    tkMoi: [5, 7, 8],
    tkLap: [3, 5, 7]
  };

  /*
   * Các bước công việc và quan hệ phụ thuộc.
   *  FS = bắt đầu sau khi bước trước xong
   *  SS = bắt đầu cùng ngày bước trước bắt đầu (+ lag ngày làm việc)
   * START = ライン投入 (chuẩn bị xong, vào chuyền) = sau khi 'hop' và 'gc' đều xong.
   * 'trienkhai' diễn ra sau START.
   */
  var STEPS = [
    { id: 'xnk', nameJa: '通関・工場への輸送', key: 'xnk', dept: 'xnk', phase: 'p0', people: '?', calendar: true,
      name: 'Nhận hàng đến ETA nhà máy', detail: 'Khai nhận, chuyển hàng đến APIC sau khi hàng đến cảng',
      note: 'Cho dù luồng đỏ, tối đa 7 ngày (tính cả ngày nghỉ)', deps: [] },
    { id: 'khui', nameJa: '荷降ろし・開梱検品', key: 'khui', dept: 'bocvac', phase: 'p1', people: '?',
      name: 'Xuống hàng, khui thùng kiểm đếm', detail: 'Xuống hàng, khui thùng, kiểm tra nội dung',
      deps: [{ id: 'xnk', type: 'FS' }] },
    { id: 'qc', nameJa: '品質検査・有効幅確認', key: 'qc', dept: 'qckho', phase: 'p1', people: '1',
      name: 'Kiểm chất lượng, báo khổ', detail: 'Kiểm xác suất chất lượng, xác nhận và báo khổ hữu dụng',
      deps: [{ id: 'khui', type: 'FS' }] },
    { id: 'test', nameJa: '生地テスト（収縮率）', key: 'test', dept: 'kythuat', phase: 'p2', people: '1',
      name: 'Test vải', detail: 'Test độ co rút', deps: [{ id: 'khui', type: 'FS' }] },
    { id: 'xa', nameJa: '放反（リラックス）', key: 'xa', dept: 'cat', phase: 'p1', people: '1',
      name: 'Xả vải', detail: 'Xả vải, để vải nghỉ trước khi cắt', deps: [{ id: 'qc', type: 'FS' }] },
    { id: 'rap', nameJa: '型紙調整（収縮率・合わせ・ノッチ）', key: 'rap', dept: 'kythuat', phase: 'p3', people: '1',
      name: 'Điều chỉnh rập theo kết quả test', detail: 'Điều chỉnh rập theo độ co rút, khớp rập, thêm dấu bấm',
      deps: [{ id: 'test', type: 'FS' }] },
    { id: 'sodo', nameJa: 'マーカー作成・型紙出力', key: 'sodo', dept: 'kythuat', phase: 'p3', people: '2',
      name: 'Đi sơ đồ + in rập cứng', detail: 'Đi sơ đồ, in rập cứng, làm rập tay (2 người)',
      note: 'Bắt đầu từ ngày thứ 2 của điều chỉnh rập', deps: [{ id: 'rap', type: 'SS', lag: 1 }] },
    { id: 'cat', nameJa: '裁断・ナンバリング・束ね', key: 'cat', dept: 'cat', phase: 'p4', people: '1',
      name: 'Cắt, đánh số, bó', detail: 'Cắt (gồm laser), đánh số, thay thân, bó bán thành phẩm',
      note: 'Nguyên liệu phải sẵn sàng ít nhất 4 ngày trước khi cắt',
      deps: [{ id: 'sodo', type: 'FS' }, { id: 'xa', type: 'FS' }, { id: 'khui', type: 'SS', lagKey: 'nlTruocCat' }] },
    { id: 'caitien', nameJa: '改良パターン作成', key: 'caitien', dept: 'kythuat', phase: 'p4', people: '1',
      name: 'Rập cải tiến', detail: 'Làm rập cải tiến song song với cắt', deps: [{ id: 'cat', type: 'SS' }] },
    { id: 'gc', nameJa: 'プリント（外注）', key: null, dept: 'giacong', phase: 'p4', people: '?', when: 'gc',
      name: 'In (gia công ngoài)', detail: 'Gửi bán thành phẩm đi in và nhận về (gồm vận chuyển)',
      deps: [{ id: 'cat', type: 'FS' }] },
    { id: 'theu', nameJa: '刺繍（社内）', key: 'theuNM', dept: 'cat', phase: 'p4', people: '?', when: 'theuIn',
      name: 'Thêu (trong nhà máy)', detail: 'Thêu bán thành phẩm tại nhà máy, không qua vận chuyển',
      deps: [{ id: 'cat', type: 'FS' }] },
    { id: 'hop', nameJa: '技術資料読込・展開用型紙・展開会議', key: 'hop', dept: 'chuyen', phase: 'p5', people: '?',
      name: 'Chuẩn bị triển khai', detail: 'Đọc tài liệu kỹ thuật, làm rập triển khai các size, họp triển khai',
      note: 'Phải xong trước ライン投入', deps: [{ id: 'cat', type: 'FS' }] },
    { id: 'trienkhai', nameJa: 'ライン展開指導（先頭5枚まで）', key: null, dept: 'chuyen', phase: 'p5', people: '1', afterStart: true,
      name: 'Kỹ thuật triển khai', detail: 'Hướng dẫn trên chuyền tới khi công nhân ra được 5 pcs đầu',
      note: 'Diễn ra sau START', deps: [{ id: 'hop', type: 'FS' }, { id: 'gc', type: 'FS' }, { id: 'theu', type: 'FS' }] }
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
   *   inan: có in, theu: có thêu, theuOut: thêu gửi gia công ngoài (chỉ khi cao điểm quá tải),
   *   sub: 'hanhdan' | 'vinhphat' (nơi gia công ngoài),
   *   skipSunday: true, holidays: ['YYYY-MM-DD', ...],
   *   standards: ghi đè DEFAULT_STANDARDS
   * }
   */
  function computeSchedule(input) {
    var band = input.band != null ? input.band : bandIndex(input.seconds);
    var std = Object.assign({}, DEFAULT_STANDARDS, input.standards || {});
    var cal = makeCalendar({ skipSunday: input.skipSunday, holidays: input.holidays });
    var base = input.base === 'factory' ? 'factory' : 'port';
    var theuOut = !!(input.theu && input.theuOut);
    var flags = { gc: !!(input.inan || theuOut), theuIn: !!(input.theu && !input.theuOut) };
    var sub = SUBCONTRACTORS.filter(function (x) { return x.id === input.sub; })[0] || SUBCONTRACTORS[0];
    var done = {};
    var rows = [];

    STEPS.forEach(function (st) {
      var key = st.key || (st.id === 'gc' ? sub.key : (input.isNew === false ? 'tkLap' : 'tkMoi'));
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
          var lag = dp.lagKey ? Math.max(0, Math.ceil(Number((std[dp.lagKey] || [])[band]) || 0)) : (dp.lag || 0);
          if (dp.type === 'SS') cands.push(cal.addWork(p.start, lag));
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
      var name = st.name, nameJa = st.nameJa, detail = st.detail;
      if (st.id === 'gc') {
        var what = input.inan && theuOut ? 'In + thêu' : (input.inan ? 'In' : 'Thêu');
        var whatJa = input.inan && theuOut ? 'プリント・刺繍' : (input.inan ? 'プリント' : '刺繍');
        name = what + ' – ' + sub.label;
        nameJa = whatJa + '（外注：' + sub.label + '）';
        detail = sub.note;
      }
      var r = {
        id: st.id, afterStart: !!st.afterStart, dept: st.dept, deptLabel: DEPTS[st.dept], phase: st.phase,
        name: name, nameJa: nameJa, deptLabelJa: DEPTS_JA[st.dept], detail: detail, note: st.note || '', people: st.people,
        days: dur, start: start, end: end || start, finish: finish, calendar: !!st.calendar,
        deps: st.deps.map(function (d) { return d.id; }).filter(function (id) { return done[id]; })
      };
      done[st.id] = r;
      if (dur > 0) rows.push(r);
    });

    var cut = done.cat;
    var btp = maxDate([cut, done.gc, done.theu].filter(Boolean).map(function (r) { return r.finish; }));
    var fabricAtFactory = done.khui ? done.khui.start : cal.nextWork(input.eta);
    var lineInput = done.trienkhai.start;      // START = ライン投入
    var firstPieces = done.trienkhai.finish;   // ra 5 pcs đầu chuyền
    return {
      band: band,
      bandLabel: BANDS[band].label,
      sub: sub,
      rows: rows,
      milestones: {
        eta: input.eta,
        fabricAtFactory: fabricAtFactory,
        cutStart: cut.start,
        cutDone: cut.finish,
        btpReady: btp,
        lineInput: lineInput,
        firstPieces: firstPieces
      },
      totalDays: diffDays(input.eta, lineInput) + 1, // đánh số như Excel: ngày ETA = N1
      workDays: cal.countWork(input.eta, lineInput),
      calendar: cal
    };
  }

  // Ngày ETA muộn nhất để mốc (mặc định: START = ライン投入) rơi đúng hoặc trước ngày target.
  // milestone: 'lineInput' | 'firstPieces'
  function latestEta(input, target, milestone) {
    var key = milestone || 'lineInput';
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
    BANDS: BANDS, SUBCONTRACTORS: SUBCONTRACTORS, DEPTS: DEPTS, DEPTS_JA: DEPTS_JA, PHASES: PHASES, STEPS: STEPS, DEFAULT_STANDARDS: DEFAULT_STANDARDS,
    bandIndex: bandIndex, computeSchedule: computeSchedule, latestEta: latestEta,
    makeCalendar: makeCalendar, addDays: addDays, diffDays: diffDays, weekday: weekday
  };
});
