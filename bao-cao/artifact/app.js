/* Giao diện báo cáo điều hành APIC – đọc PLANTABLE trong trình duyệt, lưu báo cáo + lịch sử vào db */
(function () {
  'use strict';
  const PE = window.PlanEngine;
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const isNum = Number.isFinite;
  const st = { lang: 'vi', tab: 'overview', V: null, hist: [], db: null, dl: null, buf: null, fileName: '', charts: [], cmp: null, riskFilter: 'all', riskQ: '' };
  try { st.lang = localStorage.getItem('apic-lang') || 'vi'; } catch (e) { /* bỏ qua */ }
  const T = () => window.I18N[st.lang];
  const loc = () => (st.lang === 'vi' ? 'vi-VN' : 'ja-JP');
  const nf = (x, d = 0) => (isNum(x) ? new Intl.NumberFormat(loc(), { minimumFractionDigits: d, maximumFractionDigits: d }).format(x) : '–');
  const pct = (x, d = 0) => (isNum(x) ? nf(x * 100, d) + '%' : '–');
  const spct = (x, d = 0) => (isNum(x) ? (x > 0 ? '+' : '') + nf(x * 100, d) + '%' : '–');
  const usd = x => (isNum(x) ? '$' + nf(x) : '–');
  const musd = x => (isNum(x) ? (st.lang === 'vi' ? nf(x / 1e6, 1) + ' triệu USD' : nf(x / 1e4) + '万USD') : '–');
  const kp = x => (isNum(x) ? nf(x / 1000) + 'k' : '–');
  const dt = d => { if (d == null) return ''; const [y, m, dd] = PE.ymd(d); return st.lang === 'vi' ? `${String(dd).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}` : `${y}/${String(m).padStart(2, '0')}/${String(dd).padStart(2, '0')}`; };
  const ml = mi => { const y = Math.floor(mi / 12), m = (mi % 12) + 1; return st.lang === 'vi' ? `T${m}/${y}` : `${y}/${String(m).padStart(2, '0')}`; };
  const wl = w => (st.lang === 'vi' ? `${ml(w[0])}–${ml(w[1])}` : `${ml(w[0])}–${ml(w[1])}`);
  const mkey = mi => `${Math.floor(mi / 12)}-${String((mi % 12) + 1).padStart(2, '0')}`;
  const chip = (kind, text) => `<span class="chip chip-${kind}">${esc(text)}</span>`;
  const tbl = (head, rows, cls = '') => `<div class="tw"><table class="${cls}"><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => (typeof c === 'object' && c && c.html != null ? `<td class="${c.cls || ''}" ${c.attr || ''}>${c.html}</td>` : `<td class="${typeof c === 'number' ? 'n' : ''}">${esc(c)}</td>`)).join('')}</tr>`).join('')}</tbody></table></div>`;
  const H = (html, cls = '') => ({ html, cls });
  const N = (x, d) => H(nf(x, d), 'n');
  const P = (x, d) => H(pct(x, d), 'n');
  const SP = x => H(spct(x), 'n ' + (x < -0.1 ? 'neg' : x > 0.1 ? 'pos' : ''));
  const sec = (title, body, note) => `<section class="card"><h2>${esc(title)}</h2>${note ? `<p class="note">${esc(note)}</p>` : ''}${body}</section>`;

  // ------------------------------------------------------------ KPI + văn bản
  function kpis(rows) { // [label, value, compare, note, kind, statusText]
    return `<div class="kpis">${rows.map(r => `<div class="kpi kpi-${r[4]}"><div class="kl">${esc(r[0])}</div><div class="kv">${esc(r[1])}</div><div class="kc">${esc(r[2] || '')} <span>${esc(r[3] || '')}</span></div>${chip(r[4], r[5])}</div>`).join('')}</div>`;
  }
  function bullets(items) { return `<ol class="find">${items.map(([h, b]) => `<li><b>${esc(h)}</b> ${esc(b)}</li>`).join('')}</ol>`; }

  function texts() {
    const V = st.V, K = V.K, C = V.capK, R = V.riskK, J = V.juchu, t = st.lang === 'vi';
    const ys = K.years; const yl = ys[ys.length - 1] || {}; const yp = ys[ys.length - 2] || {};
    const cur = K.cur, g = b => (K.brandGrowth.find(x => x.brand === b) || {}).g;
    const w1 = wl(J.w1), w2 = wl(J.w2), mNow = wl([V.m0, V.m0 + 2]), mFut = wl([V.m0 + 4, V.m0 + 7]);
    const big = V.season.filter(s => s.n >= 100).slice(-4);
    const F = t ? [
      ['1. Quy mô và mùa hiện tại.', `Năm ${yl.year} đạt ${kp(yl.qty)} pcs (${spct(yl.d)} so với ${yp.year}). Mùa ${cur} hiện có ${kp(K.curQty)} pcs (${spct(K.curYoy)} so với cùng kỳ, ${spct(K.curPrev)} so với mùa trước). Hàng FW mất ~${nf(K.samFW[0])}–${nf(K.samFW[1])} phút/sp, hàng SS ~${nf(K.samSS[0])}–${nf(K.samSS[1])} phút/sp, nên đo năng lực bằng giờ chuẩn. Đã bỏ ${V.counts.dotted} dòng chấm (${kp(V.counts.dottedQty)} pcs) và tách ${V.counts.sms} dòng hàng SMS.`],
      ['2. Phụ thuộc khách hàng.', `TNF chiếm ${pct(K.tnf)} tổng sản lượng và ${pct(K.tnfCur)} mùa ${cur}; top 3 brand chiếm ${pct(K.top3)}. So với cùng kỳ: HELLY HANSEN ${spct(g('HELLY HANSEN'))}, GOLDWIN ${spct(g('GOLDWIN'))}, ALLTERRAIN 81 ${spct(g('ALLTERRAIN 81'))}, MONTBELL ${spct(g('MONTBELL'))}.`],
      ['3. Công suất.', `${mNow} dùng ~${pct(C.uNow)} công suất (10 tháng trước: ${pct(C.uPast)}). ${mFut} chỉ ${pct(C.uFut)}: để đạt 90% cần thêm ~${kp(C.pcsGap)} pcs; nếu không, ~${musd(C.usdGap)} chi phí nhân công không có đơn để phân bổ.`],
      ['4. Lô nhỏ đắt.', `Lô ≤300 pcs chiếm ${pct(K.lotRows)} số dòng nhưng ${pct(K.lotQty)} sản lượng; đơn giá GC ${nf(K.upcS, 1)} USD/pcs, gấp ${nf(K.lotX, 1)} lần lô 2.001–5.000 (${nf(K.upcM, 1)}), chỉ đạt ~${pct(K.effS)} tốc độ tối đa. Chi phí chia lô cộng dồn ~${musd(K.split)}; ${big.map(s => s.season + ' ' + pct(s.splitPct)).join(', ')}.`],
      ['5. Rủi ro giao hàng.', `${R.all.n} đơn (${kp(R.all.q)} pcs) đang/sắp SX; ${R.ge8.n} đơn (${kp(R.ge8.q)} pcs, ${pct(R.ge8.q / R.all.q)}) có ETD kế hoạch trễ ≥8 ngày so với 工場ETD, ${R.crit.n} đơn trễ ≥30 ngày. ${R.matLate.n} đơn NVL về sau START, ${R.smp.n} đơn chưa duyệt mẫu, ${R.nofac.n} đơn thiếu 工場ETD.`],
      ['6. 受注進行表.', `${w1} phủ ${pct(J.k1.tot)} công suất nhưng chỉ ${pct(J.k1.conf)} đã chốt; ${w2} kể cả tạm tính chỉ ${pct(J.k2.tot)}. ${J.nFree}/30 line hết đơn trước ${dt(J.lim)}.`],
      ['7. Dữ liệu.', `${kp(V.dq.blankQ)} pcs không có brand; ${V.counts.seasonVariants} cách ghi mã mùa; ${V.counts.errCells} ô lỗi công thức; ${V.dq.noFac} đơn sắp SX chưa có 工場ETD. Xem tab Dữ liệu.`],
    ] : [
      ['1. 規模と今シーズン。', `${yl.year}年は${kp(yl.qty)} pcs(${yp.year}年比${spct(yl.d)})。${cur}は${kp(K.curQty)} pcs(前年同期比${spct(K.curYoy)}、前シーズン比${spct(K.curPrev)})。FWは約${nf(K.samFW[0])}–${nf(K.samFW[1])}分/枚、SSは約${nf(K.samSS[0])}–${nf(K.samSS[1])}分/枚のため、能力は標準時間で測るべき。網掛け${V.counts.dotted}行(${kp(V.counts.dottedQty)} pcs)を除外し、SMS品${V.counts.sms}行を別集計。`],
      ['2. 顧客依存。', `TNFが総数量の${pct(K.tnf)}、${cur}では${pct(K.tnfCur)}; 上位3ブランドで${pct(K.top3)}。前年同期比: HELLY HANSEN ${spct(g('HELLY HANSEN'))}、GOLDWIN ${spct(g('GOLDWIN'))}、ALLTERRAIN 81 ${spct(g('ALLTERRAIN 81'))}、MONTBELL ${spct(g('MONTBELL'))}。`],
      ['3. 生産能力。', `${mNow}の稼働率は約${pct(C.uNow)}(直近10か月 ${pct(C.uPast)})。${mFut}は${pct(C.uFut)}のみ: 90%達成には約${kp(C.pcsGap)} pcsの追加が必要; できなければ約${musd(C.usdGap)}の人件費が配賦先を失う。`],
      ['4. 小ロットは高コスト。', `≤300ロットは行数の${pct(K.lotRows)}だが数量は${pct(K.lotQty)}; 加工単価${nf(K.upcS, 1)} USD/pcsで2,001–5,000ロット(${nf(K.upcM, 1)})の${nf(K.lotX, 1)}倍、最大速度の約${pct(K.effS)}。ロット分割コスト累計 約${musd(K.split)}; ${big.map(s => s.season + ' ' + pct(s.splitPct)).join('、')}。`],
      ['5. 納期リスク。', `生産中/予定 ${R.all.n}件(${kp(R.all.q)} pcs); うち${R.ge8.n}件(${kp(R.ge8.q)} pcs、${pct(R.ge8.q / R.all.q)})は計画ETDが工場ETDより8日以上遅れ、${R.crit.n}件は30日以上。資材着がSTART後 ${R.matLate.n}件、サンプル未承認 ${R.smp.n}件、工場ETD未入力 ${R.nofac.n}件。`],
      ['6. 受注進行表。', `${w1}は能力の${pct(J.k1.tot)}が埋まっているが確定は${pct(J.k1.conf)}のみ; ${w2}は暫定込みでも${pct(J.k2.tot)}。${J.nFree}/30ラインが${dt(J.lim)}前に受注切れ。`],
      ['7. データ。', `ブランド空欄 ${kp(V.dq.blankQ)} pcs; シーズン表記${V.counts.seasonVariants}種; 数式エラー${V.counts.errCells}セル; 生産予定で工場ETD未入力 ${V.dq.noFac}件。データタブ参照。`],
    ];
    const A = t ? [
      [`Lấp công suất ${mFut}`, `Chốt sớm đơn mùa tới (~${kp(C.pcsGap)} pcs), ưu tiên brand đang tăng; cân nhắc giá linh hoạt cho tháng thấp điểm.`, 'Sales / MD', '4 tuần'],
      [`Xử lý ${R.ge8.n} đơn trễ ETD ≥8 ngày`, 'Theo danh sách tab Rủi ro, ưu tiên “Nghiêm trọng”: tăng ca/thêm line, tách lô hoặc đàm phán 工場ETD.', 'Kế hoạch SX + Sales', '2 tuần'],
      ['Đẩy NVL & duyệt mẫu', `${R.matLate.n} đơn NVL trễ, ${R.matUnk.n} đơn NVL 未定, ${R.smp.n} đơn chưa duyệt mẫu: checklist sẵn sàng 4 tuần trước START.`, 'Mua hàng / QA', 'Hằng tuần'],
      ['Chính sách lô tối thiểu', `MOQ hoặc phụ phí cho lô ≤300 (đắt gấp ${nf(K.lotX, 1)} lần); gom lô cùng style / 集約CODE.`, 'Ban điều hành + Sales', 'Mùa tới'],
      ['Chốt đơn tạm tính', `${pct(J.k1.tentShare)} đơn ${w1} còn tạm tính: chốt để mua NVL và giữ lịch.`, 'Sales', '2 tuần'],
      ['Chuẩn hoá PLANTABLE', 'Danh mục thả xuống cho mùa/brand/factory, bắt buộc 工場ETD, xoá dòng treo, sửa công thức cột S.', 'Kế hoạch SX / IT', '1 tháng'],
    ] : [
      [`${mFut}の能力を埋める`, `次シーズンを早期確定(約${kp(C.pcsGap)} pcs)、伸びているブランドを優先; 閑散月は柔軟な価格を検討。`, '営業 / MD', '4週間'],
      [`ETD 8日以上遅れ ${R.ge8.n}件への対応`, '納期リスクタブのリストで「重大」を優先: 残業・ライン増、ロット分割、工場ETD交渉。', '生産計画 + 営業', '2週間'],
      ['資材・サンプル承認の促進', `資材遅れ${R.matLate.n}件、資材未定${R.matUnk.n}件、サンプル未承認${R.smp.n}件: START 4週間前に準備チェック。`, '購買 / QA', '毎週'],
      ['最小ロット方針', `≤300ロット(単価${nf(K.lotX, 1)}倍)にMOQまたは割増; 同一品番・集約CODEでまとめる。`, '経営陣 + 営業', '次シーズン'],
      ['暫定オーダーの確定', `${w1}の${pct(J.k1.tentShare)}が暫定: 資材手配と日程確保のため確定。`, '営業', '2週間'],
      ['PLANTABLEの標準化', 'シーズン・ブランド・工場はドロップダウン、工場ETD必須、残存行削除、S列の数式修正。', '生産計画 / IT', '1か月'],
    ];
    return { F, A };
  }

  // ------------------------------------------------------------ biểu đồ
  const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
  function chart(id, cfg) {
    const el = document.getElementById(id); if (!el || !window.Chart) return;
    const ink = css('--ink-2'), grid = css('--rule');
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false, plugins: { legend: { labels: { color: ink, boxWidth: 12 } } } }, cfg.options || {});
    cfg.options.scales = cfg.options.scales || {};
    for (const ax of ['x', 'y']) { const s = cfg.options.scales[ax] = cfg.options.scales[ax] || {}; s.ticks = Object.assign({ color: ink }, s.ticks || {}); s.grid = Object.assign({ color: grid }, s.grid || {}); }
    st.charts.push(new window.Chart(el, cfg));
  }
  const C1 = () => css('--s1'), C2 = () => css('--s2'), C3 = () => css('--s3');

  // ------------------------------------------------------------ các tab
  function tabOverview() {
    const V = st.V, K = V.K, C = V.capK, R = V.riskK, J = V.juchu, L = T(); const ys = K.years, yl = ys[ys.length - 1] || {}; const vi = st.lang === 'vi';
    const k = [
      [vi ? `Sản lượng năm ${yl.year}` : `${yl.year}年 数量`, kp(yl.qty) + ' pcs', spct(yl.d), vi ? 'so với năm trước' : '前年比', yl.d >= 0 ? 'good' : 'warning', yl.d >= 0 ? (vi ? 'Tăng' : '増加') : (vi ? 'Giảm' : '減少')],
      [vi ? `Mùa ${K.cur}` : `${K.cur} 数量`, kp(K.curQty) + ' pcs', spct(K.curYoy), L.dYoy, K.curYoy >= 0 ? 'good' : 'warning', K.curYoy >= 0 ? (vi ? 'Tăng' : '増加') : (vi ? 'Giảm' : '減少')],
      [vi ? 'Tỷ trọng TNF' : 'TNF比率', pct(K.tnf), pct(K.tnfCur), K.cur, 'serious', vi ? 'Phụ thuộc cao' : '依存度高'],
      [vi ? `Tải ${wl([V.m0, V.m0 + 2])}` : `稼働率 ${wl([V.m0, V.m0 + 2])}`, pct(C.uNow), pct(C.uPast), vi ? '10 tháng trước' : '直近10か月', C.uNow >= .85 ? 'good' : 'warning', C.uNow >= .85 ? L.full : L.room],
      [vi ? `Cần thêm ${wl([V.m0 + 4, V.m0 + 7])}` : `追加必要 ${wl([V.m0 + 4, V.m0 + 7])}`, kp(C.pcsGap) + ' pcs', musd(C.usdGap), vi ? 'nhân công chưa có đơn' : '受注なし人件費', C.uFut < .5 ? 'critical' : 'warning', vi ? 'Lỗ hổng' : '空き大'],
      [vi ? 'Lô ≤300 đắt gấp' : '≤300ロット単価', nf(K.lotX, 1) + '×', pct(K.lotRows), vi ? 'số dòng đơn' : '行数比', 'serious', vi ? 'Mất năng suất' : '生産性ロス'],
      [vi ? 'SL trễ ETD ≥8 ngày' : 'ETD 8日以上遅れ', kp(R.ge8.q) + ' pcs', pct(R.ge8.q / R.all.q), vi ? 'SL sắp SX' : '予定数量比', 'critical', vi ? 'Rủi ro cao' : '高リスク'],
      [vi ? `受注進行表 ${wl(J.w1)}` : `受注進行表 ${wl(J.w1)}`, pct(J.k1.tot), pct(J.k1.conf), L.conf, J.k1.conf < .5 ? 'warning' : 'good', vi ? 'Chủ yếu tạm tính' : '大半が暫定'],
      [vi ? `受注進行表 ${wl(J.w2)}` : `受注進行表 ${wl(J.w2)}`, pct(J.k2.tot), nf(J.k2.free), vi ? 'line-ngày trống' : '空きライン日', J.k2.tot < .5 ? 'critical' : 'good', J.k2.tot < .5 ? L.short : L.full],
    ];
    const { F, A } = texts();
    return kpis(k) + sec(L.findings, bullets(F)) + sec(L.actions, tbl(['', L.note, L.owner, L.deadline], A.map(a => [H(`<b>${esc(a[0])}</b>`), a[1], a[2], a[3]])));
  }

  function tabCeo() {
    const V = st.V, K = V.K, L = T(), vi = st.lang === 'vi', C = V.capK, J = V.juchu;
    const yrs = V.year.filter(y => y.full);
    const sum30 = vi ? `Sản lượng ${yrs.map(y => y.year + ': ' + kp(y.qty)).join(', ')} pcs. Mùa ${K.cur} ${spct(K.curYoy)} so với cùng kỳ. Nhà máy kín ~${pct(C.uNow)} đến ${ml(V.m0 + 2)}, nhưng ${wl([V.m0 + 4, V.m0 + 7])} chỉ ${pct(C.uFut)}: cần ~${kp(C.pcsGap)} pcs, nếu không ~${musd(C.usdGap)} nhân công không có đơn. TNF ${pct(K.tnf)}; lô nhỏ đắt gấp ${nf(K.lotX, 1)} lần; ${pct(V.riskK.ge8.q / V.riskK.all.q)} SL sắp SX trễ ETD ≥8 ngày.`
      : `数量 ${yrs.map(y => y.year + '年: ' + kp(y.qty)).join('、')} pcs。${K.cur}は前年同期比${spct(K.curYoy)}。${ml(V.m0 + 2)}まで稼働率約${pct(C.uNow)}だが、${wl([V.m0 + 4, V.m0 + 7])}は${pct(C.uFut)}: 約${kp(C.pcsGap)} pcsが必要、なければ約${musd(C.usdGap)}の人件費が配賦先なし。TNF ${pct(K.tnf)}; 小ロット単価${nf(K.lotX, 1)}倍; 予定数量の${pct(V.riskK.ge8.q / V.riskK.all.q)}がETD 8日以上遅れ。`;
    const cur = K.cur, yo = PE.yoySeason(cur), pv = PE.prevSeason(cur);
    const brands = V.brandTotal.filter(b => b.brand !== '(trống)').slice(0, 8).map(b => b.brand);
    const bs = (b, s) => (V.brandSeason[b] || {})[s] || 0;
    const dec = vi ? [
      ['1. Lấp công suất', `Duyệt mục tiêu ~${kp(C.pcsGap)} pcs đơn mùa tới trong 4 tuần; giá linh hoạt cho tháng thấp điểm.`, 'Dự phòng: điều chuyển nhân lực, nghỉ phép bù, gia công ngắn hạn.'],
      ['2. Lô tối thiểu', `MOQ / phụ phí cho lô ≤300 pcs (chiếm ${pct(K.lotRows)} số đơn, ${pct(K.lotQty)} sản lượng, đơn giá ×${nf(K.lotX, 1)}).`, 'Mục tiêu chi phí chia lô ≤10%.'],
      ['3. Đa dạng khách hàng', `Brand ngoài TNF ≥30% sản lượng (hiện ${pct(1 - K.tnf)}).`, 'Theo dõi brand giảm mạnh ở bảng dưới.'],
      ['4. Đơn trễ', `Ngân sách tăng ca/thêm line cho ${V.riskK.ge8.n} đơn trễ ≥8 ngày (~${musd(V.riskK.ge8.v)} giá trị GC).`, 'So sánh với rủi ro phạt / mất đơn.'],
      ['5. Chốt đơn', `${pct(J.k1.tentShare)} đơn ${wl(J.w1)} còn tạm tính.`, 'Sales chốt trong 2 tuần.'],
    ] : [
      ['1. 能力を埋める', `次シーズン約${kp(C.pcsGap)} pcsを4週間で確定する目標を承認; 閑散月は柔軟価格。`, '代替: 配置転換、代休、短期受託。'],
      ['2. 最小ロット', `≤300ロットにMOQ/割増(行数${pct(K.lotRows)}、数量${pct(K.lotQty)}、単価×${nf(K.lotX, 1)})。`, '分割コスト ≤10% を目標。'],
      ['3. 顧客分散', `TNF以外を数量の30%以上に(現在${pct(1 - K.tnf)})。`, '下表で減少ブランドを確認。'],
      ['4. 遅延オーダー', `ETD 8日以上遅れ${V.riskK.ge8.n}件(加工金額 約${musd(V.riskK.ge8.v)})の残業・ライン増予算。`, 'ペナルティ・失注リスクと比較。'],
      ['5. 受注確定', `${wl(J.w1)}の${pct(J.k1.tentShare)}が暫定。`, '営業が2週間以内に確定。'],
    ];
    return `<section class="card lead"><p>${esc(sum30)}</p></section>` +
      `<div class="two">${sec(L.yearTable, `<div class="ch"><canvas id="cYear"></canvas></div>` + tbl([L.year, L.qty, vi ? 'Tăng/giảm' : '増減', L.stdH, L.cost, L.split], yrs.map(y => [String(y.year), N(y.qty), SP(y.d), N(y.h), H(usd(y.c2), 'n'), H(usd(y.split), 'n')])))}` +
      sec(L.byBrand, tbl([L.brand, `${cur}`, `${yo}`, `${pv}`, L.dYoy, L.dPrev], brands.map(b => [b, N(bs(b, cur)), N(bs(b, yo)), N(bs(b, pv)), SP(bs(b, yo) ? bs(b, cur) / bs(b, yo) - 1 : NaN), SP(bs(b, pv) ? bs(b, cur) / bs(b, pv) - 1 : NaN)]))) + `</div>` +
      sec(vi ? 'Quyết định cần CEO phê duyệt' : 'CEO承認事項', tbl(['', L.note, ''], dec.map(d => [H(`<b>${esc(d[0])}</b>`), d[1], d[2]])));
  }
  function afterCeo() { const y = st.V.year.filter(x => x.full); chart('cYear', { type: 'bar', data: { labels: y.map(x => x.year), datasets: [{ label: T().qty, data: y.map(x => x.qty), backgroundColor: C1(), borderRadius: 4 }] }, options: { plugins: { legend: { display: false } } } }); }

  function tabFactory() {
    const V = st.V, L = T(), C = V.capK, vi = st.lang === 'vi', K = V.K;
    const k = [
      [vi ? `Tải ${wl([V.m0, V.m0 + 2])}` : `稼働率 ${wl([V.m0, V.m0 + 2])}`, pct(C.uNow), pct(C.uPast), vi ? '10 tháng trước' : '直近10か月', 'good', L.full],
      [vi ? `Tải ${ml(V.m0 + 3)}` : `稼働率 ${ml(V.m0 + 3)}`, pct(C.uNext), '', vi ? 'kiểm tra Tết / nghỉ lễ' : 'テト・休日を確認', 'warning', vi ? 'Theo dõi' : '注意'],
      [vi ? `Tải ${wl([V.m0 + 4, V.m0 + 7])}` : `稼働率 ${wl([V.m0 + 4, V.m0 + 7])}`, pct(C.uFut), nf(C.wdGap), vi ? 'công-ngày cần thêm' : '追加必要人日', 'critical', L.short],
      [L.effReq, pct(C.effPast), '', vi ? 'giờ chuẩn / giờ công (10 tháng)' : '標準時間 / 工数', 'warning', vi ? 'Theo dõi' : '注意'],
      [vi ? 'Mã bắt đầu / tháng' : '月平均着手品番', nf(C.startsAvg), nf(C.startsAvg / 30, 1), vi ? 'lần / line / tháng' : '回/ライン/月', 'serious', vi ? 'Đổi mã nhiều' : '切替多い'],
      [vi ? 'Đơn song song / line / ngày' : '並行オーダー/ライン/日', nf(C.parallel, 2), '', '', 'warning', vi ? 'Theo dõi' : '注意'],
      [vi ? 'Hiệu suất lô ≤300' : '≤300ロット効率', pct(K.effS), pct(V.lot[4].eff), '2.001–5.000', 'serious', vi ? 'Mất năng suất' : '生産性ロス'],
    ];
    const ms = []; for (let mi = V.m0; mi <= V.m0 + 7; mi++) ms.push(mi);
    const heat = tbl([L.line, ...ms.map(ml), L.orders, L.qty, L.freeFrom, L.smallLines], V.heat.map(h => [H(`<b>${h.line}</b>`), ...h.cells.map(c => H(c == null ? '' : pct(c), 'n heat" style="--h:' + (c == null ? 0 : Math.min(c, 1.1)))), N(h.n), N(h.qty), H(dt(h.freeFrom), h.freeFrom != null && h.freeFrom < PE.mStart(V.m0 + 4) + 14 ? 'bad' : ''), h.small.join(', ')]), 'heatmap');
    const perf = tbl([vi ? 'Cụm line' : 'ライン番号', L.orders, L.qty, L.stdH, L.samMin, L.lotAvg, L.cost, L.upc, L.effLot], V.linePerf.map(p => [String(p.line), N(p.n), N(p.qty), N(p.h), N(p.sam), N(p.lot), H(usd(p.cost), 'n'), N(p.upc, 2), P(p.eff)]));
    return kpis(k) + `<div class="two">${sec(L.util, `<div class="ch"><canvas id="cCap"></canvas></div>`)}${sec(L.capacity, tbl([L.month, L.wd, L.capacity, L.occ, L.util, L.qty, L.starts], V.cap.map(c => [ml(c.mi), N(c.wd), N(c.cap), N(c.occ), H(pct(c.util), 'n ' + (c.util < .5 ? 'neg' : '')), N(c.qty), N(c.starts)])))}</div>` +
      sec(L.heatTitle, heat, L.heatNote) + sec(L.perfTitle, perf);
  }
  function afterFactory() { const c = st.V.cap; chart('cCap', { type: 'line', data: { labels: c.map(x => ml(x.mi)), datasets: [{ label: T().util, data: c.map(x => x.util), borderColor: C1(), backgroundColor: C1() + '22', fill: true, pointRadius: 3, tension: 0 }] }, options: { plugins: { legend: { display: false } }, scales: { y: { min: 0, max: 1.1, ticks: { callback: v => pct(v) } } } } }); }

  function tabProd() {
    const V = st.V, L = T(), vi = st.lang === 'vi', r = V.ready;
    const c = f => r.filter(f); const q = a => a.reduce((s, x) => s + (x.q || 0), 0);
    const k = [
      [vi ? 'Đơn bắt đầu trong 6 tuần' : '6週間以内の着手', nf(r.length), kp(q(r)) + ' pcs', '', 'good', vi ? 'Khối lượng' : '物量'],
      [L.readyY, nf(c(x => x.ok).length), pct(c(x => x.ok).length / Math.max(r.length, 1)), vi ? 'mục tiêu ≥90%' : '目標 ≥90%', 'warning', vi ? 'Cần đẩy' : '要促進'],
      [vi ? 'Vướng NVL' : '資材問題', nf(c(x => x.nvl === 'late' || x.nvl === 'unk').length), kp(q(c(x => x.nvl === 'late' || x.nvl === 'unk'))) + ' pcs', '', 'critical', vi ? 'Đôn đốc mua hàng' : '購買督促'],
      [vi ? 'Chưa duyệt mẫu' : 'サンプル未承認', nf(c(x => !x.sm).length), kp(q(c(x => !x.sm))) + ' pcs', '', 'serious', vi ? 'Đôn đốc QA' : 'QA督促'],
      [vi ? 'Trễ ETD ≥8 ngày' : 'ETD 8日以上遅れ', nf(c(x => x.etd === 'crit' || x.etd === 'high').length), kp(q(c(x => x.etd === 'crit' || x.etd === 'high'))) + ' pcs', '', 'critical', vi ? 'Cân đối line' : 'ライン再配分'],
      [vi ? 'Thiếu 工場ETD' : '工場ETD未入力', nf(c(x => x.etd === 'nofac').length), '', vi ? 'Sales nhập' : '営業が入力', 'warning', vi ? 'Bổ sung' : '補完'],
    ];
    const g = `<div class="tw"><table class="gantt"><thead><tr><th>${L.line}</th>${V.weeks.map(w => `<th>${esc(dt(w).slice(0, st.lang === 'vi' ? 5 : 10).replace(/^\d{4}\//, ''))}</th>`).join('')}</tr></thead><tbody>${V.gantt.map(row => `<tr><th>${row.line}</th>${row.cells.map(cc => cc ? `<td class="g-${cc.st}">${esc(cc.text)}</td>` : `<td class="g-empty">·</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    const rt = tbl([L.line, L.style, L.brand, L.qty, 'CUT', L.start, L.end, L.inputOk, vi ? 'NVL' : '資材', vi ? 'Mẫu' : 'サンプル', L.facEtd, L.gap, 'ETD', vi ? 'Kết luận' : '判定'], r.map(x => [
      `${x.l}${x.p || ''}`, x.st, x.b, N(x.q), dt(x.a - 10), dt(x.a), dt(x.e), x.io != null ? dt(x.io) : (x.itx || ''), H(esc(L.mat[x.nvl] || ''), x.nvl === 'late' ? 'bad' : x.nvl === 'unk' ? 'warn' : ''), H(esc(L.smp[x.sm]), x.sm ? '' : 'warn'),
      dt(x.fe), x.g == null ? '' : N(x.g), H(esc(L.lv[x.etd] || ''), 'lv-' + x.etd), H(x.ok ? chip('good', L.readyY) : esc(x.blk.map(b => L.blk[b]).join(' · ')), x.ok ? '' : 'bad')]));
    const sms = tbl([L.line, L.brand, L.style, L.qty, L.start, L.end, L.comment], V.smsFwd.map(x => [`${x.l}${x.p || ''}`, x.b, x.st, N(x.q), dt(x.a), dt(x.e), x.c]));
    return kpis(k) + sec(L.ganttTitle, g, L.ganttNote) + sec(L.readyTitle, rt) + sec(L.smsTitle, sms);
  }

  function trendRows() {
    const J = st.V.juchu; const win = (s, w) => { const ms = []; for (let mi = w[0]; mi <= w[1]; mi++) ms.push(mkey(mi)); const x = ms.map(k => s.months.find(m => m.m === k)); if (x.some(v => !v)) return [null, null]; const cap = x.reduce((a, v) => a + v.cap, 0); return [x.reduce((a, v) => a + v.tot, 0) / cap, x.reduce((a, v) => a + v.conf, 0) / cap]; };
    const snaps = st.hist.slice(); const curKey = PE.iso(st.V.snap) + '_PLANTABLE';
    if (!snaps.find(s => s.id === curKey)) snaps.push({ id: curKey, date: PE.iso(st.V.snap), src: 'PLANTABLE', months: st.V.juchu.mon.filter(m => m.mi >= st.V.m0).map(m => ({ m: mkey(m.mi), cap: m.cap, conf: m.conf, tot: m.tot })) });
    return snaps.sort((a, b) => (a.date + a.src).localeCompare(b.date + b.src)).map(s => { const a = win(s, J.w1), b = win(s, J.w2); return { date: s.date, src: s.src, t1: a[0], c1: a[1], t2: b[0], c2: b[1] }; });
  }
  function tabJuchu() {
    const V = st.V, J = V.juchu, L = T(), vi = st.lang === 'vi';
    const tr = trendRows(); const prev = tr.filter(x => x.t2 != null && x.date <= PE.iso(V.ref - 21)).slice(-1)[0];
    const k = [
      [vi ? `Độ phủ ${wl(J.w1)}` : `充足率 ${wl(J.w1)}`, pct(J.k1.tot), pct(J.k1.conf), L.conf, 'warning', vi ? 'Kín, chủ yếu tạm tính' : '満杯・大半暫定'],
      [vi ? `Độ phủ ${wl(J.w2)}` : `充足率 ${wl(J.w2)}`, pct(J.k2.tot), prev ? pct(prev.t2) : '', prev ? (vi ? 'bản ' : '') + prev.date : '', 'critical', L.short],
      [vi ? `Line-ngày trống ${wl(J.w2)}` : `空きライン日 ${wl(J.w2)}`, nf(J.k2.free), kp(J.k2.free * 32 * V.capK.prod) + ' pcs', vi ? '≈ theo năng suất' : '≈ 生産性ベース', 'critical', vi ? 'Cần đơn mới' : '新規受注必要'],
      [vi ? `Line hết đơn trước ${dt(J.lim)}` : `${dt(J.lim)}前に受注切れ`, `${J.nFree}/30`, '', '', 'serious', vi ? 'Ưu tiên giao đơn' : '優先割付'],
      [vi ? `Tạm tính trong ${wl(J.w1)}` : `${wl(J.w1)}の暫定比率`, pct(J.k1.tentShare), '', vi ? 'chữ đỏ' : '赤字', 'serious', vi ? 'Cần chốt' : '確定必要'],
    ];
    const mon = tbl([L.month, L.wd, L.capacity, L.conf, L.tent, L.empty, L.confShare, L.totShare], J.mon.map(m => [ml(m.mi), N(m.wd), N(m.cap), N(m.conf), N(m.tot - m.conf), N(Math.max(m.cap - m.tot, 0)), P(m.conf / m.cap), H(pct(m.tot / m.cap), 'n ' + (m.tot / m.cap < .5 ? 'neg' : ''))]));
    // lưới
    let grid = `<div class="tw"><table class="jgrid"><thead><tr><th>${L.line}</th>${J.per.map(p => `<th>${esc(p[2])}</th>`).join('')}</tr></thead><tbody>`;
    for (let ln = 1; ln <= 15; ln++) for (const h of ['A', 'B']) {
      grid += `<tr><th>${ln}${h}</th>`; let j = 0;
      while (j < J.per.length) { const c = J.grid[`${ln}|${h}|${j}`]; if (!c) { grid += '<td class="j-empty"></td>'; j++; continue; }
        let span = 1; while (j + span < J.per.length) { const n = J.grid[`${ln}|${h}|${j + span}`]; if (n && n.lab === c.lab && n.tent === c.tent) span++; else break; }
        const kind = c.lab.includes('_SMS') ? 'sms' : c.lab[0] === 'S' ? 'ss' : c.lab[0] === 'F' ? 'fw' : 'ot';
        grid += `<td colspan="${span}" class="j-${kind}${c.tent ? ' j-tent' : ''}">${esc(c.lab)}${c.small ? '*' : ''}${c.more ? ' +' : ''}</td>`; j += span; }
      grid += '</tr>';
    }
    grid += '</tbody></table></div>';
    const ms = []; for (let mi = V.m0 + 1; mi <= V.m0 + 7; mi++) ms.push(mi);
    const sb = tbl([L.season + ' / ' + L.brand, ...ms.map(ml), vi ? 'Tổng' : '合計', L.tentPart, '%'], J.sb.filter(x => x.tot >= 1).slice(0, 16).map(x => [x.key, ...ms.map(mi => H(x.m[mi] >= .5 ? nf(x.m[mi]) : '', 'n')), N(x.tot), N(x.tent), P(x.tent / x.tot)]));
    const lines = tbl([L.line, L.curRun, L.confEnd, L.endAll, L.freeFrom, L.lastOrder, L.smallLines], J.lines.map(l => [l.line, l.cur, dt(l.confEnd), dt(l.end), H(dt(l.end != null ? l.end + 1 : null), l.end != null && l.end < J.lim ? 'bad' : ''), l.last, l.small]));
    const trT = tbl([L.snapDate, L.source, `${wl(J.w1)} · ${vi ? 'gồm tạm tính' : '暫定込み'}`, `${wl(J.w1)} · ${L.conf}`, `${wl(J.w2)} · ${vi ? 'gồm tạm tính' : '暫定込み'}`, `${wl(J.w2)} · ${L.conf}`], tr.map(x => [x.date, x.src === 'PLANTABLE' ? 'PLANTABLE' : (vi ? '受注進行表 (file)' : '受注進行表(ファイル)'), P(x.t1), P(x.c1), P(x.t2), P(x.c2)]));
    const yrs = J.years; const by = J.brandYear; const tot = {}; for (const b in by) for (const y of yrs) tot[y] = (tot[y] || 0) + (by[b][y] || 0);
    const top = Object.keys(by).sort((a, b) => yrs.reduce((s, y) => s + (by[b][y] || 0), 0) - yrs.reduce((s, y) => s + (by[a][y] || 0), 0)).slice(0, 10);
    const mix = tbl([L.brand, ...yrs.map(String)], top.map(b => [b, ...yrs.map(y => P((by[b][y] || 0) / (tot[y] || 1)))]).concat([[L.other, ...yrs.map(y => P(1 - top.reduce((s, b) => s + (by[b][y] || 0), 0) / (tot[y] || 1)))]]));
    return kpis(k) + `<p class="note">${esc(L.juchuNote)}</p>` + `<div class="two">${sec(vi ? 'Độ phủ theo tháng (line-ngày)' : '月別受注充足(ライン日)', `<div class="ch"><canvas id="cJ"></canvas></div>`)}${sec(L.month, mon)}</div>` +
      sec('受注進行表', grid, L.gridLegend) + sec(L.sbTitle, sb) + sec(L.line, lines) +
      `<div class="two">${sec(L.trend, `<div class="ch"><canvas id="cTr"></canvas></div>`, L.trendNote)}${sec(L.trend, trT)}</div>` + sec(L.brandMix, mix);
  }
  function afterJuchu() {
    const J = st.V.juchu, L = T();
    chart('cJ', { type: 'bar', data: { labels: J.mon.map(m => ml(m.mi)), datasets: [{ label: L.conf, data: J.mon.map(m => m.conf), backgroundColor: C1(), stack: 's' }, { label: L.tent, data: J.mon.map(m => m.tot - m.conf), backgroundColor: C2(), stack: 's' }, { label: L.empty, data: J.mon.map(m => Math.max(m.cap - m.tot, 0)), backgroundColor: css('--rule'), stack: 's' }] }, options: { scales: { x: { stacked: true }, y: { stacked: true } } } });
    const tr = trendRows();
    chart('cTr', { type: 'line', data: { labels: tr.map(x => x.date.slice(5)), datasets: [{ label: wl(J.w1), data: tr.map(x => x.t1), borderColor: C1(), pointRadius: 3, spanGaps: true }, { label: wl(J.w2), data: tr.map(x => x.t2), borderColor: C2(), pointRadius: 3, spanGaps: true }] }, options: { scales: { y: { min: 0, max: 1, ticks: { callback: v => pct(v) } } } } });
  }

  function cmpStats(s) {
    const c = st.V.cube[s]; if (!c) return null;
    return { n: c.n, q: c.q, h: c.h, sam: c.qs ? c.h * 60 / c.qs : NaN, lot: c.n ? c.q / c.n : NaN, sd: c.sdn ? c.sd / c.sdn : NaN, c2: c.c2, upc: c.qc ? c.c2 / c.qc : NaN, split: c.split, small: c.q ? ((c.lot['0'] || [0, 0])[1]) / c.q : NaN, bes: c.q ? c.bes / c.q : NaN, tnf: c.q ? c.tnf / c.q : NaN, c };
  }
  function tabSeasons() {
    const V = st.V, L = T(), vi = st.lang === 'vi';
    const s = V.season;
    const tabS = tbl([L.season, L.rowsN, L.styles, L.qty, L.dPrev, L.dYoy, L.stdH, L.samMin, L.cost, L.upc, L.split, L.splitPct, L.lotMed, L.bespoke], s.map(x => [x.season, N(x.n), N(x.styles), N(x.qty), SP(x.dPrev), SP(x.dYoy), N(x.h), N(x.sam, 1), H(usd(x.c2), 'n'), N(x.upc, 2), H(usd(x.split), 'n'), P(x.splitPct, 1), N(x.lotMed), P(x.bespoke)]));
    const cur = st.cmp || [V.K.cur, PE.prevSeason(V.K.cur), PE.yoySeason(V.K.cur)];
    const sel = (id, v) => `<select id="${id}" class="sel">${V.seasons.map(x => `<option ${x === v ? 'selected' : ''}>${x}</option>`).join('')}</select>`;
    const A = cmpStats(cur[0]) || {}, Pv = cmpStats(cur[1]) || {}, Y = cmpStats(cur[2]) || {};
    const d = (a, b) => (isNum(a) && isNum(b) && b ? a / b - 1 : NaN), dp = (a, b) => (isNum(a) && isNum(b) ? a - b : NaN);
    const met = [[L.rowsN, 'n', 0], [L.qty, 'q', 0], [L.stdH, 'h', 0], [L.samMin, 'sam', 1], [L.lotAvg, 'lot', 0], [vi ? 'Số ngày may TB' : '平均縫製日数', 'sd', 1], [L.cost, 'c2', 0], [L.upc, 'upc', 2], [L.split, 'split', 0], [vi ? '% SL lô ≤300' : '≤300ロット数量比', 'small', 'p'], [L.bespoke, 'bes', 'p'], [vi ? '% TNF' : 'TNF比率', 'tnf', 'p']];
    const mt = tbl([L.kpi, cur[0], cur[1], cur[2], L.dPrev, L.dYoy], met.map(([lab, k2, dd]) => dd === 'p' ? [lab, P(A[k2], 1), P(Pv[k2], 1), P(Y[k2], 1), H(isNum(dp(A[k2], Pv[k2])) ? (dp(A[k2], Pv[k2]) > 0 ? '+' : '') + nf(dp(A[k2], Pv[k2]) * 100, 1) + (vi ? ' điểm' : 'pt') : '–', 'n'), H(isNum(dp(A[k2], Y[k2])) ? (dp(A[k2], Y[k2]) > 0 ? '+' : '') + nf(dp(A[k2], Y[k2]) * 100, 1) + (vi ? ' điểm' : 'pt') : '–', 'n')]
      : [lab, N(A[k2], dd), N(Pv[k2], dd), N(Y[k2], dd), SP(d(A[k2], Pv[k2])), SP(d(A[k2], Y[k2]))]));
    const dim = (key, labels) => tbl([vi ? 'Nhóm' : '区分', cur[0], cur[1], cur[2], '% ' + cur[0], '% ' + cur[2], L.dYoy, L.upc + ' ' + cur[0], L.upc + ' ' + cur[2]], labels.map((lab, i) => {
      const g = (S, j) => ((S.c || { [key]: {} })[key][i] || [0, 0, 0, 0, 0])[j]; const a = g(A, 1), p = g(Pv, 1), y = g(Y, 1);
      return [lab, N(a), N(p), N(y), P(a / (A.q || 1), 1), P(y / (Y.q || 1), 1), SP(y ? a / y - 1 : NaN), N(g(A, 4) ? g(A, 3) / g(A, 4) : NaN, 2), N(g(Y, 4) ? g(Y, 3) / g(Y, 4) : NaN, 2)]; }));
    const bset = [...new Set([...Object.keys((A.c || {}).brand || {}), ...Object.keys((Y.c || {}).brand || {})])].sort((a, b) => (((Y.c || {}).brand || {})[b] || [0, 0])[1] + (((A.c || {}).brand || {})[b] || [0, 0])[1] - ((((Y.c || {}).brand || {})[a] || [0, 0])[1] + (((A.c || {}).brand || {})[a] || [0, 0])[1])).slice(0, 14);
    const bt = tbl([L.brand, cur[0], cur[1], cur[2], L.dPrev, L.dYoy, L.share + ' ' + cur[0], L.samMin + ' ' + cur[0], L.upc + ' ' + cur[0], L.upc + ' ' + cur[2]], bset.map(b => { const g = (S, j) => (((S.c || {}).brand || {})[b] || [0, 0, 0, 0, 0])[j];
      return [b, N(g(A, 1)), N(g(Pv, 1)), N(g(Y, 1)), SP(g(Pv, 1) ? g(A, 1) / g(Pv, 1) - 1 : NaN), SP(g(Y, 1) ? g(A, 1) / g(Y, 1) - 1 : NaN), P(g(A, 1) / (A.q || 1), 1), N(g(A, 1) ? g(A, 2) * 60 / g(A, 1) : NaN), N(g(A, 4) ? g(A, 3) / g(A, 4) : NaN, 2), N(g(Y, 4) ? g(Y, 3) / g(Y, 4) : NaN, 2)]; }));
    return `<div class="two">${sec(L.seasonTable, `<div class="ch"><canvas id="cSea"></canvas></div>`)}${sec(L.yearTable, tbl([L.year, L.qty, vi ? 'Tăng/giảm' : '増減', L.stdH, L.cost, L.split], V.year.map(y => [String(y.year) + (y.full ? '' : ' *'), N(y.qty), SP(y.d), N(y.h), H(usd(y.c2), 'n'), H(usd(y.split), 'n')])))}</div>` + sec(L.seasonTable, tabS) +
      sec(L.cmpTitle, `<div class="cmpsel"><label for="sA">${L.cmpA}</label>${sel('sA', cur[0])}<label for="sP">${L.cmpP}</label>${sel('sP', cur[1])}<label for="sY">${L.cmpY}</label>${sel('sY', cur[2])}</div>` + mt) +
      sec(L.byBrand, bt) + `<div class="two">${sec(L.byLot, dim('lot', L.lotBands))}${sec(L.bySam, dim('sam', L.samBands))}</div>` + sec(L.byDay, dim('day', L.dayBands));
  }
  function afterSeasons() {
    const s = st.V.season; chart('cSea', { type: 'bar', data: { labels: s.map(x => x.season), datasets: [{ label: T().qty, data: s.map(x => x.qty), backgroundColor: s.map(x => x.season.endsWith('SS') ? C1() : C3()), borderRadius: 4 }] }, options: { plugins: { legend: { display: false } } } });
    for (const id of ['sA', 'sP', 'sY']) { const el = document.getElementById(id); if (el) el.onchange = () => { st.cmp = [$('#sA').value, $('#sP').value, $('#sY').value]; render(); }; }
  }

  function tabRisk() {
    const V = st.V, L = T(), R = V.riskK, vi = st.lang === 'vi';
    const k = [[vi ? 'Đang/sắp SX' : '生産中/予定', nf(R.all.n), kp(R.all.q) + ' pcs', '', 'good', vi ? 'Khối lượng' : '物量'], [L.lv.crit, nf(R.crit.n), kp(R.crit.q) + ' pcs', '≥30', 'critical', L.lv.crit], [L.lv.high, nf(R.high.n), kp(R.high.q) + ' pcs', '8–29', 'serious', L.lv.high], [L.lv.watch, nf(R.watch.n), kp(R.watch.q) + ' pcs', '1–7', 'warning', L.lv.watch], [L.lv.nofac, nf(R.nofac.n), kp(R.nofac.q) + ' pcs', '', 'warning', L.lv.nofac], [L.mat.late, nf(R.matLate.n), kp(R.matLate.q) + ' pcs', '', 'serious', L.mat.late], [L.mat.unk, nf(R.matUnk.n), kp(R.matUnk.q) + ' pcs', '', 'warning', L.mat.unk], [L.smp[0], nf(R.smp.n), kp(R.smp.q) + ' pcs', '', 'warning', L.smp[0]]];
    const q = st.riskQ.toLowerCase(); const rows = V.risk.filter(x => (st.riskFilter === 'all' || x.lv === st.riskFilter || (st.riskFilter === 'mat' && (x.mt === 'late' || x.mt === 'unk')) || (st.riskFilter === 'smp' && !x.sm)) && (!q || `${x.st} ${x.b} ${x.it}`.toLowerCase().includes(q)));
    const opts = [['all', L.all], ['crit', L.lv.crit], ['high', L.lv.high], ['watch', L.lv.watch], ['nofac', L.lv.nofac], ['mat', vi ? 'Vấn đề NVL' : '資材問題'], ['smp', L.smp[0]]];
    const filt = `<div class="cmpsel"><label for="rf">${L.filter}</label><select id="rf" class="sel">${opts.map(([v, l]) => `<option value="${v}" ${v === st.riskFilter ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select><label for="rq">${L.search}</label><input id="rq" class="sel" value="${esc(st.riskQ)}"><span class="muted">${rows.length} ${L.shown}</span></div>`;
    const t = tbl([L.line, L.season, L.brand, L.style, L.item, L.stage, L.qty, L.start, L.end, L.etdPlan, L.facEtd, L.gap, 'ETD', vi ? 'NVL' : '資材', vi ? 'Mẫu' : 'サンプル', L.valueUsd, L.comment], rows.map(x => [`${x.l}${x.p || ''}`, x.s, x.b, x.st, x.it || '', x.stg || '', N(x.q), dt(x.a), dt(x.e), dt(x.ep), dt(x.fe), x.g == null ? '' : N(x.g), H(esc(L.lv[x.lv] || ''), 'lv-' + x.lv), H(esc(L.mat[x.mt] || ''), x.mt === 'late' ? 'bad' : x.mt === 'unk' ? 'warn' : ''), H(esc(L.smp[x.sm]), x.sm ? '' : 'warn'), H(usd(x.v), 'n'), x.c]));
    return kpis(k) + sec(L.riskTitle, filt + t);
  }
  function afterRisk() { const f = $('#rf'), q = $('#rq'); if (f) f.onchange = () => { st.riskFilter = f.value; render(); }; if (q) q.onchange = () => { st.riskQ = q.value; render(); }; }

  function tabData() {
    const V = st.V, L = T(), vi = st.lang === 'vi', d = V.dq, c = V.counts;
    const issues = vi ? [
      ['Brand trống', `${d.blankN} dòng / ${nf(d.blankQ)} pcs`, 'Bắt buộc nhập Brand (danh sách thả xuống).'],
      ['Dòng chấm (chuyền lớn ghi 2 lần)', `${c.dotted} dòng / ${nf(c.dottedQty)} pcs – không cộng SL`, 'Đặt QTY dòng thứ 2 = 0 hoặc loại khỏi SUMIF; báo cáo đã tính lại lô 原価1 cho ' + c.nAdj1 + ' dòng.'],
      ['Hàng SMS', `${c.sms} dòng / ${nf(c.smsQty)} pcs – tách riêng`, 'Có thể thêm cột “Loại hàng” để phân biệt rõ.'],
      ['Mã mùa không thống nhất', `${c.seasonVariants} cách ghi`, 'Danh sách thả xuống cho cột Status.'],
      ['Ô lỗi công thức', `${c.errCells} ô`, 'Bọc IFERROR / sửa nguồn.'],
      ['Dòng mùa cũ treo lịch', `${d.stale.length} dòng (${d.stale.slice(0, 12).join(', ')})`, 'Xoá hoặc đánh dấu huỷ.'],
      ['Mã mùa không khớp lịch', `${d.chk.length} dòng (${d.chk.slice(0, 12).join(', ')})`, 'Xác nhận với Sales.'],
      ['Lệch ETD (cột S) vô nghĩa', `${d.gapBad} dòng`, 'Sửa công thức S: trả "" khi 工場ETD trống.'],
      ['Thiếu 工場ETD (đơn sắp SX)', `${d.noFac} đơn`, 'Sales nhập trước khi xếp line.'],
      ['Giai đoạn mẫu nhiều mã', `${d.stages} mã`, 'Chuẩn hoá 1st → SMS → PPS → BULK.'],
      ['Đơn giá GC bất thường', `${d.upcBad} dòng > 200 USD/pcs`, 'Tách đơn mẫu khỏi đơn sản xuất.'],
    ] : [
      ['ブランド空欄', `${d.blankN}行 / ${nf(d.blankQ)} pcs`, 'ブランドを必須入力(ドロップダウン)。'],
      ['網掛け行(大ラインの2重記載)', `${c.dotted}行 / ${nf(c.dottedQty)} pcs – 数量に加算せず`, '2行目のQTYを0にするかSUMIFから除外; 原価1ロットを' + c.nAdj1 + '行で再計算済み。'],
      ['SMS品', `${c.sms}行 / ${nf(c.smsQty)} pcs – 別集計`, '「品種区分」列の追加を推奨。'],
      ['シーズンコードの表記ゆれ', `${c.seasonVariants}種類`, 'Status列をドロップダウン化。'],
      ['数式エラーセル', `${c.errCells}セル`, 'IFERROR/元データ修正。'],
      ['旧シーズン行が残存', `${d.stale.length}行(${d.stale.slice(0, 12).join(', ')})`, '削除またはキャンセル表示。'],
      ['シーズンと日程の不一致', `${d.chk.length}行(${d.chk.slice(0, 12).join(', ')})`, '営業と確認。'],
      ['ETD差(S列)が無意味', `${d.gapBad}行`, 'S列の数式修正: 工場ETD空欄なら""。'],
      ['工場ETD未入力(予定)', `${d.noFac}件`, 'ライン割付前に営業が入力。'],
      ['サンプル段階コード多様', `${d.stages}種類`, '1st → SMS → PPS → BULKに統一。'],
      ['加工単価の異常値', `${d.upcBad}行 > 200 USD/pcs`, 'サンプルと量産を分ける。'],
    ];
    const method = vi ? [
      'Đọc sheet PLANTABLE (header dòng 9, dữ liệu từ dòng 10) ngay trong trình duyệt; file không được gửi đi đâu. Trang chỉ lưu báo cáo tóm tắt và lịch sử độ phủ.',
      'Xưởng APIC = gộp AP1A/AP1B/AP1C. Công suất 30 line × 32 CN = 960 CN; line nhỏ 小ロット = 16 CN; đơn AB chiếm cả A và B; không cộng trùng khi nhiều đơn chạy song song.',
      'Dòng chấm (QTY nền chấm) = dòng thứ 2 của chuyền lớn: không cộng SL, giờ chuẩn, chi phí; vẫn tính chiếm line.',
      'Hàng SMS: style/hợp đồng/item chứa SMS hoặc ghi chú SMS / 展示会サンプル / 展示会投入 (trừ “SMSサンプル対応”), SL ≤500.',
      'Lệch ETD = 工場ETD − ETD kế hoạch (âm = trễ): Nghiêm trọng ≤ −30, Cao −29…−8, Theo dõi −7…−1.',
      '受注進行表: đã chốt (ước tính) = đã bắt đầu may hoặc NVL OK (生地 OK & 資材 OK). Đã đối chiếu với 受注進行表 bản 09/11/2026: độ phủ lệch < 3%/tháng.',
      'Chi phí nhân công 46,29 USD/công-ngày (tham số AT5 của file gốc); mục tiêu tải 90%; các khoảng thời gian trượt theo ngày tham chiếu.',
    ] : [
      'PLANTABLEシート(見出し9行目、データ10行目〜)をブラウザ内で読取; ファイルは外部送信しない。保存するのは集計結果と充足率履歴のみ。',
      'APIC工場 = 旧AP1A/AP1B/AP1Cを統合。能力 30ライン × 32人 = 960人; 小ロット = 16人; ABは A・B 両方を占有; 並行オーダーは重複計上しない。',
      '網掛け行(QTY網掛け) = 大ラインの2行目: 数量・標準時間・費用に加算せず、ライン占有のみ計上。',
      'SMS品: 品番/契約/アイテムにSMS、またはコメントにSMS / 展示会サンプル / 展示会投入(「SMSサンプル対応」除く)、数量 ≤500。',
      'ETD差 = 工場ETD − 計画ETD(マイナス = 遅れ): 重大 ≤ −30、高 −29…−8、注意 −7…−1。',
      '受注進行表: 確定(推定) = 縫製開始済み or 資材OK(生地OK・資材OK)。2026/11/09版の受注進行表と照合: 月次差 < 3%。',
      '人件費 46.29 USD/人日(元ファイルAT5); 目標稼働率90%; 期間は基準日に応じてスライド。',
    ];
    return sec(L.dqTitle, tbl([L.issue, L.scale, L.fix], issues)) + sec(L.methodTitle, `<ul class="method">${method.map(m => `<li>${esc(m)}</li>`).join('')}</ul>`);
  }

  const TABS = { overview: [tabOverview], ceo: [tabCeo, afterCeo], factory: [tabFactory, afterFactory], prod: [tabProd], juchu: [tabJuchu, afterJuchu], seasons: [tabSeasons, afterSeasons], risk: [tabRisk, afterRisk], data: [tabData] };

  // ------------------------------------------------------------ khung trang
  function render() {
    const L = T();
    document.documentElement.lang = st.lang;
    $('#appTitle').textContent = L.appTitle; $('#appSub').textContent = L.appSub;
    $('#btnUpload').textContent = L.upload; $('#btnExcel').textContent = L.excel; $('#btnLang').textContent = L.lang; $('#lblRef').textContent = L.refDate;
    $('#dropTitle').textContent = L.dropTitle; $('#dropHint').textContent = L.dropHint;
    $('#btnExcel').disabled = !st.V || !st.dl; $('#btnExcel').hidden = !st.dl;
    $('#tabs').innerHTML = Object.keys(TABS).map(k => `<button class="tab ${k === st.tab ? 'on' : ''}" data-tab="${k}" type="button">${esc(L.tabs[k])}</button>`).join('');
    document.querySelectorAll('.tab').forEach(b => { b.onclick = () => { st.tab = b.dataset.tab; try { history.replaceState(null, '', '#' + st.tab); } catch (e) { /* */ } render(); }; });
    st.charts.forEach(c => c.destroy()); st.charts = [];
    if (!st.V) { $('#drop').hidden = false; $('#main').innerHTML = ''; $('#meta').textContent = ''; return; }
    $('#drop').hidden = true;
    const m = st.V.meta || {};
    $('#meta').textContent = `${L.lastReport}: ${m.file || '?'} · ${L.ref} ${dt(st.V.ref)} · ${L.fileDate} ${dt(st.V.snap)}${m.at ? ' · ' + L.uploadedAt + ' ' + new Date(m.at).toLocaleString(loc()) : ''}`;
    $('#refInput').value = PE.iso(st.V.ref);
    const [f, after] = TABS[st.tab] || TABS.overview;
    $('#main').innerHTML = f();
    after && after();
  }
  function status(msg, kind = 'info') { const el = $('#status'); el.textContent = msg || ''; el.className = 'status ' + kind; el.hidden = !msg; }

  // ------------------------------------------------------------ tải file
  function parseDates(name) {
    const m = /(20\d\d)[-_.]?(\d\d)[-_.]?(\d\d)/.exec(name || '');
    const today = (() => { const t = new Date(); return PE.dFromYMD(t.getFullYear(), t.getMonth() + 1, t.getDate()); })();
    if (m) { const f = PE.dFromYMD(+m[1], +m[2], +m[3]); return { snap: f, ref: f + 1 }; }
    return { snap: today, ref: today };
  }
  async function processFile(buf, name, refOverride) {
    const L = T();
    try {
      const dts = parseDates(name); if (refOverride != null) dts.ref = refOverride;
      const R = await PE.run(buf, window.JSZip, { ref: dts.ref, snap: dts.snap }, step => status(L.loading[step] || ''));
      st.V = PE.view(R, { file: name, at: Date.now() }); st.cmp = null;
      status(''); render();
      await saveReport();
    } catch (e) {
      status(e && e.message === 'NO_PLANTABLE' ? L.errNoPlantable : L.errRead + (e && e.message || e), 'err');
    }
  }
  async function saveReport() {
    const L = T(); if (!st.db) { status(L.savedLocal, 'warn'); return; }
    try {
      const parts = PE.splitView(st.V);
      for (const [k, v] of Object.entries(parts)) await st.db.doc('reports/latest-' + k).set({ json: v });
      await st.db.doc('reports/latest').set({ parts: Object.keys(parts), file: st.V.meta.file, at: st.V.meta.at, ref: PE.iso(st.V.ref), snap: PE.iso(st.V.snap) });
      const id = PE.iso(st.V.snap) + '_PLANTABLE';
      await st.db.doc('snapshots/' + id).set({ date: PE.iso(st.V.snap), src: 'PLANTABLE', months: st.V.juchu.mon.filter(m => m.mi >= st.V.m0).map(m => ({ m: mkey(m.mi), cap: m.cap, conf: Math.round(m.conf * 10) / 10, tot: Math.round(m.tot * 10) / 10 })) });
      status(`${L.savedOk} ${dt(st.V.snap)}.`, 'ok');
    } catch (e) { status(L.savedLocal, 'warn'); }
  }
  async function loadLatest() {
    if (!st.db) return;
    const timer = setTimeout(() => { if (!st.V) status(T().noDb, 'warn'); }, 10000);
    try {
      const head = await st.db.doc('reports/latest').get();
      if (head.exists && !st.V) {
        const h = head.data(); const V = {};
        for (const k of h.parts) { const p = await st.db.doc('reports/latest-' + k).get(); if (p.exists) Object.assign(V, k === 'core' ? JSON.parse(p.data().json) : JSON.parse(p.data().json)); }
        if (V.v && !st.V) { st.V = V; render(); }
      }
    } catch (e) { status(T().noDb, 'warn'); }
    clearTimeout(timer);
    try {
      st.db.collection('snapshots').onSnapshot(s => { st.hist = s.docs.map(d => Object.assign({ id: d.id }, d.data())); if (st.V && st.tab === 'juchu') render(); }, () => {});
    } catch (e) { /* bỏ qua */ }
  }

  // ------------------------------------------------------------ Excel
  async function exportExcel() {
    const L = T(), V = st.V; if (!V || !st.dl || !window.ExcelJS) return;
    status(L.excelBusy);
    try {
      const wb = new window.ExcelJS.Workbook(); wb.creator = 'APIC report';
      const font = { name: st.lang === 'vi' ? 'Arial' : 'Meiryo UI', size: 10 };
      const addSheet = (name, head, rows, widths) => {
        const ws = wb.addWorksheet(name.slice(0, 31)); ws.addRow(head);
        ws.getRow(1).eachCell(c => { c.font = Object.assign({}, font, { bold: true, color: { argb: 'FFFFFFFF' } }); c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1C5CAB' } }; c.alignment = { wrapText: true, vertical: 'middle' }; });
        for (const r of rows) ws.addRow(r).eachCell(c => { c.font = font; });
        ws.columns.forEach((col, i) => { col.width = (widths && widths[i]) || 14; }); ws.views = [{ state: 'frozen', ySplit: 1 }]; return ws;
      };
      const D = d => (d == null ? null : new Date(d * PE.DAY));
      const { F, A } = texts();
      const s1 = addSheet(L.tabs.overview, [L.findings, L.note], F.map(f => [f[0], f[1]]), [30, 140]); s1.getColumn(2).alignment = { wrapText: true };
      addSheet(L.actions, ['', L.note, L.owner, L.deadline], A, [34, 90, 22, 14]);
      addSheet(L.seasonTable, [L.season, L.rowsN, L.styles, L.qty, L.dPrev, L.dYoy, L.stdH, L.samMin, L.cost, L.upc, L.split, L.splitPct, L.lotMed, L.bespoke], V.season.map(x => [x.season, x.n, x.styles, x.qty, x.dPrev, x.dYoy, x.h, x.sam, x.c2, x.upc, x.split, x.splitPct, x.lotMed, x.bespoke]));
      addSheet(L.byBrand, [L.brand, ...V.seasons], V.brandTotal.map(b => [b.brand, ...V.seasons.map(s => (V.brandSeason[b.brand] || {})[s] || 0)]), [20]);
      addSheet(L.capacity, [L.month, L.wd, L.capacity, L.occ, L.util, L.qty, L.stdH, L.starts], V.cap.map(c => [ml(c.mi), c.wd, c.cap, c.occ, c.util, Math.round(c.qty), Math.round(c.h), c.starts]));
      const ms = []; for (let mi = V.m0; mi <= V.m0 + 7; mi++) ms.push(ml(mi));
      addSheet(L.heatTitle, [L.line, ...ms, L.orders, L.qty, L.freeFrom, L.smallLines], V.heat.map(h => [h.line, ...h.cells, h.n, h.qty, D(h.freeFrom), h.small.join(', ')]), [8]);
      const J = V.juchu;
      addSheet('受注進行表 ' + L.month, [L.month, L.wd, L.capacity, L.conf, L.tent, L.empty, L.confShare, L.totShare], J.mon.map(m => [ml(m.mi), m.wd, m.cap, m.conf, m.tot - m.conf, Math.max(m.cap - m.tot, 0), m.conf / m.cap, m.tot / m.cap]));
      const gws = wb.addWorksheet('受注進行表'); gws.addRow([L.line, ...J.per.map(p => p[2])]); gws.getRow(1).font = Object.assign({}, font, { bold: true });
      for (let ln = 1; ln <= 15; ln++) for (const h of ['A', 'B']) { const row = gws.addRow([ln + h, ...J.per.map((p, j) => { const c = J.grid[`${ln}|${h}|${j}`]; return c ? c.lab + (c.small ? '*' : '') : ''; })]);
        J.per.forEach((p, j) => { const c = J.grid[`${ln}|${h}|${j}`]; if (!c) return; const cell = row.getCell(j + 2); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c.lab.includes('_SMS') ? 'FFE4DEF7' : c.lab[0] === 'S' ? 'FFDCE9F9' : 'FFD9F2D9' } }; cell.font = Object.assign({}, font, { size: 8, color: { argb: c.tent ? 'FFD03B3B' : 'FF1F1F1F' }, bold: true }); }); }
      gws.columns.forEach((c, i) => { c.width = i ? 12 : 8; });
      addSheet(L.trend, [L.snapDate, L.source, wl(J.w1), wl(J.w1) + ' ' + L.conf, wl(J.w2), wl(J.w2) + ' ' + L.conf], trendRows().map(x => [x.date, x.src, x.t1, x.c1, x.t2, x.c2]));
      addSheet(L.riskTitle, [L.line, L.season, L.brand, L.style, L.item, L.qty, L.start, L.end, L.etdPlan, L.facEtd, L.gap, 'ETD', 'NVL', L.valueUsd, L.comment], V.risk.map(x => [`${x.l}${x.p || ''}`, x.s, x.b, x.st, x.it, x.q, D(x.a), D(x.e), D(x.ep), D(x.fe), x.g, L.lv[x.lv], L.mat[x.mt], x.v, x.c]));
      addSheet(L.readyTitle, [L.line, L.style, L.brand, L.qty, L.start, L.end, 'NVL', L.facEtd, 'ETD', L.readyY], V.ready.map(x => [`${x.l}${x.p || ''}`, x.st, x.b, x.q, D(x.a), D(x.e), L.mat[x.nvl], D(x.fe), L.lv[x.etd], x.ok ? L.readyY : x.blk.map(b => L.blk[b]).join(' · ')]));
      addSheet(L.ganttTitle, [L.line, ...V.weeks.map(w => dt(w))], V.gantt.map(g => [g.line, ...g.cells.map(c => (c ? c.text : ''))]), [8]);
      const buf = await wb.xlsx.writeBuffer();
      await st.dl.save({ filename: `${st.lang === 'vi' ? 'Bao_cao_dieu_hanh' : 'Keiei_Report'}_${PE.iso(V.ref)}.xlsx`, data: new Blob([buf]) });
      status(L.excelDone, 'ok');
    } catch (e) { status(e && e.code === 'declined' ? '' : (L.excelNo + (e && e.message ? ' (' + e.message + ')' : '')), e && e.code === 'declined' ? 'info' : 'warn'); }
  }

  // ------------------------------------------------------------ khởi động
  function init() {
    const fi = $('#file');
    $('#btnUpload').onclick = () => fi.click();
    fi.onchange = async () => { const f = fi.files[0]; if (!f) return; st.fileName = f.name; st.buf = await f.arrayBuffer(); processFile(st.buf, f.name); fi.value = ''; };
    const drop = document.body;
    drop.addEventListener('dragover', e => { e.preventDefault(); document.body.classList.add('dragging'); });
    drop.addEventListener('dragleave', e => { if (e.target === document.body) document.body.classList.remove('dragging'); });
    drop.addEventListener('drop', async e => { e.preventDefault(); document.body.classList.remove('dragging'); const f = e.dataTransfer.files[0]; if (!f) return; st.fileName = f.name; st.buf = await f.arrayBuffer(); processFile(st.buf, f.name); });
    $('#btnLang').onclick = () => { st.lang = st.lang === 'vi' ? 'ja' : 'vi'; try { localStorage.setItem('apic-lang', st.lang); } catch (e) { /* */ } render(); };
    $('#btnExcel').onclick = exportExcel;
    $('#refInput').onchange = e => { if (!st.buf || !e.target.value) return; const [y, m, d] = e.target.value.split('-').map(Number); processFile(st.buf, st.fileName, PE.dFromYMD(y, m, d)); };
    const h = (location.hash || '').slice(1); if (TABS[h]) st.tab = h;
    render();
    if (window.claude && window.claude.use) {
      window.claude.use('db').then(db => { st.db = db; if (db) loadLatest(); else if (!st.V) status(''); }).catch(() => {});
      window.claude.use('downloads').then(dl => { st.dl = dl; render(); }).catch(() => {});
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
