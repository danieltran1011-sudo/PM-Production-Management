// Chụp từng phần của trang flowchart (workflow/flowchart.html) để ghép thành PDF.
// Cần playwright-core và Chromium:
//   NODE_PATH=<thư mục có playwright-core> node scripts/chup_flowchart.js <thư_mục_ảnh> [đường_dẫn_chromium]
// Sau đó: python3 scripts/xuat_pdf.py <thư_mục_ảnh> docs/Leadtime-TIJ.pdf
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright-core');

(async () => {
  const outDir = process.argv[2];
  const exe = process.argv[3] || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
  fs.mkdirSync(outDir, { recursive: true });
  const src = 'file://' + path.resolve(__dirname, '../workflow/flowchart.html');
  const browser = await chromium.launch({ executablePath: exe });
  const page = await browser.newPage({ viewport: { width: 1240, height: 900 }, deviceScaleFactor: 3, colorScheme: 'light' });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(src);
  await page.evaluate(() => document.fonts && document.fonts.ready);
  await page.waitForTimeout(800);

  const box = (sel) => page.$eval(sel, (el) => {
    const r = el.getBoundingClientRect();
    return { x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height };
  });
  const shot = async (name, clip) => {
    await page.screenshot({ path: path.join(outDir, name), clip: { x: clip.x, y: clip.y, width: clip.w, height: clip.h }, fullPage: true });
  };
  const pad = (b, p) => ({ x: b.x - p, y: b.y - p, w: b.w + 2 * p, h: b.h + 2 * p });

  await shot('01-header.png', pad(await box('header.top'), 6));
  await shot('02-controls.png', pad(await box('section[aria-label="Điều kiện"]'), 4));
  await shot('03-thesis.png', pad(await box('section.thesis'), 4));

  // Flowchart: chia 2 phần tại khoảng trống phía trên một hàng (xả vải / sơ đồ / cắt),
  // chọn hàng làm hai phần cao gần bằng nhau; phần 2 lặp lại hàng tiêu đề bộ phận.
  const sec = await page.$eval('section.thesis + section', (el) => {
    const r = el.getBoundingClientRect();
    const flow = el.querySelector('.flow').getBoundingClientRect();
    const head = el.querySelector('.lane-head').getBoundingClientRect();
    const top = r.top + scrollY, bottom = r.bottom + scrollY, headH = head.bottom - flow.top;
    let best = null;
    ['xa', 'sodo', 'cat'].forEach((id) => {
      const n = el.querySelector('[data-node="' + id + '"]');
      if (!n) return;
      const split = n.getBoundingClientRect().top + scrollY - 13;
      const worst = Math.max(split - top, bottom - split + headH);
      if (!best || worst < best.worst) best = { split, worst };
    });
    return {
      x: r.left + scrollX, y: top, w: r.width, h: r.height,
      flowTop: flow.top + scrollY, headBottom: head.bottom + scrollY, split: best.split
    };
  });
  await shot('04-flow-a.png', { x: sec.x - 4, y: sec.y - 4, w: sec.w + 8, h: sec.split - sec.y + 4 });
  await shot('05-flow-head.png', { x: sec.x - 4, y: sec.flowTop, w: sec.w + 8, h: sec.headBottom - sec.flowTop });
  await shot('06-flow-b.png', { x: sec.x - 4, y: sec.split, w: sec.w + 8, h: sec.y + sec.h - sec.split + 4 });

  const twos = await page.$$('section.two');
  const panels = async (el) => el.evaluate((n) => {
    const r = n.getBoundingClientRect();
    return { x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height };
  });
  await shot('07-findings-cases.png', pad(await panels(twos[0]), 4));
  const reasonsHead = await page.$eval('#reasons', (el) => {
    const r = el.parentElement.getBoundingClientRect();
    return { x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height };
  });
  await shot('08-reasons.png', pad(reasonsHead, 4));
  await shot('09-asks.png', pad(await panels(twos[1]), 4));
  await shot('10-notes.png', pad(await box('section.notes'), 4));

  console.log('ok', outDir, errors.length ? 'errors: ' + errors.join(' | ') : 'no page errors');
  await browser.close();
})();
