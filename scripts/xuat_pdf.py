"""Ghép ảnh chụp trang flowchart thành PDF A4 ngang (song ngữ Nhật–Việt).

Chạy sau scripts/chup_flowchart.js:
    python3 scripts/xuat_pdf.py <thư_mục_ảnh> docs/Leadtime-TIJ.pdf
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

src, out = sys.argv[1], sys.argv[2]
DPI = 300
MM = DPI / 25.4
PAGE_W, PAGE_H = round(297 * MM), round(210 * MM)
MARGIN = round(10 * MM)
GAP = round(4 * MM)
FOOTER_H = round(6 * MM)
FONT_PATHS = ["/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"]

# Mỗi trang: danh sách ảnh xếp dọc. Ảnh cùng trang dùng chung một tỉ lệ.
PAGES = [
    ["01-header.png", "02-controls.png", "03-thesis.png"],
    ["04-flow-a.png"],
    ["05-flow-head.png", "06-flow-b.png"],
    ["07-findings-cases.png"],
    ["08-reasons.png"],
    ["09-asks.png", "10-notes.png"],
]
STACK_GAP = {("05-flow-head.png", "06-flow-b.png"): 0}

VI_FONT_PATHS = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"]


def load_font(paths):
    for p in paths:
        if os.path.exists(p):
            return ImageFont.truetype(p, round(2.6 * MM))
    return ImageFont.load_default()


font_ja = load_font(FONT_PATHS)   # chữ Nhật
font_vi = load_font(VI_FONT_PATHS)  # chữ Việt (font Nhật không có dấu tiếng Việt)

pages = []
for n, names in enumerate(PAGES, 1):
    ims = [Image.open(os.path.join(src, x)).convert("RGB") for x in names]
    gaps = [STACK_GAP.get((a, b), GAP) for a, b in zip(names, names[1:])]
    avail_w = PAGE_W - 2 * MARGIN
    avail_h = PAGE_H - 2 * MARGIN - FOOTER_H - sum(gaps)
    scale = min(avail_w / max(i.width for i in ims), avail_h / sum(i.height for i in ims))
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    y = MARGIN
    for i, im in enumerate(ims):
        w, h = round(im.width * scale), round(im.height * scale)
        page.paste(im.resize((w, h), Image.LANCZOS), ((PAGE_W - w) // 2, y))
        y += h + (gaps[i] if i < len(gaps) else 0)
    d = ImageDraw.Draw(page)
    x, y_text, color = MARGIN, PAGE_H - MARGIN - round(3 * MM), (91, 100, 130)
    for part, fnt in (("生産工程リードタイム・TIJ協議資料", font_ja),
                      ("   /   Lead time sản xuất – tài liệu thảo luận với TIJ", font_vi)):
        d.text((x, y_text), part, fill=color, font=fnt)
        x += d.textlength(part, font=fnt)
    num = f"{n} / {len(PAGES)}"
    d.text((PAGE_W - MARGIN - d.textlength(num, font=font_vi), y_text), num, fill=color, font=font_vi)
    pages.append(page)

pages[0].save(out, "PDF", resolution=DPI, save_all=True, append_images=pages[1:], quality=92)
print("saved", out, len(pages), "pages")
