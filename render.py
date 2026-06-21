# -*- coding: utf-8 -*-
import pathlib, time
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).parent
html = (HERE / "index.html").resolve().as_uri()
pdf_out = HERE / "자율전공_GATED_연구제안서.pdf"
thumb_out = HERE / "thumbnail.png"

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 1240, "height": 1754}, device_scale_factor=2)
    pg.goto(html, wait_until="networkidle")
    time.sleep(1.2)  # let webfont settle

    # PDF (A4, print CSS)
    pg.emulate_media(media="print")
    pg.pdf(path=str(pdf_out), format="A4", print_background=True,
           margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})

    # Thumbnail: cover area (screen media)
    pg.emulate_media(media="screen")
    pg.goto(html, wait_until="networkidle")
    time.sleep(0.8)
    pg.set_viewport_size({"width": 1000, "height": 1300})
    pg.screenshot(path=str(thumb_out), clip={"x": 0, "y": 0, "width": 1000, "height": 1300})
    b.close()

print("PDF  :", pdf_out, pdf_out.stat().st_size, "bytes")
print("THUMB:", thumb_out, thumb_out.stat().st_size, "bytes")
