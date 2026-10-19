"""Make a short silent demo video (1280x720 MP4, about 45 seconds) of the portfolio.

    python make_demo_video.py

For Fiverr gig videos and Upwork catalog and portfolio videos: a title card,
each demo's cover with a one-line caption and a slow zoom, the moving
dashboard, and a closing card. Uses the images in images/ (run make_covers.py
and sales-dashboard/make_demo_gif.py first). Needs imageio and imageio-ffmpeg.
Writes images/demo-video.mp4.
"""

from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageSequence

from make_covers import BG_BOTTOM, BG_TOP, FAINT, MUTED, WHITE, font

HERE = Path(__file__).parent
IMAGES = HERE / "images"
OUT = IMAGES / "demo-video.mp4"
W, H, FPS = 1280, 720, 30
ACCENT = (34, 197, 94)

DEMOS = [
    ("customer-insights", "Customer insights on 541,909 real order lines"),
    ("sales-dashboard", "An interactive sales dashboard in one file"),
    ("sales-report", "A messy store export becomes a clean Excel report"),
    ("invoice-extractor", "Supplier PDF invoices read into one checked sheet"),
    ("list-cleaner", "Three messy contact lists merged into one clean list"),
    ("price-tracker", "Competitor prices and stock, tracked daily"),
    ("ads-waste-finder", "Amazon ad searches that never sell, found"),
]


def background():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    for y in range(H):
        t = y / (H - 1)
        draw.line([(0, y), (W, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(BG_TOP, BG_BOTTOM)))
    for x in range(0, W, 32):
        for y in range(0, H, 32):
            draw.point((x, y), fill=FAINT)
    return img


def card(lines):
    """A text card: (text, size, colour, bold) per line, centred."""
    img = background()
    draw = ImageDraw.Draw(img)
    heights = [size * 1.35 for _, size, _, _ in lines]
    y = (H - sum(heights)) / 2
    for (text, size, colour, bold), h in zip(lines, heights):
        draw.text((W / 2, y + h / 2), text, font=font("seguisb.ttf" if bold else "segoeui.ttf", size),
                  fill=colour, anchor="mm")
        y += h
    return img


def framed(image, caption):
    """An image fitted above a caption strip on the dark background."""
    img = background()
    area_h = H - 110
    pic = image.convert("RGB")
    pic.thumbnail((W - 80, area_h - 30))
    img.paste(pic, ((W - pic.width) // 2, 24))
    ImageDraw.Draw(img).text((W / 2, H - 52), caption, font=font("seguisb.ttf", 34), fill=WHITE, anchor="mm")
    return img


def zoom(img, t):
    """Slow push-in: t from 0 to 1 zooms from 100% to 104%."""
    scale = 1 + 0.04 * t
    cw, ch = W / scale, H / scale
    left, top = (W - cw) / 2, (H - ch) / 2
    return img.crop((left, top, left + cw, top + ch)).resize((W, H), Image.BILINEAR)


def scenes():
    yield card([("Automation that saves you hours", 64, WHITE, True),
                ("Excel & Google Sheets  ·  Python  ·  Web scraping  ·  Dashboards", 32, MUTED, False),
                ("Seven working demos, each with automated tests", 30, ACCENT, False)]), 3.0, True
    for name, caption in DEMOS:
        yield framed(Image.open(IMAGES / f"cover-{name}.png"), caption), 4.0, True
        if name == "sales-dashboard":
            gif = Image.open(IMAGES / "sales-dashboard-demo.gif")
            for frame in ImageSequence.Iterator(gif):
                yield framed(frame.copy(), "Filter by state: every number updates"), 1.6, False
    yield card([("Fixed price  ·  Built and tested on your data  ·  Plain instructions", 34, WHITE, True),
                ("Send me a sample file and I'll tell you the fastest way to do it.", 30, MUTED, False)]), 4.0, False


def main():
    fade = int(0.4 * FPS)
    previous = None
    with imageio.get_writer(OUT, fps=FPS, codec="libx264", quality=6, macro_block_size=8,
                            ffmpeg_params=["-pix_fmt", "yuv420p", "-movflags", "+faststart"]) as video:
        for still, seconds, moving in scenes():
            count = int(seconds * FPS)
            for i in range(count):
                frame = zoom(still, i / count) if moving else still
                if previous is not None and i < fade:
                    frame = Image.blend(previous, frame, (i + 1) / fade)
                video.append_data(np.asarray(frame))
            previous = frame
    print(f"Wrote {OUT} ({OUT.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
