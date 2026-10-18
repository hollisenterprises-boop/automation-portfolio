"""Record a short animated demo of the dashboard: the state filter stepping through states.

    python make_demo_gif.py

Takes a headless Chrome screenshot of output/dashboard.html for each filter
(the page reads ?state=XX), adds a strip naming the filter, and writes
../images/sales-dashboard-demo.gif. Run build_dashboard.py first.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
PAGE = HERE / "output" / "dashboard.html"
OUT = HERE.parent / "images" / "sales-dashboard-demo.gif"
CHROME = Path("C:/Program Files/Google/Chrome/Application/chrome.exe")
FONTS = Path("C:/Windows/Fonts")
STATES = ["", "NSW", "VIC", "QLD", "WA"]
W, H, STRIP = 1600, 1200, 84
BG, INK, MUTED, ACCENT = (11, 18, 32), (241, 245, 249), (107, 127, 158), (34, 197, 94)


def screenshot(state: str, folder: Path) -> Image.Image:
    path = folder / f"{state or 'all'}.png"
    url = PAGE.resolve().as_uri() + (f"?state={state}" if state else "")
    subprocess.run([str(CHROME), "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    f"--window-size={W},{H}", "--virtual-time-budget=10000",
                    f"--user-data-dir={folder / 'profile'}", f"--screenshot={path}", url],
                   check=True, capture_output=True)
    return Image.open(path).convert("RGB")


def caption(shot: Image.Image, current: str) -> Image.Image:
    """Add a strip under the page that names each filter and highlights the current one."""
    frame = Image.new("RGB", (W, H + STRIP), BG)
    frame.paste(shot, (0, 0))
    draw = ImageDraw.Draw(frame)
    font = ImageFont.truetype(str(FONTS / "seguisb.ttf"), 30)
    x, y = 32, H + STRIP // 2
    draw.text((x, y), "State filter:", font=font, fill=MUTED, anchor="lm")
    x += draw.textlength("State filter:", font=font) + 24
    for state in STATES:
        label = state or "All states"
        width = draw.textlength(label, font=font) + 36
        if state == current:
            draw.rounded_rectangle([x, y - 24, x + width, y + 24], radius=24, fill=ACCENT)
        draw.text((x + width / 2, y), label, font=font, fill=BG if state == current else INK, anchor="mm")
        x += width + 14
    return frame


def main() -> None:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        frames = [caption(screenshot(state, Path(tmp)), state) for state in STATES]
    frames = [f.quantize(colors=256, method=Image.Quantize.MEDIANCUT) for f in frames]
    durations = [2600] + [1800] * (len(frames) - 1)
    frames[0].save(OUT, save_all=True, append_images=frames[1:], duration=durations, loop=0, optimize=True)
    print(f"{len(frames)} frames -> {OUT} ({OUT.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
