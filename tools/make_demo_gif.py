"""Record `xlgit demo` as the animated terminal at the top of the README.

    python tools/make_demo_gif.py [docs/demo.gif]

Runs the real demo and draws its real output; the only edit is shortening
the throwaway folder's path. Needs Pillow and a monospace font (Consolas on
Windows, DejaVu Sans Mono elsewhere). Regenerate it when the demo changes.
"""
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
COLS, SIZE, PAD, LINE = 104, 15, 18, 20
BG, FG, DIM = (24, 26, 31), (214, 218, 224), (128, 135, 148)
GREEN, YELLOW, BLUE, WHITE = (120, 200, 120), (230, 196, 100), (110, 170, 240), (245, 246, 248)


def font(bold=False):
    names = (["consolab.ttf", "DejaVuSansMono-Bold.ttf"] if bold else ["consola.ttf", "DejaVuSansMono.ttf"])
    for folder in ("C:/Windows/Fonts", "/usr/share/fonts/truetype/dejavu", "/Library/Fonts"):
        for n in names:
            if os.path.exists(os.path.join(folder, n)):
                return ImageFont.truetype(os.path.join(folder, n), SIZE)
    sys.exit("no monospace font found (install DejaVu Sans Mono)")


def demo_output():
    with tempfile.TemporaryDirectory() as tmp:
        folder = os.path.join(tmp, "xlgit-demo")
        r = subprocess.run([sys.executable, str(ROOT / "xlgit.py"), "demo", folder, "--no-open"],
                           capture_output=True, text=True, encoding="utf-8")
        if r.returncode:
            sys.exit(r.stdout + r.stderr)
        text = r.stdout.replace(os.path.abspath(folder), "~/xlgit-demo").replace("\\", "/")
    lines = []
    for line in text.rstrip().splitlines():
        if len(line) > COLS:
            line = line[:COLS - 1] + "…"
        lines.append(line)
    return lines


def colour(line):
    if re.match(r"^\d\. ", line):
        return WHITE, True
    if "<-" in line or "no conflicts" in line:
        return GREEN, False
    if line.startswith(("    row inserted", "    changed", "    object changed")):
        return YELLOW, False
    if line.startswith(("Visual diff", "Poke around", "To use", "A throwaway")):
        return DIM, False
    return FG, False


def main(out):
    regular, bold = font(), font(bold=True)
    char_w = regular.getbbox("M")[2]
    lines = demo_output()
    width = PAD * 2 + char_w * COLS
    height = PAD * 2 + LINE * (len(lines) + 2)
    frames, durations = [], []

    def frame(prompt_text, shown, cursor, ms):
        img = Image.new("RGB", (width, height), BG)
        d = ImageDraw.Draw(img)
        d.text((PAD, PAD), "$ ", font=bold, fill=GREEN)
        d.text((PAD + char_w * 2, PAD), prompt_text + ("_" if cursor else ""), font=bold, fill=WHITE)
        for i, line in enumerate(shown, start=1):
            fill, is_bold = colour(line)
            d.text((PAD, PAD + LINE * (i + 0.5)), line, font=bold if is_bold else regular, fill=fill)
        frames.append(img)
        durations.append(ms)

    command = "xlgit demo"
    frame("", [], True, 700)
    for i in range(1, len(command) + 1):
        frame(command[:i], [], True, 90)
    frame(command, [], False, 500)
    for n in range(1, len(lines) + 1):
        line = lines[n - 1]
        pause = 1400 if re.match(r"^\d\. ", line) else 2600 if "no conflicts" in line else 130
        frame(command, lines[:n], False, pause)
    durations[-1] = 6000
    frames[0].save(out, save_all=True, append_images=frames[1:], duration=durations, loop=0, optimize=True)
    print(f"{out}: {len(frames)} frames, {os.path.getsize(out) // 1024} KB, {width}x{height}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "docs" / "demo.gif"))
