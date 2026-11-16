# Excel Lab — Malayalam explainer clip generator
# Draws clean spreadsheet slides (Pillow canvas), shapes Malayalam text with
# HarfBuzz + renders with FreeType (Pillow alone lacks complex-script shaping
# on Windows), narrates with edge-tts (ml-IN neural voices), and assembles
# an H.264 MP4 with ffmpeg.
#
# Usage: python gen_clip_<lesson>.py  (spec file imports this module)

import os, sys, subprocess, shutil

W, H = 1280, 720
TEAL_DARK = (15, 118, 110)
TEAL = (13, 148, 136)
TEAL_SOFT = (204, 251, 241)
TEAL_FAINT = (240, 253, 250)
INK = (15, 23, 42)
INK2 = (71, 85, 105)
INK3 = (100, 116, 139)
BORDER = (226, 232, 240)
AMBER = (254, 240, 138)
CARD = (255, 255, 255)
BG = (246, 248, 250)

FONTS_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
_TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_clipwork")


# ---------------- fonts & shaped text engine ----------------
def _font_path(name):
    p = os.path.join(FONTS_DIR, name)
    return p if os.path.exists(p) else None

_REG = _font_path("Nirmala.ttf") or _font_path("segoeui.ttf") or _font_path("arial.ttf")
_BOLD = _font_path("NirmalaB.ttf") or _font_path("segoeuib.ttf") or _font_path("arialbd.ttf")
_MONO = _font_path("consola.ttf") or _REG
if not (_REG and _BOLD):
    raise SystemExit("No usable font found in " + FONTS_DIR)

# font key -> (file, px size)
FONTS = {
    "title":  (_BOLD, 44),
    "head":   (_BOLD, 30),
    "body":   (_REG, 29),
    "small":  (_REG, 24),
    "tiny":   (_BOLD, 20),
    "mono":   (_MONO, 26),
    "mono_s": (_MONO, 21),
}

import uharfbuzz as hb
import freetype as ft


class _Face:
    """One shaped/renderable font at a fixed pixel size."""
    def __init__(self, path, size):
        blob = hb.Blob.from_file_path(path)
        face = hb.Face(blob)
        upem = face.upem
        self.hb_font = hb.Font(face)
        self.hb_font.scale = (size * 64, size * 64)
        self.ft_face = ft.Face(path)
        self.ft_face.set_char_size(size * 64)
        self.ascender = self.ft_face.size.ascender / 64.0
        self._shape_cache = {}

    def shape(self, text):
        ops = self._shape_cache.get(text)
        if ops is not None:
            return ops
        buf = hb.Buffer()
        buf.add_str(text)
        buf.guess_segment_properties()
        hb.shape(self.hb_font, buf, {"kern": True, "liga": True, "clig": True})
        pen = 0.0
        ops = []
        for info, pos in zip(buf.glyph_infos, buf.glyph_positions):
            ops.append((info.codepoint, pen + pos.x_offset / 64.0,
                        pos.y_offset / 64.0, pos.x_advance / 64.0))
            pen += pos.x_advance / 64.0
        ops = (ops, pen)  # (glyph ops, total width)
        self._shape_cache[text] = ops
        return ops

    def draw(self, img, text, x, y, fill):
        """Draw text with top-left anchor matching PIL's default."""
        ops, _w = self.shape(text)
        baseline = y + self.ascender
        for gid, gx, gy, _adv in ops:
            self.ft_face.load_glyph(gid, ft.FT_LOAD_RENDER | ft.FT_LOAD_TARGET_NORMAL)
            g = self.ft_face.glyph
            bm = g.bitmap
            if bm.width == 0 or bm.rows == 0:
                continue
            gx += g.bitmap_left
            gy = baseline - gy - g.bitmap_top
            px, py = int(round(x + gx)), int(round(gy))
            if px >= img.width or py >= img.height or px + bm.width <= 0 or py + bm.rows <= 0:
                continue
            glyph_img = Image.frombytes("L", (bm.width, bm.rows), bytes(bm.buffer))
            solid = Image.new("RGB", glyph_img.size, fill)
            img.paste(solid, (px, py), glyph_img)


class _Shaper:
    def __init__(self):
        self._faces = {}

    def face(self, key):
        if key not in self._faces:
            path, size = FONTS[key]
            self._faces[key] = _Face(path, size)
        return self._faces[key]

    def width(self, text, key):
        return self.face(key).shape(text)[1]

    def draw(self, img, text, key, x, y, fill):
        self.face(key).draw(img, text, x, y, fill)

    def wrap(self, text, key, max_w):
        words = text.split()
        lines, cur = [], ""
        for w_ in words:
            t = (cur + " " + w_).strip()
            if self.width(t, key) <= max_w or not cur:
                cur = t
            else:
                lines.append(cur)
                cur = w_
        if cur:
            lines.append(cur)
        return lines


S = _Shaper()
from PIL import Image, ImageDraw  # noqa: E402


# ---------------- drawing helpers ----------------
def _rr(d, box, r, fill=None, outline=None, width=1):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


def draw_frame(slide, idx, total, lesson_title, out_png):
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    # header band
    d.rectangle([0, 0, W, 64], fill=TEAL_DARK)
    S.draw(img, "Excel Lab", "tiny", 28, 14, (255, 255, 255))
    S.draw(img, lesson_title, "tiny", W - 28 - S.width(lesson_title, "tiny"), 14, TEAL_SOFT)

    # slide title
    d.rectangle([28, 96, 36, 142], fill=TEAL)
    S.draw(img, slide["title"], "title", 52, 94, INK)

    y = 176
    grid = slide.get("grid")
    steps = slide.get("steps")
    bullets = slide.get("bullets")

    if bullets:
        for b in bullets:
            d.ellipse([62, y + 12, 76, y + 26], fill=TEAL)
            for ln in S.wrap(b, "body", 900):
                S.draw(img, ln, "body", 92, y, INK2)
                y += 44
            y += 10
        y += 8

    if steps:
        n = 0
        for s in steps:
            n += 1
            _rr(d, [56, y, W - 56, y + 54], 10, fill=CARD, outline=BORDER)
            d.ellipse([68, y + 13, 94, y + 39], fill=TEAL)
            nx = S.width(str(n), "tiny")
            S.draw(img, str(n), "tiny", 81 - nx / 2, y + 16, (255, 255, 255))
            ty = y + 26 - 15 * len(S.wrap(s, "body", 880))
            for ln in S.wrap(s, "body", 880):
                S.draw(img, ln, "body", 110, ty, INK2)
                ty += 30
            y += 66

    if grid:
        gy = max(y + 6, 196)
        _draw_grid(img, d, grid, gy)

    # caption strip
    cap = slide.get("caption")
    if cap:
        lines = S.wrap(cap, "head", 1080)
        ch = 24 + len(lines) * 40
        _rr(d, [28, H - 28 - ch, W - 28, H - 28], 14, fill=TEAL_FAINT, outline=TEAL_SOFT, width=2)
        cy = H - 28 - ch + 18
        for ln in lines:
            S.draw(img, ln, "head", 52, cy, TEAL_DARK)
            cy += 40

    # slide counter
    lbl = "%d / %d" % (idx, total)
    S.draw(img, lbl, "small", W - 40 - S.width(lbl, "small"), H - 46, INK3)

    img.save(out_png)


def _draw_grid(img, d, grid, y_top):
    headers = grid.get("headers")
    rows = grid.get("rows", [])
    hl = set(tuple(x) for x in grid.get("highlight", []))  # (row_idx, col_idx) data-only
    title = grid.get("title")

    ncols = max([len(headers) if headers else 0] + [len(r) for r in rows])

    def cellmax(ci):
        w = 0
        if headers and ci < len(headers):
            w = max(w, len(str(headers[ci])))
        for r in rows:
            if ci < len(r):
                w = max(w, len(str(r[ci])))
        return w

    colw = [max(64, 22 * (cellmax(ci) + 2)) for ci in range(ncols)]
    total_w = sum(colw) + 34
    x0 = max(60, (W - total_w) // 2)

    if title:
        S.draw(img, title, "tiny", x0 + 6, y_top - 4, INK3)
        y_top += 30

    rh = 38
    y = y_top
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    d.rectangle([x0, y, x0 + 34, y + 28], fill=(241, 245, 249), outline=BORDER)
    cx = x0 + 34
    for ci in range(ncols):
        d.rectangle([cx, y, cx + colw[ci], y + 28], fill=(241, 245, 249), outline=BORDER)
        lw = S.width(letters[ci], "tiny")
        S.draw(img, letters[ci], "tiny", cx + (colw[ci] - lw) / 2, y + 3, INK3)
        cx += colw[ci]
    y += 28

    if headers:
        d.rectangle([x0, y, x0 + 34, y + rh], fill=(241, 245, 249), outline=BORDER)
        cx = x0 + 34
        for ci, hd in enumerate(headers):
            d.rectangle([cx, y, cx + colw[ci], y + rh], fill=(241, 245, 249), outline=BORDER)
            lw = S.width(str(hd), "tiny")
            S.draw(img, str(hd), "tiny", cx + (colw[ci] - lw) / 2, y + 8, INK2)
            cx += colw[ci]
        y += rh

    for ri, row in enumerate(rows):
        d.rectangle([x0, y, x0 + 34, y + rh], fill=(241, 245, 249), outline=BORDER)
        rw = S.width(str(ri + 1), "tiny")
        S.draw(img, str(ri + 1), "tiny", x0 + 17 - rw / 2, y + 8, INK3)
        cx = x0 + 34
        for ci in range(ncols):
            val = str(row[ci]) if ci < len(row) else ""
            is_h = (ri, ci) in hl
            d.rectangle([cx, y, cx + colw[ci], y + rh],
                        fill=AMBER if is_h else CARD, outline=BORDER)
            if is_h:
                d.rectangle([cx - 1, y - 1, cx + colw[ci] + 1, y + rh + 1],
                            outline=TEAL, width=3)
            if val:
                align_r = val[:1].isdigit() or val[:1] in "+-(" or val[:1] == "="
                if align_r:
                    lw = S.width(val, "mono_s")
                    S.draw(img, val, "mono_s", cx + colw[ci] - 10 - lw, y + 8,
                           TEAL_DARK if val[:1] == "=" else INK2)
                else:
                    S.draw(img, val, "mono_s", cx + 10, y + 8, INK2)
            cx += colw[ci]
        y += rh


# ---------------- audio + assembly ----------------
def _run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        raise SystemExit("FAILED: " + " ".join(map(str, cmd)) + "\n" + (r.stderr or r.stdout)[-1500:])
    return r


def _duration(path):
    r = _run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
              "-of", "csv=p=0", path])
    return float(r.stdout.strip().splitlines()[-1])


def _narrate(text, mp3, voice, rate="+0%", tries=3):
    for t in range(1, tries + 1):
        r = subprocess.run(
            [sys.executable, "-m", "edge_tts", "--voice", voice, "--rate=" + rate,
             "--text", text, "--write-media", mp3],
            capture_output=True, text=True)
        if r.returncode == 0 and os.path.exists(mp3) and os.path.getsize(mp3) > 1000:
            return
        print("  edge-tts attempt %d failed: %s" % (t, (r.stderr or "").strip()[-300:]))
    raise SystemExit("edge-tts failed for text: " + text[:60])


def build(slides, lesson_title, out_mp4, voice="ml-IN-MidhunNeural", rate="+0%"):
    os.makedirs(_TMP, exist_ok=True)
    work = os.path.join(_TMP, os.path.splitext(os.path.basename(out_mp4))[0])
    if os.path.exists(work):
        shutil.rmtree(work)
    os.makedirs(work)
    total = len(slides)
    segs = []
    for i, s in enumerate(slides, 1):
        png = os.path.join(work, "slide_%02d.png" % i)
        draw_frame(s, i, total, lesson_title, png)
        mp3 = os.path.join(work, "narr_%02d.mp3" % i)
        _narrate(s["narr"], mp3, voice, rate)
        dur = _duration(mp3) + 0.9  # small pause after each slide
        seg = os.path.join(work, "seg_%02d.mp4" % i)
        _run(["ffmpeg", "-y", "-loop", "1", "-framerate", "6", "-i", png, "-i", mp3,
              "-filter_complex", "[1:a]apad=pad_dur=0.9[a]",
              "-map", "0:v", "-map", "[a]", "-t", "%.2f" % dur,
              "-c:v", "libx264", "-preset", "veryfast", "-tune", "stillimage",
              "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "96k", seg])
        segs.append(seg)
        print("  slide %d/%d  %.1fs" % (i, total, dur))
    lst = os.path.join(work, "list.txt")
    with open(lst, "w", encoding="utf-8") as fh:
        for seg in segs:
            fh.write("file '%s'\n" % seg.replace("\\", "/"))
    os.makedirs(os.path.dirname(out_mp4), exist_ok=True)
    _run(["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", lst,
          "-c", "copy", "-movflags", "+faststart", out_mp4])
    shutil.copyfile(os.path.join(work, "slide_01.png"),
                    os.path.splitext(out_mp4)[0] + ".jpg")
    print("OK %s (%.1f KB)" % (out_mp4, os.path.getsize(out_mp4) / 1024))
