"""
generate_contact_sheets.py — Generate visual contact sheets for site photos.
=============================================================================
Creates contact sheet PNGs (one per date) so you can visually identify which
photos belong to which construction phase before batch renaming.

Features:
  - .heic conversion via ffmpeg (Windows Pillow can't read .heic natively)
  - 3-column grid layout with photo index badges
  - Original filename labels under each thumbnail
  - Auto-fallback for corrupted/weird format images

Usage:
  1. Edit BASE to point to your photo directory (date subfolders)
  2. Edit OUT for contact sheet output location
  3. Edit FFMPEG path if needed (or install: scoop install ffmpeg)
  4. Run: python generate_contact_sheets.py
"""
import os, subprocess
from PIL import Image, ImageDraw, ImageFont

BASE = r"<PROJECTS_ROOT>\site_photos"          # <-- EDIT: photo root (date subfolders)
OUT = r"<PROJECTS_ROOT>\output\contact_sheets" # <-- EDIT: contact sheet output
os.makedirs(OUT, exist_ok=True)

FFMPEG = r"ffmpeg"  # <-- EDIT: or full path like r"C:\Tools\ffmpeg\bin\ffmpeg.exe"
TMP = os.path.join(OUT, "_conv")
os.makedirs(TMP, exist_ok=True)

def get_font(size):
    for p in [
        r"C:\Windows\Fonts\mingliu.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simsun.ttc",
    ]:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()

def to_rgb(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".heic":
        out = os.path.join(TMP, os.path.splitext(os.path.basename(path))[0] + ".png")
        if not os.path.exists(out):
            subprocess.run([FFMPEG, "-y", "-i", path, out],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        path = out
    try:
        im = Image.open(path).convert("RGB")
    except Exception:
        # fallback: try ffmpeg for any weird format
        out = os.path.join(TMP, os.path.splitext(os.path.basename(path))[0] + "_f.png")
        subprocess.run([FFMPEG, "-y", "-i", path, out],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        im = Image.open(out).convert("RGB")
    return im

dates = [d for d in sorted(os.listdir(BASE)) if os.path.isdir(os.path.join(BASE, d))]
for d in dates:
    folder = os.path.join(BASE, d)
    files = sorted([f for f in os.listdir(folder)
                    if f.lower().endswith((".jpg", ".jpeg", ".heic", ".png"))])
    if not files:
        print(f"[{d}] No photos, skipping")
        continue
    thumbs = []
    labels = []
    for f in files:
        try:
            im = to_rgb(os.path.join(folder, f))
        except Exception as e:
            print(f"  ! {d}/{f} read failed: {e}")
            continue
        im.thumbnail((360, 270))
        thumbs.append(im)
        labels.append(f)
    # build grid 3 cols
    cols = 3
    rows = (len(thumbs) + cols - 1) // cols
    cell_w = 380
    cell_h = 300
    pad = 8
    sheet_w = cols * cell_w + (cols + 1) * pad
    sheet_h = rows * cell_h + (rows + 1) * pad
    sheet = Image.new("RGB", (sheet_w, sheet_h), (255, 255, 255))
    font = get_font(15)
    for idx, (im, lab) in enumerate(zip(thumbs, labels)):
        r = idx // cols
        c = idx % cols
        x = pad + c * cell_w
        y = pad + r * cell_h
        ix = x + (cell_w - im.width) // 2
        iy = y + 4
        sheet.paste(im, (ix, iy))
        draw = ImageDraw.Draw(sheet)
        txt = f"{idx+1:02d}. {lab[:34]}"
        draw.text((x + 4, y + cell_h - 26), txt, fill=(0, 0, 0), font=font)
        draw.rectangle([x + 2, y + 2, x + 34, y + 24], fill=(220, 30, 30))
        draw.text((x + 6, y + 4), f"{idx+1:02d}", fill=(255, 255, 255), font=get_font(16))
    out_path = os.path.join(OUT, f"contact_{d}.png")
    sheet.save(out_path)
    print(f"[{d}] {len(thumbs)} photos -> {out_path}")
