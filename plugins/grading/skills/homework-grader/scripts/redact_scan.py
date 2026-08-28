#!/usr/bin/env python3
"""Black out the header band of a scanned submission, locally and destructively.

Usage:
    redact_scan.py <file_or_dir> [--band 6] [--dpi 200] [--force] [--no-preview]

A scanned page carries the student's handwritten name as pixels. Nothing in
the OOXML scrubber can touch it, so scans reach the grader identifying.

Everything here runs on this machine. That is the point: sending the page to
a vision model to find the name would transmit the very thing the redaction
exists to withhold, so detection is a fixed geometric band rather than
anything clever. The band is blunt, and a name written down the margin or on
a later page will survive it -- which is why a preview PNG is written for a
human to check locally before grading begins.

The redaction is destructive by construction. Drawing a black rectangle into
the PDF would hide the name on screen while leaving the pixels underneath
extractable; instead the page is rasterized, the band is painted onto the
bitmap, and the PDF is rebuilt from the painted image. The original pixels
are gone, and so is any OCR text layer the scanner embedded -- which is its
own leak, since scanner OCR sometimes reads a printed name correctly.

Adobe Scan and friends embed an OCR text layer over the page image, so
"does this PDF contain text?" does not distinguish a scan from a born-digital
export. The test used here is whether a raster image covers essentially the
whole page.
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

DEFAULT_BAND_PCT = 6.0
DEFAULT_DPI = 200

PDF_EXT = {".pdf"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
# Needs pillow-heif, which is usually absent; reported rather than silently missed.
UNSUPPORTED_EXT = {".heic", ".heif"}

# A raster covering at least this fraction of the page means the page is a scan.
FULL_PAGE_COVERAGE = 0.80


def _pdftoppm():
    return shutil.which("pdftoppm")


def page_sizes(path):
    """Page dimensions in points, via pypdf. Returns [] if pypdf is unavailable."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return []
    try:
        return [(float(p.mediabox.width), float(p.mediabox.height))
                for p in PdfReader(str(path)).pages]
    except Exception:
        return []


def is_scanned_pdf(path):
    """True when a raster image covers essentially the whole of any page."""
    exe = shutil.which("pdfimages")
    if not exe:
        return False
    try:
        out = subprocess.run([exe, "-list", str(path)],
                             capture_output=True, text=True, timeout=60).stdout
    except Exception:
        return False

    sizes = page_sizes(path)
    for line in out.splitlines()[2:]:
        f = line.split()
        if len(f) < 15:
            continue
        try:
            page, w_px, h_px = int(f[0]), int(f[3]), int(f[4])
            x_ppi, y_ppi = float(f[12]), float(f[13])
        except ValueError:
            continue
        if x_ppi <= 0 or y_ppi <= 0 or page > len(sizes):
            continue
        pw, ph = sizes[page - 1]
        if pw <= 0 or ph <= 0:
            continue
        # Convert the image back to points and compare against the page.
        if ((w_px / x_ppi * 72) / pw >= FULL_PAGE_COVERAGE
                and (h_px / y_ppi * 72) / ph >= FULL_PAGE_COVERAGE):
            return True
    return False


def _paint_band(im, band_pct):
    from PIL import ImageDraw
    w, h = im.size
    band = max(1, int(h * band_pct / 100.0))
    ImageDraw.Draw(im).rectangle([0, 0, w, band], fill=(0, 0, 0))
    return im


def redact_pdf(path, band_pct=DEFAULT_BAND_PCT, dpi=DEFAULT_DPI, preview=True):
    """Rasterize, paint the band on every page, rebuild the PDF in place."""
    from PIL import Image

    exe = _pdftoppm()
    if not exe:
        return "skipped", "pdftoppm not found (install poppler)"

    with tempfile.TemporaryDirectory() as tmp:
        stem = str(Path(tmp) / "pg")
        proc = subprocess.run([exe, "-r", str(dpi), "-png", str(path), stem],
                              capture_output=True, text=True, timeout=300)
        pages = sorted(Path(tmp).glob("pg*.png"))
        if proc.returncode != 0 and not pages:
            return "error", f"pdftoppm failed: {proc.stderr.strip()[:200]}"
        if not pages:
            return "error", "pdftoppm produced no pages"

        images = [_paint_band(Image.open(p).convert("RGB"), band_pct) for p in pages]
        out_tmp = path.with_suffix(".pdf.tmp")
        try:
            images[0].save(out_tmp, "PDF", resolution=float(dpi), save_all=True,
                           append_images=images[1:])
            out_tmp.replace(path)
        except Exception as exc:
            out_tmp.unlink(missing_ok=True)
            return "error", str(exc)

        if preview:
            _write_preview(path, images[0])
    return "redacted", f"{len(pages)} page(s), top {band_pct:g}%"


def redact_image(path, band_pct=DEFAULT_BAND_PCT, preview=True):
    from PIL import Image
    try:
        im = _paint_band(Image.open(path).convert("RGB"), band_pct)
        im.save(path)
        if preview:
            _write_preview(path, im)
    except Exception as exc:
        return "error", str(exc)
    return "redacted", f"top {band_pct:g}%"


def _write_preview(path, first_page):
    """Downscaled full-page PNG, so a human can spot a name outside the band."""
    prev_dir = path.parent / "redaction_previews"
    prev_dir.mkdir(exist_ok=True)
    im = first_page.copy()
    im.thumbnail((1000, 1000))
    im.save(prev_dir / f"{path.stem}_preview.png")


def redact(path, band_pct=DEFAULT_BAND_PCT, dpi=DEFAULT_DPI, force=False, preview=True):
    """Redact one file. Returns (status, detail).

    status: redacted | skipped | error
    """
    path = Path(path)
    ext = path.suffix.lower()
    try:
        import PIL  # noqa: F401
    except ImportError:
        return "skipped", "Pillow not installed"

    if ext in UNSUPPORTED_EXT:
        return "skipped", f"{ext} needs pillow-heif -- convert it by hand"
    if ext in IMAGE_EXT:
        return redact_image(path, band_pct, preview)
    if ext in PDF_EXT:
        if not force and not is_scanned_pdf(path):
            return "skipped", "not a scan (no full-page image) -- use --force to redact anyway"
        return redact_pdf(path, band_pct, dpi, preview)
    return "skipped", f"nothing to redact in {ext}"


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target", type=Path, help="a file, or a directory to sweep")
    ap.add_argument("--band", type=float, default=DEFAULT_BAND_PCT,
                    help=f"percent of page height to black out (default: {DEFAULT_BAND_PCT:g})")
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI,
                    help=f"rasterization resolution (default: {DEFAULT_DPI})")
    ap.add_argument("--force", action="store_true",
                    help="redact a PDF even if it does not look like a scan")
    ap.add_argument("--no-preview", action="store_true", help="skip preview PNGs")
    args = ap.parse_args()

    target = args.target.expanduser().resolve()
    if target.is_dir():
        files = [f for f in sorted(target.iterdir())
                 if f.is_file() and f.suffix.lower() in (PDF_EXT | IMAGE_EXT | UNSUPPORTED_EXT)]
    elif target.is_file():
        files = [target]
    else:
        sys.exit(f"error: no such file or directory: {target}")

    if not files:
        sys.exit(f"error: nothing redactable in {target}")

    counts = {}
    for f in files:
        status, detail = redact(f, args.band, args.dpi, args.force, not args.no_preview)
        counts[status] = counts.get(status, 0) + 1
        print(f"  {status:9} {f.name}  -- {detail}")

    if counts.get("redacted"):
        print(f"\nCheck {files[0].parent / 'redaction_previews'} before grading -- the band is "
              "geometric, so a name written elsewhere on the page will have survived it.")
    return 1 if counts.get("error") else 0


if __name__ == "__main__":
    sys.exit(main())
