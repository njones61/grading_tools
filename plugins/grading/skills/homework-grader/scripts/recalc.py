#!/usr/bin/env python3
"""
Recalculate Excel formulas using LibreOffice.

Usage:
    python recalc.py <path_to_xlsx>

openpyxl writes formulas as text -- they carry no cached value until Excel or
LibreOffice opens the file and computes them. Anything that reads scores.xlsx
with data_only=True sees None until that happens. This opens the file in
headless LibreOffice and re-saves it, which forces the recalculation.

LibreOffice ships its binary under different names on different platforms:
`libreoffice` on most Linux installs, `soffice` inside the app bundle on
macOS. Both are looked for on PATH first, then at the usual install paths.
"""
import os
import shutil
import subprocess
import sys
import tempfile

# Tried in order: PATH lookups first, then platform install locations.
CANDIDATES = (
    "libreoffice",
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/bin/libreoffice",
    "/snap/bin/libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
)


def find_libreoffice():
    """Return a runnable LibreOffice path, or exit with install instructions."""
    for candidate in CANDIDATES:
        found = shutil.which(candidate)
        if found:
            return found
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    sys.exit(
        "error: LibreOffice not found. Install it and re-run:\n"
        "  macOS   brew install --cask libreoffice\n"
        "  Debian  sudo apt install libreoffice-calc"
    )


def recalc(filepath):
    filepath = os.path.abspath(filepath)
    if not os.path.isfile(filepath):
        sys.exit(f"error: no such file: {filepath}")

    soffice = find_libreoffice()
    basename = os.path.basename(filepath)

    # Convert into a scratch directory rather than /tmp directly, so two
    # assignments recalculating a file both named scores.xlsx cannot collide.
    with tempfile.TemporaryDirectory() as outdir:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "xlsx",
             "--outdir", outdir, filepath],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            sys.exit(f"error: LibreOffice failed: {result.stderr.strip()}")

        converted = os.path.join(outdir, basename)
        if not os.path.exists(converted):
            sys.exit(
                "error: LibreOffice produced no output file. This usually "
                "means another LibreOffice instance is already running -- "
                "quit it and try again."
            )
        shutil.copy2(converted, filepath)

    print(f"Recalculated: {filepath}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(f"Usage: {sys.argv[0]} <path_to_xlsx>")
    recalc(sys.argv[1])
