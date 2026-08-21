#!/usr/bin/env python3
"""
Recalculate Excel formulas using LibreOffice.

Usage:
    python recalc.py <path_to_xlsx>

This script opens the xlsx file in LibreOffice headless mode, which
forces all formulas to recompute, then saves the result back.
This is needed after generating scores.xlsx because openpyxl writes
formulas as text — they don't have cached computed values until
Excel or LibreOffice opens and recalculates them.
"""
import sys
import os
import subprocess
import shutil

def recalc(filepath):
    filepath = os.path.abspath(filepath)
    dirname = os.path.dirname(filepath)
    basename = os.path.basename(filepath)

    # Use LibreOffice to open and re-save (forces formula recalculation)
    result = subprocess.run(
        ['libreoffice', '--headless', '--calc', '--convert-to', 'xlsx',
         '--outdir', '/tmp', filepath],
        capture_output=True, text=True, timeout=60
    )

    if result.returncode != 0:
        print(f"Error: {result.stderr}")
        sys.exit(1)

    # Copy back
    tmp_path = os.path.join('/tmp', basename)
    if os.path.exists(tmp_path):
        shutil.copy2(tmp_path, filepath)
        os.remove(tmp_path)
        print(f"Recalculated: {filepath}")
    else:
        print(f"Error: LibreOffice did not produce output file")
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <path_to_xlsx>")
        sys.exit(1)
    recalc(sys.argv[1])
