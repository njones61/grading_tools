#!/usr/bin/env python3
"""Bundle a folder of feedback documents into the zip Learning Suite ingests.

Usage:
    package_feedback.py <feedback_dir> [--name batch_upload.zip]

Learning Suite accepts one zip of per-student feedback files rather than
uploading them one at a time. The zip is written *into* the feedback folder,
so it excludes itself and any earlier copy -- otherwise each run would nest
the previous upload inside the next one.

anonymize.py unmask calls this automatically once names are restored. Run it
by hand if you graded without de-identifying, or after editing feedback.
"""

import argparse
import sys
import zipfile
from pathlib import Path

DEFAULT_NAME = "batch_upload.zip"

# Editor and OS droppings that must never reach the LMS.
SKIP_NAMES = {"Icon\r", "Icon", ".DS_Store", "Thumbs.db", "desktop.ini"}


def _skip(p):
    return (p.name in SKIP_NAMES or p.name.startswith("~$")
            or p.name.startswith("._") or p.name.startswith(".~lock"))


def package(feedback_dir, name=DEFAULT_NAME):
    """Zip every feedback file in feedback_dir. Returns (zip_path, file_count)."""
    fb = Path(feedback_dir).expanduser().resolve()
    if not fb.is_dir():
        raise NotADirectoryError(f"no such directory: {fb}")

    out = fb / name
    # Collect before opening the zip: the file being written must not be an input.
    files = sorted(f for f in fb.iterdir()
                   if f.is_file() and not _skip(f) and f.suffix.lower() != ".zip")
    if not files:
        raise ValueError(f"no feedback files to package in {fb}")

    tmp = out.with_suffix(".zip.tmp")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for f in files:
                z.write(f, arcname=f.name)   # flat -- no directory entries
        tmp.replace(out)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    return out, len(files)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("feedback_dir", type=Path)
    ap.add_argument("--name", default=DEFAULT_NAME,
                    help=f"zip filename (default: {DEFAULT_NAME})")
    args = ap.parse_args()
    try:
        out, n = package(args.feedback_dir, args.name)
    except Exception as exc:
        sys.exit(f"error: {exc}")
    print(f"Packaged {n} feedback file(s) -> {out}  ({out.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
