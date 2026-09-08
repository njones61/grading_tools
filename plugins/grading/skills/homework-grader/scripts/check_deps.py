#!/usr/bin/env python3
"""Verify every tool the grading pipeline shells out to is actually installed.

    check_deps.py [--workspace DIR] [--json]

Run before anything else. Each of these is discovered late otherwise, and
some of them fail in ways that look like success:

  * Without Pillow, pypdf or poppler, redaction is skipped and the scan is
    copied anyway. The filename is masked, the student's handwritten name is
    still on the page, and the run reports it in the same list that a healthy
    run fills with born-digital PDFs -- so it reads as normal.
  * Without LibreOffice, scores.xlsx keeps its formulas and none of the cached
    values, so every Total reads as blank and the upload CSV comes out empty.
    That surfaces at Phase 5, after the whole class has been graded.
  * Without the docx npm package, no feedback document can be written at all.
    That surfaces at Phase 4, after the whole class has been graded.

The cost of checking is a second. The cost of not checking is a batch graded
twice, or a name reaching the grader that was supposed to have been removed.

Exit codes:
    0  everything present
    1  something missing
"""

import argparse
import importlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

# (import name, pip name, what it does, which phase notices it missing)
PYTHON_DEPS = [
    ("openpyxl", "openpyxl", "reads and writes every .xlsx in the pipeline", "2"),
    ("docx", "python-docx", "reads feedback documents to validate and reconcile them", "4"),
    ("PIL", "pillow", "paints the redaction band onto a scanned page", "2.5"),
    ("pypdf", "pypdf", "measures PDF pages, so a scan can be told from a born-digital file", "2.5"),
]

BINARY_DEPS = [
    (["pdftoppm"], "brew install poppler",
     "rasterizes PDF pages for redaction", "2.5"),
    (["pdfimages"], "brew install poppler",
     "detects whether a PDF is a scan", "2.5"),
]


def check(workspace=None):
    """Everything missing, as [{name, install, why, phase}]."""
    missing = []

    for mod, pip_name, why, phase in PYTHON_DEPS:
        try:
            importlib.import_module(mod)
        except ImportError:
            missing.append({"name": pip_name, "install": f"pip install {pip_name}",
                            "why": why, "phase": phase})

    for names, install, why, phase in BINARY_DEPS:
        if not any(shutil.which(n) for n in names):
            missing.append({"name": names[0], "install": install,
                            "why": why, "phase": phase})

    # LibreOffice is rarely on PATH -- on macOS it lives inside the app bundle.
    # recalc.py already knows every place it hides, so ask it rather than
    # guessing here and reporting a false absence.
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from recalc import find_libreoffice
        if find_libreoffice() is None:
            missing.append({
                "name": "LibreOffice", "install": "brew install --cask libreoffice",
                "why": "recalculates scores.xlsx, without which every Total is blank",
                "phase": "5"})
    except ImportError:
        missing.append({"name": "recalc.py", "install": "reinstall the plugin",
                        "why": "could not be imported to locate LibreOffice",
                        "phase": "5"})

    # node resolves node_modules by walking up from the working directory, so
    # this has to be asked from inside the workspace where docx was installed.
    if workspace:
        ws = Path(workspace).expanduser()
        if not shutil.which("node"):
            missing.append({"name": "node", "install": "brew install node",
                            "why": "runs the feedback document generator", "phase": "4"})
        elif ws.is_dir():
            try:
                proc = subprocess.run(
                    ["node", "-e", "require.resolve('docx')"],
                    cwd=str(ws), capture_output=True, text=True, timeout=30)
                if proc.returncode != 0:
                    missing.append({
                        "name": "docx (npm)",
                        "install": f"cd {ws} && npm install docx",
                        "why": "builds every feedback .docx", "phase": "4"})
            except Exception as exc:
                missing.append({"name": "docx (npm)",
                                "install": f"cd {ws} && npm install docx",
                                "why": f"could not be checked: {exc}", "phase": "4"})
    return missing


def report(missing):
    if not missing:
        return "  dependencies   all present"
    lines = [f"  MISSING ({len(missing)}) -- install these before grading:", ""]
    for m in missing:
        lines.append(f"      {m['name']}  (needed in phase {m['phase']})")
        lines.append(f"          {m['why']}")
        lines.append(f"          {m['install']}")
        lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workspace", type=Path,
                    help="grading workspace, to check the docx npm package")
    ap.add_argument("--json", action="store_true", dest="as_json")
    args = ap.parse_args()

    missing = check(args.workspace)
    if args.as_json:
        print(json.dumps({"missing": missing, "ready": not missing}, indent=2))
    else:
        print(report(missing))
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
