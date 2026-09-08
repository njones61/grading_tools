#!/usr/bin/env python3
"""Reconcile scores.xlsx and the upload artifacts with the feedback documents.

    sync_scores.py <assignment_dir> [--apply] [--fix-totals] [--zip] [--round N]

The feedback document is what the student reads, so the number in it is the
number that has to reach the gradebook. This treats the feedback folders as
the source of truth and derives everything else from them: you edit a document
during review, run this, and scores.xlsx and scores_upload<N>.csv follow.
Nothing is transcribed by hand, so nothing drifts.

Without --apply it only reports. With --apply it writes the document scores
into scores.xlsx in place -- leaving the =SUM() formulas, the conditional
formatting and the statistics rows exactly as they are -- recalculates, and
writes the upload CSV. --zip also packages the reviewed documents for the
Learning Suite batch upload.

Feedback folders are numbered to match the downloads they came from:
submissions1 is graded into feedback1 and uploaded as scores_upload1.csv,
submissions2 into feedback2, and so on. scores.xlsx covers the whole class, so
reconciling it reads every round; the upload artifacts are per-round, and cover
only the students that round graded. Importing the whole class again on a late
batch would overwrite any score adjusted inside Learning Suite since. --round
picks a different round, and --round all exports everyone.

When a student appears in two rounds -- they resubmitted -- the later document
is the score, and the earlier one is reported as superseded.

Three numbers in each document can disagree after a hand edit: the per-section
"Score: 4 / 5" lines, the Earned column of the summary table, and the two
totals ("Total Score:" near the top and the table's TOTAL row). Editing one
and forgetting the others is the ordinary way this goes wrong, so every
document is checked against itself first. The summary table is authoritative
-- it is the part students read as the score -- and --fix-totals rewrites the
two totals from the table's own Earned column.

Runs against masked or restored feedback: a document is matched to a student
by its S-XXXXXX code, or by the NetID in its filename once unmask has run.
"""

import argparse
import re
import sys
from copy import copy
from pathlib import Path

CODE_RE = re.compile(r"^(S-[0-9A-F]{6})_")
# last_first_netid_... -- the filename unmask restores.
NETID_RE = re.compile(r"^[^_]+_[^_]+_([^_]+)_")

TOTAL_LABEL = re.compile(r"^\s*total\s*$", re.I)
SCORE_LINE = re.compile(r"Score:\s*(-?[\d.]+)\s*/\s*([\d.]+)")
TOTAL_LINE = re.compile(r"Total Score:\s*(-?[\d.]+)\s*/\s*([\d.]+)")

NON_STUDENT = re.compile(r"^(max\s*points?|average|median|std|count|total)\b", re.I)


def num(text):
    """Parse a cell's number, or None if it isn't one."""
    try:
        v = float(str(text).strip())
    except (TypeError, ValueError):
        return None
    return int(v) if v.is_integer() else v


def norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower()) if isinstance(s, str) else ""


def read_feedback(path):
    """Pull the rubric table, the section scores and both totals out of a .docx.

    Returns a dict, or raises ValueError if the document has no summary table
    in the expected three-column shape.
    """
    import docx

    doc = docx.Document(str(path))

    items, table_total, table_possible = [], None, None
    for table in doc.tables:
        head = [norm(c.text) for c in table.rows[0].cells]
        if head[:3] != ["rubricitem", "possible", "earned"]:
            continue
        for row in table.rows[1:]:
            label = row.cells[0].text.strip()
            possible, earned = num(row.cells[1].text), num(row.cells[2].text)
            if TOTAL_LABEL.match(label):
                table_total, table_possible = earned, possible
            elif earned is not None:
                items.append({"label": label, "possible": possible, "earned": earned})
        break
    if not items:
        raise ValueError("no 'Rubric Item / Possible / Earned' summary table")

    # The per-section score lines, in document order, pair up with the table
    # rows: the template emits one section per rubric item, same order.
    sections, stated_total = [], None
    for para in doc.paragraphs:
        text = para.text
        m = TOTAL_LINE.search(text)
        if m:
            stated_total = num(m.group(1))
            continue
        m = SCORE_LINE.search(text)
        if m:
            sections.append(num(m.group(1)))

    return {
        "path": path,
        "items": items,
        "table_total": table_total,
        "table_possible": table_possible,
        "stated_total": stated_total,
        "sections": sections,
        "sum_of_items": sum(i["earned"] for i in items),
    }


def check_internal(fb):
    """Ways one document can disagree with itself after a hand edit."""
    problems = []
    if fb["table_total"] is not None and fb["table_total"] != fb["sum_of_items"]:
        problems.append(f"table TOTAL says {fb['table_total']}, "
                        f"but the Earned column sums to {fb['sum_of_items']}")
    if fb["stated_total"] is not None and fb["stated_total"] != fb["sum_of_items"]:
        problems.append(f"'Total Score' says {fb['stated_total']}, "
                        f"but the Earned column sums to {fb['sum_of_items']}")
    if len(fb["sections"]) == len(fb["items"]):
        for sec, item in zip(fb["sections"], fb["items"]):
            if sec != item["earned"]:
                problems.append(f"section score {sec} disagrees with the table's "
                                f"{item['earned']} for '{item['label']}'")
    return problems


def fix_totals(fb):
    """Rewrite the two displayed totals from the table's Earned column.

    The template puts each number in a run of its own, so the replacement stays
    inside one run and no formatting is disturbed.
    """
    import docx

    target = fb["sum_of_items"]
    doc = docx.Document(str(fb["path"]))
    changed = 0

    for para in doc.paragraphs:
        if not TOTAL_LINE.search(para.text):
            continue
        for run in para.runs:
            m = re.fullmatch(r"\s*(-?[\d.]+)\s*/\s*([\d.]+)\s*", run.text)
            if m and num(m.group(1)) != target:
                run.text = f"{target} / {m.group(2)}"
                changed += 1

    for table in doc.tables:
        if [norm(c.text) for c in table.rows[0].cells][:3] != ["rubricitem", "possible", "earned"]:
            continue
        for row in table.rows[1:]:
            if not TOTAL_LABEL.match(row.cells[0].text.strip()):
                continue
            cell = row.cells[2]
            for para in cell.paragraphs:
                for run in para.runs:
                    if num(run.text) is not None and num(run.text) != target:
                        run.text = str(target)
                        changed += 1
        break

    if changed:
        doc.save(str(fb["path"]))
    return changed


def identify(name, rows):
    """The student a feedback filename names: its masked code, or its NetID.

    Independent of whether the workbook has a row for them -- a student graded
    in a later batch has a document before they have a row, and that gap is
    worth reporting rather than filing under 'unrecognized'.
    """
    m = CODE_RE.match(name)
    if m:
        return m.group(1)
    m = NETID_RE.match(name)
    if m:
        for key in rows:                        # prefer the workbook's spelling
            if key.lower() == m.group(1).lower():
                return key
        return m.group(1)
    return None


# feedback1, feedback2, ... The bare `feedback` is round 1 too, for workspaces
# created before rounds existed.
ROUND_DIR = re.compile(r"^feedback[_-]?(\d+)?$")


def round_dirs(root):
    """Every feedback folder for this assignment, oldest round first.

    Folders pair up by number: submissions1 is graded into feedback1 and
    uploaded as scores_upload1.csv, submissions2 into feedback2, and so on.
    One number, the same meaning everywhere, so a folder's name tells you which
    download it came from and which documents still need reading.
    """
    out = []
    for d in sorted(root.iterdir()):
        if not d.is_dir():
            continue
        m = ROUND_DIR.match(d.name)
        if m:
            out.append((int(m.group(1) or 1), d))
    return sorted(out)


def csv_name(rnd):
    return f"scores_upload{rnd}.csv"


ROUND_HEADER = "Round"


def ensure_round_column(ws, layout):
    """Find the Round column, adding it after Total if it isn't there yet.

    It goes at the end deliberately. sync pairs the workbook's rubric columns
    with the document's summary table rows by position, so a column inserted
    anywhere before Total would silently shift that pairing by one.
    """
    header_row = layout["header_row"]
    for col in range(1, ws.max_column + 2):
        if norm(ws.cell(header_row, col).value) == norm(ROUND_HEADER):
            return col

    col = (layout["total_col"] or ws.max_column) + 1
    src = ws.cell(header_row, layout["netid_col"])
    cell = ws.cell(header_row, col, ROUND_HEADER)
    cell._style = copy(src._style)
    ws.column_dimensions[cell.column_letter].width = 8
    return col


def load_scores(path, data_only=False):
    """Index scores.xlsx by identity. Returns (workbook, sheet, layout).

    data_only=True reads the cached values instead of the formulas, which is
    what anything analyzing the scores wants and what writing them back must
    never use -- saving a data_only workbook discards every formula in it.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=data_only)
    ws = wb.active

    header_row = netid_col = total_col = None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
        found = {}
        for cell in row:
            key = norm(cell.value)
            if key == "netid":
                found["netid"] = cell.column
            elif key == "total":
                found["total"] = cell.column
        if "netid" in found:
            header_row, netid_col = row[0].row, found["netid"]
            total_col = found.get("total")
            break
    if not header_row:
        raise ValueError(f"no 'Net ID' column in the first 10 rows of {path.name}")

    # Rubric columns are everything between Net ID and Total.
    rubric_cols = []
    if total_col:
        for col in range(netid_col + 1, total_col):
            label = ws.cell(header_row, col).value
            if label:
                rubric_cols.append((col, str(label).strip()))

    rows = {}
    for r in range(header_row + 1, ws.max_row + 1):
        key = ws.cell(r, netid_col).value
        if key is None or not str(key).strip():
            continue
        key = str(key).strip()
        if NON_STUDENT.match(key):
            continue
        rows[key] = r

    return wb, ws, {"header_row": header_row, "netid_col": netid_col,
                    "total_col": total_col, "rubric_cols": rubric_cols, "rows": rows}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("assignment_dir", type=Path,
                    help="the assignment folder holding feedback/ and scores.xlsx")
    ap.add_argument("--apply", action="store_true",
                    help="write the document scores into scores.xlsx and rebuild the CSV")
    ap.add_argument("--fix-totals", action="store_true",
                    help="rewrite each document's two totals from its own Earned column")
    ap.add_argument("--zip", action="store_true",
                    help="also package the reviewed documents for the batch upload")
    ap.add_argument("--force", action="store_true",
                    help="package even documents that contradict themselves")
    ap.add_argument("--round", default=None,
                    help="which round's upload artifacts to build: a number, or "
                         "'all' for the whole class (default: the latest round)")
    ap.add_argument("--scores", type=Path, help="scores.xlsx (default: <dir>/scores.xlsx)")
    ap.add_argument("--feedback", type=Path,
                    help="one feedback dir (feedback1, feedback2, ...) instead of "
                         "every round under <dir>")
    args = ap.parse_args()

    root = args.assignment_dir.expanduser().resolve()
    scores_path = (args.scores or root / "scores.xlsx").expanduser().resolve()
    if not scores_path.is_file():
        sys.exit(f"error: no scores workbook at {scores_path}")

    if args.feedback:
        d = args.feedback.expanduser().resolve()
        m = ROUND_DIR.match(d.name)
        rounds = [(int(m.group(1) or 1) if m else 1, d)]
    else:
        rounds = round_dirs(root)
    if not rounds or not all(d.is_dir() for _, d in rounds):
        sys.exit(f"error: no feedback directory under {root}")

    # scores.xlsx covers the whole class, so reconciling it means reading every
    # round. The upload artifacts are per-round, and built further down.
    docs = [(rnd, f) for rnd, d in rounds for f in sorted(d.iterdir())
            if f.suffix.lower() == ".docx" and not f.name.startswith("~$")]
    if not docs:
        sys.exit(f"error: no feedback documents under {root}")

    try:
        wb, ws, layout = load_scores(scores_path)
    except ImportError:
        sys.exit("error: openpyxl is not installed")
    except Exception as exc:
        sys.exit(f"error: {exc}")

    parsed, unreadable, unmatched, internal, seen = {}, [], [], [], {}
    superseded, doc_round = [], {}
    for rnd, f in docs:
        try:
            fb = read_feedback(f)
        except ImportError:
            sys.exit("error: python-docx is not installed")
        except Exception as exc:
            unreadable.append((f.name, str(exc)))
            continue
        key = identify(f.name, layout["rows"])
        if key is None:
            unmatched.append(f.name)
            continue
        if key in seen:
            prev_round, prev_name = doc_round[key], seen[key]
            if rnd > prev_round:
                # A resubmission: the later round regraded work the earlier one
                # had already covered, so the later document is the score.
                superseded.append((key, prev_name, prev_round,
                                   f"{f.parent.name}/{f.name}", rnd))
            elif rnd < prev_round:
                superseded.append((key, f"{f.parent.name}/{f.name}", rnd,
                                   prev_name, prev_round))
                continue
            else:
                internal.append((f.name, [f"a second document in the same round for "
                                          f"this student ({prev_name}) -- resolve by hand"]))
                continue
        seen[key] = f"{f.parent.name}/{f.name}"
        doc_round[key] = rnd
        problems = check_internal(fb)
        if args.fix_totals and any("Total" in p or "total" in p for p in problems):
            n = fix_totals(fb)
            if n:
                fb = read_feedback(f)
                problems = check_internal(fb)
        if problems:
            internal.append((f.name, problems))
        fb["round"] = rnd
        parsed[key] = fb

    # ---- compare against the workbook ----
    drift, missing_rows = [], []
    for key, fb in parsed.items():
        row = layout["rows"].get(key)
        if row is None:
            missing_rows.append(key)
            continue
        diffs = []
        for (col, label), item in zip(layout["rubric_cols"], fb["items"]):
            have = ws.cell(row, col).value
            if num(have) != item["earned"]:
                diffs.append((label, have, item["earned"], col))
        if len(layout["rubric_cols"]) != len(fb["items"]):
            diffs.append((f"column count: workbook has {len(layout['rubric_cols'])} "
                          f"rubric columns, the document has {len(fb['items'])} items",
                          None, None, None))
        if diffs:
            drift.append((key, seen[key], diffs))

    orphan_rows = sorted(set(layout["rows"]) - set(parsed))

    # ---- which round gets the upload artifacts ----
    latest = max(doc_round.values()) if doc_round else 1
    if args.round in (None, "", "latest"):
        target = latest
    elif str(args.round).lower() == "all":
        target = None
    else:
        try:
            target = int(args.round)
        except ValueError:
            sys.exit(f"error: --round takes a number or 'all', not {args.round!r}")
        if target not in dict((r, d) for r, d in rounds):
            sys.exit(f"error: no feedback folder for round {target}")
    by_round = {r: d for r, d in rounds}

    # ---- report ----
    where = ", ".join(d.name for _, d in rounds)
    print(f"\n{len(parsed)} feedback document(s) read from {where}")
    print(f"compared against {scores_path.name}\n")

    if superseded:
        print(f"  SUPERSEDED ({len(superseded)}) -- regraded in a later round, so the")
        print("  earlier document is stale and its score is not used:")
        for key, old, old_r, new, new_r in superseded:
            print(f"      {key}: {old}  ->  {new}")
        print()

    if unreadable:
        print(f"  UNREADABLE ({len(unreadable)}):")
        for n, why in unreadable:
            print(f"      {n}: {why}")
        print()
    if unmatched:
        print(f"  UNRECOGNIZED FILENAME ({len(unmatched)}) -- neither an S-XXXXXX code nor")
        print("  the last_first_netid_... form unmask produces:")
        for n in unmatched:
            print(f"      {n}")
        print()
    if internal:
        print(f"  DOCUMENT DISAGREES WITH ITSELF ({len(internal)}):")
        for n, problems in internal:
            print(f"      {n}")
            for p in problems:
                print(f"          {p}")
        if not args.fix_totals and any("otal" in p for _, ps in internal for p in ps):
            print("  --fix-totals rewrites the two displayed totals from the Earned column.")
        if any("section score" in p for _, ps in internal for p in ps):
            print("  Section score lines are left alone: the feedback bullets underneath one")
            print("  explain the deduction it shows, so rewriting the number without the prose")
            print("  would leave the document contradicting its own explanation. Fix by hand.")
        print()
    if missing_rows:
        print(f"  NOT IN THE WORKBOOK ({len(missing_rows)}) -- graded, but no row yet.")
        print("  These are usually a later batch; add their rows to scores.xlsx:")
        for k in missing_rows:
            print(f"      {k}  ({seen[k]})")
        print()
    if orphan_rows:
        print(f"  NO FEEDBACK DOCUMENT ({len(orphan_rows)}) -- a row in the workbook with")
        print("  nothing in feedback/ to back it:")
        for k in orphan_rows:
            print(f"      {k}")
        print()

    if drift:
        print(f"  OUT OF SYNC ({len(drift)}) -- the document and the workbook differ:")
        for key, name, diffs in drift:
            print(f"      {key}  ({name})")
            for label, have, want, _ in diffs:
                if want is None:
                    print(f"          {label}")
                else:
                    print(f"          {label}: workbook {have} -> document {want}")
        print()
    else:
        print("  Every matched row already agrees with its feedback document.\n")

    if not args.apply:
        if drift or missing_rows:
            print("  Nothing was written. Re-run with --apply to bring the workbook into line.\n")
        return 1 if (unreadable or unmatched or internal or missing_rows) else 0

    # ---- apply ----
    if missing_rows:
        print("  REFUSING TO WRITE: some students have no row in the workbook. Adding rows")
        print("  means moving the statistics block and re-ranging the =SUM() and conditional")
        print("  formatting, which is the grader's job, not a text edit. Add the rows for the")
        print("  students listed above, then run this again.\n")
        return 1

    written = 0
    for key, name, diffs in drift:
        row = layout["rows"][key]
        for label, have, want, col in diffs:
            if col is None:
                continue
            ws.cell(row, col, value=want)
            written += 1

    # Which round each student's score came from, so one workbook still shows
    # who was graded when.
    round_col = ensure_round_column(ws, layout)
    stamped = 0
    for key, rnd in doc_round.items():
        row = layout["rows"].get(key)
        if row and num(ws.cell(row, round_col).value) != rnd:
            ws.cell(row, round_col, value=rnd)
            stamped += 1
    if stamped:
        written += stamped
        print(f"  Stamped the round on {stamped} student row(s)")

    if written:
        wb.save(scores_path)
        print(f"  Wrote {written} cell(s) into {scores_path.name}")
    else:
        print(f"  {scores_path.name} needed no changes")

    # The Total column is a formula, so it has no cached value until something
    # recalculates. export_upload_csv reads exactly that cached value.
    try:
        from recalc import try_recalc
        ok, err = try_recalc(scores_path)
        print(f"  Recalculated {scores_path.name}" if ok else f"  warning: {err}")
    except ImportError:
        print(f"  warning: could not import recalc.py -- run it on {scores_path.name}")

    # One CSV per round, holding only the students that round graded. Importing
    # the whole class again would overwrite any score adjusted inside Learning
    # Suite since the last upload.
    only = None if target is None else [k for k, r in doc_round.items() if r == target]
    out_csv = scores_path.parent / (csv_name(target) if target else "scores_upload_all.csv")
    try:
        from export_upload_csv import export
        csv_path, n, skipped = export(scores_path, out=out_csv, only=only)
        scope = "the whole class" if target is None else f"round {target}"
        print(f"  Wrote {csv_path.name} -- {n} row(s) for {scope}")
        for netid, why in skipped:
            print(f"      skipped {netid}: {why}")
    except Exception as exc:
        print(f"  warning: could not write the upload CSV: {exc}")

    if args.zip and internal and not args.force:
        # The zip is what reaches students. A document showing two different
        # scores for the same rubric item should not be in it.
        print(f"  NOT PACKAGING: {len(internal)} document(s) contradict themselves and would")
        print("  go to students that way. Fix them, or pass --force to package anyway.")
    elif args.zip:
        for rnd in ([target] if target else sorted(by_round)):
            try:
                from package_feedback import package
                zip_path, n = package(by_round[rnd])
                print(f"  Packaged {n} document(s) -> {by_round[rnd].name}/{zip_path.name} "
                      f"({zip_path.stat().st_size // 1024} KB) for the batch upload")
            except Exception as exc:
                print(f"  warning: could not package {by_round[rnd].name}: {exc}")

    print()
    return 1 if (unreadable or unmatched or internal) else 0


if __name__ == "__main__":
    sys.exit(main())
