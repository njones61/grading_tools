#!/usr/bin/env python3
"""De-identify student submissions before grading, and restore names afterward.

Implements the coded de-identification pattern FERPA contemplates at
34 CFR 99.31(b)(2): each student gets an opaque code that is *not* derived
from their name, NetID, or any other personal information, and the crosswalk
mapping codes back to students never leaves this machine.

    anonymize.py mask   <submissions_dir> [--out DIR] [--roster FILE] [--fresh]
    anonymize.py unmask <feedback_dir>    --roster FILE [--scores FILE]

`mask` copies submissions into a working directory with student-identifying
filenames replaced by codes, scrubs names out of file formats that store
text, and writes the crosswalk. Grade the masked copies.

`mask` is incremental. Run it again against a later download -- late work,
resubmissions -- and it reuses each returning student's existing code, hashes
every file against the previous round, and copies forward only what actually
needs grading. `ledger.json` records the per-file hashes and submission times
that make that comparison possible; unlike the crosswalk it holds no names,
so it is the file to read when you want to know what state a batch is in.
`--fresh` ignores an existing roster and starts the assignment over.

`unmask` renames the generated feedback documents back to real student
filenames and substitutes real names into scores.xlsx, ready for upload. It
is safe to run repeatedly: documents already restored in an earlier round are
recognized and left alone.

IMPORTANT LIMITS -- read the report `mask` prints:
  * Scanned PDFs and images cannot be scrubbed. A handwritten name or a
    photo of a signed page stays identifying no matter what the file is
    called. These are listed as UNSCRUBBABLE.
  * Scrubbing is best-effort text replacement. It catches a name typed into
    a spreadsheet cell or a notebook markdown cell; it will not catch a name
    embedded in a chart image or a screenshot pasted into a workbook.
De-identification reduces exposure. It does not by itself resolve whether
sending student work to a given service is permitted -- that is a question
for your institution's registrar or FERPA officer.
"""

import argparse
import hashlib
import json
import os
import re
import secrets
import shutil
import stat
import sys
import zipfile
from datetime import datetime
from pathlib import Path

# last_first_netid_whatever.ext -- the Learning Suite download convention.
SUBMISSION_RE = re.compile(r"^(?P<last>[^_]+)_(?P<first>[^_]+)_(?P<netid>[^_]+)_(?P<rest>.+)$")

# Formats that are zipped XML and can have text rewritten in place.
OOXML = {".xlsx", ".xlsm", ".docx", ".pptx"}
# Formats that are JSON text.
JSON_LIKE = {".ipynb"}
# Formats whose content we cannot inspect or scrub.
OPAQUE = {".pdf", ".png", ".jpg", ".jpeg", ".heic", ".zip", ".vsd", ".vsdx"}

SKIP_NAMES = {"Icon\r", "Icon", ".DS_Store", "Thumbs.db"}


def skip(p):
    return (p.name in SKIP_NAMES or p.name.startswith("~$")
            or p.name.startswith("._") or p.name.startswith(".~lock"))


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def when(ts):
    """A file timestamp as a readable local string."""
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")


def load_ledger(path):
    if path.exists():
        try:
            return json.loads(path.read_text())
        except Exception:
            pass
    return {"batches": [], "students": {}}


def new_code(used):
    """Opaque, random, not derived from any student attribute."""
    while True:
        code = "S-" + secrets.token_hex(3).upper()
        if code not in used:
            used.add(code)
            return code


def name_variants(last, first, netid):
    """Strings worth scrubbing from file contents, longest first."""
    out = {
        f"{first} {last}", f"{last}, {first}", f"{first}{last}",
        f"{last} {first}", last, first, netid,
    }
    return sorted({v for v in out if len(v) >= 3}, key=len, reverse=True)


def compile_variants(variants):
    """Word-boundary patterns, longest first.

    Bare substring matching is not safe here: 'Nora' occurs inside the OOXML
    attribute name 'mc:Ignorable', and replacing it produces invalid XML that
    Excel refuses to open. Anchor every variant to word boundaries.
    """
    out = []
    for v in variants:
        lead = r"\b" if v[0].isalnum() else ""
        tail = r"\b" if v[-1].isalnum() else ""
        out.append(re.compile(lead + re.escape(v) + tail, re.IGNORECASE))
    return out


def scrub_text(text, patterns, code):
    hits = 0
    for pat in patterns:
        text, n = pat.subn(code, text)
        hits += n
    return text, hits


# XML text nodes only -- everything between '>' and the next '<'. Never
# touches tag names or attribute values, which is where the corruption risk is.
TEXT_NODE = re.compile(r">([^<>]+)<")


def scrub_xml_part(xml, patterns, code):
    hits = 0

    def repl(m):
        nonlocal hits
        scrubbed, n = scrub_text(m.group(1), patterns, code)
        hits += n
        return f">{scrubbed}<"

    return TEXT_NODE.sub(repl, xml), hits


def scrub_ooxml(path, variants, code):
    """Rewrite an OOXML package, scrubbing identifying strings from text nodes.

    Covers cell text (sharedStrings.xml), document body text, and the
    creator / lastModifiedBy fields in docProps -- all of which are text
    nodes. Markup is left untouched.
    """
    patterns = compile_variants(variants)
    tmp = path.with_suffix(path.suffix + ".tmp")
    hits = 0
    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith(".xml"):
                    try:
                        text = data.decode("utf-8")
                        text, n = scrub_xml_part(text, patterns, code)
                        hits += n
                        data = text.encode("utf-8")
                    except UnicodeDecodeError:
                        pass
                zout.writestr(item, data)
        tmp.replace(path)
        return hits, None
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return 0, str(exc)


CODE_RE = re.compile(r"S-[0-9A-F]{6}")


def restore_ooxml_text(path, by_code):
    """Replace masked codes with real names inside an OOXML package's text.

    unmask renames feedback files, but a document that shows "Student: S-7F3A2B"
    in its body still reaches the student that way. This rewrites text nodes
    only -- the same narrow surface scrub_ooxml touches -- so formatting,
    formulas, and cached values all survive untouched.
    """
    tmp = path.with_suffix(path.suffix + ".tmp")
    hits = 0

    def repl(m):
        nonlocal hits
        def sub_code(cm):
            nonlocal hits
            st = by_code.get(cm.group(0))
            if not st:
                return cm.group(0)
            hits += 1
            return f"{st['first']} {st['last']} ({st['netid']})"
        return ">" + CODE_RE.sub(sub_code, m.group(1)) + "<"

    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith(".xml"):
                    try:
                        data = TEXT_NODE.sub(repl, data.decode("utf-8")).encode("utf-8")
                    except UnicodeDecodeError:
                        pass
                zout.writestr(item, data)
        tmp.replace(path)
        return hits, None
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return 0, str(exc)


def header_columns(ws):
    """Locate the First Name / Last Name / Net ID columns in a scores sheet.

    Returns (header_row, {field: column_index}) or (None, {}) if the sheet
    has no recognizable header row.
    """
    wanted = {"firstname": "first", "lastname": "last", "netid": "netid"}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
        found = {}
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            key = re.sub(r"[^a-z]", "", cell.value.lower())
            if key in wanted:
                found[wanted[key]] = cell.column
        if "netid" in found:
            return row[0].row, found
    return None, {}


def scrub_json(path, variants, code):
    """Scrub a notebook by walking parsed JSON, touching values but never keys."""
    patterns = compile_variants(variants)
    hits = 0

    def walk(node):
        nonlocal hits
        if isinstance(node, str):
            out, n = scrub_text(node, patterns, code)
            hits += n
            return out
        if isinstance(node, list):
            return [walk(x) for x in node]
        if isinstance(node, dict):
            return {k: walk(v) for k, v in node.items()}
        return node

    try:
        doc = json.loads(path.read_text(encoding="utf-8"))
        path.write_text(json.dumps(walk(doc), indent=1), encoding="utf-8")
        return hits, None
    except Exception as exc:
        return 0, str(exc)


def scrub_copy(dest, variants, code, do_redact, band, errors, unscrubbable, redacted):
    """Scrub or redact one masked copy in place, recording what happened."""
    ext = dest.suffix.lower()
    if ext in OOXML:
        _, err = scrub_ooxml(dest, variants, code)
        if err:
            errors.append((dest.name, err))
    elif ext in JSON_LIKE:
        _, err = scrub_json(dest, variants, code)
        if err:
            errors.append((dest.name, err))
    elif ext in OPAQUE:
        # A scan carries the name as pixels. Black out the header band
        # locally -- before anything is read, and without sending the page
        # anywhere to work out where the name is.
        status, detail = ("skipped", "redaction disabled")
        if do_redact:
            try:
                from redact_scan import redact
                status, detail = redact(dest, band_pct=band)
            except ImportError:
                status, detail = ("skipped", "redact_scan.py not importable")
            except Exception as exc:
                status, detail = ("error", str(exc))
        if status == "redacted":
            redacted.append((dest.name, detail))
        elif status == "error":
            errors.append((dest.name, f"redaction failed: {detail}"))
            unscrubbable.append(dest.name)
        else:
            unscrubbable.append(dest.name)
    else:
        unscrubbable.append(dest.name)


# A file needing to be graded, versus one already accounted for last round.
NEEDS_GRADING = {"new", "resubmitted", "added"}


def mask(args):
    src = args.submissions_dir.expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"error: no such directory: {src}")

    out = (args.out or src.parent / "masked").expanduser().resolve()
    roster_path = (args.roster or src.parent / "roster.json").expanduser().resolve()
    ledger_path = (args.ledger or src.parent / "ledger.json").expanduser().resolve()

    # An existing crosswalk means this assignment has been masked before, so
    # this is a later batch: returning students keep the codes their round-one
    # feedback and scores already carry.
    prior, ledger = {}, {"batches": [], "students": {}}
    if roster_path.exists() and not args.fresh:
        try:
            doc = json.loads(roster_path.read_text())
            prior = {s["netid"].lower(): s for s in doc.get("students", [])}
        except Exception as exc:
            sys.exit(f"error: could not read {roster_path}: {exc}")
        ledger = load_ledger(ledger_path)
    incremental = bool(prior)

    if not incremental and out.exists() and any(out.iterdir()):
        sys.exit(f"error: {out} exists and is not empty -- remove it, pass --out, "
                 f"or run against a roster to add a later batch")
    out.mkdir(parents=True, exist_ok=True)

    files = [f for f in sorted(src.iterdir()) if f.is_file() and not skip(f)]
    if not files:
        sys.exit(f"error: no submissions found in {src}")

    # code -> {masked filename: record} from every earlier round.
    prior_files = {code: {rec["masked"]: rec for rec in st.get("files", [])}
                   for code, st in ledger.get("students", {}).items()}

    students = dict(prior)
    used = {s["code"] for s in prior.values()}
    unparsed, unscrubbable, redacted, errors = [], [], [], []
    seen = {}          # code -> set of masked names in this batch
    batch = {}         # code -> list of (file, dest_name, sha, mtime, status)

    for f in files:
        m = SUBMISSION_RE.match(f.stem)
        if not m:
            unparsed.append(f.name)
            continue
        last, first, netid = m["last"], m["first"], m["netid"]
        key = netid.lower()
        is_new_student = key not in students
        if is_new_student:
            students[key] = {
                "code": new_code(used), "last": last, "first": first,
                "netid": netid, "files": [],
            }
        st = students[key]
        code = st["code"]
        dest_name = f"{code}_{m['rest']}{f.suffix}"
        sha = sha256_of(f)
        pf = prior_files.get(code, {})
        seen.setdefault(code, set()).add(dest_name)

        if is_new_student:
            status = "new"
        elif dest_name in pf:
            status = "unchanged" if pf[dest_name].get("sha256") == sha else "resubmitted"
        elif any(rec.get("sha256") == sha for rec in pf.values()):
            status = "renamed"
        elif not pf and (out / dest_name).exists():
            # Masked in an earlier round that predates the ledger: the masked
            # copy is the only evidence, and it cannot be hashed against the
            # source because scrubbing changed it. Take its presence as proof.
            status = "unchanged"
        else:
            status = "added"

        batch.setdefault(code, []).append(
            {"file": f, "masked": dest_name, "sha256": sha,
             "submitted": f.stat().st_mtime, "status": status})

    # A student is regraded whole. If any one file is new or changed, every
    # file they submitted goes back to the grader, so partial credit is judged
    # against the complete submission rather than the fragment that changed.
    regrade = {code for code, recs in batch.items()
               if any(r["status"] in NEEDS_GRADING for r in recs)}

    renames, to_grade = [], {}
    for code, recs in batch.items():
        st = next(s for s in students.values() if s["code"] == code)
        variants = name_variants(st["last"], st["first"], st["netid"])
        for r in recs:
            if r["status"] == "renamed":
                # Same bytes under a new name: refresh the masked copy's name
                # so the folder matches the ledger, but nothing needs regrading.
                old = next(name for name, rec in prior_files.get(code, {}).items()
                           if rec.get("sha256") == r["sha256"])
                renames.append((old, r["masked"]))
                (out / old).unlink(missing_ok=True)
            elif code not in regrade:
                continue
            shutil.copy2(r["file"], out / r["masked"])
            scrub_copy(out / r["masked"], variants, code, not args.no_redact,
                       args.band, errors, unscrubbable, redacted)
        # Keep the crosswalk's file list a union across rounds -- unmask reads it.
        known = {e["masked"] for e in st["files"]}
        for r in recs:
            if r["masked"] not in known:
                st["files"].append({"original": r["file"].name, "masked": r["masked"]})
        if code in regrade:
            to_grade[code] = [r["masked"] for r in recs]

    # Files this batch did not include, but earlier rounds did. A file the
    # student renamed is still here under its new name, so it is not missing.
    renamed_away = {old for old, _ in renames}
    missing = [(code, name) for code, pf in prior_files.items()
               for name in pf
               if name not in seen.get(code, set()) and name not in renamed_away]

    roster = {
        "source": str(src),
        "masked": str(out),
        "students": list(students.values()),
    }
    roster_path.write_text(json.dumps(roster, indent=2))
    os.chmod(roster_path, stat.S_IRUSR | stat.S_IWUSR)  # 600 -- owner only

    # The ledger carries hashes, timestamps and status but no names, so it is
    # safe to read while grading. The crosswalk stays closed.
    if src.name not in ledger["batches"]:
        ledger["batches"].append(src.name)
    # Folders pair up by number: submissions2 is graded into feedback2 and
    # uploaded as scores_upload2.csv. One number, the same meaning everywhere.
    rnd = ledger["batches"].index(src.name) + 1
    fb_dir = f"feedback{rnd}"
    ledger["updated"] = datetime.now().isoformat(timespec="seconds")
    for code, recs in batch.items():
        entry = ledger["students"].setdefault(code, {"files": [], "status": "needs_grading"})
        by_name = {e["masked"]: e for e in entry["files"]}
        for name, _ in renames:
            by_name.pop(name, None)
        for r in recs:
            by_name[r["masked"]] = {
                "masked": r["masked"], "sha256": r["sha256"],
                "submitted": datetime.fromtimestamp(r["submitted"]).isoformat(timespec="seconds"),
                "batch": src.name,
            }
        entry["files"] = [by_name[k] for k in sorted(by_name)]
        if code in regrade:
            entry["status"] = "needs_grading"
            entry["queued_by"] = src.name
            entry["round"] = rnd
    ledger_path.write_text(json.dumps(ledger, indent=2))

    todo = out / "TO_GRADE.txt"
    if to_grade:
        todo.write_text("\n".join(
            name for code in sorted(to_grade) for name in sorted(to_grade[code])) + "\n")
    else:
        todo.write_text("")

    # ---- report ----
    print(f"\nRound {rnd}: masked {sum(len(v) for v in batch.values())} file(s) for "
          f"{len(batch)} student(s) from {src.name}/")
    print(f"  masked copies : {out}")
    print(f"  crosswalk     : {roster_path}  (mode 600)")
    print(f"  ledger        : {ledger_path}\n")

    def show(header, rows, note=None):
        if not rows:
            return
        print(f"  {header} ({len(rows)}):")
        for r in rows:
            print(f"      {r['masked']}   submitted {when(r['submitted'])}")
        if note:
            print(f"  {note}")
        print()

    def having(status, regraded=None):
        return [r for code, recs in sorted(batch.items()) for r in recs
                if r["status"] == status
                and (regraded is None or (code in regrade) == regraded)]

    show("NEW -- first submission from this student", having("new"))
    show("RESUBMITTED -- content changed since the last round", having("resubmitted"),
         "Their existing feedback document and score are now stale and are being replaced.")
    show("ADDED -- new file alongside one already graded", having("added"))
    show("UNCHANGED, REGRADED ANYWAY -- this student submitted something new, so their "
         "whole\n  submission goes back to the grader", having("unchanged", regraded=True))
    show("UNCHANGED -- already graded, not copied forward",
         having("unchanged", regraded=False))

    if renames:
        print(f"  RENAMED ({len(renames)}) -- same content, new filename:")
        for old, new in renames:
            print(f"      {old}  ->  {new}")
        print()
    if missing:
        print(f"  MISSING ({len(missing)}) -- graded earlier, absent from this download.")
        print("  Nothing was deleted; their feedback and scores stand:")
        for code, name in sorted(missing):
            print(f"      {name}")
        print()
    if unparsed:
        print(f"  COULD NOT PARSE ({len(unparsed)}) -- not copied, handle by hand:")
        for n in unparsed:
            print(f"      {n}")
        print()
    if redacted:
        print(f"  REDACTED ({len(redacted)}) -- header band blacked out, OCR text layer removed:")
        for n, d in redacted:
            print(f"      {n}  ({d})")
        print("  The band is geometric, not a name detector. Look at the previews in")
        print(f"  {out / 'redaction_previews'} before grading -- a name written down a")
        print("  margin or on a later page is still there.\n")
    if unscrubbable:
        print(f"  UNSCRUBBABLE ({len(unscrubbable)}) -- filename is masked, but content")
        print(f"  may still identify the student (handwriting, scans, embedded images):")
        for n in unscrubbable:
            print(f"      {n}")
        print()
    if errors:
        print(f"  ERRORS ({len(errors)}):")
        for n, e in errors:
            print(f"      {n}: {e}")
        print()

    if to_grade:
        print(f"  GRADE THESE {sum(len(v) for v in to_grade.values())} file(s) "
              f"for {len(to_grade)} student(s) -- listed in {todo}")
        print(f"  Write their feedback documents into {src.parent / fb_dir}/ -- one folder")
        print(f"  per round, so the user reads only what this round produced.")
    else:
        print("  Nothing to grade: every file in this batch was graded in an earlier round.")
    print("  Grade the masked copies. Do not commit the crosswalk to git.\n")
    return 1 if (unparsed or errors) else 0


def unmask(args):
    roster_path = args.roster.expanduser().resolve()
    if not roster_path.exists():
        sys.exit(f"error: no crosswalk at {roster_path}")
    roster = json.loads(roster_path.read_text())

    by_code = {s["code"]: s for s in roster["students"]}
    fb = args.feedback_dir.expanduser().resolve()
    if not fb.is_dir():
        sys.exit(f"error: no such directory: {fb}")

    # Late batches mean unmask runs more than once, so the feedback folder is a
    # mix: documents restored last round already carry real names. Recognize
    # them by their NetID rather than reporting them as failures.
    known_netids = {s["netid"].lower() for s in roster["students"]}

    def already_restored(name):
        m = SUBMISSION_RE.match(Path(name).stem)
        return bool(m) and m["netid"].lower() in known_netids

    renamed, orphans, restored_bodies, skipped = 0, [], 0, 0
    for f in sorted(fb.iterdir()):
        if not f.is_file() or skip(f):
            continue
        m = re.match(r"^(S-[0-9A-F]{6})_(.*)$", f.name)
        if not m:
            if already_restored(f.name):
                skipped += 1
            else:
                orphans.append(f.name)
            continue
        code, rest = m.group(1), m.group(2)
        st = by_code.get(code)
        if not st:
            orphans.append(f.name)
            continue
        new_name = f"{st['last']}_{st['first']}_{st['netid']}_{rest}"
        f = f.rename(f.with_name(new_name))
        renamed += 1
        # The code is in the document body too, not just the filename.
        if f.suffix.lower() in OOXML:
            hits, err = restore_ooxml_text(f, by_code)
            if err:
                print(f"warning: could not rewrite text in {f.name}: {err}")
            elif hits:
                restored_bodies += 1

    print(f"\nRestored {renamed} feedback file(s) in {fb}")
    if restored_bodies:
        print(f"  and replaced the code inside {restored_bodies} document body/bodies")
    if skipped:
        print(f"  ({skipped} already carried real names from an earlier round -- left alone)")

    if args.scores:
        scores = args.scores.expanduser().resolve()
        try:
            import openpyxl
            wb = openpyxl.load_workbook(scores)
            replaced, named = 0, 0
            for ws in wb.worksheets:
                _, cols = header_columns(ws)
                for row in ws.iter_rows():
                    for cell in row:
                        if not (isinstance(cell.value, str) and cell.value.strip() in by_code):
                            continue
                        st = by_code[cell.value.strip()]
                        in_netid_col = cols.get("netid") == cell.column
                        cell.value = st["netid"]
                        replaced += 1
                        # Only fill names alongside a code that sat in the Net ID
                        # column -- a code elsewhere is not a student row.
                        if in_netid_col:
                            for field in ("first", "last"):
                                col = cols.get(field)
                                if col:
                                    ws.cell(row=cell.row, column=col, value=st[field])
                                    named += 1
            wb.save(scores)
            msg = f"Replaced {replaced} code(s) with NetIDs in {scores.name}"
            if named:
                msg += f", and filled {named} name cell(s)"
            print(msg)

            # openpyxl writes formulas without cached values, so the totals and
            # statistics in scores.xlsx would read as blank until something
            # recalculates them. Do it here rather than leaving a silent trap.
            try:
                from recalc import try_recalc
                ok, err = try_recalc(scores)
                print("Recalculated formulas in " + scores.name if ok
                      else f"warning: {err}\n         run recalc.py on {scores.name} "
                           f"before uploading, or the totals will read as blank")
            except ImportError:
                print(f"warning: could not import recalc.py -- run it on {scores.name} "
                      "before uploading, or the totals will read as blank")
        except ImportError:
            print("warning: openpyxl not installed -- scores.xlsx left unmodified")
        except Exception as exc:
            print(f"warning: could not rewrite {scores.name}: {exc}")

    # The upload artifacts -- scores_upload.csv and batch_upload.zip -- are not
    # built here. Every feedback document gets read and often edited before it
    # goes out, and an artifact built now would be a snapshot of the draft.
    # sync_scores.py builds both from what you actually approved.
    print("\nNext: review the feedback documents, then build the upload artifacts from")
    print("      the reviewed versions:")
    print(f"      python sync_scores.py {fb.parent} --apply --zip")

    if orphans:
        print(f"\n  UNMATCHED ({len(orphans)}) -- left as-is:")
        for n in orphans:
            print(f"      {n}")
    print()
    return 1 if orphans else 0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    m = sub.add_parser("mask", help="de-identify submissions before grading")
    m.add_argument("submissions_dir", type=Path)
    m.add_argument("--out", type=Path, help="where masked copies go (default: ../masked)")
    m.add_argument("--roster", type=Path, help="crosswalk path (default: ../roster.json)")
    m.add_argument("--ledger", type=Path, help="grading state (default: ../ledger.json)")
    m.add_argument("--fresh", action="store_true",
                   help="ignore an existing roster and start the assignment over")
    m.add_argument("--band", type=float, default=6.0,
                   help="percent of page height to black out on scans (default: 6)")
    m.add_argument("--no-redact", action="store_true",
                   help="copy scans through untouched instead of redacting the header band")
    m.set_defaults(func=mask)

    u = sub.add_parser("unmask", help="restore real names to generated feedback")
    u.add_argument("feedback_dir", type=Path)
    u.add_argument("--roster", type=Path, required=True)
    u.add_argument("--scores", type=Path, help="scores.xlsx to rewrite in place")
    u.set_defaults(func=unmask)

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
