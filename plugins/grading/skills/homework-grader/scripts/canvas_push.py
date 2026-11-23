#!/usr/bin/env python3
"""Send a graded round's scores and feedback documents to Canvas.

    canvas_push.py <assignment_dir> [--round N|all] [--send]
                   [--overwrite NETID ...] [--comment TEXT]

The Canvas counterpart of the Learning Suite upload artifacts. For each
student the round graded, it enters their score on the Canvas submission and
attaches their feedback .docx as a submission comment.

Without --send it is a dry run: it lists what it would send, and to whom, and
sends nothing. Read that list before sending.

What it sends, and what it leaves alone:

  * The raw score, from the feedback document's summary table -- the score of
    record. Canvas applies the course's late policy to a late submission
    itself, so nothing is deducted here.
  * Grades are entered, never posted. The assignment must post grades
    manually in Canvas; if it posts them automatically, a grade entered here
    would be visible to the student at once, so nothing is sent until that is
    changed. Posting stays a click in the Canvas gradebook.
  * A score changed in Canvas since the last push -- or entered there without
    ever being pushed -- is not overwritten. It is listed as CHANGED IN CANVAS
    with both numbers; --overwrite <netid> (or --overwrite all) replaces it.
  * A student who resubmitted on Canvas after the pull is skipped: the score
    is for an attempt that is no longer their latest. Pull and grade again.
  * A student already up to date -- this score, this exact document -- is
    left alone, so running it twice sends nothing twice.

It reads the documents finalize has reconciled, and refuses a student whose
document contradicts itself or whose score in scores.xlsx differs from the
document: both mean finalize needs to run again. Canvas user ids come from
canvas.json, which canvas_pull.py wrote. Every send is recorded there, which
is what the changed-in-Canvas check compares against next time.

Exit codes: 0 nothing needs attention; 1 something was skipped or blocked;
2 could not run, or Canvas refused a send partway through.
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from canvas_api import (CanvasError, Canvas, load_manifest, now_iso,  # noqa: E402
                        pulled_record, read_token, save_manifest, sha256_of,
                        user_for_netid, DEFAULT_TOKEN_FILE)
from sync_scores import (CODE_RE, check_internal, load_scores, num,  # noqa: E402
                         read_feedback, round_dirs)

NAME_RE = re.compile(r"^(?P<last>[^_]+)_(?P<first>[^_]+)_(?P<netid>[^_]+)_")

SEND, UP_TO_DATE, CHANGED, RESUBMITTED, BLOCKED = (
    "send", "up_to_date", "changed", "resubmitted", "blocked")


def entered(sub):
    """The score as entered, before Canvas's late deduction."""
    v = sub.get("entered_score")
    if v is None and sub.get("score") is not None:
        v = sub["score"] + (sub.get("points_deducted") or 0)
    return v


def same(a, b):
    if a is None or b is None:
        return a is None and b is None
    return abs(float(a) - float(b)) < 1e-6


def fmt(v):
    if v is None:
        return "none"
    v = float(v)
    return str(int(v)) if v.is_integer() else f"{v:g}"


def documents(adir, which):
    """(rounds in scope, [(round, path)]) for the round(s) being pushed."""
    rounds = round_dirs(adir)
    if not rounds:
        raise CanvasError(f"no feedback folder under {adir}")
    numbers = [r for r, _ in rounds]
    if which in (None, "", "latest"):
        wanted = {max(numbers)}
    elif str(which).lower() == "all":
        wanted = set(numbers)
    else:
        try:
            wanted = {int(which)}
        except ValueError:
            raise CanvasError(f"--round takes a number or 'all', not {which!r}")
        if not wanted <= set(numbers):
            raise CanvasError(f"no feedback folder for round {which}")
    docs = [(r, f) for r, d in rounds if r in wanted for f in sorted(d.iterdir())
            if f.suffix.lower() == ".docx" and not f.name.startswith(("~$", "._"))]
    return sorted(wanted), docs


def plan(adir, manifest, docs, subs, assignment, overwrite):
    """One row per student in scope, each with a status and the reason for it."""
    scores = adir / "scores.xlsx"
    if not scores.is_file():
        raise CanvasError(f"no scores.xlsx in {adir} -- finalize first")
    try:
        wb, ws, layout = load_scores(scores, data_only=True)
    except ValueError as exc:
        raise CanvasError(str(exc))
    totals = {k.lower(): ws.cell(r, layout["total_col"]).value if layout["total_col"] else None
              for k, r in layout["rows"].items()}
    possible_canvas = assignment.get("points_possible")

    rows, by_netid = [], {}
    for rnd, f in docs:
        if CODE_RE.match(f.name):
            rows.append({"status": BLOCKED, "netid": "?", "name": f.name, "round": rnd,
                         "why": "still masked -- run unmask on this round first"})
            continue
        m = NAME_RE.match(f.name)
        if not m:
            rows.append({"status": BLOCKED, "netid": "?", "name": f.name, "round": rnd,
                         "why": "filename is not last_first_netid_..."})
            continue
        by_netid.setdefault(m["netid"].lower(), []).append((rnd, f, m))

    for key, items in sorted(by_netid.items()):
        top = max(r for r, _, _ in items)
        mine = [it for it in items if it[0] == top]
        rnd, f, m = mine[0]
        row = {"netid": m["netid"], "name": f"{m['last']}, {m['first']}", "round": rnd,
               "doc": f}
        rows.append(row)

        def block(why):
            row.update(status=BLOCKED, why=why)

        if len(mine) > 1:
            block(f"{len(mine)} documents in feedback{rnd}/ -- keep one")
            continue
        try:
            fb = read_feedback(f)
        except Exception as exc:
            block(f"unreadable: {exc}")
            continue
        problems = check_internal(fb)
        if problems:
            block("the document disagrees with itself (" + "; ".join(problems) +
                  ") -- fix it and finalize again")
            continue
        score = fb["sum_of_items"]
        possible = fb["table_possible"]
        if possible is None:
            possible = sum(i["possible"] or 0 for i in fb["items"])
        row.update(score=score, possible=possible)

        total = num(totals.get(key)) if key in totals else None
        if key not in totals:
            block("no row in scores.xlsx -- finalize first")
            continue
        if total is None:
            block("scores.xlsx has no calculated Total for this student -- finalize again")
            continue
        if not same(total, score):
            block(f"scores.xlsx says {fmt(total)}, the document says {fmt(score)} "
                  f"-- finalize again")
            continue
        if possible_canvas is not None and not same(possible, possible_canvas):
            block(f"the document is out of {fmt(possible)}, the Canvas assignment out "
                  f"of {fmt(possible_canvas)}")
            continue

        uid, newest = user_for_netid(manifest, m["netid"])
        if uid is None:
            block("not in canvas.json -- this student's work was not pulled from Canvas")
            continue
        row["uid"] = uid
        sub = subs.get(str(uid))
        if sub is None:
            block(f"Canvas user {uid} has no submission record here (dropped the course?)")
            continue

        graded = pulled_record(manifest, f"submissions{rnd}", uid) or newest
        row["late_seconds"] = (sub.get("seconds_late") or 0) if sub.get("late") else 0
        g_att, c_att = graded.get("attempt"), sub.get("attempt")
        newer = (c_att is not None and g_att is not None and c_att > g_att) or (
            c_att is None and sub.get("submitted_at") and
            sub.get("submitted_at") != graded.get("submitted_at"))
        if newer:
            row.update(status=RESUBMITTED,
                       why=f"graded attempt {g_att}, Canvas now has attempt {c_att}")
            continue

        now = entered(sub)
        last = manifest.get("pushes", {}).get(str(uid))
        row.update(canvas=now, last=last["score"] if last else None)
        if sub.get("excused"):
            changed = "excused in Canvas"
        elif last is not None:
            if same(now, last["score"]) or (
                    now is None and sub.get("grade_matches_current_submission") is False):
                changed = None
            elif now is None:
                changed = f"score removed in Canvas; the last push sent {fmt(last['score'])}"
            else:
                changed = (f"Canvas has {fmt(now)}, the last push sent "
                           f"{fmt(last['score'])}")
        else:
            changed = (f"Canvas already has {fmt(now)}, entered there rather than pushed "
                       f"from here") if now is not None else None
        if changed and not sub.get("excused") and same(now, score):
            changed = None          # Canvas already holds this round's number
        if changed and not ("all" in overwrite or m["netid"].lower() in overwrite):
            row.update(status=CHANGED, why=changed)
            continue
        row["overwriting"] = changed

        sha = sha256_of(f)
        row.update(sha256=sha, attempt=g_att,
                   send_grade=not same(now, score),
                   send_doc=not (last and last.get("sha256") == sha))
        row["status"] = SEND if (row["send_grade"] or row["send_doc"]) else UP_TO_DATE
    return rows


def assignment_blockers(a):
    """Reasons nothing can be sent to this assignment at all."""
    out = []
    if not a.get("post_manually"):
        out.append(
            "this assignment posts grades automatically, so a grade entered now would be "
            "visible to the student at once. In Canvas, set it to post manually: Grades > "
            "the assignment's menu > Grade Posting Policy > Manually (or the gradebook "
            "settings for the whole course), then run this again.")
    if a.get("grading_type") not in (None, "points"):
        out.append(f"the assignment is graded as {a.get('grading_type')!r}; this sends "
                   f"points only")
    if a.get("moderated_grading"):
        out.append("moderated grading is on; a grade sent here would be provisional")
    return out


def days(seconds):
    d = seconds / 86400
    return f"{d:.1f} day(s)" if d >= 0.1 else f"{max(1, round(seconds / 60))} min"


def print_rows(rows, status, header, line):
    sel = [r for r in rows if r.get("status") == status]
    if not sel:
        return
    print(f"  {header.format(n=len(sel))}")
    for r in sel:
        print(f"      {r['netid']:<10} {r['name'][:28]:<28} {line(r)}")
    print()


def send_line(r):
    what = " + ".join(w for w, on in (("score", r["send_grade"]),
                                      ("feedback", r["send_doc"])) if on)
    s = (f"{fmt(r['score']):>5} / {fmt(r['possible']):<4}  Canvas now: {fmt(r['canvas']):<5}"
         f"  sends {what}")
    if r.get("late_seconds"):
        s += f"  [late {days(r['late_seconds'])}]"
    if r.get("overwriting"):
        s += f"  [OVERWRITES: {r['overwriting']}]"
    return s


def report(a, rounds, rows, blockers, sending):
    scope = ", ".join(f"feedback{r}/" for r in rounds)
    print(f"\n{a.get('name')!s} -- Canvas assignment {a['id']}, from {scope}")
    if not sending:
        print("DRY RUN: nothing has been sent. --send sends the SEND list below.\n")
    else:
        print()
    print_rows(rows, SEND, "SEND ({n}) -- raw scores; Canvas applies its late policy "
               "to late work:", send_line)
    print_rows(rows, UP_TO_DATE, "UP TO DATE ({n}) -- Canvas already has this score and "
               "this document:", lambda r: f"{fmt(r['score']):>5} / {fmt(r['possible'])}")
    print_rows(rows, CHANGED, "CHANGED IN CANVAS ({n}) -- left alone; --overwrite <netid> "
               "replaces it:", lambda r: f"this round says {fmt(r['score'])}: {r['why']}")
    print_rows(rows, RESUBMITTED, "RESUBMITTED SINCE THE PULL ({n}) -- skipped; pull and "
               "grade the new attempt:", lambda r: r["why"])
    print_rows(rows, BLOCKED, "BLOCKED ({n}):", lambda r: r["why"])
    for b in blockers:
        print(f"  CANNOT SEND: {b}\n")


def send(cv, manifest, adir, a, rows, comment):
    cid, aid = manifest["course_id"], a["id"]
    sent = []
    for r in [r for r in rows if r.get("status") == SEND]:
        uid = r["uid"]
        fid = None
        if r["send_doc"]:
            fid = cv.upload_comment_file(cid, aid, uid, r["doc"])
        data = []
        if r["send_grade"]:
            data.append(("submission[posted_grade]", fmt(r["score"])))
        if fid is not None:
            data += [("comment[text_comment]", comment), ("comment[file_ids][]", str(fid))]
        resp = cv.write("PUT", f"/courses/{cid}/assignments/{aid}/submissions/{uid}", data)
        recorded = entered(resp) if r["send_grade"] else r["score"]
        if recorded is None:
            recorded = r["score"]
        rec = {"netid": r["netid"], "score": recorded, "round": r["round"],
               "document": r["doc"].name, "sha256": r["sha256"], "file_id": fid,
               "attempt": r["attempt"], "pushed_at": now_iso()}
        manifest.setdefault("pushes", {})[str(uid)] = rec
        manifest.setdefault("push_log", []).append(
            {**rec, "user_id": str(uid), "sent_grade": r["send_grade"],
             "sent_document": r["send_doc"], "overwrote": r.get("overwriting")})
        save_manifest(adir, manifest)
        sent.append(r)
        what = " + ".join(w for w, on in (("score", r["send_grade"]),
                                          ("feedback", fid is not None)) if on)
        print(f"  sent  {r['netid']:<10} {fmt(r['score']):>5} / {fmt(r['possible'])}  {what}")
    return sent


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("assignment_dir", type=Path)
    ap.add_argument("--round", default=None,
                    help="which round to send: a number, or 'all' (default: the latest)")
    ap.add_argument("--send", action="store_true",
                    help="send for real (default: dry run, nothing leaves this machine)")
    ap.add_argument("--overwrite", action="append", default=[], metavar="NETID",
                    help="replace a score changed in Canvas for this student "
                         "(repeatable; 'all' for everyone)")
    ap.add_argument("--comment", help="the comment text the feedback document rides on "
                                      "(default: 'Your feedback for <assignment> is attached.')")
    ap.add_argument("--token-file", default=None,
                    help=f"Canvas API token (default: {DEFAULT_TOKEN_FILE})")
    args = ap.parse_args()

    adir = args.assignment_dir.expanduser().resolve()
    try:
        manifest = load_manifest(adir)
        if not manifest:
            raise CanvasError(f"no canvas.json in {adir}: this assignment was not pulled "
                              f"from Canvas (canvas_pull.py writes it)")
        cv = Canvas(manifest["host"], read_token(args.token_file or DEFAULT_TOKEN_FILE))
        cid, aid = manifest["course_id"], manifest["assignment_id"]
        a = cv.get(f"/courses/{cid}/assignments/{aid}")
        subs = {str(s["user_id"]): s
                for s in cv.get(f"/courses/{cid}/assignments/{aid}/submissions")}
        rounds, docs = documents(adir, args.round)
        if not docs:
            raise CanvasError(f"no feedback documents in {', '.join(f'feedback{r}' for r in rounds)}")
        overwrite = {o.lower() for o in args.overwrite}
        rows = plan(adir, manifest, docs, subs, a, overwrite)
        blockers = assignment_blockers(a)
        to_send = [r for r in rows if r.get("status") == SEND]
        report(a, rounds, rows, blockers, args.send and not blockers)
        attention = any(r.get("status") in (CHANGED, RESUBMITTED, BLOCKED) for r in rows)

        if not args.send:
            if to_send and not blockers:
                print(f"  {len(to_send)} to send. Nothing was sent.\n")
            return 1 if (attention or blockers) else 0
        if blockers:
            print("  Nothing was sent.\n")
            return 1
        if not to_send:
            print("  Nothing to send.\n")
            return 1 if attention else 0

        comment = args.comment or f"Your feedback for {a.get('name')} is attached."
        try:
            sent = send(cv, manifest, adir, a, rows, comment)
        except CanvasError as exc:
            print(f"\n  STOPPED: {exc}")
            print("  Everything listed as sent above did reach Canvas and is recorded in "
                  "canvas.json; run this again to continue.\n")
            return 2
        print(f"\n  Sent {len(sent)}. The grades are entered but hidden: post them from "
              f"the Canvas gradebook when you are ready.\n")
        return 1 if attention else 0
    except CanvasError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
