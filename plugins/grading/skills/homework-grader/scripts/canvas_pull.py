#!/usr/bin/env python3
"""Download an assignment's submissions from Canvas into the grading workspace.

    canvas_pull.py <instructor_repo> <assignment_key> [--term TERM] [--dest DIR]
                   [--assignment-id ID]

The Canvas counterpart of downloading from Learning Suite. Every student's
current submission (their latest attempt, every file in it) lands in the next
numbered `submissions<N>/` folder of the assignment, named the way Learning
Suite names its downloads:

    last_first_netid_theirfilename.ext

so `anonymize.py mask` and everything after it runs unchanged. Each file's
modification time is set to the Canvas submission time, which is where mask
reads its timestamps from.

Like a Learning Suite download, a later pull holds the whole class rather than
the difference: mask hashes every file against the previous round and grades
only what changed, so an unchanged student is reported UNCHANGED instead of
MISSING. Files already pulled are copied from the earlier folder rather than
downloaded again, so their bytes, and their hashes, are identical.

  * Nothing new on Canvas since the last pull: no folder is created.
  * The newest folder has been neither masked nor graded: it is replaced in
    place rather than starting another round, as long as it holds only files a
    pull put there.
  * A resubmitted file Canvas renamed to avoid a clash (hw1.xlsx -> hw1-1.xlsx)
    keeps the student's earlier name, so mask sees it as RESUBMITTED rather
    than as a new file beside a missing one.

`canvas.json` in the assignment folder (mode 600) records each file's Canvas
user id, attachment id and attempt, which is how canvas_push.py finds the
student to send a grade to. Like roster.json it never goes in a repo.

The report gives counts, not names: nothing has been masked yet, and the
masking step prints the per-student detail under codes.

The assignment folder is <grading_workspace>/<term>/<assignment_key>, with the
term from --term, the course's `canvas.term`, or the Canvas course's own term
("2027 Winter" -> 2027-winter). --dest names the folder outright.

Exit codes: 0 pulled, or nothing to pull; 1 pulled, with something to look
at; 2 could not pull.
"""

import argparse
import os
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from canvas_api import (CanvasError, Canvas, canvas_settings, clean_filename,  # noqa: E402
                        load_course, load_manifest, local_filename, netid_for,
                        now_iso, parse_time, read_token, resolve_assignment,
                        save_manifest, split_name, term_slug)
from anonymize import SUBMISSION_RE, skip  # noqa: E402

ROUND_DIR = re.compile(r"^submissions(\d+)$")
# Canvas renames an upload that clashes with an earlier one: hw1.xlsx -> hw1-1.xlsx.
CANVAS_RENAME = re.compile(r"^(?P<base>.+?)-(?P<n>\d+)(?P<ext>\.[^.]+)?$")


def current_submissions(cv, cfg, assignment_id):
    """Every student's latest attempt that has files, keyed by Canvas user id."""
    subs = cv.get(f"/courses/{cfg['course_id']}/assignments/{assignment_id}/submissions",
                  {"include[]": ["user"]})
    found, none, not_upload = {}, 0, []
    for s in subs:
        uid = str(s.get("user_id"))
        if not s.get("submitted_at") or s.get("workflow_state") == "unsubmitted":
            none += 1
            continue
        atts = [a for a in (s.get("attachments") or []) if a.get("url")]
        if not atts:
            # A text entry or URL submission: there is no file to fetch.
            not_upload.append(uid)
            continue
        user = s.get("user") or {"id": s.get("user_id")}
        user.setdefault("id", s.get("user_id"))
        netid, source = netid_for(user, cfg["netid_field"], cfg["netid_domain"])
        last, first = split_name(user)
        found[uid] = {
            "netid": netid, "netid_from": source, "last": last, "first": first,
            "attempt": s.get("attempt"), "submitted_at": s.get("submitted_at"),
            "late": bool(s.get("late")), "seconds_late": s.get("seconds_late") or 0,
            "attachments": [{"id": a["id"], "size": a.get("size"), "url": a["url"],
                             "display_name": a.get("display_name") or a.get("filename")}
                            for a in atts],
        }

    seen = {}
    for uid, rec in found.items():
        other = seen.setdefault(rec["netid"].lower(), uid)
        if other != uid:
            raise CanvasError(f"Canvas users {other} and {uid} both give the NetID slot "
                              f"the same value; mask would merge them. Fix the user "
                              f"records, or set canvas.netid_field in course.yml.")
    return found, none, not_upload


def attachment_ids(rec):
    return sorted(a["id"] for a in rec["attachments"])


def pulled_ids(entry):
    return sorted(f["attachment_id"] for f in entry["files"])


def same_state(current, students):
    if set(current) != set(students):
        return False
    return all(attachment_ids(current[u]) == pulled_ids(students[u]) for u in current)


def submission_rounds(adir):
    out = []
    for d in adir.iterdir() if adir.is_dir() else []:
        m = ROUND_DIR.match(d.name)
        if d.is_dir() and m:
            out.append((int(m.group(1)), d))
    return sorted(out)


def masked_batches(adir):
    """Submission folders anonymize.py has already masked, from ledger.json."""
    import json
    p = adir / "ledger.json"
    if not p.exists():
        return set()
    try:
        return set(json.loads(p.read_text()).get("batches", []))
    except ValueError:
        return set()


def choose_folder(adir, manifest):
    """(folder to write, pull entry it replaces or None)."""
    rounds = submission_rounds(adir)
    if not rounds:
        return adir / "submissions1", None
    n, latest = rounds[-1]
    # Masked, or graded without masking (de-identification skipped): either
    # way it is a round now, and a round is never rewritten.
    if latest.name in masked_batches(adir) or (adir / f"feedback{n}").exists():
        return adir / f"submissions{n + 1}", None

    # Never masked or graded, so nothing downstream has read it. Replace it,
    # but only if it holds nothing but what an earlier pull wrote.
    entry = next((p for p in manifest["pulls"] if p["folder"] == latest.name), None)
    if entry is None:
        raise CanvasError(
            f"{latest.name}/ has not been masked yet and was not pulled from Canvas. "
            f"Mask and grade it first, or move it out of {adir}, then pull again.")
    ours = {f["file"] for st in entry["students"].values() for f in st["files"]}
    extra = sorted(f.name for f in latest.iterdir() if not skip(f) and f.name not in ours)
    if extra:
        raise CanvasError(
            f"{latest.name}/ has not been masked yet, so this pull would replace it, "
            f"but it holds {len(extra)} file(s) no pull put there. Mask it first, or "
            f"move those files out, then pull again.")
    return latest, entry


def _unique(name, used):
    if name not in used:
        return name
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    i = 2
    while True:
        cand = f"{stem} ({i})" + (f".{ext}" if ext else "")
        if cand not in used:
            return cand
        i += 1


def plan_names(rec, prev_names, known):
    """The student's own filename for each attachment: {attachment id: name}.

    A file pulled before keeps the name it had. A file Canvas renamed to avoid
    clashing with the student's earlier upload (hw1-1.xlsx) gets that earlier
    name back, so mask compares the new bytes with the old ones under one name
    and reports a resubmission. Two attachments with the same name in one
    attempt are told apart with (2), (3).
    """
    names, used = {}, set()
    atts = rec["attachments"]
    for a in atts:
        k = known.get(a["id"])
        if k and k["name"] not in used:
            names[a["id"]] = k["name"]
            used.add(k["name"])
    for a in atts:
        if a["id"] in names:
            continue
        name = clean_filename(a["display_name"])
        if (name in prev_names or not CANVAS_RENAME.match(name)) and name not in used:
            names[a["id"]] = name
            used.add(name)
    for a in atts:
        if a["id"] in names:
            continue
        name = clean_filename(a["display_name"])
        m = CANVAS_RENAME.match(name)
        if m:
            base = m["base"] + (m["ext"] or "")
            if base in prev_names and base not in used:
                name = base
        name = _unique(name, used)
        names[a["id"]] = name
        used.add(name)
    return names


def pull(cv, cfg, assignment, adir):
    """Pull the current submissions. Returns a result dict for report()."""
    aid = int(assignment["id"])
    adir = Path(adir)
    manifest = load_manifest(adir)
    if manifest and int(manifest.get("assignment_id", aid)) != aid:
        raise CanvasError(f"{adir} was pulled from Canvas assignment "
                          f"{manifest['assignment_id']}, not {aid}; refusing to mix them")

    current, none, not_upload = current_submissions(cv, cfg, aid)
    result = {"assignment": assignment, "adir": adir, "none": none,
              "not_upload": not_upload, "students": len(current)}
    if not current:
        result["status"] = "none_submitted"
        return result

    adir.mkdir(parents=True, exist_ok=True)
    manifest = manifest or {
        "host": cfg["host"], "course_id": cfg["course_id"], "assignment_id": aid,
        "pulls": [], "pushes": {}, "push_log": [],
    }
    manifest["assignment_name"] = assignment.get("name")
    manifest["points_possible"] = assignment.get("points_possible")
    pulls = manifest["pulls"]

    if pulls and same_state(current, pulls[-1]["students"]):
        result.update(status="nothing_new", folder=adir / pulls[-1]["folder"])
        return result

    folder, replacing = choose_folder(adir, manifest)
    if replacing is not None:
        i = pulls.index(replacing)
        baseline = pulls[i - 1] if i > 0 else None
    else:
        baseline = pulls[-1] if pulls else None
    base_students = baseline["students"] if baseline else {}

    # Every file any pull fetched, newest first, so a file seen before is
    # copied rather than downloaded again.
    known = {}
    for p in reversed(pulls):
        for st in p["students"].values():
            for f in st["files"]:
                src = adir / p["folder"] / f["file"]
                if f["attachment_id"] not in known and src.exists():
                    known[f["attachment_id"]] = {**f, "path": src}

    build = adir / f".{folder.name}.partial"
    if build.exists():
        shutil.rmtree(build)
    build.mkdir()
    entry = {"folder": folder.name, "pulled_at": now_iso(), "students": {}}
    changes = {"new": [], "resubmitted": [], "unchanged": []}
    fallback, late, copied, downloaded = [], [], 0, 0
    try:
        for uid, rec in sorted(current.items()):
            prev = base_students.get(uid)
            if prev is None:
                changes["new"].append(uid)
            elif pulled_ids(prev) == attachment_ids(rec):
                changes["unchanged"].append(uid)
            else:
                changes["resubmitted"].append(uid)
            if rec["netid_from"] == "user_id":
                fallback.append(uid)
            if rec["late"]:
                late.append(uid)

            prev_names = {f["name"] for f in (prev or {}).get("files", [])}
            names = plan_names(rec, prev_names, known)
            when = parse_time(rec["submitted_at"])
            files = []
            for a in rec["attachments"]:
                fname = local_filename(rec["last"], rec["first"], rec["netid"], names[a["id"]])
                m = SUBMISSION_RE.match(Path(fname).stem)
                if not m or m["netid"] != rec["netid"]:
                    raise CanvasError(f"built a filename mask cannot parse for Canvas "
                                      f"user {uid}; this is a bug in canvas_pull.py")
                dest = build / fname
                k = known.get(a["id"])
                if k:
                    shutil.copy2(k["path"], dest)
                    sha, size = k["sha256"], dest.stat().st_size
                    copied += 1
                else:
                    sha, size = cv.download(a["url"], dest)
                    downloaded += 1
                    if a.get("size") not in (None, size):
                        raise CanvasError(f"download for Canvas user {uid} came back "
                                          f"{size} bytes, Canvas says {a['size']}")
                if when:
                    os.utime(dest, (when, when))
                files.append({"file": fname, "name": names[a["id"]],
                              "canvas_name": a["display_name"], "attachment_id": a["id"],
                              "size": size, "sha256": sha})
            entry["students"][uid] = {
                "netid": rec["netid"], "netid_from": rec["netid_from"],
                "attempt": rec["attempt"], "submitted_at": rec["submitted_at"],
                "late": rec["late"], "seconds_late": rec["seconds_late"], "files": files,
            }
    except BaseException:
        shutil.rmtree(build, ignore_errors=True)
        raise

    if replacing is not None:
        shutil.rmtree(folder)
        pulls.remove(replacing)
    build.rename(folder)
    pulls.append(entry)
    save_manifest(adir, manifest)

    result.update(
        status="pulled", folder=folder, replaced=replacing is not None,
        baseline=baseline["folder"] if baseline else None, changes=changes,
        gone=sorted(set(base_students) - set(current)), fallback=fallback, late=late,
        files=sum(len(s["files"]) for s in entry["students"].values()),
        copied=copied, downloaded=downloaded)
    return result


def assignment_warnings(a):
    out = []
    if a.get("group_category_id") and not a.get("grade_group_students_individually"):
        out.append("this is a group assignment: each member's files are the group's")
    if a.get("anonymous_grading"):
        out.append("anonymous grading is on in Canvas")
    if a.get("moderated_grading"):
        out.append("moderated grading is on: canvas_push.py will not send grades")
    if "online_upload" not in (a.get("submission_types") or []):
        out.append(f"submission types are {a.get('submission_types')}, not file uploads")
    return out


def report(r):
    a = r["assignment"]
    print(f"\n{a.get('name')!s} -- Canvas assignment {a['id']}")
    for w in assignment_warnings(a):
        print(f"  note: {w}")
    if r["status"] == "none_submitted":
        print(f"  Nobody has submitted yet ({r['none']} student(s) with no submission).")
        print("  Nothing was written.\n")
        return 0
    if r["status"] == "nothing_new":
        print(f"  Nothing new since {r['folder'].name}/ -- no new or changed submission "
              f"on Canvas.")
        print("  No folder was created.\n")
        return 0

    since = r["baseline"] or "the last pull"
    verb = "replaced" if r["replaced"] else "into"
    print(f"  {r['students']} student(s), {r['files']} file(s) {verb} {r['folder']}/")
    if r["replaced"]:
        print(f"  ({r['folder'].name}/ had not been masked yet, so this pull replaced it "
              f"rather than starting another round)")
    print(f"  ({r['downloaded']} downloaded, {r['copied']} copied from earlier pulls)\n")
    c = r["changes"]
    print(f"  NEW           {len(c['new']):>4}   first submission from this student")
    print(f"  RESUBMITTED   {len(c['resubmitted']):>4}   a newer attempt than {since}")
    print(f"  UNCHANGED     {len(c['unchanged']):>4}   same files as {since}; mask skips them")
    print(f"  LATE          {len(r['late']):>4}   per Canvas, which applies the late policy "
          f"itself when a score is entered")
    print(f"  NO SUBMISSION {r['none']:>4}\n")
    attention = False
    if r["gone"]:
        print(f"  NO LONGER ON CANVAS ({len(r['gone'])}) -- in {since}, absent now. Nothing "
              f"was deleted. Canvas user id(s): {', '.join(r['gone'])}\n")
        attention = True
    if r["not_upload"]:
        print(f"  NOT A FILE UPLOAD ({len(r['not_upload'])}) -- submitted text or a URL, "
              f"so there is no file to pull. Canvas user id(s): "
              f"{', '.join(r['not_upload'])}\n")
        attention = True
    if r["fallback"]:
        print(f"  NO NETID ({len(r['fallback'])}) -- no usable login_id, so the Canvas "
              f"user id fills the NetID slot. Canvas user id(s): "
              f"{', '.join(r['fallback'])}\n")
        attention = True
    print("  Next: mask it.")
    print(f"      python anonymize.py mask {r['folder']}\n")
    return 1 if attention else 0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("instructor_repo", type=Path, help="the repo holding course.yml")
    ap.add_argument("assignment_key", help="e.g. 01_head")
    ap.add_argument("--term", help="term folder in the workspace, e.g. 2027-winter")
    ap.add_argument("--dest", type=Path, help="the assignment folder, outright")
    ap.add_argument("--assignment-id", type=int, help="the Canvas assignment id, outright")
    args = ap.parse_args()

    try:
        course = load_course(args.instructor_repo)
        cfg = canvas_settings(course, args.instructor_repo)
        cv = Canvas(cfg["host"], read_token(cfg["token_file"]))
        aid, how = resolve_assignment(cv, cfg, args.assignment_key, args.assignment_id)
        assignment = cv.get(f"/courses/{cfg['course_id']}/assignments/{aid}")
        print(f"Canvas assignment {aid} ({how}): {assignment.get('name')!r}")

        if args.dest:
            adir = args.dest.expanduser().resolve()
        else:
            if not cfg["workspace"]:
                raise CanvasError("course.yml has no grading_workspace; pass --dest")
            term = args.term or cfg["term"]
            if not term:
                c = cv.get(f"/courses/{cfg['course_id']}", {"include[]": ["term"]})
                name = (c.get("term") or {}).get("name")
                term = term_slug(name)
                if not term:
                    raise CanvasError(f"cannot make a term folder from the Canvas term "
                                      f"{name!r}; pass --term (e.g. 2027-winter)")
            adir = cfg["workspace"] / term / args.assignment_key
        return report(pull(cv, cfg, assignment, adir))
    except CanvasError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
