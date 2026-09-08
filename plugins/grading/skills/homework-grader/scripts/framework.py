#!/usr/bin/env python3
"""Stamp and verify the cached grading framework for an assignment.

    framework.py stamp <work/framework.json>
    framework.py check <work/framework.json>

Grading an assignment twice -- the first download, then the late work a week
later -- has to reach the same judgments both times. Rebuilding the rubric map
and the checker from scratch for the second batch produces different
partial-credit calls on identical work, which students notice and cannot be
defended. So the first pass writes its framework to `work/framework.json` and
later passes reuse it.

Reuse is only safe while the material it was built from has not moved. A rubric
line fixed in the content repo between batches would leave batch two graded
against a rubric that no longer exists. `stamp` records the commit that last
touched each source file; `check` re-resolves them and reports what changed.

`sources.files` is a list of {repo, path, commit}, where `repo` is a key in
`sources.repos`. Everything else in framework.json belongs to the grader.

Exit codes for `check`: 0 fresh, 1 stale, 2 could not tell.
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path


def last_commit(repo, path):
    """The commit that last touched one file, or None."""
    try:
        out = subprocess.run(
            ["git", "-C", str(repo), "log", "-1", "--format=%H", "--", str(path)],
            capture_output=True, text=True, timeout=30)
    except Exception:
        return None
    sha = out.stdout.strip()
    return sha or None


def load(path):
    p = Path(path).expanduser().resolve()
    if not p.is_file():
        sys.exit(f"error: no framework at {p}")
    try:
        doc = json.loads(p.read_text())
    except Exception as exc:
        sys.exit(f"error: could not read {p}: {exc}")
    src = doc.get("sources") or {}
    if not src.get("repos") or not src.get("files"):
        sys.exit(f"error: {p} has no sources.repos / sources.files block")
    return p, doc, src


def resolve(src, entry):
    repo = src["repos"].get(entry["repo"])
    if not repo:
        return None, f"unknown repo key {entry['repo']!r}"
    return last_commit(Path(repo).expanduser(), entry["path"]), None


def stamp(args):
    p, doc, src = load(args.framework)
    stamped, failed = 0, []
    for entry in src["files"]:
        sha, err = resolve(src, entry)
        if sha is None:
            failed.append((entry["path"], err or "not tracked in that repo"))
            continue
        entry["commit"] = sha
        stamped += 1
    p.write_text(json.dumps(doc, indent=2))
    print(f"Stamped {stamped} source file(s) in {p.name}")
    for path, why in failed:
        print(f"  COULD NOT STAMP  {path}: {why}")
    return 1 if failed else 0


def check(args):
    p, doc, src = load(args.framework)
    stale, unknown = [], []
    for entry in src["files"]:
        sha, err = resolve(src, entry)
        if sha is None:
            unknown.append((entry["path"], err or "not tracked in that repo"))
        elif entry.get("commit") != sha:
            stale.append((entry["path"], entry.get("commit"), sha))

    built = doc.get("built", "unknown")
    if not stale and not unknown:
        print(f"FRESH -- {p.name} (built {built}) still matches every source file.")
        print("Reuse the rubric map, the checker, and the decisions log.")
        return 0

    if stale:
        print(f"STALE -- {len(stale)} source file(s) changed since {p.name} was built "
              f"({built}):")
        for path, was, now in stale:
            print(f"      {path}")
            print(f"          was {(was or 'unstamped')[:12]}  now {now[:12]}")
        print("\nRe-read the changed material and rebuild the framework before grading.")
        print("Anything already graded was graded against the older version -- decide")
        print("whether those scores need revisiting.")
    if unknown:
        print(f"\nCOULD NOT CHECK ({len(unknown)}):")
        for path, why in unknown:
            print(f"      {path}: {why}")
    return 1 if stale else 2


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn, help_text in [
            ("stamp", stamp, "record the current commit for each source file"),
            ("check", check, "report whether any source file has moved since")]:
        s = sub.add_parser(name, help=help_text)
        s.add_argument("framework", type=Path)
        s.set_defaults(func=fn)
    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
