#!/usr/bin/env python3
"""Verify every repo a course depends on is clean and in sync before grading.

Grading reads the assignment and rubric from the public content repo and the
answer key from the private instructor repo. If either is stale, the whole
class gets graded against the wrong thing -- silently. This checks first.

Usage:
    preflight_sync.py <instructor_repo_path> [--pull] [--json]

    --pull   fast-forward any repo that is clean and merely behind
    --json   emit machine-readable output instead of a table

Exit codes:
    0  all repos clean and up to date (or fast-forwarded with --pull)
    1  at least one repo needs human attention
    2  bad invocation -- missing path, unreadable course.yml
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

OK, BEHIND, AHEAD, DIVERGED, DIRTY, NO_UPSTREAM, ERROR = (
    "ok", "behind", "ahead", "diverged", "dirty", "no-upstream", "error",
)

# States a human has to resolve; --pull cannot fix these.
BLOCKING = {AHEAD, DIVERGED, DIRTY, ERROR}

# Editor and OS droppings. Changes to these cannot make a rubric or key stale,
# so they are reported but do not block grading -- a check that fires on
# .DS_Store every time is a check people learn to ignore.
CRUFT = re.compile(
    r"(^|/)(\.DS_Store|Thumbs\.db|Icon\r?|desktop\.ini)$"
    r"|(^|/)__pycache__/|\.pyc$"
    r"|(^|/)\.idea/|(^|/)\.vscode/"
    r"|(^|/)~\$|(^|/)\.~lock\.|(^|/)\._"
)


def git(repo, *args):
    """Run a git command in repo. Returns (returncode, stdout, stderr)."""
    proc = subprocess.run(
        ["git", "-C", str(repo), *args],
        capture_output=True, text=True,
    )
    return proc.returncode, proc.stdout.strip(), proc.stderr.strip()


def load_course(instructor_path):
    """Read course.yml. Falls back to a minimal parser if PyYAML is absent."""
    manifest = instructor_path / "course.yml"
    if not manifest.exists():
        sys.exit(f"error: no course.yml in {instructor_path}")
    text = manifest.read_text()
    try:
        import yaml
        return yaml.safe_load(text)
    except ImportError:
        return _mini_yaml(text)


def _mini_yaml(text):
    """Parse the two-level key/value subset course.yml actually uses.

    Only exists so preflight works without PyYAML installed. Handles
    'key: value' and one level of nesting; ignores lists and comments.
    """
    root, current = {}, None
    for raw in text.splitlines():
        line = raw.split("#", 1)[0].rstrip()
        if not line.strip():
            continue
        indented = line[0] in " \t"
        key, _, value = line.strip().partition(":")
        key, value = key.strip(), value.strip().strip("'\"")
        if indented and current is not None:
            root[current][key] = value
        elif value:
            root[key] = value
        else:
            current = key
            root[key] = {}
    return root


def repos_from(course, instructor_path):
    """Yield (label, path, declared_remote) for every repo the course needs."""
    found = []
    for label, key in (("content", "content_repo"), ("instructor", "instructor_repo")):
        spec = course.get(key)
        if not isinstance(spec, dict):
            continue
        local = spec.get("local", ".")
        # Paths in course.yml are relative to the instructor repo root.
        path = (instructor_path / local).resolve()
        found.append((label, path, spec.get("remote")))
    return found


def check(path, do_pull):
    """Fetch and classify one repo's sync state."""
    if not (path / ".git").exists():
        return {"state": ERROR, "detail": f"not a git repo: {path}"}

    code, _, err = git(path, "fetch", "--quiet")
    if code != 0:
        return {"state": ERROR, "detail": f"fetch failed: {err}"}

    # Uncommitted work blocks everything -- never pull over it.
    _, porcelain, _ = git(path, "status", "--porcelain")
    cruft = []
    if porcelain:
        # Split off the two-column status code rather than slicing a fixed
        # offset: git() strips stdout, so the first line has already lost its
        # leading space. maxsplit=1 keeps filenames containing spaces intact.
        files = [ln.strip().split(maxsplit=1)[-1] for ln in porcelain.splitlines()]
        cruft = [f for f in files if CRUFT.search(f)]
        real = [f for f in files if not CRUFT.search(f)]
        if real:
            more = " (showing 10)" if len(real) > 10 else ""
            return {"state": DIRTY, "detail": f"{len(real)} uncommitted{more}",
                    "files": real[:10], "cruft": len(cruft)}

    code, branch, _ = git(path, "rev-parse", "--abbrev-ref", "HEAD")
    code, upstream, _ = git(path, "rev-parse", "--abbrev-ref", "@{upstream}")
    if code != 0:
        return {"state": NO_UPSTREAM, "branch": branch, "cruft": len(cruft),
                "detail": f"branch '{branch}' tracks nothing"}

    _, counts, _ = git(path, "rev-list", "--left-right", "--count", "@{upstream}...HEAD")
    behind, ahead = (int(n) for n in counts.split())

    result = {"branch": branch, "upstream": upstream, "behind": behind,
              "ahead": ahead, "cruft": len(cruft)}

    if behind and ahead:
        return {**result, "state": DIVERGED,
                "detail": f"{behind} behind and {ahead} ahead of {upstream}"}
    if ahead:
        return {**result, "state": AHEAD,
                "detail": f"{ahead} unpushed commit(s)"}
    if behind:
        if not do_pull:
            return {**result, "state": BEHIND,
                    "detail": f"{behind} commit(s) behind {upstream}"}
        code, _, err = git(path, "merge", "--ff-only", "@{upstream}")
        if code != 0:
            return {**result, "state": ERROR, "detail": f"fast-forward failed: {err}"}
        return {**result, "state": OK, "behind": 0,
                "detail": f"pulled {behind} commit(s)"}

    return {**result, "state": OK, "detail": "up to date"}


ADVICE = {
    DIRTY: "commit or stash before grading",
    AHEAD: "push, or confirm the unpushed work is intentional",
    DIVERGED: "reconcile with the remote by hand",
    BEHIND: "re-run with --pull",
    NO_UPSTREAM: "set an upstream, or confirm this repo is local-only",
    ERROR: "resolve the error above",
}

MARK = {OK: "ok", BEHIND: "BEHIND", AHEAD: "AHEAD", DIVERGED: "DIVERGED",
        DIRTY: "DIRTY", NO_UPSTREAM: "NO UPSTREAM", ERROR: "ERROR"}


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("instructor_repo", type=Path,
                    help="path to the private instructor repo holding course.yml")
    ap.add_argument("--pull", action="store_true",
                    help="fast-forward repos that are clean and behind")
    ap.add_argument("--json", action="store_true", dest="as_json",
                    help="emit JSON instead of a table")
    args = ap.parse_args()

    instructor_path = args.instructor_repo.expanduser().resolve()
    if not instructor_path.is_dir():
        sys.exit(f"error: no such directory: {instructor_path}")

    course = load_course(instructor_path)
    targets = repos_from(course, instructor_path)
    if not targets:
        sys.exit("error: course.yml declares no repos")

    results = {}
    for label, path, remote in targets:
        r = check(path, args.pull)
        r["path"] = str(path)
        if remote:
            r["declared_remote"] = remote
        results[label] = r

    blocked = [k for k, v in results.items() if v["state"] in BLOCKING]
    stale = [k for k, v in results.items() if v["state"] in (BEHIND, NO_UPSTREAM)]

    if args.as_json:
        print(json.dumps({
            "course": course.get("course"),
            "repos": results,
            "ready": not (blocked or stale),
        }, indent=2))
        return 1 if (blocked or stale) else 0

    print(f"\n{course.get('course', 'course')} — sync preflight\n")
    for label, r in results.items():
        print(f"  {MARK[r['state']]:<12} {label:<11} {r['detail']}")
        print(f"  {'':<12} {'':<11} {r['path']}")
        for f in r.get("files", []):
            print(f"  {'':<12} {'':<11}   {f}")
        if r.get("cruft"):
            print(f"  {'':<12} {'':<11} ({r['cruft']} ignored: .DS_Store, __pycache__, editor files)")
        if r["state"] in ADVICE:
            print(f"  {'':<12} {'':<11} -> {ADVICE[r['state']]}")
        print()

    if blocked or stale:
        print("NOT READY — resolve the above before grading.\n")
        return 1

    print("Ready to grade.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
