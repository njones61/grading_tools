#!/usr/bin/env python3
"""Shared Canvas plumbing for canvas_pull.py and canvas_push.py.

Holds the API client, the course.yml `canvas:` block, the mapping from an
assignment key to a Canvas assignment id, and the per-assignment record
(`canvas.json`) that lets the push step find each student's Canvas user id.

The API token is read from ~/.config/canvas/token (or the course's
`canvas.token_file`). It is sent only as an Authorization header, and only to
the course's own Canvas host: never in a URL, never to the file-storage hosts
Canvas redirects downloads and uploads to, and never printed or written.

course.yml, for a course that has moved to Canvas:

    lms: canvas
    canvas:
      host: https://byu.instructure.com
      course_id: 40232
      term: 2027-winter                 # optional; read from Canvas otherwise
      assignments_file: canvas/state.json   # JSON of key -> assignment id
      assignments_prefix: "hw:"             # its keys are hw:<assignment key>
      assignments:                          # explicit ids; these win
        01_head: 1532387
      netid_field: login_id             # optional; the user field holding the NetID
      netid_domain: byu.edu             # optional; strip "@byu.edu" from that field

An assignment key found in neither map is matched against the Canvas
assignment names, and the match is reported so it can be checked.
"""

import hashlib
import json
import os
import re
import stat
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit

DEFAULT_TOKEN_FILE = "~/.config/canvas/token"
MANIFEST = "canvas.json"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class CanvasError(Exception):
    """Something the user has to fix. The message says what."""


# --------------------------------------------------------------------------- #
# Token
# --------------------------------------------------------------------------- #
def read_token(path=DEFAULT_TOKEN_FILE):
    """The API token, from a file only its owner can read.

    A token readable by other accounts on the machine is refused rather than
    used: it grants everything the instructor can do in every course.
    """
    p = Path(path).expanduser()
    if not p.is_file():
        raise CanvasError(
            f"no Canvas API token at {p}. Make one in Canvas (Account > Settings > "
            "New Access Token), then:\n"
            f"    mkdir -p {p.parent} && pbpaste > {p} && chmod 600 {p}")
    mode = stat.S_IMODE(p.stat().st_mode)
    if mode & 0o077:
        raise CanvasError(f"{p} is readable by other accounts (mode {mode:o}). "
                          f"Run: chmod 600 {p}")
    token = p.read_text().strip()
    if not token:
        raise CanvasError(f"{p} is empty")
    return token


# --------------------------------------------------------------------------- #
# Client
# --------------------------------------------------------------------------- #
def _same_origin(url, host):
    a, b = urlsplit(url), urlsplit(host)
    return (a.scheme, a.netloc.lower()) == (b.scheme, b.netloc.lower())


def _short(text, n=300):
    text = (text or "").strip().replace("\n", " ")
    return text if len(text) <= n else text[:n] + "..."


class Canvas:
    """A thin Canvas REST client: paginated GETs, form writes, file transfer.

    `session` is anything with a requests-style `request()` method, which is
    how the tests substitute a fake Canvas.
    """

    RETRY_STATUS = {429, 500, 502, 503, 504}

    def __init__(self, host, token, session=None, sleep=time.sleep):
        if session is None:
            import requests
            session = requests.Session()
        self.host = host.rstrip("/")
        self.s = session
        self._token = token
        self._sleep = sleep

    def __repr__(self):
        return f"Canvas({self.host!r})"

    def url(self, path):
        return path if path.startswith("http") else f"{self.host}/api/v1{path}"

    def request(self, method, path, *, params=None, data=None, files=None,
                allow_redirects=True, stream=False, auth=True):
        url = self.url(path)
        # The token goes to the Canvas host and nowhere else. requests also
        # drops it on a redirect to another host, which is where file
        # downloads end up.
        headers = ({"Authorization": f"Bearer {self._token}"}
                   if auth and _same_origin(url, self.host) else {})
        tries = 4 if method == "GET" else 1
        for attempt in range(tries):
            r = self.s.request(method, url, params=params, data=data, files=files,
                               headers=headers, allow_redirects=allow_redirects,
                               stream=stream, timeout=120)
            throttled = (r.status_code in self.RETRY_STATUS or
                         (r.status_code == 403 and "rate limit" in (r.text or "").lower()))
            if throttled and attempt < tries - 1:
                self._sleep(2 ** attempt)
                continue
            return r
        return r

    def _check(self, r, method, path):
        if not r.ok:
            # The path only: a download URL carries a verifier in its query.
            where = urlsplit(self.url(path)).path
            raise CanvasError(f"Canvas refused {method} {where}: "
                              f"{r.status_code} {_short(r.text)}")

    def get(self, path, params=None):
        """GET one object, or every page of a list (following the Link header)."""
        params = dict(params or {})
        params.setdefault("per_page", 100)
        url, out = self.url(path), None
        while url:
            r = self.request("GET", url, params=params)
            self._check(r, "GET", path)
            data = r.json()
            if not isinstance(data, list):
                return data
            out = (out or []) + data
            url = (getattr(r, "links", None) or {}).get("next", {}).get("url")
            params = None          # the next link carries its own query
        return out or []

    def write(self, method, path, data):
        r = self.request(method, path, data=data)
        self._check(r, method, path)
        return r.json() if r.content else {}

    def download(self, url, dest):
        """Stream a file to dest. Returns (sha256, size)."""
        r = self.request("GET", url, stream=True)
        self._check(r, "GET", url)
        h, size = hashlib.sha256(), 0
        tmp = Path(str(dest) + ".part")
        try:
            with open(tmp, "wb") as fh:
                for chunk in r.iter_content(1 << 16):
                    if chunk:
                        fh.write(chunk)
                        h.update(chunk)
                        size += len(chunk)
            tmp.replace(dest)
        finally:
            tmp.unlink(missing_ok=True)
            close = getattr(r, "close", None)
            if close:
                close()
        return h.hexdigest(), size

    def upload_comment_file(self, course_id, assignment_id, user_id, path,
                            content_type=DOCX_MIME):
        """Upload a file for a submission comment. Returns its Canvas file id.

        Canvas's three-step upload: ask for an upload URL, post the bytes there
        (without the token -- it is usually another host), then confirm if
        Canvas answers with a redirect.
        """
        path = Path(path)
        pre = self.write("POST", f"/courses/{course_id}/assignments/{assignment_id}"
                                 f"/submissions/{user_id}/comments/files",
                         {"name": path.name, "size": path.stat().st_size,
                          "content_type": content_type})
        upload_url = pre.get("upload_url")
        if not upload_url:
            raise CanvasError(f"Canvas gave no upload URL for {path.name}")
        field = pre.get("file_param") or "file"
        with open(path, "rb") as fh:
            r = self.request("POST", upload_url, data=pre.get("upload_params") or {},
                             files={field: (path.name, fh, content_type)},
                             allow_redirects=False, auth=False)
        if r.status_code in (301, 302, 303):
            loc = r.headers.get("Location")
            r = self.request("GET", loc)
            self._check(r, "GET", loc)
        elif not r.ok:
            raise CanvasError(f"uploading {path.name} failed: {r.status_code} "
                              f"{_short(r.text)}")
        info = r.json()
        if "id" not in info:
            raise CanvasError(f"upload of {path.name} returned no file id")
        return info["id"]


# --------------------------------------------------------------------------- #
# course.yml
# --------------------------------------------------------------------------- #
def load_course(instructor_repo):
    repo = Path(instructor_repo).expanduser().resolve()
    manifest = repo / "course.yml"
    if not manifest.is_file():
        raise CanvasError(f"no course.yml in {repo}")
    try:
        import yaml
    except ImportError:
        raise CanvasError("PyYAML is not installed: pip install pyyaml")
    return yaml.safe_load(manifest.read_text()) or {}


def lms_of(course):
    """learning_suite unless course.yml says otherwise."""
    return str(course.get("lms") or "learning_suite").strip().lower()


def canvas_settings(course, instructor_repo):
    """The course's `canvas:` block, checked and with defaults filled in."""
    c = course.get("canvas")
    if not isinstance(c, dict):
        raise CanvasError("course.yml has no `canvas:` block (host, course_id)")
    host = str(c.get("host") or "").rstrip("/")
    if not host.startswith("https://"):
        raise CanvasError(f"canvas.host must be an https:// URL, not {host!r}")
    try:
        course_id = int(c["course_id"])
    except (KeyError, TypeError, ValueError):
        raise CanvasError("canvas.course_id must be the numeric Canvas course id")

    repo = Path(instructor_repo).expanduser().resolve()
    explicit = {}
    for k, v in (c.get("assignments") or {}).items():
        try:
            explicit[str(k)] = int(v)
        except (TypeError, ValueError):
            raise CanvasError(f"canvas.assignments.{k} is not a numeric assignment id: {v!r}")
    afile = c.get("assignments_file")
    return {
        "host": host,
        "course_id": course_id,
        "term": str(c["term"]) if c.get("term") else None,
        "assignments": explicit,
        "assignments_file": (repo / Path(afile).expanduser()) if afile else None,
        "assignments_prefix": str(c.get("assignments_prefix") or ""),
        "netid_field": str(c.get("netid_field") or "login_id"),
        "netid_domain": str(c.get("netid_domain") or "byu.edu"),
        "token_file": str(c.get("token_file") or DEFAULT_TOKEN_FILE),
        "workspace": (Path(course["grading_workspace"]).expanduser()
                      if course.get("grading_workspace") else None),
    }


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def match_by_name(assignments, key):
    """The one Canvas assignment whose name fits an assignment key, or an error.

    Exact name first, then every word of the key (01_head -> head) inside the
    name. Ambiguity is an error that lists the candidates, never a guess.
    """
    uploads = [a for a in assignments
               if "online_upload" in (a.get("submission_types") or [])] or assignments
    exact = [a for a in uploads if _norm(a.get("name")) == _norm(key)]
    if len(exact) == 1:
        return exact[0]
    words = [w for w in re.split(r"[^a-z]+", key.lower()) if len(w) >= 3]

    def fits(a):
        n = _norm(a.get("name"))
        return bool(words) and all(w in n or (w.endswith("s") and w[:-1] in n)
                                   for w in words)

    hits = [a for a in uploads if fits(a)]
    if len(hits) == 1:
        return hits[0]
    shown = hits or uploads
    listing = "\n".join(f"    {a['id']}  {a.get('name')}" for a in shown[:40])
    why = (f"{len(hits)} Canvas assignments match" if hits
           else "no Canvas assignment name matches")
    raise CanvasError(
        f"{why} the key {key!r}. Add it to course.yml under canvas.assignments "
        f"({key}: <id>), or pass --assignment-id. Candidates:\n{listing}")


def resolve_assignment(cv, cfg, key, override=None):
    """(assignment id, how it was found) for an assignment key."""
    if override:
        return int(override), "--assignment-id"
    if key in cfg["assignments"]:
        return cfg["assignments"][key], "course.yml canvas.assignments"
    af = cfg["assignments_file"]
    if af:
        if not af.is_file():
            raise CanvasError(f"canvas.assignments_file {af} does not exist")
        try:
            data = json.loads(af.read_text())
        except ValueError as exc:
            raise CanvasError(f"could not read {af}: {exc}")
        for k in (cfg["assignments_prefix"] + key, key):
            v = data.get(k) if isinstance(data, dict) else None
            if v is not None and str(v).isdigit():
                return int(v), f"{af.name} ({k})"
    found = match_by_name(cv.get(f"/courses/{cfg['course_id']}/assignments"), key)
    return found["id"], f"matched by name: {found.get('name')!r}"


def term_slug(name):
    """'2027 Winter' or 'Winter 2027' -> '2027-winter', the workspace's term folder."""
    year = re.search(r"\b(20\d\d)\b", name or "")
    season = re.search(r"\b(winter|spring|summer|fall|autumn)\b", (name or "").lower())
    if not (year and season):
        return None
    s = "fall" if season.group(1) == "autumn" else season.group(1)
    return f"{year.group(1)}-{s}"


# --------------------------------------------------------------------------- #
# Students and filenames
# --------------------------------------------------------------------------- #
NETID_OK = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]*$")


def netid_for(user, field="login_id", domain="byu.edu"):
    """(netid, source) for a Canvas user.

    The NetID comes from `field` (login_id at BYU), with a trailing @domain
    removed: some accounts log in as netid@byu.edu. When that gives nothing
    usable, the Canvas user id stands in, and `source` says so.
    """
    raw = str(user.get(field) or "").strip()
    if "@" in raw:
        local, _, dom = raw.partition("@")
        raw = local if domain and dom.lower() == domain.lower() else ""
    if raw and NETID_OK.match(raw):
        return raw, field
    return str(user["id"]), "user_id"


def split_name(user):
    """(last, first) from a Canvas user, preferring sortable_name 'Last, First'."""
    sortable = str(user.get("sortable_name") or "")
    if ", " in sortable:
        last, first = sortable.split(", ", 1)
        return last, first
    parts = str(user.get("name") or user.get("short_name") or "").split()
    if len(parts) >= 2:
        return parts[-1], " ".join(parts[:-1])
    return (parts[0] if parts else "Unknown"), "Unknown"


def clean_name_part(s):
    """A name made safe for the last_first_netid_ prefix.

    Underscores would shift the fields anonymize.py splits on, and a dot can be
    taken for the start of an extension, so both become spaces.
    """
    s = unicodedata.normalize("NFC", str(s or ""))
    s = re.sub(r"[_./\\:\x00-\x1f]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or "Unknown"


def clean_filename(s, limit=180):
    """A student's own filename, safe to write: no separators, no leading dot."""
    s = unicodedata.normalize("NFC", str(s or ""))
    s = re.sub(r"[/\\:\x00-\x1f]", "-", s).strip().lstrip(".").strip()
    if not s:
        s = "submission"
    if len(s.encode("utf-8")) > limit:
        stem, dot, ext = s.rpartition(".")
        if not dot or len(ext) > 10:
            stem, ext = s, ""
        while len((stem + ("." + ext if ext else "")).encode("utf-8")) > limit:
            stem = stem[:-1]
        s = stem + ("." + ext if ext else "")
    return s


def local_filename(last, first, netid, theirs):
    """The Learning Suite download convention: last_first_netid_theirfilename."""
    return f"{clean_name_part(last)}_{clean_name_part(first)}_{netid}_{theirs}"


def parse_time(iso):
    """A Canvas timestamp ('2027-01-15T06:59:59Z') as epoch seconds, or None."""
    if not iso:
        return None
    try:
        return datetime.fromisoformat(str(iso).replace("Z", "+00:00")).timestamp()
    except ValueError:
        return None


def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------- #
# canvas.json -- what was pulled and pushed, per assignment folder
# --------------------------------------------------------------------------- #
def load_manifest(adir):
    p = Path(adir) / MANIFEST
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text())
    except ValueError as exc:
        raise CanvasError(f"could not read {p}: {exc}")


def save_manifest(adir, doc):
    """Write canvas.json atomically, owner-only: it ties NetIDs to Canvas ids."""
    p = Path(adir) / MANIFEST
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(doc, indent=2) + "\n")
    os.chmod(tmp, stat.S_IRUSR | stat.S_IWUSR)
    tmp.replace(p)
    return p


def user_for_netid(manifest, netid):
    """(user id, the newest pull record for that student) by NetID, or (None, None)."""
    want = str(netid).lower()
    for pull in reversed(manifest.get("pulls", [])):
        for uid, rec in pull.get("students", {}).items():
            if str(rec.get("netid", "")).lower() == want:
                return uid, rec
    return None, None


def pulled_record(manifest, folder, uid):
    """What a given pull recorded for a student, or None."""
    for pull in manifest.get("pulls", []):
        if pull.get("folder") == folder:
            return pull.get("students", {}).get(str(uid))
    return None
