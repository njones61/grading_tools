"""An in-memory Canvas that answers the handful of REST calls the scripts make.

It stands in for requests.Session: Canvas(host, token, session=FakeCanvas())
sends every call here instead of to the network. Every call is recorded with
its headers, so the tests can check what was written, and that the token only
ever went to the Canvas host.
"""

import copy
import json
import math
from urllib.parse import parse_qs, urlsplit

HOST = "https://canvas.test"
UPLOAD_HOST = "https://upload.test"
TOKEN = "tok-DO-NOT-LEAK-123"


class Resp:
    def __init__(self, status=200, body=None, content=None, headers=None, links=None):
        self.status_code = status
        if content is None:
            content = json.dumps(body).encode() if body is not None else b""
        self.content = content
        self.headers = headers or {}
        self.links = links or {}

    @property
    def ok(self):
        return 200 <= self.status_code < 400

    @property
    def text(self):
        return self.content.decode("utf-8", errors="replace")

    def json(self):
        return json.loads(self.content)

    def iter_content(self, n):
        for i in range(0, len(self.content), n):
            yield self.content[i:i + n]

    def close(self):
        pass


class FakeCanvas:
    def __init__(self, course_id=7, page_size=2):
        self.course_id = course_id
        self.page_size = page_size
        self.term = "2027 Winter"
        self.assignments = {}
        self.subs = {}          # assignment id -> {user id: submission}
        self.blobs = {}         # attachment id -> bytes
        self.uploads = {}       # file id -> (name, bytes)
        self.calls = []         # (method, url, headers, data)
        self._next_att = 500
        self._next_file = 9000

    # ---- setting up the course ---------------------------------------------
    def add_assignment(self, aid, name="HW 1: Hydraulic Head", points=30.0,
                       post_manually=True, **extra):
        self.assignments[aid] = {
            "id": aid, "name": name, "points_possible": points,
            "post_manually": post_manually, "grading_type": "points",
            "submission_types": ["online_upload"], "published": True, **extra}
        self.subs.setdefault(aid, {})

    def add_student(self, aid, uid, login_id, sortable_name):
        last, first = sortable_name.split(", ")
        self.subs[aid][uid] = {
            "id": uid * 10, "user_id": uid, "assignment_id": aid,
            "workflow_state": "unsubmitted", "attempt": None, "submitted_at": None,
            "attachments": [], "late": False, "seconds_late": 0, "excused": None,
            "entered_score": None, "score": None, "points_deducted": None,
            "grade_matches_current_submission": True,
            "user": {"id": uid, "login_id": login_id, "sortable_name": sortable_name,
                     "name": f"{first} {last}", "short_name": f"{first} {last}"}}

    def submit(self, aid, uid, files, submitted_at, seconds_late=0):
        """A new attempt with these (display name, bytes) files."""
        sub = self.subs[aid][uid]
        atts = []
        for name, data in files:
            self._next_att += 1
            self.blobs[self._next_att] = data
            atts.append({"id": self._next_att, "display_name": name, "filename": name,
                         "size": len(data), "content-type": "application/octet-stream",
                         "url": f"{HOST}/files/{self._next_att}/download?verifier=v{self._next_att}"})
        sub.update(attachments=atts, attempt=(sub["attempt"] or 0) + 1,
                   submitted_at=submitted_at, late=seconds_late > 0,
                   seconds_late=seconds_late, submission_type="online_upload",
                   workflow_state="submitted",
                   grade_matches_current_submission=sub["entered_score"] is None)

    def set_score(self, aid, uid, entered):
        """Someone changes the grade by hand in Canvas."""
        self._grade(aid, uid, entered)

    def _grade(self, aid, uid, entered):
        sub = self.subs[aid][uid]
        points = self.assignments[aid]["points_possible"]
        deducted = 0.0
        if entered is not None and sub["seconds_late"]:
            days = math.ceil(sub["seconds_late"] / 86400)
            deducted = min(0.1 * days, 0.5) * points
        sub.update(entered_score=entered, points_deducted=deducted if entered is not None else None,
                   score=None if entered is None else entered - deducted,
                   workflow_state="graded", grade_matches_current_submission=True)

    # ---- what the tests inspect ---------------------------------------------
    def writes(self):
        return [c for c in self.calls if c[0] != "GET"]

    def puts(self):
        return [c for c in self.calls if c[0] == "PUT"]

    # ---- the transport ----------------------------------------------------------
    def request(self, method, url, params=None, data=None, files=None, headers=None,
                allow_redirects=True, stream=False, timeout=None):
        headers = headers or {}
        self.calls.append((method, url, dict(headers), data))
        assert TOKEN not in url, "token in a URL"
        sp = urlsplit(url)
        origin = f"{sp.scheme}://{sp.netloc}"
        # Merge the URL's own query with params, keeping repeated keys as lists,
        # the way Canvas sees them.
        query = {k: list(v) for k, v in parse_qs(sp.query).items()}
        for k, v in (params or {}).items():
            query[k] = list(v) if isinstance(v, list) else [str(v)]

        if origin == UPLOAD_HOST:
            assert "Authorization" not in headers, "token sent to the upload host"
            (field, (name, fh, _ctype)), = files.items()
            self._next_file += 1
            self.uploads[self._next_file] = (name, fh.read())
            return Resp(201, {"id": self._next_file, "display_name": name})

        assert origin == HOST, f"unexpected host {origin}"
        assert headers.get("Authorization") == f"Bearer {TOKEN}"
        parts = sp.path.strip("/").split("/")

        if parts[0] == "files":
            return Resp(200, content=self.blobs[int(parts[1])])

        assert parts[:2] == ["api", "v1"], sp.path
        p = parts[2:]
        if method == "GET" and p == ["courses", str(self.course_id)]:
            return Resp(200, {"id": self.course_id, "term": {"name": self.term}})
        if method == "GET" and p == ["courses", str(self.course_id), "assignments"]:
            return self._page(sp.path, list(self.assignments.values()), query)
        if p[:3] == ["courses", str(self.course_id), "assignments"] and len(p) >= 4:
            aid = int(p[3])
            if aid not in self.assignments:
                return Resp(404, {"errors": [{"message": "not found"}]})
            if method == "GET" and len(p) == 4:
                return Resp(200, self.assignments[aid])
            if method == "GET" and p[4:] == ["submissions"]:
                rows = [copy.deepcopy(s) for _, s in sorted(self.subs[aid].items())]
                if "user" not in query.get("include[]", []):
                    for r in rows:
                        r.pop("user", None)
                return self._page(sp.path, rows, query)
            if method == "POST" and p[4] == "submissions" and p[6:] == ["comments", "files"]:
                return Resp(200, {"upload_url": f"{UPLOAD_HOST}/up/{p[5]}",
                                  "upload_params": {"key": "abc"}, "file_param": "file"})
            if method == "PUT" and p[4] == "submissions" and len(p) == 6:
                uid = int(p[5])
                fields = dict((k, v) for k, v in data if not k.endswith("[]"))
                if "submission[posted_grade]" in fields:
                    self._grade(aid, uid, float(fields["submission[posted_grade]"]))
                if "comment[text_comment]" in fields:
                    self.subs[aid][uid].setdefault("comments", []).append(
                        (fields["comment[text_comment]"],
                         [v for k, v in data if k == "comment[file_ids][]"]))
                out = copy.deepcopy(self.subs[aid][uid])
                out.pop("user", None)
                return Resp(200, out)
        raise AssertionError(f"no fake route for {method} {sp.path}")

    def _page(self, path, rows, query):
        """One page, with a next link that repeats the query, as Canvas's does."""
        page = int(query.get("page", ["1"])[0])
        n = self.page_size
        chunk = rows[(page - 1) * n: page * n]
        links = {}
        if page * n < len(rows):
            keep = [(k, v) for k, vs in query.items() if k not in ("page", "per_page")
                    for v in vs]
            q = "&".join(f"{k}={v}" for k, v in keep + [("page", page + 1), ("per_page", n)])
            links["next"] = {"url": f"{HOST}{path}?{q}"}
        return Resp(200, chunk, links=links)
