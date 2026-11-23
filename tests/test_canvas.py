"""Canvas pull and push, against an in-memory Canvas.

There are no real submissions to test against, so fake_canvas.FakeCanvas plays
the API. The pull tests run the real anonymize.py mask on what was pulled, so
"the existing round logic handles it" is checked by that logic itself.
"""

import json
import os
import stat
import sys
from argparse import Namespace
from datetime import datetime
from pathlib import Path

import pytest

import anonymize
import canvas_api
import canvas_pull
import canvas_push
from canvas_api import Canvas, CanvasError
from fake_canvas import HOST, TOKEN, FakeCanvas

AID = 1532387
CFG = {
    "host": HOST, "course_id": 7, "term": None, "assignments": {},
    "assignments_file": None, "assignments_prefix": "", "netid_field": "login_id",
    "netid_domain": "byu.edu", "token_file": "unused", "workspace": None,
}


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
@pytest.fixture
def fake():
    f = FakeCanvas(course_id=7, page_size=2)       # small pages: pagination is exercised
    f.add_assignment(AID)
    f.add_student(AID, 101, "aaa11", "Avery, Ann")
    f.add_student(AID, 102, "bbb22@byu.edu", "Brooks, Ben")
    f.add_student(AID, 103, "ccc33", "Cruz, Cal")
    return f


def client(fake):
    return Canvas(HOST, TOKEN, session=fake, sleep=lambda s: None)


def do_pull(fake, adir):
    return canvas_pull.pull(client(fake), CFG, fake.assignments[AID], adir)


def do_mask(folder):
    args = Namespace(submissions_dir=folder, out=None, roster=None, ledger=None,
                     fresh=False, domain=None, band=6.0, no_redact=False)
    return anonymize.mask(args)


def ledger_changes(adir):
    led = json.loads((adir / "ledger.json").read_text())
    return {rec["masked"].split("_", 1)[1] + "@" + code: rec["change"]
            for code, st in led["students"].items() for rec in st["files"]}


def to_grade(adir):
    return (adir / "masked" / "TO_GRADE.txt").read_text().split()


def write_feedback(path, items):
    """A feedback document in the shape the template produces and sync reads."""
    import docx
    doc = docx.Document()
    earned = sum(e for _, _, e in items)
    possible = sum(p for _, p, _ in items)
    doc.add_paragraph("Homework Feedback: HW 1")
    doc.add_paragraph(f"Total Score: {earned} / {possible}")
    t = doc.add_table(rows=len(items) + 2, cols=3)
    for c, h in enumerate(("Rubric Item", "Possible", "Earned")):
        t.cell(0, c).text = h
    for i, (label, p, e) in enumerate(items, start=1):
        t.cell(i, 0).text, t.cell(i, 1).text, t.cell(i, 2).text = label, str(p), str(e)
    last = len(items) + 1
    t.cell(last, 0).text, t.cell(last, 1).text, t.cell(last, 2).text = (
        "TOTAL", str(possible), str(earned))
    doc.add_paragraph("Points Missed")
    for label, p, e in items:
        if e < p:
            doc.add_paragraph(label)
            doc.add_paragraph(f"Score: {e} / {p}")
            doc.add_paragraph("You used the wrong datum.")
    doc.save(str(path))


def write_scores(path, rows):
    """scores.xlsx after finalize: values in the Total column, as recalc leaves them."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["HW 1"])
    ws.append(["First Name", "Last Name", "Net ID", "Part 1", "Part 2", "Total", "Round"])
    ws.append([None, None, "Max Points", 15, 15, 30, None])
    for netid, p1, p2, rnd in rows:
        ws.append(["x", "y", netid, p1, p2, p1 + p2, rnd])
    wb.save(str(path))


def graded_round(adir, fake, scores, rnd=1):
    """Feedback documents and scores.xlsx for a round, as unmask + finalize leave them."""
    fb = adir / f"feedback{rnd}"
    fb.mkdir(exist_ok=True)
    rows = []
    manifest = json.loads((adir / "canvas.json").read_text())
    pull = next(p for p in manifest["pulls"] if p["folder"] == f"submissions{rnd}")
    for uid, (p1, p2) in scores.items():
        st = pull["students"][str(uid)]
        stem = Path(st["files"][0]["file"]).stem
        write_feedback(fb / f"{stem}_FEEDBACK.docx",
                       [("Part 1", 15, p1), ("Part 2", 15, p2)])
        rows.append((st["netid"], p1, p2, rnd))
    write_scores(adir / "scores.xlsx", rows)


def run_push(monkeypatch, capsys, fake, adir, *argv):
    monkeypatch.setattr(canvas_push, "read_token", lambda path: TOKEN)
    monkeypatch.setattr(canvas_push, "Canvas",
                        lambda host, token: Canvas(host, token, session=fake,
                                                   sleep=lambda s: None))
    monkeypatch.setattr(sys, "argv", ["canvas_push.py", str(adir), *argv])
    code = canvas_push.main()
    return code, capsys.readouterr().out


# --------------------------------------------------------------------------- #
# filenames and NetIDs
# --------------------------------------------------------------------------- #
def test_netid_comes_from_login_id():
    assert canvas_api.netid_for({"id": 1, "login_id": "abc12"}) == ("abc12", "login_id")
    # Some BYU accounts log in as netid@byu.edu.
    assert canvas_api.netid_for({"id": 1, "login_id": "abc12@byu.edu"}) == ("abc12", "login_id")
    # Anything else falls back to the Canvas user id, and says so.
    assert canvas_api.netid_for({"id": 77, "login_id": "x@gmail.com"}) == ("77", "user_id")
    assert canvas_api.netid_for({"id": 77, "login_id": ""}) == ("77", "user_id")
    assert canvas_api.netid_for({"id": 77, "login_id": "a_b"}) == ("77", "user_id")


@pytest.mark.parametrize("sortable,theirs", [
    ("Avery, Ann", "hw1.xlsx"),
    ("Van der Berg, Mary Ann", "HW 1 final.xlsx"),
    ("St. John, Jo_Ellen", "head calcs v2.final.xlsx"),
    ("Núñez, José", "tarea/uno.ipynb"),
])
def test_filename_round_trips_through_mask(sortable, theirs):
    last, first = canvas_api.split_name({"sortable_name": sortable})
    name = canvas_api.clean_filename(theirs)
    fname = canvas_api.local_filename(last, first, "abc12", name)
    m = anonymize.SUBMISSION_RE.match(Path(fname).stem)
    assert m, fname
    assert m["netid"] == "abc12"
    assert f"{m['rest']}{Path(fname).suffix}" == name
    assert "/" not in fname


def test_names_fall_back_without_sortable_name():
    assert canvas_api.split_name({"name": "Ann Marie Avery"}) == ("Avery", "Ann Marie")


def test_term_slug():
    assert canvas_api.term_slug("2027 Winter") == "2027-winter"
    assert canvas_api.term_slug("Fall 2026") == "2026-fall"
    assert canvas_api.term_slug("Default Term") is None


def test_plan_names_undoes_canvas_rename_only_against_an_earlier_name():
    rec = {"attachments": [{"id": 1, "display_name": "hw1-1.xlsx"},
                           {"id": 2, "display_name": "notes-2.pdf"}]}
    names = canvas_pull.plan_names(rec, prev_names={"hw1.xlsx"}, known={})
    assert names == {1: "hw1.xlsx", 2: "notes-2.pdf"}


# --------------------------------------------------------------------------- #
# pull
# --------------------------------------------------------------------------- #
def test_pull_multi_file_submission(fake, tmp_path):
    fake.submit(AID, 101, [("hw1.xlsx", b"one"), ("work.pdf", b"two"), ("hw1.xlsx", b"three")],
                "2027-01-15T06:59:00Z")
    fake.submit(AID, 102, [("Brooks hw.xlsx", b"ben")], "2027-01-16T08:00:00Z",
                seconds_late=3600)
    adir = tmp_path / "01_head"
    r = do_pull(fake, adir)

    assert r["status"] == "pulled" and r["folder"].name == "submissions1"
    files = sorted(p.name for p in (adir / "submissions1").iterdir())
    assert files == ["Avery_Ann_aaa11_hw1 (2).xlsx", "Avery_Ann_aaa11_hw1.xlsx",
                     "Avery_Ann_aaa11_work.pdf", "Brooks_Ben_bbb22_Brooks hw.xlsx"]
    assert (adir / "submissions1/Avery_Ann_aaa11_hw1 (2).xlsx").read_bytes() == b"three"

    # mask reads the submission time from the file's mtime.
    want = datetime.fromisoformat("2027-01-15T06:59:00+00:00").timestamp()
    assert (adir / "submissions1/Avery_Ann_aaa11_work.pdf").stat().st_mtime == want

    m = json.loads((adir / "canvas.json").read_text())
    assert stat.S_IMODE((adir / "canvas.json").stat().st_mode) == 0o600
    assert set(m["pulls"][0]["students"]) == {"101", "102"}
    assert m["pulls"][0]["students"]["102"]["late"] is True
    assert canvas_api.user_for_netid(m, "BBB22")[0] == "102"
    assert r["none"] == 1 and len(r["late"]) == 1

    # Token went only to the Canvas host, never in a URL, and nothing was written.
    assert not fake.writes()
    assert all(TOKEN not in url for _, url, _, _ in fake.calls)


def test_second_pull_feeds_the_round_logic(fake, tmp_path):
    adir = tmp_path / "01_head"
    fake.submit(AID, 101, [("hw1.py", b"print(1)")], "2027-01-15T06:00:00Z")
    fake.submit(AID, 102, [("a.py", b"a = 1"), ("b.py", b"b = 1")], "2027-01-15T07:00:00Z")
    do_pull(fake, adir)
    assert do_mask(adir / "submissions1") == 0
    assert len(to_grade(adir)) == 3

    # Ann resubmits, and Canvas renames her clashing upload to hw1-1.py.
    # Cal submits for the first time. Ben changes nothing.
    fake.submit(AID, 101, [("hw1-1.py", b"print(2)")], "2027-01-17T09:00:00Z")
    fake.submit(AID, 103, [("cal.py", b"c = 1")], "2027-01-18T09:00:00Z")
    downloads_before = sum(1 for c in fake.calls if "/files/" in c[1])
    r = do_pull(fake, adir)

    assert r["folder"].name == "submissions2"
    assert len(r["changes"]["new"]) == 1
    assert len(r["changes"]["resubmitted"]) == 1
    assert len(r["changes"]["unchanged"]) == 1
    # The whole class is there, as in a Learning Suite download, but Ben's
    # unchanged files were copied from submissions1, not fetched again.
    s2 = sorted(p.name for p in (adir / "submissions2").iterdir())
    assert s2 == ["Avery_Ann_aaa11_hw1.py", "Brooks_Ben_bbb22_a.py",
                  "Brooks_Ben_bbb22_b.py", "Cruz_Cal_ccc33_cal.py"]
    assert sum(1 for c in fake.calls if "/files/" in c[1]) - downloads_before == 2

    # The existing round logic sees exactly what happened.
    assert do_mask(adir / "submissions2") == 0
    changes = ledger_changes(adir)
    assert sorted(changes.values()) == ["new", "resubmitted", "unchanged", "unchanged"]
    assert [v for k, v in changes.items() if k.startswith("hw1.py")] == ["resubmitted"]
    graded = to_grade(adir)
    assert len(graded) == 2 and not any(n.endswith(("_a.py", "_b.py")) for n in graded)
    out_round2 = json.loads((adir / "ledger.json").read_text())["batches"]
    assert out_round2 == ["submissions1", "submissions2"]

    # Nothing new on Canvas: no submissions3.
    r = do_pull(fake, adir)
    assert r["status"] == "nothing_new"
    assert not (adir / "submissions3").exists()


def test_pull_replaces_a_folder_not_yet_masked(fake, tmp_path):
    adir = tmp_path / "01_head"
    fake.submit(AID, 101, [("hw1.py", b"1")], "2027-01-15T06:00:00Z")
    do_pull(fake, adir)
    fake.submit(AID, 102, [("b.py", b"2")], "2027-01-16T06:00:00Z")
    r = do_pull(fake, adir)
    assert r["folder"].name == "submissions1" and r["replaced"]
    assert not (adir / "submissions2").exists()
    assert len(list((adir / "submissions1").iterdir())) == 2
    assert len(json.loads((adir / "canvas.json").read_text())["pulls"]) == 1

    # A file nobody pulled is never deleted to make room.
    (adir / "submissions1" / "Hand_Added_xyz_extra.py").write_text("x")
    fake.submit(AID, 103, [("c.py", b"3")], "2027-01-17T06:00:00Z")
    with pytest.raises(CanvasError, match="no pull put there"):
        do_pull(fake, adir)
    assert (adir / "submissions1" / "Hand_Added_xyz_extra.py").exists()


def test_pull_with_no_submissions_writes_nothing(fake, tmp_path):
    adir = tmp_path / "01_head"
    r = do_pull(fake, adir)
    assert r["status"] == "none_submitted" and r["none"] == 3
    assert not adir.exists()


def test_pull_falls_back_to_user_id(fake, tmp_path):
    fake.subs[AID][103]["user"]["login_id"] = "cal@gmail.com"
    fake.submit(AID, 103, [("c.py", b"3")], "2027-01-17T06:00:00Z")
    r = do_pull(fake, tmp_path / "x")
    assert r["fallback"] == ["103"]
    assert [p.name for p in (tmp_path / "x/submissions1").iterdir()] == ["Cruz_Cal_103_c.py"]


# --------------------------------------------------------------------------- #
# assignment ids and the token
# --------------------------------------------------------------------------- #
def test_resolve_assignment(fake, tmp_path):
    fake.add_assignment(1532388, name="HW 2: Darcy's Law")
    fake.add_assignment(1532389, name="HW 3: Flow Nets")
    fake.add_assignment(1532390, name="HW 4: Flow Nets, Part 2")
    state = tmp_path / "state.json"
    state.write_text(json.dumps({"hw:01_head": AID, "group:homework": 5}))
    cfg = dict(CFG, assignments={"02_darcy": 42}, assignments_file=state,
               assignments_prefix="hw:")
    cv = client(fake)
    assert canvas_api.resolve_assignment(cv, cfg, "02_darcy")[0] == 42
    assert canvas_api.resolve_assignment(cv, cfg, "01_head") == (AID, "state.json (hw:01_head)")
    aid, how = canvas_api.resolve_assignment(cv, dict(CFG), "02_darcy")
    assert aid == 1532388 and how.startswith("matched by name")
    with pytest.raises(CanvasError, match="2 Canvas assignments match"):
        canvas_api.resolve_assignment(cv, dict(CFG), "04_flownets")
    assert canvas_api.resolve_assignment(cv, cfg, "x", override=9)[0] == 9


def test_canvas_settings_from_course_yml(tmp_path):
    course = {"lms": "canvas", "grading_workspace": "~/grading/ce544",
              "canvas": {"host": "https://byu.instructure.com/", "course_id": 40232,
                         "assignments_file": "canvas/state.json",
                         "assignments_prefix": "hw:"}}
    cfg = canvas_api.canvas_settings(course, tmp_path)
    assert cfg["host"] == "https://byu.instructure.com"
    assert cfg["assignments_file"] == tmp_path / "canvas/state.json"
    assert canvas_api.lms_of(course) == "canvas" and canvas_api.lms_of({}) == "learning_suite"
    with pytest.raises(CanvasError):
        canvas_api.canvas_settings({"canvas": {"host": "http://x", "course_id": 1}}, tmp_path)


def test_token_must_be_owner_only(tmp_path):
    t = tmp_path / "token"
    t.write_text("secret\n")
    os.chmod(t, 0o644)
    with pytest.raises(CanvasError, match="chmod 600") as e:
        canvas_api.read_token(t)
    assert "secret" not in str(e.value)
    os.chmod(t, 0o600)
    assert canvas_api.read_token(t) == "secret"


# --------------------------------------------------------------------------- #
# push
# --------------------------------------------------------------------------- #
@pytest.fixture
def graded(fake, tmp_path):
    """Round 1 pulled and graded: Ann 27/30 on time, Ben 24/30 a day and a half late."""
    fake.submit(AID, 101, [("hw1.xlsx", b"ann")], "2027-01-15T06:00:00Z")
    fake.submit(AID, 102, [("hw1.xlsx", b"ben")], "2027-01-16T18:00:00Z",
                seconds_late=36 * 3600)
    adir = tmp_path / "01_head"
    do_pull(fake, adir)
    graded_round(adir, fake, {101: (12, 15), 102: (15, 9)})
    return adir


def test_push_dry_run_lists_and_sends_nothing(fake, graded, monkeypatch, capsys):
    code, out = run_push(monkeypatch, capsys, fake, graded)
    assert code == 0
    assert "DRY RUN" in out and "SEND (2)" in out
    assert "aaa11" in out and "27 / 30" in out
    assert "bbb22" in out and "24 / 30" in out and "late 1.5 day(s)" in out
    assert "2 to send. Nothing was sent." in out
    assert not fake.writes()
    assert TOKEN not in out


def test_push_sends_raw_scores_and_attaches_feedback(fake, graded, monkeypatch, capsys):
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert code == 0, out
    puts = {c[1].rsplit("/", 1)[1]: dict(c[3]) for c in fake.puts()}
    # Raw: Ben's 24 goes up as 24, and Canvas takes its own late deduction.
    assert puts["101"]["submission[posted_grade]"] == "27"
    assert puts["102"]["submission[posted_grade]"] == "24"
    assert fake.subs[AID][102]["entered_score"] == 24 and fake.subs[AID][102]["score"] < 24
    assert all("comment[file_ids][]" in p for p in puts.values())
    assert sorted(n for n, _ in fake.uploads.values()) == [
        "Avery_Ann_aaa11_hw1_FEEDBACK.docx", "Brooks_Ben_bbb22_hw1_FEEDBACK.docx"]
    # Nothing posts grades: no call touches a posting endpoint or field.
    assert not any("post" in str(c[3]).lower().replace("posted_grade", "")
                   for c in fake.writes())
    m = json.loads((graded / "canvas.json").read_text())
    assert m["pushes"]["102"]["score"] == 24 and len(m["push_log"]) == 2

    # Running it again sends nothing twice.
    n = len(fake.writes())
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert "UP TO DATE (2)" in out and len(fake.writes()) == n


def test_push_will_not_overwrite_a_score_changed_in_canvas(fake, graded, monkeypatch, capsys):
    run_push(monkeypatch, capsys, fake, graded, "--send")
    fake.set_score(AID, 101, 29)                    # Norm bumps Ann in the gradebook
    graded_round(graded, fake, {101: (13, 15), 102: (15, 9)})   # and her document too

    n = len(fake.puts())
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert code == 1
    assert "CHANGED IN CANVAS (1)" in out and "Canvas has 29, the last push sent 27" in out
    assert len(fake.puts()) == n                    # Ann untouched, Ben up to date
    assert fake.subs[AID][101]["entered_score"] == 29

    code, out = run_push(monkeypatch, capsys, fake, graded, "--send", "--overwrite", "aaa11")
    assert "OVERWRITES" in out
    assert fake.subs[AID][101]["entered_score"] == 28


def test_push_will_not_overwrite_a_score_entered_only_in_canvas(fake, graded, monkeypatch,
                                                                capsys):
    fake.set_score(AID, 102, 20)
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert "Canvas already has 20, entered there" in out
    assert [c[1].rsplit("/", 1)[1] for c in fake.puts()] == ["101"]


def test_push_refuses_while_grades_post_automatically(fake, graded, monkeypatch, capsys):
    fake.assignments[AID]["post_manually"] = False
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert code == 1
    assert "CANNOT SEND" in out and "posts grades automatically" in out
    assert not fake.writes()


def test_push_skips_a_student_who_resubmitted_after_the_pull(fake, graded, monkeypatch, capsys):
    fake.submit(AID, 101, [("hw1.xlsx", b"ann v2")], "2027-01-17T06:00:00Z")
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert "RESUBMITTED SINCE THE PULL (1)" in out
    assert "graded attempt 1, Canvas now has attempt 2" in out
    assert [c[1].rsplit("/", 1)[1] for c in fake.puts()] == ["102"]


def test_push_blocks_when_finalize_is_stale(fake, graded, monkeypatch, capsys):
    write_scores(graded / "scores.xlsx", [("aaa11", 12, 15, 1), ("bbb22", 15, 8, 1)])
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert "scores.xlsx says 23, the document says 24 -- finalize again" in out
    assert [c[1].rsplit("/", 1)[1] for c in fake.puts()] == ["101"]


def test_push_blocks_masked_documents(fake, graded, monkeypatch, capsys):
    fb = graded / "feedback1"
    for f in fb.iterdir():
        f.rename(fb / ("S-ABC123_" + f.name.split("_", 3)[3]))
        break
    code, out = run_push(monkeypatch, capsys, fake, graded)
    assert "still masked -- run unmask" in out


def test_push_round_two_sends_only_that_round(fake, graded, monkeypatch, capsys):
    run_push(monkeypatch, capsys, fake, graded, "--send")
    # Ann resubmits; round 2 pulls and regrades her alone.
    fake.submit(AID, 101, [("hw1-1.xlsx", b"ann v2")], "2027-01-17T06:00:00Z")
    r = do_pull(fake, graded)
    # submissions1 was graded (feedback1 exists), so even unmasked it is a round
    # and is never rewritten.
    assert r["folder"].name == "submissions2" and not r["replaced"]
    graded_round(graded, fake, {101: (15, 15)}, rnd=2)
    write_scores(graded / "scores.xlsx", [("aaa11", 15, 15, 2), ("bbb22", 15, 9, 1)])

    n = len(fake.puts())
    code, out = run_push(monkeypatch, capsys, fake, graded, "--send")
    assert code == 0, out
    assert "feedback2/" in out and "SEND (1)" in out
    new = fake.puts()[n:]
    assert [c[1].rsplit("/", 1)[1] for c in new] == ["101"]
    assert dict(new[0][3])["submission[posted_grade]"] == "30"


def test_check_deps_wants_a_token_only_for_canvas(tmp_path):
    import check_deps
    t = tmp_path / "token"
    assert not [m for m in check_deps.check(lms="learning_suite", token_file=t)
                if "token" in m["name"]]
    assert [m["name"] for m in check_deps.check(lms="canvas", token_file=t)
            if "token" in m["name"]] == ["Canvas API token"]
    t.write_text("x")
    os.chmod(t, 0o644)
    assert "chmod 600" in check_deps.check_canvas_token(t)["install"]
    os.chmod(t, 0o600)
    assert check_deps.check_canvas_token(t) is None
