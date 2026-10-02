# grading_tools

Private Claude Code plugin marketplace holding the teaching tools shared across BYU courses. Currently one plugin: **grading**, which provides the `homework-grader` skill.

**This file is the grading manual for every course.** The course repos hold keys and course policy; how the grading workflow itself works is documented here, once, in the same repo as the scripts it describes.

| Course | Content repo | Instructor repo | Workspace |
|---|---|---|---|
| CE 544 | `njones61/ce544` | `njones61/ce544_private` | `~/grading/ce544` |
| CE 547 | `njones61/ce547` | `njones61/ce547_private` | `~/grading/ce547` |
| CCE 270 | `byu-cce270/content` | `byu-cce270/instructor` | `~/grading/cce270` |

---

## One-Time Setup

### 1. Install the plugin

```
/plugin marketplace add njones61/grading_tools
/plugin install grading@grading_tools
```

Claude Code clones this repo with your existing `gh` credentials, so a private repo works without extra setup.

**`/plugin install` enables the plugin for the current project only.** It writes `.claude/settings.json` into whatever repo you ran it from, which limits the skill to that one course and leaves an untracked file that fails the Phase 0 preflight. To get it everywhere, put this in `~/.claude/settings.json` and delete any `.claude/settings.json` the install left behind:

```json
"enabledPlugins": { "grading@grading_tools": true }
```

### Updating later

Two steps, and both are needed:

```
/plugin marketplace update grading_tools
/plugin update grading@grading_tools
```

Then apply it to the running session:

```
/reload-plugins
```

Restart Claude Code if that doesn't pick it up.

The first command refreshes the catalog; the second refreshes the copy you actually run. Doing only the first leaves you on whatever you installed originally, with the marketplace sitting current beside it. That is what happened to everyone between v1.5.0 and v2.0.0.

To see which version you are on:

```bash
ls ~/.claude/plugins/cache/grading_tools/grading/
```

The cache keeps every version ever installed, so more than one directory here is normal; the skill runs from the newest.

### 2. Check out the repos side by side

`course.yml` points at the content repo with a relative path, so they must be siblings:

```bash
cd ~/python_projects
git clone https://github.com/njones61/ce544.git
git clone https://github.com/njones61/ce544_private.git
```

### 3. Install the dependencies

The scripts shell out to these.

```bash
brew install --cask libreoffice      # recalculates scores.xlsx formulas
brew install poppler                 # renders scanned PDFs for redaction
pip install openpyxl python-docx     # reads submissions, validates feedback
pip install pillow pypdf             # blacks out the name band on scans
cd ~/grading && npm install docx     # builds the feedback .docx files
```

Check them at any time:

```bash
python "$GT/check_deps.py" --workspace ~/grading
```

The Phase 0 preflight runs this before it fetches a repo, and refuses to grade until everything is present. Two of these otherwise fail in ways that look like success: without the redaction stack a scanned submission is copied through with the student's handwritten name still on it, and without LibreOffice every total in `scores.xlsx` comes out blank at Phase 5, once the class has already been graded.

Install `docx` in `~/grading` rather than globally: node resolves `node_modules` by walking up from the script, so one install there covers every course and assignment beneath it.

### 4. Create the grading workspace

Student work never goes in a git repo:

```bash
mkdir -p ~/grading/ce544
```

Back this directory up — Time Machine, or a Drive folder outside any repo.

---

## The Repo Model

Course material lives in three places. Nothing is duplicated.

```
  content repo (public)         instructor repo (private)      grading workspace
  ────────────────────          ─────────────────────────      ─────────────────
  assignment + rubric           answer keys                    submissions
  background reading            course.yml                     feedback
  in-class material             grading_guide.md               scores.xlsx
                                key.md per assignment          roster.json
```

The assignment students read *is* the file the grader grades against — a markdown file with the rubric inline. There is no PDF to print, download, or keep in sync.

**Student submissions and feedback never enter a git repo.** They are FERPA-protected education records and live only in the grading workspace.

---

## Grading an Assignment

### 1. Download submissions from Learning Suite

```
~/grading/ce544/2026-fall/01_head/
└── submissions1/
    ├── Last_First_netid_theirfilename.xlsx
    └── ...
```

Number the folder. A later download goes in `submissions2`, and the numbers pair up all the way through: `submissions1` is graded into `feedback1` and uploaded as `scores_upload1.csv`.

Keep Learning Suite's `last_first_netid_...` filenames — the de-identification step parses them.

### 2. Ask Claude Code to grade it

From the instructor repo:

```
grade 01_head
```

The skill runs the workflow: sync check, assemble the assignment and key, de-identify, grade, generate feedback, restore names. It stops there, with everything ready for you to read — the upload files come after your review, in step 4.

### 3. Review the feedback

```
~/grading/ce544/2026-fall/01_head/
├── feedback1/      one .docx per student, real names restored
└── scores.xlsx     class summary with statistics
```

Read the documents and fix the ones you disagree with. **Edit the document, not the spreadsheet** — the document is what the student reads, and step 4 copies its scores into `scores.xlsx` for you.

Each document opens with the score table, then has a short section for each item where the student lost points. Items with full credit get no section.

### 4. Build the upload artifacts

Ask for it the same way you asked for the grading:

```
finalize 01_head
```

"sync the scores", "I'm done reviewing", "I bumped a couple of scores" all mean the same thing. Nothing is regraded — the documents on disk are already the answer, and this reads them.

It writes their scores into `scores.xlsx`, recalculates, and produces:

```
├── feedback1/batch_upload.zip   ← upload this, not the files individually
└── scores_upload1.csv           ← Net ID + score, for the Grades import
```

It also catches the mistake that is easy to make by hand: bumping a score in the summary table and leaving the section line or the totals showing the old number. It reports those rather than guessing, and can recompute the totals from the table if you ask. To see what it found before anything is written, say `check 01_head` instead.

Both artifacts are snapshots. Edit another document afterward and say `finalize 01_head` again.

By hand, if you'd rather:

```bash
python "$GT/sync_scores.py" ~/grading/ce544/2026-fall/01_head --apply --zip
```

### 5. Late work and resubmissions

Download again into the next numbered folder:

```
~/grading/ce544/2026-fall/01_head/
├── submissions1/       the first download
└── submissions2/       the second one
```

Then ask for the same thing — `grade 01_head`. Learning Suite hands back the whole class on every download, so the grader hashes each file against the previous round and only grades what actually changed. Students keep the codes their existing feedback carries, and the report tells you exactly what happened:

| | |
|---|---|
| **NEW** | first submission from this student |
| **RESUBMITTED** | content changed — their old feedback and score are being replaced |
| **ADDED** | a new file alongside one already graded |
| **UNCHANGED** | already graded, skipped |
| **UNCHANGED, REGRADED ANYWAY** | this student submitted something else, so their whole set comes back |
| **RENAMED** | same content under a new filename, nothing to regrade |
| **MISSING** | graded earlier, not in this download; nothing deleted |

Submission timestamps come with it, so you can see what was late. No penalty is applied automatically — that policy lives in your `grading_guide.md` and stays your call.

Nothing is ever deleted, and the first download stays exactly as it was.

Round 2's work lands in its own folders, so review means reading the three late submissions rather than the whole class again:

```
~/grading/ce544/2026-fall/01_head/
├── feedback1/          round 1, already uploaded
├── feedback2/          ← the only documents to read this time
├── scores_upload1.csv  round 1's import, already done
├── scores_upload2.csv  ← the only students to import this time
├── round2_summary.txt  who was new, and whose grade is changing
└── scores.xlsx         the whole class, with a Round column
```

`round2_summary.txt` is the one to read first. The folder shows who was graded; it can't show why, and a student who resubmitted is not the same as a new one — their grade is already posted, so their row is a correction to something they have seen:

```
Round 2 — graded 8 Sep 2026

NEW (1) — first submission from this student
  Park, Rui (rpark3)     submitted 7 Sep 23:14

RESUBMITTED (1) — already graded and uploaded; these scores are changing
  Ngo, Alex (ango7)      was 1 Sep 10:02, now 7 Sep 22:40
```

Then `finalize 01_head` again. It reconciles `scores.xlsx` across every round — one workbook for the class, with a **Round** column recording who was graded when — and writes upload artifacts covering only the round just graded. Importing all forty students again would overwrite any score you had adjusted inside Learning Suite since the first upload; `--round all` exports everyone if you ever want that.

If someone resubmitted, they have a document in both folders. The later one is the score, and `sync_scores.py` says so rather than letting the stale one count.

---

## What Happens Under the Hood

You can run any phase by hand if something goes wrong. These commands need the plugin's script directory:

```bash
GT=$(ls -d ~/.claude/plugins/cache/grading_tools/grading/*/skills/homework-grader/scripts | sort -V | tail -1)
echo "$GT"    # should end in .../<version>/skills/homework-grader/scripts
```

The cache keeps every version ever installed, so this selects the newest deliberately. Do not search for the directory by name and take the first hit — that returns the marketplace clone and every stale version in arbitrary order, and will happily run a year-old script.

(Claude Code knows this path as `${CLAUDE_PLUGIN_ROOT}` and resolves it itself. You only need `$GT` when running a step by hand.)

### Phase 0 — Preflight

```bash
python "$GT/preflight_sync.py" ~/python_projects/ce544_private --pull
```

Checks the tooling, then the repos, and stops on either.

Dependencies come first and are reported together, with the install command and the phase each would have broken. Then both repos must be clean and current, or you risk grading against a stale rubric. `--pull` fast-forwards a repo that is clean and merely behind; it **stops** rather than pulling if a repo has uncommitted changes or has diverged — those need you.

### Phase 1 — Resolve the assignment

First it checks whether an earlier round already worked this assignment out:

```bash
python "$GT/framework.py" check ~/grading/ce544/2026-fall/01_head/work/framework.json
```

The first pass writes its rubric map, its checker script, and a log of the judgment calls it made to `work/`, and stamps the commit each source file was at. A later round reuses all of it, so the late submissions are graded the same way the first batch was — rebuilding the framework from scratch would produce different partial-credit calls on identical work. If the assignment or key has been edited in between, `check` reports STALE and names the file, and the framework gets rebuilt.

Then it reads `course.yml`, finds the key folder, and reads its `key.md`:

```markdown
---
assignment: docs/unit1/01_head/head_hw.md
background:
  - docs/unit1/01_head/head_read.md
---
Assignment-specific grading notes go here.
```

Paths in the frontmatter are relative to the **content repo** root.

### Phase 2 — Build the grading framework

Maps the rubric, works out what to check for each item, writes the checker script, and saves all of it to `work/framework.json` with the commit each source file was at. That is what a later round reuses.

### Phase 2.5 — De-identify

```bash
python "$GT/anonymize.py" mask ~/grading/ce544/2026-fall/01_head/submissions1
```

Each student gets a random opaque code (`S-7F3A2B`). Names and NetIDs are scrubbed from spreadsheet cells, document text, and notebook content. The crosswalk lands in `roster.json` at mode 600. Grading then runs against `masked/`, never `submissions1/`.

Point it at a later download and it becomes a diff: returning students keep their codes, every file is hashed against the previous round, and only new or changed work is copied forward. `masked/TO_GRADE.txt` lists exactly what that round covers. `ledger.json` holds the hashes, submission times and status — keyed by code, no names, so it is safe to read. `--fresh` throws the history away and regrades everyone.

**Read the three categories it reports:**

- **COULD NOT PARSE** — filenames not matching `last_first_netid_...`. These are *not* copied. Rename them by hand and re-run.
- **REDACTED** — scans and images whose header band was blacked out. This is destructive: the page is rasterized, the band is painted onto the bitmap, and the scanner's OCR text layer goes with it. **Look at `masked/redaction_previews/` before grading** — the band is geometry, not a name detector, so a name written down a margin or on a later page is still there. `--band <percent>` adjusts the height; `--no-redact` turns it off.
- **UNSCRUBBABLE** — what redaction could not cover, plus born-digital PDFs, which are skipped so they keep their selectable text. These may still identify the student.
- **NAME IS ALSO A COURSE WORD** — a student surnamed Wells, Head or Brooks in a course whose assignment uses that word. Scrubbing the bare token would replace it throughout their own submission and hand the grader prose about "the S-7F3A2B screened at 40 ft", so it is left in place. The full name is still scrubbed; a lone surname on the page is not.

Detection is deliberately local. Sending a page to a vision model to locate the name would transmit the very thing the redaction exists to withhold.

### Phases 3–4 — Grade and generate

Grades the files `TO_GRADE.txt` lists, writes one feedback `.docx` per student into that round's `feedback<N>/`, and adds their rows to `scores.xlsx`. Everything it generates along the way — the checker script, the intermediate JSON, the notes — goes in `work/`, so nothing lands in a git repo.

### Phase 5 — Restore identities

```bash
python "$GT/anonymize.py" unmask ~/grading/ce544/2026-fall/01_head/feedback1 \
    --roster ~/grading/ce544/2026-fall/01_head/roster.json \
    --scores ~/grading/ce544/2026-fall/01_head/scores.xlsx
```

It:

1. renames feedback files to `last_first_netid_...`
2. replaces the code *inside* each document with `First Last (netid)`
3. turns codes in `scores.xlsx` into NetIDs and fills the First/Last Name columns
4. recalculates `scores.xlsx` — rewriting it clears the cached formula values, so skipping this leaves every total blank
5. writes `round<N>_summary.txt` — who this round covered, and whose grade is being corrected rather than posted for the first time

Safe to run again on a later round: documents restored earlier are recognized and left alone.

### Phase 6 — Reconcile after your review

```bash
python "$GT/sync_scores.py" ~/grading/ce544/2026-fall/01_head --apply --zip
```

This is what `finalize 01_head` runs. The feedback document is the score of record, so it reads them and derives everything else. Without `--apply` it only reports. It:

1. checks each document against itself — the section score lines, the summary table, and the two totals, which a hand edit usually leaves disagreeing, and flags a deduction with no section explaining it (`--fix-totals` recomputes the totals from the table; a section line is never rewritten automatically, since the bullets under it explain the number it shows)
2. writes each document's scores into `scores.xlsx`, leaving the formulas, conditional formatting and statistics rows untouched
3. stamps the **Round** column, adding it after Total if it isn't there
4. recalculates, then writes `scores_upload<N>.csv` and `feedback<N>/batch_upload.zip`

The CSV and zip cover only the round just graded. It refuses to build the zip while any document contradicts itself, since that zip is what reaches students.

### Phase 7 — Debrief

```bash
python "$GT/debrief_stats.py" ~/grading/ce544/2026-fall/01_head
```

The per-item numbers behind the write-up. The prose is the grader's work — see [Debriefing an Assignment](#debriefing-an-assignment) above for what it produces and where it lands.

---

## Debriefing an Assignment

Once the scores are final:

```
debrief 01_head
```

You get a prose write-up of how the assignment went: the distribution, where students struggled, and — the part worth having — where the *assignment* is at fault rather than the students.

That distinction rests on a number rather than an impression. For each rubric item the grader checks whether the strong students missed it too, ranking the class by everything *except* that item so it can't place students in its own halves. An item that costs real points but that the top half missed about as often as the bottom half isn't measuring understanding, and the instruction is the first suspect:

```
  Rubric item        max   mean   lost   top½   bot½      r   reading
  P3                   8    6.4    20%    91%    70%   0.74   genuinely hard, or under-taught
  P5                  10    6.0    40%    61%    59%   0.02   SUSPECT THE WORDING
```

The numbers say where to look. What went wrong comes from reading the feedback documents, so every claim in the debrief names the specific misconception and quotes the instruction that invited it — "fourteen of thirty-eight measured head from the well screen rather than the datum", not "students found Part 2 hard". Where the wording is at fault it drafts the replacement sentence for you to take or reject; it never edits the content repo.

Under about eight students it says the class was too small to tell rather than reporting noise. And an assignment that worked is a valid finding — it's meant to say so, not to manufacture a problem every time.

**It writes two copies:**

```
~/grading/ce544/2026-fall/01_head/debrief.md     with the rest of the grading
~/python_projects/ce544_private/keys/unit1/01_head/debrief.md
```

The second is the one that matters. The workspace is organized by term, so you'd never open `2026-fall/` again — but you will open the key folder the next time you teach this. It holds aggregate statistics and no student data, and it's written fresh rather than copied out of the workspace, so nothing that was ever a student record enters a repo. It arrives uncommitted; committing it is yours.

---

## Where Guidance Lives

Four tiers, most specific wins:

| Tier | Location | Scope |
|---|---|---|
| 1 | `key.md` body in the instructor repo | one assignment |
| 2 | `grading_guide.md` in the instructor repo | one course |
| 3 | `plugins/grading/skills/homework-grader/references/grading_*.md` | one artifact type |
| 4 | `SKILL.md` grading principles | everything |

Tier 3 is deliberately not per-course: CCE 270 and CE 544 both assign Python, and CE 544 assigns both spreadsheets and Python in the same term. How to grade Python is a Python fact, not a course fact.

**When you find yourself correcting the grader twice on the same point, write it down at the right tier.** That is the maintenance loop this structure exists to support.

---

## Rules That Do Not Bend

- **Student submissions and feedback never go in a git repo.** They are FERPA-protected education records and live only in the grading workspace.
- **`roster.json` is never shared or committed.** It maps codes back to real students; it is the one file that re-identifies everything.
- **Never pull over uncommitted changes.** The preflight refuses to, and so should you.

---

## Releasing a New Version

**Bump `version` in `plugins/grading/.claude-plugin/plugin.json` on every release.** Users are pinned to the version they installed, so an unchanged version string means nobody gets the change — the marketplace updates, the installed copy does not, and the only symptom is that a documented feature appears not to exist.

Use semver against the *user's* workflow: a bump to the minor version for new capability, to the major for anything that changes what they do by hand or removes an artifact they relied on. v2.0.0 was major because `unmask` stopped producing the upload zip and CSV.

After committing the bump:

```bash
claude plugin validate .
claude plugin tag plugins/grading --push
```

`validate` checks the marketplace manifest. `tag` creates `grading--v<version>`, refusing if `plugin.json` and the marketplace entry disagree or the tree is dirty — which is the check that would have caught this being forgotten.

Then tell people to run both update commands and `/reload-plugins`.

### Why not drop the version field entirely

Claude Code falls back to the commit SHA when `plugin.json` declares no `version`, so every push would reach everyone with no bump to remember — which would make the failure above impossible.

The cost is that the install cache is then named by SHA rather than by version. The `$GT` lookup below picks the newest directory with `sort -V`, which is meaningful for `2.0.0` and meaningless for a hash, and there would be no version number to ask a colleague for when something goes wrong. Explicit versions stay, and `claude plugin tag` guards the bump.

## Adding a Course

1. Create a private instructor repo beside the public content repo.
2. Add `course.yml` naming both repos and the grading workspace.
3. Add `grading_guide.md` for course-level policy.
4. Put keys under `keys/<unit>/<topic>/`, each with a `key.md` pointing at its assignment.

---

## Layout

```
grading_tools/
├── .claude-plugin/marketplace.json     catalog
└── plugins/grading/
    ├── .claude-plugin/plugin.json
    └── skills/homework-grader/
        ├── SKILL.md                    workflow + grading principles
        ├── references/
        │   ├── grading_code.md          Python, notebooks, scripts
        │   ├── api_reference.md         openpyxl patterns
        │   ├── feedback-template.md     docx template
        │   └── scores-template.md       xlsx template
        └── scripts/
            ├── preflight_sync.py        verify tooling installed, repos in sync
            ├── check_deps.py            verify every shelled-out tool is present
            ├── framework.py             stamp / check the cached rubric map
            ├── anonymize.py             mask / unmask, and diff later downloads
            ├── redact_scan.py           black out the header band on scans
            ├── sync_scores.py           feedback docs -> scores.xlsx -> uploads
            ├── debrief_stats.py         per-item stats for the debrief
            ├── recalc.py                recalculate xlsx formulas
            ├── package_feedback.py      build feedbackN/batch_upload.zip
            └── export_upload_csv.py     build scores_uploadN.csv
```
