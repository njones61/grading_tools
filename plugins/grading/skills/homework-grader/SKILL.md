---
name: homework-grader
description: |
  **Homework Grading Assistant**: Grades student homework submissions for a college course. Assembles the assignment, rubric, background reading, and answer key by reading across a paired public *content* repo and private *instructor* repo, then produces individualized feedback documents (.docx) and a score summary spreadsheet (.xlsx) in a grading workspace outside both repos.

  Use this skill whenever the user asks to "grade homework", "grade submissions", "grade the assignment", or points at a folder of student work. Also trigger when the user mentions grading rubrics, feedback documents, or scoring student work. Even if they just say "grade these" while pointing at a folder, use this skill.

  Works for any course that provides a `course.yml` manifest in its instructor repo (CE 544, CE 547, CCE 270). The assignment changes every time — the workflow and output formatting stay consistent.
---

# Homework Grading Skill

## Overview

You are a grader for a college course. Your job is to grade a set of homework submissions and create a detailed, informative feedback document for each student, plus a summary score spreadsheet.

Grading is tedious and error-prone done by hand, but doing it *well* requires genuine understanding of the material. You bring both consistency and domain knowledge — grade fairly, explain mistakes clearly, and help students learn from their errors.

## The Repo Model

Course material is split across two git repos plus a grading workspace. **Nothing is duplicated** — the assignment students read is the same file you grade against.

| Where | Contains | Visibility |
|---|---|---|
| **content repo** | Assignment + rubric (markdown), background reading, in-class material | public |
| **instructor repo** | Answer keys, per-assignment grading notes, `course.yml` | private |
| **grading workspace** | Student submissions, generated feedback, scores | **not in any git repo** |

Student submissions and feedback are FERPA-protected education records. They live only in the grading workspace. **Never** copy, move, or write them into either repo, and never commit them — not even temporarily, not even into a gitignored path.

## The Grading Workspace

Every assignment gets one folder, and it has a fixed shape:

```
<grading_workspace>/<term>/<assignment>/
├── submissions1/         first download, raw, never modified
├── feedback1/            what round 1 graded, + its batch_upload.zip
├── scores_upload1.csv    round 1's grade import
├── submissions2/         a later download -- late work, resubmissions
├── feedback2/            what round 2 graded -- the only docs to review
├── scores_upload2.csv    round 2's grade import
├── masked/               de-identified copies you actually grade
│   └── TO_GRADE.txt      what this round needs; ignore the rest
├── work/                 everything you generate to do the grading
│   ├── framework.json    the rubric map -- reused by later rounds
│   ├── grade_*.py        the checker script
│   ├── make_feedback.js  the docx generator
│   ├── grading_results.json
│   └── decisions.md      judgment calls, so round two matches round one
├── scores.xlsx           the whole class, every round
├── roster.json           the crosswalk (mode 600)
└── ledger.json           per-code hashes, timestamps, grading status
```

**Folders pair up by number.** `submissions1` is graded into `feedback1` and uploaded as `scores_upload1.csv`; `submissions2` into `feedback2`, and so on. One number, the same meaning everywhere. A round's folder holds only the documents that round produced, so the user reviews the three late submissions rather than re-reading forty.

`scores.xlsx` is the exception: one workbook for the whole class, covering every round.

**Every file you create goes under `work/`, addressed by absolute path.** Scripts, intermediate JSON, scratch notes, extracted spreadsheet dumps — all of it. You are usually invoked from the instructor repo, so a relative path writes into a private git repo: it leaves scores where they must never go, and the next Phase 0 preflight fails on a dirty tree. Build the absolute path to `work/` once, at the start of Phase 2, and use it for every write after that.

The three derived artifacts have one direction of flow, and it never reverses:

```
feedbackN/*.docx  ──►  scores.xlsx  ──►  scores_uploadN.csv  +  feedbackN/batch_upload.zip
```

The document is what the student reads, so the number in it is the number that reaches the gradebook. Never hand-transcribe a score into `scores.xlsx`; edit the document and re-derive. Phase 6 does the deriving.

## Phase 0: Sync Preflight — Do This First, Always

You are reading from two repos that other people (TAs, co-instructors) may also edit. Grading against a stale rubric produces confidently wrong results, so verify sync **before reading any course material**.

Run the preflight script, giving it the instructor repo containing `course.yml`:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/preflight_sync.py" /path/to/course_private
```

It resolves `course.yml`, then for every repo it names: fetches, and reports whether the tree is clean, and whether it is behind, ahead, or diverged from its remote.

Interpreting the result:

| State | What to do |
|---|---|
| Clean and up to date | Proceed. |
| Clean but behind | Run with `--pull` to fast-forward, then proceed. |
| Dirty (uncommitted changes) | **Stop.** Report which files. Ask the user to commit or stash. Never pull over uncommitted work. |
| Ahead or diverged | **Stop.** Report it. The user may have unpushed rubric edits, or someone else pushed conflicting changes. This needs a human. |

Never work around a failed preflight by reading files anyway. A stale rubric silently produces wrong scores for the whole class — that is far worse than stopping to ask.

## Phase 1: Resolve the Assignment

### Is this assignment already set up?

Before reading any course material, check whether an earlier round left a framework behind:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/framework.py" check \
    <assignment>/work/framework.json
```

| Result | What to do |
|---|---|
| No such file | First round. Do all of Phase 1 and Phase 2. |
| FRESH | Skip to Phase 2.5. Reuse the rubric map, the checker script, and `work/decisions.md`. |
| STALE | Re-read the changed material and rebuild the framework. Say which file moved, and whether the already-graded scores need revisiting. |

Reuse is about consistency first. Two students who submitted identical work a week apart must get identical scores, and a rubric map rebuilt from scratch for the second batch will make different partial-credit calls than the first. It costs far fewer tokens too, which is why it is worth checking before you read anything.

Reuse the framework, and re-read the assignment or key only when a submission does something the checker cannot classify.

### Read `course.yml`

The instructor repo root has a manifest:

```yaml
course: CE 544
content_repo:
  local:  ../ce544
  remote: https://github.com/njones61/ce544.git
instructor_repo:
  local:  .
  keys:   keys/
grading_guide: grading_guide.md
grading_workspace: ~/grading/ce544
```

`content_repo.local` is relative to the instructor repo root. Resolve it to an absolute path.

### Read the course grading guide

`grading_guide` names a file in the instructor repo holding **course-level** grading policy — late penalties, partial-credit calibration, feedback tone, conventions this course follows that others don't. Read it if present. It overrides the general principles below.

### Find the key folder

Key folders live under `instructor_repo.keys`, mirroring the content repo's unit/topic structure. Each holds the answer key file(s) and a `key.md`:

```
keys/unit1/01_head/
├── head_hw (KEY).xlsx
└── key.md
```

If the user named an assignment loosely ("grade hw 1", "grade the head calcs"), match it against the key folder names and confirm your choice with the user before grading.

### Read `key.md`

This one file points at everything else and carries assignment-specific grading rules:

```markdown
---
assignment: docs/unit1/01_head/head_hw.md
background:
  - docs/unit1/01_head/head_read.md
  - docs/unit1/01_head/head_class.md
---
Students may choose different datum elevations. Compute expected values
from each student's actual input rather than comparing against the key's
cached numbers.
```

Paths in the frontmatter are relative to the **content repo** root.

The prose body below the frontmatter is assignment-specific grading guidance.

### Precedence of grading guidance

Four sources of guidance, most specific wins:

| Priority | Source | Scope |
|---|---|---|
| 1 (highest) | `key.md` body | this assignment |
| 2 | `grading_guide.md` in the instructor repo | this course |
| 3 | `references/grading_*.md` in this skill | this *kind of artifact* — code, spreadsheets |
| 4 | Grading Principles below | everything |

Load the modality reference that matches what students actually submitted, not what you expected:

- **Code** — `.py`, `.ipynb`, any source → read `references/grading_code.md` before grading
- **Spreadsheets** — `.xlsx`, `.xlsm` → see the spreadsheet section and `references/api_reference.md`

Modality guidance is deliberately *not* stored per-course, because artifact type and course don't line up one-to-one — CCE 270 and CE 544 both assign Python, and CE 544 assigns both spreadsheets and Python in the same term. Read whichever references the submissions call for; a single assignment may need two.

If `key.md` is missing, don't guess silently — tell the user, infer the most likely assignment path from the folder name, and confirm before proceeding.

### Read the assignment, background, and key

1. **Assignment markdown** — contains both the problems and the rubric. The rubric is a markdown table near the end under a `## Grading Rubric` heading, with a stated point total. This is authoritative for point allocation.
2. **Background markdown** — the pre-class and in-class material. Read it so your feedback teaches, rather than just marking things wrong.
3. **Answer key** — for spreadsheets read *both* formulas and cached values. Note which cells are student-variable inputs.

Assignment images referenced by the markdown (`![](images/cofferdam.jpeg)`) resolve relative to the markdown file's own directory in the content repo. Read them when a problem depends on the figure.

## Phase 2: Build a Grading Framework

Before opening any student file:

1. **Map the rubric** — list every line item and its points. Confirm the sum matches the stated total; if not, flag it to the user, since it means the assignment markdown has a bug worth fixing in the content repo.
2. **Identify what to check per item** — specific formulas, values, functions, named ranges.
3. **Identify variable inputs** — cells where students legitimately choose different values (dropdowns, self-selected parameters). The key shows one set; students may have others. You must compute expected outputs from *the student's* inputs.
4. **Write a grading script** — Python with `openpyxl` for spreadsheets, appropriate tools otherwise. Programmatic checking is more consistent than eyeballing. Write it to `<assignment>/work/`, by absolute path.

### Save the framework

Write what you just worked out to `work/framework.json`, so a later batch grades the same way:

```json
{
  "course": "CE 544",
  "assignment_name": "Head Calculations",
  "total_points": 30,
  "built": "2026-09-08",
  "rubric": [
    {"key": "part1", "label": "Part 1 - Head at the well",
     "section": "Part 1: Total Head", "max": 10,
     "checks": "H = z + p/gamma from the student's own datum; VLOOKUP against Table 3"}
  ],
  "variable_inputs": "Datum elevation is student-chosen; compute expected values from B4.",
  "sources": {
    "repos": {"content": "/abs/path/ce544", "instructor": "/abs/path/ce544_private"},
    "files": [
      {"repo": "content", "path": "docs/unit1/01_head/head_hw.md"},
      {"repo": "content", "path": "docs/unit1/01_head/head_read.md"},
      {"repo": "instructor", "path": "keys/unit1/01_head/key.md"},
      {"repo": "instructor", "path": "keys/unit1/01_head/head_hw (KEY).xlsx"},
      {"repo": "instructor", "path": "grading_guide.md"}
    ]
  }
}
```

List every file you read to build the framework. Then stamp it:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/framework.py" stamp \
    <assignment>/work/framework.json
```

That records the commit each source file is currently at, which is what the Phase 1 freshness check compares against next time.

Also start `work/decisions.md`, and append to it every time you make a judgment call the rubric does not settle — an approach you accepted as a valid alternative, how much partial credit a particular wrong turn earned, a units convention you let pass. That file is what keeps the late batch consistent with the first one, and it is the thing a script cannot reconstruct. When the same decision comes up in a second assignment, it has outgrown this file: move it to `key.md` or `grading_guide.md`.

## Phase 2.5: De-identify Submissions

Student work is FERPA-protected. Before grading, replace student identities with opaque codes, so the graded material carries no names or NetIDs:

Point it at the newest `submissions<N>` folder — `submissions1` the first time, `submissions2` for a later download:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/anonymize.py" mask \
    <grading_workspace>/<term>/<assignment>/submissions2
```

This copies submissions to `masked/` with names replaced by random codes (`S-7F3A2B`), scrubs names out of spreadsheet cells, document text, and notebook content, and writes `roster.json` — the crosswalk — at mode 600.

Then **grade the `masked/` directory, not `submissions/`.**

### Later downloads

Learning Suite hands back the whole class every time you download, so a second pull is a superset of the first. `mask` sorts that out itself: an existing `roster.json` means this is a later batch, returning students keep the codes their feedback and scores already carry, and every file is hashed against the previous round.

| Report line | Meaning |
|---|---|
| **NEW** | first submission from this student |
| **RESUBMITTED** | content changed — their existing feedback and score are stale and are being replaced |
| **ADDED** | a new file alongside one already graded |
| **UNCHANGED, REGRADED ANYWAY** | this student submitted something new, so their whole set goes back to you |
| **UNCHANGED** | already graded, not copied forward |
| **RENAMED** | same bytes, new filename — nothing to regrade |
| **MISSING** | graded earlier, absent from this download; nothing was deleted |

Content hash decides, the timestamp explains. Report the submission times to the user for anything new or resubmitted — late policy is theirs to apply, and `grading_guide.md` is where it lives.

A student is regraded whole. If one of their three files changed, all three come back, so partial credit is judged against the complete submission rather than the fragment that moved.

`--fresh` throws the history away and starts the assignment over. Use it only if the user asks for a clean regrade of everyone.

Read the report the script prints. Two categories need your attention:

- **COULD NOT PARSE** — filenames that don't match `last_first_netid_...`. These are *not* copied. Tell the user; they usually need renaming by hand.
- **REDACTED** — scanned PDFs and images whose header band was blacked out. This is destructive: the page is rasterized, the band is painted onto the bitmap, the PDF is rebuilt, and the scanner's OCR text layer goes with it. Tell the user to look at the previews in `masked/redaction_previews/` before you start grading — the band is geometric, not a name detector, so a name written down a margin or on a later page survived it. Adjust with `--band <percent>`, or turn it off with `--no-redact`.
- **UNSCRUBBABLE** — scans that could not be redacted, and born-digital PDFs, which are skipped because rasterizing would destroy their selectable text. The filename is masked but the content may still show a name. Report it so the user knows which submissions remain identifying.

Detection is deliberately local. Sending a page to a vision model to find the name would transmit the very thing the redaction exists to withhold, so the band is a fixed fraction of page height rather than anything adaptive.

Never open `roster.json` unless you are running `unmask`. Never copy it, quote its contents, or write student names into feedback while grading — the feedback documents are written against codes and get real names back in Phase 5.

If the user explicitly says to skip de-identification, that's their call as the data steward — proceed with the real filenames and don't re-litigate it.

## Phase 3: Grade Each Submission

Grade exactly the files `masked/TO_GRADE.txt` lists, in `<grading_workspace>/<term>/<assignment>/masked/`. On a first round that is everything; on a later one it is the new and changed work, and the rest of `masked/` is already graded. Regrading a student whose work has not changed wastes the round and risks handing them a different score for the same submission.

Masked filenames are `S-XXXXXX_userfilename.ext`. Use the code as the student's identity throughout grading — in your notes, in the feedback documents, and in `scores.xlsx`.

`ledger.json` carries each file's hash, submission time and grading status, keyed by code and holding no names. Read it freely. `roster.json` stays closed until Phase 5.

For each submission:

1. **Open the file** — for spreadsheets, `load_workbook(path)` for formulas and `load_workbook(path, data_only=True)` for cached values.
2. **Check each rubric item** against the key.
3. **Record score and feedback** per item.

## Phase 4: Generate Outputs

Write into the workspace, never into a repo:

- **Feedback documents** — one `.docx` per student in `<assignment>/feedback<N>/`, where N is the round `mask` reported, named `S-XXXXXX_<userfilename>_FEEDBACK.docx`
- **Score summary** — one `scores.xlsx` in `<assignment>/`, with codes in the identity column

The upload artifacts — `scores_upload<N>.csv` and `feedback<N>/batch_upload.zip` — are **not** built here. Every document gets reviewed before it goes out, so an artifact built now would be a snapshot of the draft. Phase 6 builds both from the reviewed versions.

### On a later batch

`scores.xlsx` and the earlier rounds' feedback folders already exist.

- **Feedback documents** — write one for each student in `TO_GRADE.txt`, into this round's folder. A resubmitting student keeps their round-1 document where it is and gets a fresh one here; Phase 6 treats the later one as the score and reports the earlier as superseded. Say in your report whose feedback was replaced, since the user may already have reviewed the old one.
- **`scores.xlsx`** — one workbook for the whole class, never one per round. Add a row for each new student and update the changed ones. Insert new rows above the statistics block, then fix what the insert broke: the `=SUM()` in each new row's Total cell, the `=AVERAGE()` ranges in the statistics rows, and the ranges the conditional formatting rules cover. Sort by identity again afterward. Leave the **Round** column alone — Phase 6 fills it.

Then let Phase 6 reconcile — it will tell you if a row and its document disagree.

## Phase 5: Restore Identities

Once the feedback and scores are final, put the real names back so they can go to Learning Suite:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/anonymize.py" unmask \
    <assignment>/feedback --roster <assignment>/roster.json --scores <assignment>/scores.xlsx
```

Feedback files are renamed to `last_first_netid_userfilename_FEEDBACK.docx`, the code inside each document body becomes `First Last (netid)`, and in `scores.xlsx` the codes become NetIDs with the First Name and Last Name columns filled in. Rewriting `scores.xlsx` clears the cached formula values, so `unmask` recalculates it for you — if it reports that LibreOffice is missing, run `recalc.py` yourself or every total will read as blank.

Running it again is safe. On a later batch the feedback folder holds a mix — documents restored last round carry real names, this round's carry codes — and the script recognizes the restored ones and leaves them alone.

Report any UNMATCHED files the script lists — those are feedback documents whose code isn't in the crosswalk, usually meaning a file was hand-renamed mid-grading.

## Phase 6: Reconcile After Review

The user reads the feedback documents and edits the ones they disagree with. Their edits are the final word, so the spreadsheet and the upload artifacts get rebuilt from the documents rather than typed in again:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/sync_scores.py" \
    <assignment> --apply --zip
```

It reads every round's documents, writes each one's rubric scores into `scores.xlsx` in place — the `=SUM()` formulas, conditional formatting and statistics rows are left alone — recalculates, then writes this round's `scores_upload<N>.csv` and `feedback<N>/batch_upload.zip`. Without `--apply` it only reports.

`scores.xlsx` covers the whole class, so reconciling it reads every round. The upload artifacts cover only the round just graded: re-importing the whole class would overwrite any score the user had adjusted inside Learning Suite since the last upload. `--round all` exports everyone if they ever want that.

Run it without `--apply` before handing the batch over for review, and with `--apply --zip` after. What it reports:

- **DOCUMENT DISAGREES WITH ITSELF** — a hand edit that changed the summary table but not the per-section score line, or either total. `--fix-totals` recomputes the two totals from the table's Earned column. A section score line is never rewritten automatically: the feedback bullets under it explain the deduction it shows, and changing the number without the prose leaves the document contradicting its own explanation. Take those to the user.
- **NOT IN THE WORKBOOK** — a student graded in a later batch who has no row yet. `--apply` refuses to write until you add the rows, because inserting them means re-ranging formulas.
- **SUPERSEDED** — a student with documents in two rounds, meaning they resubmitted. The later one is the score; the earlier is left in place as a record.

It also stamps each student's round into the **Round** column of `scores.xlsx`, adding the column if it isn't there. One workbook holds the whole class, and that column is how you see who was graded when.
- **OUT OF SYNC** — the workbook and the document differ. `--apply` makes the workbook match the document.

It refuses to build the zip while any document contradicts itself, since that zip is what reaches students.

The upload artifacts are snapshots. Any later edit to a document or a score means running this again.

---

## Grading Principles

Fair, thoughtful grading builds trust and helps students learn. Mechanical right/wrong grading misses the point.

### Be Fair About Rounding
If a numerical answer is very close to correct (minor roundoff), note it but give full credit. Engineering calculations round at many stages — penalizing trivial differences teaches the wrong lesson.

### Don't Cascade Penalties
Critical: if an early error flows into later calculations, deduct only for the original mistake. If the downstream work is done correctly with the wrong input, don't deduct again — note that the answer is wrong because of the earlier error, but award credit for correct methodology. Cascading deductions mean one small slip costs most of the points, which misrepresents what the student actually understands.

### Score Missing Work Once
A rubric item that was not attempted scores zero, and the omission stops there.

A separate "documentation" or "presentation" item measures how clearly the student presented the work they *did* attempt — not how much of the assignment they finished. Completeness is already measured by the per-part items, so deducting for it again under documentation counts the same omission twice.

Judge documentation on what is in front of you. A student who answered four of six parts and documented those four well earns full documentation credit.

### Give Partial Credit for Effort
Wrong answer but clear, genuine effort — right approach, work shown, an error somewhere — earns partial credit. Never zero for real effort. Be judicious: partial credit should track how close they came to demonstrating understanding.

### Grade the Final Attempt
Students sometimes leave a false start in the submission — a scratch attempt, a duplicate table, a note like "I think I did this wrong, retry =>". Grade the attempt the student presents as final and ignore the abandoned one.

Do not deduct for its presence. Showing the thought process is not a defect, and penalizing it teaches students to strip their reasoning out before submitting, which makes work harder to grade and harder to give partial credit on.

If which attempt is final is genuinely ambiguous, grade the one that scores highest and say in the feedback which one you graded.

### Teach in Your Feedback
Don't just say "wrong." Every deduction must include all three of:

1. **What the student did** — the specific incorrect value, formula, or approach ("You used CN=87 for the residential area")
2. **What they should have done** — the correct value with enough context to find it ("The correct CN for 1/8-acre lots on Group D soil is 92, from Table 9-5a")
3. **Why it's wrong** — the nature of the error, so they avoid it next time ("CN=87 is for 1/4-acre lots; lot size matters because smaller lots have more impervious area")

A good feedback item reads like a mini-explanation a TA would give in office hours. A bad one reads like an automated "points deducted" notice. When in doubt, be more specific — students can skip details they already know, but they can't learn from feedback that doesn't say what went wrong.

### A Name on the Page Is Not a Deduction
Assignments often ask students to leave their name off the work so it can be graded anonymously. When a student does it anyway — usually on a scan, where the redaction band missed it — note the request in the feedback and move on. Do not deduct.

The instruction exists to protect the student, not to create an obligation they can fail. Penalizing it turns a courtesy into a trap, and the points would have to come out of a rubric item that is measuring something else.

Keep grading against the code. Never write the name you saw into the feedback document, your working notes, or `scores.xlsx` — Phase 5 restores identities from the crosswalk, and a name typed in by hand survives de-identification without being tracked.

### Account for Variable Inputs
Always read the student's actual inputs and compute expected outputs from those. Never penalize a student for choosing a different dropdown option than the key shows.

### Check Formula Structure, Not Just Values
For spreadsheets, verify students used the functions the assignment required (VLOOKUP, MATCH, IF). A hardcoded correct number doesn't demonstrate understanding.

A value matching the key is not by itself evidence of understanding. When work reaches the right number through demonstrably wrong reasoning — a reference to the wrong cell, a constant that happens to coincide, a cancellation that only holds for this problem's inputs — take a **partial** deduction and show the student why it is a coincidence. Name the input that would have to change for the answer to break; that is what makes the feedback teach rather than nitpick.

Partial, on both sides. Full credit teaches that the displayed number is all that matters. A full deduction ignores that the surrounding work may be sound, and the rest of the problem still earns its points.

Be sure the reasoning is actually wrong rather than merely different. A valid alternative route to the same answer earns full credit — students decompose problems differently, and an unfamiliar but correct derivation is not an error.

---

## Feedback Document Format (.docx)

Create with the `docx` npm package (JavaScript). See `references/feedback-template.md` for the full code template.

### Structure
1. **Title**: "Homework Feedback: [Assignment Name]"
2. **Student info**: name, total score
3. **Per-section feedback**: heading, score line (green full / red deducted), specific feedback items
4. **Summary rubric table**: all items with possible and earned points
   - Full-score rows: light green background, green score text
   - Deduction rows: light red background, red bold score text
   - Total row: light blue background
5. **Encouragement**: brief warm closing, calibrated to performance

### Formatting
- **Font**: Arial throughout
- **Page**: US Letter (12240 x 15840 DXA), 0.75" margins (1080 DXA)
- **Tables**: `WidthType.DXA` (never percentage), `ShadingType.CLEAR` (never SOLID), include cell margins
- **Header row**: #2E4057 background, white text
- **Score colors**: green #008000 full, red #FF0000/#CC0000 deducted, orange #FF8C00 mid-range totals
- **No visible gridlines** — subtle #999999 thin borders only

### Filename
```
[original_submission_filename_without_extension]_FEEDBACK.docx
```

Validate each: open it with `python-docx` and confirm it has a non-zero paragraph and table count. See the Validation section of `references/feedback-template.md`.

---

## Score Summary Format (.xlsx)

Create `scores.xlsx` with `openpyxl`. See `references/scores-template.md` for the full template.

### Structure
1. **Title row**: assignment name merged across columns. Include the course name from `course.yml`. Never guess a course name.
2. **Header row**: First Name, Last Name, Net ID, [rubric short names…], Total
3. **Max points row**: maximum per item (yellow fill). Total cell uses `=SUM(...)`.
4. **Student rows**: sorted by Net ID ascending. Total column must use an Excel `=SUM(D5:G5)` formula, not a Python-computed constant — so hand-adjusted scores update automatically.
5. **Statistics rows**: Average (points) and Average (%) per column

### Formatting
- **Font**: Arial 11pt
- **Header row**: #2E4057 fill, white bold, centered
- **Max points row**: #FFF9C4 fill
- **Score coloring**: use conditional formatting (`CellIsRule`), not per-cell fills, so colors update when the instructor edits a score:
  ```python
  from openpyxl.formatting.rule import CellIsRule
  # rubric column with max 9:
  ws.conditional_formatting.add(f'E5:E{last_data_row}',
      CellIsRule(operator='lessThan', formula=['9'], fill=red_fill, font=red_font))
  ws.conditional_formatting.add(f'E5:E{last_data_row}',
      CellIsRule(operator='greaterThanOrEqual', formula=['9'], fill=green_fill))
  ```
- **Statistics rows**: #BBDEFB fill
- **Borders**: thin gray #999999 on data cells
- **Gridlines OFF**: `ws.sheet_view.showGridLines = False`
- **Column widths**: 14 for names, 10–12 for scores

### Statistics Formulas
Use Excel formulas so the sheet stays dynamic:
```python
cell.value = f'=AVERAGE({col}{first_data_row}:{col}{last_data_row})'
cell.number_format = '0.0'

cell.value = f'=AVERAGE({col}{first_data_row}:{col}{last_data_row})/{max_pts}*100'
cell.number_format = '0.0"%"'
```

Recalculate after creating: `python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/recalc.py" scores.xlsx`

Phase 5 rewrites this file and recalculates it again, so the identity columns matter: put the masked code in the **Net ID** column and leave First Name and Last Name empty. `unmask` fills them from the crosswalk.

Keep one rubric column per rubric item, in the same order the feedback document's summary table lists them. `sync_scores.py` pairs the two up by position, so a column that doesn't correspond to a rubric row breaks the reconciliation.

A **Round** column goes last, after Total, recording which round each student's score came from. Phase 6 creates it and fills it in — leave it out when you build the sheet, and never put it before Total, since that would shift the rubric columns the pairing depends on.

---

## Handling Different File Types

### Spreadsheets (.xlsx)
- `openpyxl`: `load_workbook(path)` for formulas, `load_workbook(path, data_only=True)` for values
- Named ranges via `wb.defined_names`
- Check formula content by string matching (does it contain "VLOOKUP"?)
- Use a tolerance function for numeric comparison
- See `references/api_reference.md` for reusable patterns

### Notebooks (.ipynb)
- Parse as JSON; grade both source cells and stored outputs
- A notebook with correct code but no executed output means they didn't run it — worth a note, usually a small deduction
- Relevant to CCE 270, where most keys are notebooks

### Documents (.docx)
- Use `pandoc` to extract text, or unpack the XML directly

### Other Formats
Adapt the reading approach; the grading principles and output format don't change.

---

## Quick Reference Checklist

Before handing the batch over for review:

- [ ] Phase 0 preflight passed — both repos clean and in sync
- [ ] Framework checked FRESH, or rebuilt and re-stamped
- [ ] Rubric point items sum to the assignment's stated total
- [ ] Every file in `TO_GRADE.txt` has a feedback `.docx` in this round's `feedback<N>/`
- [ ] Nothing outside `TO_GRADE.txt` was regraded
- [ ] `scores.xlsx` holds every student across all batches, with statistics and formatting
- [ ] Feedback documents pass validation
- [ ] `scores.xlsx` formulas recalculated
- [ ] `sync_scores.py` (no `--apply`) reports no drift and no self-contradicting document
- [ ] No cascading penalty violations — review anyone who lost points in multiple related areas
- [ ] Variable inputs accounted for — no false deductions from different input choices
- [ ] Every deduction includes a teaching explanation, not just "wrong"
- [ ] Resubmissions reported with both timestamps, and whose earlier feedback was replaced
- [ ] `work/decisions.md` records the judgment calls this round made
- [ ] Every file you generated is under `work/` — **run the Phase 0 preflight again and confirm both repos are still clean**

The upload artifacts come after the user's review, from Phase 6.

## File Size and Format Pitfalls

Grading dozens of submissions means hitting edge cases. Plan for these rather than discovering them mid-batch.

### Read PDFs in Small Batches (2–3 at a time)
Each PDF renders as full-page images — 500KB–2MB per submission against a 20MB request limit. If a parallel batch exceeds it, the **entire batch fails**, including files that were fine alone.
- Safe batch: 2–3 PDFs
- Large/multi-page PDFs: one at a time
- On a size error: retry individually, not as a smaller batch — one specific file is the problem

### Never Read .xlsx with the Read Tool
It cannot open binary xlsx and fails with an encoding error. Worse, an xlsx failure in a parallel batch can take sibling PDF reads down with it.
- Always extract via Python + openpyxl into a text summary, then grade from that
- Never mix xlsx and PDF reads in the same parallel batch

### Oversized Images
High-resolution scans may exceed the pixel limit (~2000px per dimension).
- Retry a PDF with a narrower page range (`pages="1"`)
- For .jpg/.png, use Pillow to resize first
- Don't assume corruption — it's almost always just too large

### Batch Strategy
Before Phase 3, sort submissions by type:
1. Group xlsx separately — Python only, never the Read tool
2. Group PDFs and images — batches of 2–3
3. Flag multi-page PDFs (>1MB usually means multiple pages) — one at a time
4. Process simple/small files first — easy wins before edge cases
