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
4. **Write a grading script** — Python with `openpyxl` for spreadsheets, appropriate tools otherwise. Programmatic checking is more consistent than eyeballing. Keep the script in the workspace, not in a repo.

## Phase 2.5: De-identify Submissions

Student work is FERPA-protected. Before grading, replace student identities with opaque codes, so the graded material carries no names or NetIDs:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/anonymize.py" mask \
    <grading_workspace>/<term>/<assignment>/submissions
```

This copies submissions to `masked/` with names replaced by random codes (`S-7F3A2B`), scrubs names out of spreadsheet cells, document text, and notebook content, and writes `roster.json` — the crosswalk — at mode 600.

Then **grade the `masked/` directory, not `submissions/`.**

Read the report the script prints. Two categories need your attention:

- **COULD NOT PARSE** — filenames that don't match `last_first_netid_...`. These are *not* copied. Tell the user; they usually need renaming by hand.
- **UNSCRUBBABLE** — scanned PDFs and images. The filename is masked but the content may still show a handwritten name. Nothing can fix this automatically. Report it so the user knows which submissions remain identifying.

Never open `roster.json` unless you are running `unmask`. Never copy it, quote its contents, or write student names into feedback while grading — the feedback documents are written against codes and get real names back in Phase 5.

If the user explicitly says to skip de-identification, that's their call as the data steward — proceed with the real filenames and don't re-litigate it.

## Phase 3: Grade Each Submission

Grade the files in `<grading_workspace>/<term>/<assignment>/masked/`.

Masked filenames are `S-XXXXXX_userfilename.ext`. Use the code as the student's identity throughout grading — in your notes, in the feedback documents, and in `scores.xlsx`.

For each submission:

1. **Open the file** — for spreadsheets, `load_workbook(path)` for formulas and `load_workbook(path, data_only=True)` for cached values.
2. **Check each rubric item** against the key.
3. **Record score and feedback** per item.

## Phase 4: Generate Outputs

Write into the workspace, never into a repo:

- **Feedback documents** — one `.docx` per student in `<assignment>/feedback/`, named `S-XXXXXX_<userfilename>_FEEDBACK.docx`
- **Score summary** — one `scores.xlsx` in `<assignment>/`, with codes in the identity column

Phase 5 adds `feedback/batch_upload.zip` for Learning Suite. If you graded without de-identifying, `unmask` never runs, so build it yourself:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/package_feedback.py" <assignment>/feedback
```

Rebuild it after any edit to a feedback document — the zip is a snapshot, not a live view.

## Phase 5: Restore Identities

Once the feedback and scores are final, put the real names back so they can go to Learning Suite:

```bash
python "${CLAUDE_PLUGIN_ROOT}/skills/homework-grader/scripts/anonymize.py" unmask \
    <assignment>/feedback --roster <assignment>/roster.json --scores <assignment>/scores.xlsx
```

Feedback files are renamed to `last_first_netid_userfilename_FEEDBACK.docx`, the code inside each document body becomes `First Last (netid)`, and in `scores.xlsx` the codes become NetIDs with the First Name and Last Name columns filled in. The restored documents are then zipped into `feedback/batch_upload.zip`, which is what Learning Suite ingests -- one upload instead of one per student. Rewriting `scores.xlsx` clears the cached formula values, so `unmask` recalculates it for you — if it reports that LibreOffice is missing, run `recalc.py` yourself before uploading or every total will read as blank.

Do this **last**, after any regrading. If you regrade after unmasking, re-run `mask` rather than editing de-identified and identified files side by side.

Report any UNMATCHED files the script lists — those are feedback documents whose code isn't in the crosswalk, usually meaning a file was hand-renamed mid-grading.

---

## Grading Principles

Fair, thoughtful grading builds trust and helps students learn. Mechanical right/wrong grading misses the point.

### Be Fair About Rounding
If a numerical answer is very close to correct (minor roundoff), note it but give full credit. Engineering calculations round at many stages — penalizing trivial differences teaches the wrong lesson.

### Don't Cascade Penalties
Critical: if an early error flows into later calculations, deduct only for the original mistake. If the downstream work is done correctly with the wrong input, don't deduct again — note that the answer is wrong because of the earlier error, but award credit for correct methodology. Cascading deductions mean one small slip costs most of the points, which misrepresents what the student actually understands.

### Give Partial Credit for Effort
Wrong answer but clear, genuine effort — right approach, work shown, an error somewhere — earns partial credit. Never zero for real effort. Be judicious: partial credit should track how close they came to demonstrating understanding.

### Teach in Your Feedback
Don't just say "wrong." Every deduction must include all three of:

1. **What the student did** — the specific incorrect value, formula, or approach ("You used CN=87 for the residential area")
2. **What they should have done** — the correct value with enough context to find it ("The correct CN for 1/8-acre lots on Group D soil is 92, from Table 9-5a")
3. **Why it's wrong** — the nature of the error, so they avoid it next time ("CN=87 is for 1/4-acre lots; lot size matters because smaller lots have more impervious area")

A good feedback item reads like a mini-explanation a TA would give in office hours. A bad one reads like an automated "points deducted" notice. When in doubt, be more specific — students can skip details they already know, but they can't learn from feedback that doesn't say what went wrong.

### Account for Variable Inputs
Always read the student's actual inputs and compute expected outputs from those. Never penalize a student for choosing a different dropdown option than the key shows.

### Check Formula Structure, Not Just Values
For spreadsheets, verify students used the functions the assignment required (VLOOKUP, MATCH, IF). A hardcoded correct number doesn't demonstrate understanding.

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

Before delivering results:

- [ ] Phase 0 preflight passed — both repos clean and in sync
- [ ] Rubric point items sum to the assignment's stated total
- [ ] Every submission has a feedback `.docx` in `feedback/`
- [ ] `scores.xlsx` exists with all students, statistics, and formatting
- [ ] Feedback documents pass validation
- [ ] `scores.xlsx` formulas recalculated
- [ ] `feedback/batch_upload.zip` exists and holds every feedback document
- [ ] No cascading penalty violations — review anyone who lost points in multiple related areas
- [ ] Variable inputs accounted for — no false deductions from different input choices
- [ ] Every deduction includes a teaching explanation, not just "wrong"
- [ ] **No student file was written into either git repo**

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
