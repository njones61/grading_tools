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

To pick up later changes:

```
/plugin marketplace update grading_tools
```

Use `marketplace update`. `/plugin update grading` opens the plugin browser without updating anything.

### 2. Check out the repos side by side

`course.yml` points at the content repo with a relative path, so they must be siblings:

```bash
cd ~/python_projects
git clone https://github.com/njones61/ce544.git
git clone https://github.com/njones61/ce544_private.git
```

### 3. Install the dependencies

The scripts shell out to these. Grading fails partway through Phase 4 without them.

```bash
brew install --cask libreoffice      # recalculates scores.xlsx formulas
pip install openpyxl python-docx     # reads submissions, validates feedback
cd ~/grading && npm install docx     # builds the feedback .docx files
```

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
└── submissions/
    ├── Last_First_netid_theirfilename.xlsx
    └── ...
```

Keep Learning Suite's `last_first_netid_...` filenames — the de-identification step parses them.

### 2. Ask Claude Code to grade it

From the instructor repo:

```
grade 01_head
```

That's it. The skill runs the whole workflow: sync check, assemble the assignment and key, de-identify, grade, generate feedback, restore names, and build the two upload artifacts.

### 3. Upload the results

```
~/grading/ce544/2026-fall/01_head/
├── feedback/
│   ├── <one .docx per student, real names restored>
│   └── batch_upload.zip    ← upload this, not the files individually
├── scores.xlsx             class summary with statistics, for you
└── scores_upload.csv       ← Net ID + score, for the Grades import
```

Learning Suite takes the zip for feedback and the CSV for scores. `scores.xlsx` is your working copy and is not uploaded — but if you adjust a score in it, regenerate the CSV so the two agree.

---

## What Happens Under the Hood

You can run any phase by hand if something goes wrong. These commands need the plugin's script directory:

```bash
GT=$(ls -d ~/.claude/plugins/cache/grading_tools/grading/*/skills/homework-grader/scripts | sort -V | tail -1)
echo "$GT"    # should end in .../<version>/skills/homework-grader/scripts
```

The cache keeps every version ever installed, so this selects the newest deliberately. Do not search for the directory by name and take the first hit — that returns the marketplace clone and every stale version in arbitrary order, and will happily run a year-old script.

(Claude Code knows this path as `${CLAUDE_PLUGIN_ROOT}` and resolves it itself. You only need `$GT` when running a step by hand.)

### Phase 0 — Sync preflight

Both repos must be clean and current, or you risk grading against a stale rubric.

```bash
python "$GT/preflight_sync.py" ~/python_projects/ce544_private --pull
```

`--pull` fast-forwards a repo that is clean and merely behind. It **stops** rather than pulling if a repo has uncommitted changes or has diverged — those need you.

### Phase 1 — Resolve the assignment

Reads `course.yml`, finds the key folder, reads its `key.md`:

```markdown
---
assignment: docs/unit1/01_head/head_hw.md
background:
  - docs/unit1/01_head/head_read.md
---
Assignment-specific grading notes go here.
```

Paths in the frontmatter are relative to the **content repo** root.

### Phase 2.5 — De-identify

```bash
python "$GT/anonymize.py" mask ~/grading/ce544/2026-fall/01_head/submissions
```

Each student gets a random opaque code (`S-7F3A2B`). Names and NetIDs are scrubbed from spreadsheet cells, document text, and notebook content. The crosswalk lands in `roster.json` at mode 600. Grading then runs against `masked/`, never `submissions/`.

**Read the three categories it reports:**

- **COULD NOT PARSE** — filenames not matching `last_first_netid_...`. These are *not* copied. Rename them by hand and re-run.
- **REDACTED** — scans and images whose header band was blacked out. This is destructive: the page is rasterized, the band is painted onto the bitmap, and the scanner's OCR text layer goes with it. **Look at `masked/redaction_previews/` before grading** — the band is geometry, not a name detector, so a name written down a margin or on a later page is still there. `--band <percent>` adjusts the height; `--no-redact` turns it off.
- **UNSCRUBBABLE** — what redaction could not cover, plus born-digital PDFs, which are skipped so they keep their selectable text. These may still identify the student.

Detection is deliberately local. Sending a page to a vision model to locate the name would transmit the very thing the redaction exists to withhold.

### Phases 3–4 — Grade and generate

Grades the masked copies against the rubric, writes one feedback `.docx` per student, and builds `scores.xlsx`.

### Phase 5 — Restore identities

```bash
python "$GT/anonymize.py" unmask ~/grading/ce544/2026-fall/01_head/feedback \
    --roster ~/grading/ce544/2026-fall/01_head/roster.json \
    --scores ~/grading/ce544/2026-fall/01_head/scores.xlsx
```

Do this **last**, after any regrading. It:

1. renames feedback files to `last_first_netid_...`
2. replaces the code *inside* each document with `First Last (netid)`
3. turns codes in `scores.xlsx` into NetIDs and fills the First/Last Name columns
4. recalculates `scores.xlsx` — rewriting it clears the cached formula values, so skipping this leaves every total blank
5. writes `scores_upload.csv`
6. builds `feedback/batch_upload.zip`

Steps 5 and 6 are snapshots. Edit a feedback document or a score afterward and you must rebuild them:

```bash
python "$GT/package_feedback.py"   ~/grading/ce544/2026-fall/01_head/feedback
python "$GT/export_upload_csv.py"  ~/grading/ce544/2026-fall/01_head/scores.xlsx
```

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
            ├── preflight_sync.py        verify repos clean and in sync
            ├── anonymize.py             mask / unmask student identities
            ├── redact_scan.py           black out the header band on scans
            ├── recalc.py                recalculate xlsx formulas
            ├── package_feedback.py      build feedback/batch_upload.zip
            └── export_upload_csv.py     build scores_upload.csv
```
