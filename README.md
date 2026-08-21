# grading_tools

Private Claude Code plugin marketplace holding the teaching tools shared across BYU courses. Currently one plugin: **grading**, which provides the `homework-grader` skill.

## Install

Once per machine:

```
/plugin marketplace add njones61/grading_tools
/plugin install grading@grading_tools
```

The skill is then available in every Claude Code project on that machine — `ce544_private`, `ce547_private`, `instructor`, anywhere.

After pushing a change here:

```
/plugin update grading
```

Claude Code clones this repo with your existing `gh` credentials, so a private repo works without extra setup. Anyone you add as a collaborator can run the same two commands.

## The Repo Model

Course material lives in three places. Nothing is duplicated.

```
  content repo (public)         instructor repo (private)      grading workspace
  ────────────────────          ─────────────────────────      ─────────────────
  assignment + rubric           answer keys                    submissions
  background reading            course.yml                     feedback
  in-class material             grading_guide.md               scores.xlsx
                                key.md per assignment          roster.json

  njones61/ce544                njones61/ce544_private         ~/grading/ce544
  njones61/ce547                njones61/ce547_private         ~/grading/ce547
  byu-cce270/content            byu-cce270/instructor          ~/grading/cce270
```

The assignment students read *is* the file the grader grades against — a markdown file with the rubric inline. There is no PDF to print, download, or keep in sync.

**Student submissions and feedback never enter a git repo.** They are FERPA-protected education records and live only in the grading workspace.

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

## Adding a Course

1. Create a private instructor repo beside the public content repo.
2. Add `course.yml` naming both repos and the grading workspace.
3. Add `grading_guide.md` for course-level policy.
4. Put keys under `keys/<unit>/<topic>/`, each with a `key.md` pointing at its assignment.

## Layout

```
grading_tools/
├── .claude-plugin/marketplace.json     catalog
└── plugins/grading/
    ├── .claude-plugin/plugin.json
    └── skills/homework-grader/
        ├── SKILL.md
        ├── references/
        │   ├── grading_code.md          Python, notebooks, scripts
        │   ├── api_reference.md         openpyxl patterns
        │   ├── feedback-template.md     docx template
        │   └── scores-template.md       xlsx template
        └── scripts/
            ├── preflight_sync.py        verify repos clean and in sync
            ├── anonymize.py             mask / unmask student identities
            └── recalc.py                recalculate xlsx formulas
```

## Scripts

```bash
# Before grading: confirm both repos are clean and current
python preflight_sync.py ~/python_projects/ce544_private [--pull]

# De-identify submissions, then grade the masked copies
python anonymize.py mask ~/grading/ce544/2026-fall/01_head/submissions

# After grading: restore real names for Learning Suite upload
python anonymize.py unmask ~/grading/ce544/2026-fall/01_head/feedback \
    --roster ~/grading/ce544/2026-fall/01_head/roster.json \
    --scores ~/grading/ce544/2026-fall/01_head/scores.xlsx
```

`anonymize.py mask` reports two things you must read: filenames it could not parse (not copied), and files it could not scrub (scanned PDFs and images, whose content may still identify the student regardless of filename).
