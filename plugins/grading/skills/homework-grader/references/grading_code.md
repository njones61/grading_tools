# Grading Code Submissions (Python, notebooks, scripts)

Load this reference when submissions are `.py`, `.ipynb`, or any other source code. It applies to any course that assigns programming — CCE 270's Python units, CE 544's xslope work, and anything similar. It is not course-specific.

Course-level overrides live in the instructor repo's `grading_guide.md`; assignment-level overrides live in the `key.md` body. Both outrank this file.

## The Core Principle: Grade the Reasoning, Not the Diff

There is no single correct program. Two students can write completely different code that both fully satisfy the assignment. Never diff a submission against the key and deduct for divergence — read the code, decide whether it does the right thing, and grade *that*.

Deduct for: wrong output, wrong approach, code that doesn't run, missing required constructs.
Do **not** deduct for: different variable names, different loop style, extra helper functions, fewer lines, more lines, a different but valid algorithm.

## Always Execute the Code

Static reading misses real bugs and invents fake ones. Run every submission.

```bash
python3 submission.py                    # scripts
jupyter nbconvert --to notebook --execute --inplace --allow-errors copy.ipynb
```

Rules:
- **Work on a copy** in the grading workspace. Never modify the original submission.
- **Run with a timeout** (30–60s). Student code contains infinite loops; that is a finding, not a hang.
- **`--allow-errors`** so one broken cell doesn't hide the rest of the notebook's results.
- If the code needs a data file the student didn't submit, supply it from the assignment's `files/` folder in the content repo before concluding it's broken.

## Separate the Four Failure Modes

When output is wrong, diagnose *which* of these it is — the feedback and the deduction differ sharply:

| Mode | Meaning | Typical treatment |
|---|---|---|
| **Doesn't run** | Syntax error, bad import, crash | Find the fault, note the exact line. Fix a *trivial* fault (a typo, a missing import) in your copy and grade the rest of the work on its merits — a one-character slip should not zero an otherwise correct assignment. |
| **Runs, wrong answer** | Logic error | Locate the specific wrong line. Full credit for the parts that are right. |
| **Runs, right answer, wrong method** | Hardcoded result, or ignored a required construct | Deduct for the requirement they skipped, not for the answer. |
| **Runs, right answer, right method** | Full credit | Say something specific about what they did well. |

## Check Required Constructs

Assignments usually mandate specific language features ("use a `for` loop", "write a function that returns…", "use a dictionary"). Verify structurally, not by grepping text — a comment mentioning "dictionary" isn't a dictionary.

```python
import ast
tree = ast.parse(source)
has_loop = any(isinstance(n, (ast.For, ast.While)) for n in ast.walk(tree))
funcs = [n.name for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)]
```

A correct number produced by hardcoding rather than the required construct earns the points for the answer but not for the method. Say so plainly: the point of the exercise was the mechanism.

## Notebooks Specifically

- **Parse as JSON.** Grade both `source` (what they wrote) and `outputs` (what it produced).
- **Empty outputs mean they never ran it.** Note it; a small deduction is usually right, since submitting unverified work is the actual mistake.
- **Check execution order.** Non-sequential `execution_count` values mean cells were run out of order — results may not be reproducible from a clean restart. Worth a note when the output depends on it.
- **Cells that error** show as `"output_type": "error"` with a traceback. Read the traceback; it usually names the exact problem.
- **Markdown cells** often hold the written answer to a question. Don't grade only the code cells and miss the prose.

## Plots and Figures

For matplotlib and similar output, check what the assignment actually required — usually axis labels, a title, a legend when multiple series are present, and the right chart type. Render the figure and look at it rather than only reading the plotting call; a correct-looking `plt.plot()` can still produce an empty or nonsensical chart.

## Style

Default: **do not deduct for style unless the assignment's rubric has a line item for it.** Comment on it warmly instead — naming, structure, and readability are worth mentioning as growth, not as penalty. If a rubric item does cover documentation or comments, grade it against what the rubric says, not against your own taste.

## Feedback for Code

The general three-part rule from SKILL.md still holds — what they did, what they should have done, why it matters — but for code, **quote the specific line**:

> **Problem 3 (−2):** Your loop was `for i in range(len(data))` and then indexed `data[i]` inside. That works, but line 14 reads `total += data[i+1]`, which steps one element past the end on the final iteration and raises `IndexError`. Iterating directly with `for value in data:` avoids the whole class of off-by-one error, since there's no index to get wrong.

Quoting the line makes the feedback checkable. Paraphrasing makes students hunt for what you meant.

## Never Run Untrusted Code Carelessly

Student code is untrusted input. Before executing, skim for anything that touches the filesystem outside its own directory, opens a network connection, or calls `subprocess`/`os.system`. Assignments rarely need any of those. If you see one, don't run it — read it, grade it statically, and flag it to the instructor. This is almost always a student copying something off the internet rather than malice, but running it blind is still the wrong move.
