# Grading Script Patterns

This reference covers common patterns for writing the Python grading script that reads student submissions and compares them against the answer key.

## Table of Contents

1. [Reading Spreadsheet Submissions](#reading-spreadsheet-submissions)
2. [Tolerance Comparison](#tolerance-comparison)
3. [Formula Checking](#formula-checking)
4. [Variable Input Handling](#variable-input-handling)
5. [Output Format](#output-format)

---

## Reading Spreadsheet Submissions

Always open each submission twice — once for formulas and once for cached values:

```python
import openpyxl

# Formulas (raw cell content)
wb_f = openpyxl.load_workbook(filepath)

# Cached values (what Excel computed)
wb_v = openpyxl.load_workbook(filepath, data_only=True)
```

For named ranges, check via:
```python
defined_names = {}
for name, defn in wb_f.defined_names.items():
    defined_names[name.lower()] = defn.attr_text
```

---

## Tolerance Comparison

Numerical comparisons need tolerance because students may round at different stages:

```python
def close_enough(a, b, tol=0.02):
    """
    Check if two numbers are close enough.
    Returns True if within 2% relative difference or 0.01 absolute difference.
    """
    if a is None or b is None:
        return False
    try:
        a, b = float(a), float(b)
    except (ValueError, TypeError):
        return False
    if abs(b) < 0.001:
        return abs(a - b) < 0.01
    return abs(a - b) / max(abs(b), 0.001) < tol or abs(a - b) < 0.01
```

---

## Formula Checking

Check whether a cell contains a formula and what functions it uses:

```python
def is_formula(val):
    if val is None:
        return False
    return str(val).startswith('=')

def check_vlookup_match(formula):
    if formula is None:
        return False, False
    f = str(formula).upper()
    return 'VLOOKUP' in f, 'MATCH' in f

def check_if_statement(formula):
    if formula is None:
        return False
    return 'IF' in str(formula).upper()
```

---

## Variable Input Handling

This is one of the most important patterns. When assignments have dropdown menus or user-entered values, you cannot just compare against the key's cached values. You must:

1. Read the student's actual input values
2. Compute what the expected outputs should be for those inputs
3. Compare the student's outputs against your computed expected values

Example for a VLOOKUP-dependent cell:

```python
# Read the student's selected gravel type (dropdown in B2)
student_gravel = me_v['B2'].value

# Look up the expected cost per ton for that gravel type
GRAVEL_COSTS = {
    'Crushed Stone': 70.0,
    'Pea Gravel': 55.0,
    'River Rock': 85.0,
    # ... from the answer key's lookup table
}
expected_cost = GRAVEL_COSTS.get(student_gravel, None)

# Now compare the student's B7 value against the expected cost
# for THEIR gravel selection, not the key's
```

Example for computed values that depend on a user input:

```python
# Read the student's x value
student_x = float(sb_v['B10'].value)
student_a = float(sb_v['B7'].value)  # a might also be variable

# Compute expected deflection for the student's x
if student_x <= student_a:
    expected_v = (P * b * student_x) / (6 * E * Iu * L) * (b**2 + student_x**2 - L**2)
else:
    expected_v = -P * b / (6 * E * Iu * L) * ((L/b)*(student_x - student_a)**3
                  + (L**2 - b**2)*student_x - student_x**3)

# Compare against student's answer using close_enough()
```

---

## Output Format

The grading script should output a JSON file with this structure:

```json
[
    {
        "filename": "student_submission.xlsx",
        "scores": {
            "rubric_key_1": 5.0,
            "rubric_key_2": 1.5
        },
        "feedback": {
            "rubric_key_1": ["Correct! Great use of VLOOKUP + MATCH."],
            "rubric_key_2": ["Formula missing in B7. Expected =VLOOKUP(...)", "Values are correct despite formula issue."]
        },
        "total": 28.5
    }
]
```

Each entry in the `scores` dict uses the same rubric keys defined in the grading framework. Each entry in `feedback` is a list of strings — one per observation for that rubric item. The `total` is the sum of all scores.

This JSON file is then consumed by both the feedback document generator and the scores spreadsheet generator.
