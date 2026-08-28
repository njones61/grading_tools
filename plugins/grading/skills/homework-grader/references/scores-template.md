# Score Summary Template (openpyxl / Python)

This is the code template for generating the `scores.xlsx` class summary spreadsheet using Python `openpyxl`. Adapt the rubric items, column headers, and total points to match each assignment.

## Table of Contents

1. [Dependencies](#dependencies)
2. [Data Structures to Adapt](#data-structures-to-adapt)
3. [Complete Template Code](#complete-template-code)
4. [Formatting Constants](#formatting-constants)
5. [Recalculation](#recalculation)

---

## Dependencies

```bash
pip install openpyxl --break-system-packages
```

The script reads a `grading_results.json` file (produced by the grading script) and creates `scores.xlsx` in the assignment folder.

---

## Data Structures to Adapt

These must be updated for each assignment:

### `rubric_keys` — Internal keys matching grading_results.json
```python
rubric_keys = ['item_key_1', 'item_key_2', ...]
```

### `rubric_short` — Short column header labels
```python
rubric_short = ['P1 Task', 'P2 Task', ...]
```

### `max_pts` — Point values (same order as rubric_keys)
```python
max_pts = [5, 2, ...]
```

### `TOTAL_POINTS` — Sum of all max_pts
```python
TOTAL_POINTS = sum(max_pts)
```

---

## Complete Template Code

```python
#!/usr/bin/env python3
"""Create scores.xlsx with summary table and statistics"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json

# ======== ADAPT THESE FOR EACH ASSIGNMENT ========

COURSE_NAME = 'Course Name'
ASSIGNMENT_NAME = 'Assignment Name'
RESULTS_PATH = 'grading_results.json'   # Set to actual path
OUTPUT_PATH = 'scores.xlsx'              # Set to actual path

rubric_keys = [
    # Must match keys in grading_results.json scores dict
]

rubric_short = [
    # Short display names for column headers
]

max_pts = [
    # Point values, same order as rubric_keys
]

TOTAL_POINTS = sum(max_pts)

# ======== LOAD DATA ========

with open(RESULTS_PATH) as f:
    results = json.load(f)

# ======== STYLES (do not change) ========

header_font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='2E4057')
data_font = Font(name='Arial', size=11)
bold_font = Font(name='Arial', bold=True, size=11)
stat_font = Font(name='Arial', bold=True, size=11, color='2E4057')
green_fill = PatternFill('solid', fgColor='C8E6C9')
red_fill = PatternFill('solid', fgColor='FFCDD2')
yellow_fill = PatternFill('solid', fgColor='FFF9C4')
blue_fill = PatternFill('solid', fgColor='BBDEFB')
thin_border = Border(
    left=Side(style='thin', color='999999'),
    right=Side(style='thin', color='999999'),
    top=Side(style='thin', color='999999'),
    bottom=Side(style='thin', color='999999')
)
center = Alignment(horizontal='center', vertical='center')

# ======== NAME PARSING (adapt for each assignment) ========

def parse_name(filename):
    """
    Extract first name, last name, and net ID from submission filename.
    Adapt this function for each assignment's filename convention.
    Common format: last_first_netid_userfilename.xlsx
    Students often deviate — handle edge cases.
    """
    name = filename.replace('.xlsx', '').replace('.XLSX', '')
    # Remove assignment-specific suffixes
    # name = name.replace('-HW-Name-Here', '')

    if '_' in name:
        parts = name.split('_')
        if len(parts) >= 3:
            return parts[1].strip(), parts[0].strip(), parts[2].strip()
        elif len(parts) == 2:
            return parts[1].strip(), parts[0].strip(), ''
    elif '-' in name:
        parts = name.split('-')
        if len(parts) >= 2:
            return parts[-1].strip(), parts[0].strip(), ''
    elif ' ' in name:
        parts = name.split()
        if len(parts) >= 2:
            return parts[0].strip(), parts[-1].strip(), ''
    return name, '', ''

# ======== BUILD WORKBOOK ========

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scores"

# IMPORTANT: Turn off gridlines for a clean look
ws.sheet_view.showGridLines = False

# --- Title row ---
num_cols = 3 + len(rubric_short) + 1  # name cols + rubric cols + total
last_col_letter = get_column_letter(num_cols)
ws.merge_cells(f'A1:{last_col_letter}1')
ws['A1'] = f'{COURSE_NAME} - {ASSIGNMENT_NAME} - Score Summary'  # Only include COURSE_NAME if found in assignment materials; leave blank if not
ws['A1'].font = Font(name='Arial', bold=True, size=14, color='2E4057')
ws['A1'].alignment = Alignment(horizontal='center')

# --- Header row (row 3) ---
headers = ['First Name', 'Last Name', 'Net ID'] + rubric_short + ['Total']
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = thin_border

# --- Max points row (row 4) ---
ws.cell(row=4, column=1, value='Max Points').font = bold_font
for col_idx in range(1, 4):
    ws.cell(row=4, column=col_idx).fill = yellow_fill
    ws.cell(row=4, column=col_idx).border = thin_border

for i, mp in enumerate(max_pts):
    cell = ws.cell(row=4, column=4 + i, value=mp)
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.alignment = center
    cell.border = thin_border

# Max points total uses SUM formula
total_col = 4 + len(max_pts)
total_col_letter = get_column_letter(total_col)
first_rubric_letter = get_column_letter(4)
last_rubric_letter = get_column_letter(3 + len(max_pts))
cell = ws.cell(row=4, column=total_col)
cell.value = f'=SUM({first_rubric_letter}4:{last_rubric_letter}4)'
cell.font = bold_font
cell.fill = yellow_fill
cell.alignment = center
cell.border = thin_border

# --- Student data rows (starting row 5) ---
for row_idx, r in enumerate(results, 5):
    first, last, netid = parse_name(r['filename'])

    ws.cell(row=row_idx, column=1, value=first).font = data_font
    ws.cell(row=row_idx, column=1).border = thin_border
    ws.cell(row=row_idx, column=2, value=last).font = data_font
    ws.cell(row=row_idx, column=2).border = thin_border
    ws.cell(row=row_idx, column=3, value=netid).font = data_font
    ws.cell(row=row_idx, column=3).border = thin_border

    for i, key in enumerate(rubric_keys):
        score = r['scores'][key]
        cell = ws.cell(row=row_idx, column=4 + i, value=score)
        cell.font = data_font
        cell.alignment = center
        cell.border = thin_border
        # NOTE: Don't apply fills here — use conditional formatting below

    # Total column uses SUM formula (not hardcoded value)
    total_cell = ws.cell(row=row_idx, column=total_col)
    total_cell.value = f'=SUM({first_rubric_letter}{row_idx}:{last_rubric_letter}{row_idx})'
    total_cell.font = Font(name='Arial', bold=True, size=11)
    total_cell.alignment = center
    total_cell.border = thin_border

last_data_row = 4 + len(results)

# --- Conditional formatting for score cells ---
# This ensures colors update automatically if the instructor edits a score by hand.
from openpyxl.formatting.rule import CellIsRule
red_font = Font(name='Arial', size=11, color='CC0000', bold=True)

for i, mp in enumerate(max_pts):
    col_letter = get_column_letter(4 + i)
    cell_range = f'{col_letter}5:{col_letter}{last_data_row}'
    ws.conditional_formatting.add(cell_range,
        CellIsRule(operator='lessThan', formula=[str(mp)],
                   fill=red_fill, font=red_font))
    ws.conditional_formatting.add(cell_range,
        CellIsRule(operator='greaterThanOrEqual', formula=[str(mp)],
                   fill=green_fill))

# Total column conditional formatting
total_range = f'{total_col_letter}5:{total_col_letter}{last_data_row}'
ws.conditional_formatting.add(total_range,
    CellIsRule(operator='lessThan', formula=[str(TOTAL_POINTS)],
               fill=red_fill))
ws.conditional_formatting.add(total_range,
    CellIsRule(operator='greaterThanOrEqual', formula=[str(TOTAL_POINTS)],
               fill=green_fill))

# --- Statistics rows ---
stat_row = last_data_row + 2
stat_labels = ['Average (pts)', 'Average (%)']

for label_idx, label in enumerate(stat_labels):
    row = stat_row + label_idx

    ws.cell(row=row, column=1, value=label).font = stat_font
    ws.cell(row=row, column=1).fill = blue_fill
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=2).fill = blue_fill
    ws.cell(row=row, column=2).border = thin_border
    ws.cell(row=row, column=3).fill = blue_fill
    ws.cell(row=row, column=3).border = thin_border

    for col_idx in range(4, 4 + len(rubric_keys) + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.fill = blue_fill
        cell.border = thin_border
        cell.alignment = center
        cell.font = stat_font

        col_letter = get_column_letter(col_idx)
        data_range = f'{col_letter}5:{col_letter}{last_data_row}'

        if label_idx == 0:  # Average points
            cell.value = f'=AVERAGE({data_range})'
            cell.number_format = '0.0'
        else:  # Average percentage
            max_col_idx = col_idx - 4
            if max_col_idx < len(max_pts):
                mp = max_pts[max_col_idx]
            else:
                mp = TOTAL_POINTS
            cell.value = f'=AVERAGE({data_range})/{mp}*100'
            cell.number_format = '0.0"%"'

# --- Column widths ---
ws.column_dimensions['A'].width = 14
ws.column_dimensions['B'].width = 14
ws.column_dimensions['C'].width = 10
for i in range(len(rubric_short)):
    ws.column_dimensions[get_column_letter(4 + i)].width = 12
ws.column_dimensions[get_column_letter(4 + len(rubric_short))].width = 10

# --- Save ---
wb.save(OUTPUT_PATH)
print(f"Saved: {OUTPUT_PATH}")
```

---

## Formatting Constants

| Element | Value | Notes |
|---------|-------|-------|
| Font | Arial, 11pt | Throughout |
| Header row fill | #2E4057 | Dark blue-gray, white bold text |
| Max points row fill | #FFF9C4 | Yellow |
| Full-score cell fill | #C8E6C9 | Light green |
| Deduction cell fill | #FFCDD2 | Light red |
| Deduction text color | #CC0000 | Dark red, bold |
| Statistics row fill | #BBDEFB | Light blue |
| Border style | thin, #999999 | All data cells |
| Gridlines | OFF | `ws.sheet_view.showGridLines = False` |
| Column widths | 14 (names), 10 (net ID), 12 (scores), 10 (total) | Adjust as needed |

---

## Learning Suite Upload CSV

`scores.xlsx` is the instructor's working sheet; Learning Suite's grade import wants something much narrower, and rejects the rest:

```
Net ID,Homework 1 - Head Calculations
mattd1,29
jra94,30
```

`Net ID` in the top-left cell, the assignment title beside it, plain numbers, one row per student. Both columns need a header or nothing imports. The max-points row, the per-item columns, and the average rows must all stay out — an `Average (points)` row imports as a student who does not exist.

Do not write this from `grading_results.json`: at that point the identities are still codes. `export_upload_csv.py` derives it from the finished `scores.xlsx`, after `unmask` has restored NetIDs and recalculated the totals, so it also reflects any score you adjusted by hand.

---

## Recalculation

After creating `scores.xlsx`, recalculate formulas so cached values are correct:

```bash
python scripts/recalc.py scores.xlsx
```

If `recalc.py` is not available, use LibreOffice:
```bash
libreoffice --headless --calc --convert-to xlsx --outdir /tmp scores.xlsx
cp /tmp/scores.xlsx scores.xlsx
```

This ensures the AVERAGE formulas in the statistics rows have computed values, not just formula text.
