# Feedback Document Template (docx-js / JavaScript)

This is the code template for generating per-student feedback `.docx` files using the `docx` npm package. Adapt the rubric items, section headers, and labels to match each assignment's rubric.

## Table of Contents

1. [Dependencies and Setup](#dependencies-and-setup)
2. [Data Structures to Adapt](#data-structures-to-adapt)
3. [Complete Template Code](#complete-template-code)
4. [Formatting Constants](#formatting-constants)
5. [Validation](#validation)

---

## Dependencies and Setup

Install the docx package:
```bash
npm install docx
```

The script reads `work/grading_results.json` (produced by the grading script) and generates one `.docx` file per student in the round's feedback folder — `feedback1/` on the first pass, `feedback2/` for a later batch of late work. `anonymize.py mask` names the round it just prepared; write into that folder, so the instructor reviews only what this round produced.

The script itself belongs in `work/`, addressed by absolute path, along with everything else you generate.

---

## Data Structures to Adapt

These two objects must be updated for each assignment. They define what rubric items exist and their display labels.

### `maxPts` — Point values per rubric item
```javascript
const maxPts = {
    item_key_1: 5,
    item_key_2: 2,
    // ... one entry per rubric item, keys must match grading_results.json
};
```

### `rubricLabels` — Labels for the summary table and the section headings
```javascript
const rubricLabels = {
    item_key_1: 'Part 1 - Description of what was graded',
    item_key_2: 'Part 2 - Description of what was graded',
    // ... same keys as maxPts
};
```

The same label heads the table row and any section about that item under
Points Missed or Other Feedback. `sync_scores.py` pairs a Points Missed section
with its table row by that text, so keep them identical.

---

## Complete Template Code

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, BorderStyle, WidthType, ShadingType, AlignmentType } = require('docx');
const fs = require('fs');

// ======== ADAPT THESE FOR EACH ASSIGNMENT ========

const ASSIGNMENT_NAME = 'Assignment Name Here';
const TOTAL_POINTS = 30; // Sum of all maxPts values

const maxPts = {
    // rubric_key: point_value
};

const rubricLabels = {
    // rubric_key: 'Label for the summary table row and the section heading'
};

// ======== CONSTANTS (do not change) ========

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

// Absolute paths. A relative one resolves against whatever directory the
// grader was invoked from -- usually the instructor repo, where student
// scores must never land.
const RESULTS_PATH = '<assignment>/work/grading_results.json';
const OUT_DIR      = '<assignment>/feedback1/';   // feedback2/ on a later round

// ======== HELPER FUNCTIONS ========

// Submissions arrive as .xlsx, .pdf, .docx, .ipynb -- strip whatever extension
// is on the file, not just .xlsx.
function stripExt(filename) {
    return filename.replace(/\.[A-Za-z0-9]+$/, '');
}

// While grading is de-identified the student IS the code. Emit the bare code
// here, with nothing appended: anonymize.py unmask finds S-XXXXXX in the
// document body and swaps in "First Last (netid)", so anything glued to the
// code would survive into the delivered feedback.
function extractStudentName(filename) {
    const m = filename.match(/^(S-[0-9A-F]{6})/);
    return m ? m[1] : stripExt(filename);
}

function scoreColor(total, maxTotal) {
    const pct = total / maxTotal;
    if (pct >= 0.9) return "008000";      // green
    if (pct >= 0.73) return "FF8C00";     // orange
    return "FF0000";                       // red
}

// ======== MAIN DOCUMENT BUILDER ========

async function createFeedback(result) {
    const studentName = extractStudentName(result.filename);
    const total = result.total;
    const children = [];

    // --- Title ---
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({
            text: `Homework Feedback: ${ASSIGNMENT_NAME}`,
            bold: true, size: 28, font: "Arial"
        })]
    }));

    // --- Student info ---
    children.push(new Paragraph({
        spacing: { after: 100 },
        children: [
            new TextRun({ text: `Student: `, bold: true, size: 22, font: "Arial" }),
            new TextRun({ text: studentName, size: 22, font: "Arial" }),
        ]
    }));
    children.push(new Paragraph({
        spacing: { after: 200 },
        children: [
            new TextRun({ text: `Total Score: `, bold: true, size: 22, font: "Arial" }),
            new TextRun({
                text: `${total} / ${TOTAL_POINTS}`,
                size: 22, font: "Arial",
                color: scoreColor(total, TOTAL_POINTS)
            }),
        ]
    }));

    // --- Summary rubric table (first, so students see the scores at a glance) ---
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
        children: [new TextRun({
            text: "Grading Summary", bold: true, size: 28, font: "Arial"
        })]
    }));

    // Header row
    const headerRow = new TableRow({
        children: [
            new TableCell({
                borders, width: { size: 5500, type: WidthType.DXA },
                shading: { fill: "2E4057", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ children: [new TextRun({
                    text: "Rubric Item", bold: true, color: "FFFFFF",
                    size: 20, font: "Arial"
                })] })]
            }),
            new TableCell({
                borders, width: { size: 1400, type: WidthType.DXA },
                shading: { fill: "2E4057", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ alignment: AlignmentType.CENTER,
                    children: [new TextRun({
                        text: "Possible", bold: true, color: "FFFFFF",
                        size: 20, font: "Arial"
                    })] })]
            }),
            new TableCell({
                borders, width: { size: 1400, type: WidthType.DXA },
                shading: { fill: "2E4057", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ alignment: AlignmentType.CENTER,
                    children: [new TextRun({
                        text: "Earned", bold: true, color: "FFFFFF",
                        size: 20, font: "Arial"
                    })] })]
            }),
        ]
    });

    // Data rows (one per rubric item)
    const dataRows = [];
    for (const key of Object.keys(rubricLabels)) {
        const score = result.scores[key];
        const max = maxPts[key];
        const isFullScore = score >= max;
        const fillColor = isFullScore ? "E8F5E9" : "FFEBEE";
        const scoreTextColor = isFullScore ? "008000" : "FF0000";

        dataRows.push(new TableRow({
            children: [
                new TableCell({
                    borders, width: { size: 5500, type: WidthType.DXA },
                    shading: { fill: fillColor, type: ShadingType.CLEAR },
                    margins: cellMargins,
                    children: [new Paragraph({ children: [new TextRun({
                        text: rubricLabels[key], size: 19, font: "Arial"
                    })] })]
                }),
                new TableCell({
                    borders, width: { size: 1400, type: WidthType.DXA },
                    shading: { fill: fillColor, type: ShadingType.CLEAR },
                    margins: cellMargins,
                    children: [new Paragraph({ alignment: AlignmentType.CENTER,
                        children: [new TextRun({
                            text: `${max}`, size: 19, font: "Arial"
                        })] })]
                }),
                new TableCell({
                    borders, width: { size: 1400, type: WidthType.DXA },
                    shading: { fill: fillColor, type: ShadingType.CLEAR },
                    margins: cellMargins,
                    children: [new Paragraph({ alignment: AlignmentType.CENTER,
                        children: [new TextRun({
                            text: `${score}`, size: 19, font: "Arial",
                            color: scoreTextColor, bold: !isFullScore
                        })] })]
                }),
            ]
        }));
    }

    // Total row (blue)
    dataRows.push(new TableRow({
        children: [
            new TableCell({
                borders, width: { size: 5500, type: WidthType.DXA },
                shading: { fill: "E3F2FD", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ children: [new TextRun({
                    text: "TOTAL", bold: true, size: 20, font: "Arial"
                })] })]
            }),
            new TableCell({
                borders, width: { size: 1400, type: WidthType.DXA },
                shading: { fill: "E3F2FD", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ alignment: AlignmentType.CENTER,
                    children: [new TextRun({
                        text: `${TOTAL_POINTS}`, bold: true,
                        size: 20, font: "Arial"
                    })] })]
            }),
            new TableCell({
                borders, width: { size: 1400, type: WidthType.DXA },
                shading: { fill: "E3F2FD", type: ShadingType.CLEAR },
                margins: cellMargins,
                children: [new Paragraph({ alignment: AlignmentType.CENTER,
                    children: [new TextRun({
                        text: `${total}`, bold: true, size: 20, font: "Arial",
                        color: scoreColor(total, TOTAL_POINTS)
                    })] })]
            }),
        ]
    }));

    const rubricTable = new Table({
        width: { size: 8300, type: WidthType.DXA },
        columnWidths: [5500, 1400, 1400],
        rows: [headerRow, ...dataRows]
    });
    children.push(rubricTable);

    const partHeading = (text) => new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 120 },
        children: [new TextRun({ text, bold: true, size: 28, font: "Arial" })]
    });
    const itemHeading = (text) => new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 300, after: 80 },
        children: [new TextRun({ text, bold: true, size: 22, font: "Arial" })]
    });
    const bullet = (text) => new Paragraph({
        spacing: { after: 60 },
        indent: { left: 360 },
        children: [new TextRun({ text, size: 20, font: "Arial" })]
    });
    const keys = Object.keys(rubricLabels);

    // --- Points Missed: one section per item that lost points ---
    // Full-credit items get no section here; the table already shows them.
    // The heading is the table label, which sync_scores.py matches on.
    const missed = keys.filter(k => result.scores[k] < maxPts[k]);
    if (missed.length) children.push(partHeading("Points Missed"));
    for (const key of missed) {
        const score = result.scores[key];
        const max = maxPts[key];
        children.push(itemHeading(rubricLabels[key]));
        children.push(new Paragraph({
            spacing: { after: 80 },
            children: [
                new TextRun({ text: `Score: `, bold: true, size: 20, font: "Arial" }),
                new TextRun({
                    text: `${score} / ${max}`,
                    bold: true, size: 20, font: "Arial", color: "FF0000"
                }),
            ]
        }));
        for (const item of result.feedback[key] || []) children.push(bullet(item));
    }

    // --- Other Feedback: better habits on full-credit work, and general notes ---
    // No score line here, so sync_scores.py never reads these as deductions.
    const suggestions = result.suggestions || {};
    const suggested = keys.filter(k => result.scores[k] >= maxPts[k] && (suggestions[k] || []).length);
    const notes = result.notes || [];
    if (suggested.length || notes.length) children.push(partHeading("Other Feedback"));
    for (const key of suggested) {
        children.push(itemHeading(rubricLabels[key]));
        for (const item of suggestions[key]) children.push(bullet(item));
    }
    if (notes.length) {
        children.push(itemHeading("General"));
        for (const note of notes) children.push(bullet(note));
    }

    // --- Encouragement ---
    children.push(new Paragraph({ spacing: { before: 300 }, children: [] }));
    const pct = total / TOTAL_POINTS;
    let encouragement;
    if (pct >= 0.93) {
        encouragement = "Excellent work! You demonstrated a strong understanding of the material. Keep it up!";
    } else if (pct >= 0.83) {
        encouragement = "Good job overall! You have a solid foundation. Review the Points Missed section above.";
    } else if (pct >= 0.67) {
        encouragement = "You made a good effort on this assignment. Please read the Points Missed section above. Feel free to reach out during office hours for help.";
    } else {
        encouragement = "Please review the assignment instructions carefully and revisit the concepts from the reading material. Office hours are a great resource for getting help.";
    }
    children.push(new Paragraph({
        spacing: { before: 200 },
        children: [new TextRun({
            text: encouragement, italics: true,
            size: 20, font: "Arial", color: "555555"
        })]
    }));

    // --- Build document ---
    const doc = new Document({
        styles: {
            default: { document: { run: { font: "Arial", size: 20 } } },
            paragraphStyles: [
                {
                    id: "Heading1", name: "Heading 1",
                    basedOn: "Normal", next: "Normal", quickFormat: true,
                    run: { size: 28, bold: true, font: "Arial" },
                    paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
                },
                {
                    id: "Heading2", name: "Heading 2",
                    basedOn: "Normal", next: "Normal", quickFormat: true,
                    run: { size: 24, bold: true, font: "Arial" },
                    paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 }
                },
            ]
        },
        sections: [{
            properties: {
                page: {
                    size: { width: 12240, height: 15840 },  // US Letter
                    margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }  // 0.75"
                }
            },
            children
        }]
    });

    // --- Save ---
    const outName = `${stripExt(result.filename)}_FEEDBACK.docx`;
    const outPath = `${OUT_DIR}/${outName}`;
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buffer);
    console.log(`Created: ${outName}`);
}

// ======== RUN ========

const results = JSON.parse(fs.readFileSync(RESULTS_PATH, 'utf8'));
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

async function main() {
    for (const r of results) {
        await createFeedback(r);
    }
    console.log(`\nAll ${results.length} feedback documents created in ${OUT_DIR}/`);
}

main().catch(console.error);
```

---

## Formatting Constants

| Element | Value | Notes |
|---------|-------|-------|
| Font | Arial | Throughout the document |
| Page size | 12240 × 15840 DXA | US Letter |
| Margins | 1080 DXA (0.75 in) | All sides |
| Header row fill | #2E4057 | Dark blue-gray, white text |
| Full-score fill | #E8F5E9 | Light green |
| Deduction fill | #FFEBEE | Light red |
| Total row fill | #E3F2FD | Light blue |
| Full-score text | #008000 | Green |
| Deduction text | #FF0000 | Red |
| Orange (mid-range) | #FF8C00 | Used for totals in the B-C range |
| Cell borders | #999999 thin | Subtle gray borders |
| Cell margins | 60 top/bottom, 100 left/right | DXA units |
| Table widths | DXA only | Never use percentage |
| Shading type | ShadingType.CLEAR | Never use SOLID |

---

## Validation

After creating feedback documents, confirm every one opens and carries the content it should:

```bash
python3 -c '
import sys
from docx import Document
for path in sys.argv[1:]:
    doc = Document(path)
    status = "ok" if doc.paragraphs and doc.tables else "SUSPECT"
    print(f"{status:8} {len(doc.paragraphs):3} paragraphs  {len(doc.tables)} tables  {path}")
' <assignment>/feedback1/*.docx
```

A document with zero tables never got its rubric summary; regenerate it. A document
that raises on open is corrupt -- usually a template value that was `undefined` when
the docx package serialized it.

Spot-check one visually by converting it to PDF:

```bash
soffice --headless --convert-to pdf <feedback_file.docx>
# macOS, where soffice is not on PATH:
# /Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf <file>
```
