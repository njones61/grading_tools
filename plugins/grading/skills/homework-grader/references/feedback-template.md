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

The script reads a `grading_results.json` file (produced by the grading script) and generates one `.docx` file per student in a `feedback/` subfolder.

---

## Data Structures to Adapt

These three objects must be updated for each assignment. They define what rubric items exist, their display labels, and section groupings.

### `maxPts` — Point values per rubric item
```javascript
const maxPts = {
    item_key_1: 5,
    item_key_2: 2,
    // ... one entry per rubric item, keys must match grading_results.json
};
```

### `rubricLabels` — Labels for the summary table
```javascript
const rubricLabels = {
    item_key_1: 'Part 1 - Description of what was graded',
    item_key_2: 'Part 2 - Description of what was graded',
    // ... same keys as maxPts
};
```

### `sectionHeaders` — Headers for detailed feedback sections
```javascript
const sectionHeaders = {
    item_key_1: 'Part 1: Topic - Subtopic',
    item_key_2: 'Part 2: Topic - Subtopic',
    // ... same keys as maxPts
};
```

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
    // rubric_key: 'Display label for summary table'
};

const sectionHeaders = {
    // rubric_key: 'Section heading for detailed feedback'
};

// ======== CONSTANTS (do not change) ========

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

const RESULTS_PATH = 'grading_results.json';  // Set to actual path
const OUT_DIR = 'feedback/';                    // Set to actual path

// ======== HELPER FUNCTIONS ========

function extractStudentName(filename) {
    // Adapt name parsing for each assignment's filename convention
    let name = filename.replace(/\.xlsx$/i, '');
    name = name.replace(/^\d+-/, '');
    return name;
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

    // --- Per-section detailed feedback ---
    const keys = Object.keys(sectionHeaders);
    for (const key of keys) {
        const score = result.scores[key];
        const max = maxPts[key];
        const feedback = result.feedback[key] || [];

        // Section heading
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
            children: [new TextRun({
                text: sectionHeaders[key],
                bold: true, size: 22, font: "Arial"
            })]
        }));

        // Score line (color-coded)
        const isFullScore = score >= max;
        children.push(new Paragraph({
            spacing: { after: 80 },
            children: [
                new TextRun({ text: `Score: `, bold: true, size: 20, font: "Arial" }),
                new TextRun({
                    text: `${score} / ${max}`,
                    bold: true, size: 20, font: "Arial",
                    color: isFullScore ? "008000" : "FF0000"
                }),
            ]
        }));

        // Feedback bullet items
        for (const item of feedback) {
            children.push(new Paragraph({
                spacing: { after: 60 },
                indent: { left: 360 },
                children: [new TextRun({ text: item, size: 20, font: "Arial" })]
            }));
        }
    }

    // --- Summary rubric table ---
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 200 },
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

    // --- Encouragement ---
    children.push(new Paragraph({ spacing: { before: 300 }, children: [] }));
    const pct = total / TOTAL_POINTS;
    let encouragement;
    if (pct >= 0.93) {
        encouragement = "Excellent work! You demonstrated a strong understanding of the material. Keep it up!";
    } else if (pct >= 0.83) {
        encouragement = "Good job overall! You have a solid foundation. Review the feedback above for areas where you can improve.";
    } else if (pct >= 0.67) {
        encouragement = "You made a good effort on this assignment. Please review the specific feedback for each section and make sure you understand the concepts. Feel free to reach out during office hours for help.";
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
    let outName = result.filename.replace(/\.xlsx$/i, '_FEEDBACK.docx');
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

After creating feedback documents, validate each one:
```bash
python scripts/office/validate.py <feedback_file.docx>
```

If `validate.py` is not available, at minimum check:
1. File size is > 0 bytes
2. File can be opened by `python-docx` or LibreOffice without errors
3. Convert one to PDF and visually inspect:
   ```bash
   libreoffice --headless --convert-to pdf <feedback_file.docx>
   ```
