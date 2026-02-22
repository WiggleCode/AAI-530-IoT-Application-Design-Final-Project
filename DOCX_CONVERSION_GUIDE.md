# MD to APA7 DOCX Conversion Guide

**Project**: AAI-530 IoT Agriculture Final Report
**Team**: Dylan Scott-Dawkins, Francisco Monarrez Felix, Jeffery Smith

---

## Files Involved

```
tiot/
├── IoT_Agriculture_Final_Report.md          ← source (edited report)
├── reference.docx                           ← APA7 style template
├── create_apa7_reference.py                 ← generates reference.docx
├── apa7_format.py                           ← APA7 post-processing script
└── IoT_Agriculture_Final_Report_Final.docx  ← final output
```

---

## One-Time Setup

### Install Pandoc
```bash
brew install pandoc
pandoc --version   # verify
```

### Regenerate reference.docx (only needed if you change base styles)
```bash
python3 create_apa7_reference.py
```

`reference.docx` is already generated and ready to use. Only re-run this if
you want to change the base font, spacing, or heading styles.

---

## Conversion (Run Every Time)

### Step 1 — Convert Markdown to raw DOCX
```bash
pandoc IoT_Agriculture_Final_Report.md \
  --reference-doc=reference.docx \
  -o paper_raw.docx
```

Pandoc maps markdown elements to Word styles defined in `reference.docx`:

| Markdown | Word Style |
|---|---|
| `# Heading` | Heading 1 — bold, centered |
| `## Heading` | Heading 2 — bold, left-aligned |
| `### Heading` | Heading 3 — bold italic, left-aligned |
| Regular text | Normal — TNR 12pt, double-spaced |
| `> blockquote` | Block Text — 0.5" indent both sides |
| ` ```code``` ` | Source Code — Courier New 10pt, single-spaced |
| Tables | Word table (borders fixed in Step 2) |
| References section | Bibliography — hanging indent |

### Step 2 — Apply remaining APA7 rules
```bash
python3 apa7_format.py paper_raw.docx IoT_Agriculture_Final_Report_Final.docx
```

This script enforces rules that cannot live in a Word style:
- Table borders: top, below-header row, and bottom only (no vertical lines)
- Section-aware first-line indentation (body, abstract, references)
- 0.5" hanging indent on every reference entry
- No indent on first paragraph after each heading
- No indent on Table/Figure labels, *Note.* lines, Keywords line
- Font and double-spacing as a safety net over Pandoc output
- **Removes all bookmarks** inserted automatically by Pandoc (Pandoc adds
  a bookmark to every heading for internal linking — these are stripped)

### Combined single command
```bash
cd /Users/dscottdawkins/usd/tiot

pandoc IoT_Agriculture_Final_Report.md \
  --reference-doc=reference.docx \
  -o paper_raw.docx \
&& python3 apa7_format.py paper_raw.docx IoT_Agriculture_Final_Report_Final.docx \
&& rm paper_raw.docx
```

The `&&` ensures each step only runs if the previous one succeeded.
`rm paper_raw.docx` removes the intermediate file automatically.

---

## Step 3 — Manual Steps in Word

Open `IoT_Agriculture_Final_Report_Final.docx` and complete these four steps:

1. **Page numbers**
   Insert → Page Number → Top of Page → Plain Number 3 (top right corner)

2. **Title page**
   Add: student names, course (AAI-530: IoT Application Design),
   instructor name, University of San Diego, submission date

3. **Figure labels**
   Find each figure caption (e.g. `Figure 1`) and change the label text
   from *italic* to **bold** — APA7 requires bold figure labels

4. **Note. lines**
   Below each table, confirm only the word `Note.` is italic,
   not the full sentence that follows it

---

## What Each Layer Guarantees

| Layer | Handles |
|---|---|
| `reference.docx` | Base styles — TNR 12pt, double spacing, heading alignment baked into every Word style Pandoc maps to |
| Pandoc | Structure — maps markdown elements to the correct pre-styled Word styles |
| `apa7_format.py` | Rules — table borders, indent logic, hanging indents, section detection |
| Word (manual) | Page-level — page numbers, title page, figure label boldness |

---

## APA7 Styles in reference.docx

| Style | Font | Spacing | Indent |
|---|---|---|---|
| Normal | TNR 12pt | Double | 0.50" first-line |
| Body Text | TNR 12pt | Double | 0.50" first-line |
| First Paragraph | TNR 12pt | Double | 0" (no indent) |
| Heading 1 | TNR 12pt bold | Double | 0" centered |
| Heading 2 | TNR 12pt bold | Double | 0" left |
| Heading 3 | TNR 12pt bold-italic | Double | 0" left |
| Heading 4 | TNR 12pt italic | Double | 0.50" first-line |
| Block Text | TNR 12pt | Double | 0.50" left + right |
| Source Code | Courier New 10pt | Single | 0.50" left |
| Caption | TNR 12pt | Double | 0" |
| Bibliography | TNR 12pt | Double | −0.50" hanging |
| Verbatim Char | Courier New 10pt | — | character style |

---

## Troubleshooting

**Pandoc not found**
```bash
brew install pandoc
```

**python-docx not installed**
```bash
pip3 install python-docx
```

**reference.docx missing**
```bash
python3 create_apa7_reference.py
```

**Table borders still showing vertical lines**
The `apa7_format.py` script handles this automatically. If borders are wrong,
re-run Step 2. If still wrong, select the table in Word →
Table Design → Borders → No Border, then manually add top, header-bottom,
and bottom borders.

**Font not applying to some paragraphs**
Some styled elements (e.g. code blocks) are intentionally left unchanged.
If body text shows the wrong font, select all (Cmd+A) in Word and apply
Times New Roman 12pt, then reapply double spacing.
