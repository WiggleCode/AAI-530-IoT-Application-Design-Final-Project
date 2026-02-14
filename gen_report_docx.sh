#!/bin/bash

# gen_report_docx.sh
# Script to convert IoT Agriculture Report from Markdown to Word (DOCX)
# This is easier than PDF as it doesn't require LaTeX installation
# Usage: bash gen_report_docx.sh

echo "Converting IoT Agriculture Report to Word (DOCX)..."

# Check if pandoc is installed
if ! command -v pandoc &> /dev/null
then
    echo "Error: pandoc is not installed."
    echo "Install with: brew install pandoc (macOS) or apt-get install pandoc (Linux)"
    exit 1
fi

# Input and output files
INPUT_FILE="IoT_Agriculture_Final_Report.md"
OUTPUT_FILE="IoT_Agriculture_Final_Report.docx"

# Check if input file exists
if [ ! -f "$INPUT_FILE" ]; then
    echo "Error: $INPUT_FILE not found!"
    exit 1
fi

# Convert markdown to DOCX with proper formatting
pandoc "$INPUT_FILE" \
    -o "$OUTPUT_FILE" \
    --toc \
    --toc-depth=2 \
    --number-sections \
    --highlight-style=tango \
    --standalone

# Check if conversion was successful
if [ $? -eq 0 ]; then
    echo "✓ Word document generated successfully: $OUTPUT_FILE"
    echo ""
    echo "Report includes:"
    echo "  - Full text with formatting"
    echo "  - 12 embedded figures with captions"
    echo "  - 4 comprehensive results tables"
    echo "  - Table of contents"
    echo "  - Numbered sections"
    echo ""
    echo "File size:"
    du -h "$OUTPUT_FILE"
    echo ""
    echo "Next steps:"
    echo "  1. Open $OUTPUT_FILE in Microsoft Word or Google Docs"
    echo "  2. Adjust formatting to match APA7 requirements (fonts, spacing)"
    echo "  3. Export as PDF from Word/Docs"
else
    echo "✗ Error: DOCX conversion failed"
    exit 1
fi
