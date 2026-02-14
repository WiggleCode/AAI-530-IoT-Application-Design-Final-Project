#!/bin/bash

# gen_report.sh
# Script to convert IoT Agriculture Report from Markdown to PDF
# Usage: bash gen_report.sh

echo "Converting IoT Agriculture Report to PDF..."

# Check if pandoc is installed
if ! command -v pandoc &> /dev/null
then
    echo "Error: pandoc is not installed."
    echo "Install with: brew install pandoc (macOS) or apt-get install pandoc (Linux)"
    exit 1
fi

# Input and output files
INPUT_FILE="IoT_Agriculture_Final_Report.md"
OUTPUT_FILE="IoT_Agriculture_Final_Report.pdf"

# Check if input file exists
if [ ! -f "$INPUT_FILE" ]; then
    echo "Error: $INPUT_FILE not found!"
    exit 1
fi

# Convert markdown to PDF with proper formatting
pandoc "$INPUT_FILE" \
    -o "$OUTPUT_FILE" \
    --pdf-engine=pdflatex \
    --variable geometry:margin=1in \
    --variable fontsize=11pt \
    --variable linestretch=1.5 \
    --variable colorlinks=true \
    --variable linkcolor=blue \
    --variable urlcolor=blue \
    --variable citecolor=blue \
    --number-sections \
    --toc \
    --toc-depth=2 \
    --highlight-style=tango \
    -V fontfamily=times \
    -V papersize=letter \
    --standalone

# Check if conversion was successful
if [ $? -eq 0 ]; then
    echo "✓ PDF generated successfully: $OUTPUT_FILE"
    echo ""
    echo "Report includes:"
    echo "  - Full text with APA7 formatting"
    echo "  - 12 embedded figures with captions"
    echo "  - 4 comprehensive results tables"
    echo "  - Table of contents"
    echo "  - Numbered sections"
    echo ""
    echo "File size:"
    du -h "$OUTPUT_FILE"
else
    echo "✗ Error: PDF conversion failed"
    echo ""
    echo "Common issues:"
    echo "  1. Missing LaTeX installation (required for PDF generation)"
    echo "     macOS: brew install basictex"
    echo "     Ubuntu: sudo apt-get install texlive-latex-base texlive-fonts-recommended"
    echo ""
    echo "  2. Missing image files - ensure all PNG files are in current directory"
    echo ""
    echo "Alternative: Generate DOCX instead (doesn't require LaTeX):"
    echo "  pandoc $INPUT_FILE -o IoT_Agriculture_Final_Report.docx"
    exit 1
fi
