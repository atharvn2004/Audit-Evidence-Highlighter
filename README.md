# Audit Evidence Highlighter

**A Python automation tool designed to streamline the audit documentation process.**

## Project Overview
This tool automates the manual process of highlighting audit evidence. It takes a source file (PDF, Image, Word, or Excel) and a search string, locating the text and overlaying a red bounding box (or border) to flag the evidence without altering the original content.

## Key Features
- **PDF Analysis:** Uses Coordinate Geometry (PyMuPDF) to draw vector bounding boxes over text.
- **OCR Integration:** Uses Tesseract and OpenCV to detect and highlight text in raster images.
- **Office Support:** XML manipulation for MS Word and Cell Styling for MS Excel.
- **Audit Compliance:** Ensures the underlying document content remains editable and valid.

## Tech Stack
- **Language:** Python 3.12
- **Libraries:** `pymupdf` (PDF), `opencv-python` (Images), `pytesseract` (OCR), `python-docx` (Word), `openpyxl` (Excel).

## Usage
1. Run `audit_tool.py`.
2. Input the target filename.
3. Input the text to search (e.g., "Morgan").
4. The tool generates a new file in `audit_results/` with the evidence highlighted.
