# PowerPoint Generator for PDF Model Walkthroughs

This repository converts a PDF walkthrough into a PowerPoint deck using Python.

## What it does

- Reads `content/model_walkthrough.pdf`.
- Extracts page text and embedded images using **PyMuPDF**.
- Detects chart-like visuals with a simple image heuristic and prioritizes those visuals on slides.
- Builds a presentation with **python-pptx**.
- Saves the output to `output/model_walkthrough_presentation.pptx`.

## Project Structure

```text
content/
  model_walkthrough.pdf
scripts/
  generate_ppt.py
output/
```

## Setup

1. Create and activate a virtual environment (recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Place your source PDF at:

   ```text
   content/model_walkthrough.pdf
   ```

   > Fallback behavior: if that file is missing and there is exactly one PDF in `content/`, the script will use it.

2. Run the generator:

   ```bash
   python scripts/generate_ppt.py
   ```

3. Open the generated file:

   ```text
   output/model_walkthrough_presentation.pptx
   ```

## Notes

- Text extraction quality depends on the PDF (image-only/scanned PDFs may need OCR).
- Embedded images are extracted and one prioritized visual is placed per page slide.
- Chart detection is heuristic-based and may require manual slide edits after generation.
