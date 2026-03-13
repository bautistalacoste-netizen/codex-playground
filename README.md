# PowerPoint Generator for Model Walkthroughs

This repository converts a Markdown model walkthrough into a clean, corporate PowerPoint presentation.

## Project Structure

```text
content/
  model_walkthrough.md
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

1. Edit `content/model_walkthrough.md` with your walkthrough content.
2. Run the generator:

   ```bash
   python scripts/generate_ppt.py
   ```

3. Open the generated file:

   ```text
   output/model_walkthrough_presentation.pptx
   ```

## Notes

- The script expects top-level title (`#`) and section headings (`##`) in Markdown.
- Bullet points (`- item`) and numbered items (`1. item`) are converted into slide content.
- The generator creates a title slide, agenda slide, section slides, and a closing "Next Steps" slide.
