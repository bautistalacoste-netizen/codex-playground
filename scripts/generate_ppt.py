from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import List

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

INPUT_PATH = Path("content/model_walkthrough.md")
OUTPUT_PATH = Path("output/model_walkthrough_presentation.pptx")


@dataclass
class Section:
    title: str
    bullets: List[str] = field(default_factory=list)
    paragraphs: List[str] = field(default_factory=list)


def parse_markdown(markdown_text: str) -> tuple[str, str, List[Section]]:
    lines = markdown_text.splitlines()

    title = "Model Walkthrough"
    subtitle = "Executive Presentation"
    sections: List[Section] = []
    current_section: Section | None = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        h1_match = re.match(r"^#\s+(.+)", line)
        if h1_match:
            title = h1_match.group(1).strip()
            continue

        h2_match = re.match(r"^##\s+(.+)", line)
        if h2_match:
            current_section = Section(title=h2_match.group(1).strip())
            sections.append(current_section)
            continue

        if line.startswith("- ") or re.match(r"^\d+\.\s+", line):
            bullet_text = re.sub(r"^(?:-\s+|\d+\.\s+)", "", line).strip()
            if current_section is None:
                current_section = Section(title="Overview")
                sections.append(current_section)
            current_section.bullets.append(bullet_text)
            continue

        if current_section is None:
            subtitle = line
        else:
            current_section.paragraphs.append(line)

    if not sections:
        sections = [
            Section(
                title="Overview",
                bullets=["Add section headings and bullets to content/model_walkthrough.md"],
            )
        ]

    return title, subtitle, sections


def apply_corporate_style(prs: Presentation) -> None:
    # Use 16:9 layout for modern presentation displays
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)


def add_title_slide(prs: Presentation, title: str, subtitle: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_box = slide.shapes.title
    subtitle_box = slide.placeholders[1]

    title_box.text = title
    subtitle_box.text = subtitle

    title_p = title_box.text_frame.paragraphs[0]
    title_p.font.name = "Calibri"
    title_p.font.size = Pt(40)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(15, 56, 107)

    subtitle_p = subtitle_box.text_frame.paragraphs[0]
    subtitle_p.font.name = "Calibri"
    subtitle_p.font.size = Pt(20)
    subtitle_p.font.color.rgb = RGBColor(90, 90, 90)


def add_agenda_slide(prs: Presentation, section_titles: List[str]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Agenda"
    body = slide.shapes.placeholders[1].text_frame
    body.clear()

    for idx, section_title in enumerate(section_titles, start=1):
        p = body.add_paragraph() if idx > 1 else body.paragraphs[0]
        p.text = f"{idx}. {section_title}"
        p.font.name = "Calibri"
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(40, 40, 40)


def add_content_slide(prs: Presentation, section: Section) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = section.title

    title_p = slide.shapes.title.text_frame.paragraphs[0]
    title_p.font.name = "Calibri"
    title_p.font.size = Pt(32)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(15, 56, 107)

    text_frame = slide.shapes.placeholders[1].text_frame
    text_frame.clear()

    content_items = section.paragraphs + section.bullets
    if not content_items:
        content_items = ["No details provided in this section."]

    for idx, item in enumerate(content_items):
        paragraph = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
        paragraph.text = item
        paragraph.level = 0
        paragraph.font.name = "Calibri"
        paragraph.font.size = Pt(20)
        paragraph.font.color.rgb = RGBColor(50, 50, 50)


def add_closing_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.title
    title_shape.text = "Next Steps"

    p = title_shape.text_frame.paragraphs[0]
    p.font.name = "Calibri"
    p.font.size = Pt(42)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 56, 107)
    p.alignment = PP_ALIGN.CENTER


def generate_presentation(input_path: Path = INPUT_PATH, output_path: Path = OUTPUT_PATH) -> Path:
    markdown_text = input_path.read_text(encoding="utf-8")
    title, subtitle, sections = parse_markdown(markdown_text)

    prs = Presentation()
    apply_corporate_style(prs)
    add_title_slide(prs, title, subtitle)
    add_agenda_slide(prs, [section.title for section in sections])

    for section in sections:
        add_content_slide(prs, section)

    add_closing_slide(prs)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(
            "Missing input file: content/model_walkthrough.md. "
            "Add your walkthrough markdown before running this script."
        )

    generated_path = generate_presentation()
    print(f"Presentation generated at: {generated_path}")


if __name__ == "__main__":
    main()
