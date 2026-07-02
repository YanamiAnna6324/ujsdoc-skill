---
name: ujs-thesis-format
description: Format, edit, create, and verify Jiangsu University undergraduate graduation design/thesis documents according to the uploaded 2026 江苏大学毕业设计（论文）写作规范 PDF. Use when working on DOCX/PDF thesis drafts, thesis templates, captions, tables, figures, references, abstracts, page layout, or final formatting for Jiangsu University undergraduate graduation design or thesis documents.
---

# UJS Thesis Format

## Required Workflow

1. Read `references/format-requirements.md` before creating, editing, or checking a document.
2. Treat `assets/jiangsu-university-thesis-writing-standard-2026.pdf` as the source of truth. If a user provides a newer college template or attachment template, prefer that template for exact styles and use this skill for the remaining rules.
3. For new DOCX files, start from `assets/ujs-thesis-template.docx` when possible. If it is missing or unsuitable, run `python scripts/create_docx_template.py <output.docx>` to regenerate a starter template.
4. For existing DOCX files, preserve content first, then normalize page setup, headings, captions, references, figures, tables, and image placement to the specification.
5. Before delivery, run `python scripts/validate_docx_format.py <document.docx>`. Fix all `ERROR` findings and review every `WARN` finding.
6. If delivering PDF, export the final DOCX to PDF, render representative pages, and visually inspect margins, header line, footer page numbers, captions, tables, images, and Chinese text rendering.

## Formatting Priorities

- Use A4 pages with top/left margins of 25 mm and bottom/right margins of 20 mm.
- Keep left-side binding in mind; do not move page content into the binding area.
- Use Chinese thesis typography: FangSong-style body text, SimHei-style headings, Times New Roman for Latin text and numbers where appropriate. The source PDF also uses FZXiaoBiaoSong, KaiTi, FangSong, SimHei, and Times New Roman.
- Unless the user or an official college template explicitly requires another color, set all font colors to black (`#000000`), including body text, headings, captions, tables, references, headers, and footers.
- Keep a left-aligned header with a horizontal rule and centered Arabic page numbers in the footer when matching the source PDF style.
- Use first-line indentation for Chinese paragraphs, consistent line spacing, and no overlapping or clipped text.
- Number figures as `图 1`, `图 2`, ... with titles below figures.
- Number tables as `表 1`, `表 2`, ... with titles above tables.
- Number formulas consecutively with Arabic numerals and place formula numbers at the far right of the formula line.
- Use GB/T 7714-2015 reference formatting and bracketed citation numbers such as `[1]`.

## Image And Table Rules

- Keep every image sharp, centered, and inside the printable text area.
- Add a concise figure title under every image. Do not leave decorative or unexplained images in the thesis.
- Refer to each figure/table from the surrounding text before or near its appearance.
- Keep table numbers and titles above tables. Put table notes directly below the title or below the table as needed.
- Align numbers vertically within the same table column.
- Never use `同上`, `同左`, `；`, or similar placeholders inside tables; write the actual value.

## Bundled Resources

- `references/format-requirements.md`: concise requirements extracted from the 2026 PDF.
- `assets/jiangsu-university-thesis-writing-standard-2026.pdf`: uploaded source PDF.
- `assets/ujs-thesis-template.docx`: starter DOCX template with A4 layout, margins, header/footer, and baseline styles.
- `scripts/create_docx_template.py`: regenerate the starter DOCX template.
- `scripts/validate_docx_format.py`: check a DOCX for page setup, required sections, figure/table captions, images, keywords, and references.
