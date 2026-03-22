from __future__ import annotations

from docx.oxml.ns import qn

_DEFAULT_FONT_NAME = "Microsoft YaHei"


def rgb(value: str):
    from docx.shared import RGBColor

    return RGBColor.from_string(value)


def clear_paragraph(paragraph) -> None:
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)


def format_run(
    run,
    *,
    font_name: str,
    font_size,
    bold: bool,
    color=None,
    italic: bool = False,
    underline: bool = False,
) -> None:
    run.font.name = font_name
    r_pr = run._element.get_or_add_rPr()
    r_pr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color is not None:
        run.font.color.rgb = color


def resolve_alignment(value: str | None, *, default):
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    return mapping.get(value or "", default)
