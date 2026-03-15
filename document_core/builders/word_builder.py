from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..models import blocks as block_models
from ..models.document import DocumentModel

if TYPE_CHECKING:
    from docx.document import Document as WordDocument

HeadingBlock = block_models.HeadingBlock
ImageBlock = block_models.ImageBlock
ParagraphBlock = block_models.ParagraphBlock
TableBlock = block_models.TableBlock
SummaryCardBlock = getattr(block_models, "SummaryCardBlock", None)

THEMES = {
    "business_report": {
        "accent": "1F4E79",
        "accent_soft": "DCE6F1",
        "title_size": 20,
        "heading_size": 14,
        "body_size": 11,
        "table_style": "report_grid",
        "summary_fill": "EEF4FA",
    },
    "project_review": {
        "accent": "0F766E",
        "accent_soft": "D9F3EE",
        "title_size": 19,
        "heading_size": 13,
        "body_size": 10.5,
        "table_style": "metrics_compact",
        "summary_fill": "E8F6F3",
    },
    "executive_brief": {
        "accent": "B45309",
        "accent_soft": "FDEBD8",
        "title_size": 18.5,
        "heading_size": 12.5,
        "body_size": 10.5,
        "table_style": "minimal",
        "summary_fill": "FFF7ED",
    },
}


class WordDocumentBuilder:
    def build(self, document_model: DocumentModel, output_path: Path) -> Path:
        try:
            from docx import Document
            from docx.enum.section import WD_SECTION
            from docx.shared import Cm
        except ImportError as exc:
            raise RuntimeError(
                "python-docx is required to export Word documents."
            ) from exc

        output_path.parent.mkdir(parents=True, exist_ok=True)

        doc = Document()
        section = doc.sections[0]
        section.start_type = WD_SECTION.NEW_PAGE
        theme = self._resolve_theme(document_model)
        section.top_margin = Cm(theme["margins"]["top"])
        section.bottom_margin = Cm(theme["margins"]["bottom"])
        section.left_margin = Cm(theme["margins"]["left"])
        section.right_margin = Cm(theme["margins"]["right"])

        if document_model.metadata.title:
            self._add_title(doc, document_model.metadata.title, theme)

        for block in document_model.blocks:
            if isinstance(block, HeadingBlock):
                self._add_heading(doc, block, theme)
            elif isinstance(block, ParagraphBlock):
                self._add_paragraph(doc, block, theme)
            elif isinstance(block, TableBlock):
                self._add_table(doc, block, theme, document_model)
            elif isinstance(block, ImageBlock):
                self._add_image(doc, block)
            elif SummaryCardBlock is not None and isinstance(block, SummaryCardBlock):
                self._add_summary_card(doc, block, theme)

        doc.save(str(output_path))
        return output_path

    def _resolve_theme(self, document_model: DocumentModel) -> dict:
        theme_name = getattr(document_model.metadata, "theme_name", "business_report")
        theme = dict(THEMES.get(theme_name, THEMES["business_report"]))
        density = getattr(document_model.metadata, "density", "comfortable")
        if density == "compact":
            theme.update(
                {
                    "margins": {"top": 2.2, "bottom": 2.1, "left": 2.4, "right": 2.3},
                    "title_spacing_after": 14,
                    "heading_space_before": 10,
                    "heading_space_after": 5,
                    "body_indent": 18,
                    "body_space_after": 6,
                    "body_line_spacing": 1.2,
                    "table_font_size": 9.5,
                }
            )
        else:
            theme.update(
                {
                    "margins": {"top": 2.8, "bottom": 2.6, "left": 2.8, "right": 2.6},
                    "title_spacing_after": 18,
                    "heading_space_before": 14,
                    "heading_space_after": 8,
                    "body_indent": 24,
                    "body_space_after": 10,
                    "body_line_spacing": 1.5,
                    "table_font_size": 10.5,
                }
            )

        accent_color = getattr(document_model.metadata, "accent_color", "")
        if accent_color:
            theme["accent"] = accent_color
            theme["accent_soft"] = self._blend_hex(accent_color, "FFFFFF", 0.84)
            theme["summary_fill"] = self._blend_hex(accent_color, "FFFFFF", 0.92)

        return theme

    def _add_title(self, doc: WordDocument, text: str, theme: dict) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_after = Pt(theme["title_spacing_after"])
        run = paragraph.add_run(text)
        self._format_run(
            run,
            font_name="Microsoft YaHei",
            font_size=Pt(theme["title_size"]),
            bold=True,
            color=self._rgb(theme["accent"]),
        )

    def _add_heading(self, doc: WordDocument, block: HeadingBlock, theme: dict) -> None:
        from docx.shared import Pt

        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(theme["heading_space_before"])
        paragraph.paragraph_format.space_after = Pt(theme["heading_space_after"])
        run = paragraph.add_run(block.text)
        self._format_run(
            run,
            font_name="Microsoft YaHei",
            font_size=Pt(
                theme["heading_size"]
                if block.level <= 1
                else max(theme["body_size"] + 1, 11.5)
            ),
            bold=True,
            color=self._rgb(theme["accent"]),
        )

    def _add_paragraph(
        self, doc: WordDocument, block: ParagraphBlock, theme: dict
    ) -> None:
        from docx.enum.text import WD_LINE_SPACING
        from docx.shared import Pt

        variant = getattr(block, "variant", "body")
        if variant in {"summary_box", "key_takeaway"}:
            card_title = getattr(block, "title", "") or (
                "Summary" if variant == "summary_box" else "Key Takeaway"
            )
            self._add_summary_card(
                doc,
                SimpleSummaryCard(title=card_title, items=[block.text]),
                theme,
            )
            return

        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.first_line_indent = Pt(theme["body_indent"])
        paragraph.paragraph_format.space_after = Pt(theme["body_space_after"])
        paragraph.paragraph_format.line_spacing = theme["body_line_spacing"]
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        run = paragraph.add_run(block.text)
        self._format_run(
            run,
            font_name="Microsoft YaHei",
            font_size=Pt(theme["body_size"]),
            bold=False,
        )

    def _add_table(
        self,
        doc: WordDocument,
        block: TableBlock,
        theme: dict,
        document_model: DocumentModel,
    ) -> None:
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        column_count = 0
        if block.headers:
            column_count = len(block.headers)
        elif block.rows:
            column_count = max(len(row) for row in block.rows)

        if column_count <= 0:
            return

        row_count = len(block.rows) + (1 if block.headers else 0)
        table = doc.add_table(rows=row_count, cols=column_count)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        style_name = getattr(block, "table_style", "") or getattr(
            document_model.metadata, "table_template", ""
        )
        style_name = style_name or theme["table_style"]
        try:
            table.style = "Table Grid"
        except (KeyError, ValueError):
            pass

        row_index = 0
        if block.headers:
            for col_index, value in enumerate(block.headers):
                self._set_cell_text(
                    table.rows[0].cells[col_index],
                    value,
                    bold=True,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER,
                    font_size=Pt(self._table_font_size(style_name, theme, header=True)),
                    color=self._table_header_color(style_name, theme),
                    background=self._header_fill(style_name, theme),
                )
            row_index = 1

        for current_row, values in enumerate(block.rows, start=row_index):
            for col_index in range(column_count):
                self._set_cell_text(
                    table.rows[current_row].cells[col_index],
                    values[col_index] if col_index < len(values) else "",
                    bold=False,
                    alignment=(
                        WD_ALIGN_PARAGRAPH.CENTER
                        if style_name == "metrics_compact" and col_index > 0
                        else WD_ALIGN_PARAGRAPH.LEFT
                    ),
                    font_size=Pt(
                        self._table_font_size(style_name, theme, header=False)
                    ),
                    background=self._table_row_fill(style_name, current_row),
                )

    @staticmethod
    def _set_cell_text(
        cell,
        value: str,
        *,
        bold: bool,
        alignment,
        font_size,
        color=None,
        background: str | None = None,
    ) -> None:
        text = value or ""
        paragraph = cell.paragraphs[0]
        paragraph.alignment = alignment
        paragraph.clear()
        run = paragraph.add_run(text)
        WordDocumentBuilder._format_run(
            run,
            font_name="Microsoft YaHei",
            font_size=font_size,
            bold=bold,
            color=color,
        )
        paragraph.paragraph_format.space_after = 0
        if background:
            WordDocumentBuilder._set_cell_background(cell, background)

    @staticmethod
    def _format_run(run, *, font_name: str, font_size, bold: bool, color=None) -> None:
        from docx.oxml.ns import qn

        run.font.name = font_name
        r_pr = run._element.get_or_add_rPr()
        r_pr.rFonts.set(qn("w:eastAsia"), font_name)
        run.font.size = font_size
        run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color

    def _add_summary_card(self, doc: WordDocument, block, theme: dict) -> None:
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        items = list(getattr(block, "items", []) or [])
        if not items:
            return

        variant = getattr(block, "variant", "summary") or "summary"
        table = doc.add_table(rows=len(items) + 1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        try:
            table.style = "Table Grid"
        except (KeyError, ValueError):
            pass

        title = getattr(block, "title", "") or (
            "Conclusion" if variant == "conclusion" else "Summary"
        )
        self._set_cell_text(
            table.rows[0].cells[0],
            title,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            font_size=Pt(11),
            color=self._summary_title_color(variant, theme),
            background=self._summary_title_fill(variant, theme),
        )

        for row_index, item in enumerate(items, start=1):
            self._set_cell_text(
                table.rows[row_index].cells[0],
                f"• {item}",
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                font_size=Pt(theme["body_size"]),
                color=self._summary_body_color(variant, theme),
                background=self._summary_body_fill(variant, theme),
            )

    @staticmethod
    def _rgb(value: str):
        from docx.shared import RGBColor

        return RGBColor.from_string(value)

    @staticmethod
    def _header_fill(style_name: str, theme: dict) -> str | None:
        if style_name == "minimal":
            return None
        return theme["accent"]

    @staticmethod
    def _table_header_color(style_name: str, theme: dict):
        if style_name == "minimal":
            return WordDocumentBuilder._rgb(theme["accent"])
        return WordDocumentBuilder._rgb("FFFFFF")

    @staticmethod
    def _table_row_fill(style_name: str, row_index: int) -> str | None:
        if style_name == "report_grid" and row_index % 2 == 1:
            return "F7FBFF"
        return None

    @staticmethod
    def _table_font_size(style_name: str, theme: dict, *, header: bool) -> float:
        base_size = theme["table_font_size"]
        if style_name == "metrics_compact":
            return max(base_size - 0.5, 9)
        if style_name == "minimal" and header:
            return max(base_size, theme["body_size"])
        return base_size

    @staticmethod
    def _summary_title_fill(variant: str, theme: dict) -> str:
        if variant == "conclusion":
            return theme["accent_soft"]
        return theme["accent"]

    @staticmethod
    def _summary_body_fill(variant: str, theme: dict) -> str:
        if variant == "conclusion":
            return "FFFFFF"
        return theme["summary_fill"]

    @staticmethod
    def _summary_title_color(variant: str, theme: dict):
        if variant == "conclusion":
            return WordDocumentBuilder._rgb(theme["accent"])
        return WordDocumentBuilder._rgb("FFFFFF")

    @staticmethod
    def _summary_body_color(variant: str, theme: dict):
        if variant == "conclusion":
            return WordDocumentBuilder._rgb("444444")
        return None

    @staticmethod
    def _set_cell_background(cell, fill: str) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        tc_pr = cell._tc.get_or_add_tcPr()
        shd = tc_pr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tc_pr.append(shd)
        shd.set(qn("w:fill"), fill)

    @staticmethod
    def _blend_hex(source: str, target: str, ratio: float) -> str:
        ratio = min(max(ratio, 0.0), 1.0)
        source_channels = [int(source[index : index + 2], 16) for index in (0, 2, 4)]
        target_channels = [int(target[index : index + 2], 16) for index in (0, 2, 4)]
        blended = [
            round(src * (1 - ratio) + dst * ratio)
            for src, dst in zip(source_channels, target_channels, strict=False)
        ]
        return "".join(f"{channel:02X}" for channel in blended)

    def _add_image(self, doc: WordDocument, block: ImageBlock) -> None:
        image_path = Path(block.path)
        if not image_path.exists():
            return
        doc.add_picture(str(image_path))
        if block.caption:
            doc.add_paragraph(block.caption)


class SimpleSummaryCard:
    def __init__(self, title: str, items: list[str]) -> None:
        self.title = title
        self.items = items
