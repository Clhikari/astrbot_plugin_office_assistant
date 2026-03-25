from __future__ import annotations

from .docx_utils import (
    _DEFAULT_FONT_NAME,
    clear_paragraph,
    format_run,
    resolve_alignment,
    rgb,
)

DOCX_TABLE_STYLES = {
    "report_grid": "Table Grid",
    "metrics_compact": "Light List Accent 1",
    "minimal": "Table Grid",
}


class TableRenderer:
    def render(
        self,
        doc,
        block,
        theme: dict,
        *,
        default_table_style: str,
    ) -> None:
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm, Pt

        column_count = 0
        if block.headers:
            column_count = len(block.headers)
        elif block.rows:
            column_count = max(len(row) for row in block.rows)

        if column_count <= 0:
            return

        title_row_offset = 1 if block.caption else 0
        header_row_offset = 1 if block.headers else 0
        row_count = len(block.rows) + title_row_offset + header_row_offset
        table = doc.add_table(rows=row_count, cols=column_count)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        style_name = (
            getattr(block.style, "table_grid", None)
            or getattr(block, "table_style", "")
            or default_table_style
            or theme["table_style"]
        )
        resolved_style = self.resolve_docx_table_style(style_name)
        if resolved_style:
            try:
                table.style = resolved_style
            except (KeyError, ValueError):
                if style_name == "metrics_compact":
                    try:
                        table.style = DOCX_TABLE_STYLES["report_grid"]
                    except (KeyError, ValueError):
                        pass

        body_alignment = getattr(block.style, "cell_align", None)
        numeric_columns = set(getattr(block, "numeric_columns", []) or [])

        if getattr(block, "column_widths", None):
            table.autofit = False
            for col_index, width_cm in enumerate(block.column_widths[:column_count]):
                if width_cm <= 0:
                    continue
                target_width = Cm(width_cm)
                for row in table.rows:
                    row.cells[col_index].width = target_width

        row_index = 0
        if block.caption:
            title_cell = table.rows[0].cells[0]
            for col_index in range(1, column_count):
                title_cell = title_cell.merge(table.rows[0].cells[col_index])
            self._set_cell_text(
                title_cell,
                block.caption,
                bold=True,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                font_size=Pt(max(theme["body_size"], 11)),
                font_name=theme["font_name"],
                color=rgb(theme["accent"]),
                background=theme["accent_soft"],
            )
            row_index = 1

        if block.headers:
            header_row = table.rows[row_index]
            for col_index, value in enumerate(block.headers):
                self._set_cell_text(
                    header_row.cells[col_index],
                    value,
                    bold=True,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER,
                    font_size=Pt(self._table_font_size(style_name, theme, header=True)),
                    font_name=theme["font_name"],
                    color=self._table_header_color(style_name, theme),
                    background=self._header_fill(style_name, theme),
                )
            row_index += 1

        for data_row_index, values in enumerate(block.rows):
            current_row = row_index + data_row_index
            for col_index in range(column_count):
                default_alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                    if style_name == "metrics_compact" and col_index > 0
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                if body_alignment is None and col_index in numeric_columns:
                    default_alignment = WD_ALIGN_PARAGRAPH.RIGHT
                self._set_cell_text(
                    table.rows[current_row].cells[col_index],
                    values[col_index] if col_index < len(values) else "",
                    bold=False,
                    alignment=resolve_alignment(
                        body_alignment,
                        default=default_alignment,
                    ),
                    font_size=Pt(
                        self._table_font_size(style_name, theme, header=False)
                    ),
                    font_name=theme["font_name"],
                    background=self._table_row_fill(style_name, data_row_index + 1),
                )

    @staticmethod
    def resolve_docx_table_style(style_name: str) -> str | None:
        candidate = (style_name or "").strip()
        if not candidate:
            return DOCX_TABLE_STYLES["report_grid"]
        return DOCX_TABLE_STYLES.get(candidate, candidate)

    @staticmethod
    def _set_cell_text(
        cell,
        value: str,
        *,
        bold: bool,
        alignment,
        font_size,
        font_name: str = _DEFAULT_FONT_NAME,
        color=None,
        background: str | None = None,
    ) -> None:
        text = value or ""
        paragraph = cell.paragraphs[0]
        paragraph.alignment = alignment
        clear_paragraph(paragraph)
        run = paragraph.add_run(text)
        format_run(
            run,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            color=color,
        )
        paragraph.paragraph_format.space_after = 0
        if background:
            TableRenderer._set_cell_background(cell, background)

    @staticmethod
    def _header_fill(style_name: str, theme: dict) -> str | None:
        if style_name == "minimal":
            return theme["accent_soft"]
        return theme["accent"]

    @staticmethod
    def _table_header_color(style_name: str, theme: dict):
        if style_name == "minimal":
            return rgb(theme["accent"])
        return rgb("FFFFFF")

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
    def _set_cell_background(cell, fill: str) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        tc_pr = cell._tc.get_or_add_tcPr()
        shd = tc_pr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tc_pr.append(shd)
        shd.set(qn("w:fill"), fill)
