from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..macros import build_summary_card_group, expand_summary_card_block
from ..models import blocks as block_models
from ..models.document import DocumentModel
from .docx_utils import (
    _DEFAULT_CODE_FONT_NAME,
    _DEFAULT_FONT_NAME,
    format_run,
    resolve_alignment,
    rgb,
)
from .table_renderer import TableRenderer

if TYPE_CHECKING:
    from docx.document import Document as WordDocument

ColumnsBlock = block_models.ColumnsBlock
GroupBlock = block_models.GroupBlock
HeadingBlock = block_models.HeadingBlock
ImageBlock = block_models.ImageBlock
ListBlock = block_models.ListBlock
PageBreakBlock = block_models.PageBreakBlock
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
        "font_name": _DEFAULT_FONT_NAME,
    },
    "project_review": {
        "accent": "0F766E",
        "accent_soft": "D9F3EE",
        "title_size": 19,
        "heading_size": 13,
        "body_size": 10.5,
        "table_style": "metrics_compact",
        "summary_fill": "E8F6F3",
        "font_name": _DEFAULT_FONT_NAME,
    },
    "executive_brief": {
        "accent": "B45309",
        "accent_soft": "FDEBD8",
        "title_size": 18.5,
        "heading_size": 12.5,
        "body_size": 10.5,
        "table_style": "minimal",
        "summary_fill": "FFF7ED",
        "font_name": _DEFAULT_FONT_NAME,
    },
}


class WordDocumentBuilder:
    def __init__(self, table_renderer: TableRenderer | None = None) -> None:
        self._table_renderer = table_renderer or TableRenderer()

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

        workspace_dir = output_path.parent.resolve()
        for block in document_model.blocks:
            self._append_block(doc, block, theme, document_model, workspace_dir)

        doc.save(str(output_path))
        return output_path

    def _append_block(
        self,
        doc: WordDocument,
        block,
        theme: dict,
        document_model: DocumentModel,
        workspace_dir: Path,
    ) -> None:
        if isinstance(block, HeadingBlock):
            self._add_heading(doc, block, theme)
        elif isinstance(block, ParagraphBlock):
            self._add_paragraph(doc, block, theme, document_model, workspace_dir)
        elif isinstance(block, ListBlock):
            self._add_list(doc, block, theme)
        elif isinstance(block, TableBlock):
            self._table_renderer.render(
                doc,
                block,
                theme,
                default_table_style=getattr(
                    document_model.metadata, "table_template", ""
                ),
            )
        elif isinstance(block, ImageBlock):
            self._add_image(doc, block, workspace_dir)
        elif isinstance(block, GroupBlock):
            self._add_group(doc, block, theme, document_model, workspace_dir)
        elif isinstance(block, ColumnsBlock):
            self._add_columns(doc, block, theme, document_model, workspace_dir)
        elif isinstance(block, PageBreakBlock):
            self._add_page_break(doc)
        elif SummaryCardBlock is not None and isinstance(block, SummaryCardBlock):
            self._add_group(
                doc,
                expand_summary_card_block(block),
                theme,
                document_model,
                workspace_dir,
            )

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
                    "list_space_after": 4,
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
                    "list_space_after": 6,
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
        format_run(
            run,
            font_name=theme["font_name"],
            font_size=Pt(theme["title_size"]),
            bold=True,
            color=rgb(theme["accent"]),
        )

    def _add_heading(self, doc: WordDocument, block: HeadingBlock, theme: dict) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        paragraph = doc.add_paragraph()
        paragraph.alignment = resolve_alignment(
            getattr(block.style, "align", None),
            default=WD_ALIGN_PARAGRAPH.LEFT,
        )
        paragraph.paragraph_format.space_before = Pt(
            self._resolved_spacing(
                getattr(block.layout, "spacing_before", None),
                theme["heading_space_before"],
            )
        )
        paragraph.paragraph_format.space_after = Pt(
            self._resolved_spacing(
                getattr(block.layout, "spacing_after", None),
                theme["heading_space_after"],
            )
        )
        run = paragraph.add_run(block.text)
        format_run(
            run,
            font_name=theme["font_name"],
            font_size=Pt(
                self._scaled_size(
                    theme["heading_size"]
                    if block.level <= 1
                    else max(theme["body_size"] + 1, 11.5),
                    getattr(block.style, "font_scale", None),
                )
            ),
            bold=True,
            color=self._resolve_text_color(
                theme=theme,
                emphasis=getattr(block.style, "emphasis", None),
                default_color=rgb(theme["accent"]),
            ),
        )

    def _add_paragraph(
        self,
        doc: WordDocument,
        block: ParagraphBlock,
        theme: dict,
        document_model: DocumentModel,
        workspace_dir: Path,
    ) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.shared import Pt

        variant = getattr(block, "variant", "body")
        paragraph_text = self._paragraph_text(block)
        if variant in {"summary_box", "key_takeaway"}:
            card_title = getattr(block, "title", "") or (
                "Summary" if variant == "summary_box" else "Key Takeaway"
            )
            self._add_group(
                doc,
                build_summary_card_group(
                    title=card_title,
                    items=[paragraph_text],
                    variant="summary",
                    style=block.style,
                    layout=block.layout,
                ),
                theme,
                document_model,
                workspace_dir=workspace_dir,
            )
            return

        paragraph = doc.add_paragraph()
        paragraph.alignment = resolve_alignment(
            getattr(block.style, "align", None),
            default=WD_ALIGN_PARAGRAPH.LEFT,
        )
        paragraph.paragraph_format.first_line_indent = Pt(theme["body_indent"])
        paragraph.paragraph_format.space_before = Pt(
            self._resolved_spacing(getattr(block.layout, "spacing_before", None), 0)
        )
        paragraph.paragraph_format.space_after = Pt(
            self._resolved_spacing(
                getattr(block.layout, "spacing_after", None),
                theme["body_space_after"],
            )
        )
        paragraph.paragraph_format.line_spacing = theme["body_line_spacing"]
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        self._append_paragraph_runs(paragraph, block, theme)

    def _add_list(self, doc: WordDocument, block: ListBlock, theme: dict) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        for index, item in enumerate(block.items, start=1):
            paragraph = doc.add_paragraph()
            paragraph.alignment = resolve_alignment(
                getattr(block.style, "align", None),
                default=WD_ALIGN_PARAGRAPH.LEFT,
            )
            paragraph.paragraph_format.left_indent = Pt(
                max(theme["body_indent"] - 6, 12)
            )
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.space_before = Pt(
                self._resolved_spacing(
                    getattr(block.layout, "spacing_before", None),
                    0 if index > 1 else 0,
                )
            )
            paragraph.paragraph_format.space_after = Pt(
                self._resolved_spacing(
                    getattr(block.layout, "spacing_after", None),
                    theme["list_space_after"],
                )
            )
            marker = f"{index}. " if block.ordered else "• "
            run = paragraph.add_run(f"{marker}{item}")
            format_run(
                run,
                font_name=theme["font_name"],
                font_size=Pt(
                    self._scaled_size(
                        theme["body_size"], getattr(block.style, "font_scale", None)
                    )
                ),
                bold=self._resolved_bold(False, getattr(block.style, "emphasis", None)),
                color=self._resolve_text_color(
                    theme=theme,
                    emphasis=getattr(block.style, "emphasis", None),
                ),
            )

    def _add_group(
        self,
        doc: WordDocument,
        block: GroupBlock,
        theme: dict,
        document_model: DocumentModel,
        workspace_dir: Path,
    ) -> None:
        for child in block.blocks:
            self._append_block(doc, child, theme, document_model, workspace_dir)

    def _add_columns(
        self,
        doc: WordDocument,
        block: ColumnsBlock,
        theme: dict,
        document_model: DocumentModel,
        workspace_dir: Path,
    ) -> None:
        # Word multi-column layout requires section-level changes. For now we
        # keep the primitive in the model and render it sequentially as a safe
        # fallback instead of introducing layout-specific template branches.
        for column_index, column in enumerate(block.columns):
            if column_index > 0:
                doc.add_paragraph("")
            for child in column.blocks:
                self._append_block(doc, child, theme, document_model, workspace_dir)

    @staticmethod
    def _add_page_break(doc: WordDocument) -> None:
        doc.add_page_break()

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

    def _add_image(
        self, doc: WordDocument, block: ImageBlock, workspace_dir: Path
    ) -> None:
        from docx.shared import Inches

        image_path = self._resolve_workspace_image_path(block.path, workspace_dir)
        if image_path is None:
            return
        section = doc.sections[0]
        max_width = section.page_width - section.left_margin - section.right_margin
        picture = doc.add_picture(str(image_path))

        if block.width_px is not None:
            picture.width = min(Inches(block.width_px / 96.0), max_width)
        elif picture.width > max_width:
            picture.width = max_width

        if block.caption:
            doc.add_paragraph(block.caption)

    @staticmethod
    def _resolve_workspace_image_path(
        path_value: str, workspace_dir: Path
    ) -> Path | None:
        workspace_root = workspace_dir.resolve()
        candidate = Path(path_value)
        if not candidate.is_absolute():
            candidate = workspace_root / candidate

        resolved_candidate = candidate.resolve(strict=False)
        try:
            resolved_candidate.relative_to(workspace_root)
        except ValueError:
            return None

        return resolved_candidate if resolved_candidate.is_file() else None

    @staticmethod
    def _resolved_spacing(value: float | None, default: float) -> float:
        return default if value is None else value

    @staticmethod
    def _scaled_size(base_size: float, scale: float | None) -> float:
        return base_size if scale is None else base_size * scale

    @staticmethod
    def _resolved_bold(default_bold: bool, emphasis: str | None) -> bool:
        if emphasis == "strong":
            return True
        return default_bold

    @staticmethod
    def _paragraph_text(block: ParagraphBlock) -> str:
        if block.runs:
            return "".join(run.text for run in block.runs)
        return block.text

    def _append_paragraph_runs(self, paragraph, block: ParagraphBlock, theme: dict) -> None:
        from docx.shared import Pt

        font_size = Pt(
            self._scaled_size(
                theme["body_size"], getattr(block.style, "font_scale", None)
            )
        )
        default_color = self._resolve_text_color(
            theme=theme,
            emphasis=getattr(block.style, "emphasis", None),
            default_color=None,
        )

        if block.runs:
            for run_block in block.runs:
                font_name = (
                    _DEFAULT_CODE_FONT_NAME if run_block.code else theme["font_name"]
                )
                run = paragraph.add_run(run_block.text)
                format_run(
                    run,
                    font_name=font_name,
                    font_size=font_size,
                    bold=self._resolved_bold(
                        run_block.bold, getattr(block.style, "emphasis", None)
                    ),
                    italic=run_block.italic,
                    underline=run_block.underline,
                    color=default_color,
                )
        else:
            run = paragraph.add_run(self._paragraph_text(block))
            format_run(
                run,
                font_name=theme["font_name"],
                font_size=font_size,
                bold=self._resolved_bold(
                    False, getattr(block.style, "emphasis", None)
                ),
                color=default_color,
            )

    def _resolve_text_color(
        self,
        *,
        theme: dict,
        emphasis: str | None,
        default_color=None,
    ):
        if emphasis == "subtle":
            return rgb(theme["accent"])
        return default_color
