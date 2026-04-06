from __future__ import annotations

from collections.abc import Callable, Sequence
from datetime import datetime, timedelta, timezone
from pathlib import Path
import re
from threading import RLock
from uuid import uuid4

from astrbot.api import logger
from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path

from ...document_core.macros import summary_card_defaults_from_config
from ...document_core.models.blocks import (
    AccentBoxBlock,
    ColumnBlock,
    ColumnsBlock,
    DocumentBlock,
    GroupBlock,
    HeadingBlock,
    HeroBannerBlock,
    ListBlock,
    MetricCard,
    MetricCardsBlock,
    PageBreakBlock,
    ParagraphBlock,
    SectionBreakBlock,
    SummaryCardBlock,
    TableBlock,
    TocBlock,
)
from ...document_core.models.document import (
    DocumentMetadata,
    DocumentModel,
    DocumentStatus,
    DocumentSummaryCardDefaults,
)
from .contracts import (
    AddBlocksRequest,
    AddHeadingRequest,
    AddListRequest,
    AddPageBreakRequest,
    AddParagraphRequest,
    AddSectionBundleRequest,
    AddSummaryCardRequest,
    AddTableRequest,
    BlockInput,
    BlockColumnsInput,
    BlockGroupInput,
    BlockHeadingInput,
    CreateDocumentRequest,
    ExportDocumentRequest,
    FinalizeDocumentRequest,
    SectionBreakInput,
    SectionAccentBoxInput,
    SectionCardInput,
    SectionHeroBannerInput,
    SectionListInput,
    SectionMetricCardsInput,
    SectionPageBreakInput,
    SectionParagraphInput,
    SectionTableInput,
    TocInput,
    _normalize_docx_filename,
)
from .hooks import (
    BlockNormalizationContext,
    BlockNormalizeHook,
    run_block_normalize_hooks,
)

PLUGIN_NAME = "astrbot_plugin_office_assistant"

BLOCK_TYPE_HEADING = "heading"
BLOCK_TYPE_PARAGRAPH = "paragraph"
BLOCK_TYPE_LIST = "list"
BLOCK_TYPE_TABLE = "table"
BLOCK_TYPE_SUMMARY_CARD = "summary_card"
BLOCK_TYPE_PAGE_BREAK = "page_break"
MAX_HEADING_LENGTH_FOR_TABLE_TITLE = 24
TABLE_CAPTION_NUMBERED_HEADING_RE = re.compile(
    r"^(?:[IVXLCDM]+\.\s+|\d+(?:\.\d+)*[.)]?\s+|[一二三四五六七八九十百千万]+、)"
)
SummaryCardDefaultsResolver = Callable[
    [DocumentSummaryCardDefaults | None],
    dict[str, object | None],
]
RuntimeBlock = DocumentBlock


def _default_workspace_dir() -> Path:
    return Path(get_astrbot_plugin_data_path()) / PLUGIN_NAME / "documents"


def _is_within_workspace(path: Path, workspace_dir: Path) -> bool:
    try:
        path.relative_to(workspace_dir)
        return True
    except ValueError:
        return False


def _looks_like_numbered_heading(text: str) -> bool:
    return bool(TABLE_CAPTION_NUMBERED_HEADING_RE.match(text.strip()))


class DocumentSessionStore:
    def __init__(
        self,
        workspace_dir: Path | None = None,
        *,
        max_documents: int | None = 256,
        ttl: timedelta | None = None,
        normalize_block_hooks: Sequence[BlockNormalizeHook] | None = None,
        summary_card_defaults_resolver: SummaryCardDefaultsResolver = (
            summary_card_defaults_from_config
        ),
    ) -> None:
        self._lock = RLock()
        self._documents: dict[str, DocumentModel] = {}
        self.workspace_dir = workspace_dir or _default_workspace_dir()
        self.workspace_dir.mkdir(parents=True, exist_ok=True)
        self._max_documents = max_documents
        self._ttl = ttl
        self._summary_card_defaults_resolver = summary_card_defaults_resolver
        self._normalize_block_hooks = normalize_block_hooks or [
            self._drop_duplicate_document_title_headings,
            self._move_landscape_intro_paragraphs_before_section_break,
            self._promote_heading_before_table_to_caption,
        ]

    def _evict_expired_locked(self) -> None:
        if self._ttl is None:
            return

        now = datetime.now(timezone.utc)
        expired_ids = [
            document_id
            for document_id, document in self._documents.items()
            if document.metadata.updated_at + self._ttl < now
        ]
        for document_id in expired_ids:
            self._documents.pop(document_id, None)

    def _evict_excess_locked(self) -> None:
        if self._max_documents is None:
            return

        excess = len(self._documents) - self._max_documents
        if excess <= 0:
            return

        oldest_documents = sorted(
            self._documents.items(),
            key=lambda item: item[1].metadata.updated_at,
        )
        for document_id, _ in oldest_documents[:excess]:
            self._documents.pop(document_id, None)

    def _prune_locked(self) -> None:
        self._evict_expired_locked()
        self._evict_excess_locked()

    def create_document(self, request: CreateDocumentRequest) -> DocumentModel:
        with self._lock:
            document_id = uuid4().hex
            document = DocumentModel(
                document_id=document_id,
                session_id=request.session_id,
                metadata=DocumentMetadata(
                    title=request.title,
                    preferred_filename=request.output_name,
                    theme_name=request.theme_name,
                    table_template=request.table_template,
                    density=request.density,
                    accent_color=request.accent_color,
                    header_footer=request.header_footer,
                    document_style=request.document_style,
                ),
            )
            self._documents[document_id] = document
            self._prune_locked()
            return document

    def get_document(self, document_id: str) -> DocumentModel | None:
        with self._lock:
            self._evict_expired_locked()
            return self._documents.get(document_id)

    def require_document(self, document_id: str) -> DocumentModel:
        document = self.get_document(document_id)
        if document is None:
            raise KeyError(f"Document not found: {document_id}")
        return document

    def add_blocks(self, request: AddBlocksRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            self._append_blocks_locked(document, request.blocks)
            return document

    def _append_blocks_locked(
        self, document: DocumentModel, blocks: list[BlockInput]
    ) -> None:
        normalized_blocks: list[BlockInput] = run_block_normalize_hooks(
            self._normalize_block_hooks,
            BlockNormalizationContext(
                document=document,
                incoming_blocks=list(blocks),
                source="agent_tool_or_mcp",
            ),
        )
        for block in normalized_blocks:
            runtime_block = self._build_runtime_block(block, document)
            document.add_block(runtime_block)

    @staticmethod
    def _drop_duplicate_document_title_headings(
        context: BlockNormalizationContext,
    ) -> list[BlockInput]:
        normalized: list[BlockInput] = []
        normalized_document_title = context.document.metadata.title.strip()
        for current in context.incoming_blocks:
            current_text = (
                current.text.strip() if isinstance(current, BlockHeadingInput) else ""
            )
            if (
                isinstance(current, BlockHeadingInput)
                and not normalized
                and normalized_document_title
                and current_text == normalized_document_title
            ):
                continue
            normalized.append(current)
        return normalized

    @staticmethod
    def _move_landscape_intro_paragraphs_before_section_break(
        context: BlockNormalizationContext,
    ) -> list[BlockInput]:
        normalized: list[BlockInput] = []
        index = 0
        blocks = context.incoming_blocks

        while index < len(blocks):
            current = blocks[index]
            heading_block = blocks[index + 1] if index + 1 < len(blocks) else None
            if not (
                isinstance(current, SectionBreakInput)
                and current.page_orientation == "landscape"
                and isinstance(heading_block, BlockHeadingInput)
            ):
                normalized.append(current)
                index += 1
                continue

            paragraph_start = index + 2
            paragraph_end = paragraph_start
            while paragraph_end < len(blocks):
                paragraph_block = blocks[paragraph_end]
                if not isinstance(paragraph_block, SectionParagraphInput):
                    break
                if paragraph_block.variant != "body":
                    break
                if not paragraph_block.text.strip() and not paragraph_block.runs:
                    break
                paragraph_end += 1

            table_block = blocks[paragraph_end] if paragraph_end < len(blocks) else None
            if paragraph_end == paragraph_start or not isinstance(
                table_block, SectionTableInput
            ):
                normalized.append(current)
                index += 1
                continue

            normalized.extend(blocks[paragraph_start:paragraph_end])
            normalized.append(current)
            normalized.append(heading_block)
            normalized.append(table_block)
            index = paragraph_end + 1

        return normalized

    @staticmethod
    def _promote_heading_before_table_to_caption(
        context: BlockNormalizationContext,
    ) -> list[BlockInput]:
        normalized: list[BlockInput] = []
        index = 0
        blocks = context.incoming_blocks
        while index < len(blocks):
            current = blocks[index]
            previous_block = blocks[index - 1] if index > 0 else None
            next_block = blocks[index + 1] if index + 1 < len(blocks) else None
            current_text = (
                current.text.strip() if isinstance(current, BlockHeadingInput) else ""
            )
            next_caption = (
                (next_block.caption or "").strip()
                if isinstance(next_block, SectionTableInput)
                else ""
            )
            next_title = (
                (next_block.title or "").strip()
                if isinstance(next_block, SectionTableInput)
                else ""
            )
            uses_executive_brief_table_titles = (
                context.document.metadata.theme_name == "executive_brief"
            )
            follows_landscape_section = (
                isinstance(previous_block, SectionBreakInput)
                and previous_block.page_orientation == "landscape"
            )

            if (
                isinstance(current, BlockHeadingInput)
                and isinstance(next_block, SectionTableInput)
                and not (next_caption or next_title)
                and not follows_landscape_section
                and len(current_text) <= MAX_HEADING_LENGTH_FOR_TABLE_TITLE
                and (
                    (
                        not current.bottom_border
                        and not _looks_like_numbered_heading(current_text)
                    )
                    or (
                        current.bottom_border
                        and _looks_like_numbered_heading(current_text)
                    )
                    or (
                        current.bottom_border and uses_executive_brief_table_titles
                    )
                )
            ):
                normalized.append(
                    next_block.model_copy(
                        update={
                            "caption": current_text,
                            "title": current_text,
                        }
                    )
                )
                index += 2
                continue

            normalized.append(current)
            index += 1

        return normalized

    def _build_runtime_block(
        self,
        block: BlockInput,
        document: DocumentModel,
    ) -> RuntimeBlock:
        if isinstance(block, SectionHeroBannerInput):
            return HeroBannerBlock(
                title=block.title,
                subtitle=block.subtitle,
                theme_color=block.theme_color,
                text_color=block.text_color,
                subtitle_color=block.subtitle_color,
                min_height_pt=block.min_height_pt,
                full_width=block.full_width,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, BlockHeadingInput):
            return HeadingBlock(
                text=block.text,
                level=block.level,
                bottom_border=block.bottom_border,
                bottom_border_color=block.bottom_border_color,
                bottom_border_size_pt=block.bottom_border_size_pt,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionParagraphInput):
            return ParagraphBlock(
                text=block.text,
                variant=block.variant,
                title=block.title,
                runs=block.runs,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionAccentBoxInput):
            return AccentBoxBlock(
                title=block.title,
                text=block.text,
                runs=[run.model_copy(deep=True) for run in block.runs],
                items=[
                    item if isinstance(item, str) else item.model_copy(deep=True)
                    for item in block.items
                ],
                accent_color=block.accent_color,
                fill_color=block.fill_color,
                title_color=block.title_color,
                border_color=block.border_color,
                border_width_pt=block.border_width_pt,
                accent_border_width_pt=block.accent_border_width_pt,
                padding_pt=block.padding_pt,
                title_font_scale=block.title_font_scale,
                body_font_scale=block.body_font_scale,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionListInput):
            return ListBlock(
                items=block.items,
                ordered=block.ordered,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionTableInput):
            return TableBlock(
                headers=block.headers,
                rows=block.rows,
                header_groups=block.header_groups,
                table_style=block.table_style or document.metadata.table_template,
                caption=block.caption or block.title,
                column_widths=block.column_widths,
                numeric_columns=block.numeric_columns,
                header_fill=block.header_fill,
                header_fill_enabled=block.header_fill_enabled,
                header_text_color=block.header_text_color,
                header_bold=block.header_bold,
                banded_rows=block.banded_rows,
                banded_row_fill=block.banded_row_fill,
                first_column_bold=block.first_column_bold,
                table_align=block.table_align,
                border_style=block.border_style,
                caption_emphasis=block.caption_emphasis,
                cell_padding_horizontal_pt=block.cell_padding_horizontal_pt,
                cell_padding_vertical_pt=block.cell_padding_vertical_pt,
                header_font_scale=block.header_font_scale,
                body_font_scale=block.body_font_scale,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionMetricCardsInput):
            return MetricCardsBlock(
                metrics=[
                    MetricCard(
                        label=metric.label,
                        value=metric.value,
                        delta=metric.delta,
                        note=metric.note,
                        value_color=metric.value_color,
                        delta_color=metric.delta_color,
                        fill_color=metric.fill_color,
                        label_color=metric.label_color,
                        note_color=metric.note_color,
                        value_font_scale=metric.value_font_scale,
                        delta_font_scale=metric.delta_font_scale,
                    )
                    for metric in block.metrics
                ],
                accent_color=block.accent_color,
                fill_color=block.fill_color,
                label_color=block.label_color,
                border_color=block.border_color,
                border_width_pt=block.border_width_pt,
                divider_color=block.divider_color,
                divider_width_pt=block.divider_width_pt,
                padding_pt=block.padding_pt,
                label_font_scale=block.label_font_scale,
                value_font_scale=block.value_font_scale,
                delta_font_scale=block.delta_font_scale,
                note_font_scale=block.note_font_scale,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionCardInput):
            expanded = self._expand_summary_card_block(block, document)
            if len(expanded) != 1:
                return GroupBlock(blocks=expanded)
            return expanded[0]
        if isinstance(block, SectionPageBreakInput):
            return PageBreakBlock()
        if isinstance(block, SectionBreakInput):
            return SectionBreakBlock(
                start_type=block.start_type,
                inherit_header_footer=block.inherit_header_footer,
                page_orientation=block.page_orientation,
                margins=block.margins,
                restart_page_numbering=block.restart_page_numbering,
                page_number_start=block.page_number_start,
                header_footer=block.header_footer,
            )
        if isinstance(block, TocInput):
            return TocBlock(
                title=block.title,
                levels=block.levels,
                start_on_new_page=block.start_on_new_page,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, BlockGroupInput):
            return GroupBlock(
                blocks=[
                    self._build_runtime_block(child, document) for child in block.blocks
                ],
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, BlockColumnsInput):
            return ColumnsBlock(
                columns=[
                    ColumnBlock(
                        blocks=[
                            self._build_runtime_block(child, document)
                            for child in column.blocks
                        ]
                    )
                    for column in block.columns
                ],
                style=block.style,
                layout=block.layout,
            )
        raise TypeError(f"Unsupported block input: {type(block)!r}")

    def _expand_summary_card_block(
        self, block: SectionCardInput, document: DocumentModel
    ) -> list:
        document_style = getattr(document.metadata, "document_style", None)
        summary_card_config = getattr(document_style, "summary_card_defaults", None)
        try:
            summary_card_defaults = self._summary_card_defaults_resolver(
                summary_card_config
            )
        except Exception:
            logger.exception(
                "Failed to resolve summary card defaults for document %s",
                document.document_id,
            )
            summary_card_defaults = {}
        try:
            from ...document_core.macros import build_summary_card_group
        except ImportError:
            return [
                SummaryCardBlock(
                    title=block.title,
                    items=block.items,
                    variant=block.variant,
                    style=block.style,
                    layout=block.layout,
                    **summary_card_defaults,
                )
            ]

        try:
            return [
                build_summary_card_group(
                    title=block.title,
                    items=block.items,
                    variant=block.variant,
                    style=block.style,
                    layout=block.layout,
                    **summary_card_defaults,
                )
            ]
        except Exception:
            logger.exception(
                "Failed to expand summary card block for document %s",
                document.document_id,
            )
            return [
                SummaryCardBlock(
                    title=block.title,
                    items=block.items,
                    variant=block.variant,
                    style=block.style,
                    layout=block.layout,
                )
            ]

    def add_heading(self, request: AddHeadingRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    BlockHeadingInput(
                        text=request.text,
                        level=request.level,
                        style=request.style.model_copy(deep=True),
                        layout=request.layout.model_copy(deep=True),
                    )
                ],
            )
        )

    def add_paragraph(self, request: AddParagraphRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    SectionParagraphInput(
                        text=request.text,
                        variant=request.variant,
                        title=request.title,
                        runs=[run.model_copy(deep=True) for run in request.runs],
                        style=request.style.model_copy(deep=True),
                        layout=request.layout.model_copy(deep=True),
                    )
                ],
            )
        )

    def add_list(self, request: AddListRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    SectionListInput(
                        items=list(request.items),
                        ordered=request.ordered,
                        style=request.style.model_copy(deep=True),
                        layout=request.layout.model_copy(deep=True),
                    )
                ],
            )
        )

    def add_table(self, request: AddTableRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    SectionTableInput(
                        headers=[*request.headers],
                        rows=[[*row] for row in request.rows],
                        header_groups=[
                            group.model_copy(deep=True)
                            for group in request.header_groups
                        ],
                        table_style=request.table_style,
                        caption=request.caption or request.title,
                        column_widths=[*request.column_widths],
                        numeric_columns=[*request.numeric_columns],
                        header_fill=request.header_fill,
                        header_fill_enabled=request.header_fill_enabled,
                        header_text_color=request.header_text_color,
                        header_bold=request.header_bold,
                        banded_rows=request.banded_rows,
                        banded_row_fill=request.banded_row_fill,
                        first_column_bold=request.first_column_bold,
                        table_align=request.table_align,
                        border_style=request.border_style,
                        caption_emphasis=request.caption_emphasis,
                        cell_padding_horizontal_pt=request.cell_padding_horizontal_pt,
                        cell_padding_vertical_pt=request.cell_padding_vertical_pt,
                        header_font_scale=request.header_font_scale,
                        body_font_scale=request.body_font_scale,
                        style=request.style.model_copy(deep=True),
                        layout=request.layout.model_copy(deep=True),
                    )
                ],
            )
        )

    def add_summary_card(self, request: AddSummaryCardRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    {
                        "type": BLOCK_TYPE_SUMMARY_CARD,
                        "title": request.title,
                        "items": request.items,
                        "variant": request.variant,
                        "style": request.style.model_dump(exclude_none=True),
                        "layout": request.layout.model_dump(exclude_none=True),
                    }
                ],
            )
        )

    def add_page_break(self, request: AddPageBreakRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[{"type": BLOCK_TYPE_PAGE_BREAK}],
            )
        )

    def add_section_bundle(self, request: AddSectionBundleRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    {
                        "type": BLOCK_TYPE_HEADING,
                        "text": request.heading,
                        "level": request.level,
                    },
                    *[block.model_dump(exclude_none=True) for block in request.blocks],
                ],
            )
        )

    def finalize_document(self, request: FinalizeDocumentRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.status = DocumentStatus.FINALIZED
            document.touch()
            return document

    def build_prompt_summary(self, document_id: str) -> dict[str, object]:
        with self._lock:
            document = self.require_document(document_id)
            return self._build_prompt_summary_locked(document)

    @staticmethod
    def _build_prompt_summary_locked(document: DocumentModel) -> dict[str, object]:
        latest_block_types = [
            DocumentSessionStore._summarize_runtime_block_type(block)
            for block in document.blocks[-3:]
        ]
        next_allowed_actions: list[str]
        if document.status == DocumentStatus.DRAFT:
            next_allowed_actions = ["add_blocks", "finalize_document"]
        elif document.status == DocumentStatus.FINALIZED:
            next_allowed_actions = ["export_document"]
        else:
            next_allowed_actions = []
        return {
            "document_id": document.document_id,
            "title": document.metadata.title,
            "status": document.status.value,
            "block_count": len(document.blocks),
            "latest_block_types": latest_block_types,
            "next_allowed_actions": next_allowed_actions,
        }

    @staticmethod
    def _summarize_runtime_block_type(block: object) -> str:
        if isinstance(block, HeadingBlock):
            return BLOCK_TYPE_HEADING
        if isinstance(block, ParagraphBlock):
            return BLOCK_TYPE_PARAGRAPH
        if isinstance(block, ListBlock):
            return BLOCK_TYPE_LIST
        if isinstance(block, TableBlock):
            return BLOCK_TYPE_TABLE
        if isinstance(block, SummaryCardBlock):
            return BLOCK_TYPE_SUMMARY_CARD
        if isinstance(block, PageBreakBlock):
            return BLOCK_TYPE_PAGE_BREAK
        if isinstance(block, SectionBreakBlock):
            return "section_break"
        if isinstance(block, TocBlock):
            return "toc"
        if isinstance(block, GroupBlock):
            return "group"
        if isinstance(block, ColumnsBlock):
            return "columns"
        return getattr(block, "type", "unknown")

    def prepare_export_path(
        self, request: ExportDocumentRequest
    ) -> tuple[DocumentModel, Path]:
        with self._lock:
            document = self.require_document(request.document_id)
            preferred_name = request.output_name or document.metadata.preferred_filename
            file_name = _normalize_docx_filename(preferred_name)

            workspace_dir = self.workspace_dir.resolve()
            output_dir = (
                (workspace_dir / request.output_dir).resolve()
                if request.output_dir
                else workspace_dir
            )
            if not _is_within_workspace(output_dir, workspace_dir):
                raise ValueError("output_dir cannot escape the document workspace")
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = (output_dir / file_name).resolve()
            if not _is_within_workspace(output_path, workspace_dir):
                raise ValueError("output_path cannot escape the document workspace")
            document.output_path = str(output_path)
            document.touch()
            return document, output_path

    def complete_export(self, document_id: str) -> DocumentModel:
        with self._lock:
            document = self.require_document(document_id)
            document.status = DocumentStatus.EXPORTED
            document.touch()
            return document


__all__ = [
    "BLOCK_TYPE_HEADING",
    "BLOCK_TYPE_LIST",
    "BLOCK_TYPE_PAGE_BREAK",
    "BLOCK_TYPE_PARAGRAPH",
    "BLOCK_TYPE_SUMMARY_CARD",
    "BLOCK_TYPE_TABLE",
    "DocumentSessionStore",
    "MAX_HEADING_LENGTH_FOR_TABLE_TITLE",
    "_default_workspace_dir",
    "_is_within_workspace",
]
