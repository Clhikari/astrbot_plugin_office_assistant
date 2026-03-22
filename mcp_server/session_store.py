from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path
from threading import RLock
from uuid import uuid4

from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path

from ..document_core.models.blocks import (
    ColumnBlock,
    ColumnsBlock,
    GroupBlock,
    HeadingBlock,
    ListBlock,
    PageBreakBlock,
    ParagraphBlock,
    SummaryCardBlock,
    TableBlock,
)
from ..document_core.models.document import (
    DocumentMetadata,
    DocumentModel,
    DocumentStatus,
)
from .schemas import (
    AddBlocksRequest,
    AddHeadingRequest,
    AddListRequest,
    AddPageBreakRequest,
    AddParagraphRequest,
    AddSectionBundleRequest,
    AddSummaryCardRequest,
    AddTableRequest,
    BlockColumnsInput,
    BlockGroupInput,
    BlockHeadingInput,
    CreateDocumentRequest,
    ExportDocumentRequest,
    FinalizeDocumentRequest,
    SectionCardInput,
    SectionListInput,
    SectionPageBreakInput,
    SectionParagraphInput,
    SectionTableInput,
    _normalize_docx_filename,
)

PLUGIN_NAME = "astrbot_plugin_office_assistant"

# Block type identifiers — single source of truth for the serialization protocol.
BLOCK_TYPE_HEADING = "heading"
BLOCK_TYPE_PARAGRAPH = "paragraph"
BLOCK_TYPE_LIST = "list"
BLOCK_TYPE_TABLE = "table"
BLOCK_TYPE_SUMMARY_CARD = "summary_card"
BLOCK_TYPE_PAGE_BREAK = "page_break"
MAX_HEADING_LENGTH_FOR_TABLE_TITLE = 24


def _default_workspace_dir() -> Path:
    return Path(get_astrbot_plugin_data_path()) / PLUGIN_NAME / "documents"


def _is_within_workspace(path: Path, workspace_dir: Path) -> bool:
    try:
        path.relative_to(workspace_dir)
        return True
    except ValueError:
        return False


class DocumentSessionStore:
    def __init__(
        self,
        workspace_dir: Path | None = None,
        *,
        max_documents: int | None = 256,
        ttl: timedelta | None = None,
    ) -> None:
        self._lock = RLock()
        self._documents: dict[str, DocumentModel] = {}
        self.workspace_dir = workspace_dir or _default_workspace_dir()
        self.workspace_dir.mkdir(parents=True, exist_ok=True)
        self._max_documents = max_documents
        self._ttl = ttl

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

    def _append_blocks_locked(self, document: DocumentModel, blocks: list) -> None:
        normalized_blocks = self._normalize_table_title_blocks(
            blocks, document.metadata.title
        )
        for block in normalized_blocks:
            runtime_block = self._build_runtime_block(block, document)
            document.add_block(runtime_block)

    @staticmethod
    def _normalize_table_title_blocks(blocks: list, document_title: str = ""):
        normalized: list = []
        index = 0
        normalized_document_title = document_title.strip()
        while index < len(blocks):
            current = blocks[index]
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

            if (
                isinstance(current, BlockHeadingInput)
                and isinstance(next_block, SectionTableInput)
                and not (next_caption or next_title)
                and len(current_text) <= MAX_HEADING_LENGTH_FOR_TABLE_TITLE
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

            if (
                isinstance(current, BlockHeadingInput)
                and not normalized
                and normalized_document_title
                and current_text == normalized_document_title
            ):
                index += 1
                continue

            normalized.append(current)
            index += 1

        return normalized

    def _build_runtime_block(self, block, document: DocumentModel):
        if isinstance(block, BlockHeadingInput):
            return HeadingBlock(
                text=block.text,
                level=block.level,
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
                table_style=block.table_style or document.metadata.table_template,
                caption=block.caption or block.title,
                column_widths=block.column_widths,
                numeric_columns=block.numeric_columns,
                style=block.style,
                layout=block.layout,
            )
        if isinstance(block, SectionCardInput):
            expanded = self._expand_summary_card_block(block)
            if len(expanded) != 1:
                return GroupBlock(blocks=expanded)
            return expanded[0]
        if isinstance(block, SectionPageBreakInput):
            return PageBreakBlock()
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

    def _expand_summary_card_block(self, block: SectionCardInput) -> list:
        try:
            from ..document_core.macros import build_summary_card_group

            return [
                build_summary_card_group(
                    title=block.title,
                    items=block.items,
                    variant=block.variant,
                    style=block.style,
                    layout=block.layout,
                )
            ]
        except Exception:
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
                    {
                        "type": BLOCK_TYPE_HEADING,
                        "text": request.text,
                        "level": request.level,
                        "style": request.style.model_dump(exclude_none=True),
                        "layout": request.layout.model_dump(exclude_none=True),
                    }
                ],
            )
        )

    def add_paragraph(self, request: AddParagraphRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    {
                        "type": BLOCK_TYPE_PARAGRAPH,
                        "text": request.text,
                        "variant": request.variant,
                        "title": request.title,
                        "runs": [
                            run.model_dump(exclude_none=True) for run in request.runs
                        ],
                        "style": request.style.model_dump(exclude_none=True),
                        "layout": request.layout.model_dump(exclude_none=True),
                    }
                ],
            )
        )

    def add_list(self, request: AddListRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    {
                        "type": BLOCK_TYPE_LIST,
                        "items": request.items,
                        "ordered": request.ordered,
                        "style": request.style.model_dump(exclude_none=True),
                        "layout": request.layout.model_dump(exclude_none=True),
                    }
                ],
            )
        )

    def add_table(self, request: AddTableRequest) -> DocumentModel:
        return self.add_blocks(
            AddBlocksRequest(
                document_id=request.document_id,
                blocks=[
                    {
                        "type": BLOCK_TYPE_TABLE,
                        "headers": request.headers,
                        "rows": request.rows,
                        "table_style": request.table_style,
                        "caption": request.caption or request.title,
                        "column_widths": request.column_widths,
                        "numeric_columns": request.numeric_columns,
                        "style": request.style.model_dump(exclude_none=True),
                        "layout": request.layout.model_dump(exclude_none=True),
                    }
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
        """Mark a document as exported after the builder has written the file."""
        with self._lock:
            document = self.require_document(document_id)
            document.status = DocumentStatus.EXPORTED
            document.touch()
            return document
