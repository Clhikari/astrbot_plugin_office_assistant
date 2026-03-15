from __future__ import annotations

from pathlib import Path
from threading import RLock
from uuid import uuid4

from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path

from ..document_core.models.blocks import (
    HeadingBlock,
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
    AddHeadingRequest,
    AddParagraphRequest,
    AddSectionBundleRequest,
    AddSummaryCardRequest,
    AddTableRequest,
    CreateDocumentRequest,
    ExportDocumentRequest,
    FinalizeDocumentRequest,
    SectionCardInput,
    SectionParagraphInput,
    SectionTableInput,
    _normalize_docx_filename,
)

PLUGIN_NAME = "astrbot_plugin_office_assistant"


def _default_workspace_dir() -> Path:
    return Path(get_astrbot_plugin_data_path()) / PLUGIN_NAME / "documents"


def _is_within_workspace(path: Path, workspace_dir: Path) -> bool:
    try:
        path.relative_to(workspace_dir)
        return True
    except ValueError:
        return False


class DocumentSessionStore:
    def __init__(self, workspace_dir: Path | None = None) -> None:
        self._lock = RLock()
        self._documents: dict[str, DocumentModel] = {}
        self.workspace_dir = workspace_dir or _default_workspace_dir()
        self.workspace_dir.mkdir(parents=True, exist_ok=True)

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
            return document

    def get_document(self, document_id: str) -> DocumentModel | None:
        with self._lock:
            return self._documents.get(document_id)

    def require_document(self, document_id: str) -> DocumentModel:
        document = self.get_document(document_id)
        if document is None:
            raise KeyError(f"Document not found: {document_id}")
        return document

    def add_heading(self, request: AddHeadingRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.add_block(HeadingBlock(text=request.text, level=request.level))
            return document

    def add_paragraph(self, request: AddParagraphRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.add_block(ParagraphBlock(text=request.text))
            return document

    def add_table(self, request: AddTableRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.add_block(
                TableBlock(
                    headers=request.headers,
                    rows=request.rows,
                    table_style=request.table_style or document.metadata.table_template,
                )
            )
            return document

    def add_summary_card(self, request: AddSummaryCardRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.add_block(
                SummaryCardBlock(
                    title=request.title,
                    items=request.items,
                    variant=request.variant,
                )
            )
            return document

    def add_section_bundle(self, request: AddSectionBundleRequest) -> DocumentModel:
        with self._lock:
            document = self.require_document(request.document_id)
            document.add_block(HeadingBlock(text=request.heading, level=request.level))
            for block in request.blocks:
                if isinstance(block, SectionParagraphInput):
                    document.add_block(ParagraphBlock(text=block.text))
                elif isinstance(block, SectionTableInput):
                    document.add_block(
                        TableBlock(
                            headers=block.headers,
                            rows=block.rows,
                            table_style=block.table_style
                            or document.metadata.table_template,
                        )
                    )
                elif isinstance(block, SectionCardInput):
                    document.add_block(
                        SummaryCardBlock(
                            title=block.title,
                            items=block.items,
                            variant=block.variant,
                        )
                    )
            return document

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
