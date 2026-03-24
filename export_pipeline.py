from __future__ import annotations

from collections.abc import Awaitable, Callable
from pathlib import Path

from .document_core.builders.word_builder import WordDocumentBuilder
from .internal_hooks import (
    AfterExportContext,
    AfterExportHook,
    BeforeExportHook,
    ExportPreparationContext,
    run_after_export_hooks,
    run_before_export_hooks,
)
from .mcp_server.schemas import ExportDocumentRequest
from .mcp_server.session_store import DocumentSessionStore


async def export_document_via_pipeline(
    *,
    store: DocumentSessionStore,
    builder: WordDocumentBuilder,
    request: ExportDocumentRequest,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    source: str,
) -> tuple[object, Path]:
    document, output_path = store.prepare_export_path(request)
    export_context = ExportPreparationContext(
        document=document,
        output_path=output_path,
        source=source,
    )
    if before_export_hooks:
        export_context = await run_before_export_hooks(
            before_export_hooks,
            export_context,
        )
    builder.build(export_context.document, export_context.output_path)
    document = store.complete_export(request.document_id)
    after_context = AfterExportContext(
        document=document,
        output_path=Path(export_context.output_path),
        source=source,
    )
    if after_export_hooks:
        after_context = await run_after_export_hooks(
            after_export_hooks,
            after_context,
        )
    document.output_path = str(Path(after_context.output_path))
    document.touch()
    return document, Path(after_context.output_path)
