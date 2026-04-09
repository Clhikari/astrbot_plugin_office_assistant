from __future__ import annotations

from pathlib import Path
from collections.abc import Sequence

from astrbot.api import logger

from ...document_core.models.document import DocumentModel
from .contracts import ExportDocumentRequest
from .hooks import (
    AfterExportContext,
    AfterExportHook,
    BeforeExportHook,
    ExportPreparationContext,
    run_after_export_hooks,
    run_before_export_hooks,
)
from .render_backends import DocumentRenderBackend
from .render_backends import render_document_with_backends
from .session_store import DocumentSessionStore


async def export_document_via_pipeline(
    *,
    store: DocumentSessionStore,
    render_backends: Sequence[DocumentRenderBackend],
    request: ExportDocumentRequest,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    source: str,
) -> tuple[DocumentModel, Path]:
    document, output_path = store.prepare_export_path(request)
    logger.debug(
        "[office-assistant] export pipeline prepared document=%s source=%s output=%s",
        document.document_id,
        source,
        output_path,
    )
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
        logger.debug(
            "[office-assistant] before_export hooks completed for document=%s output=%s",
            export_context.document.document_id,
            export_context.output_path,
        )
    result = render_document_with_backends(
        export_context.document,
        export_context.output_path,
        render_backends,
    )
    logger.debug(
        "[office-assistant] document build completed for document=%s output=%s backend=%s",
        export_context.document.document_id,
        export_context.output_path,
        result.backend_name,
    )
    document = store.complete_export(request.document_id)
    after_context = AfterExportContext(
        document=document,
        output_path=export_context.output_path,
        source=source,
    )
    if after_export_hooks:
        after_context = await run_after_export_hooks(
            after_export_hooks,
            after_context,
        )
        logger.debug(
            "[office-assistant] after_export hooks completed for document=%s output=%s",
            after_context.document.document_id,
            after_context.output_path,
        )
    document.output_path = str(after_context.output_path)
    document.touch()
    logger.debug(
        "[office-assistant] export pipeline finished document=%s output=%s",
        document.document_id,
        after_context.output_path,
    )
    return document, after_context.output_path


__all__ = ["export_document_via_pipeline"]
