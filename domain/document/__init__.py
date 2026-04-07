from __future__ import annotations

__all__ = [
    "AddBlocksRequest",
    "AfterExportHook",
    "BeforeExportHook",
    "BlockNormalizationContext",
    "BlockNormalizeHook",
    "CreateDocumentRequest",
    "DocumentFormat",
    "DocumentRenderBackend",
    "DocumentRenderBackendConfig",
    "DocumentRenderBackendError",
    "DocumentSessionStore",
    "DocumentSummary",
    "ExportDocumentRequest",
    "ExportDocumentResult",
    "ExportPreparationContext",
    "NodeDocumentRenderBackend",
    "PythonExcelRenderBackend",
    "PythonPptRenderBackend",
    "FinalizeDocumentRequest",
    "RenderResult",
    "ToolResult",
    "build_document_summary",
    "build_document_render_backends",
    "build_header_footer_schema",
    "export_document_via_pipeline",
    "normalize_create_document_kwargs",
    "normalize_raw_block_payloads",
    "run_after_export_hooks",
    "run_before_export_hooks",
    "run_block_normalize_hooks",
]


def __getattr__(name: str):
    if name == "DocumentSessionStore":
        from .session_store import DocumentSessionStore

        return DocumentSessionStore
    if name == "export_document_via_pipeline":
        from .export_pipeline import export_document_via_pipeline

        return export_document_via_pipeline
    if name in {
        "DocumentFormat",
        "DocumentRenderBackend",
        "DocumentRenderBackendConfig",
        "DocumentRenderBackendError",
        "NodeDocumentRenderBackend",
        "PythonExcelRenderBackend",
        "PythonPptRenderBackend",
        "RenderResult",
        "build_document_render_backends",
    }:
        from . import render_backends as render_backends_module

        return getattr(render_backends_module, name)
    if name in {
        "AddBlocksRequest",
        "CreateDocumentRequest",
        "DocumentSummary",
        "ExportDocumentRequest",
        "ExportDocumentResult",
        "FinalizeDocumentRequest",
        "ToolResult",
        "build_document_summary",
        "build_header_footer_schema",
        "normalize_create_document_kwargs",
        "normalize_raw_block_payloads",
    }:
        from .contracts import __dict__ as contracts_namespace

        return contracts_namespace[name]
    if name in {
        "AfterExportHook",
        "BeforeExportHook",
        "BlockNormalizationContext",
        "BlockNormalizeHook",
        "ExportPreparationContext",
        "run_after_export_hooks",
        "run_before_export_hooks",
        "run_block_normalize_hooks",
    }:
        from .hooks import __dict__ as hooks_namespace

        return hooks_namespace[name]
    raise AttributeError(name)
