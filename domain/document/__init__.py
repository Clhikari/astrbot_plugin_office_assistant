__all__ = [
    "AddBlocksRequest",
    "AfterExportHook",
    "BeforeExportHook",
    "BlockNormalizationContext",
    "BlockNormalizeHook",
    "CreateDocumentRequest",
    "DocumentSessionStore",
    "DocumentSummary",
    "ExportDocumentRequest",
    "ExportDocumentResult",
    "ExportPreparationContext",
    "FinalizeDocumentRequest",
    "ToolResult",
    "build_document_summary",
    "build_header_footer_schema",
    "export_document_via_pipeline",
    "normalize_raw_block_payloads",
    "run_after_export_hooks",
    "run_before_export_hooks",
    "run_block_normalize_hooks",
]


# Keep __all__ and __getattr__ in sync when exports change.
# This module uses lazy imports both for startup cost and for compatibility
# exports, so missing an entry in either place can surface as AttributeError.
def __getattr__(name: str):
    if name == "DocumentSessionStore":
        from .session_store import DocumentSessionStore

        return DocumentSessionStore
    if name == "export_document_via_pipeline":
        from .export_pipeline import export_document_via_pipeline

        return export_document_via_pipeline
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
