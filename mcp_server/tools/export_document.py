from mcp.server.fastmcp import FastMCP

from ...domain.document.export_pipeline import export_document_via_pipeline
from ...domain.document.hooks import AfterExportHook, BeforeExportHook
from ...domain.document.render_backends import (
    DocumentRenderBackend,
    DocumentRenderBackendConfig,
    build_document_render_backends,
)
from ...domain.document.session_store import DocumentSessionStore
from ...domain.document.contracts import (
    ExportDocumentRequest,
    ExportDocumentResult,
    build_document_summary,
)


def register_export_document_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    render_backends: list[DocumentRenderBackend] | None = None,
    render_backend_config: DocumentRenderBackendConfig | None = None,
) -> None:
    resolved_render_backends = render_backends or build_document_render_backends("word")

    @server.tool(
        name="export_document",
        description=(
            "Export the current Word draft to a .docx file and return the file path."
        ),
        structured_output=True,
    )
    async def export_document(
        document_id: str,
        output_dir: str = "",
        output_name: str = "",
    ) -> ExportDocumentResult:
        request = ExportDocumentRequest(
            document_id=document_id,
            output_dir=output_dir,
            output_name=output_name,
        )
        document_for_routing = store.require_document(document_id)
        current_render_backends = (
            build_document_render_backends(
                document_for_routing.format,
                render_backend_config,
            )
            if render_backend_config is not None or document_for_routing.format != "word"
            else resolved_render_backends
        )
        document, output_path = await export_document_via_pipeline(
            store=store,
            render_backends=current_render_backends,
            request=request,
            before_export_hooks=before_export_hooks or [],
            after_export_hooks=after_export_hooks or [],
            source="mcp",
        )
        return ExportDocumentResult(
            success=True,
            message="Document exported.",
            document=build_document_summary(document),
            file_path=str(output_path),
        )
