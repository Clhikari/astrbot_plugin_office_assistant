from mcp.server.fastmcp import FastMCP

from ...document_core.builders.word_builder import WordDocumentBuilder
from ...export_pipeline import export_document_via_pipeline
from ...internal_hooks import AfterExportHook, BeforeExportHook
from ..schemas import (
    ExportDocumentRequest,
    ExportDocumentResult,
    build_document_summary,
)
from ..session_store import DocumentSessionStore


def register_export_document_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> None:
    builder = WordDocumentBuilder()

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
        document, output_path = await export_document_via_pipeline(
            store=store,
            builder=builder,
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
