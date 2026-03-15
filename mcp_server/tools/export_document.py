from mcp.server.fastmcp import FastMCP

from ...document_core.builders.word_builder import WordDocumentBuilder
from ...document_core.models.document import DocumentStatus
from ..schemas import (
    ExportDocumentRequest,
    ExportDocumentResult,
    build_document_summary,
)
from ..session_store import DocumentSessionStore


def register_export_document_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    builder = WordDocumentBuilder()

    @server.tool(
        name="export_document",
        description="Export the current Word draft to a .docx file and return the file path.",
        structured_output=True,
    )
    def export_document(
        document_id: str,
        output_dir: str = "",
        output_name: str = "",
    ) -> ExportDocumentResult:
        request = ExportDocumentRequest(
            document_id=document_id,
            output_dir=output_dir,
            output_name=output_name,
        )
        document, output_path = store.prepare_export_path(request)
        builder.build(document, output_path)
        document.status = DocumentStatus.EXPORTED
        document.touch()
        return ExportDocumentResult(
            success=True,
            message="Document exported.",
            document=build_document_summary(document),
            file_path=str(output_path),
        )
