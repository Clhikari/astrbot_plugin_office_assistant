from mcp.server.fastmcp import FastMCP

from ..schemas import FinalizeDocumentRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_finalize_document_tool(
    server: FastMCP, store: DocumentSessionStore
) -> None:
    @server.tool(
        name="finalize_document",
        description="Mark a document draft as finalized before export.",
        structured_output=True,
    )
    def finalize_document(document_id: str) -> ToolResult:
        request = FinalizeDocumentRequest(document_id=document_id)
        document = store.finalize_document(request)
        return ToolResult(
            success=True,
            message="Document finalized.",
            document=build_document_summary(document),
        )
