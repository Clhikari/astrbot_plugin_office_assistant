from mcp.server.fastmcp import FastMCP

from ..schemas import AddParagraphRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_paragraph_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_paragraph",
        description="Append a paragraph block to the current draft document.",
        structured_output=True,
    )
    def add_paragraph(document_id: str, text: str) -> ToolResult:
        request = AddParagraphRequest(document_id=document_id, text=text)
        document = store.add_paragraph(request)
        return ToolResult(
            success=True,
            message="Paragraph added.",
            document=build_document_summary(document),
        )
