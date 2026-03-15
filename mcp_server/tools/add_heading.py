from mcp.server.fastmcp import FastMCP

from ..schemas import AddHeadingRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_heading_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_heading",
        description="Append a heading block to the current draft document.",
        structured_output=True,
    )
    def add_heading(document_id: str, text: str, level: int = 1) -> ToolResult:
        request = AddHeadingRequest(document_id=document_id, text=text, level=level)
        document = store.add_heading(request)
        return ToolResult(
            success=True,
            message="Heading added.",
            document=build_document_summary(document),
        )
