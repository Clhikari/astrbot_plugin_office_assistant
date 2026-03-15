from mcp.server.fastmcp import FastMCP

from ..schemas import AddTableRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_table_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_table",
        description="Append a table block to the current draft document.",
        structured_output=True,
    )
    def add_table(
        document_id: str,
        headers: list[str] | None = None,
        rows: list[list[str]] | None = None,
    ) -> ToolResult:
        request = AddTableRequest(
            document_id=document_id,
            headers=headers or [],
            rows=rows or [],
        )
        document = store.add_table(request)
        return ToolResult(
            success=True,
            message="Table added.",
            document=build_document_summary(document),
        )
