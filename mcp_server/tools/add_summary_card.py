from mcp.server.fastmcp import FastMCP

from ..schemas import AddSummaryCardRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_summary_card_tool(
    server: FastMCP, store: DocumentSessionStore
) -> None:
    @server.tool(
        name="add_summary_card",
        description="Append a summary or conclusion card block to the current draft document.",
        structured_output=True,
    )
    def add_summary_card(
        document_id: str,
        title: str,
        items: list[str],
        variant: str = "summary",
    ) -> ToolResult:
        request = AddSummaryCardRequest(
            document_id=document_id,
            title=title,
            items=items,
            variant=variant,
        )
        document = store.add_summary_card(request)
        return ToolResult(
            success=True,
            message="Summary card added.",
            document=build_document_summary(document),
        )
