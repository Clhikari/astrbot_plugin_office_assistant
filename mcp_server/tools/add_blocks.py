from mcp.server.fastmcp import FastMCP

from ..schemas import AddBlocksRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_blocks_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_blocks",
        description=(
            "Append one or more blocks in order. Use this for mixed content such as "
            "heading, paragraph, list, table, summary_card, page_break, group, or columns."
        ),
        structured_output=True,
    )
    def add_blocks(document_id: str, blocks: list[dict]) -> ToolResult:
        request = AddBlocksRequest(
            document_id=document_id,
            blocks=blocks,
        )
        document = store.add_blocks(request)
        return ToolResult(
            success=True,
            message="Blocks added.",
            document=build_document_summary(document),
        )
