from mcp.server.fastmcp import FastMCP

from ...domain.document.session_store import DocumentSessionStore
from ...domain.document.contracts import (
    ToolResult,
    execute_add_slides,
)


def register_add_slides_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_slides",
        description=(
            "Append slides to a PPT document. Only use for documents created with format='ppt'. "
            "Supported slide types: title_slide, content_slide, table_slide, image_slide. "
            "For content_slide, bullets should be a string array. Objects with a 'text' key and numbers are also accepted."
        ),
        structured_output=True,
    )
    def add_slides(document_id: str, slides: list[dict]) -> ToolResult:
        return execute_add_slides(store, document_id, slides)
