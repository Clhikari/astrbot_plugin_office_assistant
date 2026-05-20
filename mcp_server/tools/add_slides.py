from mcp.server.fastmcp import FastMCP

from ...domain.document.session_store import DocumentSessionStore
from ...domain.document.contracts import (
    AddBlocksRequest,
    ToolResult,
    build_document_summary,
    normalize_slide_bullets,
)


def register_add_slides_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="add_slides",
        description=(
            "Append slides to a PPT document. Only use for documents created with format='ppt'. "
            "Supported slide types: title_slide, content_slide, table_slide, image_slide."
        ),
        structured_output=True,
    )
    def add_slides(document_id: str, slides: list[dict]) -> ToolResult:
        if not document_id:
            return ToolResult(
                success=False,
                message="document_id is required.",
            )
        doc = store.get_document(document_id)
        if doc is None:
            return ToolResult(
                success=False,
                message=f"document_id={document_id} not found.",
            )
        if doc.format != "ppt":
            return ToolResult(
                success=False,
                message="add_slides 仅用于 PPT 文档。Word 文档请使用 add_blocks。",
            )
        normalized = normalize_slide_bullets(slides)
        try:
            request = AddBlocksRequest(
                document_id=document_id,
                blocks=normalized,
            )
            document = store.add_blocks(request)
        except Exception as exc:
            return ToolResult(
                success=False,
                message=str(exc),
            )
        return ToolResult(
            success=True,
            message="Slides added.",
            document=build_document_summary(document),
        )
