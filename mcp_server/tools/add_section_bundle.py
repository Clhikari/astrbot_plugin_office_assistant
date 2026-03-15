from mcp.server.fastmcp import FastMCP

from ..schemas import AddSectionBundleRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_add_section_bundle_tool(
    server: FastMCP, store: DocumentSessionStore
) -> None:
    @server.tool(
        name="add_section_bundle",
        description=(
            "Append one complete report section in a single tool call. "
            "Use this to reduce tool-call count for complex Word reports."
        ),
        structured_output=True,
    )
    def add_section_bundle(
        document_id: str,
        heading: str,
        blocks: list[dict],
        level: int = 1,
    ) -> ToolResult:
        request = AddSectionBundleRequest(
            document_id=document_id,
            heading=heading,
            level=level,
            blocks=blocks,
        )
        document = store.add_section_bundle(request)
        return ToolResult(
            success=True,
            message="Section bundle added.",
            document=build_document_summary(document),
        )
