from mcp.server.fastmcp import FastMCP

from ...domain.document.session_store import DocumentSessionStore
from ...domain.document.contracts import (
    CreateDocumentRequest,
    ToolResult,
    build_document_summary,
)


def register_create_document_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="create_document",
        description=(
            "Create a draft document session and return its document_id. "
            "Use format='word' for Word documents, format='ppt' for PowerPoint presentations."
        ),
        structured_output=True,
    )
    def create_document(
        session_id: str = "",
        format: str = "word",
        title: str = "",
        output_name: str = "",
        theme_name: str = "business_report",
        table_template: str = "report_grid",
        density: str = "comfortable",
        accent_color: str = "",
        document_style: dict | None = None,
    ) -> ToolResult:
        request = CreateDocumentRequest(
            session_id=session_id,
            format=format,
            title=title,
            output_name=output_name,
            theme_name=theme_name,
            table_template=table_template,
            density=density,
            accent_color=accent_color,
            document_style=document_style or {},
        )
        document = store.create_document(request)
        return ToolResult(
            success=True,
            message="Document session created.",
            document=build_document_summary(document),
        )
