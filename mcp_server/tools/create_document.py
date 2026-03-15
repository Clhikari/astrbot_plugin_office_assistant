from mcp.server.fastmcp import FastMCP

from ..schemas import CreateDocumentRequest, ToolResult, build_document_summary
from ..session_store import DocumentSessionStore


def register_create_document_tool(server: FastMCP, store: DocumentSessionStore) -> None:
    @server.tool(
        name="create_document",
        description="Create a draft Word document session and return its document_id.",
        structured_output=True,
    )
    def create_document(
        session_id: str = "",
        title: str = "",
        output_name: str = "document.docx",
        theme_name: str = "business_report",
        table_template: str = "report_grid",
        density: str = "comfortable",
        accent_color: str = "",
    ) -> ToolResult:
        request = CreateDocumentRequest(
            session_id=session_id,
            title=title,
            output_name=output_name,
            theme_name=theme_name,
            table_template=table_template,
            density=density,
            accent_color=accent_color,
        )
        document = store.create_document(request)
        return ToolResult(
            success=True,
            message="Document session created.",
            document=build_document_summary(document),
        )
