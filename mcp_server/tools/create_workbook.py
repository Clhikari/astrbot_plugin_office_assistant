from mcp.server.fastmcp import FastMCP

from ...domain.workbook.session_store import WorkbookSessionStore
from ...domain.workbook.contracts import (
    CreateWorkbookRequest,
    ToolResult,
    build_workbook_summary,
)


def register_create_workbook_tool(server: FastMCP, store: WorkbookSessionStore) -> None:
    @server.tool(
        name="create_workbook",
        description=(
            "Create a draft workbook session and return workbook_id for structured Excel generation."
        ),
        structured_output=True,
    )
    def create_workbook(
        session_id: str = "",
        title: str = "",
        filename: str = "workbook.xlsx",
    ) -> ToolResult:
        request = CreateWorkbookRequest(
            session_id=session_id,
            title=title,
            filename=filename,
        )
        workbook = store.create_workbook(request)
        return ToolResult(
            success=True,
            message="Workbook session created.",
            workbook=build_workbook_summary(workbook),
        )
