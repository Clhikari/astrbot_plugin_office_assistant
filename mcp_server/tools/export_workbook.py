from mcp.server.fastmcp import FastMCP

from ...domain.workbook.exporter import export_workbook_to_xlsx
from ...domain.workbook.session_store import WorkbookSessionStore
from ...domain.workbook.contracts import (
    ExportWorkbookRequest,
    ExportWorkbookResult,
    build_workbook_summary,
)


def register_export_workbook_tool(server: FastMCP, store: WorkbookSessionStore) -> None:
    @server.tool(
        name="export_workbook",
        description="Export workbook draft to .xlsx and return file path.",
        structured_output=True,
    )
    async def export_workbook(
        workbook_id: str,
        output_name: str = "",
    ) -> ExportWorkbookResult:
        request = ExportWorkbookRequest(
            workbook_id=workbook_id,
            output_name=output_name,
        )
        workbook, output_path = store.prepare_export_path(request)
        export_workbook_to_xlsx(workbook, output_path)
        workbook = store.complete_export(workbook.workbook_id)
        return ExportWorkbookResult(
            success=True,
            message="Workbook exported.",
            workbook=build_workbook_summary(workbook),
            file_path=str(output_path),
        )
