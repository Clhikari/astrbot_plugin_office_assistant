from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import ValidationError

from ...domain.workbook.session_store import WorkbookSessionStore
from ...domain.workbook.contracts import (
    ToolResult,
    WriteRowsRequest,
    build_workbook_summary,
)


def register_write_rows_tool(server: FastMCP, store: WorkbookSessionStore) -> None:
    @server.tool(
        name="write_rows",
        description="Write row data into one worksheet. Sheet is auto-created when missing.",
        structured_output=True,
    )
    def write_rows(
        workbook_id: str,
        sheet: str,
        rows: list[list[Any]],
        start_row: int = 1,
    ) -> ToolResult:
        try:
            request = WriteRowsRequest(
                workbook_id=workbook_id,
                sheet=sheet,
                rows=rows,
                start_row=start_row,
            )
            workbook = store.write_rows(request)
        except ValidationError as exc:
            return ToolResult(
                success=False,
                message=(
                    "write_rows failed. Retry write_rows with the same workbook_id "
                    f"and only fix invalid fields. Original error: {exc}"
                ),
            )
        except (KeyError, ValueError) as exc:
            return ToolResult(
                success=False,
                message=f"write_rows failed: {exc}",
            )
        return ToolResult(
            success=True,
            message="Rows written.",
            workbook=build_workbook_summary(workbook),
        )
