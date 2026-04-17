from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ValidationError
from pydantic import ConfigDict, Field
from pydantic.dataclasses import dataclass

from astrbot import logger
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import FunctionTool, ToolExecResult
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.workbook.contracts import (
    CreateWorkbookRequest,
    ExportWorkbookRequest,
    ExportWorkbookResult,
    ToolResult,
    WriteRowsRequest,
    build_workbook_summary,
)
from ..domain.workbook.session_store import WorkbookSessionStore


def _build_default_store() -> WorkbookSessionStore:
    from ..tools.registry import build_workbook_store

    return build_workbook_store()


def _dump_result(result: ToolResult) -> str:
    return result.model_dump_json(exclude_none=True)


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class WorkbookToolBase(FunctionTool[AstrAgentContext]):
    store: WorkbookSessionStore = Field(default_factory=_build_default_store)


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class CreateWorkbookTool(WorkbookToolBase):
    name: str = "create_workbook"
    description: str = (
        "Create a draft workbook session and return workbook_id for structured Excel generation."
    )
    parameters: dict[str, Any] = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "session_id": {
                    "type": "string",
                    "description": "Optional conversation session id.",
                },
                "title": {
                    "type": "string",
                    "description": "Optional workbook title for summary context.",
                },
                "filename": {
                    "type": "string",
                    "description": "Preferred output filename. Defaults to workbook.xlsx.",
                },
            },
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs: Any
    ) -> ToolExecResult:
        try:
            request = CreateWorkbookRequest(
                session_id=str(kwargs.get("session_id") or ""),
                title=str(kwargs.get("title") or ""),
                filename=str(kwargs.get("filename") or "workbook.xlsx"),
            )
            workbook = self.store.create_workbook(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        return _dump_result(
            ToolResult(
                success=True,
                message=(
                    "Workbook session created. Continue with write_rows for each batch, "
                    "then call export_workbook after all sheets are complete."
                ),
                workbook=build_workbook_summary(workbook),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class WriteRowsTool(WorkbookToolBase):
    name: str = "write_rows"
    description: str = (
        "Write row data into one worksheet. Sheet is auto-created when missing."
    )
    parameters: dict[str, Any] = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "workbook_id": {
                    "type": "string",
                    "description": "Target workbook_id returned by create_workbook.",
                },
                "sheet": {
                    "type": "string",
                    "description": "Target worksheet name.",
                },
                "rows": {
                    "type": "array",
                    "items": {"type": "array"},
                    "description": "2D row array. Cells can be string/number/bool/null.",
                },
                "start_row": {
                    "type": "integer",
                    "minimum": 1,
                    "description": "1-based start row. Defaults to 1.",
                },
            },
            "required": ["workbook_id", "sheet", "rows"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs: Any
    ) -> ToolExecResult:
        try:
            request = WriteRowsRequest(
                workbook_id=str(kwargs.get("workbook_id") or ""),
                sheet=str(kwargs.get("sheet") or ""),
                rows=list(kwargs.get("rows") or []),
                start_row=int(kwargs.get("start_row") or 1),
            )
            workbook = self.store.write_rows(request)
        except ValidationError as exc:
            return _dump_result(
                ToolResult(
                    success=False,
                    message=(
                        "write_rows failed. Retry write_rows with the same workbook_id "
                        f"and only fix invalid fields. Original error: {exc}"
                    ),
                )
            )
        except Exception as exc:
            return _dump_result(
                ToolResult(
                    success=False,
                    message=f"write_rows failed: {exc}",
                )
            )
        return _dump_result(
            ToolResult(
                success=True,
                message=(
                    "Rows written. Continue calling write_rows for additional sheets "
                    "or batches, then call export_workbook."
                ),
                workbook=build_workbook_summary(workbook),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class ExportWorkbookTool(WorkbookToolBase):
    name: str = "export_workbook"
    description: str = "Export workbook draft to .xlsx and return file path."
    parameters: dict[str, Any] = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "workbook_id": {
                    "type": "string",
                    "description": "Target workbook_id returned by create_workbook.",
                },
                "output_name": {
                    "type": "string",
                    "description": "Optional output filename.",
                },
            },
            "required": ["workbook_id"],
        }
    )
    after_export: (
        Callable[[ContextWrapper[AstrAgentContext], str], Awaitable[str | None]] | None
    ) = None

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs: Any
    ) -> ToolExecResult:
        try:
            request = ExportWorkbookRequest(
                workbook_id=str(kwargs.get("workbook_id") or ""),
                output_name=str(kwargs.get("output_name") or ""),
            )
            workbook, output_path = await asyncio.to_thread(
                self.store.export_workbook,
                request,
            )
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))

        callback_message = ""
        delivery_handled = False
        if self.after_export is not None and context is not None:
            try:
                logger.debug(
                    "[office-assistant] invoking workbook after_export callback for workbook=%s output=%s",
                    workbook.workbook_id,
                    output_path,
                )
                callback_message = (
                    await self.after_export(context, str(output_path)) or ""
                )
                delivery_handled = True
                logger.debug(
                    "[office-assistant] workbook after_export callback completed for workbook=%s output=%s delivered=%s",
                    workbook.workbook_id,
                    output_path,
                    delivery_handled,
                )
            except Exception as exc:
                logger.warning(
                    "[office-assistant] workbook after_export callback failed for %s: %s",
                    output_path,
                    exc,
                )
                callback_message = (
                    f"Workbook exported, but post-export delivery failed: {exc}"
                )
        if delivery_handled:
            return None

        return _dump_result(
            ExportWorkbookResult(
                success=True,
                message=callback_message or "Workbook exported.",
                workbook=build_workbook_summary(workbook),
                file_path=str(output_path),
            )
        )


__all__ = ["CreateWorkbookTool", "ExportWorkbookTool", "WriteRowsTool"]
