from __future__ import annotations

__all__ = [
    "CreateWorkbookRequest",
    "ExportWorkbookRequest",
    "ExportWorkbookResult",
    "ToolResult",
    "WorkbookModel",
    "WorkbookSessionStore",
    "WorkbookStatus",
    "WorkbookSummary",
    "WriteRowsRequest",
    "build_workbook_summary",
    "export_workbook_to_xlsx",
]


def __getattr__(name: str):
    if name == "WorkbookSessionStore":
        from .session_store import WorkbookSessionStore

        return WorkbookSessionStore
    if name in {
        "WorkbookModel",
        "WorkbookStatus",
    }:
        from . import models as models_module

        return getattr(models_module, name)
    if name in {
        "CreateWorkbookRequest",
        "ExportWorkbookRequest",
        "ExportWorkbookResult",
        "ToolResult",
        "WorkbookSummary",
        "WriteRowsRequest",
        "build_workbook_summary",
    }:
        from .contracts import __dict__ as contracts_namespace

        return contracts_namespace[name]
    if name == "export_workbook_to_xlsx":
        from .exporter import export_workbook_to_xlsx

        return export_workbook_to_xlsx
    raise AttributeError(name)
