from __future__ import annotations

from typing import TYPE_CHECKING

from mcp.server.fastmcp import FastMCP

from ...domain.document.hooks import AfterExportHook, BeforeExportHook
from ...domain.document.session_store import DocumentSessionStore
from ...tools.mcp_adapter import register_document_tools_from_registry
from ...tools.mcp_adapter import register_workbook_tools_from_registry

if TYPE_CHECKING:
    from ...domain.workbook.session_store import WorkbookSessionStore


def register_document_tools(
    server: FastMCP,
    store: DocumentSessionStore,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> None:
    register_document_tools_from_registry(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
    )


def register_workbook_tools(
    server: FastMCP,
    store: WorkbookSessionStore,
) -> None:
    register_workbook_tools_from_registry(server, store)
