from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from .registry import get_document_tool_specs


def register_document_tools_from_registry(
    server: FastMCP,
    store: Any,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> None:
    resolved_before_export_hooks = before_export_hooks or []
    resolved_after_export_hooks = after_export_hooks or []
    for spec in get_document_tool_specs():
        spec.mcp_registrar(
            server,
            store,
            resolved_before_export_hooks,
            resolved_after_export_hooks,
        )
