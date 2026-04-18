from __future__ import annotations

import importlib
from pathlib import Path

from astrbot.api import logger
from mcp.server.fastmcp import FastMCP

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import (
    DocumentRenderBackendConfig,
    attach_render_backend_config,
)
from ..domain.document.session_store import (
    DocumentSessionStore,
    attach_document_style_defaults,
)
from .tools import register_document_tools, register_workbook_tools


def _load_workbook_session_store():
    try:
        workbook_module = importlib.import_module(
            "astrbot_plugin_office_assistant.domain.workbook.session_store"
        )
    except (ImportError, ModuleNotFoundError) as exc:  # pragma: no cover
        logger.warning(
            "[office-assistant] workbook support disabled during MCP server startup: %s",
            exc,
        )
        return None
    return workbook_module.WorkbookSessionStore


def create_server(
    workspace_dir: Path | None = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    render_backend_config: DocumentRenderBackendConfig | None = None,
    default_document_style: dict[str, object] | None = None,
) -> FastMCP:
    workbook_session_store_cls = _load_workbook_session_store()
    workbook_supported = workbook_session_store_cls is not None
    instructions = (
        "Stateful Office builder for structured Word generation. "
        "Use create_document/add_blocks/finalize_document/export_document for Word."
    )
    if workbook_supported:
        instructions += (
            " Use create_workbook/write_rows/export_workbook for Excel."
        )
    server = FastMCP(
        name="astrbot-office-assistant",
        instructions=instructions,
    )
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    attach_render_backend_config(store, render_backend_config)
    attach_document_style_defaults(store, default_document_style)
    register_document_tools(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
    )
    if workbook_supported:
        workbook_store = workbook_session_store_cls(workspace_dir=workspace_dir)
        register_workbook_tools(server, workbook_store)
    return server


def main() -> None:
    create_server().run("stdio")


if __name__ == "__main__":
    main()
