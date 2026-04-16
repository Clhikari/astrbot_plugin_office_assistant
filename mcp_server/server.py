from __future__ import annotations

from pathlib import Path

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

try:
    from ..domain.workbook.session_store import WorkbookSessionStore
except Exception:  # pragma: no cover - workbook domain may be provided by another worker.
    WorkbookSessionStore = None  # type: ignore[assignment]


def create_server(
    workspace_dir: Path | None = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    render_backend_config: DocumentRenderBackendConfig | None = None,
    default_document_style: dict[str, object] | None = None,
) -> FastMCP:
    server = FastMCP(
        name="astrbot-office-assistant",
        instructions=(
            "Stateful Office builder for structured Word and Excel generation. "
            "Use create_document/add_blocks/finalize_document/export_document for Word, "
            "and use create_workbook/write_rows/export_workbook for Excel."
        ),
    )
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    workbook_store = WorkbookSessionStore(workspace_dir=workspace_dir)
    attach_render_backend_config(store, render_backend_config)
    attach_document_style_defaults(store, default_document_style)
    register_document_tools(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
    )
    if WorkbookSessionStore is not None:
        workbook_store = WorkbookSessionStore(workspace_dir=workspace_dir)
        register_workbook_tools(server, workbook_store)
    return server


def main() -> None:
    create_server().run("stdio")


if __name__ == "__main__":
    main()
