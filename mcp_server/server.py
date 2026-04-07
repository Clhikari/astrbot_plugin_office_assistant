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
from .tools import register_document_tools


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
            "Stateful document builder for complex Word generation. "
            "Use create_document first, then append content blocks, finalize, and export."
        ),
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
    return server


def main() -> None:
    create_server().run("stdio")


if __name__ == "__main__":
    main()
