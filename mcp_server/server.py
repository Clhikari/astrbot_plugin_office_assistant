from __future__ import annotations

from pathlib import Path

from mcp.server.fastmcp import FastMCP

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.session_store import DocumentSessionStore
from .tools import register_document_tools


def create_server(
    workspace_dir: Path | None = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> FastMCP:
    server = FastMCP(
        name="astrbot-office-assistant",
        instructions=(
            "Stateful document builder for complex Word generation. "
            "Use create_document first, then append content blocks, finalize, and export."
        ),
    )
    store = DocumentSessionStore(workspace_dir=workspace_dir)
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
