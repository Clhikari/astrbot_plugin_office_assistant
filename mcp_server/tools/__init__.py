from mcp.server.fastmcp import FastMCP

from ...domain.document.hooks import AfterExportHook, BeforeExportHook
from ...domain.document.session_store import DocumentSessionStore
from ...tools.mcp_adapter import register_document_tools_from_registry


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
