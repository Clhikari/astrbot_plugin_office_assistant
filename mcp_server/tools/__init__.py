from mcp.server.fastmcp import FastMCP

from ...internal_hooks import AfterExportHook, BeforeExportHook
from ..session_store import DocumentSessionStore
from .add_blocks import register_add_blocks_tool
from .create_document import register_create_document_tool
from .export_document import register_export_document_tool
from .finalize_document import register_finalize_document_tool


def register_document_tools(
    server: FastMCP,
    store: DocumentSessionStore,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> None:
    register_create_document_tool(server, store)
    register_add_blocks_tool(server, store)
    register_finalize_document_tool(server, store)
    register_export_document_tool(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
    )
