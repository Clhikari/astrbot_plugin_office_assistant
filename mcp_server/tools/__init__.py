from mcp.server.fastmcp import FastMCP

from ..session_store import DocumentSessionStore
from .add_blocks import register_add_blocks_tool
from .create_document import register_create_document_tool
from .export_document import register_export_document_tool
from .finalize_document import register_finalize_document_tool


def register_document_tools(server: FastMCP, store: DocumentSessionStore) -> None:
    register_create_document_tool(server, store)
    register_add_blocks_tool(server, store)
    register_finalize_document_tool(server, store)
    register_export_document_tool(server, store)
