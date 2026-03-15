from mcp.server.fastmcp import FastMCP

from ..session_store import DocumentSessionStore
from .add_heading import register_add_heading_tool
from .add_paragraph import register_add_paragraph_tool
from .add_summary_card import register_add_summary_card_tool
from .add_table import register_add_table_tool
from .create_document import register_create_document_tool
from .export_document import register_export_document_tool
from .finalize_document import register_finalize_document_tool


def register_document_tools(server: FastMCP, store: DocumentSessionStore) -> None:
    register_create_document_tool(server, store)
    register_add_heading_tool(server, store)
    register_add_paragraph_tool(server, store)
    register_add_table_tool(server, store)
    register_add_summary_card_tool(server, store)
    register_finalize_document_tool(server, store)
    register_export_document_tool(server, store)
