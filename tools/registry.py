from __future__ import annotations

from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Protocol

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import (
    attach_render_backend_config,
    build_word_render_backends,
    get_render_backend_config,
)
from ..domain.document.session_store import DocumentSessionStore

if TYPE_CHECKING:
    from mcp.server.fastmcp import FastMCP


class AstrBotDocumentTool(Protocol):
    name: str


AfterExportCallback = Callable[
    [ContextWrapper[AstrAgentContext], str], Awaitable[str | None]
]
AstrBotToolFactory = Callable[
    [
        DocumentSessionStore,
        list[BeforeExportHook],
        list[AfterExportHook],
        AfterExportCallback | None,
    ],
    AstrBotDocumentTool,
]
McpToolRegistrar = Callable[
    ["FastMCP", DocumentSessionStore, list[BeforeExportHook], list[AfterExportHook]],
    None,
]


@dataclass(frozen=True)
class DocumentToolSpec:
    name: str
    astrbot_factory: AstrBotToolFactory
    mcp_registrar: McpToolRegistrar


def _build_create_document_tool(
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
    _after_export: AfterExportCallback | None,
) -> AstrBotDocumentTool:
    from ..agent_tools.document_tools import CreateDocumentTool

    return CreateDocumentTool(store=store)


def _build_add_blocks_tool(
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
    _after_export: AfterExportCallback | None,
) -> AstrBotDocumentTool:
    from ..agent_tools.document_tools import AddBlocksTool

    return AddBlocksTool(store=store)


def _build_finalize_document_tool(
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
    _after_export: AfterExportCallback | None,
) -> AstrBotDocumentTool:
    from ..agent_tools.document_tools import FinalizeDocumentTool

    return FinalizeDocumentTool(store=store)


def _build_export_document_tool(
    store: DocumentSessionStore,
    before_export_hooks: list[BeforeExportHook],
    after_export_hooks: list[AfterExportHook],
    after_export: AfterExportCallback | None,
) -> AstrBotDocumentTool:
    from ..agent_tools.document_tools import ExportDocumentTool

    render_backend_config = get_render_backend_config(store)
    return ExportDocumentTool(
        store=store,
        render_backends=build_word_render_backends(render_backend_config),
        render_backend_config=render_backend_config,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
        after_export=after_export,
    )


def _register_create_document_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
) -> None:
    from ..mcp_server.tools.create_document import register_create_document_tool

    register_create_document_tool(server, store)


def _register_add_blocks_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
) -> None:
    from ..mcp_server.tools.add_blocks import register_add_blocks_tool

    register_add_blocks_tool(server, store)


def _register_finalize_document_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    _before_export_hooks: list[BeforeExportHook],
    _after_export_hooks: list[AfterExportHook],
) -> None:
    from ..mcp_server.tools.finalize_document import register_finalize_document_tool

    register_finalize_document_tool(server, store)


def _register_export_document_tool(
    server: FastMCP,
    store: DocumentSessionStore,
    before_export_hooks: list[BeforeExportHook],
    after_export_hooks: list[AfterExportHook],
) -> None:
    from ..mcp_server.tools.export_document import register_export_document_tool

    render_backend_config = get_render_backend_config(store)
    register_export_document_tool(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
        render_backends=build_word_render_backends(render_backend_config),
        render_backend_config=render_backend_config,
    )


_DOCUMENT_TOOL_SPECS: tuple[DocumentToolSpec, ...] = (
    DocumentToolSpec(
        name="create_document",
        astrbot_factory=_build_create_document_tool,
        mcp_registrar=_register_create_document_tool,
    ),
    DocumentToolSpec(
        name="add_blocks",
        astrbot_factory=_build_add_blocks_tool,
        mcp_registrar=_register_add_blocks_tool,
    ),
    DocumentToolSpec(
        name="finalize_document",
        astrbot_factory=_build_finalize_document_tool,
        mcp_registrar=_register_finalize_document_tool,
    ),
    DocumentToolSpec(
        name="export_document",
        astrbot_factory=_build_export_document_tool,
        mcp_registrar=_register_export_document_tool,
    ),
)


def get_document_tool_specs() -> tuple[DocumentToolSpec, ...]:
    return _DOCUMENT_TOOL_SPECS


def build_document_store(
    workspace_dir: Path | None = None,
    *,
    render_backend_config=None,
) -> DocumentSessionStore:
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    attach_render_backend_config(store, render_backend_config)
    return store
