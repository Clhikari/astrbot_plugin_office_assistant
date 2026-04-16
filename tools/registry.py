from __future__ import annotations

from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Protocol

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import (
    DocumentRenderBackendConfig,
    attach_render_backend_config,
    build_document_render_backends,
    get_render_backend_config,
)
from ..domain.document.session_store import DocumentSessionStore
from ..domain.document.session_store import attach_document_style_defaults

if TYPE_CHECKING:
    from mcp.server.fastmcp import FastMCP
    from ..domain.workbook.session_store import WorkbookSessionStore


class AstrBotDocumentTool(Protocol):
    name: str


class AstrBotWorkbookTool(Protocol):
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

WorkbookAfterExportCallback = Callable[
    [ContextWrapper[AstrAgentContext], str], Awaitable[str | None]
]
AstrBotWorkbookToolFactory = Callable[
    ["WorkbookSessionStore", WorkbookAfterExportCallback | None],
    AstrBotWorkbookTool,
]
WorkbookMcpToolRegistrar = Callable[["FastMCP", "WorkbookSessionStore"], None]


@dataclass(frozen=True)
class DocumentToolSpec:
    name: str
    astrbot_factory: AstrBotToolFactory
    mcp_registrar: McpToolRegistrar


@dataclass(frozen=True)
class WorkbookToolSpec:
    name: str
    astrbot_factory: AstrBotWorkbookToolFactory
    mcp_registrar: WorkbookMcpToolRegistrar


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
    document_format = "word"
    return ExportDocumentTool(
        store=store,
        render_backends=build_document_render_backends(
            document_format,
            render_backend_config,
        ),
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
    document_format = "word"
    register_export_document_tool(
        server,
        store,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
        render_backends=build_document_render_backends(
            document_format,
            render_backend_config,
        ),
        render_backend_config=render_backend_config,
    )


def _build_create_workbook_tool(
    store: "WorkbookSessionStore",
    _after_export: WorkbookAfterExportCallback | None,
) -> AstrBotWorkbookTool:
    from ..agent_tools.workbook_tools import CreateWorkbookTool

    return CreateWorkbookTool(store=store)


def _build_write_rows_tool(
    store: "WorkbookSessionStore",
    _after_export: WorkbookAfterExportCallback | None,
) -> AstrBotWorkbookTool:
    from ..agent_tools.workbook_tools import WriteRowsTool

    return WriteRowsTool(store=store)


def _build_export_workbook_tool(
    store: "WorkbookSessionStore",
    after_export: WorkbookAfterExportCallback | None,
) -> AstrBotWorkbookTool:
    from ..agent_tools.workbook_tools import ExportWorkbookTool

    return ExportWorkbookTool(store=store, after_export=after_export)


def _register_create_workbook_tool(
    server: FastMCP,
    store: "WorkbookSessionStore",
) -> None:
    from ..mcp_server.tools.create_workbook import register_create_workbook_tool

    register_create_workbook_tool(server, store)


def _register_write_rows_tool(
    server: FastMCP,
    store: "WorkbookSessionStore",
) -> None:
    from ..mcp_server.tools.write_rows import register_write_rows_tool

    register_write_rows_tool(server, store)


def _register_export_workbook_tool(
    server: FastMCP,
    store: "WorkbookSessionStore",
) -> None:
    from ..mcp_server.tools.export_workbook import register_export_workbook_tool

    register_export_workbook_tool(server, store)


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

_WORKBOOK_TOOL_SPECS: tuple[WorkbookToolSpec, ...] = (
    WorkbookToolSpec(
        name="create_workbook",
        astrbot_factory=_build_create_workbook_tool,
        mcp_registrar=_register_create_workbook_tool,
    ),
    WorkbookToolSpec(
        name="write_rows",
        astrbot_factory=_build_write_rows_tool,
        mcp_registrar=_register_write_rows_tool,
    ),
    WorkbookToolSpec(
        name="export_workbook",
        astrbot_factory=_build_export_workbook_tool,
        mcp_registrar=_register_export_workbook_tool,
    ),
)


def get_document_tool_specs() -> tuple[DocumentToolSpec, ...]:
    return _DOCUMENT_TOOL_SPECS


def get_workbook_tool_specs() -> tuple[WorkbookToolSpec, ...]:
    return _WORKBOOK_TOOL_SPECS


def build_document_store(
    workspace_dir: Path | None = None,
    *,
    render_backend_config=None,
    default_document_style: dict[str, object] | None = None,
) -> DocumentSessionStore:
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    attach_render_backend_config(
        store,
        render_backend_config or DocumentRenderBackendConfig(),
    )
    attach_document_style_defaults(store, default_document_style)
    return store


def build_workbook_store(
    workspace_dir: Path | None = None,
) -> "WorkbookSessionStore":
    # Lazy import keeps the document-only path free from workbook dependencies.
    from ..domain.workbook.session_store import WorkbookSessionStore

    return WorkbookSessionStore(workspace_dir=workspace_dir)
