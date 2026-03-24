from collections.abc import Awaitable, Callable
from pathlib import Path

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import ToolSet
from astrbot.core.astr_agent_context import AstrAgentContext

from ..internal_hooks import AfterExportHook, BeforeExportHook
from ..mcp_server.session_store import DocumentSessionStore
from .document_tools import (
    AddBlocksTool,
    CreateDocumentTool,
    ExportDocumentTool,
    FinalizeDocumentTool,
)


def build_document_toolset(
    workspace_dir: Path | None = None,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    after_export: (
        Callable[[ContextWrapper[AstrAgentContext], str], Awaitable[str | None]] | None
    ) = None,
) -> ToolSet:
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    return ToolSet(
        [
            CreateDocumentTool(store=store),
            AddBlocksTool(store=store),
            FinalizeDocumentTool(store=store),
            ExportDocumentTool(
                store=store,
                before_export_hooks=before_export_hooks or [],
                after_export_hooks=after_export_hooks or [],
                after_export=after_export,
            ),
        ]
    )


__all__ = ["build_document_toolset"]
