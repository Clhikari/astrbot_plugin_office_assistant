from __future__ import annotations

from collections.abc import Awaitable, Callable
from pathlib import Path

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import ToolSet
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from .registry import build_document_store, get_document_tool_specs


def build_document_toolset_from_registry(
    workspace_dir: Path | None = None,
    after_export: (
        Callable[[ContextWrapper[AstrAgentContext], str], Awaitable[str | None]] | None
    ) = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
) -> ToolSet:
    store = build_document_store(workspace_dir=workspace_dir)
    resolved_before_export_hooks = before_export_hooks or []
    resolved_after_export_hooks = after_export_hooks or []
    tools = [
        spec.astrbot_factory(
            store,
            resolved_before_export_hooks,
            resolved_after_export_hooks,
            after_export,
        )
        for spec in get_document_tool_specs()
    ]
    toolset = ToolSet(tools)
    toolset.document_store = store
    return toolset
