from collections.abc import Awaitable, Callable
from pathlib import Path

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import ToolSet
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import DocumentRenderBackendConfig
from ..tools.astrbot_adapter import build_document_toolset_from_registry


def build_document_toolset(
    workspace_dir: Path | None = None,
    after_export: (
        Callable[[ContextWrapper[AstrAgentContext], str], Awaitable[str | None]] | None
    ) = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    render_backend_config: DocumentRenderBackendConfig | None = None,
) -> ToolSet:
    return build_document_toolset_from_registry(
        workspace_dir=workspace_dir,
        after_export=after_export,
        before_export_hooks=before_export_hooks,
        after_export_hooks=after_export_hooks,
        render_backend_config=render_backend_config,
    )


__all__ = ["build_document_toolset"]
