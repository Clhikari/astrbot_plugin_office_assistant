from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING

from astrbot.core.agent.tool import ToolSet

from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import DocumentRenderBackendConfig
from .registry import (
    AfterExportCallback,
    AstrBotDocumentTool,
    build_document_store,
    get_document_tool_specs,
)

if TYPE_CHECKING:
    from ..domain.document.session_store import DocumentSessionStore


class DocumentToolSet(ToolSet):
    document_store: DocumentSessionStore

    def __init__(
        self,
        tools: Sequence[AstrBotDocumentTool],
        *,
        document_store: DocumentSessionStore,
    ) -> None:
        super().__init__(list(tools))
        self.document_store = document_store


def build_document_toolset_from_registry(
    workspace_dir: Path | None = None,
    after_export: AfterExportCallback | None = None,
    *,
    before_export_hooks: list[BeforeExportHook] | None = None,
    after_export_hooks: list[AfterExportHook] | None = None,
    render_backend_config: DocumentRenderBackendConfig | None = None,
) -> DocumentToolSet:
    store = build_document_store(
        workspace_dir=workspace_dir,
        render_backend_config=render_backend_config,
    )
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
    return DocumentToolSet(tools, document_store=store)
