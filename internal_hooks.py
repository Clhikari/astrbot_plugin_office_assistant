from __future__ import annotations

from collections.abc import Awaitable, Callable, Sequence
from dataclasses import dataclass, field
from inspect import isawaitable
from typing import Any

from .domain.document.hooks import (
    AfterExportContext,
    AfterExportHook,
    BeforeExportHook,
    BlockNormalizationContext,
    BlockNormalizeHook,
    ExportPreparationContext,
    run_after_export_hooks,
    run_before_export_hooks,
    run_block_normalize_hooks,
)

__all__ = [
    "AfterExportContext",
    "AfterExportHook",
    "BeforeExportHook",
    "BlockNormalizationContext",
    "BlockNormalizeHook",
    "ExportPreparationContext",
    "NoticeBuildContext",
    "NoticeBuildHook",
    "ToolExposureContext",
    "ToolExposureHook",
    "run_after_export_hooks",
    "run_before_export_hooks",
    "run_block_normalize_hooks",
    "run_notice_hooks",
    "run_tool_exposure_hooks",
]


@dataclass(slots=True)
class NoticeBuildContext:
    event: Any
    request: Any
    should_expose: bool
    can_process_upload: bool
    explicit_tool_name: str | None
    notices: list[str] = field(default_factory=list)
    section_names: list[str] = field(default_factory=list)
    system_notices: list[str] = field(default_factory=list)
    system_section_names: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ToolExposureContext:
    event: Any
    request: Any
    should_expose: bool
    can_process_upload: bool
    explicit_tool_name: str | None


NoticeBuildHook = Callable[
    [NoticeBuildContext],
    NoticeBuildContext | None | Awaitable[NoticeBuildContext | None],
]
ToolExposureHook = Callable[
    [ToolExposureContext],
    ToolExposureContext | None | Awaitable[ToolExposureContext | None],
]


async def run_notice_hooks(
    hooks: Sequence[NoticeBuildHook],
    context: NoticeBuildContext,
) -> NoticeBuildContext:
    current = context
    for hook in hooks:
        result = hook(current)
        if isawaitable(result):
            result = await result
        if result is not None:
            current = result
    return current


async def run_tool_exposure_hooks(
    hooks: Sequence[ToolExposureHook],
    context: ToolExposureContext,
) -> ToolExposureContext:
    current = context
    for hook in hooks:
        result = hook(current)
        if isawaitable(result):
            result = await result
        if result is not None:
            current = result
    return current
