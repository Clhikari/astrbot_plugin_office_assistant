from __future__ import annotations

from collections.abc import Awaitable, Callable, Sequence
from dataclasses import dataclass, field
from inspect import isawaitable
from typing import Any


@dataclass(slots=True)
class NoticeBuildContext:
    event: Any
    request: Any
    should_expose: bool
    can_process_upload: bool
    explicit_tool_name: str | None
    notices: list[str] = field(default_factory=list)
    section_names: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ToolExposureContext:
    event: Any
    request: Any
    should_expose: bool
    can_process_upload: bool
    explicit_tool_name: str | None


@dataclass(slots=True)
class BlockNormalizationContext:
    document: Any
    incoming_blocks: list[Any]
    source: str


@dataclass(slots=True)
class ExportPreparationContext:
    document: Any
    output_path: Any
    source: str


@dataclass(slots=True)
class AfterExportContext:
    document: Any
    output_path: Any
    source: str


NoticeBuildHook = Callable[
    [NoticeBuildContext],
    NoticeBuildContext | None | Awaitable[NoticeBuildContext | None],
]
ToolExposureHook = Callable[
    [ToolExposureContext],
    ToolExposureContext | None | Awaitable[ToolExposureContext | None],
]
BlockNormalizeHook = Callable[[BlockNormalizationContext], list[Any] | None]
BeforeExportHook = Callable[
    [ExportPreparationContext],
    ExportPreparationContext | None | Awaitable[ExportPreparationContext | None],
]
AfterExportHook = Callable[
    [AfterExportContext],
    AfterExportContext | None | Awaitable[AfterExportContext | None],
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


def run_block_normalize_hooks(
    hooks: Sequence[BlockNormalizeHook],
    context: BlockNormalizationContext,
) -> list[Any]:
    blocks = list(context.incoming_blocks)
    for hook in hooks:
        result = hook(
            BlockNormalizationContext(
                document=context.document,
                incoming_blocks=blocks,
                source=context.source,
            )
        )
        if result is not None:
            blocks = list(result)
    return blocks


async def run_before_export_hooks(
    hooks: Sequence[BeforeExportHook],
    context: ExportPreparationContext,
) -> ExportPreparationContext:
    current = context
    for hook in hooks:
        result = hook(current)
        if isawaitable(result):
            result = await result
        if result is not None:
            current = result
    return current


async def run_after_export_hooks(
    hooks: Sequence[AfterExportHook],
    context: AfterExportContext,
) -> AfterExportContext:
    current = context
    for hook in hooks:
        result = hook(current)
        if isawaitable(result):
            result = await result
        if result is not None:
            current = result
    return current
