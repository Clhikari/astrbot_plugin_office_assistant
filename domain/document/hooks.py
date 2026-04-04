from __future__ import annotations

from collections.abc import Awaitable, Callable, Sequence
from dataclasses import dataclass
from inspect import isawaitable
from typing import Any, TypeVar

from astrbot.api import logger

from .contracts import BlockInput


@dataclass(slots=True)
class BlockNormalizationContext:
    document: Any
    incoming_blocks: list[BlockInput]
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


BlockNormalizeHook = Callable[
    [BlockNormalizationContext], Sequence[BlockInput] | None
]
BeforeExportHook = Callable[
    [ExportPreparationContext],
    ExportPreparationContext | None | Awaitable[ExportPreparationContext | None],
]
AfterExportHook = Callable[
    [AfterExportContext],
    AfterExportContext | None | Awaitable[AfterExportContext | None],
]

HookContextT = TypeVar("HookContextT")


def run_block_normalize_hooks(
    hooks: Sequence[BlockNormalizeHook],
    context: BlockNormalizationContext,
) -> list[BlockInput]:
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
    logger.debug(
        "[office-assistant] running %d before_export hooks for document=%s source=%s output=%s",
        len(hooks),
        getattr(context.document, "document_id", ""),
        context.source,
        context.output_path,
    )
    return await _run_async_hooks(hooks, context)


async def run_after_export_hooks(
    hooks: Sequence[AfterExportHook],
    context: AfterExportContext,
) -> AfterExportContext:
    logger.debug(
        "[office-assistant] running %d after_export hooks for document=%s source=%s output=%s",
        len(hooks),
        getattr(context.document, "document_id", ""),
        context.source,
        context.output_path,
    )
    return await _run_async_hooks(hooks, context)


async def _run_async_hooks(
    hooks: Sequence[Callable[[HookContextT], HookContextT | None | Awaitable[HookContextT | None]]],
    context: HookContextT,
) -> HookContextT:
    current = context
    for hook in hooks:
        logger.debug(
            "[office-assistant] executing hook=%s context=%s",
            getattr(hook, "__name__", hook.__class__.__name__),
            type(current).__name__,
        )
        result = hook(current)
        if isawaitable(result):
            result = await result
        if result is not None:
            current = result
    logger.debug(
        "[office-assistant] completed %d hooks for context=%s",
        len(hooks),
        type(current).__name__,
    )
    return current


__all__ = [
    "AfterExportContext",
    "AfterExportHook",
    "BeforeExportHook",
    "BlockNormalizationContext",
    "BlockNormalizeHook",
    "ExportPreparationContext",
    "run_after_export_hooks",
    "run_before_export_hooks",
    "run_block_normalize_hooks",
]
