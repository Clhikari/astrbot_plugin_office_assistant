from __future__ import annotations

from collections.abc import Awaitable, Callable, Sequence
from dataclasses import dataclass
from inspect import isawaitable
from typing import Any


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


BlockNormalizeHook = Callable[[BlockNormalizationContext], list[Any] | None]
BeforeExportHook = Callable[
    [ExportPreparationContext],
    ExportPreparationContext | None | Awaitable[ExportPreparationContext | None],
]
AfterExportHook = Callable[
    [AfterExportContext],
    AfterExportContext | None | Awaitable[AfterExportContext | None],
]


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
