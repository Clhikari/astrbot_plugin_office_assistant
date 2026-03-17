from __future__ import annotations

from ..models.blocks import (
    BlockLayout,
    BlockStyle,
    GroupBlock,
    ListBlock,
    ParagraphBlock,
    SummaryCardBlock,
)


def _merge_style(
    base: BlockStyle | None,
    *,
    emphasis: str | None = None,
    font_scale: float | None = None,
) -> BlockStyle:
    return BlockStyle(
        align=getattr(base, "align", None),
        emphasis=emphasis if emphasis is not None else getattr(base, "emphasis", None),
        font_scale=(
            font_scale if font_scale is not None else getattr(base, "font_scale", None)
        ),
        table_grid=getattr(base, "table_grid", None),
        cell_align=getattr(base, "cell_align", None),
    )


def _merge_layout(
    base: BlockLayout | None,
    *,
    spacing_before: float | None = None,
    spacing_after: float | None = None,
) -> BlockLayout:
    return BlockLayout(
        spacing_before=(
            spacing_before
            if spacing_before is not None
            else getattr(base, "spacing_before", None)
        ),
        spacing_after=(
            spacing_after
            if spacing_after is not None
            else getattr(base, "spacing_after", None)
        ),
    )


def build_summary_card_group(
    *,
    title: str,
    items: list[str],
    variant: str = "summary",
    style: BlockStyle | None = None,
    layout: BlockLayout | None = None,
) -> GroupBlock:
    normalized_variant = "conclusion" if variant == "conclusion" else "summary"
    title_style = _merge_style(
        style,
        emphasis="subtle" if normalized_variant == "conclusion" else "strong",
        font_scale=1.05,
    )
    body_style = _merge_style(
        style,
        emphasis="subtle" if normalized_variant == "conclusion" else None,
    )
    title_layout = _merge_layout(layout, spacing_before=6, spacing_after=2)
    list_layout = _merge_layout(layout, spacing_before=0, spacing_after=6)
    return GroupBlock(
        blocks=[
            ParagraphBlock(
                text=title,
                style=title_style,
                layout=title_layout,
            ),
            ListBlock(
                items=items,
                ordered=False,
                style=body_style,
                layout=list_layout,
            ),
        ]
    )


def expand_summary_card_block(block: SummaryCardBlock) -> GroupBlock:
    return build_summary_card_group(
        title=block.title,
        items=list(block.items),
        variant=block.variant,
        style=block.style,
        layout=block.layout,
    )
