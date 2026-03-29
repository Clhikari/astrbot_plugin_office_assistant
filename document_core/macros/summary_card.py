from __future__ import annotations

from ..models.blocks import (
    BlockLayout,
    BlockStyle,
    GroupBlock,
    ListBlock,
    ParagraphBlock,
    SummaryCardBlock,
)

_SUMMARY_CARD_DEFAULT_FIELDS = (
    "title_align",
    "title_emphasis",
    "title_font_scale",
    "title_space_before",
    "title_space_after",
    "list_space_after",
)


def _summary_card_default_kwargs(
    values: dict[str, object | None],
) -> dict[str, object | None]:
    return {field: values.get(field) for field in _SUMMARY_CARD_DEFAULT_FIELDS}


def summary_card_defaults_from_config(config) -> dict[str, object | None]:
    if config is None:
        return _summary_card_default_kwargs({})
    return _summary_card_default_kwargs(
        {field: getattr(config, field, None) for field in _SUMMARY_CARD_DEFAULT_FIELDS}
    )


def summary_card_defaults_from_theme(theme: dict) -> dict[str, object | None]:
    return _summary_card_default_kwargs(
        {
            "title_align": theme.get("summary_card_title_align"),
            "title_emphasis": theme.get("summary_card_title_emphasis"),
            "title_font_scale": theme.get("summary_card_title_font_scale"),
            "title_space_before": theme.get("summary_card_title_space_before"),
            "title_space_after": theme.get("summary_card_title_space_after"),
            "list_space_after": theme.get("summary_card_list_space_after"),
        }
    )


def _merge_style(
    base: BlockStyle | None,
    *,
    align: str | None = None,
    emphasis: str | None = None,
    font_scale: float | None = None,
) -> BlockStyle:
    base_align = getattr(base, "align", None)
    base_emphasis = getattr(base, "emphasis", None)
    base_font_scale = getattr(base, "font_scale", None)
    # Summary card block styles stay at the top of the precedence chain.
    # Document defaults and variant defaults only fill missing values.
    return BlockStyle(
        align=base_align if base_align is not None else align,
        emphasis=base_emphasis if base_emphasis is not None else emphasis,
        font_scale=base_font_scale if base_font_scale is not None else font_scale,
        table_grid=getattr(base, "table_grid", None),
        cell_align=getattr(base, "cell_align", None),
    )


def _merge_layout(
    base: BlockLayout | None,
    *,
    spacing_before: float | None = None,
    spacing_after: float | None = None,
) -> BlockLayout:
    base_spacing_before = getattr(base, "spacing_before", None)
    base_spacing_after = getattr(base, "spacing_after", None)
    return BlockLayout(
        spacing_before=(
            base_spacing_before if base_spacing_before is not None else spacing_before
        ),
        spacing_after=(
            base_spacing_after if base_spacing_after is not None else spacing_after
        ),
    )


def build_summary_card_group(
    *,
    title: str,
    items: list[str],
    variant: str = "summary",
    style: BlockStyle | None = None,
    layout: BlockLayout | None = None,
    title_align: str | None = None,
    title_emphasis: str | None = None,
    title_font_scale: float | None = None,
    title_space_before: float | None = None,
    title_space_after: float | None = None,
    list_space_after: float | None = None,
) -> GroupBlock:
    normalized_variant = "conclusion" if variant == "conclusion" else "summary"
    title_style = _merge_style(
        style,
        align=title_align,
        emphasis=(
            title_emphasis
            if title_emphasis is not None
            else ("subtle" if normalized_variant == "conclusion" else "strong")
        ),
        font_scale=title_font_scale if title_font_scale is not None else 1.05,
    )
    body_style = _merge_style(
        style,
        emphasis="subtle" if normalized_variant == "conclusion" else None,
    )
    title_layout = _merge_layout(
        layout,
        spacing_before=title_space_before if title_space_before is not None else 6,
        spacing_after=title_space_after if title_space_after is not None else 2,
    )
    list_layout = _merge_layout(
        layout,
        spacing_before=0,
        spacing_after=list_space_after if list_space_after is not None else 6,
    )
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


def expand_summary_card_block(
    block: SummaryCardBlock,
    *,
    title_align: str | None = None,
    title_emphasis: str | None = None,
    title_font_scale: float | None = None,
    title_space_before: float | None = None,
    title_space_after: float | None = None,
    list_space_after: float | None = None,
) -> GroupBlock:
    return build_summary_card_group(
        title=block.title,
        items=list(block.items),
        variant=block.variant,
        style=block.style,
        layout=block.layout,
        title_align=title_align,
        title_emphasis=title_emphasis,
        title_font_scale=title_font_scale,
        title_space_before=title_space_before,
        title_space_after=title_space_after,
        list_space_after=list_space_after,
    )
