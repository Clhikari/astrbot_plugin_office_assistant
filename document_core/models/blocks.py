from __future__ import annotations

from typing import Annotated, Literal
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class BlockLayout(BaseModel):
    model_config = ConfigDict(extra="forbid")

    spacing_before: float | None = Field(default=None, ge=0, le=72)
    spacing_after: float | None = Field(default=None, ge=0, le=72)


class BlockStyle(BaseModel):
    model_config = ConfigDict(extra="forbid")

    align: Literal["left", "center", "right", "justify"] | None = None
    emphasis: Literal["normal", "strong", "subtle"] | None = None
    font_scale: float | None = Field(default=None, ge=0.75, le=2.0)
    table_grid: Literal["report_grid", "metrics_compact", "minimal"] | None = None
    cell_align: Literal["left", "center", "right"] | None = None


class BlockBase(BaseModel):
    model_config = ConfigDict(extra="forbid")

    block_id: str = Field(default_factory=lambda: uuid4().hex)
    type: str
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class HeadingBlock(BlockBase):
    type: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)


class ParagraphRun(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = Field(min_length=1)
    bold: bool = False
    italic: bool = False
    underline: bool = False
    code: bool = False


class ParagraphBlock(BlockBase):
    type: Literal["paragraph"] = "paragraph"
    text: str = ""
    variant: Literal["body", "summary_box", "key_takeaway"] = "body"
    title: str = ""
    runs: list[ParagraphRun] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> ParagraphBlock:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("paragraph requires text or runs")


class ListBlock(BlockBase):
    type: Literal["list"] = "list"
    items: list[str] = Field(min_length=1)
    ordered: bool = False

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned


class TableHeaderGroup(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    span: int = Field(ge=1)


def resolve_table_column_count(headers: list[str], rows: list[list[str]]) -> int:
    if headers:
        return len(headers)
    if rows:
        return max(len(row) for row in rows)
    return 0


def validate_table_structure(
    headers: list[str],
    rows: list[list[str]],
    header_groups: list[TableHeaderGroup] | None = None,
) -> None:
    column_count = resolve_table_column_count(headers, rows)
    if column_count <= 0:
        if header_groups:
            raise ValueError(
                "header_groups require at least one column from headers or rows "
                f"(column_count={column_count})"
            )
        return
    if not header_groups:
        return
    total_span = sum(group.span for group in header_groups)
    if total_span != column_count:
        raise ValueError(
            f"header_groups span total ({total_span}) must equal column count "
            f"({column_count})"
        )


class TableBlock(BlockBase):
    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: Literal["report_grid", "metrics_compact", "minimal"] = "report_grid"
    caption: str = ""
    column_widths: list[float] = Field(default_factory=list)
    numeric_columns: list[int] = Field(default_factory=list)

    @field_validator("column_widths")
    @classmethod
    def validate_column_widths(cls, value: list[float]) -> list[float]:
        return [width if width > 0 else 0 for width in value]

    @field_validator("numeric_columns")
    @classmethod
    def validate_numeric_columns(cls, value: list[int]) -> list[int]:
        cleaned = sorted({index for index in value if index >= 0})
        return cleaned

    @model_validator(mode="after")
    def validate_table_shape(self) -> TableBlock:
        validate_table_structure(self.headers, self.rows, self.header_groups)
        return self


class SummaryCardBlock(BlockBase):
    # Compatibility-only block. Writers may still emit it, but renderers should
    # expand it into standard primitives instead of treating it as a core
    # rendering branch.
    type: Literal["summary_card"] = "summary_card"
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: Literal["summary", "conclusion"] = "summary"


class ImageBlock(BlockBase):
    type: Literal["image"] = "image"
    path: str = Field(min_length=1)
    caption: str = ""
    width_px: int | None = Field(default=None, gt=0)


class GroupBlock(BlockBase):
    type: Literal["group"] = "group"
    blocks: list[DocumentBlock] = Field(min_length=1)


class ColumnBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blocks: list[DocumentBlock] = Field(min_length=1)


class ColumnsBlock(BlockBase):
    type: Literal["columns"] = "columns"
    columns: list[ColumnBlock] = Field(min_length=1, max_length=3)


class PageBreakBlock(BlockBase):
    type: Literal["page_break"] = "page_break"


DocumentBlock = Annotated[
    HeadingBlock
    | ParagraphBlock
    | ListBlock
    | TableBlock
    | SummaryCardBlock
    | ImageBlock
    | GroupBlock
    | ColumnsBlock
    | PageBreakBlock,
    Field(discriminator="type"),
]

GroupBlock.model_rebuild()
ColumnBlock.model_rebuild()
ColumnsBlock.model_rebuild()
