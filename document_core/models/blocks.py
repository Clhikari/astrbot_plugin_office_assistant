from __future__ import annotations

from typing import Annotated, Literal
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field


class BlockBase(BaseModel):
    model_config = ConfigDict(extra="forbid")

    block_id: str = Field(default_factory=lambda: uuid4().hex)
    type: str


class HeadingBlock(BlockBase):
    type: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)


class ParagraphBlock(BlockBase):
    type: Literal["paragraph"] = "paragraph"
    text: str = Field(min_length=1)


class TableBlock(BlockBase):
    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    table_style: Literal["report_grid", "metrics_compact", "minimal"] = "report_grid"


class SummaryCardBlock(BlockBase):
    type: Literal["summary_card"] = "summary_card"
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: Literal["summary", "conclusion"] = "summary"


class ImageBlock(BlockBase):
    type: Literal["image"] = "image"
    path: str = Field(min_length=1)
    caption: str = ""
    width_px: int | None = Field(default=None, gt=0)


DocumentBlock = Annotated[
    HeadingBlock | ParagraphBlock | TableBlock | SummaryCardBlock | ImageBlock,
    Field(discriminator="type"),
]
