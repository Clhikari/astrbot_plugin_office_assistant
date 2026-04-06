from __future__ import annotations

import warnings

from .models.blocks import (
    HeadingBlock,
    ImageBlock,
    ParagraphBlock,
    TableBlock,
)
from .models.document import DocumentModel, DocumentStatus

__all__ = [
    "DocumentModel",
    "DocumentStatus",
    "HeadingBlock",
    "ParagraphBlock",
    "TableBlock",
    "ImageBlock",
]


def __getattr__(name: str):
    if name == "WordDocumentBuilder":
        warnings.warn(
            "document_core.WordDocumentBuilder is legacy. Use the document render backend pipeline instead.",
            DeprecationWarning,
            stacklevel=2,
        )
        from .builders.word_builder import WordDocumentBuilder

        return WordDocumentBuilder
    raise AttributeError(name)
