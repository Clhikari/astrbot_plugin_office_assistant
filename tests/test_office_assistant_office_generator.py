from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest
from astrbot_plugin_office_assistant.document_core.models.blocks import (
    ParagraphBlock,
    TableBlock,
)
from astrbot_plugin_office_assistant.constants import OfficeType
from astrbot_plugin_office_assistant.office_generator import (
    OfficeGenerator,
)


def test_build_word_document_model_skips_invalid_blocks(workspace_root: Path):
    generator = OfficeGenerator(data_path=workspace_root)
    file_path = workspace_root / "report.docx"

    document = generator._build_word_document_model(
        file_path,
        {
            "metadata": {"title": "Quarterly Report"},
            "blocks": [
                {"type": "paragraph", "text": "Keep this paragraph."},
                {"type": "unknown", "text": "Skip this unsupported block."},
                {"type": "heading", "text": "", "level": 1},
                {
                    "type": "table",
                    "headers": ["Metric", "Value"],
                    "rows": [["Users", "42"]],
                },
            ],
        },
    )

    assert document.metadata.preferred_filename == "report.docx"
    assert len(document.blocks) == 2
    assert isinstance(document.blocks[0], ParagraphBlock)
    assert document.blocks[0].text == "Keep this paragraph."
    assert isinstance(document.blocks[1], TableBlock)
    assert document.blocks[1].rows == [["Users", "42"]]


@pytest.mark.asyncio
async def test_generate_uses_explicit_filename_fallback(workspace_root: Path):
    generator = OfficeGenerator(data_path=workspace_root)
    generator.support = {OfficeType.WORD: True}
    generator._generate_word = AsyncMock()
    event = MagicMock()

    output_path = await generator.generate(
        event,
        OfficeType.WORD,
        "report",
        {"content": {"title": "Quarterly Report"}},
    )

    generator._generate_word.assert_awaited_once()
    assert output_path == workspace_root / "report.docx"
