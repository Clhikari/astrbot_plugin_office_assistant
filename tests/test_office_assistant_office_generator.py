import tempfile
from pathlib import Path

from astrbot_plugin_office_assistant.document_core.models.blocks import (
    ParagraphBlock,
    TableBlock,
)
from astrbot_plugin_office_assistant.office_generator import (
    OfficeGenerator,
)


def test_build_word_document_model_skips_invalid_blocks():
    with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
        temp_path = Path(temp_dir)
        generator = OfficeGenerator(data_path=temp_path)
        file_path = temp_path / "report.docx"

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
