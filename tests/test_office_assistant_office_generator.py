import shutil
from pathlib import Path
from uuid import uuid4

import pytest

from astrbot_plugin_office_assistant.document_core.models.blocks import (
    ParagraphBlock,
    TableBlock,
)
from astrbot_plugin_office_assistant.office_generator import (
    OfficeGenerator,
)


@pytest.fixture
def workspace_root() -> Path:
    workspace_base = Path(__file__).resolve().parent / ".tmp_office_generator"
    workspace_base.mkdir(parents=True, exist_ok=True)
    workspace_dir = workspace_base / f"workspace-root-{uuid4().hex}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    try:
        yield workspace_dir
    finally:
        shutil.rmtree(workspace_dir, ignore_errors=True)


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
