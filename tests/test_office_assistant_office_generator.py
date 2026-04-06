from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest
from astrbot_plugin_office_assistant.domain.document.render_backends import (
    DocumentRenderBackendConfig,
    RenderResult,
)
from astrbot_plugin_office_assistant.document_core.models.blocks import (
    HeadingBlock,
    GroupBlock,
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


def test_build_word_document_model_reuses_session_store_normalization(
    workspace_root: Path,
):
    generator = OfficeGenerator(data_path=workspace_root)
    file_path = workspace_root / "report.docx"

    document = generator._build_word_document_model(
        file_path,
        {
            "metadata": {"title": "Quarterly Report"},
            "blocks": [
                {"type": "heading", "text": "一、经营总览", "level": 1},
                {
                    "type": "table",
                    "headers": ["日期", "时间", "内容"],
                    "rows": [
                        [{"text": "第一天", "row_span": 2}, "09:00", "课程 A"],
                        ["13:00", "课程 B"],
                    ],
                },
            ],
        },
    )

    assert len(document.blocks) == 2
    assert isinstance(document.blocks[0], HeadingBlock)
    assert document.blocks[0].text == "一、经营总览"
    assert isinstance(document.blocks[1], TableBlock)
    assert document.blocks[1].caption in (None, "")
    assert document.blocks[1].rows[0][0].row_span == 2


def test_build_word_document_model_expands_summary_card_blocks(
    workspace_root: Path,
):
    generator = OfficeGenerator(data_path=workspace_root)
    file_path = workspace_root / "report.docx"

    document = generator._build_word_document_model(
        file_path,
        {
            "metadata": {"title": "Quarterly Report"},
            "blocks": [
                {
                    "type": "summary_card",
                    "title": "Highlights",
                    "items": ["Stable revenue", "Lower churn"],
                    "variant": "conclusion",
                }
            ],
        },
    )

    assert len(document.blocks) == 1
    assert isinstance(document.blocks[0], GroupBlock)
    assert document.blocks[0].blocks[0].text == "Highlights"
    assert document.blocks[0].blocks[1].items == [
        "Stable revenue",
        "Lower churn",
    ]


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


def test_generate_word_sync_uses_document_render_backends(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    config = DocumentRenderBackendConfig(preferred_backend="node")
    generator = OfficeGenerator(
        data_path=workspace_root,
        render_backend_config=config,
    )
    file_path = workspace_root / "report.docx"
    captured: dict[str, object] = {}
    backends = [MagicMock(name="node-backend")]

    def _fake_build_backends(document_format, resolved_config):
        captured["document_format"] = document_format
        captured["config"] = resolved_config
        return backends

    def _fake_render(document, output_path, resolved_backends):
        captured["document"] = document
        captured["output_path"] = output_path
        captured["backends"] = resolved_backends
        return RenderResult(backend_name="node", output_path=output_path)

    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.office_generator.build_document_render_backends",
        _fake_build_backends,
    )
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.office_generator.render_document_with_backends",
        _fake_render,
    )

    generator._generate_word_sync(
        file_path,
        {"metadata": {"title": "Quarterly Report"}, "blocks": []},
    )

    assert captured["document_format"] == "word"
    assert captured["config"] is config
    assert captured["output_path"] == file_path
    assert captured["backends"] == backends
    assert captured["document"].metadata.preferred_filename == "report.docx"


def test_generate_word_sync_propagates_backend_failure(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    generator = OfficeGenerator(
        data_path=workspace_root,
        render_backend_config=DocumentRenderBackendConfig(preferred_backend="node"),
    )
    file_path = workspace_root / "report.docx"

    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.office_generator.build_document_render_backends",
        lambda *_args, **_kwargs: [MagicMock(name="node-backend")],
    )
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.office_generator.render_document_with_backends",
        MagicMock(side_effect=RuntimeError("node renderer unavailable")),
    )

    with pytest.raises(RuntimeError, match="node renderer unavailable"):
        generator._generate_word_sync(
            file_path,
            {"metadata": {"title": "Quarterly Report"}, "blocks": []},
        )
