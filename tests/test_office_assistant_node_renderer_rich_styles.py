import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _docx_test_helpers import (
    _business_report_metadata,
    _find_paragraph,
    _paragraph_bottom_border_color,
    _paragraph_bottom_border_size,
    _render_structured_payload_with_node,
)


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _paragraph_xml_by_text(document_xml: str, paragraph_text: str) -> ET.Element:
    root = ET.fromstring(document_xml)
    for paragraph in root.findall(".//w:p", NS):
        text = "".join(
            node.text or ""
            for node in paragraph.findall(".//w:t", NS)
            if node.text is not None
        )
        if text == paragraph_text:
            return paragraph
    raise AssertionError(f"Paragraph not found: {paragraph_text}")


def _run_props_by_text(paragraph: ET.Element, run_text: str) -> ET.Element:
    for run in paragraph.findall(".//w:r", NS):
        text = "".join(
            node.text or "" for node in run.findall(".//w:t", NS) if node.text is not None
        )
        if text == run_text:
            return run
    raise AssertionError(f"Run not found: {run_text}")


def test_node_renderer_renders_rich_paragraph_runs_and_border(workspace_root: Path):
    loaded_doc, output_path = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-rich-paragraph",
        {
            "document_id": "rich-paragraph",
            "metadata": _business_report_metadata(title="富文本段落"),
            "blocks": [
                {
                    "type": "paragraph",
                    "border": {
                        "bottom": {
                            "style": "single",
                            "color": "1F4E79",
                            "width_pt": 1.0,
                        }
                    },
                    "runs": [
                        {
                            "text": "普通",
                            "font_name": "SimSun",
                            "color": "0F4C81",
                        },
                        {
                            "text": "强调",
                            "font_name": "Source Han Sans SC",
                            "font_scale": 1.25,
                            "italic": True,
                            "strikethrough": True,
                            "color": "C2410C",
                        },
                        {
                            "text": "链接",
                            "url": "https://example.com/docs",
                            "font_name": "Microsoft YaHei",
                            "underline": True,
                        },
                    ],
                }
            ],
        },
    )

    paragraph = _find_paragraph(loaded_doc, "普通强调链接")
    assert _paragraph_bottom_border_color(paragraph) == "1F4E79"
    assert _paragraph_bottom_border_size(paragraph) == "8"

    with zipfile.ZipFile(output_path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        rels_xml = archive.read("word/_rels/document.xml.rels").decode("utf-8")

    paragraph_xml = _paragraph_xml_by_text(document_xml, "普通强调链接")
    first_run = _run_props_by_text(paragraph_xml, "普通")
    second_run = _run_props_by_text(paragraph_xml, "强调")
    third_run = _run_props_by_text(paragraph_xml, "链接")

    assert "https://example.com/docs" in rels_xml
    assert paragraph_xml.find(".//w:hyperlink", NS) is not None

    first_fonts = first_run.find("w:rPr/w:rFonts", NS)
    second_fonts = second_run.find("w:rPr/w:rFonts", NS)
    third_fonts = third_run.find("w:rPr/w:rFonts", NS)

    assert first_fonts is not None
    assert second_fonts is not None
    assert third_fonts is not None
    assert first_fonts.get(f"{{{NS['w']}}}ascii") == "SimSun"
    assert second_fonts.get(f"{{{NS['w']}}}ascii") == "Source Han Sans SC"
    assert third_fonts.get(f"{{{NS['w']}}}ascii") == "Microsoft YaHei"
    assert second_run.find("w:rPr/w:i", NS) is not None
    assert second_run.find("w:rPr/w:strike", NS) is not None
    assert second_run.find("w:rPr/w:color", NS).get(f"{{{NS['w']}}}val") == "C2410C"
    assert third_run.find("w:rPr/w:u", NS) is not None
    assert "w:hyperlink" in document_xml


def test_node_renderer_keeps_text_only_paragraph_and_list_intact(workspace_root: Path):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-text-only-paragraph-list",
        {
            "document_id": "text-only-paragraph-list",
            "metadata": _business_report_metadata(title="普通文本"),
            "blocks": [
                {
                    "type": "paragraph",
                    "text": "纯文本段落",
                },
                {
                    "type": "list",
                    "items": ["第一项", "第二项"],
                },
            ],
        },
    )

    paragraph = _find_paragraph(loaded_doc, "纯文本段落")
    list_item_one = _find_paragraph(loaded_doc, "• 第一项")
    list_item_two = _find_paragraph(loaded_doc, "• 第二项")

    assert paragraph.text == "纯文本段落"
    assert list_item_one.text == "• 第一项"
    assert list_item_two.text == "• 第二项"
    assert _paragraph_bottom_border_color(paragraph) is None
    assert _paragraph_bottom_border_size(paragraph) is None
    assert _paragraph_bottom_border_color(list_item_one) is None
    assert _paragraph_bottom_border_color(list_item_two) is None


def test_node_renderer_honors_empty_paragraph_border_side_defaults(
    workspace_root: Path,
):
    loaded_doc, output_path = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-default-paragraph-border-side",
        {
            "document_id": "default-paragraph-border-side",
            "metadata": _business_report_metadata(title="默认段落边框标题"),
            "blocks": [
                {
                    "type": "paragraph",
                    "text": "默认段落边框正文",
                    "border": {
                        "bottom": {},
                    },
                }
            ],
        },
    )

    paragraph = _find_paragraph(loaded_doc, "默认段落边框正文")
    assert _paragraph_bottom_border_size(paragraph) == "4"
    assert _paragraph_bottom_border_color(paragraph) is None

    with zipfile.ZipFile(output_path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")

    paragraph_xml = _paragraph_xml_by_text(document_xml, "默认段落边框正文")
    assert paragraph_xml.find("./w:pPr/w:pBdr/w:top", NS) is None
    assert paragraph_xml.find("./w:pPr/w:pBdr/w:left", NS) is None
    assert paragraph_xml.find("./w:pPr/w:pBdr/w:right", NS) is None
