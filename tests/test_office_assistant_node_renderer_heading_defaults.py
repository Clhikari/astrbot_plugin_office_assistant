from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _docx_test_helpers import (
    _business_report_metadata,
    _find_paragraph,
    _paragraph_after,
    _paragraph_bottom_border_color,
    _paragraph_bottom_border_size,
    _render_structured_payload_with_node,
)


def test_node_renderer_does_not_add_heading_divider_by_default(
    workspace_root: Path,
):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-heading-divider-default-off",
        {
            "document_id": "heading-divider-default-off",
            "metadata": _business_report_metadata(title="标题默认无线"),
            "blocks": [
                {
                    "type": "heading",
                    "text": "经营概览",
                    "level": 1,
                },
                {
                    "type": "paragraph",
                    "text": "正文段落",
                },
            ],
        },
    )

    heading = _find_paragraph(loaded_doc, "经营概览")
    body = _find_paragraph(loaded_doc, "正文段落")

    assert _paragraph_bottom_border_color(heading) is None
    assert _paragraph_bottom_border_size(heading) is None
    assert _paragraph_after(loaded_doc, heading)._p is body._p


def test_node_renderer_keeps_explicit_heading_divider(
    workspace_root: Path,
):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-heading-divider-explicit-on",
        {
            "document_id": "heading-divider-explicit-on",
            "metadata": _business_report_metadata(title="标题显式分割线"),
            "blocks": [
                {
                    "type": "heading",
                    "text": "经营概览",
                    "level": 1,
                    "bottom_border": True,
                    "bottom_border_color": "1F4E79",
                    "bottom_border_size_pt": 1.5,
                },
                {
                    "type": "paragraph",
                    "text": "正文段落",
                },
            ],
        },
    )

    heading = _find_paragraph(loaded_doc, "经营概览")
    divider = _paragraph_after(loaded_doc, heading)
    body = _find_paragraph(loaded_doc, "正文段落")

    assert _paragraph_bottom_border_color(heading) is None
    assert _paragraph_bottom_border_color(divider) == "1F4E79"
    assert _paragraph_bottom_border_size(divider) == "12"
    assert _paragraph_after(loaded_doc, divider, offset=2)._p is body._p
