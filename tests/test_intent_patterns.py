from astrbot_plugin_office_assistant.services.intent_patterns import (
    detect_explicit_file_tool,
    extract_document_id,
    has_document_intent,
    has_file_intent,
    has_pdf_conversion_intent,
    looks_like_document_followup,
    should_inject_document_tool_detail,
    should_use_active_document_summary,
)


def test_extract_document_id_returns_expected_value():
    assert extract_document_id('继续完善 document_id="doc-123" 的内容') == "doc-123"
    assert extract_document_id("hello") is None


def test_detect_explicit_file_tool_respects_negative_prefix():
    file_tools = ("read_file", "create_document", "create_office_file")

    assert (
        detect_explicit_file_tool(
            "请调用 create_document，title=季度复盘",
            file_tools,
        )
        == "create_document"
    )
    assert (
        detect_explicit_file_tool(
            "不要调用 create_document，先告诉我可用工具。",
            file_tools,
        )
        is None
    )


def test_pdf_conversion_intent_stays_distinct_from_generic_document_intent():
    text = "先导出为pdf再转word"

    assert has_pdf_conversion_intent(text) is True
    assert has_file_intent(text) is True
    assert has_document_intent(text) is False


def test_document_followup_requires_document_semantics_for_topic_switch():
    assert looks_like_document_followup("继续") is True
    assert looks_like_document_followup("再加一章关于销售的") is True
    assert looks_like_document_followup("继续讲个笑话") is False


def test_active_document_summary_only_applies_to_followup_requests():
    assert should_use_active_document_summary("继续完善上一版报告") is True
    assert should_use_active_document_summary("再加一章关于销售的") is True
    assert should_use_active_document_summary("请生成一份新的报告") is False


def test_document_tool_detail_hint_matches_detail_request_only():
    assert (
        should_inject_document_tool_detail(
            request_text="请生成一份 Word 报告，accent_color=112233",
            document_id=None,
        )
        is True
    )
    assert (
        should_inject_document_tool_detail(
            request_text='继续完善 document_id="doc-1" 的内容',
            document_id="doc-1",
        )
        is False
    )
