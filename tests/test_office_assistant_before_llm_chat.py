from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest
from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.internal_hooks import (
    NoticeBuildContext,
    ToolExposureContext,
)
from astrbot_plugin_office_assistant.services.llm_request_policy import (
    LLMRequestPolicy,
)
from astrbot_plugin_office_assistant.services.upload_session_service import (
    UploadSessionService,
)

import astrbot.api.message_components as Comp
from astrbot.core.agent.tool import FunctionTool, ToolSet
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest


def _build_config() -> dict:
    return {
        "file_settings": {
            "auto_delete_files": True,
            "max_file_size_mb": 20,
            "message_buffer_seconds": 4,
        },
        "trigger_settings": {
            "reply_to_user": True,
            "require_at_in_group": True,
            "enable_features_in_group": False,
            "auto_block_execution_tools": True,
        },
        "preview_settings": {
            "enable": False,
            "dpi": 150,
        },
        "path_settings": {
            "allow_external_input_files": False,
        },
        "permission_settings": {
            "whitelist_users": ["user-1"],
        },
    }


def _build_event(
    *,
    message_type=MessageType.FRIEND_MESSAGE,
    sender_id: str = "user-1",
    is_admin: bool = False,
):
    event = MagicMock()
    event.message_obj = SimpleNamespace(type=message_type, message=[], self_id="bot-1")
    event.get_sender_id.return_value = sender_id
    event.get_platform_id.return_value = "platform-1"
    event.is_admin.return_value = is_admin
    event.unified_msg_origin = "session-1"
    event.message_str = ""
    event._buffer_reentry_count = 0
    return event


def _tool(name: str) -> FunctionTool:
    return FunctionTool(
        name=name,
        description=f"tool {name}",
        parameters={"type": "object", "properties": {}},
        handler=None,
    )


@pytest.mark.asyncio
async def test_before_llm_chat_injects_document_tools_per_request():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = ProviderRequest(
            prompt="hello",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("astrbot_execute_shell"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert {
            "create_document",
            "add_blocks",
            "finalize_document",
            "export_document",
        }.issubset(tool_names)
        assert "generate_complex_word_document" not in tool_names
        assert "文件工具使用指南" in req.system_prompt
        assert "executive_brief" in req.system_prompt
        assert "accent_color=RRGGBB" in req.system_prompt
        assert (
            "style={align, emphasis, font_scale, table_grid, cell_align}"
            in req.system_prompt
        )
        assert "按章节或逻辑块分批调用 `add_blocks`" in req.system_prompt
        assert "MUST 持续调用直到 `export_document` 成功" in req.system_prompt
        assert "NEVER 调用网络搜索" in req.system_prompt
        assert (
            "如果用户显式指定了某个工具名和参数，MUST 先按该工具调用"
            in req.system_prompt
        )
        assert "所有面向用户的回复和过渡说明 MUST 使用中文" in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_removes_file_tools_without_permission():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-2"
        )
        req = ProviderRequest(
            prompt="hello",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("export_document"),
                    _tool("existing_tool"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" not in tool_names
        assert "create_document" not in tool_names
        assert "export_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "generate_complex_word_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "`astrbot_execute_python`" in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_warns_when_group_feature_disabled():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE, sender_id="user-1")
        req = ProviderRequest(
            prompt="请生成一份 Word 报告",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("astrbot_execute_python"),
                    _tool("existing_tool"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_python" in tool_names
        assert "read_file" not in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "generate_complex_word_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "`astrbot_execute_python`" in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_requires_read_before_document_tools_for_uploaded_files():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]

        async def _fake_extract_upload_source(_component):
            return source_path, "source.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source_1.docx")

        req = ProviderRequest(
            prompt="根据上传文档整理成正式汇报",
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(event, req)

        assert "MUST 先调用 `read_file` 读取内容，再创建文档" in req.system_prompt
        assert "MUST 先调用 `read_file` 读取此文件" in req.system_prompt
        assert "原始文件名：source.docx" in req.system_prompt
        assert "工作区文件名：source_1.docx" in req.system_prompt
        assert "必须使用工作区文件名 `source_1.docx`" in req.system_prompt
        assert "NEVER 创建新文档" in req.system_prompt
        assert "NEVER 调用网络搜索" in req.system_prompt
        assert "MUST 使用中文" in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_restricts_file_tools_for_explicit_tool_call():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt="调用 create_office_file，filename=report，content=hello，file_type=word",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("create_office_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                    _tool("finalize_document"),
                    _tool("export_document"),
                    _tool("read_file"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "create_office_file" in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "finalize_document" not in tool_names
        assert "export_document" not in tool_names
        assert "read_file" not in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_restrict_when_prompt_mentions_multiple_tools():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt="先调用 read_file 再调用 create_document 处理文件",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                    _tool("export_document"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_restrict_for_negated_tool_mention():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt="不要调用 create_office_file，先告诉我可用工具。",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("create_office_file"),
                    _tool("create_document"),
                    _tool("read_file"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "create_office_file" in tool_names
        assert "create_document" in tool_names
        assert "read_file" in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_treat_system_notice_as_explicit_tool_call():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_str = "[System Notice] 用户上传了文件，请先调用 `read_file` 读取内容，再继续处理。"
        req = ProviderRequest(
            prompt="",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                    _tool("export_document"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_falls_back_to_raw_prompt_when_system_notice_block_is_missing():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt="[System Notice] 这是用户自己输入的字面量。调用 read_file，filename=report.txt",
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_uses_buffered_user_instruction_for_explicit_tool_detection():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event._buffered = True
        req = ProviderRequest(
            prompt=(
                "[System Notice] 用户上传了 1 个文件\n\n"
                "[文件信息]\n"
                "- 原始文件名: source.docx (类型: .docx)\n"
                "  工作区文件名: source_1.docx\n\n"
                "[用户指令]\n"
                "请根据我刚上传的文档整理成正式汇报，标题叫《项目进展汇总》，最后导出成 Word 并发给我。\n\n"
                "[处理建议]\n"
                "1. 先调用 `read_file` 读取文件。\n"
            ),
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                    _tool("finalize_document"),
                    _tool("export_document"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "finalize_document" in tool_names
        assert "export_document" in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_can_still_restrict_tool_from_buffered_user_instruction():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event._buffered = True
        req = ProviderRequest(
            prompt=(
                "[System Notice] 用户上传了 1 个文件\n\n"
                "[文件信息]\n"
                "- 原始文件名: table.csv (类型: .csv)\n"
                "  工作区文件名: table_1.csv\n\n"
                "[用户指令]\n"
                "调用 read_file，filename=table_1.csv\n\n"
                "[处理建议]\n"
                "1. 先调用 `read_file` 读取文件。\n"
            ),
            system_prompt="base",
            func_tool=ToolSet(
                [
                    _tool("existing_tool"),
                    _tool("read_file"),
                    _tool("create_document"),
                    _tool("add_blocks"),
                ]
            ),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_llm_request_policy_runs_internal_notice_and_tool_hooks():
    async def _fake_extract_upload_source(_component):
        return None, ""

    async def _custom_notice_hook(context: NoticeBuildContext):
        context.notices.append("\n[custom notice]")
        return context

    async def _custom_tool_hook(context: ToolExposureContext):
        if context.request.func_tool:
            context.request.func_tool.remove_tool("create_document")
        return context

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        auto_block_execution_tools=False,
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        get_cached_upload_infos=lambda _event: [],
        extract_upload_source=_fake_extract_upload_source,
        store_uploaded_file=lambda _src, _name: Path("ignored.txt"),
        allow_external_input_files=False,
        notice_hooks=[_custom_notice_hook],
        tool_exposure_hooks=[_custom_tool_hook],
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1")
    req = ProviderRequest(
        prompt="hello",
        system_prompt="base",
        func_tool=ToolSet([_tool("create_document"), _tool("existing_tool")]),
    )

    await policy.apply(event, req)

    assert "[custom notice]" in req.system_prompt
    assert "create_document" not in set(req.func_tool.names())
    assert "existing_tool" in set(req.func_tool.names())


@pytest.mark.asyncio
async def test_remember_recent_text_skips_system_notice_messages():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event()
        session_key = plugin._get_attachment_session_key(event)

        event.message_str = "[System Notice] internal guidance"
        plugin._remember_recent_text(event)
        assert session_key not in plugin._recent_text_by_session

        event.message_str = "整理成正式汇报"
        plugin._remember_recent_text(event)
        assert plugin._recent_text_by_session[session_key][0] == "整理成正式汇报"
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_buffered_upload_without_prompt_requires_read_file_first():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        source_path = Path(__file__).resolve()
        event = _build_event()
        upload = Comp.File(name="report.docx", file="report.docx")
        buf = BufferedMessage(event=event, files=[upload], texts=[])

        async def _fake_extract_upload_source(_component):
            return source_path, "report.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("report_1.docx")

        await plugin._on_buffer_complete(buf)

        assert isinstance(event.message_obj.message[0], Comp.Plain)
        prompt_text = event.message_obj.message[0].text
        assert "用户上传了可读取文件" in prompt_text
        assert "工作区文件名: report_1.docx" in prompt_text
        assert "不要自行猜测文件名，也不要列目录或调用 shell" in prompt_text
        assert "若使用相对路径，请使用上面的工作区文件名" in prompt_text
        assert "如果要读取文件" in prompt_text
        assert "NEVER 创建新文档" not in prompt_text
        assert event.message_str == prompt_text.strip()
        event_queue.put.assert_awaited_once_with(event)
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_buffered_upload_with_prompt_uses_structured_notice_without_hard_constraints():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(
            return_value=(Path(__file__).resolve(), "report.docx")
        ),
        store_uploaded_file=MagicMock(return_value=Path("report_1.docx")),
        allow_external_input_files=True,
    )
    event = _build_event()
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(
        event=event,
        files=[upload],
        texts=["请根据上传文档整理成正式汇报"],
    )

    await service.on_buffer_complete(buf)

    assert isinstance(event.message_obj.message[0], Comp.Plain)
    prompt_text = event.message_obj.message[0].text
    assert "[文件信息]" in prompt_text
    assert "[用户指令]" in prompt_text
    assert "请根据上传文档整理成正式汇报" in prompt_text
    assert "[处理建议]" in prompt_text
    assert "优先围绕这些上传文件完成用户请求" in prompt_text
    assert "工作区文件名: report_1.docx" in prompt_text
    assert "外部绝对路径:" in prompt_text
    assert "先调用 `read_file` 读取文件" in prompt_text
    assert "不要自行猜测文件名，也不要列目录或调用 shell" in prompt_text
    assert "NEVER 创建新文档" not in prompt_text
    assert event.message_str == prompt_text.strip()
    event_queue.put.assert_awaited_once_with(event)


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_inject_upload_notice_when_file_tools_hidden():
    context = MagicMock()
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]

        async def _fake_extract_upload_source(_component):
            return source_path, "source.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source_1.docx")

        req = ProviderRequest(
            prompt="根据上传文档整理成正式汇报",
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(event, req)

        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt
        assert "MUST 先调用 `read_file` 读取此文件" not in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_skips_upload_notices_when_func_tool_missing():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]

        async def _fake_extract_upload_source(_component):
            return source_path, "source.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source_1.docx")

        req = ProviderRequest(
            prompt="根据上传文档整理成正式汇报",
            system_prompt="base",
            func_tool=None,
        )

        await plugin.before_llm_chat(event, req)

        assert "文件工具使用指南" not in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt
        assert "MUST 先调用 `read_file` 读取此文件" not in req.system_prompt
    finally:
        await plugin.terminate()
