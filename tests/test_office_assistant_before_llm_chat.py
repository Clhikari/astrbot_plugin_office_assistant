from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest
from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage

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
        assert "处理 Word 文档时，请使用有状态文档工具链" in req.system_prompt
        assert "executive_brief" in req.system_prompt
        assert "accent_color=RRGGBB" in req.system_prompt
        assert (
            "style={align, emphasis, font_scale, table_grid, cell_align}"
            in req.system_prompt
        )
        assert "最好按章节或逻辑块调用 `add_blocks`" in req.system_prompt
        assert "继续调用文档工具，直到 `export_document` 成功" in req.system_prompt
        assert "不要调用网络搜索" in req.system_prompt
        assert (
            "如果在调用工具前需要先给用户一句过渡说明，也请使用中文"
            in req.system_prompt
        )
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
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source.docx")

        req = ProviderRequest(
            prompt="根据上传文档整理成正式汇报",
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(event, req)

        assert (
            "如果用户请求依赖上传的可读文件，先调用 `read_file`，再调用 `create_document`"
            in req.system_prompt
        )
        assert (
            "如果用户请求依赖这个上传文件，先调用 `read_file`，再调用 `create_document`"
            in req.system_prompt
        )
        assert "在至少读取一次上传源文件之前，不要先创建新文档。" in req.system_prompt
        assert "不要调用网络搜索" in req.system_prompt
        assert "如果要先给用户一句过渡说明，也请使用中文" in req.system_prompt
    finally:
        await plugin.terminate()


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
        event = _build_event()
        upload = Comp.File(name="report.docx", file="report.docx")
        buf = BufferedMessage(event=event, files=[upload], texts=[])

        await plugin._on_buffer_complete(buf)

        assert isinstance(event.message_obj.message[0], Comp.Plain)
        prompt_text = event.message_obj.message[0].text
        assert "请现在调用 `read_file`。" in prompt_text
        assert "读取上传源文件前，不要先创建新文档。" in prompt_text
        assert "目前用户意图还不够明确，读取后再用中文追问。" in prompt_text
        assert event.message_str == prompt_text.strip()
        event_queue.put.assert_awaited_once_with(event)
    finally:
        await plugin.terminate()
