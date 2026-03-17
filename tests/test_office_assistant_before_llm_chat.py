from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

import astrbot.api.message_components as Comp
from astrbot.core.agent.tool import FunctionTool, ToolSet
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest
from data.plugins.plugin_upload_astrbot_plugin_office_assistant.main import (
    FileOperationPlugin,
)


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
    event.is_admin.return_value = is_admin
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
        assert "For Word documents" in req.system_prompt
        assert "executive_brief" in req.system_prompt
        assert "accent_color=RRGGBB" in req.system_prompt
        assert (
            "style={align, emphasis, font_scale, table_grid, cell_align}"
            in req.system_prompt
        )
        assert (
            "Prefer one `add_blocks` call per section or logical chunk"
            in req.system_prompt
        )
        assert (
            "Continue calling document tools until `export_document` succeeds"
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
        assert "File/Office/PDF actions are unavailable" in req.system_prompt
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
        assert "File/Office/PDF actions are unavailable" in req.system_prompt
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
            "If the user request depends on uploaded readable files, call `read_file` before `create_document` or `create_office_file`."
            in req.system_prompt
        )
        assert (
            "If the user request depends on this uploaded file, call `read_file` before `create_document` or `create_office_file`."
            in req.system_prompt
        )
        assert (
            "Do not create a new document before reading the uploaded source at least once."
            in req.system_prompt
        )
    finally:
        await plugin.terminate()
