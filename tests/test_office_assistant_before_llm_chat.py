from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from astrbot_plugin_office_assistant.constants import DOC_COMMAND_TRIGGER_EVENT_KEY
from astrbot_plugin_office_assistant.internal_hooks import (
    NoticeBuildContext,
    ToolExposureContext,
)
from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.services.llm_request_policy import (
    LLMRequestPolicy,
)
from astrbot_plugin_office_assistant.services.request_hook_service import (
    RequestHookService,
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
    extras: dict[str, object] = {}
    event.message_obj = SimpleNamespace(type=message_type, message=[], self_id="bot-1")
    event.get_sender_id.return_value = sender_id
    event.get_platform_id.return_value = "platform-1"
    event.is_admin.return_value = is_admin
    event.unified_msg_origin = "session-1"
    event.message_str = ""
    event._buffer_reentry_count = 0
    event.set_extra.side_effect = lambda key, value: extras.__setitem__(key, value)
    event.get_extra.side_effect = lambda key=None, default=None: (
        dict(extras) if key is None else extras.get(key, default)
    )
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
        assert "document_style={brief, heading_color, title_align" in req.system_prompt
        assert "paragraph_space_after, list_space_after" in req.system_prompt
        assert "summary_card_defaults, table_defaults" in req.system_prompt
        assert "header_fill" in req.system_prompt
        assert "banded_rows" in req.system_prompt
        assert "first_column_bold" in req.system_prompt
        assert (
            "style={align, emphasis, font_scale, table_grid, cell_align}"
            in req.system_prompt
        )
        assert (
            "横向页面、节级页眉页脚、页码重置 MUST 使用独立的 `section_break` block"
            in req.system_prompt
        )
        assert (
            "不要把 `page_orientation`、`start_type`、`restart_page_numbering`"
            in req.system_prompt
        )
        assert (
            "`toc` 只使用 `title`、`levels`、`start_on_new_page`" in req.system_prompt
        )
        assert (
            "表格列标题使用 `headers`；不要给 `table` 传 `columns`" in req.system_prompt
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
@pytest.mark.parametrize(
    ("prompt", "expected_tool"),
    [
        (
            "调用 create_office_file，filename=report，content=hello，file_type=word",
            "create_office_file",
        ),
        (
            "reviewpr请求create_office_file，filename=report，content=hello，file_type=word",
            "create_office_file",
        ),
        (
            "please call `create_office_file` with filename=report, content=hello, file_type=word",
            "create_office_file",
        ),
    ],
)
async def test_before_llm_chat_restricts_file_tools_for_explicit_tool_call(
    prompt: str,
    expected_tool: str,
):
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt=prompt,
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
        assert expected_tool in tool_names
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
async def test_before_llm_chat_does_not_restrict_for_question_style_tool_mention():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt="请问 create_office_file 怎么用？先告诉我可用工具。",
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
@pytest.mark.parametrize(
    "prompt",
    [
        "不要调用 create_office_file，先告诉我可用工具。",
        "请问 create_office_file 怎么用？先告诉我可用工具。",
    ],
)
async def test_before_llm_chat_does_not_restrict_for_non_explicit_tool_mentions(
    prompt: str,
):
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = ProviderRequest(
            prompt=prompt,
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
    async def _custom_notice_hook(context: NoticeBuildContext):
        context.notices.append("\n[custom notice]")
        return context

    async def _custom_tool_hook(context: ToolExposureContext):
        if context.request.func_tool:
            context.request.func_tool.remove_tool("create_document")
        return context

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
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
async def test_llm_request_policy_does_not_build_request_hook_service_when_hooks_are_provided():
    with patch(
        "astrbot_plugin_office_assistant.services.llm_request_policy.RequestHookService"
    ) as request_hook_service_cls:
        policy = LLMRequestPolicy(
            document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
            require_at_in_group=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            is_bot_mentioned=lambda _event: True,
            notice_hooks=[],
            tool_exposure_hooks=[],
        )

    assert policy is not None
    request_hook_service_cls.assert_not_called()


@pytest.mark.asyncio
async def test_llm_request_policy_uses_injected_request_hook_service_for_default_hooks():
    request_hook_service = RequestHookService(
        auto_block_execution_tools=True,
        get_cached_upload_infos=lambda _event: [],
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
        allow_external_input_files=False,
    )
    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        request_hook_service=request_hook_service,
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1")
    req = ProviderRequest(
        prompt="hello",
        system_prompt="base",
        func_tool=ToolSet(
            [
                _tool("create_document"),
                _tool("existing_tool"),
                _tool("astrbot_execute_shell"),
            ]
        ),
    )

    await policy.apply(event, req)

    tool_names = set(req.func_tool.names())
    assert "create_document" in tool_names
    assert "existing_tool" in tool_names
    assert "astrbot_execute_shell" not in tool_names
    assert "文件工具使用指南" in req.system_prompt


def test_llm_request_policy_requires_hook_pairs():
    with pytest.raises(ValueError, match="must be provided together"):
        LLMRequestPolicy(
            document_toolset=SimpleNamespace(tools=[]),
            require_at_in_group=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            is_bot_mentioned=lambda _event: True,
            notice_hooks=[],
            tool_exposure_hooks=None,
        )


@pytest.mark.asyncio
async def test_runtime_bundle_does_not_expose_recent_text_cache():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        assert hasattr(plugin._runtime, "upload_session_service") is True
        assert hasattr(plugin._runtime, "recent_text_by_session") is False
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_plugin_terminate_allows_missing_runtime():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    await plugin.terminate()


@pytest.mark.asyncio
async def test_plugin_terminate_cleans_runtime_resources():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        message_buffer=MagicMock(set_complete_callback=MagicMock()),
        office_gen=MagicMock(cleanup=MagicMock()),
        pdf_converter=MagicMock(cleanup=MagicMock()),
        executor=MagicMock(shutdown=MagicMock()),
        temp_dir=MagicMock(cleanup=MagicMock()),
    )
    runtime = plugin._runtime

    await plugin.terminate()

    runtime.message_buffer.set_complete_callback.assert_called_once_with(None)
    runtime.office_gen.cleanup.assert_called_once_with()
    runtime.pdf_converter.cleanup.assert_called_once_with()
    runtime.executor.shutdown.assert_called_once_with(wait=False)
    runtime.temp_dir.cleanup.assert_called_once_with()
    assert plugin._runtime is None


@pytest.mark.asyncio
async def test_plugin_terminate_tolerates_temp_dir_permission_error():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        message_buffer=MagicMock(set_complete_callback=MagicMock()),
        office_gen=MagicMock(cleanup=MagicMock()),
        pdf_converter=MagicMock(cleanup=MagicMock()),
        executor=MagicMock(shutdown=MagicMock()),
        temp_dir=MagicMock(cleanup=MagicMock(side_effect=PermissionError("locked"))),
    )
    runtime = plugin._runtime

    await plugin.terminate()

    runtime.message_buffer.set_complete_callback.assert_called_once_with(None)
    runtime.office_gen.cleanup.assert_called_once_with()
    runtime.pdf_converter.cleanup.assert_called_once_with()
    runtime.executor.shutdown.assert_called_once_with(wait=False)
    runtime.temp_dir.cleanup.assert_called_once_with()
    assert plugin._runtime is None


@pytest.mark.asyncio
async def test_on_buffer_complete_ignores_released_runtime():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = None
    buf = MagicMock()

    with patch("astrbot_plugin_office_assistant.main.logger.warning") as warning:
        await plugin._on_buffer_complete(buf)

    warning.assert_called_once()


@pytest.mark.asyncio
async def test_handle_exported_document_tool_uses_bound_service_after_runtime_release():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    service = MagicMock()
    service.handle_exported_document_tool = AsyncMock(return_value="sent")
    plugin._post_export_hook_service = service
    plugin._runtime = None
    context = MagicMock()

    result = await plugin._handle_exported_document_tool(
        context,
        "/tmp/exported.docx",
    )

    service.handle_exported_document_tool.assert_awaited_once_with(
        context,
        "/tmp/exported.docx",
    )
    assert result == "sent"


@pytest.mark.asyncio
async def test_doc_list_command_stops_event_after_sending():
    context = MagicMock()
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.set_result = MagicMock()

        await plugin.doc_list(event)

        event.set_result.assert_called_once()
        result = event.set_result.call_args.args[0]
        assert result.get_plain_text() == "当前没有可处理的上传文件。"
        assert result.is_stopped() is True
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_doc_use_command_stops_event_after_requeue():
    context = MagicMock()
    plugin = FileOperationPlugin(context=context, config=_build_config())
    try:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.stop_event = MagicMock()
        plugin._runtime.command_service.doc_use = AsyncMock(return_value=None)

        await plugin.doc_use(event, "f1 根据这份文件整理")

        plugin._runtime.command_service.doc_use.assert_awaited_once_with(
            event,
            "f1 根据这份文件整理",
        )
        event.stop_event.assert_called_once_with()
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_doc_clear_command_sets_stopped_result():
    context = MagicMock()
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.set_result = MagicMock()

        await plugin.doc_clear(event, "")

        event.set_result.assert_called_once()
        result = event.set_result.call_args.args[0]
        assert result.get_plain_text() == "❌ 当前没有可处理的上传文件。"
        assert result.is_stopped() is True
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_buffered_upload_without_prompt_only_caches_upload_infos():
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

        event_queue.put.assert_not_awaited()
        upload_infos = plugin._runtime.upload_session_service.list_session_upload_infos(
            event
        )
        assert len(upload_infos) == 1
        assert upload_infos[0]["original_name"] == "report.docx"
        assert upload_infos[0]["stored_name"] == "report_1.docx"
        assert upload_infos[0]["file_id"] == "f1"
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
        upload_session_ttl_seconds=300,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        upload_session_cleanup_interval_seconds=30,
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

    queued_event = event_queue.put.await_args.args[0]
    assert queued_event is not event
    assert isinstance(queued_event.message_obj.message[0], Comp.Plain)
    prompt_text = queued_event.message_obj.message[0].text
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
    assert not queued_event.message_str.startswith("/")
    assert queued_event.message_str.endswith(prompt_text.strip())
    event_queue.put.assert_awaited_once()


@pytest.mark.asyncio
async def test_before_llm_chat_exposes_file_tools_for_buffered_group_upload_when_mentioned():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        raw_message = SimpleNamespace(mentions=[SimpleNamespace(id="bot-1")])
        event.message_obj.raw_message = raw_message
        event.is_mentioned.side_effect = lambda: hasattr(
            event.message_obj.raw_message, "mentions"
        ) and any(
            str(mention.id) == str(event.message_obj.self_id)
            for mention in event.message_obj.raw_message.mentions
        )
        upload = Comp.File(name="source.docx", file=str(source_path))
        buf = BufferedMessage(
            event=event,
            files=[upload],
            texts=["请根据上传文档整理成正式汇报"],
        )

        async def _fake_extract_upload_source(_component):
            return source_path, "source.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source_1.docx")

        await plugin._on_buffer_complete(buf)
        queued_event = event_queue.put.await_args.args[0]

        req = ProviderRequest(
            prompt=queued_event.message_str,
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(queued_event, req)

        tool_names = set(req.func_tool.names())
        assert queued_event.message_obj.raw_message is raw_message
        assert queued_event.is_mentioned() is True
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" not in req.system_prompt
        assert "工作区文件名：source_1.docx" in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_hides_file_tools_for_buffered_group_upload_when_not_mentioned():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        raw_message = SimpleNamespace(mentions=[])
        event.message_obj.raw_message = raw_message
        event.is_mentioned.side_effect = lambda: hasattr(
            event.message_obj.raw_message, "mentions"
        ) and any(
            str(mention.id) == str(event.message_obj.self_id)
            for mention in event.message_obj.raw_message.mentions
        )
        upload = Comp.File(name="source.docx", file=str(source_path))
        buf = BufferedMessage(
            event=event,
            files=[upload],
            texts=["请根据上传文档整理成正式汇报"],
        )

        async def _fake_extract_upload_source(_component):
            return source_path, "source.docx"

        plugin._extract_upload_source = _fake_extract_upload_source
        plugin._store_uploaded_file = lambda *_args, **_kwargs: Path("source_1.docx")

        await plugin._on_buffer_complete(buf)

        req = ProviderRequest(
            prompt=event.message_str,
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert event.message_obj.raw_message is raw_message
        assert event.is_mentioned() is False
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "export_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt
    finally:
        await plugin.terminate()


@pytest.mark.asyncio
async def test_before_llm_chat_exposes_file_tools_for_group_doc_command_without_mention():
    context = MagicMock()
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    plugin = FileOperationPlugin(context=context, config=config)
    try:
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        event.is_mentioned.return_value = False
        event._buffered = True
        event.set_extra(DOC_COMMAND_TRIGGER_EVENT_KEY, True)
        event.message_str = (
            "[System Notice] 用户上传了 1 个文件\n\n"
            "[文件信息]\n"
            "- 原始文件名: B.xlsx (类型: .xlsx)\n"
            "  工作区文件名: B_1.xlsx\n\n"
            "[用户指令]\n"
            "根据这份文件整理成正式汇报\n\n"
            "[处理建议]\n"
            "1. 优先围绕这些上传文件完成用户请求。\n"
        )
        req = ProviderRequest(
            prompt=event.message_str,
            system_prompt="base",
            func_tool=ToolSet([_tool("existing_tool")]),
        )

        await plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" not in req.system_prompt
    finally:
        await plugin.terminate()


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
        event.is_mentioned.return_value = False
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
