import shutil
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch
from uuid import uuid4

import pytest
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.services import (
    AccessPolicyService,
    CommandService,
    DeliveryService,
    ErrorHookService,
    FileToolService,
    IncomingMessageService,
    UploadSessionService,
    WorkspaceService,
    build_plugin_runtime,
)

import astrbot.api.message_components as Comp
from astrbot.core.platform.message_type import MessageType


def _build_event(
    *,
    sender_id: str = "user-1",
    message_type=MessageType.FRIEND_MESSAGE,
):
    event = MagicMock()
    event.message_obj = SimpleNamespace(type=message_type, message=[], self_id="bot-1")
    event.get_sender_id.return_value = sender_id
    event.get_platform_id.return_value = "platform-1"
    event.unified_msg_origin = "session-1"
    event.message_str = ""
    event.is_admin.return_value = False
    event._buffer_reentry_count = 0
    event._buffered = False
    return event


def _make_workspace(name: str) -> Path:
    workspace_base = Path(__file__).resolve().parent / ".tmp_services"
    workspace_base.mkdir(parents=True, exist_ok=True)
    workspace_dir = workspace_base / f"{name}-{uuid4().hex}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    return workspace_dir


def test_access_policy_service_handles_whitelist_and_group_flags():
    service = AccessPolicyService(
        whitelist_users=["user-1"],
        enable_features_in_group=False,
    )
    friend_event = _build_event(message_type=MessageType.FRIEND_MESSAGE)
    group_event = _build_event(message_type=MessageType.GROUP_MESSAGE)

    assert service.check_permission(friend_event) is True
    assert service.is_group_feature_enabled(friend_event) is True
    assert service.is_group_feature_enabled(group_event) is False


def test_access_policy_service_detects_bot_mention():
    service = AccessPolicyService(
        whitelist_users=["user-1"],
        enable_features_in_group=True,
    )
    event = _build_event()
    event.message_obj.message = [Comp.At(qq="bot-1")]

    assert service.is_bot_mentioned(event) is True


def test_build_plugin_runtime_returns_temp_workspace_and_services():
    context = MagicMock()
    config = {
        "file_settings": {
            "auto_delete_files": True,
            "max_file_size_mb": 8,
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
    runtime = build_plugin_runtime(
        context=context,
        config=config,
        handle_exported_document_tool=AsyncMock(),
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
    )

    try:
        assert runtime.settings.auto_delete is True
        assert runtime.plugin_data_path.exists()
        assert runtime.temp_dir is not None
        assert runtime.workspace_service.plugin_data_path == runtime.plugin_data_path
        assert runtime.message_buffer is not None
        assert runtime.incoming_message_service is not None
    finally:
        runtime.executor.shutdown(wait=False)
        runtime.office_gen.cleanup()
        runtime.pdf_converter.cleanup()
        if runtime.temp_dir is not None:
            try:
                runtime.temp_dir.cleanup()
            except PermissionError:
                pass


def test_build_plugin_runtime_uses_persistent_workspace_when_auto_delete_disabled(
    monkeypatch: pytest.MonkeyPatch,
):
    data_root = _make_workspace("runtime-builder-data-root")
    context = MagicMock()
    config = {
        "file_settings": {
            "auto_delete_files": False,
            "max_file_size_mb": 16,
            "message_buffer_seconds": 7,
        },
        "trigger_settings": {
            "reply_to_user": False,
            "require_at_in_group": False,
            "enable_features_in_group": True,
            "auto_block_execution_tools": False,
        },
        "preview_settings": {
            "enable": True,
            "dpi": 180,
        },
        "path_settings": {
            "allow_external_input_files": True,
        },
        "permission_settings": {
            "whitelist_users": ["user-2"],
        },
        "feature_settings": {
            "enable_office_files": True,
        },
    }
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.services.runtime_builder.StarTools.get_data_dir",
        lambda: data_root,
    )

    runtime = build_plugin_runtime(
        context=context,
        config=config,
        handle_exported_document_tool=AsyncMock(),
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
    )

    try:
        assert runtime.temp_dir is None
        assert runtime.plugin_data_path == data_root / "files"
        assert runtime.plugin_data_path.exists()
        assert runtime.settings.auto_delete is False
        assert runtime.settings.max_file_size == 16 * 1024 * 1024
        assert runtime.settings.reply_to_user is False
        assert runtime.settings.require_at_in_group is False
        assert runtime.settings.enable_features_in_group is True
        assert runtime.settings.auto_block_execution_tools is False
        assert runtime.settings.enable_preview is True
        assert runtime.settings.preview_dpi == 180
        assert runtime.settings.allow_external_input_files is True
        assert runtime.settings.recent_text_ttl_seconds == 17
        assert runtime.settings.recent_text_cleanup_interval_seconds == 17
        assert runtime.command_service._plugin_data_path == data_root / "files"
        assert runtime.workspace_service.plugin_data_path == data_root / "files"
    finally:
        runtime.executor.shutdown(wait=False)
        runtime.office_gen.cleanup()
        runtime.pdf_converter.cleanup()
        shutil.rmtree(data_root, ignore_errors=True)


def test_workspace_service_pre_check_rejects_outside_workspace():
    workspace_dir = Path(__file__).resolve().parent / "workspace-root"
    outside_file = Path(__file__).resolve()

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024,
            feature_settings={},
        )

        ok, resolved_path, err = service.pre_check(
            _build_event(),
            str(outside_file),
            require_exists=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission_fn=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        assert not ok
        assert resolved_path is None
        assert err == "错误：非法路径：禁止访问工作区外的文件"
    finally:
        executor.shutdown(wait=False)


def test_upload_session_service_skips_system_notice_messages():
    service = UploadSessionService(
        context=MagicMock(),
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
    )
    event = _build_event()
    session_key = service.get_attachment_session_key(event)

    event.message_str = "[System Notice] internal guidance"
    service.remember_recent_text(event)
    assert session_key not in service.recent_text_by_session

    event.message_str = "整理成正式汇报"
    service.remember_recent_text(event)
    assert service.recent_text_by_session[session_key][0] == "整理成正式汇报"


@pytest.mark.asyncio
async def test_upload_session_service_builds_read_first_prompt_for_buffered_upload():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
    )
    event = _build_event()
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(event=event, files=[upload], texts=[])

    await service.on_buffer_complete(buf)

    assert isinstance(event.message_obj.message[0], Comp.Plain)
    prompt_text = event.message_obj.message[0].text
    assert "请现在调用 `read_file`。" in prompt_text
    assert "读取上传源文件前，不要先创建新文档。" in (prompt_text)
    assert "目前用户意图还不够明确，读取后再用中文追问。" in prompt_text
    assert event.message_str == prompt_text.strip()
    event_queue.put.assert_awaited_once_with(event)


@pytest.mark.asyncio
async def test_delivery_service_sends_message_and_file_for_existing_path():
    event = _build_event()
    event.send = AsyncMock()
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=True,
    )
    file_path = Path(__file__).resolve()
    try:
        await service.send_file_with_preview(event, file_path, "✅ 已发送")
    finally:
        service._executor.shutdown(wait=False)

    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_delivery_service_returns_missing_message_for_absent_export():
    event = _build_event()
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=False,
    )
    missing_path = Path(__file__).resolve().parent / "missing-output.docx"
    try:
        result = await service.send_exported_document(
            event,
            missing_path,
            "✅ exported",
        )
    finally:
        service._executor.shutdown(wait=False)

    assert "but the file does not exist" in result


@pytest.mark.asyncio
async def test_delivery_service_handles_exported_document_tool_via_context_event():
    event = _build_event()
    event.send = AsyncMock()
    context = SimpleNamespace(context=SimpleNamespace(event=event))
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=False,
    )
    file_path = Path(__file__).resolve()
    try:
        result = await service.handle_exported_document_tool(
            context,
            str(file_path),
            "✅ exported",
        )
    finally:
        service._executor.shutdown(wait=False)

    assert result == f"Document exported and sent to the user: {file_path.name}"
    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_delivery_service_logs_preview_generation_failure():
    event = _build_event()
    event.send = AsyncMock()
    preview_generator = MagicMock()
    preview_generator.generate_preview.side_effect = RuntimeError("preview boom")
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=preview_generator,
        enable_preview=True,
        auto_delete=False,
        reply_to_user=False,
    )
    file_path = Path(__file__).resolve()
    try:
        with patch(
            "astrbot_plugin_office_assistant.services.delivery_service.logger.warning"
        ) as warning_mock:
            await service.send_file_with_preview(event, file_path, "✅ 已发送")
    finally:
        service._executor.shutdown(wait=False)

    warning_mock.assert_called_once()
    assert "生成预览图失败" in warning_mock.call_args.args[0]
    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_delivery_service_logs_file_cleanup_failure():
    event = _build_event()
    event.send = AsyncMock()
    workspace_dir = _make_workspace("delivery-cleanup")
    preview_path = workspace_dir / "preview.png"
    preview_path.write_text("preview", encoding="utf-8")
    file_path = workspace_dir / "result.docx"
    file_path.write_text("result", encoding="utf-8")

    preview_generator = MagicMock()
    preview_generator.generate_preview.return_value = preview_path

    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=preview_generator,
        enable_preview=True,
        auto_delete=True,
        reply_to_user=False,
    )

    with patch.object(Path, "unlink", side_effect=PermissionError("locked")):
        with patch(
            "astrbot_plugin_office_assistant.services.delivery_service.logger.warning"
        ) as warning_mock:
            try:
                await service.send_file_with_preview(event, file_path, "✅ 已发送")
            finally:
                service._executor.shutdown(wait=False)
                shutil.rmtree(workspace_dir, ignore_errors=True)

    assert warning_mock.call_count == 2
    logged_messages = [call.args[0] for call in warning_mock.call_args_list]
    assert any("自动删除预览文件失败" in message for message in logged_messages)
    assert any("自动删除文件失败" in message for message in logged_messages)


@pytest.mark.asyncio
async def test_incoming_message_service_buffers_supported_file_and_stops_event():
    message_buffer = MagicMock()
    message_buffer.add_message = AsyncMock(return_value=True)
    message_buffer.is_buffering.return_value = False
    remember_recent_text = MagicMock()
    service = IncomingMessageService(
        message_buffer=message_buffer,
        remember_recent_text=remember_recent_text,
        is_group_feature_enabled=lambda _event: True,
    )
    event = _build_event()
    event.stop_event = MagicMock()
    event.message_obj.message = [Comp.File(name="report.docx", file="report.docx")]

    await service.handle_file_message(event)

    remember_recent_text.assert_called_once_with(event)
    message_buffer.add_message.assert_awaited_once_with(event)
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_incoming_message_service_keeps_existing_buffer_for_unsupported_file():
    message_buffer = MagicMock()
    message_buffer.add_message = AsyncMock(return_value=True)
    message_buffer.is_buffering.return_value = True
    remember_recent_text = MagicMock()
    service = IncomingMessageService(
        message_buffer=message_buffer,
        remember_recent_text=remember_recent_text,
        is_group_feature_enabled=lambda _event: True,
    )
    event = _build_event()
    event.stop_event = MagicMock()
    event.message_obj.message = [Comp.File(name="avatar.png", file="avatar.png")]

    await service.handle_file_message(event)

    remember_recent_text.assert_called_once_with(event)
    message_buffer.is_buffering.assert_called_once_with(event)
    message_buffer.add_message.assert_awaited_once_with(event)
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_error_hook_service_uses_event_session_when_target_not_configured():
    context = MagicMock()
    context.send_message = AsyncMock(return_value=True)
    service = ErrorHookService(
        context=context,
        config={},
        plugin_name="astrbot_plugin_office_assistant",
    )
    event = _build_event()
    event.stop_event = MagicMock()

    await service.handle_plugin_error(
        event,
        "astrbot_plugin_office_assistant",
        "handler_name",
        RuntimeError("boom"),
        "traceback line 1\ntraceback line 2",
    )

    context.send_message.assert_awaited_once()
    args, _kwargs = context.send_message.await_args
    assert args[0] == "session-1"
    assert "handler=handler_name" in args[1].chain[0].text
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_file_tool_service_reads_text_from_workspace_file():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.read_file(event, Path(__file__).resolve().name)
    finally:
        executor.shutdown(wait=False)

    assert result is not None
    assert "[文件:" in result
    assert "test_file_tool_service_reads_text_from_workspace_file" in result


@pytest.mark.asyncio
async def test_file_tool_service_returns_local_guidance_for_missing_file():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.read_file(event, "CLAUDE.md")
    finally:
        executor.shutdown(wait=False)

    assert result is not None
    assert "错误：文件 'CLAUDE.md' 不存在。" in result
    assert "不要联网搜索" in result
    assert "重新上传文件或提供正确的本地路径" in result


@pytest.mark.asyncio
async def test_file_tool_service_creates_office_file_via_generator_and_delivery():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    event.send = AsyncMock()
    office_generator = MagicMock()
    office_generator.generate = AsyncMock(return_value=Path(__file__).resolve())
    delivery_service = MagicMock()
    delivery_service.send_file_with_preview = AsyncMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report.docx",
            content="hello world",
            file_type="word",
        )
    finally:
        executor.shutdown(wait=False)

    office_generator.generate.assert_awaited_once()
    delivery_service.send_file_with_preview.assert_awaited_once()
    assert result is None


@pytest.mark.asyncio
async def test_file_tool_service_create_office_file_returns_error_without_sending():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    event.send = AsyncMock()
    office_generator = MagicMock()
    delivery_service = MagicMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report.unsupported",
            content="hello world",
            file_type="unknown",
        )
    finally:
        executor.shutdown(wait=False)

    assert result == "错误：不支持的文件类型 'unknown'"
    event.send.assert_not_called()
    office_generator.generate.assert_not_called()
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_file_tool_service_convert_to_pdf_returns_none_after_delivery():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    source_path = workspace_dir / "convert-source.docx"
    source_path.write_text("demo", encoding="utf-8")
    output_path = workspace_dir / "convert-source.pdf"
    pdf_converter = MagicMock()
    pdf_converter.is_available.return_value = True
    pdf_converter.office_to_pdf = AsyncMock(return_value=output_path)
    delivery_service = MagicMock()
    delivery_service.send_file_with_preview = AsyncMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_pdf_conversion": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=pdf_converter,
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )
        output_path.write_text("pdf", encoding="utf-8")

        result = await service.convert_to_pdf(
            event,
            filename=source_path.name,
        )
    finally:
        source_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        executor.shutdown(wait=False)

    pdf_converter.office_to_pdf.assert_awaited_once_with(source_path)
    delivery_service.send_file_with_preview.assert_awaited_once()
    assert result is None


def test_command_service_lists_office_files_in_workspace():
    workspace_dir = Path(__file__).resolve().parent / "workspace-command-list"
    workspace_dir.mkdir(exist_ok=True)
    sample_file = workspace_dir / "report.docx"
    sample_file.write_text("demo", encoding="utf-8")

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        pdf_converter = MagicMock()
        pdf_converter.capabilities = {
            "office_to_pdf": False,
            "pdf_to_word": False,
            "pdf_to_excel": False,
        }
        service = CommandService(
            workspace_service=workspace_service,
            pdf_converter=pdf_converter,
            plugin_data_path=workspace_dir,
            auto_delete=False,
            allow_external_input_files=False,
            enable_features_in_group=True,
            auto_block_execution_tools=True,
            reply_to_user=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = service.list_files(_build_event())
    finally:
        sample_file.unlink(missing_ok=True)
        workspace_dir.rmdir()
        executor.shutdown(wait=False)

    assert "机器人工作区 Office 文件列表" in result
    assert "report.docx" in result


def test_command_service_builds_pdf_status_summary():
    executor = ThreadPoolExecutor(max_workers=1)
    try:
        workspace_service = WorkspaceService(
            plugin_data_path=Path(__file__).resolve().parent,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        pdf_converter = MagicMock()
        pdf_converter.get_detailed_status.return_value = {
            "capabilities": {
                "office_to_pdf": True,
                "pdf_to_word": False,
                "pdf_to_excel": True,
            },
            "office_to_pdf_backend": "libreoffice",
            "word_backend": None,
            "excel_backend": "tabula",
            "is_windows": False,
            "java_available": True,
            "libreoffice_path": "/usr/bin/libreoffice",
            "libs": {"pdf2docx": False, "tabula-py": True},
        }
        pdf_converter.get_missing_dependencies.return_value = ["pdf2docx"]
        service = CommandService(
            workspace_service=workspace_service,
            pdf_converter=pdf_converter,
            plugin_data_path=Path(__file__).resolve().parent,
            auto_delete=False,
            allow_external_input_files=False,
            enable_features_in_group=True,
            auto_block_execution_tools=True,
            reply_to_user=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = service.pdf_status(_build_event())
    finally:
        executor.shutdown(wait=False)

    assert "Office→PDF: ✅ 可用 (libreoffice)" in result
    assert "PDF→Excel:  ✅ 可用 (tabula)" in result
    assert "缺失依赖" in result
    assert "pdf2docx" in result
