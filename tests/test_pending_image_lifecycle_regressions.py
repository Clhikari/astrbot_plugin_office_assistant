import asyncio
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import astrbot.api.message_components as Comp
import pytest
from astrbot.api.event import AstrMessageEvent
from astrbot.api.platform import (
    AstrBotMessage,
    MessageMember,
    MessageType,
    PlatformMetadata,
)
from PIL import Image

from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.services.command_service import CommandService
from astrbot_plugin_office_assistant.services.image_asset_service import (
    ImageAssetService,
)
from astrbot_plugin_office_assistant.services.incoming_message_service import (
    IncomingMessageService,
)
from astrbot_plugin_office_assistant.services.message_buffer import MessageBuffer
from astrbot_plugin_office_assistant.services.upload_session_service import (
    UploadSessionService,
)
from astrbot_plugin_office_assistant.services import upload_session_service


def _event(components=(), *, sender="image-user"):
    message = AstrBotMessage()
    message.type = MessageType.FRIEND_MESSAGE
    message.sender = MessageMember(sender, sender)
    message.message = list(components)
    message.message_str = "image lifecycle fixture"
    message.message_id = "image-lifecycle-message"
    message.self_id = "fixture-bot"
    return AstrMessageEvent(
        message.message_str,
        message,
        PlatformMetadata("webchat", "fixture", "webchat"),
        "fixture-session",
    )


@pytest.fixture
def service():
    queue = asyncio.Queue()

    async def extract(component):
        return Path(await component.get_file()), component.name

    result = UploadSessionService(
        context=SimpleNamespace(
            get_event_queue=lambda: queue, get_config=lambda *_: {}
        ),
        recent_text_ttl_seconds=60,
        upload_session_ttl_seconds=120,
        recent_text_max_entries=100,
        recent_text_cleanup_interval_seconds=10,
        upload_session_cleanup_interval_seconds=10,
        extract_upload_source=extract,
        store_uploaded_file=lambda path, _name: path,
        allow_external_input_files=False,
    )
    yield result
    result.cleanup()


@pytest.fixture
def source_image(tmp_path):
    source = tmp_path / "中文 空格 %.png"
    Image.new("RGB", (7, 5), "orange").save(source)
    return source, source.read_bytes()


def _plugin(service, buffer, image_service):
    command = object.__new__(CommandService)
    command._require_access = lambda _: None
    command._upload_session_service = service
    command._image_asset_service = image_service
    plugin = object.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        upload_session_service=service, message_buffer=buffer, command_service=command
    )
    return plugin


@pytest.mark.asyncio
@pytest.mark.parametrize("mixed", [False, True])
async def test_real_sdk_event_cleanup_does_not_destroy_pending_image(
    service, tmp_path, monkeypatch, source_image, mixed
):
    if not hasattr(AstrMessageEvent, "cleanup_temporary_local_files"):
        pytest.skip("SDK does not clean event media")
    try:
        from astrbot.core.utils import media_utils
        from astrbot.core.pipeline.preprocess_stage import stage as preprocess_module
    except ImportError:
        pytest.skip("SDK does not normalize images with MediaResolver")
    if not hasattr(preprocess_module.PreProcessStage, "_track_temp_media"):
        pytest.skip("SDK does not track normalized image media for event cleanup")

    sdk_temp = tmp_path / "sdk-temp"
    sdk_temp.mkdir()
    monkeypatch.setattr(media_utils, "get_astrbot_temp_path", lambda: str(sdk_temp))
    monkeypatch.setattr(
        preprocess_module, "get_astrbot_temp_path", lambda: str(sdk_temp)
    )
    source, original_bytes = source_image
    component = Comp.Image.fromFileSystem(str(source))
    components = [component]
    if mixed:
        document = tmp_path / "mixed.docx"
        document.write_bytes(b"buffer fixture")
        components.insert(0, Comp.File(name=document.name, file=str(document)))
    event = _event(components)
    preprocess = preprocess_module.PreProcessStage()
    await preprocess.initialize(SimpleNamespace(astrbot_config={}, plugin_manager=None))
    await preprocess.process(event)
    normalized = Path(await component.convert_to_file_path())
    expected_bytes = normalized.read_bytes()
    assert normalized.suffix == ".jpg"

    buffer = MessageBuffer(wait_seconds=60)
    incoming = IncomingMessageService(
        message_buffer=buffer,
        remember_recent_text=service.remember_recent_text,
        is_group_feature_enabled=lambda _: True,
        cache_pending_image_resource=service.cache_pending_image_resource,
    )
    images = ImageAssetService(plugin_data_path=tmp_path / "pool")
    plugin = _plugin(service, buffer, images)
    try:
        await incoming.handle_file_message(event)
        event.cleanup_temporary_local_files()
        assert not normalized.exists()
        assert source.read_bytes() == original_bytes
        await plugin.img_add(_event(), "lifecycle")
        registered = images.list_images(service.get_attachment_session_key(event))
        assert len(registered) == 1
        assert (tmp_path / "pool" / registered[0]["ref"]).read_bytes() == expected_bytes
        assert service.get_pending_image_resources(event) == []
        assert service._pending_image_snapshots == {}
        assert list(Path(service._pending_image_temp_dir.name).iterdir()) == []
        assert list(sdk_temp.iterdir()) == []
    finally:
        await buffer.cancel_buffer(event)
        event.cleanup_temporary_local_files()


@pytest.mark.parametrize(
    "removal", ["expiry", "session_capacity", "global_capacity", "clear", "terminate"]
)
def test_pending_snapshot_cleanup_never_deletes_external_source(
    service, monkeypatch, source_image, removal
):
    source, expected = source_image
    resource = Comp.Image.fromFileSystem(str(source))
    event = _event()
    other_event = _event(sender="other-user")
    service.cache_pending_image_resource(event, resource)
    snapshot = service.get_pending_image_snapshot(event, resource)
    assert snapshot is not None and snapshot != source
    assert snapshot.read_bytes() == expected
    assert service.get_pending_image_snapshot(other_event, resource) is None
    service.clear_pending_image_resources(other_event, [resource])
    assert snapshot.exists()
    service.cache_pending_image_resource(event, resource)
    assert service.get_pending_image_resources(event) == [resource]
    assert service.get_pending_image_snapshot(event, resource) == snapshot

    if removal == "expiry":
        cached_at = service._pending_images_by_session[
            service.get_attachment_session_key(event)
        ][0][1]
        monkeypatch.setattr(
            upload_session_service.time, "time", lambda: cached_at + 120
        )
        assert service.get_pending_image_resources(event) == []
    elif removal == "session_capacity":
        monkeypatch.setattr(service, "_MAX_PENDING_IMAGES_PER_SESSION", 1)
        service.cache_pending_image_resource(event, object())
    elif removal == "global_capacity":
        monkeypatch.setattr(service, "_MAX_PENDING_IMAGES_TOTAL", 1)
        service.cache_pending_image_resource(other_event, object())
    elif removal == "clear":
        service.clear_pending_images(event)
    else:
        service.cleanup()
    assert not snapshot.exists()
    assert source.read_bytes() == expected
    assert service._pending_image_snapshots == {}


@pytest.mark.asyncio
async def test_failed_image_registration_keeps_snapshot_for_retry(
    service, tmp_path, monkeypatch, source_image
):
    source, _ = source_image
    event = _event()
    service.cache_pending_image_resource(event, source)
    snapshot = service.get_pending_image_snapshot(event, source)
    images = ImageAssetService(plugin_data_path=tmp_path / "pool")
    buffer = MessageBuffer()
    plugin = _plugin(service, buffer, images)
    real_register = images.register_image

    monkeypatch.setattr(
        images,
        "register_image",
        Mock(side_effect=ValueError("temporary registration failure")),
    )
    await plugin.img_add(event, "retry")
    assert service.get_pending_image_resources(event) == [source]
    assert snapshot.exists()
    assert not service.were_pending_image_resources_consumed(event, [source])
    monkeypatch.setattr(images, "register_image", real_register)
    await plugin.img_add(event, "retry")
    assert len(images.list_images(service.get_attachment_session_key(event))) == 1
    assert not snapshot.exists()
    assert source.exists()
    assert service.get_pending_image_resources(event) == []


@pytest.mark.parametrize("resource_type", ["path", "image", "file"])
def test_pending_snapshot_is_independent_of_external_file(
    service, source_image, resource_type
):
    source, expected = source_image
    resource = {
        "path": source,
        "image": Comp.Image.fromFileSystem(str(source)),
        "file": Comp.File(name=source.name, file=source.as_uri()),
    }[resource_type]
    event = _event()
    service.cache_pending_image_resource(event, resource)
    source.unlink()
    snapshot = service.get_pending_image_snapshot(event, resource)
    assert snapshot is not None
    assert snapshot.read_bytes() == expected
    service.cache_pending_image_resource(event, resource)
    assert service.get_pending_image_snapshot(event, resource) == snapshot


def test_snapshot_copy_and_cleanup_failures_do_not_interrupt_cache(
    service, monkeypatch, source_image
):
    source, expected = source_image
    with monkeypatch.context() as scoped:
        scoped.setattr(
            upload_session_service.shutil,
            "copyfile",
            Mock(side_effect=OSError("copy failed")),
        )
        scoped.setattr(Path, "unlink", Mock(side_effect=OSError("cleanup failed")))
        event = _event()
        service.cache_pending_image_resource(event, source)
    assert service.get_pending_image_resources(event) == [source]
    assert service.get_pending_image_snapshot(event, source) is None
    assert source.read_bytes() == expected


@pytest.mark.asyncio
async def test_snapshot_cleanup_failure_does_not_skip_other_plugin_shutdown():
    released = []
    plugin = object.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        message_buffer=None,
        upload_session_service=SimpleNamespace(
            cleanup=Mock(side_effect=OSError("snapshot still in use"))
        ),
        office_gen=SimpleNamespace(cleanup=lambda: released.append("office")),
        pdf_converter=SimpleNamespace(cleanup=lambda: released.append("pdf")),
        executor=SimpleNamespace(shutdown=lambda **_: released.append("executor")),
        temp_dir=None,
    )
    await plugin.terminate()
    assert released == ["office", "pdf", "executor"]
    assert plugin._runtime is None
