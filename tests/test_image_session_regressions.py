from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import astrbot.api.message_components as Comp
import pytest
from astrbot.core.platform.message_type import MessageType
from PIL import Image

from astrbot_plugin_office_assistant.services import upload_session_service
from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.services.command_service import CommandService
from astrbot_plugin_office_assistant.services.image_asset_service import (
    ImageAssetService,
)
from astrbot_plugin_office_assistant.services.image_file_utils import (
    is_image_file_component,
)
from astrbot_plugin_office_assistant.services.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.services.upload_session_service import (
    UploadSessionService,
)


def _event(sender="user-1"):
    event = MagicMock()
    event.get_platform_id.return_value = "audit-platform"
    event.get_sender_id.return_value = sender
    event.unified_msg_origin = "audit-group"
    event._buffered = False
    event._buffer_reentry_count = 0
    event.message_obj = SimpleNamespace(type=MessageType.GROUP_MESSAGE, message=[])
    return event


@pytest.fixture
def clock(monkeypatch):
    now = [1000.0]
    monkeypatch.setattr(upload_session_service.time, "time", lambda: now[0])
    return now


@pytest.fixture
def service():
    result = UploadSessionService(
        context=MagicMock(),
        recent_text_ttl_seconds=60,
        upload_session_ttl_seconds=120,
        recent_text_max_entries=100,
        recent_text_cleanup_interval_seconds=10,
        upload_session_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(),
        store_uploaded_file=lambda path, _name: path,
        allow_external_input_files=False,
    )
    yield result
    result.cleanup()


@pytest.mark.asyncio
@pytest.mark.parametrize("image_name", ["attachment.bin", ""])
async def test_mixed_buffer_retains_mime_image_outside_doc_list(
    service, tmp_path, image_name
):
    document_path = tmp_path / "report.txt"
    document_path.write_text("fixture", encoding="utf-8")
    image_path = tmp_path / "download.bin"
    Image.new("RGB", (5, 5)).save(image_path, format="PNG")
    image_file = Comp.File(name=image_name, file=str(image_path))
    object.__setattr__(image_file, "mime_type", "image/png")
    assert is_image_file_component(image_file)
    document_file = Comp.File(name="report.txt", file=str(document_path))
    service._extract_upload_source.side_effect = [
        (document_path, document_file.name),
        (image_path, image_name),
    ]
    event = _event()

    await service.on_buffer_complete(
        BufferedMessage(event=event, files=[document_file, image_file])
    )

    assert service.get_pending_image_resources(event) == [image_path]
    assert [
        info["original_name"] for info in service.list_session_upload_infos(event)
    ] == ["report.txt"]
    assert service.get_pending_images(event)[0].read_bytes() == image_path.read_bytes()


@pytest.mark.parametrize("activity", ["cache_image", "read_images", "text"])
def test_activity_expires_idle_sessions(service, clock, activity):
    idle_event = _event("idle-user")
    idle_key = service.get_attachment_session_key(idle_event)
    service.cache_pending_image_resource(idle_event, object())
    clock[0] += 120
    active_event = _event("active-user")

    if activity == "cache_image":
        service.cache_pending_image_resource(active_event, object())
    elif activity == "read_images":
        service.get_pending_image_resources(active_event)
    else:
        active_event.message_obj.message = [Comp.Plain("hello")]
        service.remember_recent_text(active_event)

    assert idle_key not in service._pending_images_by_session


@pytest.mark.parametrize("activity", ["read", "write"])
def test_busy_session_expires_old_items_between_global_cleanups(
    service, clock, activity
):
    event = _event()
    assert service.get_pending_image_resources(event) == []
    assert service._pending_images_by_session == {}
    service.cache_pending_image_resource(event, object())
    clock[0] += 119
    service.list_session_upload_infos(_event("another-user"))
    clock[0] += 1
    if activity == "write":
        fresh_resource = object()
        service.cache_pending_image_resource(event, fresh_resource)
        session_key = service.get_attachment_session_key(event)
        assert [
            resource for resource, _ in service._pending_images_by_session[session_key]
        ] == [fresh_resource]
    else:
        assert service.get_pending_image_resources(event) == []
        assert service._pending_images_by_session == {}


def test_global_cleanup_keeps_fresh_resources_and_selective_clear(service, clock):
    event = _event()
    service.cache_pending_image_resource(event, object())
    clock[0] += 60
    first_fresh, second_fresh = object(), object()
    service.cache_pending_image_resource(event, first_fresh)
    service.cache_pending_image_resource(event, second_fresh)
    clock[0] += 60

    service.list_session_upload_infos(_event("another-user"))
    service.clear_pending_image_resources(event, [first_fresh])

    assert service.get_pending_image_resources(event) == [second_fresh]


@pytest.mark.parametrize(
    "scope,senders,limit,retained",
    [
        ("PER_SESSION", ["only"] * 3, 2, {"only": [1, 2]}),
        (
            "TOTAL",
            ["oldest", "middle", "newest", "newest"],
            3,
            {"oldest": [], "middle": [1], "newest": [2, 3]},
        ),
    ],
)
def test_pending_image_limits_preserve_newest(
    service, clock, monkeypatch, scope, senders, limit, retained
):
    monkeypatch.setattr(service, f"_MAX_PENDING_IMAGES_{scope}", limit)
    events = {sender: _event(sender) for sender in senders}
    resources = [object() for _ in senders]
    for sender, resource in zip(senders, resources):
        service.cache_pending_image_resource(events[sender], resource)
        clock[0] += 1
    for sender, indices in retained.items():
        event = events[sender]
        assert service.get_pending_image_resources(event) == [
            resources[i] for i in indices
        ]
        if not indices:
            assert (
                service.get_attachment_session_key(event)
                not in service._pending_images_by_session
            )


@pytest.fixture
def waiting_request(service, tmp_path):
    event = _event()
    image_path = tmp_path / "original.png"
    Image.new("RGB", (5, 5)).save(image_path)
    service.cache_pending_image_resource(event, image_path)
    event._has_pending_images = True
    event._pending_image_resources = [image_path]
    plugin = object.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        upload_session_service=service,
        settings=SimpleNamespace(image_llm_delay=3),
        llm_request_policy=SimpleNamespace(apply=AsyncMock()),
        message_buffer=SimpleNamespace(pop_images=AsyncMock(return_value=[])),
    )
    return plugin, event


@pytest.mark.asyncio
@pytest.mark.parametrize("removal_reason", ["expiry", "capacity"])
async def test_image_request_continues_after_cache_eviction(
    service, clock, monkeypatch, waiting_request, removal_reason
):
    plugin, event = waiting_request
    request = object()

    async def evict_while_waiting(_delay):
        if removal_reason == "expiry":
            clock[0] += 120
        else:
            monkeypatch.setattr(service, "_MAX_PENDING_IMAGES_PER_SESSION", 1)
            service.cache_pending_image_resource(event, object())

    monkeypatch.setattr("asyncio.sleep", evict_while_waiting)

    await plugin.before_llm_chat(event, request)

    event.stop_event.assert_not_called()
    plugin._runtime.llm_request_policy.apply.assert_awaited_once_with(event, request)


@pytest.mark.asyncio
async def test_img_add_consumption_still_stops_waiting_image_request(
    service, tmp_path, monkeypatch, waiting_request
):
    plugin, event = waiting_request
    image_service = ImageAssetService(plugin_data_path=tmp_path / "pool")
    command_service = object.__new__(CommandService)
    command_service._require_access = lambda _event: None
    command_service._upload_session_service = service
    command_service._image_asset_service = image_service
    plugin._runtime.command_service = command_service

    async def add_image_while_waiting(_delay):
        await plugin.img_add(_event(), "cover")

    monkeypatch.setattr("asyncio.sleep", add_image_while_waiting)

    await plugin.before_llm_chat(event, object())

    assert (
        len(image_service.list_images(service.get_attachment_session_key(event))) == 1
    )
    event.stop_event.assert_called_once()
    plugin._runtime.llm_request_policy.apply.assert_not_awaited()


def test_image_consumption_records_are_bounded_and_expire(service, clock, monkeypatch):
    monkeypatch.setattr(service, "_MAX_PENDING_IMAGES_TOTAL", 2)
    event = _event()
    resources = [object() for _ in range(3)]
    for resource in resources:
        service.cache_pending_image_resource(event, resource)
        service.clear_pending_image_resources(event, [resource])
        clock[0] += 1

    assert not service.were_pending_image_resources_consumed(event, resources[:1])
    assert service.were_pending_image_resources_consumed(event, resources[-2:])
    assert len(service._consumed_pending_images) == 2
    clock[0] += 120
    service.list_session_upload_infos(_event("another-user"))
    assert service._consumed_pending_images == {}


def test_reused_image_resource_does_not_inherit_consumption_marker(service, clock):
    event = _event()
    resource = object()
    service.cache_pending_image_resource(event, resource)
    service.clear_pending_image_resources(event, [resource])
    assert service.were_pending_image_resources_consumed(event, [resource])
    assert not service.were_pending_image_resources_consumed(
        _event("other"), [resource]
    )

    service.cache_pending_image_resource(event, resource)

    assert not service.were_pending_image_resources_consumed(event, [resource])
    assert service.get_pending_image_resources(event) == [resource]


@pytest.mark.parametrize("clear_all", [False, True])
def test_clearing_expired_images_does_not_mark_them_consumed(service, clock, clear_all):
    event = _event()
    resource = object()
    service.cache_pending_image_resource(event, resource)
    clock[0] += 120

    if clear_all:
        service.clear_pending_images(event)
    else:
        service.clear_pending_image_resources(event, [resource])

    assert not service.were_pending_image_resources_consumed(event, [resource])
    assert service._pending_images_by_session == {}
