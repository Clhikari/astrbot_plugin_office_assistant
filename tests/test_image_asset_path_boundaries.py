from __future__ import annotations

import json
import os
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from astrbot_plugin_office_assistant.services.image_asset_service import (
    ImageAssetService,
)

SESSION_A = ("platform", "owner", "origin")
SESSION_B = ("platform", "other", "origin")


@pytest.fixture
def image_pool(tmp_path: Path):
    images = tmp_path / "images"
    external = tmp_path / "external"
    images.mkdir()
    external.mkdir()
    for path in (images / "own.png", images / "other.png", external / "target.png"):
        Image.new("RGB", (2, 2), color="red").save(path)
    return SimpleNamespace(root=tmp_path, images=images, external=external)


def _load_service(pool, ref: str) -> ImageAssetService:
    records = [
        {
            "ref": item_ref,
            "original_name": "source.png",
            "note": "",
            "width": 2,
            "height": 2,
            "format": "PNG",
            "size_bytes": (pool.images / "own.png").stat().st_size,
            "registered_at": 1750000000.0,
            "session_key": list(session),
        }
        for item_ref, session in (
            ("images/own.png", SESSION_A),
            ("images/other.png", SESSION_B),
            (ref, SESSION_A),
        )
    ]
    (pool.images / "index.json").write_text(json.dumps(records), encoding="utf-8")
    return ImageAssetService(plugin_data_path=pool.root)


@contextmanager
def _directory_link(link: Path, target: Path):
    if os.name == "nt":
        import _winapi

        _winapi.CreateJunction(str(target), str(link))
    else:
        link.symlink_to(target, target_is_directory=True)
    try:
        yield
    finally:
        if os.name == "nt":
            link.rmdir()
        else:
            link.unlink()


def _assert_cannot_resolve(service: ImageAssetService, ref: str) -> None:
    with pytest.raises(ValueError):
        service.resolve_ref(ref, session_key=SESSION_A)
    assert service.ref_exists(ref, session_key=SESSION_A) is False


@pytest.mark.parametrize("link_after_load", [False, True], ids=["persisted-link", "link-after-load"])
def test_directory_link_cannot_resolve_outside_pool(image_pool, link_after_load):
    ref = "images/link/target.png"
    service = _load_service(image_pool, ref) if link_after_load else None
    with _directory_link(image_pool.images / "link", image_pool.external):
        service = service or _load_service(image_pool, ref)
        _assert_cannot_resolve(service, ref)


@pytest.mark.parametrize("clear_all", [False, True], ids=["clear-ref", "clear-session"])
def test_clear_directory_link_preserves_external_file_and_other_session(image_pool, clear_all):
    ref = "images/link/target.png"
    service = _load_service(image_pool, ref)
    external_file = image_pool.external / "target.png"
    original_bytes = external_file.read_bytes()
    own_file = image_pool.images / "own.png"
    own_bytes = own_file.read_bytes()
    other_file = image_pool.images / "other.png"
    other_bytes = other_file.read_bytes()

    with _directory_link(image_pool.images / "link", image_pool.external):
        service.clear_images(session_key=SESSION_A, ref=None if clear_all else ref)
        assert external_file.read_bytes() == original_bytes
        assert own_file.exists() is not clear_all
        if not clear_all:
            assert own_file.read_bytes() == own_bytes
        assert other_file.read_bytes() == other_bytes
        assert service.resolve_ref("images/other.png", session_key=SESSION_B) == other_file.resolve()
        assert [item["ref"] for item in service.list_images(SESSION_A)] == (
            [] if clear_all else ["images/own.png"]
        )


@pytest.mark.parametrize("external_target", [False, True], ids=["internal-target", "external-target"])
def test_final_file_symlink_is_rejected_without_deleting_target(image_pool, external_target):
    ref = "images/linked.png"
    service = _load_service(image_pool, ref)
    target = image_pool.external / "target.png" if external_target else image_pool.images / "own.png"
    original_bytes = target.read_bytes()
    link = image_pool.images / "linked.png"
    try:
        link.symlink_to(target)
    except OSError as exc:
        if os.name == "nt" and exc.winerror == 1314:
            pytest.skip("Windows does not permit file symlinks for this test process")
        raise
    try:
        _assert_cannot_resolve(service, ref)
        service.clear_images(session_key=SESSION_A, ref=ref)
        assert target.read_bytes() == original_bytes
    finally:
        link.unlink(missing_ok=True)


def test_regular_nested_image_can_resolve_and_clear(image_pool):
    nested = image_pool.images / "nested" / "image.png"
    nested.parent.mkdir()
    nested.write_bytes((image_pool.images / "own.png").read_bytes())
    ref = "images/nested/image.png"
    service = _load_service(image_pool, ref)

    assert service.resolve_ref(ref, session_key=SESSION_A) == nested.resolve()
    assert service.ref_exists(ref, session_key=SESSION_A) is True
    assert service.clear_images(session_key=SESSION_A, ref=ref) == 1
    assert not nested.exists()
    assert (image_pool.images / "own.png").is_file()
    assert (image_pool.images / "other.png").is_file()
