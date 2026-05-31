from __future__ import annotations

from pathlib import Path

import pytest

from astrbot_plugin_office_assistant.services.image_asset_service import (
    ImageAssetService,
)


@pytest.fixture()
def image_service(tmp_path: Path) -> ImageAssetService:
    return ImageAssetService(plugin_data_path=tmp_path)


@pytest.fixture()
def sample_png(tmp_path: Path) -> Path:
    from PIL import Image

    img = Image.new("RGB", (100, 50), color="red")
    path = tmp_path / "sample.png"
    img.save(path, format="PNG")
    return path


@pytest.fixture()
def sample_jpeg(tmp_path: Path) -> Path:
    from PIL import Image

    img = Image.new("RGB", (200, 100), color="blue")
    path = tmp_path / "sample.jpg"
    img.save(path, format="JPEG")
    return path


SESSION_A = ("platform1", "user1", "origin1")
SESSION_B = ("platform1", "user2", "origin2")


class TestImageAssetServiceRegistration:
    def test_register_png(self, image_service, sample_png):
        info = image_service.register_image(
            sample_png,
            session_key=SESSION_A,
            note="test image",
            original_name="sample.png",
        )
        assert info["ref"].startswith("images/img_")
        assert info["ref"].endswith(".png")
        assert info["width"] == 100
        assert info["height"] == 50
        assert info["format"] == "PNG"
        assert info["note"] == "test image"
        assert info["original_name"] == "sample.png"
        assert info["session_key"] == list(SESSION_A)

    def test_register_jpeg(self, image_service, sample_jpeg):
        info = image_service.register_image(sample_jpeg, session_key=SESSION_A)
        assert info["ref"].endswith(".jpg")
        assert info["format"] == "JPEG"
        assert info["width"] == 200
        assert info["height"] == 100

    def test_register_webp_converts_to_png(self, image_service, tmp_path):
        from PIL import Image

        img = Image.new("RGB", (120, 60), color="green")
        webp_file = tmp_path / "sample.webp"
        img.save(webp_file, format="WEBP")

        info = image_service.register_image(webp_file, session_key=SESSION_A)
        assert info["ref"].endswith(".png")
        assert info["format"] == "PNG"
        assert info["width"] == 120
        assert info["height"] == 60

        stored = image_service.resolve_ref(info["ref"], session_key=SESSION_A)
        with Image.open(stored) as reopened:
            assert reopened.format == "PNG"

    def test_register_corrects_extension(self, image_service, tmp_path):
        from PIL import Image

        img = Image.new("RGB", (10, 10))
        wrong_ext = tmp_path / "actually_png.bmp"
        img.save(wrong_ext, format="PNG")

        info = image_service.register_image(wrong_ext, session_key=SESSION_A)
        assert info["ref"].endswith(".png")
        assert info["format"] == "PNG"

    def test_register_rejects_svg(self, image_service, tmp_path):
        svg_file = tmp_path / "diagram.svg"
        svg_file.write_text("<svg></svg>")

        with pytest.raises(ValueError, match="不支持 SVG"):
            image_service.register_image(svg_file, session_key=SESSION_A)

    def test_register_rejects_svgz(self, image_service, tmp_path):
        svgz_file = tmp_path / "diagram.svgz"
        svgz_file.write_bytes(b"\x1f\x8b")

        with pytest.raises(ValueError, match="不支持 SVG"):
            image_service.register_image(svgz_file, session_key=SESSION_A)

    def test_register_rejects_invalid_image(self, image_service, tmp_path):
        bad_file = tmp_path / "not_an_image.png"
        bad_file.write_text("this is not an image")

        with pytest.raises(ValueError, match="无法识别为有效图片"):
            image_service.register_image(bad_file, session_key=SESSION_A)

    def test_register_rejects_missing_file(self, image_service, tmp_path):
        missing = tmp_path / "does_not_exist.png"

        with pytest.raises(FileNotFoundError, match="源文件不存在"):
            image_service.register_image(missing, session_key=SESSION_A)


class TestImageAssetServiceSessionIsolation:
    def test_list_only_shows_own_session(self, image_service, sample_png):
        image_service.register_image(sample_png, session_key=SESSION_A, note="A")
        image_service.register_image(sample_png, session_key=SESSION_B, note="B")

        images_a = image_service.list_images(SESSION_A)
        images_b = image_service.list_images(SESSION_B)

        assert len(images_a) == 1
        assert images_a[0]["note"] == "A"
        assert len(images_b) == 1
        assert images_b[0]["note"] == "B"

    def test_clear_only_affects_own_session(self, image_service, sample_png):
        image_service.register_image(sample_png, session_key=SESSION_A)
        image_service.register_image(sample_png, session_key=SESSION_B)

        cleared = image_service.clear_images(session_key=SESSION_A)
        assert cleared == 1
        assert len(image_service.list_images(SESSION_A)) == 0
        assert len(image_service.list_images(SESSION_B)) == 1

    def test_resolve_rejects_other_session_ref(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)

        with pytest.raises(ValueError, match="不在当前会话的资源池中"):
            image_service.resolve_ref(info["ref"], session_key=SESSION_B)


class TestImageAssetServiceResolveRef:
    def test_resolve_valid_ref(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)
        resolved = image_service.resolve_ref(info["ref"], session_key=SESSION_A)
        assert resolved.exists()

    def test_resolve_rejects_non_images_prefix(self, image_service):
        with pytest.raises(ValueError, match="images/"):
            image_service.resolve_ref("other/file.png", session_key=SESSION_A)

    def test_resolve_rejects_traversal(self, image_service):
        with pytest.raises(ValueError, match="traversal|目录遍历"):
            image_service.resolve_ref("images/../etc/passwd", session_key=SESSION_A)

    def test_resolve_rejects_unregistered_ref(self, image_service):
        with pytest.raises(ValueError, match="不在当前会话的资源池中"):
            image_service.resolve_ref("images/fake.png", session_key=SESSION_A)

    def test_resolve_rejects_deleted_file(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)
        file_path = image_service._images_dir.parent / info["ref"]
        file_path.unlink()

        with pytest.raises(FileNotFoundError, match="图片文件不存在"):
            image_service.resolve_ref(info["ref"], session_key=SESSION_A)

    def test_ref_exists_returns_true_for_valid(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)
        assert image_service.ref_exists(info["ref"], session_key=SESSION_A) is True

    def test_ref_exists_returns_false_for_invalid(self, image_service):
        assert (
            image_service.ref_exists("images/nope.png", session_key=SESSION_A) is False
        )


class TestImageAssetServiceNoteAndClear:
    def test_update_note(self, image_service, sample_png):
        info = image_service.register_image(
            sample_png, session_key=SESSION_A, note="old"
        )
        assert image_service.update_note(info["ref"], "new", session_key=SESSION_A)
        images = image_service.list_images(SESSION_A)
        assert images[0]["note"] == "new"

    def test_update_note_wrong_session(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)
        assert not image_service.update_note(info["ref"], "hack", session_key=SESSION_B)

    def test_clear_specific_ref(self, image_service, sample_png, sample_jpeg):
        info1 = image_service.register_image(sample_png, session_key=SESSION_A)
        image_service.register_image(sample_jpeg, session_key=SESSION_A)

        cleared = image_service.clear_images(session_key=SESSION_A, ref=info1["ref"])
        assert cleared == 1
        remaining = image_service.list_images(SESSION_A)
        assert len(remaining) == 1
        assert remaining[0]["format"] == "JPEG"

    def test_clear_all(self, image_service, sample_png, sample_jpeg):
        image_service.register_image(sample_png, session_key=SESSION_A)
        image_service.register_image(sample_jpeg, session_key=SESSION_A)

        cleared = image_service.clear_images(session_key=SESSION_A)
        assert cleared == 2
        assert len(image_service.list_images(SESSION_A)) == 0

    def test_clear_deletes_file(self, image_service, sample_png):
        info = image_service.register_image(sample_png, session_key=SESSION_A)
        file_path = image_service._images_dir.parent / info["ref"]
        assert file_path.exists()

        image_service.clear_images(session_key=SESSION_A)
        assert not file_path.exists()

    def test_set_active_images_limits_workflow_refs(self, image_service, sample_png):
        old_info = image_service.register_image(sample_png, session_key=SESSION_A)
        new_info = image_service.register_image(sample_png, session_key=SESSION_A)

        active = image_service.set_active_images(
            [new_info["ref"]],
            session_key=SESSION_A,
        )

        assert [info["ref"] for info in active] == [new_info["ref"]]
        assert [info["ref"] for info in image_service.list_images(SESSION_A)] == [
            old_info["ref"],
            new_info["ref"],
        ]
        assert [
            info["ref"] for info in image_service.list_active_images(SESSION_A)
        ] == [new_info["ref"]]

    def test_clear_removes_refs_from_active_set(self, image_service, sample_png):
        old_info = image_service.register_image(sample_png, session_key=SESSION_A)
        new_info = image_service.register_image(sample_png, session_key=SESSION_A)
        image_service.set_active_images(
            [old_info["ref"], new_info["ref"]],
            session_key=SESSION_A,
        )

        image_service.clear_images(session_key=SESSION_A, ref=old_info["ref"])

        assert [
            info["ref"] for info in image_service.list_active_images(SESSION_A)
        ] == [new_info["ref"]]


class TestImageAssetServicePersistence:
    def test_index_survives_reload(self, tmp_path, sample_png):
        service1 = ImageAssetService(plugin_data_path=tmp_path)
        info = service1.register_image(
            sample_png, session_key=SESSION_A, note="persist"
        )

        service2 = ImageAssetService(plugin_data_path=tmp_path)
        images = service2.list_images(SESSION_A)
        assert len(images) == 1
        assert images[0]["ref"] == info["ref"]
        assert images[0]["note"] == "persist"

    def test_active_set_survives_reload(self, tmp_path, sample_png):
        service1 = ImageAssetService(plugin_data_path=tmp_path)
        info = service1.register_image(sample_png, session_key=SESSION_A)
        service1.set_active_images([info["ref"]], session_key=SESSION_A)

        service2 = ImageAssetService(plugin_data_path=tmp_path)

        active_images = service2.list_active_images(SESSION_A)
        assert len(active_images) == 1
        assert active_images[0]["ref"] == info["ref"]


class TestImageBlockValidation:
    def test_image_block_rejects_non_asset_path(self):
        from astrbot_plugin_office_assistant.document_core.models.blocks import (
            ImageBlock,
        )

        with pytest.raises(Exception, match="images/"):
            ImageBlock(path="some/other/path.png")

    def test_image_block_accepts_asset_ref(self):
        from astrbot_plugin_office_assistant.document_core.models.blocks import (
            ImageBlock,
        )

        block = ImageBlock(path="images/img_20260521_abc12345.png")
        assert block.path == "images/img_20260521_abc12345.png"

    def test_image_slide_block_rejects_non_asset_path(self):
        from astrbot_plugin_office_assistant.document_core.models.blocks import (
            ImageSlideBlock,
        )

        with pytest.raises(Exception, match="images/"):
            ImageSlideBlock(image_path="photo.jpg", title="test")

    def test_image_slide_block_accepts_asset_ref(self):
        from astrbot_plugin_office_assistant.document_core.models.blocks import (
            ImageSlideBlock,
        )

        block = ImageSlideBlock(
            image_path="images/img_20260521_def67890.jpg", title="test"
        )
        assert block.image_path == "images/img_20260521_def67890.jpg"
