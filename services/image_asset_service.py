from __future__ import annotations

import json
import shutil
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TypedDict

from astrbot.api import logger

from ..document_core.models.blocks import validate_image_asset_ref

ALLOWED_FORMATS = {"PNG", "JPEG", "WEBP"}
SVG_EXTENSIONS = {".svg", ".svgz"}


class ImageAssetInfo(TypedDict):
    ref: str
    original_name: str
    note: str
    width: int
    height: int
    format: str
    size_bytes: int
    registered_at: float
    session_key: list[str]


class ImageAssetActiveSet(TypedDict):
    session_key: list[str]
    refs: list[str]


class ImageAssetService:
    def __init__(self, *, plugin_data_path: Path) -> None:
        self._images_dir = plugin_data_path / "images"
        self._images_dir.mkdir(parents=True, exist_ok=True)
        self._index_path = self._images_dir / "index.json"
        self._active_path = self._images_dir / "active.json"
        self._index: list[ImageAssetInfo] = self._load_index()
        self._by_session_ref: dict[tuple[str, str, str], dict[str, ImageAssetInfo]] = (
            self._build_lookup(self._index)
        )
        self._active_refs_by_session: dict[tuple[str, str, str], list[str]] = (
            self._load_active_refs()
        )

    def _load_index(self) -> list[ImageAssetInfo]:
        if self._index_path.exists():
            try:
                data = json.loads(self._index_path.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return data
            except (json.JSONDecodeError, OSError):
                logger.warning("image asset index corrupted, starting fresh")
        return []

    def _save_index(self) -> None:
        tmp_path = self._index_path.with_suffix(self._index_path.suffix + ".tmp")
        data = json.dumps(self._index, ensure_ascii=False, indent=2)
        try:
            tmp_path.write_text(data, encoding="utf-8")
            tmp_path.replace(self._index_path)
        except OSError as exc:
            logger.warning("failed to persist image asset index: %s", exc)

    def _load_active_refs(self) -> dict[tuple[str, str, str], list[str]]:
        if not self._active_path.exists():
            return {}
        try:
            data = json.loads(self._active_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            logger.warning("image asset active set corrupted, starting fresh")
            return {}
        if not isinstance(data, list):
            return {}

        active_refs: dict[tuple[str, str, str], list[str]] = {}
        for item in data:
            if not isinstance(item, dict):
                continue
            raw_refs = item.get("refs")
            if not isinstance(raw_refs, list):
                continue
            session_key = self._normalize_session_key(item.get("session_key", []))
            refs = [str(ref) for ref in raw_refs if isinstance(ref, str)]
            if refs:
                active_refs[session_key] = refs
        return active_refs

    def _save_active_refs(self) -> None:
        tmp_path = self._active_path.with_suffix(self._active_path.suffix + ".tmp")
        data: list[ImageAssetActiveSet] = [
            {"session_key": list(session_key), "refs": refs}
            for session_key, refs in self._active_refs_by_session.items()
            if refs
        ]
        try:
            tmp_path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            tmp_path.replace(self._active_path)
        except OSError as exc:
            logger.warning("failed to persist image asset active set: %s", exc)

    @staticmethod
    def _normalize_session_key(
        session_key: tuple[str, str, str] | list[str],
    ) -> tuple[str, str, str]:
        platform_id, sender_id, origin = (list(session_key) + ["", "", ""])[:3]
        return (str(platform_id), str(sender_id), str(origin))

    @classmethod
    def _build_lookup(
        cls,
        items: list[ImageAssetInfo],
    ) -> dict[tuple[str, str, str], dict[str, ImageAssetInfo]]:
        lookup: dict[tuple[str, str, str], dict[str, ImageAssetInfo]] = {}
        for item in items:
            session_key = cls._normalize_session_key(item["session_key"])
            lookup.setdefault(session_key, {})[item["ref"]] = item
        return lookup

    def register_image(
        self,
        source_path: Path,
        *,
        session_key: tuple[str, str, str],
        note: str = "",
        original_name: str = "",
    ) -> ImageAssetInfo:
        session_key = self._normalize_session_key(session_key)
        source_path = Path(source_path)
        if not source_path.exists():
            raise FileNotFoundError(f"源文件不存在: {source_path}")

        suffix = source_path.suffix.lower()
        if suffix in SVG_EXTENSIONS:
            raise ValueError(
                "不支持 SVG 格式。请提供 PNG、JPEG 或 WebP 格式的位图文件。"
            )

        from PIL import Image

        try:
            with Image.open(source_path) as img:
                img.verify()
            with Image.open(source_path) as img:
                pil_format = img.format
                width, height = img.size
        except Exception as exc:
            raise ValueError(f"无法识别为有效图片: {exc}") from exc

        if pil_format not in ALLOWED_FORMATS:
            raise ValueError(
                f"不支持的图片格式: {pil_format}。仅接受 PNG、JPEG、WebP。"
            )

        ext_map = {"PNG": ".png", "JPEG": ".jpg", "WEBP": ".png"}
        real_ext = ext_map[pil_format]
        stored_format = "PNG" if pil_format == "WEBP" else pil_format

        timestamp = datetime.now(timezone.utc).strftime("%Y%m%d")
        short_id = uuid.uuid4().hex[:8]
        filename = f"img_{timestamp}_{short_id}{real_ext}"
        dest_path = self._images_dir / filename

        if pil_format == "WEBP":
            # docx/pptx 无法嵌入 webp，注册时转存为 png 保证文档生成兼容
            with Image.open(source_path) as img:
                img.save(dest_path, format="PNG")
        else:
            shutil.copy2(source_path, dest_path)

        ref = f"images/{filename}"
        info: ImageAssetInfo = {
            "ref": ref,
            "original_name": original_name or "",
            "note": note,
            "width": width,
            "height": height,
            "format": stored_format,
            "size_bytes": dest_path.stat().st_size,
            "registered_at": time.time(),
            "session_key": list(session_key),
        }
        self._index.append(info)
        self._by_session_ref.setdefault(session_key, {})[ref] = info
        self._save_index()
        return info

    def list_images(self, session_key: tuple[str, str, str]) -> list[ImageAssetInfo]:
        session_key = self._normalize_session_key(session_key)
        return list(self._by_session_ref.get(session_key, {}).values())

    def list_active_images(
        self,
        session_key: tuple[str, str, str],
    ) -> list[ImageAssetInfo]:
        session_key = self._normalize_session_key(session_key)
        session_items = self._by_session_ref.get(session_key, {})
        active_refs = self._active_refs_by_session.get(session_key, [])
        active_images = [
            session_items[ref] for ref in active_refs if ref in session_items
        ]
        if len(active_images) != len(active_refs):
            if active_images:
                self._active_refs_by_session[session_key] = [
                    item["ref"] for item in active_images
                ]
            else:
                self._active_refs_by_session.pop(session_key, None)
            self._save_active_refs()
        return active_images

    def set_active_images(
        self,
        refs: list[str],
        *,
        session_key: tuple[str, str, str],
    ) -> list[ImageAssetInfo]:
        session_key = self._normalize_session_key(session_key)
        session_items = self._by_session_ref.get(session_key, {})
        active_refs: list[str] = []
        for ref in refs:
            if ref not in session_items or ref in active_refs:
                continue
            active_refs.append(ref)
        if active_refs:
            self._active_refs_by_session[session_key] = active_refs
        else:
            self._active_refs_by_session.pop(session_key, None)
        self._save_active_refs()
        return [session_items[ref] for ref in active_refs]

    def update_note(
        self,
        ref: str,
        note: str,
        *,
        session_key: tuple[str, str, str],
    ) -> bool:
        session_key = self._normalize_session_key(session_key)
        item = self._by_session_ref.get(session_key, {}).get(ref)
        if item is None:
            return False
        item["note"] = note
        self._save_index()
        return True

    def clear_images(
        self,
        *,
        session_key: tuple[str, str, str],
        ref: str | None = None,
    ) -> int:
        session_key = self._normalize_session_key(session_key)
        session_items = self._by_session_ref.get(session_key, {})
        if ref is None:
            to_remove = list(session_items.values())
        else:
            item = session_items.get(ref)
            to_remove = [item] if item is not None else []
        if not to_remove:
            return 0

        for item in to_remove:
            file_path = self._images_dir.parent / item["ref"]
            if file_path.exists():
                try:
                    file_path.unlink()
                except OSError:
                    logger.warning(f"failed to delete image file: {file_path}")

        removed_refs = {item["ref"] for item in to_remove}
        self._index = [
            item
            for item in self._index
            if not (
                self._normalize_session_key(item["session_key"]) == session_key
                and item["ref"] in removed_refs
            )
        ]
        if ref is None:
            self._by_session_ref.pop(session_key, None)
        else:
            session_items.pop(ref, None)
            if not session_items:
                self._by_session_ref.pop(session_key, None)

        active_refs = self._active_refs_by_session.get(session_key)
        if active_refs is not None:
            remaining_active_refs = [
                active_ref
                for active_ref in active_refs
                if active_ref not in removed_refs
            ]
            if remaining_active_refs:
                self._active_refs_by_session[session_key] = remaining_active_refs
            else:
                self._active_refs_by_session.pop(session_key, None)
        self._save_index()
        self._save_active_refs()
        return len(to_remove)

    def resolve_ref(
        self,
        ref: str,
        *,
        session_key: tuple[str, str, str],
    ) -> Path:
        session_key = self._normalize_session_key(session_key)
        try:
            ref = validate_image_asset_ref(ref)
        except ValueError as exc:
            raise ValueError(f"无效的图片引用: {ref}。{exc}") from exc

        if ref not in self._by_session_ref.get(session_key, {}):
            raise ValueError(
                f"图片引用 {ref} 不在当前会话的资源池中。请先使用 /img add 注册图片。"
            )

        file_path = self._images_dir.parent / ref
        if not file_path.exists():
            raise FileNotFoundError(
                f"图片文件不存在: {ref}。可能已被清理，请重新上传。"
            )
        return file_path.resolve()

    def ref_exists(
        self,
        ref: str,
        *,
        session_key: tuple[str, str, str],
    ) -> bool:
        try:
            self.resolve_ref(ref, session_key=session_key)
            return True
        except (ValueError, FileNotFoundError):
            return False
