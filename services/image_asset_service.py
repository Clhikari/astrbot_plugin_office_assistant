from __future__ import annotations

import json
import shutil
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TypedDict

from astrbot.api import logger

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


class ImageAssetService:
    def __init__(self, *, plugin_data_path: Path) -> None:
        self._images_dir = plugin_data_path / "images"
        self._images_dir.mkdir(parents=True, exist_ok=True)
        self._index_path = self._images_dir / "index.json"
        self._index: list[ImageAssetInfo] = self._load_index()

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

    def register_image(
        self,
        source_path: Path,
        *,
        session_key: tuple[str, str, str],
        note: str = "",
        original_name: str = "",
    ) -> ImageAssetInfo:
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
        self._save_index()
        return info

    def list_images(self, session_key: tuple[str, str, str]) -> list[ImageAssetInfo]:
        sk = list(session_key)
        return [item for item in self._index if item["session_key"] == sk]

    def update_note(
        self,
        ref: str,
        note: str,
        *,
        session_key: tuple[str, str, str],
    ) -> bool:
        sk = list(session_key)
        for item in self._index:
            if item["ref"] == ref and item["session_key"] == sk:
                item["note"] = note
                self._save_index()
                return True
        return False

    def clear_images(
        self,
        *,
        session_key: tuple[str, str, str],
        ref: str | None = None,
    ) -> int:
        sk = list(session_key)
        to_remove: list[ImageAssetInfo] = []
        kept: list[ImageAssetInfo] = []

        for item in self._index:
            if item["session_key"] != sk:
                kept.append(item)
                continue
            if ref is None or item["ref"] == ref:
                to_remove.append(item)
            else:
                kept.append(item)

        for item in to_remove:
            file_path = self._images_dir.parent / item["ref"]
            if file_path.exists():
                try:
                    file_path.unlink()
                except OSError:
                    logger.warning(f"failed to delete image file: {file_path}")

        self._index = kept
        self._save_index()
        return len(to_remove)

    def resolve_ref(
        self,
        ref: str,
        *,
        session_key: tuple[str, str, str],
    ) -> Path:
        if not ref.startswith("images/"):
            raise ValueError(f"无效的图片引用: {ref}。引用必须以 images/ 开头。")
        if ".." in ref.replace("\\", "/").split("/"):
            raise ValueError("图片引用不允许包含目录遍历 (..)")

        sk = list(session_key)
        found = any(
            item["ref"] == ref and item["session_key"] == sk for item in self._index
        )
        if not found:
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
