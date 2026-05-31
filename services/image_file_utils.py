from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

import astrbot.api.message_components as Comp

IMAGE_FILE_SUFFIXES = frozenset({".png", ".jpg", ".jpeg", ".webp"})
IMAGE_MIME_TYPES = frozenset({"image/png", "image/jpeg", "image/webp"})


def is_supported_image_filename(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_FILE_SUFFIXES


def is_supported_image_reference(value: str) -> bool:
    candidate = str(value or "").strip()
    if not candidate:
        return False
    parsed = urlparse(candidate)
    path_value = parsed.path if parsed.scheme and parsed.path else candidate
    return is_supported_image_filename(unquote(path_value))


def is_supported_image_mime_type(mime_type: str) -> bool:
    return (
        str(mime_type or "").split(";", maxsplit=1)[0].strip().lower()
        in IMAGE_MIME_TYPES
    )


def is_image_file_component(component: Any) -> bool:
    if not isinstance(component, Comp.File):
        return False
    for attr_name in ("name", "file", "file_", "url"):
        value = getattr(component, attr_name, "") or ""
        if value and is_supported_image_reference(str(value)):
            return True
    for attr_name in ("mime_type", "mimetype", "mime", "content_type"):
        value = getattr(component, attr_name, "") or ""
        if value and is_supported_image_mime_type(str(value)):
            return True
    return False
