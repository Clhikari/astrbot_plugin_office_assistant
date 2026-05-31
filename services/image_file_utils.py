from pathlib import Path
from urllib.parse import unquote, urlparse

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
