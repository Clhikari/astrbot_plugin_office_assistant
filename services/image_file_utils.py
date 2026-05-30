from pathlib import Path

IMAGE_FILE_SUFFIXES = frozenset({".png", ".jpg", ".jpeg", ".webp"})


def is_supported_image_filename(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_FILE_SUFFIXES
