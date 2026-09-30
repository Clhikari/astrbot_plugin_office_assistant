"""Small ownership records for immutable exported files."""

import hashlib
import json
import re
from contextlib import suppress
from pathlib import Path

EXPORT_DIRECTORY_PREFIX = ".office-export-"
EXPORT_OWNER_FILENAME = ".office-owner.json"
MAX_EXPORT_OWNER_BYTES = 4096


def _owner_digest(owner_key: tuple[str, str, str]) -> str:
    encoded = json.dumps(list(owner_key), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _read_owner_record(path: Path) -> dict | None:
    try:
        if path.is_symlink():
            return None
        with path.open("rb") as source:
            raw = source.read(MAX_EXPORT_OWNER_BYTES + 1)
        if len(raw) > MAX_EXPORT_OWNER_BYTES:
            return None
        record = json.loads(raw)
    except (OSError, ValueError, UnicodeError, RecursionError):
        return None
    if not isinstance(record, dict) or set(record) != {
        "version",
        "owner_sha256",
        "basename",
    }:
        return None
    basename = record["basename"]
    digest = record["owner_sha256"]
    if (
        type(record["version"]) is not int
        or record["version"] != 1
        or not isinstance(basename, str)
        or not basename
        or basename in {".", ".."}
        or any(char in basename for char in "/\\:\0")
        or not isinstance(digest, str)
        or re.fullmatch(r"[0-9a-f]{64}", digest) is None
    ):
        return None
    return record


def write_owned_export_metadata(
    output_path: Path,
    owner_key: tuple[str, str, str] | None,
    workspace_dir: Path,
) -> None:
    if owner_key is None:
        return
    output_path = output_path.resolve()
    if (
        not output_path.parent.name.startswith(EXPORT_DIRECTORY_PREFIX)
        or not output_path.is_relative_to(workspace_dir.resolve())
        or not output_path.is_file()
    ):
        raise ValueError(
            "Owned export metadata requires an existing private workspace file"
        )
    metadata_path = output_path.parent / EXPORT_OWNER_FILENAME
    temporary_path = metadata_path.with_suffix(".json.tmp")
    record = {
        "version": 1,
        "owner_sha256": _owner_digest(owner_key),
        "basename": output_path.name,
    }
    try:
        temporary_path.write_text(
            json.dumps(record, ensure_ascii=False), encoding="utf-8"
        )
        temporary_path.replace(metadata_path)
    finally:
        with suppress(OSError):
            temporary_path.unlink(missing_ok=True)


def list_owned_export_paths(
    workspace_dir: Path, owner_key: tuple[str, str, str]
) -> list[Path]:
    workspace_dir = workspace_dir.resolve()
    digest = _owner_digest(owner_key)
    paths: set[Path] = set()
    for metadata_path in workspace_dir.rglob(EXPORT_OWNER_FILENAME):
        parent = metadata_path.parent.resolve()
        if not parent.name.startswith(
            EXPORT_DIRECTORY_PREFIX
        ) or not parent.is_relative_to(workspace_dir):
            continue
        record = _read_owner_record(metadata_path)
        if record is None or record["owner_sha256"] != digest:
            continue
        output_path = (parent / record["basename"]).resolve()
        if output_path.parent == parent and output_path.is_file():
            paths.add(output_path)
    return sorted(paths)


def cleanup_owned_export_directory(output_path: Path) -> None:
    """Remove this deleted artifact's ownership record, then only an empty directory."""
    parent = output_path.parent
    if not parent.name.startswith(EXPORT_DIRECTORY_PREFIX) or output_path.exists():
        return
    metadata_path = parent / EXPORT_OWNER_FILENAME
    record = _read_owner_record(metadata_path)
    if record is not None and record["basename"] == output_path.name:
        with suppress(OSError):
            metadata_path.unlink()
    with suppress(OSError):
        parent.rmdir()
