import importlib.util
import shutil
import sys
from collections.abc import Iterator
from pathlib import Path
from uuid import uuid4

import pytest


def _ensure_local_package_alias() -> None:
    package_name = "astrbot_plugin_office_assistant"
    if package_name in sys.modules:
        return

    project_root = Path(__file__).resolve().parent.parent
    spec = importlib.util.spec_from_file_location(
        package_name,
        project_root / "__init__.py",
        submodule_search_locations=[str(project_root)],
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法为 {package_name} 创建本地包映射")

    module = importlib.util.module_from_spec(spec)
    sys.modules[package_name] = module
    spec.loader.exec_module(module)


_ensure_local_package_alias()


def build_notice_once_callback():
    seen: dict[tuple[str, str, str], set[str]] = {}

    def consume(event, notice_key: str) -> bool:
        session_key = (
            str(event.get_platform_id() or ""),
            str(event.get_sender_id() or ""),
            str(event.unified_msg_origin or ""),
        )
        used = seen.setdefault(session_key, set())
        if notice_key in used:
            return False
        used.add(notice_key)
        return True

    return consume


@pytest.fixture
def workspace_root() -> Iterator[Path]:
    workspace_base = Path(__file__).resolve().parent / ".tmp_test_workspaces"
    workspace_base.mkdir(parents=True, exist_ok=True)
    workspace_dir = workspace_base / f"workspace-root-{uuid4().hex}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    try:
        yield workspace_dir
    finally:
        shutil.rmtree(workspace_dir, ignore_errors=True)
