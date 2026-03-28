import shutil
from collections.abc import Iterator
from pathlib import Path
from uuid import uuid4

import pytest


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
