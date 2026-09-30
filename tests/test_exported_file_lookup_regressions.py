import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from astrbot_plugin_office_assistant.domain.document.contracts import (
    CreateDocumentRequest,
    ExportDocumentRequest,
)
from astrbot_plugin_office_assistant.domain.document.session_store import (
    DocumentSessionStore,
)
from astrbot_plugin_office_assistant.domain.workbook.contracts import (
    CreateWorkbookRequest,
    ExportWorkbookRequest,
    WriteRowsRequest,
)
from astrbot_plugin_office_assistant.domain.workbook.session_store import (
    WorkbookSessionStore,
)
from astrbot_plugin_office_assistant.services.command_service import CommandService
from astrbot_plugin_office_assistant.services.workspace_service import WorkspaceService
from astrbot_plugin_office_assistant.services import runtime_builder
from astrbot_plugin_office_assistant.domain.export_artifacts import (
    EXPORT_OWNER_FILENAME,
    cleanup_owned_export_directory,
    list_owned_export_paths,
    write_owned_export_metadata,
)


def _event(user="alice"):
    return SimpleNamespace(
        get_sender_id=lambda: user,
        get_platform_id=lambda: "audit",
        unified_msg_origin="group",
    )


def _export_document(store, user="alice", name="report.docx"):
    document = store.create_document(
        CreateDocumentRequest(title="Report", output_name=name)
    )
    document._owner_key = ("audit", user, "group")
    _, path = store.prepare_export_path(
        ExportDocumentRequest(document_id=document.document_id)
    )
    path.write_bytes(b"generated document")
    store.complete_export(document.document_id, path)
    return path


def _export_workbook(store, user="alice", name="report.xlsx"):
    workbook = store.create_workbook(CreateWorkbookRequest(filename=name))
    workbook._owner_key = ("audit", user, "group")
    store.write_rows(
        WriteRowsRequest(workbook_id=workbook.workbook_id, sheet="Data", rows=[[user]])
    )
    _, path = store.export_workbook(
        ExportWorkbookRequest(workbook_id=workbook.workbook_id)
    )
    return path


def _workspace(root):
    workspace = WorkspaceService(
        plugin_data_path=root, executor=None, office_libs={}, max_file_size=1_000_000
    )
    workspace.set_exported_paths_lookup(
        runtime_builder._build_exported_paths_lookup(workspace_dir=root)
    )
    return workspace


def _command(workspace, root=None):
    command = object.__new__(CommandService)
    command._workspace_service = workspace
    command._plugin_data_path = root
    command._auto_delete = False
    command._require_access = lambda event: None
    return command


def _check(workspace, filename, user="alice", **kwargs):
    return workspace.pre_check(
        _event(user),
        filename,
        require_exists=kwargs.pop("require_exists", True),
        is_group_feature_enabled=lambda event: True,
        check_permission_fn=lambda event: True,
        group_feature_disabled_error=lambda: "disabled",
        **kwargs,
    )


@pytest.mark.parametrize("kind", ["document", "workbook"])
def test_export_list_is_scoped_and_skips_deleted_files(tmp_path, kind):
    store = (
        DocumentSessionStore(tmp_path)
        if kind == "document"
        else WorkbookSessionStore(tmp_path)
    )
    export = _export_document if kind == "document" else _export_workbook
    own = export(store)
    other = export(store, "bob")
    workspace = _workspace(tmp_path)

    assert workspace.list_exported_paths(_event()) == [own]
    assert workspace.list_exported_paths(_event("bob")) == [other]
    own.unlink()
    assert workspace.list_exported_paths(_event()) == []


def test_list_and_basename_lookup_survive_store_rebuild_without_cross_user_leak(
    tmp_path,
):
    documents = DocumentSessionStore(tmp_path)
    workbooks = WorkbookSessionStore(tmp_path)
    own_document = _export_document(documents)
    own_workbook = _export_workbook(workbooks)
    other = _export_workbook(workbooks, "bob", "private.xlsx")
    del documents, workbooks  # Recovery must depend only on persisted owner records.
    workspace = _workspace(tmp_path)
    command = _command(workspace, tmp_path)

    listing = command.list_files(_event())

    assert own_document.relative_to(tmp_path).as_posix() in listing
    assert own_workbook.relative_to(tmp_path).as_posix() in listing
    assert other.relative_to(tmp_path).as_posix() not in listing
    assert _check(workspace, own_document.name)[:2] == (True, own_document)
    assert _check(workspace, own_workbook.name)[:2] == (True, own_workbook)
    assert _check(workspace, other.name)[0] is False


def test_existing_root_file_wins_and_duplicate_owned_basenames_require_path(tmp_path):
    workbooks = WorkbookSessionStore(tmp_path)
    first = _export_workbook(workbooks)
    second = _export_workbook(workbooks)
    workspace = _workspace(tmp_path)

    ok, _, error = _check(workspace, "report.xlsx")
    assert ok is False
    assert "多个" in error
    assert first.relative_to(tmp_path).as_posix() in error
    assert second.relative_to(tmp_path).as_posix() in error
    assert _check(workspace, first.relative_to(tmp_path).as_posix())[:2] == (
        True,
        first,
    )
    root_file = tmp_path / "report.xlsx"
    root_file.write_bytes(b"existing root file")
    assert _check(workspace, "report.xlsx")[:2] == (True, root_file)


@pytest.mark.parametrize(
    "filename,options,allowed",
    [
        ("../report.xlsx", {}, False),
        ("sub/report.xlsx", {}, False),
        ("./report.xlsx", {}, False),
        ("sub\\report.xlsx", {}, False),
        ("C:report.xlsx", {}, False),
        (Path("report.xlsx"), {}, False),
        (Path("../missing-report.xlsx"), {"allow_external_path": True}, False),
        ("report.xlsx", {"require_exists": False}, True),
    ],
)
def test_non_bare_and_creation_paths_never_use_export_aliases(
    tmp_path, filename, options, allowed
):
    workspace = _workspace(tmp_path)
    lookup = MagicMock(return_value=[])
    workspace.set_exported_paths_lookup(lookup)
    # Path cases are absolute references; strings preserve their supplied spelling.
    reference = str(tmp_path / filename) if isinstance(filename, Path) else filename
    result = _check(workspace, reference, **options)
    assert result[0] is allowed
    if allowed:
        assert result[1] == tmp_path / "report.xlsx"
    lookup.assert_not_called()


def test_lookup_deduplicates_results_and_rejects_external_paths(tmp_path):
    workbooks = WorkbookSessionStore(tmp_path)
    own = _export_workbook(workbooks)
    workspace = _workspace(tmp_path)

    assert workspace.list_exported_paths(_event()) == [own]
    workspace.set_exported_paths_lookup(
        lambda event: [own, own, Path(__file__).resolve()]
    )
    assert workspace.list_exported_paths(_event()) == [own]
    assert _check(workspace, own.name)[:2] == (True, own)
    workspace.set_exported_paths_lookup(lambda event: [Path(__file__).resolve()])
    assert workspace.list_exported_paths(_event()) == []


@pytest.mark.parametrize("absolute", [False, True])
def test_known_private_path_read_and_delete_require_owner_after_restart(
    tmp_path, absolute
):
    workbooks = WorkbookSessionStore(tmp_path)
    output = _export_workbook(workbooks, "bob", "private.xlsx")
    workspace = _workspace(tmp_path)
    reference = str(output) if absolute else output.relative_to(tmp_path).as_posix()
    command = _command(workspace)

    ok, _, error = _check(workspace, reference)
    assert ok is False
    assert "权限不足" in error
    assert "权限不足" in command.delete_file(_event(), f"/delete_file {reference}")
    assert output.exists()
    assert _check(workspace, reference, user="bob")[:2] == (True, output)
    assert "已删除" in command.delete_file(_event("bob"), f"/delete_file {reference}")
    assert not output.exists()
    assert not output.parent.exists()


def test_export_metadata_contains_only_digest_and_is_published_after_success(tmp_path):
    store = DocumentSessionStore(tmp_path)
    document = store.create_document(CreateDocumentRequest(title="Report"))
    document._owner_key = ("audit-platform", "alice-private-id", "private-origin")
    _, output = store.prepare_export_path(
        ExportDocumentRequest(document_id=document.document_id)
    )
    sidecar = output.parent / EXPORT_OWNER_FILENAME
    assert not sidecar.exists()
    with pytest.raises(ValueError, match="existing private workspace file"):
        write_owned_export_metadata(output, document._owner_key, tmp_path)
    assert not sidecar.exists()
    output.write_bytes(b"finished file")

    store.complete_export(document.document_id, output)

    record_text = sidecar.read_text(encoding="utf-8")
    record = json.loads(record_text)
    assert set(record) == {"version", "owner_sha256", "basename"}
    assert len(record["owner_sha256"]) == 64
    assert record["basename"] == output.name
    assert all(part not in record_text for part in document._owner_key)
    assert not list(output.parent.glob("*.tmp"))


@pytest.mark.parametrize(
    "corruption",
    [
        "traversal",
        "absolute",
        "oversized",
        "invalid",
        "version",
        "extra",
        "missing_file",
    ],
)
def test_bad_export_metadata_is_ignored_without_hiding_other_files(
    tmp_path, corruption
):
    store = WorkbookSessionStore(tmp_path)
    damaged = _export_workbook(store, name="damaged.xlsx")
    healthy = _export_workbook(store, name="healthy.xlsx")
    sidecar = damaged.parent / EXPORT_OWNER_FILENAME
    record = json.loads(sidecar.read_text(encoding="utf-8"))
    if corruption == "traversal":
        record["basename"] = "../healthy.xlsx"
    elif corruption == "absolute":
        record["basename"] = str(healthy)
    elif corruption == "version":
        record["version"] = True
    elif corruption == "extra":
        record["owner_key"] = ["audit", "alice", "group"]
    elif corruption == "missing_file":
        damaged.unlink()
    if corruption == "oversized":
        sidecar.write_bytes(b" " * 5000)
    elif corruption == "invalid":
        sidecar.write_text("broken", encoding="utf-8")
    else:
        sidecar.write_text(json.dumps(record), encoding="utf-8")

    assert list_owned_export_paths(tmp_path, ("audit", "alice", "group")) == [healthy]


def test_cleanup_owned_metadata_preserves_live_outputs_and_unrelated_files(tmp_path):
    store = WorkbookSessionStore(tmp_path)
    output = _export_workbook(store)
    sidecar = output.parent / EXPORT_OWNER_FILENAME
    cleanup_owned_export_directory(output)
    assert sidecar.exists() and output.exists()
    sibling = output.parent / "keep.txt"
    sibling.write_text("unrelated", encoding="utf-8")
    output.unlink()

    cleanup_owned_export_directory(output)

    assert not sidecar.exists()
    assert sibling.exists()
    sibling.unlink()
    cleanup_owned_export_directory(output)
    assert not output.parent.exists()
