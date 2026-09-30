import asyncio
import json
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from astrbot_plugin_office_assistant.domain.document.contracts import (
    AddBlocksRequest,
    CreateDocumentRequest,
    ExportDocumentRequest,
    execute_add_slides,
)
from astrbot_plugin_office_assistant.domain.document.export_pipeline import (
    export_document_via_pipeline,
)
from astrbot_plugin_office_assistant.domain.document.render_backends import (
    DocumentRenderBackendError,
    NodeDocumentRenderBackend,
    RenderResult,
    build_document_render_payload,
    render_document_with_backends,
)
from astrbot_plugin_office_assistant.domain.document.session_store import (
    DocumentSessionStore,
)
from astrbot_plugin_office_assistant.mcp_server.tools.add_blocks import (
    register_add_blocks_tool,
)


def _draft(tmp_path, *, document_format="word"):
    store = DocumentSessionStore(workspace_dir=tmp_path)
    document = store.create_document(
        CreateDocumentRequest(title="Audit", format=document_format)
    )
    return store, document


def _export(store, document, backend, **options):
    return export_document_via_pipeline(
        store=store,
        render_backends=[backend],
        request=ExportDocumentRequest(document_id=document.document_id, **options),
        source="test",
    )


class _BytesBackend:
    name = "test-bytes"

    def __init__(self, content, *, fail=False):
        self.content = content
        self.fail = fail

    def render(self, document, output_path):
        output_path.write_bytes(self.content)
        if self.fail:
            raise RuntimeError("render failed")
        return RenderResult(self.name, output_path)


@pytest.mark.asyncio
@pytest.mark.parametrize("retry_name", ["", "different.docx"])
async def test_failed_reexport_preserves_previous_file_and_document_state(
    tmp_path, retry_name
):
    store, document = _draft(tmp_path)
    _, previous = await _export(store, document, _BytesBackend(b"previous export"))
    previous_state = document.model_dump()
    with pytest.raises(RuntimeError, match="render failed"):
        await _export(
            store,
            document,
            _BytesBackend(b"partial output", fail=True),
            output_name=retry_name,
        )
    assert previous.read_bytes() == b"previous export"
    assert document.model_dump() == previous_state
    assert list(tmp_path.iterdir()) == [previous]


def test_sync_render_fallback_commits_only_successful_output(tmp_path):
    _, document = _draft(tmp_path)
    output = tmp_path / "report.docx"
    output.write_bytes(b"previous export")

    class _InspectFallback(_BytesBackend):
        def render(self, document, output_path):
            assert output.read_bytes() == b"previous export"
            assert not output_path.exists()
            return super().render(document, output_path)

    result = render_document_with_backends(
        document,
        output,
        [_BytesBackend(b"partial", fail=True), _InspectFallback(b"complete")],
    )
    assert result.output_path == output
    assert output.read_bytes() == b"complete"
    assert list(tmp_path.iterdir()) == [output]


@pytest.mark.asyncio
@pytest.mark.parametrize("same_document", [False, True])
async def test_owned_document_exports_preserve_files_awaiting_delivery(
    tmp_path, same_document
):
    store, alice = _draft(tmp_path)
    alice._owner_key = ("platform", "Alice", "session")
    bob = (
        alice
        if same_document
        else store.create_document(CreateDocumentRequest(title="Bob report"))
    )
    if not same_document:
        bob._owner_key = ("platform", "Bob", "session")
    alice_exported = asyncio.Event()
    resume_delivery = asyncio.Event()

    async def alice_delivery():
        _, path = await _export(
            store,
            alice,
            _BytesBackend(b"Alice's original export"),
            output_dir="reports/q1",
        )
        alice_exported.set()
        await resume_delivery.wait()
        return path, path.read_bytes()

    delivery = asyncio.create_task(alice_delivery())
    await asyncio.wait_for(alice_exported.wait(), timeout=3)
    try:
        _, newer_path = await _export(
            store, bob, _BytesBackend(b"newer export"), output_dir="reports/q1"
        )
    finally:
        resume_delivery.set()
    original_path, delivered_bytes = await delivery
    assert delivered_bytes == b"Alice's original export"
    assert newer_path.read_bytes() == b"newer export"
    assert original_path != newer_path
    assert original_path.name == newer_path.name == "document.docx"
    assert (
        original_path.parent.parent
        == newer_path.parent.parent
        == (tmp_path / "reports" / "q1")
    )


@pytest.mark.asyncio
async def test_failed_owned_export_removes_its_empty_private_directory(tmp_path):
    store, document = _draft(tmp_path)
    document._owner_key = ("platform", "Alice", "session")
    with pytest.raises(RuntimeError, match="render failed"):
        await _export(store, document, _BytesBackend(b"partial", fail=True))
    assert list(tmp_path.iterdir()) == []
    assert document.output_path == ""


def _node_script(tmp_path, body):
    if shutil.which("node") is None:
        pytest.skip("Node is unavailable")
    script = tmp_path / "renderer.cjs"
    script.write_text(body, encoding="utf-8")
    return script


@pytest.mark.asyncio
async def test_node_export_keeps_event_loop_responsive(tmp_path):
    store, document = _draft(tmp_path)
    script = _node_script(
        tmp_path,
        "setTimeout(() => { require('fs').writeFileSync(process.argv[3], 'done'); }, 450);",
    )
    task = asyncio.create_task(
        _export(store, document, NodeDocumentRenderBackend(script))
    )
    await asyncio.sleep(0.05)
    was_running = not task.done()
    _, output = await task
    assert was_running, "Node rendering blocked the event loop until it completed"
    assert output.read_text() == "done"


@pytest.mark.asyncio
@pytest.mark.parametrize("stop", ["timeout", "cancel"])
async def test_node_export_stop_reaps_process_and_preserves_output(
    tmp_path, monkeypatch, stop
):
    store, document = _draft(tmp_path)
    output = tmp_path / "document.docx"
    output.write_bytes(b"previous")
    script = _node_script(
        tmp_path,
        "require('fs').writeFileSync(process.argv[3], 'partial'); setInterval(() => {}, 1000);",
    )
    processes = []
    staged_paths = []
    started = asyncio.Event()
    real_spawn = asyncio.create_subprocess_exec

    async def spawn(*args, **kwargs):
        process = await real_spawn(*args, **kwargs)
        processes.append(process)
        staged_paths.append(Path(args[3]))
        started.set()
        return process

    monkeypatch.setattr(asyncio, "create_subprocess_exec", spawn)
    backend = NodeDocumentRenderBackend(
        script, timeout_seconds=0.15 if stop == "timeout" else 3
    )
    task = asyncio.create_task(_export(store, document, backend))
    await asyncio.wait_for(started.wait(), timeout=3)
    if stop == "cancel":

        async def wait_for_partial_output():
            while not staged_paths[0].exists():
                await asyncio.sleep(0.01)

        await asyncio.wait_for(wait_for_partial_output(), timeout=2)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
    else:
        with pytest.raises(DocumentRenderBackendError, match="timed out"):
            await task
    assert processes[0].returncode is not None
    assert output.read_bytes() == b"previous"
    assert document.status.value == "draft"
    assert document.output_path == ""
    assert sorted(p.name for p in tmp_path.iterdir()) == [
        "document.docx",
        "renderer.cjs",
    ]
    # Cancellation and timeout must also release the document's export guard.
    await _export(store, document, _BytesBackend(b"retried"))
    assert output.read_bytes() == b"retried"


@pytest.mark.asyncio
async def test_document_rejects_mutation_and_second_export_while_rendering(tmp_path):
    store, document = _draft(tmp_path)
    started = asyncio.Event()
    finish = asyncio.Event()

    class _PausedBackend(_BytesBackend):
        async def render_async(self, document, output_path):
            started.set()
            await finish.wait()
            return self.render(document, output_path)

    backend = _PausedBackend(b"complete")
    task = asyncio.create_task(_export(store, document, backend))
    try:
        await asyncio.wait_for(started.wait(), timeout=2)
        with pytest.raises(ValueError, match="export.*in progress"):
            await _export(store, document, backend)
        with pytest.raises(ValueError, match="export.*in progress"):
            store.add_blocks(
                AddBlocksRequest(
                    document_id=document.document_id,
                    blocks=[{"type": "paragraph", "text": "racing edit"}],
                )
            )
    finally:
        finish.set()
        await task
    assert document.status.value == "exported"


def test_mcp_add_blocks_failure_is_atomic_and_retry_does_not_duplicate(tmp_path):
    store, document = _draft(tmp_path)

    class _Server:
        def tool(self, **kwargs):
            def register(function):
                self.add_blocks = function
                return function

            return register

    server = _Server()
    register_add_blocks_tool(server, store)
    blocks = [
        {"type": "paragraph", "text": "Only once"},
        {"type": "image", "path": "images/missing.png"},
    ]
    before = document.model_dump()
    for _ in range(2):
        with pytest.raises(ValueError, match="图片文件不存在"):
            server.add_blocks(document.document_id, blocks)
        assert document.model_dump() == before
    (tmp_path / "images").mkdir()
    (tmp_path / "images" / "missing.png").write_bytes(b"present")
    server.add_blocks(document.document_id, blocks)
    assert len(document.blocks) == 2
    assert document.blocks[0].text == "Only once"


def _table_slide(*, wide=True):
    return {
        "type": "table_slide",
        "headers": ["NAME", "VALUE"],
        "rows": [["Alice", "42", "must not disappear"] if wide else ["Alice"]],
    }


def test_add_slides_rejects_wide_rows_without_mutating_document(tmp_path):
    store, document = _draft(tmp_path, document_format="ppt")
    result = execute_add_slides(store, document.document_id, [_table_slide()])
    assert not result.success
    assert "row 1" in result.message and "2" in result.message
    assert document.blocks == []


@pytest.mark.parametrize("wide", [True, False])
def test_ppt_cli_rejects_wide_rows_and_preserves_short_rows(tmp_path, wide):
    if shutil.which("node") is None:
        pytest.skip("Node is unavailable")
    entry = NodeDocumentRenderBackend().entry_path
    if not entry.exists():
        pytest.skip("Build the Node renderer before running renderer tests")
    _, document = _draft(tmp_path, document_format="ppt")
    payload = build_document_render_payload(document)
    payload["blocks"] = [_table_slide(wide=wide)]
    source = tmp_path / "payload.json"
    output = tmp_path / "result.pptx"
    source.write_text(json.dumps(payload), encoding="utf-8")
    result = subprocess.run(
        ["node", str(entry), str(source), str(output)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=10,
    )
    if wide:
        assert result.returncode != 0
        assert json.loads(result.stderr)["code"] == "TABLE_ROW_TOO_WIDE"
        assert not output.exists()
    else:
        assert result.returncode == 0, result.stderr
        with zipfile.ZipFile(output) as archive:
            slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        texts = [
            element.text
            for element in slide.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}t"
            )
        ]
        assert "Alice" in texts
