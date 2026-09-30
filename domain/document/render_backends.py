from __future__ import annotations

import asyncio
import json
import math
import os
import shutil
import subprocess
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator, Literal, Protocol, Sequence

from astrbot.api import logger

from ...document_core.models.document import DocumentModel

RenderableFormat = Literal["word", "ppt", "excel"]
RenderBackendKind = Literal["python", "node"]

_STORE_RENDER_BACKEND_CONFIG_ATTR = "_document_render_backend_config"
_LEGACY_STORE_RENDER_BACKEND_CONFIG_ATTR = "_legacy_document_render_backend_config"
DEFAULT_NODE_RENDER_TIMEOUT_SECONDS = 120.0


@dataclass(slots=True)
class RenderResult:
    backend_name: str
    output_path: Path


@dataclass(slots=True)
class DocumentRenderBackendConfig:
    preferred_backend: RenderBackendKind = "node"
    fallback_enabled: bool = True
    node_renderer_entry: str = ""
    ppt_preferred_backend: RenderBackendKind = "node"
    ppt_fallback_enabled: bool = False
    excel_preferred_backend: RenderBackendKind = "python"
    excel_fallback_enabled: bool = False

    def preferred_backend_for(
        self, document_format: RenderableFormat
    ) -> RenderBackendKind:
        if document_format == "ppt":
            return self.ppt_preferred_backend
        if document_format == "excel":
            return self.excel_preferred_backend
        return self.preferred_backend

    def fallback_enabled_for(self, document_format: RenderableFormat) -> bool:
        if document_format == "ppt":
            return self.ppt_fallback_enabled
        if document_format == "excel":
            return self.excel_fallback_enabled
        return self.fallback_enabled

    @property
    def js_renderer_entry(self) -> str:
        return self.node_renderer_entry


class DocumentRenderBackend(Protocol):
    name: str

    def render(self, document: DocumentModel, output_path: Path) -> RenderResult: ...


class DocumentRenderBackendError(RuntimeError):
    def __init__(self, backend_name: str, message: str):
        super().__init__(message)
        self.backend_name = backend_name


def attach_render_backend_config(
    store: object,
    config: DocumentRenderBackendConfig | None,
) -> None:
    setattr(store, _STORE_RENDER_BACKEND_CONFIG_ATTR, config)
    setattr(store, _LEGACY_STORE_RENDER_BACKEND_CONFIG_ATTR, config)


def get_render_backend_config(
    store: object,
) -> DocumentRenderBackendConfig | None:
    config = getattr(store, _STORE_RENDER_BACKEND_CONFIG_ATTR, None)
    if isinstance(config, DocumentRenderBackendConfig):
        return config
    legacy_config = getattr(store, _LEGACY_STORE_RENDER_BACKEND_CONFIG_ATTR, None)
    if isinstance(legacy_config, DocumentRenderBackendConfig):
        return legacy_config
    return None


def build_document_render_payload(document: DocumentModel) -> dict[str, Any]:
    metadata = document.metadata.model_dump(mode="json")
    metadata["document_style"] = document.metadata.document_style.model_dump(
        mode="json",
        exclude_unset=True,
    )
    metadata["header_footer"] = document.metadata.header_footer.model_dump(
        mode="json",
        exclude_unset=True,
    )

    blocks: list[dict[str, Any]] = []
    for block in document.blocks:
        block_payload = block.model_dump(
            mode="json",
            exclude_unset=True,
            exclude_none=True,
        )
        _fixup_block_payload(block_payload, block)
        blocks.append(block_payload)

    return {
        "version": "v1",
        "render_mode": "structured",
        "document_id": document.document_id,
        "session_id": document.session_id,
        "format": document.format,
        "status": document.status.value,
        "metadata": metadata,
        "blocks": blocks,
    }


def _resolve_document_workspace_dir(document: DocumentModel, output_path: Path) -> Path:
    workspace_dir = str(getattr(document, "_workspace_dir", "") or "").strip()
    if workspace_dir:
        return Path(workspace_dir).resolve()
    return output_path.parent.resolve()


def _fixup_block_payload(block_payload: dict[str, Any], block: object) -> None:
    """Recursively ensure ``type`` is present and ``block_id`` is stripped."""
    block_payload.pop("block_id", None)
    if hasattr(block, "type"):
        block_payload["type"] = block.type
    if getattr(block, "type", "") == "page_template" and hasattr(block, "data"):
        block_payload["data"] = block.data.model_dump(  # type: ignore[union-attr]
            mode="json",
            exclude_none=True,
        )
    if hasattr(block, "blocks") and isinstance(block_payload.get("blocks"), list):
        child_blocks = getattr(block, "blocks", [])
        for idx, child_payload in enumerate(block_payload["blocks"]):
            if isinstance(child_payload, dict) and idx < len(child_blocks):
                _fixup_block_payload(child_payload, child_blocks[idx])
    if hasattr(block, "columns") and isinstance(block_payload.get("columns"), list):
        model_columns = getattr(block, "columns", [])
        for col_idx, col_payload in enumerate(block_payload["columns"]):
            if isinstance(col_payload, dict) and col_idx < len(model_columns):
                col_blocks = getattr(model_columns[col_idx], "blocks", [])
                child_payloads = col_payload.get("blocks", [])
                if isinstance(child_payloads, list):
                    for child_idx, child_payload in enumerate(child_payloads):
                        if isinstance(child_payload, dict) and child_idx < len(
                            col_blocks
                        ):
                            _fixup_block_payload(child_payload, col_blocks[child_idx])


class NodeDocumentRenderBackend:
    name = "node"

    def __init__(
        self,
        entry_path: str | Path | None = None,
        *,
        timeout_seconds: float = DEFAULT_NODE_RENDER_TIMEOUT_SECONDS,
    ) -> None:
        if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
            raise ValueError(
                "Node renderer timeout_seconds must be positive and finite"
            )
        self._timeout_seconds = timeout_seconds
        self._entry_path = (
            Path(entry_path).resolve()
            if entry_path and str(entry_path).strip()
            else self._default_entry_path()
        )

    @staticmethod
    def _default_entry_path() -> Path:
        package_root = Path(__file__).resolve().parents[2]
        return package_root / "word_renderer_js" / "dist" / "cli.js"

    @property
    def entry_path(self) -> Path:
        return self._entry_path

    def is_available(self) -> bool:
        return self._entry_path.exists() and shutil.which("node") is not None

    @contextmanager
    def _payload_file(
        self, document: DocumentModel, output_path: Path
    ) -> Iterator[Path]:
        entry_path = self._entry_path
        if not entry_path.exists():
            raise DocumentRenderBackendError(
                self.name,
                f"Node renderer entry not found: {entry_path}",
            )

        payload = build_document_render_payload(document)
        payload["workspace_dir"] = str(
            _resolve_document_workspace_dir(document, output_path)
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        payload_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                suffix=".json",
                encoding="utf-8",
                delete=False,
                dir=output_path.parent,
            ) as payload_file:
                payload_path = Path(payload_file.name)
                json.dump(payload, payload_file, ensure_ascii=False)
                payload_file.flush()
                os.fsync(payload_file.fileno())
            yield payload_path
        finally:
            if payload_path is not None:
                payload_path.unlink(missing_ok=True)

    def render(self, document: DocumentModel, output_path: Path) -> RenderResult:
        with self._payload_file(document, output_path) as payload_path:
            command = [
                "node",
                str(self._entry_path),
                str(payload_path),
                str(output_path),
            ]
            logger.debug(
                "[office-assistant] invoking js renderer entry=%s payload=%s output=%s",
                self._entry_path,
                payload_path,
                output_path,
            )
            try:
                completed = subprocess.run(
                    command,
                    cwd=str(self._entry_path.parent),
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    timeout=self._timeout_seconds,
                )
            except subprocess.TimeoutExpired as exc:
                raise self._timeout_error() from exc
            except OSError as exc:
                raise DocumentRenderBackendError(
                    self.name, f"Failed to start js renderer: {exc}"
                ) from exc

        return self._render_result(
            output_path, completed.returncode, completed.stdout, completed.stderr
        )

    async def render_async(
        self, document: DocumentModel, output_path: Path
    ) -> RenderResult:
        with self._payload_file(document, output_path) as payload_path:
            try:
                process = await asyncio.create_subprocess_exec(
                    "node",
                    str(self._entry_path),
                    str(payload_path),
                    str(output_path),
                    cwd=str(self._entry_path.parent),
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                )
            except OSError as exc:
                raise DocumentRenderBackendError(
                    self.name, f"Failed to start js renderer: {exc}"
                ) from exc
            try:
                stdout, stderr = await asyncio.wait_for(
                    process.communicate(), timeout=self._timeout_seconds
                )
            except (TimeoutError, asyncio.CancelledError) as exc:
                if process.returncode is None:
                    try:
                        process.kill()
                    except ProcessLookupError:
                        pass
                # Reap the child before removing its payload or partial output.
                await process.communicate()
                if isinstance(exc, asyncio.CancelledError):
                    raise
                raise self._timeout_error() from exc

        return self._render_result(
            output_path,
            process.returncode,
            stdout.decode("utf-8", errors="replace"),
            stderr.decode("utf-8", errors="replace"),
        )

    def _timeout_error(self) -> DocumentRenderBackendError:
        return DocumentRenderBackendError(
            self.name, f"JS renderer timed out after {self._timeout_seconds:g} seconds"
        )

    def _render_result(
        self,
        output_path: Path,
        returncode: int | None,
        stdout: str,
        stderr: str,
    ) -> RenderResult:
        if returncode != 0:
            detail = _extract_renderer_error_detail(
                stderr,
                stdout,
                returncode if returncode is not None else -1,
            )
            raise DocumentRenderBackendError(
                self.name,
                f"JS renderer failed: {detail}",
            )
        if not output_path.exists():
            raise DocumentRenderBackendError(
                self.name,
                f"JS renderer completed without output: {output_path}",
            )
        return RenderResult(backend_name=self.name, output_path=output_path)


class PythonExcelRenderBackend:
    name = "python-excel"

    def render(self, document: DocumentModel, output_path: Path) -> RenderResult:
        raise DocumentRenderBackendError(
            self.name,
            (
                "Excel render backend is reserved for Python implementation, "
                "but the actual exporter is not implemented yet"
            ),
        )


class PythonPptRenderBackend:
    name = "python-ppt"

    def render(self, document: DocumentModel, output_path: Path) -> RenderResult:
        raise DocumentRenderBackendError(
            self.name,
            "PPT render backend is planned for JS and is not implemented in Python",
        )


def _extract_renderer_error_detail(
    stderr: str | None,
    stdout: str | None,
    returncode: int,
) -> str:
    for raw_output in (stderr, stdout):
        text = (raw_output or "").strip()
        if not text:
            continue
        try:
            payload = json.loads(text)
        except json.JSONDecodeError:
            return text
        if isinstance(payload, dict):
            message = payload.get("message")
            if isinstance(message, str) and message.strip():
                return message.strip()
        return text
    return f"exit code {returncode}"


@contextmanager
def _staged_output_path(output_path: Path) -> Iterator[Path]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, filename = tempfile.mkstemp(
        prefix=".office-render-", suffix=output_path.suffix, dir=output_path.parent
    )
    os.close(descriptor)
    staged_path = Path(filename)
    staged_path.unlink()
    try:
        yield staged_path
    finally:
        staged_path.unlink(missing_ok=True)


def _commit_rendered_output(
    staged_path: Path, output_path: Path, result: RenderResult
) -> RenderResult:
    if not staged_path.is_file():
        raise DocumentRenderBackendError(
            result.backend_name, f"Renderer completed without output: {staged_path}"
        )
    staged_path.replace(output_path)
    return RenderResult(result.backend_name, output_path)


def render_document_with_backends(
    document: DocumentModel,
    output_path: Path,
    render_backends: Sequence[DocumentRenderBackend],
) -> RenderResult:
    if not render_backends:
        raise RuntimeError(
            f"No render backend configured for document format: {document.format}"
        )

    last_error: Exception | None = None
    for index, backend in enumerate(render_backends):
        try:
            with _staged_output_path(output_path) as staged_path:
                result = backend.render(document, staged_path)
                result = _commit_rendered_output(staged_path, output_path, result)
            logger.debug(
                "[office-assistant] document render completed document=%s format=%s output=%s backend=%s",
                document.document_id,
                document.format,
                output_path,
                result.backend_name,
            )
            return result
        except Exception as exc:
            last_error = exc
            has_fallback = index < len(render_backends) - 1
            logger.warning(
                "[office-assistant] render backend failed document=%s format=%s backend=%s fallback=%s error=%s",
                document.document_id,
                document.format,
                getattr(backend, "name", backend.__class__.__name__),
                has_fallback,
                exc,
            )
            if not has_fallback:
                raise

    raise RuntimeError(
        f"Rendering failed for document format: {document.format}"
    ) from last_error


async def _render_backend_async(
    backend: DocumentRenderBackend, document: DocumentModel, output_path: Path
) -> RenderResult:
    render_async = getattr(backend, "render_async", None)
    if callable(render_async):
        return await render_async(document, output_path)
    # Keep custom synchronous backends compatible without blocking the event loop.
    task = asyncio.create_task(asyncio.to_thread(backend.render, document, output_path))
    try:
        return await asyncio.shield(task)
    except asyncio.CancelledError:
        # A worker thread cannot be cancelled. Wait before cleaning its staging
        # path, and never publish the result of a cancelled export.
        try:
            await task
        except Exception:
            pass
        raise


async def render_document_with_backends_async(
    document: DocumentModel,
    output_path: Path,
    render_backends: Sequence[DocumentRenderBackend],
) -> RenderResult:
    if not render_backends:
        raise RuntimeError(
            f"No render backend configured for document format: {document.format}"
        )
    for index, backend in enumerate(render_backends):
        try:
            with _staged_output_path(output_path) as staged_path:
                result = await _render_backend_async(backend, document, staged_path)
                return _commit_rendered_output(staged_path, output_path, result)
        except Exception as exc:
            has_fallback = index < len(render_backends) - 1
            logger.warning(
                "[office-assistant] async render backend failed document=%s format=%s backend=%s fallback=%s error=%s",
                document.document_id,
                document.format,
                getattr(backend, "name", backend.__class__.__name__),
                has_fallback,
                exc,
            )
            if not has_fallback:
                raise
    raise RuntimeError(f"Rendering failed for document format: {document.format}")


def build_document_render_backends(
    document_format: RenderableFormat,
    config: DocumentRenderBackendConfig | None = None,
) -> list[DocumentRenderBackend]:
    resolved = config or DocumentRenderBackendConfig()

    if document_format == "word":
        return [
            NodeDocumentRenderBackend(entry_path=resolved.js_renderer_entry or None)
        ]

    if document_format == "ppt":
        if resolved.preferred_backend_for("ppt") == "python":
            return [PythonPptRenderBackend()]
        node_backend = NodeDocumentRenderBackend(
            entry_path=resolved.js_renderer_entry or None
        )
        if resolved.fallback_enabled_for("ppt"):
            return [node_backend, PythonPptRenderBackend()]
        return [node_backend]

    if document_format == "excel":
        return [PythonExcelRenderBackend()]

    raise ValueError(f"Unsupported document format: {document_format}")


__all__ = [
    "RenderableFormat",
    "DocumentRenderBackend",
    "DocumentRenderBackendConfig",
    "DocumentRenderBackendError",
    "NodeDocumentRenderBackend",
    "PythonExcelRenderBackend",
    "PythonPptRenderBackend",
    "RenderBackendKind",
    "RenderResult",
    "attach_render_backend_config",
    "build_document_render_backends",
    "build_document_render_payload",
    "render_document_with_backends",
    "render_document_with_backends_async",
    "get_render_backend_config",
]
