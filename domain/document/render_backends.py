from __future__ import annotations

import json
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Protocol, Sequence

from astrbot.api import logger

from ...document_core.models.document import DocumentModel

DocumentFormat = Literal["word", "ppt", "excel"]
RenderBackendKind = Literal["python", "node"]

_STORE_RENDER_BACKEND_CONFIG_ATTR = "_document_render_backend_config"
_LEGACY_STORE_RENDER_BACKEND_CONFIG_ATTR = "_legacy_document_render_backend_config"


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

    def preferred_backend_for(self, document_format: DocumentFormat) -> RenderBackendKind:
        if document_format == "ppt":
            return self.ppt_preferred_backend
        if document_format == "excel":
            return self.excel_preferred_backend
        return self.preferred_backend

    def fallback_enabled_for(self, document_format: DocumentFormat) -> bool:
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
        block_payload = block.model_dump(mode="json", exclude_unset=True)
        if getattr(block, "type", "") == "page_template" and hasattr(block, "data"):
            block_payload["data"] = block.data.model_dump(mode="json")
        block_payload["type"] = block.type
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


class NodeDocumentRenderBackend:
    name = "node"

    def __init__(self, entry_path: str | Path | None = None) -> None:
        self._entry_path = (
            Path(entry_path).resolve()
            if entry_path and str(entry_path).strip()
            else self._default_entry_path()
        )

    @staticmethod
    def _default_entry_path() -> Path:
        package_root = Path(__file__).resolve().parents[2]
        return package_root / "word_renderer_js" / "dist" / "cli.js"

    def render(self, document: DocumentModel, output_path: Path) -> RenderResult:
        entry_path = self._entry_path
        if not entry_path.exists():
            raise DocumentRenderBackendError(
                self.name,
                f"Node renderer entry not found: {entry_path}",
            )

        payload = build_document_render_payload(document)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".json",
            encoding="utf-8",
            delete=False,
            dir=output_path.parent,
        ) as payload_file:
            payload_path = Path(payload_file.name)
            json.dump(payload, payload_file, ensure_ascii=False)

        command = ["node", str(entry_path), str(payload_path), str(output_path)]
        cwd = entry_path.parent
        try:
            logger.debug(
                "[office-assistant] invoking js renderer entry=%s payload=%s output=%s",
                entry_path,
                payload_path,
                output_path,
            )
            completed = subprocess.run(
                command,
                cwd=str(cwd),
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
        except OSError as exc:
            raise DocumentRenderBackendError(
                self.name,
                f"Failed to start js renderer: {exc}",
            ) from exc
        finally:
            payload_path.unlink(missing_ok=True)

        if completed.returncode != 0:
            stderr = (completed.stderr or "").strip()
            stdout = (completed.stdout or "").strip()
            detail = stderr or stdout or f"exit code {completed.returncode}"
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
            output_path.unlink(missing_ok=True)
            result = backend.render(document, output_path)
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
            output_path.unlink(missing_ok=True)
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


def build_document_render_backends(
    document_format: DocumentFormat,
    config: DocumentRenderBackendConfig | None = None,
) -> list[DocumentRenderBackend]:
    resolved = config or DocumentRenderBackendConfig()

    if document_format == "word":
        return [NodeDocumentRenderBackend(entry_path=resolved.js_renderer_entry or None)]

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
    "DocumentFormat",
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
    "get_render_backend_config",
]
