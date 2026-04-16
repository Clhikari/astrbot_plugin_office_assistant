from __future__ import annotations

from pathlib import Path
from threading import RLock

try:
    from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path
except ModuleNotFoundError:  # pragma: no cover - test fallback
    def get_astrbot_plugin_data_path() -> str:
        return str(Path.cwd() / ".tmp_services")

from .contracts import (
    CreateWorkbookRequest,
    ExportWorkbookRequest,
    WriteRowsRequest,
    _normalize_xlsx_filename,
)
from .exporter import export_workbook_to_xlsx
from .models import WorkbookMetadata, WorkbookModel, WorkbookStatus, WorksheetModel

PLUGIN_NAME = "astrbot_plugin_office_assistant"


def _default_workspace_dir() -> Path:
    return Path(get_astrbot_plugin_data_path()) / PLUGIN_NAME / "workbooks"


def _is_within_workspace(path: Path, workspace_dir: Path) -> bool:
    try:
        path.relative_to(workspace_dir)
        return True
    except ValueError:
        return False


class WorkbookSessionStore:
    def __init__(self, workspace_dir: Path | None = None) -> None:
        self._lock = RLock()
        self._workbooks: dict[str, WorkbookModel] = {}
        self._next_workbook_id = 1
        self.workspace_dir = workspace_dir or _default_workspace_dir()
        self.workspace_dir.mkdir(parents=True, exist_ok=True)

    def _allocate_workbook_id_locked(self) -> str:
        workbook_id = f"wb-{self._next_workbook_id}"
        self._next_workbook_id += 1
        return workbook_id

    def create_workbook(self, request: CreateWorkbookRequest) -> WorkbookModel:
        with self._lock:
            workbook_id = self._allocate_workbook_id_locked()
            preferred_filename = _normalize_xlsx_filename(request.filename)
            workbook = WorkbookModel(
                workbook_id=workbook_id,
                session_id=request.session_id,
                metadata=WorkbookMetadata(
                    title=Path(preferred_filename).stem,
                    preferred_filename=preferred_filename,
                ),
            )
            self._workbooks[workbook_id] = workbook
            return workbook

    def get_workbook(self, workbook_id: str) -> WorkbookModel | None:
        with self._lock:
            return self._workbooks.get(workbook_id)

    def require_workbook(self, workbook_id: str) -> WorkbookModel:
        workbook = self.get_workbook(workbook_id)
        if workbook is None:
            raise KeyError(f"Workbook not found: {workbook_id}")
        return workbook

    def write_rows(self, request: WriteRowsRequest) -> WorkbookModel:
        with self._lock:
            workbook = self.require_workbook(request.workbook_id)
            if workbook.status != WorkbookStatus.DRAFT:
                raise ValueError(
                    "write_rows is only allowed while the workbook status is draft"
                )

            worksheet = workbook.get_sheet(request.sheet)
            if worksheet is None:
                worksheet = WorksheetModel(name=request.sheet)
                workbook.worksheets.append(worksheet)

            start_index = request.start_row - 1
            required_size = start_index + len(request.rows)
            while len(worksheet.rows) < required_size:
                worksheet.rows.append([])

            for offset, row in enumerate(request.rows):
                worksheet.rows[start_index + offset] = list(row)

            workbook.remember_written_sheet(worksheet.name)
            return workbook

    def prepare_export_path(
        self,
        request: ExportWorkbookRequest,
    ) -> tuple[WorkbookModel, Path]:
        with self._lock:
            workbook = self.require_workbook(request.workbook_id)
            if workbook.status != WorkbookStatus.DRAFT:
                raise ValueError(
                    "export_workbook is only allowed while the workbook status is draft"
                )
            preferred_name = request.output_name or workbook.metadata.preferred_filename
            file_name = _normalize_xlsx_filename(preferred_name)

            workspace_dir = self.workspace_dir.resolve()
            output_dir = workspace_dir
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = (output_dir / file_name).resolve()
            if not _is_within_workspace(output_path, workspace_dir):
                raise ValueError("output_path cannot escape the workbook workspace")
            workbook.output_path = str(output_path)
            workbook.touch()
            return workbook, output_path

    def export_workbook(
        self,
        request: ExportWorkbookRequest,
    ) -> tuple[WorkbookModel, Path]:
        workbook, output_path = self.prepare_export_path(request)
        export_workbook_to_xlsx(workbook, output_path)
        exported_workbook = self.complete_export(workbook.workbook_id)
        return exported_workbook, output_path

    def complete_export(self, workbook_id: str) -> WorkbookModel:
        with self._lock:
            workbook = self.require_workbook(workbook_id)
            workbook.status = WorkbookStatus.EXPORTED
            workbook.touch()
            return workbook

    def build_prompt_summary(self, workbook_id: str) -> dict[str, object]:
        with self._lock:
            workbook = self.require_workbook(workbook_id)
            return self._build_prompt_summary_locked(workbook)

    def _build_prompt_summary_locked(self, workbook: WorkbookModel) -> dict[str, object]:
        next_allowed_actions: list[str]
        if workbook.status == WorkbookStatus.DRAFT:
            next_allowed_actions = ["write_rows", "export_workbook"]
        else:
            next_allowed_actions = []
        return {
            "workbook_id": workbook.workbook_id,
            "title": workbook.metadata.title,
            "status": workbook.status.value,
            "sheet_names": [worksheet.name for worksheet in workbook.worksheets],
            "sheet_count": len(workbook.worksheets),
            "latest_written_sheets": list(workbook.latest_written_sheets),
            "next_allowed_actions": next_allowed_actions,
        }


__all__ = [
    "WorkbookSessionStore",
    "_default_workspace_dir",
    "_is_within_workspace",
]
