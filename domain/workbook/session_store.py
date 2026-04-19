from __future__ import annotations

from datetime import datetime, timedelta, timezone
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
    _build_workbook_summary_payload,
    _normalize_xlsx_filename,
    _normalize_xlsx_output_path,
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
    def __init__(
        self,
        workspace_dir: Path | None = None,
        *,
        max_workbooks: int | None = 256,
        ttl: timedelta | None = None,
    ) -> None:
        self._lock = RLock()
        self._workbooks: dict[str, WorkbookModel] = {}
        self._next_workbook_id = 1
        self.workspace_dir = workspace_dir or _default_workspace_dir()
        self.workspace_dir.mkdir(parents=True, exist_ok=True)
        self._max_workbooks = max_workbooks
        self._ttl = ttl

    def _evict_expired_locked(self) -> None:
        if self._ttl is None:
            return

        now = datetime.now(timezone.utc)
        expired_ids = [
            workbook_id
            for workbook_id, workbook in self._workbooks.items()
            if workbook.status != WorkbookStatus.EXPORTING
            if workbook.metadata.updated_at + self._ttl < now
        ]
        for workbook_id in expired_ids:
            self._workbooks.pop(workbook_id, None)

    def _evict_excess_locked(
        self,
        *,
        protected_workbook_ids: set[str] | None = None,
    ) -> None:
        if self._max_workbooks is None:
            return

        excess = len(self._workbooks) - self._max_workbooks
        if excess <= 0:
            return

        protected_ids = protected_workbook_ids or set()
        oldest_workbooks = sorted(
            (
                item
                for item in self._workbooks.items()
                if item[1].status != WorkbookStatus.EXPORTING
                if item[0] not in protected_ids
            ),
            key=lambda item: (
                0 if item[1].status == WorkbookStatus.EXPORTED else 1,
                item[1].metadata.updated_at,
            ),
        )
        for workbook_id, _ in oldest_workbooks[:excess]:
            self._workbooks.pop(workbook_id, None)

    def _prune_locked(
        self,
        *,
        protected_workbook_ids: set[str] | None = None,
    ) -> None:
        self._evict_expired_locked()
        self._evict_excess_locked(protected_workbook_ids=protected_workbook_ids)

    @staticmethod
    def _compact_workbook_after_export_locked(workbook: WorkbookModel) -> None:
        for worksheet in workbook.worksheets:
            worksheet.rows = []

    def _allocate_workbook_id_locked(self) -> str:
        workbook_id = f"wb-{self._next_workbook_id}"
        self._next_workbook_id += 1
        return workbook_id

    def create_workbook(self, request: CreateWorkbookRequest) -> WorkbookModel:
        with self._lock:
            workbook_id = self._allocate_workbook_id_locked()
            preferred_filename = _normalize_xlsx_filename(request.filename)
            title = request.title.strip() or Path(preferred_filename).stem
            workbook = WorkbookModel(
                workbook_id=workbook_id,
                session_id=request.session_id,
                metadata=WorkbookMetadata(
                    title=title,
                    preferred_filename=preferred_filename,
                ),
            )
            self._workbooks[workbook_id] = workbook
            self._prune_locked(protected_workbook_ids={workbook_id})
            return workbook

    def get_workbook(self, workbook_id: str) -> WorkbookModel | None:
        with self._lock:
            self._evict_expired_locked()
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

    def _prepare_export_path_locked(
        self,
        request: ExportWorkbookRequest,
    ) -> tuple[WorkbookModel, Path]:
        workbook = self.require_workbook(request.workbook_id)
        if workbook.status != WorkbookStatus.DRAFT:
            raise ValueError(
                "export_workbook is only allowed while the workbook status is draft"
            )
        preferred_name = request.output_name or workbook.metadata.preferred_filename
        file_name = _normalize_xlsx_output_path(preferred_name)

        workspace_dir = self.workspace_dir.resolve()
        output_dir = workspace_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = (output_dir / file_name).resolve()
        if not _is_within_workspace(output_path, workspace_dir):
            raise ValueError("output_path cannot escape the workbook workspace")
        workbook.output_path = str(output_path)
        workbook.touch()
        return workbook, output_path

    def _reset_failed_export_locked(self, workbook_id: str) -> None:
        workbook = self.require_workbook(workbook_id)
        workbook.status = WorkbookStatus.DRAFT
        workbook.output_path = ""
        workbook.touch()
        self._prune_locked(protected_workbook_ids={workbook_id})

    def prepare_export_path(
        self,
        request: ExportWorkbookRequest,
    ) -> tuple[WorkbookModel, Path]:
        with self._lock:
            return self._prepare_export_path_locked(request)

    def export_workbook(
        self,
        request: ExportWorkbookRequest,
    ) -> tuple[WorkbookModel, Path]:
        with self._lock:
            workbook, output_path = self._prepare_export_path_locked(request)
            workbook.status = WorkbookStatus.EXPORTING
            workbook.touch()
        try:
            export_workbook_to_xlsx(workbook, output_path)
        except Exception:
            with self._lock:
                self._reset_failed_export_locked(request.workbook_id)
            raise
        with self._lock:
            workbook = self.require_workbook(request.workbook_id)
            workbook.status = WorkbookStatus.EXPORTED
            workbook.touch()
            self._compact_workbook_after_export_locked(workbook)
            self._prune_locked()
            return workbook, output_path

    def complete_export(self, workbook_id: str) -> WorkbookModel:
        with self._lock:
            workbook = self.require_workbook(workbook_id)
            workbook.status = WorkbookStatus.EXPORTED
            workbook.touch()
            self._compact_workbook_after_export_locked(workbook)
            self._prune_locked()
            return workbook

    def build_prompt_summary(self, workbook_id: str) -> dict[str, object]:
        with self._lock:
            workbook = self.require_workbook(workbook_id)
            return self._build_prompt_summary_locked(workbook)

    def _build_prompt_summary_locked(self, workbook: WorkbookModel) -> dict[str, object]:
        summary_payload = _build_workbook_summary_payload(workbook)
        next_allowed_actions: list[str]
        if workbook.status == WorkbookStatus.DRAFT:
            next_allowed_actions = ["write_rows", "export_workbook"]
        else:
            next_allowed_actions = []
        prompt_summary = {
            key: summary_payload[key]
            for key in (
                "workbook_id",
                "title",
                "status",
                "sheet_names",
                "sheet_count",
                "latest_written_sheets",
            )
        }
        prompt_summary["next_allowed_actions"] = next_allowed_actions
        return prompt_summary


__all__ = [
    "WorkbookSessionStore",
    "_default_workspace_dir",
    "_is_within_workspace",
]
