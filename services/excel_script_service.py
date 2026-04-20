from __future__ import annotations

import asyncio
import json
import subprocess
import sys
import tempfile
import textwrap
from dataclasses import dataclass
from pathlib import Path

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent


@dataclass(frozen=True, slots=True)
class ScriptProcessResult:
    success: bool
    mode: str
    result_text: str | None = None
    output_path: Path | None = None
    error: str | None = None
    traceback: str | None = None
    script: str | None = None


class ExcelScriptService:
    _EXCEL_SUFFIXES = frozenset({".xlsx", ".xls"})
    _MAX_SCRIPT_RETRIES = 2
    _RETRY_COUNT_EVENT_KEY = "office_assistant_excel_script_retry_failures"

    def __init__(
        self,
        *,
        workspace_service,
        file_delivery_service,
        allow_external_input_files: bool,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._workspace_service = workspace_service
        self._file_delivery_service = file_delivery_service
        self._allow_external_input_files = allow_external_input_files
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    @staticmethod
    def _build_error_result(
        *,
        script: str,
        error: str,
        traceback_text: str = "",
        retry_count: int = 0,
        retry_exhausted: bool = False,
    ) -> str:
        return json.dumps(
            {
                "success": False,
                "error": error,
                "traceback": traceback_text,
                "script": script,
                "retry_count": retry_count,
                "max_retries": ExcelScriptService._MAX_SCRIPT_RETRIES,
                "retry_exhausted": retry_exhausted,
            },
            ensure_ascii=False,
        )

    @classmethod
    def _get_failure_count(cls, event: AstrMessageEvent) -> int:
        get_extra = getattr(event, "get_extra", None)
        if not callable(get_extra):
            return 0
        value = get_extra(cls._RETRY_COUNT_EVENT_KEY, 0)
        try:
            return max(int(value), 0)
        except (TypeError, ValueError):
            return 0

    @classmethod
    def _set_failure_count(cls, event: AstrMessageEvent, count: int) -> None:
        set_extra = getattr(event, "set_extra", None)
        if callable(set_extra):
            set_extra(cls._RETRY_COUNT_EVENT_KEY, max(count, 0))

    @classmethod
    def _reset_failure_count(cls, event: AstrMessageEvent) -> None:
        cls._set_failure_count(event, 0)

    @classmethod
    def _retry_used_count(cls, failure_count: int) -> int:
        return max(failure_count - 1, 0)

    @classmethod
    def _build_retry_exhausted_error(cls, error: str) -> str:
        return (
            f"{error}；已超过最多 {cls._MAX_SCRIPT_RETRIES} 次脚本重试，请停止重试，"
            "直接向用户说明失败原因，并建议缩小需求或改走原语路径。"
        )

    @classmethod
    def _build_exhausted_result(
        cls,
        *,
        script: str,
        error: str,
        traceback_text: str = "",
        failure_count: int,
    ) -> str:
        return cls._build_error_result(
            script=script,
            error=cls._build_retry_exhausted_error(error),
            traceback_text=traceback_text,
            retry_count=cls._retry_used_count(failure_count),
            retry_exhausted=True,
        )

    def _build_failure_result(
        self,
        event: AstrMessageEvent,
        *,
        script: str,
        error: str,
        traceback_text: str = "",
    ) -> str:
        failure_count = self._get_failure_count(event) + 1
        self._set_failure_count(event, failure_count)
        if failure_count > self._MAX_SCRIPT_RETRIES:
            return self._build_exhausted_result(
                script=script,
                error=error,
                traceback_text=traceback_text,
                failure_count=failure_count,
            )
        return self._build_error_result(
            script=script,
            error=error,
            traceback_text=traceback_text,
            retry_count=self._retry_used_count(failure_count),
            retry_exhausted=False,
        )

    async def execute_excel_script(
        self,
        event: AstrMessageEvent,
        *,
        script: str,
        input_files: list[str] | None = None,
        output_name: str | None = None,
    ) -> str:
        normalized_script = (script or "").strip()
        current_failure_count = self._get_failure_count(event)
        if current_failure_count > self._MAX_SCRIPT_RETRIES:
            return self._build_exhausted_result(
                script=normalized_script or script or "",
                error="脚本重试次数已用尽",
                failure_count=current_failure_count,
            )
        if not normalized_script:
            return self._build_failure_result(
                event,
                script=script or "",
                error="缺少 script 参数",
            )

        if not self._is_group_feature_enabled(event):
            return self._build_failure_result(
                event,
                script=normalized_script,
                error=self._group_feature_disabled_error(),
            )

        if not self._check_permission(event):
            return self._build_failure_result(
                event,
                script=normalized_script,
                error="错误：权限不足",
            )

        resolved_input_files: list[Path] = []
        for filename in input_files or []:
            ok, resolved_path, err = self._workspace_service.pre_check(
                event,
                filename,
                require_exists=True,
                allowed_suffixes=self._EXCEL_SUFFIXES,
                allow_external_path=self._allow_external_input_files,
                is_group_feature_enabled=self._is_group_feature_enabled,
                check_permission_fn=self._check_permission,
                group_feature_disabled_error=self._group_feature_disabled_error,
            )
            if not ok:
                return self._build_failure_result(
                    event,
                    script=normalized_script,
                    error=err or "错误：未知错误",
                )
            if resolved_path is None:
                return self._build_failure_result(
                    event,
                    script=normalized_script,
                    error="错误：输入文件路径解析失败",
                )
            resolved_input_files.append(resolved_path)

        resolved_output_path: Path | None = None
        if output_name:
            ok, resolved_path, err = self._workspace_service.pre_check(
                event,
                output_name,
                require_exists=False,
                allowed_suffixes=self._EXCEL_SUFFIXES,
                allow_external_path=False,
                is_group_feature_enabled=self._is_group_feature_enabled,
                check_permission_fn=self._check_permission,
                group_feature_disabled_error=self._group_feature_disabled_error,
            )
            if not ok:
                return self._build_failure_result(
                    event,
                    script=normalized_script,
                    error=err or "错误：未知错误",
                )
            if resolved_path is None:
                return self._build_failure_result(
                    event,
                    script=normalized_script,
                    error="错误：输出文件路径解析失败",
                )
            resolved_output_path = resolved_path
            resolved_output_path.parent.mkdir(parents=True, exist_ok=True)

        process_result = await asyncio.to_thread(
            self._run_script_process,
            script=normalized_script,
            input_files=resolved_input_files,
            output_path=resolved_output_path,
        )

        if not process_result.success:
            return self._build_failure_result(
                event,
                script=process_result.script or normalized_script,
                error=process_result.error or "脚本执行失败",
                traceback_text=process_result.traceback or "",
            )

        if process_result.mode == "text":
            self._reset_failure_count(event)
            return json.dumps(
                {
                    "success": True,
                    "mode": "text",
                    "result_text": process_result.result_text or "",
                },
                ensure_ascii=False,
            )

        delivery_error = await self._file_delivery_service.deliver_generated_file(
            event,
            process_result.output_path,
            missing_message="错误：Excel 脚本执行成功，但未找到生成的文件",
            oversized_template=(
                "错误：生成的 Excel 文件大小 {file_size} 超过限制 {max_size}"
            ),
        )
        if delivery_error:
            return self._build_failure_result(
                event,
                script=normalized_script,
                error=delivery_error,
            )

        self._reset_failure_count(event)
        return json.dumps(
            {
                "success": True,
                "mode": "file",
                "output_name": process_result.output_path.name
                if process_result.output_path is not None
                else "",
            },
            ensure_ascii=False,
        )

    def _run_script_process(
        self,
        *,
        script: str,
        input_files: list[Path],
        output_path: Path | None,
    ) -> ScriptProcessResult:
        workspace_root = self._workspace_service.plugin_data_path.resolve()
        with tempfile.TemporaryDirectory(
            prefix="excel_script_",
            dir=str(workspace_root),
        ) as temp_dir:
            temp_root = Path(temp_dir)
            runner_path = temp_root / "runner.py"
            result_path = temp_root / "result.json"
            runner_path.write_text(
                self._build_runner_script(
                    script=script,
                    input_files=input_files,
                    output_path=output_path,
                    result_path=result_path,
                ),
                encoding="utf-8",
            )
            try:
                completed = subprocess.run(
                    [sys.executable, str(runner_path)],
                    cwd=str(workspace_root),
                    capture_output=True,
                    text=True,
                    timeout=30,
                )
            except subprocess.TimeoutExpired as exc:
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本执行超时",
                    traceback=str(exc),
                    script=script,
                )

            if not result_path.exists():
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本未返回结果",
                    traceback=(completed.stderr or completed.stdout or "").strip(),
                    script=script,
                )

            payload = json.loads(result_path.read_text(encoding="utf-8"))
            if not payload.get("success"):
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error=str(payload.get("error", "脚本执行失败")),
                    traceback=str(payload.get("traceback", "")),
                    script=script,
                )

            mode = str(payload.get("mode", ""))
            if mode == "text":
                return ScriptProcessResult(
                    success=True,
                    mode="text",
                    result_text=str(payload.get("result_text", "")),
                )
            if mode == "file":
                resolved_output_path = (
                    Path(str(payload.get("output_path")))
                    if payload.get("output_path")
                    else output_path
                )
                return ScriptProcessResult(
                    success=True,
                    mode="file",
                    output_path=resolved_output_path,
                )
            return ScriptProcessResult(
                success=False,
                mode="error",
                error=f"未知脚本返回模式: {mode}",
                traceback="",
                script=script,
            )

    @staticmethod
    def _build_runner_script(
        *,
        script: str,
        input_files: list[Path],
        output_path: Path | None,
        result_path: Path,
    ) -> str:
        serialized_input_files = json.dumps(
            [str(path.resolve()) for path in input_files],
            ensure_ascii=False,
        )
        serialized_output_path = json.dumps(
            str(output_path.resolve()) if output_path is not None else None,
            ensure_ascii=False,
        )
        serialized_result_path = json.dumps(str(result_path.resolve()), ensure_ascii=False)
        serialized_script = json.dumps(script, ensure_ascii=False)
        return textwrap.dedent(
            f"""
            import json
            import traceback
            from pathlib import Path

            input_files = [Path(path) for path in json.loads({serialized_input_files!r})]
            output_path_value = json.loads({serialized_output_path!r})
            output_path = Path(output_path_value) if output_path_value else None
            result_path = Path(json.loads({serialized_result_path!r}))
            script = json.loads({serialized_script!r})
            initial_exists = output_path.exists() if output_path is not None else False
            initial_size = output_path.stat().st_size if initial_exists else None
            initial_mtime_ns = output_path.stat().st_mtime_ns if initial_exists else None

            try:
                import openpyxl
                from openpyxl import Workbook, load_workbook

                namespace = {{
                    "__name__": "__main__",
                    "openpyxl": openpyxl,
                    "Workbook": Workbook,
                    "load_workbook": load_workbook,
                    "Path": Path,
                    "input_files": input_files,
                    "output_path": output_path,
                    "result_text": None,
                }}

                exec(script, namespace, namespace)
                result_text = namespace.get("result_text")
                has_text = result_text is not None
                has_file = False
                if output_path is not None and output_path.exists():
                    if not initial_exists:
                        has_file = True
                    else:
                        has_file = (
                            output_path.stat().st_size != initial_size
                            or output_path.stat().st_mtime_ns != initial_mtime_ns
                        )
                if has_text and has_file:
                    payload = {{
                        "success": False,
                        "error": "脚本不能同时设置 result_text 并写出 output_path",
                        "traceback": "",
                    }}
                elif not has_text and not has_file:
                    payload = {{
                        "success": False,
                        "error": "脚本执行完成，但既没有设置 result_text，也没有写出 output_path",
                        "traceback": "",
                    }}
                elif has_text:
                    payload = {{
                        "success": True,
                        "mode": "text",
                        "result_text": str(result_text),
                    }}
                else:
                    payload = {{
                        "success": True,
                        "mode": "file",
                        "output_path": str(output_path),
                    }}
            except Exception as exc:
                payload = {{
                    "success": False,
                    "error": str(exc),
                    "traceback": traceback.format_exc(),
                }}

            result_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            """
        ).strip()


__all__ = ["ExcelScriptService", "ScriptProcessResult"]
