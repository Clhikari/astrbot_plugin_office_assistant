from __future__ import annotations

import asyncio
import json
import shutil
import subprocess
import sys
import tempfile
import textwrap
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from uuid import uuid4

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    EXCEL_SCRIPT_RETRY_EXHAUSTED_EVENT_KEY,
    EXCEL_SCRIPT_RETRY_FAILURES_EVENT_KEY,
    EXCEL_SUFFIXES,
)
from .excel_script_templates import (
    build_cleanup_script,
    build_prepare_script,
    build_runner_script,
    build_script_helper_template,
)
from .runtime_config import (
    SUPPORTED_COMPUTER_RUNTIME_MODES,
    resolve_computer_runtime_mode,
)

try:
    from astrbot.core.computer.computer_client import (
        get_booter as _get_computer_booter,
    )
except ImportError:  # pragma: no cover - depends on AstrBot runtime
    _get_computer_booter = None


@dataclass(frozen=True, slots=True)
class ScriptProcessResult:
    success: bool
    mode: str
    result_text: str | None = None
    output_path: Path | None = None
    error: str | None = None
    traceback: str | None = None
    script: str | None = None
    retryable: bool = True


@dataclass(frozen=True, slots=True)
class SandboxScriptPaths:
    exec_dir: str
    result_path: str
    remote_result_path: str
    input_files: list[str]
    remote_input_files: list[str]
    output_path: str | None = None
    remote_output_path: str | None = None


class ExcelScriptService:
    _EXCEL_SUFFIXES = EXCEL_SUFFIXES
    _MAX_SCRIPT_RETRIES = 3
    _RETRY_COUNT_EVENT_KEY = EXCEL_SCRIPT_RETRY_FAILURES_EVENT_KEY
    _SCRIPT_TIMEOUT_SECONDS = 30
    _SANDBOX_EXEC_ROOT = PurePosixPath(".office_assistant") / "excel_scripts"

    def __init__(
        self,
        *,
        astrbot_context=None,
        auto_block_execution_tools: bool = False,
        workspace_service,
        file_delivery_service,
        allow_external_input_files: bool,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._astrbot_context = astrbot_context
        self._auto_block_execution_tools = auto_block_execution_tools
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
        user_message: str | None = None,
    ) -> str:
        payload = {
            "success": False,
            "error": error,
            "traceback": traceback_text,
            "script": script,
            "retry_count": retry_count,
            "max_retries": ExcelScriptService._MAX_SCRIPT_RETRIES,
            "retry_exhausted": retry_exhausted,
        }
        if user_message:
            payload["user_message"] = user_message
        return json.dumps(payload, ensure_ascii=False)

    @classmethod
    def _normalize_script_text(cls, script: str) -> str:
        normalized = textwrap.dedent(script or "").strip()
        try:
            compile(normalized, "<excel-script>", "exec")
        except SyntaxError:
            repaired = normalized.replace('\\"', '"').replace("\\'", "'")
            if repaired != normalized:
                try:
                    compile(repaired, "<excel-script>", "exec")
                except SyntaxError:
                    pass
                else:
                    return repaired
        return normalized

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
        cls._set_retry_exhausted(event, False)

    @classmethod
    def _set_retry_exhausted(cls, event: AstrMessageEvent, exhausted: bool) -> None:
        set_extra = getattr(event, "set_extra", None)
        if callable(set_extra):
            set_extra(EXCEL_SCRIPT_RETRY_EXHAUSTED_EVENT_KEY, exhausted)

    @classmethod
    def _retry_used_count(cls, failure_count: int) -> int:
        return max(failure_count - 1, 0)

    @classmethod
    def _build_retry_exhausted_error(cls, error: str) -> str:
        return (
            f"{error}；已超过最多 {cls._MAX_SCRIPT_RETRIES} 次脚本重试，"
            "请停止调用工具，直接向用户说明最后一次失败原因。"
        )

    @classmethod
    def _build_retry_exhausted_user_message(cls, error: str) -> str:
        return (
            f"Excel 脚本已经达到最多 {cls._MAX_SCRIPT_RETRIES} 次重试，"
            "本次没有生成合格文件。\n\n"
            f"最后一次失败：{error}"
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
            user_message=cls._build_retry_exhausted_user_message(error),
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
            self._set_retry_exhausted(event, True)
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

    def _build_non_retry_result(
        self,
        event: AstrMessageEvent,
        *,
        script: str,
        error: str,
        traceback_text: str = "",
        retry_exhausted: bool = False,
    ) -> str:
        failure_count = self._get_failure_count(event)
        return self._build_error_result(
            script=script,
            error=error,
            traceback_text=traceback_text,
            retry_count=self._retry_used_count(failure_count),
            retry_exhausted=retry_exhausted,
        )

    async def execute_excel_script(
        self,
        event: AstrMessageEvent,
        *,
        script: str,
        input_files: list[str] | None = None,
        output_name: str | None = None,
    ) -> str:
        normalized_script = self._normalize_script_text(script)
        current_failure_count = self._get_failure_count(event)
        if current_failure_count > self._MAX_SCRIPT_RETRIES:
            self._set_retry_exhausted(event, True)
            return self._build_exhausted_result(
                script=normalized_script or script or "",
                error="脚本重试次数已用尽",
                failure_count=current_failure_count,
            )
        if not normalized_script:
            return self._build_non_retry_result(
                event,
                script=script or "",
                error="缺少 script 参数",
            )

        if not self._is_group_feature_enabled(event):
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error=self._group_feature_disabled_error(),
                retry_exhausted=True,
            )

        if not self._check_permission(event):
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error="错误：权限不足",
                retry_exhausted=True,
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
                return self._build_non_retry_result(
                    event,
                    script=normalized_script,
                    error=err or "错误：未知错误",
                )
            if resolved_path is None:
                return self._build_non_retry_result(
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
                return self._build_non_retry_result(
                    event,
                    script=normalized_script,
                    error=err or "错误：未知错误",
                )
            if resolved_path is None:
                return self._build_non_retry_result(
                    event,
                    script=normalized_script,
                    error="错误：输出文件路径解析失败",
                )
            resolved_output_path = resolved_path
            resolved_output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            runtime_mode = self._resolve_runtime_mode(event)
        except Exception as exc:
            logger.warning(f"[文件管理] 读取 Excel runtime 配置失败: {exc}")
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error="错误：读取当前会话 computer runtime 失败",
                traceback_text=str(exc),
                retry_exhausted=True,
            )
        if runtime_mode not in SUPPORTED_COMPUTER_RUNTIME_MODES:
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error=(
                    f"错误：不支持的 computer runtime 配置：{runtime_mode}，"
                    "请使用 local、sandbox 或 none"
                ),
                retry_exhausted=True,
            )
        if runtime_mode == "none":
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error=(
                    "错误：当前 computer runtime 为 none，无法执行 Excel 脚本，"
                    "请启用 local 或 sandbox，或改走原语路径"
                ),
                retry_exhausted=True,
            )
        if self._auto_block_execution_tools and runtime_mode != "sandbox":
            return self._build_non_retry_result(
                event,
                script=normalized_script,
                error=(
                    "错误：当前已启用执行类工具自动屏蔽，"
                    "execute_excel_script 仅允许在 sandbox runtime 下执行"
                ),
                retry_exhausted=True,
            )

        booter = None
        if runtime_mode == "sandbox":
            booter, booter_error = await self._acquire_sandbox_booter(
                event,
                runtime_mode=runtime_mode,
            )
            if booter is None:
                return self._build_non_retry_result(
                    event,
                    script=normalized_script,
                    error=booter_error or "错误：Excel sandbox 不可用",
                    retry_exhausted=True,
                )

        process_result = await self._run_script_process(
            script=normalized_script,
            input_files=resolved_input_files,
            output_path=resolved_output_path,
            runtime_mode=runtime_mode,
            booter=booter,
        )

        if not process_result.success:
            if getattr(process_result, "retryable", True):
                return self._build_failure_result(
                    event,
                    script=process_result.script or normalized_script,
                    error=process_result.error or "脚本执行失败",
                    traceback_text=process_result.traceback or "",
                )
            return self._build_non_retry_result(
                event,
                script=process_result.script or normalized_script,
                error=process_result.error or "脚本执行失败",
                traceback_text=process_result.traceback or "",
                retry_exhausted=True,
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

        (
            delivery_error,
            delivery_result,
        ) = await self._file_delivery_service.deliver_generated_file_with_result(
            event,
            process_result.output_path,
            missing_message="错误：Excel 脚本执行成功，但未找到生成的文件",
            oversized_template=(
                "错误：生成的 Excel 文件大小 {file_size} 超过限制 {max_size}"
            ),
            block_quality_warnings=True,
            quality_warning_input_paths=resolved_input_files,
        )
        if delivery_error:
            return self._build_failure_result(
                event,
                script=normalized_script,
                error=delivery_error,
            )

        self._reset_failure_count(event)
        success_payload = {
            "success": True,
            "mode": "file",
            "output_name": process_result.output_path.name
            if process_result.output_path is not None
            else "",
        }
        if delivery_result.quality_summary is not None:
            success_payload["quality_summary"] = delivery_result.quality_summary
            quality_warnings = list(delivery_result.quality_summary.get("warnings", []))
            if quality_warnings:
                success_payload["requires_review"] = True
                success_payload["quality_warnings"] = quality_warnings
                success_payload["message"] = (
                    "文件已生成并发送，但质量摘要存在警告；如果能根据 "
                    "quality_warnings 修正，继续调用 execute_excel_script 生成新版本；"
                    "无法修正时，回复用户必须逐条列出 quality_warnings 中的每一条，"
                    "不要遗漏，也不要表述为完全完成。"
                )
        return json.dumps(
            success_payload,
            ensure_ascii=False,
        )

    def _resolve_runtime_mode(self, event: AstrMessageEvent) -> str:
        return resolve_computer_runtime_mode(self._astrbot_context, event)

    async def _acquire_sandbox_booter(
        self,
        event: AstrMessageEvent,
        *,
        runtime_mode: str | None = None,
    ) -> tuple[object | None, str | None]:
        if self._astrbot_context is None:
            return None, "错误：当前服务未注入 AstrBot context，无法使用 sandbox"
        if _get_computer_booter is None:
            return None, "错误：当前 AstrBot 版本未提供 sandbox 执行接口"
        if runtime_mode is None:
            try:
                runtime_mode = self._resolve_runtime_mode(event)
            except Exception as exc:
                logger.warning(f"[文件管理] 读取 Excel runtime 配置失败: {exc}")
                return None, "错误：读取当前会话 computer runtime 失败"
        if runtime_mode != "sandbox":
            return (
                None,
                "错误：execute_excel_script 仅支持 sandbox runtime，请在 AstrBot 中启用 sandbox",
            )
        session_id = str(getattr(event, "unified_msg_origin", "") or "").strip()
        if not session_id:
            return None, "错误：缺少会话标识，无法初始化 Excel sandbox"
        try:
            booter = await _get_computer_booter(self._astrbot_context, session_id)
        except Exception as exc:
            logger.warning(f"[文件管理] Excel sandbox 初始化失败: {exc}")
            return None, f"错误：Excel sandbox 初始化失败：{exc}"

        capabilities = getattr(booter, "capabilities", None)
        if capabilities is not None:
            required_capabilities = {"python", "filesystem"}
            if isinstance(capabilities, dict):
                missing = sorted(
                    capability
                    for capability in required_capabilities
                    if not capabilities.get(capability)
                )
            else:
                missing = sorted(required_capabilities.difference(capabilities))
            if missing:
                missing_text = ", ".join(missing)
                return (
                    None,
                    f"错误：当前 sandbox profile 缺少必要能力：{missing_text}",
                )
        return booter, None

    @classmethod
    def _build_sandbox_paths(
        cls,
        _booter,
        *,
        input_files: list[Path],
        output_path: Path | None,
    ) -> SandboxScriptPaths:
        exec_relative_dir = cls._SANDBOX_EXEC_ROOT / uuid4().hex
        runtime_input_files: list[str] = []
        remote_input_files: list[str] = []
        for index, original_path in enumerate(input_files):
            relative_input_path = (
                exec_relative_dir / "_inputs" / str(index) / original_path.name
            )
            runtime_input_files.append(
                relative_input_path.relative_to(exec_relative_dir).as_posix()
            )
            remote_input_files.append(relative_input_path.as_posix())

        relative_result_path = exec_relative_dir / "result.json"
        runtime_result_path = relative_result_path.as_posix()

        runtime_output_path: str | None = None
        remote_output_path: str | None = None
        if output_path is not None:
            relative_output_path = exec_relative_dir / "_output" / "output.xlsx"
            runtime_output_path = relative_output_path.as_posix()
            remote_output_path = runtime_output_path

        return SandboxScriptPaths(
            exec_dir=exec_relative_dir.as_posix(),
            result_path=runtime_result_path,
            remote_result_path=relative_result_path.as_posix(),
            input_files=runtime_input_files,
            remote_input_files=remote_input_files,
            output_path=runtime_output_path,
            remote_output_path=remote_output_path,
        )

    @staticmethod
    def _build_prepare_script(file_paths: list[str]) -> str:
        return build_prepare_script(file_paths)

    @staticmethod
    def _build_cleanup_script(directory_path: str) -> str:
        return build_cleanup_script(directory_path)

    @staticmethod
    def _build_script_helper_template() -> str:
        return build_script_helper_template()

    @staticmethod
    def _format_exec_traceback(exec_result: dict) -> str:
        parts: list[str] = []
        error_text = str(exec_result.get("error", "") or "").strip()
        if error_text:
            parts.append(error_text)
        output_text = str(exec_result.get("output", "") or "").strip()
        if not output_text:
            data = exec_result.get("data")
            if isinstance(data, dict):
                output_payload = data.get("output")
                if isinstance(output_payload, dict):
                    output_text = str(output_payload.get("text", "") or "").strip()
        if output_text and output_text not in parts:
            parts.append(output_text)
        return "\n".join(parts).strip()

    async def _cleanup_sandbox_exec_dir(self, booter, exec_dir: str) -> None:
        try:
            await booter.python.exec(
                self._build_cleanup_script(exec_dir),
                timeout=10,
                silent=True,
            )
        except (
            Exception
        ) as exc:  # pragma: no cover - cleanup failure should not mask result
            logger.warning(f"[文件管理] 清理 Excel sandbox 临时目录失败: {exc}")

    async def _run_script_process(
        self,
        *,
        script: str,
        input_files: list[Path],
        output_path: Path | None,
        runtime_mode: str,
        booter=None,
    ) -> ScriptProcessResult:
        if runtime_mode == "sandbox":
            if booter is None:
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel sandbox 不可用",
                    traceback="Missing sandbox booter",
                    script=script,
                    retryable=False,
                )
            return await self._run_sandbox_script_process(
                booter=booter,
                script=script,
                input_files=input_files,
                output_path=output_path,
            )
        if runtime_mode != "local":
            return ScriptProcessResult(
                success=False,
                mode="error",
                error=(
                    f"错误：不支持的 computer runtime 配置：{runtime_mode}，"
                    "请使用 local、sandbox 或 none"
                ),
                traceback="",
                script=script,
                retryable=False,
            )
        return await asyncio.to_thread(
            self._run_local_script_process,
            script=script,
            input_files=input_files,
            output_path=output_path,
        )

    async def _run_sandbox_script_process(
        self,
        *,
        booter,
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
            sandbox_paths = self._build_sandbox_paths(
                booter,
                input_files=input_files,
                output_path=output_path,
            )
            result_path = temp_root / "result.json"
            try:
                try:
                    prepare_result = await booter.python.exec(
                        self._build_prepare_script(
                            [
                                path
                                for path in [
                                    *sandbox_paths.remote_input_files,
                                    sandbox_paths.result_path,
                                    sandbox_paths.output_path,
                                ]
                                if path
                            ]
                        ),
                        timeout=10,
                        silent=True,
                    )
                except Exception as exc:
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel sandbox 初始化失败",
                        traceback=str(exc),
                        script=script,
                    )
                if not prepare_result.get("success", False):
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel sandbox 初始化失败",
                        traceback=self._format_exec_traceback(prepare_result),
                        script=script,
                    )

                for local_path, remote_path in zip(
                    input_files,
                    sandbox_paths.remote_input_files,
                    strict=True,
                ):
                    try:
                        upload_result = await booter.upload_file(
                            str(local_path),
                            remote_path,
                        )
                    except Exception as exc:
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="Excel 输入文件上传失败",
                            traceback=str(exc),
                            script=script,
                        )
                    if not upload_result.get("success", False):
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="Excel 输入文件上传失败",
                            traceback=str(upload_result.get("message", "") or ""),
                            script=script,
                            retryable=False,
                        )

                try:
                    exec_result = await booter.python.exec(
                        self._build_runner_script(
                            script=script,
                            exec_dir=sandbox_paths.exec_dir,
                            input_files=sandbox_paths.input_files,
                            output_path=sandbox_paths.output_path,
                            result_path=sandbox_paths.result_path,
                        ),
                        timeout=self._SCRIPT_TIMEOUT_SECONDS,
                        silent=True,
                    )
                except Exception as exc:
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel 脚本执行失败",
                        traceback=str(exc),
                        script=script,
                    )
                if not exec_result.get("success", False):
                    traceback_text = self._format_exec_traceback(exec_result)
                    error_text = "Excel 脚本执行失败"
                    if "timeout" in traceback_text.lower():
                        error_text = "Excel 脚本执行超时"
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error=error_text,
                        traceback=traceback_text,
                        script=script,
                    )

                try:
                    await booter.download_file(
                        sandbox_paths.remote_result_path,
                        str(result_path),
                    )
                except Exception as exc:
                    traceback_text = str(exc)
                    exec_traceback = self._format_exec_traceback(exec_result)
                    if exec_traceback:
                        traceback_text = f"{traceback_text}\n{exec_traceback}".strip()
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel 脚本未返回结果",
                        traceback=traceback_text,
                        script=script,
                    )

                try:
                    payload = json.loads(result_path.read_text(encoding="utf-8"))
                except (json.JSONDecodeError, OSError) as exc:
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel 脚本结果解析失败",
                        traceback=str(exc),
                        script=script,
                    )
                if not isinstance(payload, dict):
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="Excel 脚本结果解析失败",
                        traceback=f"Unexpected JSON payload type: {type(payload).__name__}",
                        script=script,
                    )
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
                    output_path_value = payload.get("output_path")
                    if (
                        not isinstance(output_path_value, str)
                        or not output_path_value.strip()
                    ):
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="脚本返回了 file 模式，但缺少 output_path",
                            traceback="",
                            script=script,
                        )
                    if output_path is None or sandbox_paths.output_path is None:
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="脚本返回了 file 模式，但当前请求未提供 output_path",
                            traceback="",
                            script=script,
                        )
                    if output_path_value != sandbox_paths.output_path:
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="脚本返回了无效的 output_path",
                            traceback=f"Unexpected output_path: {output_path_value}",
                            script=script,
                        )
                    try:
                        await booter.download_file(
                            sandbox_paths.remote_output_path or "",
                            str(output_path),
                        )
                    except FileNotFoundError:
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="脚本返回了 file 模式，但 output_path 不存在",
                            traceback="",
                            script=script,
                        )
                    except Exception as exc:
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="Excel 脚本结果下载失败",
                            traceback=str(exc),
                            script=script,
                        )
                    if not output_path.exists():
                        return ScriptProcessResult(
                            success=False,
                            mode="error",
                            error="脚本返回了 file 模式，但 output_path 不存在",
                            traceback="",
                            script=script,
                        )
                    return ScriptProcessResult(
                        success=True,
                        mode="file",
                        output_path=output_path,
                    )
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error=f"未知脚本返回模式: {mode}",
                    traceback="",
                    script=script,
                )
            finally:
                await self._cleanup_sandbox_exec_dir(booter, sandbox_paths.exec_dir)

    def _run_local_script_process(
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
            copied_input_files: list[Path] = []
            inputs_root = temp_root / "_inputs"
            inputs_root.mkdir()
            for index, original_path in enumerate(input_files):
                copied_parent = inputs_root / str(index)
                copied_parent.mkdir()
                copied_path = copied_parent / original_path.name
                shutil.copy2(original_path, copied_path)
                copied_input_files.append(copied_path)

            runner_path = temp_root / "runner.py"
            result_path = temp_root / "result.json"
            runner_path.write_text(
                self._build_runner_script(
                    script=script,
                    exec_dir=str(temp_root.resolve()),
                    input_files=[
                        path.relative_to(temp_root).as_posix()
                        for path in copied_input_files
                    ],
                    output_path=(
                        str(output_path.resolve()) if output_path is not None else None
                    ),
                    result_path=str(result_path.resolve()),
                ),
                encoding="utf-8",
            )
            try:
                completed = subprocess.run(
                    [sys.executable, str(runner_path)],
                    cwd=str(workspace_root),
                    capture_output=True,
                    text=True,
                    shell=False,
                    timeout=self._SCRIPT_TIMEOUT_SECONDS,
                )
            except subprocess.TimeoutExpired as exc:
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本执行超时",
                    traceback=str(exc),
                    script=script,
                )
            except OSError as exc:
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本执行失败",
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

            try:
                payload = json.loads(result_path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError) as exc:
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本结果解析失败",
                    traceback=str(exc),
                    script=script,
                )
            if not isinstance(payload, dict):
                return ScriptProcessResult(
                    success=False,
                    mode="error",
                    error="Excel 脚本结果解析失败",
                    traceback=f"Unexpected JSON payload type: {type(payload).__name__}",
                    script=script,
                )
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
                output_path_value = payload.get("output_path")
                if (
                    not isinstance(output_path_value, str)
                    or not output_path_value.strip()
                ):
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="脚本返回了 file 模式，但缺少 output_path",
                        traceback="",
                        script=script,
                    )
                if output_path is None:
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="脚本返回了 file 模式，但当前请求未提供 output_path",
                        traceback="",
                        script=script,
                    )
                expected_output_path = str(output_path.resolve())
                if output_path_value != expected_output_path:
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="脚本返回了无效的 output_path",
                        traceback=f"Unexpected output_path: {output_path_value}",
                        script=script,
                    )
                if not output_path.exists():
                    return ScriptProcessResult(
                        success=False,
                        mode="error",
                        error="脚本返回了 file 模式，但 output_path 不存在",
                        traceback="",
                        script=script,
                    )
                return ScriptProcessResult(
                    success=True,
                    mode="file",
                    output_path=output_path,
                )
            return ScriptProcessResult(
                success=False,
                mode="error",
                error=f"未知脚本返回模式: {mode}",
                traceback="",
                script=script,
            )

    @classmethod
    def _build_runner_script(
        cls,
        *,
        script: str,
        exec_dir: str,
        input_files: list[str],
        output_path: str | None,
        result_path: str,
    ) -> str:
        return build_runner_script(
            script=script,
            exec_dir=exec_dir,
            input_files=input_files,
            output_path=output_path,
            result_path=result_path,
            helper_script=cls._build_script_helper_template(),
        )


__all__ = ["ExcelScriptService", "ScriptProcessResult"]
