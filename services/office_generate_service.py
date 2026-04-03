from pathlib import Path

from astrbot.api.event import AstrMessageEvent

from ..constants import (
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    OFFICE_LIBS,
    OFFICE_TYPE_MAP,
    SUFFIX_TO_OFFICE_TYPE,
)


class OfficeGenerateService:
    def __init__(
        self,
        *,
        workspace_service,
        office_generator,
        file_delivery_service,
        office_libs: dict,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._workspace_service = workspace_service
        self._office_generator = office_generator
        self._file_delivery_service = file_delivery_service
        self._office_libs = office_libs
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    def _is_explicit_tool_locked(
        self,
        event: AstrMessageEvent,
        tool_name: str,
    ) -> bool:
        get_extra = getattr(event, "get_extra", None)
        if not callable(get_extra):
            return False
        try:
            explicit_tool_name = get_extra(EXPLICIT_FILE_TOOL_EVENT_KEY)
        except TypeError:
            return False
        return isinstance(explicit_tool_name, str) and explicit_tool_name == tool_name

    def _finalize_error(
        self,
        event: AstrMessageEvent,
        message: str,
    ) -> str | None:
        if self._is_explicit_tool_locked(event, "create_office_file"):
            return event.plain_result(message)
        return message

    async def create_office_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        content: str = "",
        file_type: str = "",
    ) -> str | None:
        ok, _, err = self._workspace_service.pre_check(
            event,
            feature_key="enable_office_files",
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            return self._finalize_error(event, err or "错误：未知错误")

        if not content:
            return self._finalize_error(event, "错误：请提供 content（文件内容）")

        filename = Path(filename).name if filename else ""
        if not filename:
            return self._finalize_error(event, "错误：请提供 filename（文件名）")

        allowed_fallback_types = "/".join(
            office_name for office_name in OFFICE_TYPE_MAP if office_name != "word"
        )
        normalized_file_type = str(file_type or "").strip().lower()
        suffix = Path(filename).suffix.lower()
        if suffix in SUFFIX_TO_OFFICE_TYPE:
            office_type = SUFFIX_TO_OFFICE_TYPE[suffix]
        else:
            if not normalized_file_type:
                return self._finalize_error(
                    event,
                    "错误：未指定文件类型。请提供带后缀的文件名，"
                    f"或显式传入 file_type（{allowed_fallback_types}）。",
                )
            if normalized_file_type == "word":
                return self._finalize_error(
                    event,
                    "错误：Word 文档请直接提供 .docx/.doc 文件名，"
                    "或改用 create_document → add_blocks → finalize_document → "
                    "export_document。",
                )
            office_type = OFFICE_TYPE_MAP.get(normalized_file_type)

        if not office_type:
            return self._finalize_error(
                event,
                f"错误：不支持的文件类型 '{normalized_file_type}'。"
                f"允许值：{allowed_fallback_types}",
            )

        module_name = OFFICE_LIBS[office_type][0]
        if not self._office_libs.get(module_name):
            package_name = OFFICE_LIBS[office_type][1]
            return self._finalize_error(event, f"错误：需要安装 {package_name}")

        file_info = {"type": office_type, "filename": filename, "content": content}
        try:
            output_path = await self._office_generator.generate(
                event, file_info["type"], filename, file_info
            )
            delivery_error = await self._file_delivery_service.deliver_generated_file(
                event,
                output_path,
                missing_message="错误：文件生成失败，未找到输出文件",
                oversized_template="错误：文件过大 ({file_size})，超过限制 {max_size}",
            )
            if delivery_error:
                return self._finalize_error(event, delivery_error)
            return None
        except Exception as exc:
            return self._finalize_error(event, f"错误：文件操作异常: {exc}")
