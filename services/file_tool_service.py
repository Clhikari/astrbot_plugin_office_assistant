from collections.abc import AsyncGenerator
import warnings

import mcp
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    DEFAULT_MAX_EXCEL_PREVIEW_CHARS,
    DEFAULT_MAX_EXCEL_PREVIEW_ROWS,
    DEFAULT_MAX_EXCEL_PREVIEW_SHEETS,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
)
from .file_delivery_service import FileDeliveryService
from .file_read_service import FileReadService
from .generated_file_delivery_service import GeneratedFileDeliveryService
from .excel_script_service import ExcelScriptService
from .office_generate_service import OfficeGenerateService
from .pdf_convert_service import PdfConvertService
from .word_read_service import WordReadService


class FileToolService:
    @staticmethod
    def _raise_missing_dependencies(
        service_name: str,
        dependencies: list[str],
    ) -> None:
        missing = ", ".join(dependencies)
        raise ValueError(
            f"{service_name} requires injected service or dependencies: {missing}"
        )

    @classmethod
    def _require_permission_callbacks(cls) -> None:
        cls._raise_missing_dependencies(
            "permission callbacks",
            [
                "is_group_feature_enabled",
                "check_permission",
                "group_feature_disabled_error",
            ],
        )

    def __init__(
        self,
        *,
        astrbot_context=None,
        auto_block_execution_tools: bool = False,
        allow_local_excel_script: bool = False,
        workspace_service=None,
        office_generator=None,
        pdf_converter=None,
        delivery_service=None,
        generated_file_delivery_service=None,
        word_read_service=None,
        office_libs: dict | None = None,
        allow_external_input_files: bool = False,
        enable_docx_image_review: bool = True,
        max_inline_docx_image_bytes: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_MB
        * 1024
        * 1024,
        max_inline_docx_image_count: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
        max_excel_preview_rows: int = DEFAULT_MAX_EXCEL_PREVIEW_ROWS,
        max_excel_preview_chars: int = DEFAULT_MAX_EXCEL_PREVIEW_CHARS,
        max_excel_preview_sheets: int = DEFAULT_MAX_EXCEL_PREVIEW_SHEETS,
        is_group_feature_enabled=None,
        check_permission=None,
        group_feature_disabled_error=None,
        file_read_service=None,
        excel_script_service=None,
        office_generate_service=None,
        pdf_convert_service=None,
    ) -> None:
        needs_permission_callbacks = (
            file_read_service is None
            or office_generate_service is None
            or pdf_convert_service is None
            or (excel_script_service is None and workspace_service is not None)
        )
        if needs_permission_callbacks and (
            is_group_feature_enabled is None
            or check_permission is None
            or group_feature_disabled_error is None
        ):
            self._require_permission_callbacks()

        if file_read_service is None:
            if workspace_service is None:
                self._raise_missing_dependencies(
                    "file_read_service",
                    ["workspace_service"],
                )
            if word_read_service is None:
                word_read_service = WordReadService(
                    workspace_service=workspace_service,
                    enable_docx_image_review=enable_docx_image_review,
                    max_inline_docx_image_bytes=max_inline_docx_image_bytes,
                    max_inline_docx_image_count=max_inline_docx_image_count,
                )

        if office_generate_service is None or pdf_convert_service is None:
            if generated_file_delivery_service is None:
                missing_delivery_dependencies: list[str] = []
                if workspace_service is None:
                    missing_delivery_dependencies.append("workspace_service")
                if delivery_service is None:
                    missing_delivery_dependencies.append("delivery_service")
                if missing_delivery_dependencies:
                    self._raise_missing_dependencies(
                        "generated_file_delivery_service",
                        missing_delivery_dependencies,
                    )
                generated_file_delivery_service = GeneratedFileDeliveryService(
                    workspace_service=workspace_service,
                    delivery_service=delivery_service,
                )

        generated_output_delivery_service = FileDeliveryService(
            generated_file_delivery_service=generated_file_delivery_service,
        )
        self._file_read_service = file_read_service or FileReadService(
            workspace_service=workspace_service,
            word_read_service=word_read_service,
            allow_external_input_files=allow_external_input_files,
            is_group_feature_enabled=is_group_feature_enabled,
            check_permission=check_permission,
            group_feature_disabled_error=group_feature_disabled_error,
            max_excel_preview_rows=max_excel_preview_rows,
            max_excel_preview_chars=max_excel_preview_chars,
            max_excel_preview_sheets=max_excel_preview_sheets,
        )
        self._excel_script_service = excel_script_service
        if self._excel_script_service is None and workspace_service is not None:
            self._excel_script_service = ExcelScriptService(
                astrbot_context=astrbot_context,
                auto_block_execution_tools=auto_block_execution_tools,
                allow_local_excel_script=allow_local_excel_script,
                workspace_service=workspace_service,
                file_delivery_service=generated_output_delivery_service,
                allow_external_input_files=allow_external_input_files,
                is_group_feature_enabled=is_group_feature_enabled,
                check_permission=check_permission,
                group_feature_disabled_error=group_feature_disabled_error,
            )
        self._office_generate_service = (
            office_generate_service
            or OfficeGenerateService(
                workspace_service=workspace_service,
                office_generator=office_generator,
                file_delivery_service=generated_output_delivery_service,
                office_libs=office_libs or {},
                is_group_feature_enabled=is_group_feature_enabled,
                check_permission=check_permission,
                group_feature_disabled_error=group_feature_disabled_error,
            )
        )
        self._pdf_convert_service = pdf_convert_service or PdfConvertService(
            workspace_service=workspace_service,
            pdf_converter=pdf_converter,
            file_delivery_service=generated_output_delivery_service,
            allow_external_input_files=allow_external_input_files,
            is_group_feature_enabled=is_group_feature_enabled,
            check_permission=check_permission,
            group_feature_disabled_error=group_feature_disabled_error,
        )

    async def iter_read_file_tool_results(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        async for result in self._file_read_service.iter_read_file_tool_results(
            event,
            filename,
        ):
            yield result

    async def read_file(
        self,
        event,
        filename: str = "",
    ) -> str | None:
        return await self._file_read_service.read_file(event, filename)

    async def iter_read_workbook_tool_results(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        async for result in self._file_read_service.iter_read_workbook_tool_results(
            event,
            filename,
        ):
            yield result

    async def read_workbook(
        self,
        event,
        filename: str = "",
    ) -> str | None:
        return await self._file_read_service.read_workbook(event, filename)

    async def execute_excel_script(
        self,
        event,
        *,
        script: str,
        input_files: list[str] | None = None,
        output_name: str | None = None,
    ) -> str:
        if self._excel_script_service is None:
            self._raise_missing_dependencies(
                "excel_script_service",
                [
                    "workspace_service",
                    "is_group_feature_enabled",
                    "check_permission",
                    "group_feature_disabled_error",
                ],
            )
        return await self._excel_script_service.execute_excel_script(
            event,
            script=script,
            input_files=input_files,
            output_name=output_name,
        )

    async def create_office_file(
        self,
        event,
        filename: str = "",
        content: str = "",
        file_type: str = "",
    ) -> str | None:
        warnings.warn(
            "create_office_file is deprecated for Word documents. "
            "Use create_document -> add_blocks -> finalize_document -> export_document instead.",
            DeprecationWarning,
            stacklevel=2,
        )
        return await self._office_generate_service.create_office_file(
            event,
            filename=filename,
            content=content,
            file_type=file_type,
        )

    async def convert_to_pdf(
        self,
        event,
        filename: str = "",
        file_path: str = "",
    ) -> str | None:
        return await self._pdf_convert_service.convert_to_pdf(
            event,
            filename=filename,
            file_path=file_path,
        )

    async def convert_from_pdf(
        self,
        event,
        filename: str = "",
        target_format: str = "word",
        file_id: str = "",
    ) -> str | None:
        return await self._pdf_convert_service.convert_from_pdf(
            event,
            filename=filename,
            target_format=target_format,
            file_id=file_id,
        )
