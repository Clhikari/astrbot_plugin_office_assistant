import warnings

from .file_delivery_service import FileDeliveryService
from .file_read_service import FileReadService
from .generated_file_delivery_service import GeneratedFileDeliveryService
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

    def __init__(
        self,
        *,
        workspace_service=None,
        office_generator=None,
        pdf_converter=None,
        delivery_service=None,
        generated_file_delivery_service=None,
        word_read_service=None,
        office_libs: dict | None = None,
        allow_external_input_files: bool = False,
        enable_docx_image_review: bool = True,
        max_inline_docx_image_bytes: int = 0,
        max_inline_docx_image_count: int = 0,
        is_group_feature_enabled=None,
        check_permission=None,
        group_feature_disabled_error=None,
        file_read_service=None,
        office_generate_service=None,
        pdf_convert_service=None,
    ) -> None:
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
        event,
        filename: str = "",
    ):
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
