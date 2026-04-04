from __future__ import annotations

import tempfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .settings import PluginSettings

if TYPE_CHECKING:
    from ..message_buffer import MessageBuffer
    from ..office_generator import OfficeGenerator
    from ..pdf_converter import PDFConverter
    from ..preview_generator import PreviewGenerator
    from ..services.access_policy_service import AccessPolicyService
    from ..services.command_service import CommandService
    from ..services.delivery_service import DeliveryService
    from ..services.error_hook_service import ErrorHookService
    from ..services.file_delivery_service import FileDeliveryService
    from ..services.file_read_service import FileReadService
    from ..services.file_tool_service import FileToolService
    from ..services.generated_file_delivery_service import (
        GeneratedFileDeliveryService,
    )
    from ..services.incoming_message_service import IncomingMessageService
    from ..services.llm_request_policy import LLMRequestPolicy
    from ..services.office_generate_service import OfficeGenerateService
    from ..services.pdf_convert_service import PdfConvertService
    from ..services.post_export_hook_service import PostExportHookService
    from ..services.prompt_context_service import PromptContextService
    from ..services.request_hook_service import RequestHookService
    from ..services.upload_session_service import UploadSessionService
    from ..services.word_read_service import WordReadService
    from ..services.workspace_service import WorkspaceService


@dataclass(slots=True)
class PluginRuntimeBundle:
    settings: PluginSettings
    temp_dir: tempfile.TemporaryDirectory | None
    plugin_data_path: Path
    executor: ThreadPoolExecutor
    office_gen: OfficeGenerator
    pdf_converter: PDFConverter
    preview_gen: PreviewGenerator
    office_libs: dict
    workspace_service: WorkspaceService
    access_policy_service: AccessPolicyService
    upload_session_service: UploadSessionService
    document_toolset: Any
    llm_request_policy: LLMRequestPolicy
    prompt_context_service: PromptContextService
    delivery_service: DeliveryService
    generated_file_delivery_service: GeneratedFileDeliveryService
    file_delivery_service: FileDeliveryService
    word_read_service: WordReadService
    file_read_service: FileReadService
    office_generate_service: OfficeGenerateService
    pdf_convert_service: PdfConvertService
    post_export_hook_service: PostExportHookService
    file_tool_service: FileToolService
    command_service: CommandService
    error_hook_service: ErrorHookService
    message_buffer: MessageBuffer
    incoming_message_service: IncomingMessageService


@dataclass(slots=True)
class AdminUsersResolver:
    context: object
    admin_users: set[str]

    def __call__(self) -> set[str]:
        return set(self.admin_users)

    def refresh(self) -> set[str]:
        from ..services.runtime_builder import (
            _extract_admin_users,
            _resolve_root_config,
        )

        self.admin_users = _extract_admin_users(_resolve_root_config(self.context))
        return self()


@dataclass(slots=True)
class RequestPipelineServices:
    prompt_context_service: PromptContextService
    request_hook_service: RequestHookService
    llm_request_policy: LLMRequestPolicy


@dataclass(slots=True)
class FileProcessingServices:
    generated_file_delivery_service: GeneratedFileDeliveryService
    file_delivery_service: FileDeliveryService
    word_read_service: WordReadService
    file_read_service: FileReadService
    office_generate_service: OfficeGenerateService
    pdf_convert_service: PdfConvertService
    file_tool_service: FileToolService


__all__ = [
    "AdminUsersResolver",
    "FileProcessingServices",
    "PluginRuntimeBundle",
    "PluginSettings",
    "RequestPipelineServices",
]
