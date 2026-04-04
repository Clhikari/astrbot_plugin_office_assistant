from .access_policy_service import AccessPolicyService
from .command_service import CommandService
from .delivery_service import DeliveryService
from .error_hook_service import ErrorHookService
from .export_hook_service import ExportHookService
from .file_delivery_service import FileDeliveryService
from .file_read_service import FileReadService
from .file_tool_service import FileToolService
from .generated_file_delivery_service import (
    GeneratedFileDeliveryResult,
    GeneratedFileDeliveryService,
)
from .incoming_message_service import IncomingMessageService
from .llm_request_policy import LLMRequestPolicy
from .office_generate_service import OfficeGenerateService
from .pdf_convert_service import PdfConvertService
from .post_export_hook_service import PostExportHookService
from .prompt_context_service import PromptContextService
from .request_hook_service import RequestHookService
from .upload_prompt_service import UploadPromptService
from .upload_types import UploadInfo
from .upload_session_service import UploadSessionService
from .workspace_service import WorkspaceService
from .word_read_service import WordReadService

__all__ = [
    "AccessPolicyService",
    "CommandService",
    "DeliveryService",
    "ErrorHookService",
    "ExportHookService",
    "FileDeliveryService",
    "FileReadService",
    "FileToolService",
    "GeneratedFileDeliveryResult",
    "GeneratedFileDeliveryService",
    "IncomingMessageService",
    "LLMRequestPolicy",
    "OfficeGenerateService",
    "PdfConvertService",
    "PostExportHookService",
    "PluginRuntimeBundle",
    "PluginSettings",
    "PromptContextService",
    "RequestHookService",
    "UploadPromptService",
    "UploadInfo",
    "UploadSessionService",
    "WorkspaceService",
    "WordReadService",
    "build_plugin_runtime",
]


def __getattr__(name: str):
    if name == "PluginSettings":
        from ..app.settings import PluginSettings

        return PluginSettings
    if name == "PluginRuntimeBundle":
        from ..app.runtime import __dict__ as runtime_namespace

        return runtime_namespace[name]
    if name == "build_plugin_runtime":
        from .runtime_builder import build_plugin_runtime

        return build_plugin_runtime
    raise AttributeError(name)
