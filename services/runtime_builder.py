import importlib
import tempfile
from collections.abc import Iterable
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from astrbot.api import logger
from astrbot.api.star import StarTools

from ..agent_tools import build_document_toolset
from ..app.runtime import (
    AdminUsersResolver,
    FileProcessingServices,
    PluginRuntimeBundle,
    RequestPipelineServices,
)
from ..app.settings import PluginSettings, load_plugin_settings
from ..constants import (
    MSG_DOCUMENT_EXPORTED,
    OFFICE_LIBS,
)
from ..domain.document.render_backends import DocumentRenderBackendConfig
from ..message_buffer import MessageBuffer
from ..office_generator import OfficeGenerator
from ..pdf_converter import PDFConverter
from ..preview_generator import PreviewGenerator
from .access_policy_service import AccessPolicyService
from .command_service import CommandService
from .delivery_service import DeliveryService
from .error_hook_service import ErrorHookService
from .file_delivery_service import FileDeliveryService
from .file_read_service import FileReadService
from .file_tool_service import FileToolService
from .generated_file_delivery_service import GeneratedFileDeliveryService
from .incoming_message_service import IncomingMessageService
from .llm_request_policy import LLMRequestPolicy
from .office_generate_service import OfficeGenerateService
from .pdf_convert_service import PdfConvertService
from .post_export_hook_service import PostExportHookService
from .prompt_context_service import PromptContextService
from .request_hook_service import RequestHookService
from .upload_session_service import UploadSessionService
from .word_read_service import WordReadService
from .workspace_service import WorkspaceService


def build_plugin_runtime(
    *,
    context,
    config,
    plugin_name: str,
    handle_exported_document_tool,
    extract_upload_source,
    store_uploaded_file,
) -> PluginRuntimeBundle:
    settings = load_plugin_settings(config)
    root_config = _resolve_root_config(context)
    admin_users = _extract_admin_users(root_config)
    get_admin_users = _build_admin_users_resolver(
        context,
        initial_admin_users=admin_users,
    )
    temp_dir, plugin_data_path = _prepare_workspace(
        settings.auto_delete, plugin_name=plugin_name
    )
    executor = ThreadPoolExecutor(max_workers=4)
    office_gen = OfficeGenerator(plugin_data_path, executor=executor)
    pdf_converter = PDFConverter(plugin_data_path, executor=executor)
    preview_gen = PreviewGenerator(dpi=settings.preview_dpi)
    office_libs = _check_office_libs()

    workspace_service = WorkspaceService(
        plugin_data_path=plugin_data_path,
        executor=executor,
        office_libs=office_libs,
        max_file_size=settings.max_file_size,
        feature_settings=settings.feature_settings,
    )
    access_policy_service = AccessPolicyService(
        whitelist_users=config.get("permission_settings", {}).get(
            "whitelist_users", []
        ),
        admin_users=list(admin_users),
        get_admin_users=get_admin_users,
        enable_features_in_group=settings.enable_features_in_group,
    )
    upload_session_service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=settings.recent_text_ttl_seconds,
        upload_session_ttl_seconds=settings.upload_session_ttl_seconds,
        recent_text_max_entries=settings.recent_text_max_entries,
        recent_text_cleanup_interval_seconds=settings.recent_text_cleanup_interval_seconds,
        upload_session_cleanup_interval_seconds=settings.upload_session_cleanup_interval_seconds,
        extract_upload_source=extract_upload_source,
        store_uploaded_file=store_uploaded_file,
        allow_external_input_files=settings.allow_external_input_files,
    )
    delivery_service = DeliveryService(
        executor=executor,
        preview_generator=preview_gen,
        enable_preview=settings.enable_preview,
        auto_delete=settings.auto_delete,
        reply_to_user=settings.reply_to_user,
    )
    post_export_hook_service = PostExportHookService(
        executor=executor,
        preview_generator=preview_gen,
        enable_preview=settings.enable_preview,
        auto_delete=settings.auto_delete,
        reply_to_user=settings.reply_to_user,
        exported_message=MSG_DOCUMENT_EXPORTED,
    )
    document_toolset = build_document_toolset(
        workspace_dir=plugin_data_path,
        after_export=handle_exported_document_tool,
        render_backend_config=_build_word_render_backend_config(settings),
    )
    request_pipeline_services = _build_request_pipeline_services(
        settings=settings,
        upload_session_service=upload_session_service,
        access_policy_service=access_policy_service,
        document_toolset=document_toolset,
        extract_upload_source=extract_upload_source,
        store_uploaded_file=store_uploaded_file,
    )
    file_processing_services = _build_file_processing_services(
        settings=settings,
        workspace_service=workspace_service,
        delivery_service=delivery_service,
        office_gen=office_gen,
        pdf_converter=pdf_converter,
        office_libs=office_libs,
        access_policy_service=access_policy_service,
    )
    command_service = CommandService(
        workspace_service=workspace_service,
        pdf_converter=pdf_converter,
        plugin_data_path=plugin_data_path,
        auto_delete=settings.auto_delete,
        allow_external_input_files=settings.allow_external_input_files,
        enable_features_in_group=settings.enable_features_in_group,
        auto_block_execution_tools=settings.auto_block_execution_tools,
        reply_to_user=settings.reply_to_user,
        upload_session_service=upload_session_service,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
    )
    error_hook_service = ErrorHookService(
        context=context,
        config=config,
        plugin_name=plugin_name,
    )
    message_buffer = MessageBuffer(wait_seconds=settings.buffer_wait)
    incoming_message_service = IncomingMessageService(
        message_buffer=message_buffer,
        remember_recent_text=upload_session_service.remember_recent_text,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
    )

    return PluginRuntimeBundle(
        settings=settings,
        temp_dir=temp_dir,
        plugin_data_path=plugin_data_path,
        executor=executor,
        office_gen=office_gen,
        pdf_converter=pdf_converter,
        preview_gen=preview_gen,
        office_libs=office_libs,
        workspace_service=workspace_service,
        access_policy_service=access_policy_service,
        upload_session_service=upload_session_service,
        document_toolset=document_toolset,
        llm_request_policy=request_pipeline_services.llm_request_policy,
        prompt_context_service=request_pipeline_services.prompt_context_service,
        delivery_service=delivery_service,
        generated_file_delivery_service=file_processing_services.generated_file_delivery_service,
        file_delivery_service=file_processing_services.file_delivery_service,
        word_read_service=file_processing_services.word_read_service,
        file_read_service=file_processing_services.file_read_service,
        office_generate_service=file_processing_services.office_generate_service,
        pdf_convert_service=file_processing_services.pdf_convert_service,
        post_export_hook_service=post_export_hook_service,
        file_tool_service=file_processing_services.file_tool_service,
        command_service=command_service,
        error_hook_service=error_hook_service,
        message_buffer=message_buffer,
        incoming_message_service=incoming_message_service,
    )


def _build_word_render_backend_config(
    settings: PluginSettings,
) -> DocumentRenderBackendConfig:
    return DocumentRenderBackendConfig(
        preferred_backend=settings.word_render_backend,
        fallback_enabled=settings.word_render_fallback_enabled,
        ppt_preferred_backend=settings.ppt_render_backend,
        excel_preferred_backend=settings.excel_render_backend,
        node_renderer_entry=settings.js_renderer_entry,
    )


def _build_request_pipeline_services(
    *,
    settings: PluginSettings,
    upload_session_service: UploadSessionService,
    access_policy_service: AccessPolicyService,
    document_toolset,
    extract_upload_source,
    store_uploaded_file,
    ) -> RequestPipelineServices:
    prompt_context_service = PromptContextService(
        allow_external_input_files=settings.allow_external_input_files
    )
    document_store = getattr(document_toolset, "document_store", None)

    def _get_document_prompt_summary(document_id: str) -> dict[str, object] | None:
        if document_store is None:
            return None
        try:
            return document_store.build_prompt_summary(document_id)
        except KeyError:
            return None

    request_hook_service = RequestHookService(
        auto_block_execution_tools=settings.auto_block_execution_tools,
        get_cached_upload_infos=upload_session_service.get_cached_upload_infos,
        extract_upload_source=extract_upload_source,
        store_uploaded_file=store_uploaded_file,
        allow_external_input_files=settings.allow_external_input_files,
        get_document_prompt_summary=_get_document_prompt_summary,
        prompt_context_service=prompt_context_service,
    )
    llm_request_policy = LLMRequestPolicy(
        document_toolset=document_toolset,
        require_at_in_group=settings.require_at_in_group,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        is_bot_mentioned=access_policy_service.is_bot_mentioned,
        request_hook_service=request_hook_service,
        prompt_context_service=prompt_context_service,
    )
    return RequestPipelineServices(
        prompt_context_service=prompt_context_service,
        request_hook_service=request_hook_service,
        llm_request_policy=llm_request_policy,
    )


def _build_file_processing_services(
    *,
    settings: PluginSettings,
    workspace_service: WorkspaceService,
    delivery_service: DeliveryService,
    office_gen: OfficeGenerator,
    pdf_converter: PDFConverter,
    office_libs: dict,
    access_policy_service: AccessPolicyService,
) -> FileProcessingServices:
    generated_file_delivery_service = GeneratedFileDeliveryService(
        workspace_service=workspace_service,
        delivery_service=delivery_service,
    )
    file_delivery_service = FileDeliveryService(
        generated_file_delivery_service=generated_file_delivery_service,
    )
    word_read_service = WordReadService(
        workspace_service=workspace_service,
        enable_docx_image_review=settings.enable_docx_image_review,
        max_inline_docx_image_bytes=settings.max_inline_docx_image_bytes,
        max_inline_docx_image_count=settings.max_inline_docx_image_count,
    )
    file_read_service = FileReadService(
        workspace_service=workspace_service,
        word_read_service=word_read_service,
        allow_external_input_files=settings.allow_external_input_files,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
    )
    office_generate_service = OfficeGenerateService(
        workspace_service=workspace_service,
        office_generator=office_gen,
        file_delivery_service=file_delivery_service,
        office_libs=office_libs,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
    )
    pdf_convert_service = PdfConvertService(
        workspace_service=workspace_service,
        pdf_converter=pdf_converter,
        file_delivery_service=file_delivery_service,
        allow_external_input_files=settings.allow_external_input_files,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
    )
    file_tool_service = FileToolService(
        workspace_service=workspace_service,
        office_generator=office_gen,
        pdf_converter=pdf_converter,
        delivery_service=delivery_service,
        generated_file_delivery_service=generated_file_delivery_service,
        word_read_service=word_read_service,
        office_libs=office_libs,
        allow_external_input_files=settings.allow_external_input_files,
        enable_docx_image_review=settings.enable_docx_image_review,
        max_inline_docx_image_bytes=settings.max_inline_docx_image_bytes,
        max_inline_docx_image_count=settings.max_inline_docx_image_count,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
        file_read_service=file_read_service,
        office_generate_service=office_generate_service,
        pdf_convert_service=pdf_convert_service,
    )
    return FileProcessingServices(
        generated_file_delivery_service=generated_file_delivery_service,
        file_delivery_service=file_delivery_service,
        word_read_service=word_read_service,
        file_read_service=file_read_service,
        office_generate_service=office_generate_service,
        pdf_convert_service=pdf_convert_service,
        file_tool_service=file_tool_service,
    )


def _resolve_root_config(context) -> dict:
    get_config = getattr(context, "get_config", None)
    if callable(get_config):
        config = get_config()
        if isinstance(config, dict):
            return config

    legacy_config = getattr(context, "astrbot_config", None)
    if isinstance(legacy_config, dict):
        return legacy_config

    return {}


def _extract_admin_users(root_config: dict) -> set[str]:
    if not isinstance(root_config, dict):
        return set()

    raw_admins = root_config.get("admins_id", [])
    if raw_admins is None:
        return set()
    if isinstance(raw_admins, (str, bytes)):
        return {str(raw_admins)}
    if isinstance(raw_admins, Iterable):
        return {str(admin_id) for admin_id in raw_admins}
    return {str(raw_admins)}


def _build_admin_users_resolver(context, *, initial_admin_users: set[str]):
    return AdminUsersResolver(
        context=context,
        admin_users=set(initial_admin_users),
    )


def _prepare_workspace(
    auto_delete: bool,
    *,
    plugin_name: str,
) -> tuple[tempfile.TemporaryDirectory | None, Path]:
    if auto_delete:
        temp_dir = tempfile.TemporaryDirectory(prefix="astrbot_file_")
        return temp_dir, Path(temp_dir.name)

    plugin_data_path = StarTools.get_data_dir(plugin_name) / "files"
    plugin_data_path.mkdir(parents=True, exist_ok=True)
    return None, plugin_data_path


def _check_office_libs() -> dict:
    libs = {}
    for office_type, (module_name, package_name) in OFFICE_LIBS.items():
        try:
            libs[module_name] = importlib.import_module(module_name)
            logger.debug(f"[文件管理] {package_name} 已加载")
        except ImportError:
            libs[module_name] = None
            logger.warning(f"[文件管理] {package_name} 未安装")
    return libs
