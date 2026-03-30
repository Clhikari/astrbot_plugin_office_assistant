import importlib
import tempfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path

from astrbot.api import logger
from astrbot.api.star import StarTools

from ..agent_tools import build_document_toolset
from ..constants import (
    DEFAULT_MAX_FILE_SIZE_MB,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
    MSG_DOCUMENT_EXPORTED,
    OFFICE_LIBS,
)
from ..message_buffer import MessageBuffer
from ..office_generator import OfficeGenerator
from ..pdf_converter import PDFConverter
from ..preview_generator import PreviewGenerator
from .access_policy_service import AccessPolicyService
from .command_service import CommandService
from .delivery_service import DeliveryService
from .error_hook_service import ErrorHookService
from .file_tool_service import FileToolService
from .generated_file_delivery_service import GeneratedFileDeliveryService
from .incoming_message_service import IncomingMessageService
from .llm_request_policy import LLMRequestPolicy
from .post_export_hook_service import PostExportHookService
from .request_hook_service import RequestHookService
from .upload_session_service import UploadSessionService
from .word_read_service import WordReadService
from .workspace_service import WorkspaceService


@dataclass(slots=True)
class PluginSettings:
    auto_delete: bool
    max_file_size: int
    enable_docx_image_review: bool
    max_inline_docx_image_bytes: int
    max_inline_docx_image_count: int
    buffer_wait: int
    reply_to_user: bool
    require_at_in_group: bool
    enable_features_in_group: bool
    auto_block_execution_tools: bool
    enable_preview: bool
    preview_dpi: int
    allow_external_input_files: bool
    feature_settings: dict
    recent_text_ttl_seconds: int
    recent_text_max_entries: int
    recent_text_cleanup_interval_seconds: int


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
    recent_text_by_session: dict
    document_toolset: object
    llm_request_policy: LLMRequestPolicy
    delivery_service: DeliveryService
    post_export_hook_service: PostExportHookService
    file_tool_service: FileToolService
    command_service: CommandService
    error_hook_service: ErrorHookService
    message_buffer: MessageBuffer
    incoming_message_service: IncomingMessageService


def build_plugin_runtime(
    *,
    context,
    config,
    plugin_name: str,
    handle_exported_document_tool,
    extract_upload_source,
    store_uploaded_file,
) -> PluginRuntimeBundle:
    settings = _load_settings(config)
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
        recent_text_max_entries=settings.recent_text_max_entries,
        recent_text_cleanup_interval_seconds=settings.recent_text_cleanup_interval_seconds,
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
    )
    request_hook_service = RequestHookService(
        auto_block_execution_tools=settings.auto_block_execution_tools,
        get_cached_upload_infos=upload_session_service.get_cached_upload_infos,
        extract_upload_source=extract_upload_source,
        store_uploaded_file=store_uploaded_file,
        allow_external_input_files=settings.allow_external_input_files,
    )
    llm_request_policy = LLMRequestPolicy(
        document_toolset=document_toolset,
        require_at_in_group=settings.require_at_in_group,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        is_bot_mentioned=access_policy_service.is_bot_mentioned,
        request_hook_service=request_hook_service,
    )
    file_tool_service = FileToolService(
        workspace_service=workspace_service,
        office_generator=office_gen,
        pdf_converter=pdf_converter,
        delivery_service=delivery_service,
        generated_file_delivery_service=GeneratedFileDeliveryService(
            workspace_service=workspace_service,
            delivery_service=delivery_service,
        ),
        word_read_service=WordReadService(
            workspace_service=workspace_service,
            enable_docx_image_review=settings.enable_docx_image_review,
            max_inline_docx_image_bytes=settings.max_inline_docx_image_bytes,
            max_inline_docx_image_count=settings.max_inline_docx_image_count,
        ),
        office_libs=office_libs,
        allow_external_input_files=settings.allow_external_input_files,
        enable_docx_image_review=settings.enable_docx_image_review,
        max_inline_docx_image_bytes=settings.max_inline_docx_image_bytes,
        max_inline_docx_image_count=settings.max_inline_docx_image_count,
        is_group_feature_enabled=access_policy_service.is_group_feature_enabled,
        check_permission=access_policy_service.check_permission,
        group_feature_disabled_error=access_policy_service.group_feature_disabled_error,
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
        recent_text_by_session=upload_session_service.recent_text_by_session,
        document_toolset=document_toolset,
        llm_request_policy=llm_request_policy,
        delivery_service=delivery_service,
        post_export_hook_service=post_export_hook_service,
        file_tool_service=file_tool_service,
        command_service=command_service,
        error_hook_service=error_hook_service,
        message_buffer=message_buffer,
        incoming_message_service=incoming_message_service,
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
    return {str(admin_id) for admin_id in root_config.get("admins_id", [])}


def _build_admin_users_resolver(context, *, initial_admin_users: set[str]):
    cache = {"admin_users": set(initial_admin_users)}

    def get_admin_users() -> set[str]:
        return set(cache["admin_users"])

    def refresh() -> set[str]:
        cache["admin_users"] = _extract_admin_users(_resolve_root_config(context))
        return get_admin_users()

    get_admin_users.refresh = refresh
    return get_admin_users


def _load_settings(config) -> PluginSettings:
    file_settings = config.get("file_settings", {})
    trigger_settings = config.get("trigger_settings", {})
    preview_settings = config.get("preview_settings", {})
    path_settings = config.get("path_settings", {})

    auto_delete = file_settings.get("auto_delete_files", True)
    max_file_size = (
        file_settings.get("max_file_size_mb", DEFAULT_MAX_FILE_SIZE_MB) * 1024 * 1024
    )
    enable_docx_image_review = file_settings.get("enable_docx_image_review", True)
    max_inline_docx_image_bytes = (
        file_settings.get("max_inline_docx_image_mb", DEFAULT_MAX_INLINE_DOCX_IMAGE_MB)
        * 1024
        * 1024
    )
    max_inline_docx_image_count = file_settings.get(
        "max_inline_docx_image_count",
        DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    )
    buffer_wait = file_settings.get("message_buffer_seconds", 4)
    reply_to_user = trigger_settings.get("reply_to_user", True)
    require_at_in_group = trigger_settings.get("require_at_in_group", True)
    enable_features_in_group = trigger_settings.get("enable_features_in_group", False)
    auto_block_execution_tools = trigger_settings.get(
        "auto_block_execution_tools", True
    )
    enable_preview = preview_settings.get("enable", True)
    preview_dpi = preview_settings.get("dpi", 150)
    allow_external_input_files = path_settings.get("allow_external_input_files", False)
    feature_settings = config.get("feature_settings", {})
    recent_text_ttl_seconds = max(20, int(buffer_wait) + 10)
    recent_text_max_entries = 512
    recent_text_cleanup_interval_seconds = max(5, min(60, recent_text_ttl_seconds))

    return PluginSettings(
        auto_delete=auto_delete,
        max_file_size=max_file_size,
        enable_docx_image_review=enable_docx_image_review,
        max_inline_docx_image_bytes=max_inline_docx_image_bytes,
        max_inline_docx_image_count=max_inline_docx_image_count,
        buffer_wait=buffer_wait,
        reply_to_user=reply_to_user,
        require_at_in_group=require_at_in_group,
        enable_features_in_group=enable_features_in_group,
        auto_block_execution_tools=auto_block_execution_tools,
        enable_preview=enable_preview,
        preview_dpi=preview_dpi,
        allow_external_input_files=allow_external_input_files,
        feature_settings=feature_settings,
        recent_text_ttl_seconds=recent_text_ttl_seconds,
        recent_text_max_entries=recent_text_max_entries,
        recent_text_cleanup_interval_seconds=recent_text_cleanup_interval_seconds,
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
    for office_type in OFFICE_LIBS:
        try:
            module_name, package_name = OFFICE_LIBS[office_type]
            libs[module_name] = importlib.import_module(module_name)
            logger.debug(f"[文件管理] {package_name} 已加载")
        except ImportError:
            libs[module_name] = None
            logger.warning(f"[文件管理] {package_name} 未安装")
    return libs
