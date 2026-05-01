from dataclasses import dataclass

from ..constants import (
    DEFAULT_MAX_EXCEL_PREVIEW_CHARS,
    DEFAULT_MAX_EXCEL_PREVIEW_ROWS,
    DEFAULT_MAX_EXCEL_PREVIEW_SHEETS,
    DEFAULT_MAX_FILE_SIZE_MB,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
)


@dataclass(slots=True)
class PluginSettings:
    auto_delete: bool
    max_file_size: int
    enable_docx_image_review: bool
    max_inline_docx_image_bytes: int
    max_inline_docx_image_count: int
    max_excel_preview_rows: int
    max_excel_preview_chars: int
    max_excel_preview_sheets: int
    buffer_wait: int
    reply_to_user: bool
    require_at_in_group: bool
    enable_features_in_group: bool
    auto_block_execution_tools: bool
    allow_local_excel_script: bool
    enable_preview: bool
    preview_dpi: int
    allow_external_input_files: bool
    feature_settings: dict
    recent_text_ttl_seconds: int
    upload_session_ttl_seconds: int
    recent_text_max_entries: int
    recent_text_cleanup_interval_seconds: int
    upload_session_cleanup_interval_seconds: int
    js_renderer_entry: str
    default_word_font_name: str
    default_word_heading_font_name: str
    default_word_table_font_name: str
    default_word_code_font_name: str


def load_plugin_settings(config) -> PluginSettings:
    file_settings = config.get("file_settings", {})
    upload_session_settings = config.get("upload_session_settings", {})
    read_settings = config.get("read_settings", {})
    trigger_settings = config.get("trigger_settings", {})
    preview_settings = config.get("preview_settings", {})
    path_settings = config.get("path_settings", {})
    render_settings = config.get("render_settings", {})
    word_style_settings = config.get("word_style_settings", {})

    auto_delete = file_settings.get("auto_delete_files", True)
    max_file_size = (
        file_settings.get("max_file_size_mb", DEFAULT_MAX_FILE_SIZE_MB) * 1024 * 1024
    )
    enable_docx_image_review = read_settings.get(
        "enable_docx_image_review",
        file_settings.get("enable_docx_image_review", True),
    )
    max_inline_docx_image_bytes = (
        read_settings.get(
            "max_inline_docx_image_mb",
            file_settings.get(
                "max_inline_docx_image_mb",
                DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
            ),
        )
        * 1024
        * 1024
    )
    max_inline_docx_image_count = read_settings.get(
        "max_inline_docx_image_count",
        file_settings.get(
            "max_inline_docx_image_count",
            DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
        ),
    )
    max_excel_preview_rows = max(
        0,
        int(
            read_settings.get(
                "max_excel_preview_rows",
                file_settings.get(
                    "max_excel_preview_rows",
                    DEFAULT_MAX_EXCEL_PREVIEW_ROWS,
                ),
            )
        ),
    )
    max_excel_preview_chars = max(
        0,
        int(
            read_settings.get(
                "max_excel_preview_chars",
                file_settings.get(
                    "max_excel_preview_chars",
                    DEFAULT_MAX_EXCEL_PREVIEW_CHARS,
                ),
            )
        ),
    )
    max_excel_preview_sheets = max(
        0,
        int(
            read_settings.get(
                "max_excel_preview_sheets",
                file_settings.get(
                    "max_excel_preview_sheets",
                    DEFAULT_MAX_EXCEL_PREVIEW_SHEETS,
                ),
            )
        ),
    )
    buffer_wait = upload_session_settings.get(
        "message_buffer_seconds",
        file_settings.get("message_buffer_seconds", 4),
    )
    reply_to_user = trigger_settings.get("reply_to_user", True)
    require_at_in_group = trigger_settings.get("require_at_in_group", True)
    enable_features_in_group = trigger_settings.get("enable_features_in_group", False)
    auto_block_execution_tools = trigger_settings.get(
        "auto_block_execution_tools", True
    )
    allow_local_excel_script = trigger_settings.get("allow_local_excel_script", False)
    enable_preview = preview_settings.get("enable", True)
    preview_dpi = preview_settings.get("dpi", 150)
    allow_external_input_files = read_settings.get(
        "allow_external_input_files",
        path_settings.get("allow_external_input_files", False),
    )
    feature_settings = config.get("feature_settings", {})
    recent_text_ttl_seconds = max(
        20,
        int(
            upload_session_settings.get(
                "recent_text_ttl_seconds",
                file_settings.get("recent_text_ttl_seconds", int(buffer_wait) + 10),
            )
        ),
    )
    upload_session_ttl_seconds = max(
        60,
        int(
            upload_session_settings.get(
                "upload_session_ttl_seconds",
                file_settings.get("upload_session_ttl_seconds", 600),
            )
        ),
    )
    recent_text_max_entries = 512
    recent_text_cleanup_interval_seconds = max(5, min(60, recent_text_ttl_seconds))
    upload_session_cleanup_interval_seconds = max(
        10,
        min(300, upload_session_ttl_seconds),
    )
    js_renderer_entry = str(
        render_settings.get(
            "js_renderer_entry",
            render_settings.get("node_renderer_entry", ""),
        )
    ).strip()
    default_word_font_name = str(
        word_style_settings.get("default_font_name", "")
    ).strip()
    default_word_heading_font_name = str(
        word_style_settings.get("default_heading_font_name", "")
    ).strip()
    default_word_table_font_name = str(
        word_style_settings.get("default_table_font_name", "")
    ).strip()
    default_word_code_font_name = str(
        word_style_settings.get("default_code_font_name", "")
    ).strip()

    return PluginSettings(
        auto_delete=auto_delete,
        max_file_size=max_file_size,
        enable_docx_image_review=enable_docx_image_review,
        max_inline_docx_image_bytes=max_inline_docx_image_bytes,
        max_inline_docx_image_count=max_inline_docx_image_count,
        max_excel_preview_rows=max_excel_preview_rows,
        max_excel_preview_chars=max_excel_preview_chars,
        max_excel_preview_sheets=max_excel_preview_sheets,
        buffer_wait=buffer_wait,
        reply_to_user=reply_to_user,
        require_at_in_group=require_at_in_group,
        enable_features_in_group=enable_features_in_group,
        auto_block_execution_tools=auto_block_execution_tools,
        allow_local_excel_script=allow_local_excel_script,
        enable_preview=enable_preview,
        preview_dpi=preview_dpi,
        allow_external_input_files=allow_external_input_files,
        feature_settings=feature_settings,
        recent_text_ttl_seconds=recent_text_ttl_seconds,
        upload_session_ttl_seconds=upload_session_ttl_seconds,
        recent_text_max_entries=recent_text_max_entries,
        recent_text_cleanup_interval_seconds=recent_text_cleanup_interval_seconds,
        upload_session_cleanup_interval_seconds=upload_session_cleanup_interval_seconds,
        js_renderer_entry=js_renderer_entry,
        default_word_font_name=default_word_font_name,
        default_word_heading_font_name=default_word_heading_font_name,
        default_word_table_font_name=default_word_table_font_name,
        default_word_code_font_name=default_word_code_font_name,
    )


__all__ = ["PluginSettings", "load_plugin_settings"]
