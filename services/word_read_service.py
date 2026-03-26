import base64
import mimetypes
from collections.abc import AsyncGenerator
from pathlib import Path

import mcp

from astrbot.api import logger

from ..constants import (
    DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
)
from ..utils import (
    WORD_ITEM_IMAGE,
    WORD_ITEM_TEXT,
    format_file_size,
    safe_error_message,
)


class WordReadService:
    def __init__(
        self,
        *,
        workspace_service,
        enable_docx_image_review: bool = True,
        max_inline_docx_image_bytes: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_MB
        * 1024
        * 1024,
        max_inline_docx_image_count: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    ) -> None:
        self._workspace_service = workspace_service
        self._enable_docx_image_review = bool(enable_docx_image_review)
        self._max_inline_docx_image_bytes = max(0, int(max_inline_docx_image_bytes))
        self._max_inline_docx_image_count = max(0, int(max_inline_docx_image_count))

    def _build_image_tool_result(
        self,
        image_paths: list[Path],
    ) -> mcp.types.CallToolResult | None:
        content: list[mcp.types.ImageContent] = []
        for image_path in image_paths:
            try:
                base64_data = base64.b64encode(image_path.read_bytes()).decode("utf-8")
            except Exception as exc:
                logger.warning(
                    "[文件管理] 读取嵌入图片失败: %s",
                    safe_error_message(exc, "读取嵌入图片失败"),
                )
                continue

            mime_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
            content.append(
                mcp.types.ImageContent(
                    type="image",
                    data=base64_data,
                    mimeType=mime_type,
                )
            )

        if not content:
            return None

        return mcp.types.CallToolResult(content=content)

    def _plan_inline_word_images(
        self,
        image_paths: list[Path],
    ) -> dict[int, str]:
        skipped_reasons: dict[int, str] = {}
        selected_count = 0

        logger.info("[文件管理] Word 嵌入图片提取数量: %d", len(image_paths))

        for index, image_path in enumerate(image_paths, start=1):
            try:
                image_size = image_path.stat().st_size
            except OSError as exc:
                skipped_reasons[index] = (
                    f"未注入模型上下文（读取失败：{safe_error_message(exc)}）。"
                )
                continue

            logger.info(
                "[文件管理] Word 嵌入图片%d 实际大小: %s",
                index,
                format_file_size(image_size),
            )

            if selected_count >= self._max_inline_docx_image_count:
                skipped_reasons[index] = (
                    f"未注入模型上下文（超过单文档最多 {self._max_inline_docx_image_count} 张限制）。"
                )
                continue

            if image_size > self._max_inline_docx_image_bytes:
                skipped_reasons[index] = (
                    "未注入模型上下文（文件大小 "
                    f"{format_file_size(image_size)} 超过 "
                    f"{format_file_size(self._max_inline_docx_image_bytes)} 限制）。"
                )
                continue

            selected_count += 1

        return skipped_reasons

    async def iter_word_results(
        self,
        resolved_path: Path,
        display_name: str,
        suffix: str,
        file_size: int,
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        extracted = self._workspace_service.extract_word_content(
            resolved_path,
            include_images=self._enable_docx_image_review,
        )
        if extracted is None:
            yield f"错误：文件 '{display_name}' 无法读取，可能未安装对应解析库"
            return

        image_count = getattr(extracted, "image_count", 0) or len(
            getattr(extracted, "image_paths", [])
        )

        if not self._enable_docx_image_review:
            text_chunks: list[str] = []
            for item in getattr(extracted, "items", []):
                if item.type != WORD_ITEM_TEXT:
                    continue
                text = (item.text or "").strip()
                if text:
                    text_chunks.append(text)
            formatted = "\n".join(text_chunks) if text_chunks else None
            if not formatted and extracted.text:
                formatted = extracted.text.strip() or None
            if formatted:
                yield self._workspace_service.format_file_result(
                    display_name, suffix, file_size, formatted
                )
            elif image_count > 0:
                yield self._workspace_service.format_file_result(
                    display_name,
                    suffix,
                    file_size,
                    "该 Word 文档仅包含图片内容，当前未启用图片理解。",
                )
            return

        skipped_image_reasons = self._plan_inline_word_images(extracted.image_paths)
        fallback_image_index = 0
        text_chunks: list[str] = []
        selected_image_paths: list[Path] = []

        if extracted.items:
            for item in extracted.items:
                if item.type == WORD_ITEM_TEXT:
                    text = (item.text or "").strip()
                    if text:
                        text_chunks.append(text)
                    continue

                if item.type != WORD_ITEM_IMAGE or item.image_path is None:
                    continue

                image_index = getattr(item, "image_index", None)
                if image_index is None:
                    fallback_image_index += 1
                    image_index = fallback_image_index

                reason = skipped_image_reasons.get(image_index)
                if reason:
                    text_chunks.append(f"[插图{image_index}]（{reason}）")
                    continue

                text_chunks.append(f"[插图{image_index}]")
                selected_image_paths.append(item.image_path)

            final_text = "\n".join(part.strip() for part in text_chunks if part.strip())
            if final_text:
                yield self._workspace_service.format_file_result(
                    display_name, suffix, file_size, final_text
                )
            image_result = self._build_image_tool_result(selected_image_paths)
            if image_result is not None:
                yield image_result
            return

        formatted = self._workspace_service.format_word_content(extracted)
        if formatted:
            yield self._workspace_service.format_file_result(
                display_name, suffix, file_size, formatted
            )
