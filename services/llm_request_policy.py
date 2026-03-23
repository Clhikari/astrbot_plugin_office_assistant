from collections.abc import Awaitable, Callable
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from ..constants import (
    ALL_OFFICE_SUFFIXES,
    EXECUTION_TOOLS,
    FILE_TOOLS,
    NOTICE_DOCUMENT_TOOLS_GUIDE,
    NOTICE_TOOLS_DENIED,
    NOTICE_UPLOADED_FILE_TEMPLATE,
    PDF_SUFFIX,
    TEXT_SUFFIXES,
)


class LLMRequestPolicy:
    def __init__(
        self,
        *,
        document_toolset,
        auto_block_execution_tools: bool,
        require_at_in_group: bool,
        is_group_feature_enabled: Callable[[AstrMessageEvent], bool],
        check_permission: Callable[[AstrMessageEvent], bool],
        is_bot_mentioned: Callable[[AstrMessageEvent], bool],
        extract_upload_source: Callable[
            [Comp.File], Awaitable[tuple[Path | None, str]]
        ],
        store_uploaded_file: Callable[[Path, str], Path],
    ) -> None:
        self._document_toolset = document_toolset
        self._auto_block_execution_tools = auto_block_execution_tools
        self._require_at_in_group = require_at_in_group
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._is_bot_mentioned = is_bot_mentioned
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file

    async def apply(self, event: AstrMessageEvent, req: ProviderRequest) -> None:
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = event.message_obj.type == MessageType.FRIEND_MESSAGE

        if not self._is_group_feature_enabled(event):
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)
            req.system_prompt = (req.system_prompt or "") + NOTICE_TOOLS_DENIED
            logger.debug("[文件管理] 群聊总开关关闭，已隐藏全部文件工具")
            return

        has_permission = self._check_permission(event)
        can_process_upload = has_permission or event.is_admin()
        should_expose = (is_friend and event.is_admin()) or (
            has_permission
            and (
                not is_group
                or not self._require_at_in_group
                or self._is_bot_mentioned(event)
            )
        )

        if not should_expose:
            logger.debug(
                f"[文件管理] 用户 {event.get_sender_id()} 权限不足，已隐藏文件工具"
            )
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)
            req.system_prompt = (req.system_prompt or "") + NOTICE_TOOLS_DENIED

        if should_expose and req.func_tool:
            for tool in self._document_toolset.tools:
                req.func_tool.add_tool(tool)

        if should_expose and req.func_tool and self._auto_block_execution_tools:
            for tool_name in EXECUTION_TOOLS:
                req.func_tool.remove_tool(tool_name)
            logger.debug("[文件管理] 已自动屏蔽 shell/python 执行类工具")

        if should_expose and req.func_tool:
            req.system_prompt = (req.system_prompt or "") + NOTICE_DOCUMENT_TOOLS_GUIDE

        if not can_process_upload:
            return

        for component in getattr(event.message_obj, "message", None) or []:
            if not isinstance(component, Comp.File):
                continue

            try:
                src_path, original_name = await self._extract_upload_source(component)
                if not src_path or not src_path.exists():
                    continue

                stored_path = self._store_uploaded_file(src_path, original_name)
                file_suffix = stored_path.suffix.lower()
                type_desc = ""
                is_supported = False

                if file_suffix in ALL_OFFICE_SUFFIXES:
                    type_desc = "Office文档 (Word/Excel/PPT)"
                    is_supported = True
                elif file_suffix in TEXT_SUFFIXES:
                    type_desc = "文本/代码文件"
                    is_supported = True
                elif file_suffix == PDF_SUFFIX:
                    type_desc = "PDF文档"
                    is_supported = True

                if not is_supported:
                    logger.info(
                        f"[文件管理] 文件 {original_name} 格式不支持 ({file_suffix})，跳过处理"
                    )
                    continue

                if not should_expose or not req.func_tool:
                    logger.info(
                        "[文件管理] 文件 %s 已保存为 %s，但当前轮未暴露文件工具或未附加函数工具，跳过上传文件提示注入",
                        original_name,
                        stored_path.name,
                    )
                    continue

                prompt = NOTICE_UPLOADED_FILE_TEMPLATE.format(
                    type_desc=type_desc,
                    original_name=original_name,
                    file_suffix=file_suffix,
                    stored_name=stored_path.name,
                )

                req.system_prompt = (req.system_prompt or "") + prompt
                logger.info(
                    f"[文件管理] 收到文件 {original_name}，已保存为 {stored_path.name}。"
                )
            except Exception as exc:
                logger.error(f"[文件管理] 处理上传文件失败: {exc}")
