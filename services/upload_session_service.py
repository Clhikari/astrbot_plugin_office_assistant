import time
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.api.star import Context

from ..constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX, TEXT_SUFFIXES
from ..message_buffer import BufferedMessage


class UploadSessionService:
    def __init__(
        self,
        *,
        context: Context,
        recent_text_ttl_seconds: int,
        recent_text_max_entries: int,
        recent_text_cleanup_interval_seconds: int,
    ) -> None:
        self._context = context
        self._recent_text_ttl_seconds = recent_text_ttl_seconds
        self._recent_text_max_entries = recent_text_max_entries
        self._recent_text_cleanup_interval_seconds = (
            recent_text_cleanup_interval_seconds
        )
        self._recent_text_by_session: dict[tuple[str, str, str], tuple[str, float]] = {}
        self._recent_text_last_cleanup_ts = 0.0

    @property
    def recent_text_by_session(self) -> dict[tuple[str, str, str], tuple[str, float]]:
        return self._recent_text_by_session

    def get_attachment_session_key(
        self, event: AstrMessageEvent
    ) -> tuple[str, str, str]:
        platform_id = event.get_platform_id() or ""
        sender_id_obj = event.get_sender_id()
        sender_id = str(sender_id_obj) if sender_id_obj is not None else ""
        origin = event.unified_msg_origin or ""

        message_obj = getattr(event, "message_obj", None)
        if message_obj is not None:
            if not platform_id:
                platform_id = str(
                    getattr(message_obj, "platform", "")
                    or getattr(message_obj, "platform_id", "")
                    or getattr(message_obj, "self_id", "")
                )
            if not sender_id:
                sender_id = str(
                    getattr(message_obj, "sender_id", "")
                    or getattr(message_obj, "user_id", "")
                    or getattr(message_obj, "qq", "")
                )
            if not origin:
                origin = str(
                    getattr(message_obj, "session_id", "")
                    or getattr(message_obj, "conversation_id", "")
                    or getattr(message_obj, "group_id", "")
                    or getattr(message_obj, "channel_id", "")
                    or getattr(message_obj, "target_id", "")
                )

        return (
            platform_id or f"unknown_platform:{id(self)}",
            sender_id
            or f"unknown_sender:{id(message_obj) if message_obj else id(event)}",
            origin or f"unknown_origin:{id(event)}",
        )

    def cleanup_recent_text_cache(self, now: float, *, force: bool = False) -> None:
        if (
            not force
            and now - self._recent_text_last_cleanup_ts
            < self._recent_text_cleanup_interval_seconds
            and len(self._recent_text_by_session) <= self._recent_text_max_entries
        ):
            return

        expire_before = now - self._recent_text_ttl_seconds
        expired_keys = [
            key
            for key, (_, ts) in self._recent_text_by_session.items()
            if ts <= expire_before
        ]
        for key in expired_keys:
            self._recent_text_by_session.pop(key, None)

        overflow = len(self._recent_text_by_session) - self._recent_text_max_entries
        if overflow > 0:
            oldest_keys = sorted(
                self._recent_text_by_session.items(), key=lambda item: item[1][1]
            )[:overflow]
            for key, _ in oldest_keys:
                self._recent_text_by_session.pop(key, None)

        self._recent_text_last_cleanup_ts = now

    def remember_recent_text(self, event: AstrMessageEvent) -> None:
        text = str(event.message_str or "").strip()
        if not text or text.startswith("[System Notice]"):
            return
        now = time.time()
        self.cleanup_recent_text_cache(now)
        session_key = self.get_attachment_session_key(event)
        self._recent_text_by_session[session_key] = (text, now)
        if len(self._recent_text_by_session) > self._recent_text_max_entries:
            self.cleanup_recent_text_cache(now, force=True)

    def pop_recent_text(self, event: AstrMessageEvent) -> str:
        now = time.time()
        self.cleanup_recent_text_cache(now)
        session_key = self.get_attachment_session_key(event)
        item = self._recent_text_by_session.get(session_key)
        if not item:
            return ""

        text, ts = item
        if now - ts > self._recent_text_ttl_seconds:
            self._recent_text_by_session.pop(session_key, None)
            return ""

        self._recent_text_by_session.pop(session_key, None)
        return text

    async def on_buffer_complete(self, buf: BufferedMessage) -> None:
        event = buf.event
        files = buf.files
        texts = buf.texts

        logger.info(f"[消息缓冲] 缓冲完成，文件数: {len(files)}, 文本数: {len(texts)}")

        reentry_count = getattr(event, "_buffer_reentry_count", 0)
        if reentry_count >= 3:
            logger.warning("[消息缓冲] 事件重入次数过多，停止处理")
            return

        file_info_list = []
        has_readable_file = False
        for file_component in files:
            name = ""
            if isinstance(file_component, Comp.File):
                name = file_component.name or ""
            name = name or "未命名文件"
            suffix = Path(name).suffix.lower() if name else ""
            file_info_list.append(f"文件名: {name} (类型: {suffix})")
            if (
                suffix in ALL_OFFICE_SUFFIXES
                or suffix in TEXT_SUFFIXES
                or suffix == PDF_SUFFIX
            ):
                has_readable_file = True

        user_instruction = " ".join(texts) if texts else ""
        if not user_instruction:
            user_instruction = self.pop_recent_text(event)

        if has_readable_file and user_instruction:
            prompt_text = (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                "\n"
                "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                "\n"
                f"[用户指令]\n"
                f"{user_instruction}\n"
                "\n"
                "[处理建议]\n"
                "1. 优先围绕这些上传文件完成用户请求。\n"
                "2. 如果后续系统提示提供了工作区文件名，按该文件名处理。\n"
                "3. 所有面向用户的回复 MUST 使用中文。"
            )
        elif has_readable_file:
            prompt_text = (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                "\n"
                "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                "\n"
                "[处理建议]\n"
                "1. 用户上传了可读取文件，后续应优先围绕这些文件处理。\n"
                "2. 如果后续系统提示提供了工作区文件名，按该文件名处理。\n"
                "3. 用户意图尚不明确时，再用中文询问用户想要如何处理。"
            )
        else:
            prompt_text = (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                "\n"
                "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                "\n"
                "[操作要求]\n"
                "请根据用户要求处理这些文件，使用中文与用户沟通。"
            )

        new_chain = [Comp.Plain(prompt_text)]
        for file_component in files:
            new_chain.append(file_component)

        event.message_obj.message = new_chain
        if hasattr(event.message_obj, "raw_message"):
            event.message_obj.raw_message = prompt_text
        event.message_str = prompt_text.strip()

        logger.info(f"[消息缓冲] 已合并消息，提示: {prompt_text[:50]}...")

        try:
            setattr(event, "_buffered", True)
            setattr(event, "_buffer_reentry_count", reentry_count + 1)
            event._result = None
            event.is_wake = True
            event.is_at_or_wake_command = True
            event_queue = self._context.get_event_queue()
            await event_queue.put(event)
            logger.debug("[消息缓冲] 事件已重新放入队列")
        except Exception as exc:
            logger.error(f"[消息缓冲] 重新分发事件失败: {exc}")
