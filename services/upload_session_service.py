import time
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.api.star import Context

from ..constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX, TEXT_SUFFIXES
from ..message_buffer import BufferedMessage

EVENT_UPLOAD_CACHE_ATTR = "_office_assistant_uploaded_files"


class UploadSessionService:
    def __init__(
        self,
        *,
        context: Context,
        recent_text_ttl_seconds: int,
        recent_text_max_entries: int,
        recent_text_cleanup_interval_seconds: int,
        extract_upload_source,
        store_uploaded_file,
        allow_external_input_files: bool,
    ) -> None:
        self._context = context
        self._recent_text_ttl_seconds = recent_text_ttl_seconds
        self._recent_text_max_entries = recent_text_max_entries
        self._recent_text_cleanup_interval_seconds = (
            recent_text_cleanup_interval_seconds
        )
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._allow_external_input_files = allow_external_input_files
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

    def get_cached_upload_infos(self, event: AstrMessageEvent) -> list[dict]:
        cached = getattr(event, EVENT_UPLOAD_CACHE_ATTR, None)
        if isinstance(cached, list):
            return cached
        return []

    async def _ensure_upload_infos(
        self,
        event: AstrMessageEvent,
        files: list,
    ) -> list[dict]:
        cached = self.get_cached_upload_infos(event)
        if cached:
            return cached

        upload_infos: list[dict] = []
        for file_component in files:
            name = ""
            if isinstance(file_component, Comp.File):
                name = file_component.name or ""
            name = name or "未命名文件"
            suffix = Path(name).suffix.lower() if name else ""
            info = {
                "original_name": name,
                "file_suffix": suffix,
                "type_desc": "",
                "is_supported": False,
                "stored_name": "",
                "source_path": "",
            }

            if suffix in ALL_OFFICE_SUFFIXES:
                info["type_desc"] = "Office文档 (Word/Excel/PPT)"
                info["is_supported"] = True
            elif suffix in TEXT_SUFFIXES:
                info["type_desc"] = "文本/代码文件"
                info["is_supported"] = True
            elif suffix == PDF_SUFFIX:
                info["type_desc"] = "PDF文档"
                info["is_supported"] = True

            if info["is_supported"] and isinstance(file_component, Comp.File):
                try:
                    src_path, original_name = await self._extract_upload_source(
                        file_component
                    )
                    if original_name:
                        info["original_name"] = original_name
                        info["file_suffix"] = Path(original_name).suffix.lower()
                    if src_path and src_path.exists():
                        info["source_path"] = str(src_path.resolve())
                        stored_path = self._store_uploaded_file(
                            src_path, info["original_name"]
                        )
                        info["stored_name"] = stored_path.name
                except Exception as exc:
                    logger.error(f"[消息缓冲] 解析上传文件失败: {exc}")

            upload_infos.append(info)

        setattr(event, EVENT_UPLOAD_CACHE_ATTR, upload_infos)
        return upload_infos

    async def on_buffer_complete(self, buf: BufferedMessage) -> None:
        event = buf.event
        files = buf.files
        texts = buf.texts

        logger.info(f"[消息缓冲] 缓冲完成，文件数: {len(files)}, 文本数: {len(texts)}")

        reentry_count = getattr(event, "_buffer_reentry_count", 0)
        if reentry_count >= 3:
            logger.warning("[消息缓冲] 事件重入次数过多，停止处理")
            return

        upload_infos = await self._ensure_upload_infos(event, files)
        file_info_list = []
        has_readable_file = False
        for info in upload_infos:
            file_lines = [
                f"原始文件名: {info['original_name']} (类型: {info['file_suffix']})"
            ]
            if info["stored_name"]:
                file_lines.append(f"  工作区文件名: {info['stored_name']}")
            if self._allow_external_input_files and info["source_path"]:
                file_lines.append(f"  外部绝对路径: {info['source_path']}")
            file_info_list.append("\n".join(file_lines))
            if info["is_supported"]:
                has_readable_file = True

        user_instruction = " ".join(texts) if texts else ""
        if not user_instruction:
            user_instruction = self.pop_recent_text(event)

        relative_path_guidance = "3. 若使用相对路径，请使用上面的工作区文件名。\\n"
        if self._allow_external_input_files:
            relative_path_guidance = (
                "3. 若使用相对路径，请使用上面的工作区文件名；"
                "如果已提供外部绝对路径，则可直接使用该绝对路径。\\n"
            )

        if has_readable_file and user_instruction:
            prompt_text = (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                + "\n"
                + "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                + "\n"
                + "[用户指令]\n"
                + f"{user_instruction}\n"
                + "\n"
                + "[处理建议]\n"
                + "1. 优先围绕这些上传文件完成用户请求。\n"
                + "2. 先调用 `read_file` 读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
                + relative_path_guidance
                + "4. 所有面向用户的回复 MUST 使用中文。"
            )
        elif has_readable_file:
            prompt_text = (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                + "\n"
                + "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                + "\n"
                + "[处理建议]\n"
                + "1. 用户上传了可读取文件，后续应优先围绕这些文件处理。\n"
                + "2. 如果要读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
                + relative_path_guidance
                + "4. 用户意图尚不明确时，再用中文询问用户想要如何处理。"
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
