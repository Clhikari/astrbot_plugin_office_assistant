import copy
import time
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.api.star import Context

from ..constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX, TEXT_SUFFIXES
from ..message_buffer import BufferedMessage
from .upload_prompt_service import UploadInfo, UploadPromptService

EVENT_UPLOAD_CACHE_ATTR = "_office_assistant_uploaded_files"


class UploadSessionService:
    def __init__(
        self,
        *,
        context: Context,
        recent_text_ttl_seconds: int,
        upload_session_ttl_seconds: int,
        recent_text_max_entries: int,
        recent_text_cleanup_interval_seconds: int,
        upload_session_cleanup_interval_seconds: int,
        extract_upload_source,
        store_uploaded_file,
        allow_external_input_files: bool,
    ) -> None:
        self._context = context
        self._recent_text_ttl_seconds = recent_text_ttl_seconds
        self._upload_session_ttl_seconds = upload_session_ttl_seconds
        self._recent_text_max_entries = recent_text_max_entries
        self._recent_text_cleanup_interval_seconds = (
            recent_text_cleanup_interval_seconds
        )
        self._upload_session_cleanup_interval_seconds = (
            upload_session_cleanup_interval_seconds
        )
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._allow_external_input_files = allow_external_input_files
        self._prompt_service = UploadPromptService(
            allow_external_input_files=allow_external_input_files
        )
        self._recent_text_by_session: dict[tuple[str, str, str], tuple[str, float]] = {}
        self._recent_text_last_cleanup_ts = 0.0
        self._session_uploads_by_session: dict[
            tuple[str, str, str], list[tuple[UploadInfo, float]]
        ] = {}
        self._session_uploads_last_cleanup_ts = 0.0

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

    def _cleanup_session_upload_cache(self, now: float, *, force: bool = False) -> None:
        if (
            not force
            and now - self._session_uploads_last_cleanup_ts
            < self._upload_session_cleanup_interval_seconds
        ):
            return

        expire_before = now - self._upload_session_ttl_seconds
        expired_keys: list[tuple[str, str, str]] = []
        for session_key, items in self._session_uploads_by_session.items():
            kept = [(info, ts) for info, ts in items if ts > expire_before]
            if kept:
                self._session_uploads_by_session[session_key] = kept
            else:
                expired_keys.append(session_key)

        for session_key in expired_keys:
            self._session_uploads_by_session.pop(session_key, None)

        self._session_uploads_last_cleanup_ts = now

    def _allocate_file_id(self, session_key: tuple[str, str, str]) -> str:
        existing_ids = {
            info.get("file_id", "")
            for info, _ in self._session_uploads_by_session.get(session_key, [])
            if info.get("file_id")
        }
        next_index = 1
        while True:
            candidate = f"f{next_index}"
            if candidate not in existing_ids:
                return candidate
            next_index += 1

    def _cache_session_upload_infos(
        self,
        event: AstrMessageEvent,
        upload_infos: list[UploadInfo],
    ) -> list[UploadInfo]:
        now = time.time()
        self._cleanup_session_upload_cache(now)
        session_key = self.get_attachment_session_key(event)
        cached_infos = self._session_uploads_by_session.setdefault(session_key, [])

        for info in upload_infos:
            file_id = self._allocate_file_id(session_key)
            info["file_id"] = file_id
            cached_infos.append((dict(info), now))

        return upload_infos

    def _get_wake_prefix(self, event: AstrMessageEvent) -> str:
        get_config = getattr(self._context, "get_config", None)
        config = None
        if callable(get_config):
            try:
                config = get_config(event.unified_msg_origin)
            except TypeError:
                config = get_config()

        if not isinstance(config, dict):
            legacy_config = getattr(self._context, "astrbot_config", None)
            config = legacy_config if isinstance(legacy_config, dict) else {}

        wake_prefixes = config.get("wake_prefix", [])
        if isinstance(wake_prefixes, str):
            wake_prefixes = [wake_prefixes]

        for wake_prefix in wake_prefixes:
            if isinstance(wake_prefix, str) and wake_prefix:
                return wake_prefix
        return "/"

    def list_session_upload_infos(self, event: AstrMessageEvent) -> list[UploadInfo]:
        now = time.time()
        self._cleanup_session_upload_cache(now)
        session_key = self.get_attachment_session_key(event)
        return [
            dict(info)
            for info, _ in self._session_uploads_by_session.get(session_key, [])
        ]

    def get_session_upload_info(
        self,
        event: AstrMessageEvent,
        file_id: str,
    ) -> UploadInfo | None:
        normalized_id = file_id.strip().lower()
        if not normalized_id:
            return None

        for info in self.list_session_upload_infos(event):
            if info.get("file_id", "").lower() == normalized_id:
                return info
        return None

    def clear_session_upload_infos(
        self,
        event: AstrMessageEvent,
        *,
        file_id: str | None = None,
    ) -> int:
        now = time.time()
        self._cleanup_session_upload_cache(now)
        session_key = self.get_attachment_session_key(event)
        items = self._session_uploads_by_session.get(session_key, [])
        if not items:
            return 0

        if file_id is None:
            removed = len(items)
            self._session_uploads_by_session.pop(session_key, None)
            return removed

        normalized_id = file_id.strip().lower()
        kept = [
            (info, ts)
            for info, ts in items
            if info.get("file_id", "").lower() != normalized_id
        ]
        removed = len(items) - len(kept)
        if kept:
            self._session_uploads_by_session[session_key] = kept
        else:
            self._session_uploads_by_session.pop(session_key, None)
        return removed

    async def requeue_upload_request(
        self,
        event: AstrMessageEvent,
        *,
        upload_infos: list[UploadInfo],
        user_instruction: str,
        file_components: list | None = None,
        reentry_count: int = 0,
    ) -> None:
        prompt_text = self._prompt_service.build_prompt(
            upload_infos=upload_infos,
            user_instruction=user_instruction,
        )
        await self._requeue_event(
            event,
            prompt_text=prompt_text,
            upload_infos=upload_infos,
            file_components=file_components,
            reentry_count=reentry_count,
            force_command_wake=True,
        )

    async def requeue_buffered_upload_request(
        self,
        event: AstrMessageEvent,
        *,
        upload_infos: list[UploadInfo],
        user_instruction: str,
        file_components: list | None = None,
        reentry_count: int = 0,
    ) -> None:
        prompt_text = self._prompt_service.build_prompt(
            upload_infos=upload_infos,
            user_instruction=user_instruction,
        )
        await self._requeue_event(
            event,
            prompt_text=prompt_text,
            upload_infos=upload_infos,
            file_components=file_components,
            reentry_count=reentry_count,
            force_command_wake=False,
        )

    async def _requeue_event(
        self,
        event: AstrMessageEvent,
        *,
        prompt_text: str,
        upload_infos: list[UploadInfo],
        file_components: list | None,
        reentry_count: int,
        force_command_wake: bool,
    ) -> None:
        requeued_event = copy.copy(event)
        requeued_event._extras = dict(getattr(event, "_extras", {}))

        new_chain = [Comp.Plain(prompt_text)]
        for file_component in file_components or []:
            new_chain.append(file_component)

        requeued_message_obj = copy.copy(event.message_obj)
        requeued_message_obj.message = list(new_chain)
        requeued_event.message_obj = requeued_message_obj
        if force_command_wake:
            wake_prefix = self._get_wake_prefix(event)
            requeued_event.message_str = f"{wake_prefix}{prompt_text.strip()}"
        else:
            requeued_event.message_str = prompt_text.strip()
        setattr(requeued_event, EVENT_UPLOAD_CACHE_ATTR, upload_infos)

        logger.info(f"[消息缓冲] 已合并消息，提示: {prompt_text[:50]}...")

        try:
            setattr(requeued_event, "_buffered", True)
            setattr(requeued_event, "_buffer_reentry_count", reentry_count + 1)
            requeued_event._result = None
            if force_command_wake:
                requeued_event.is_wake = True
                requeued_event.is_at_or_wake_command = True
            event_queue = self._context.get_event_queue()
            await event_queue.put(requeued_event)
            logger.debug("[消息缓冲] 事件已重新放入队列")
        except Exception as exc:
            logger.error(f"[消息缓冲] 重新分发事件失败: {exc}")

    def remember_recent_text(self, event: AstrMessageEvent) -> None:
        return None

    def pop_recent_text(self, event: AstrMessageEvent) -> str:
        return ""

    def get_cached_upload_infos(self, event: AstrMessageEvent) -> list[UploadInfo]:
        cached = getattr(event, EVENT_UPLOAD_CACHE_ATTR, None)
        if isinstance(cached, list):
            return cached
        return []

    def _resolve_upload_type(self, filename: str) -> tuple[str, bool]:
        suffix = Path(filename).suffix.lower() if filename else ""
        if suffix in ALL_OFFICE_SUFFIXES:
            return "Office文档 (Word/Excel/PPT)", True
        if suffix in TEXT_SUFFIXES:
            return "文本/代码文件", True
        if suffix == PDF_SUFFIX:
            return "PDF文档", True
        return "", False

    async def _ensure_upload_infos(
        self,
        event: AstrMessageEvent,
        files: list,
    ) -> list[UploadInfo]:
        cached = self.get_cached_upload_infos(event)
        if cached:
            return cached

        upload_infos: list[UploadInfo] = []
        for file_component in files:
            name = ""
            if isinstance(file_component, Comp.File):
                name = file_component.name or ""
            name = name or "未命名文件"
            type_desc, is_supported = self._resolve_upload_type(name)
            info: UploadInfo = {
                "original_name": name,
                "file_suffix": Path(name).suffix.lower() if name else "",
                "type_desc": type_desc,
                "is_supported": is_supported,
                "stored_name": "",
                "source_path": "",
            }

            if isinstance(file_component, Comp.File):
                try:
                    src_path, original_name = await self._extract_upload_source(
                        file_component
                    )
                    if original_name:
                        info["original_name"] = original_name
                        info["file_suffix"] = Path(original_name).suffix.lower()
                        (
                            info["type_desc"],
                            info["is_supported"],
                        ) = self._resolve_upload_type(original_name)
                    if src_path and src_path.exists():
                        info["source_path"] = str(src_path.resolve())
                        if info["is_supported"]:
                            stored_path = self._store_uploaded_file(
                                src_path, info["original_name"]
                            )
                            info["stored_name"] = stored_path.name
                except Exception as exc:
                    logger.error(f"[消息缓冲] 解析上传文件失败: {exc}")

            upload_infos.append(info)

        upload_infos = self._cache_session_upload_infos(event, upload_infos)
        setattr(event, EVENT_UPLOAD_CACHE_ATTR, upload_infos)
        return upload_infos

    async def on_buffer_complete(self, buf: BufferedMessage) -> None:
        event = buf.event
        files = buf.files
        texts = buf.texts

        buffered_text_count = len(texts)

        reentry_count = getattr(event, "_buffer_reentry_count", 0)
        if reentry_count >= 3:
            logger.warning("[消息缓冲] 事件重入次数过多，停止处理")
            return

        upload_infos = await self._ensure_upload_infos(event, files)

        logger.info(
            "[消息缓冲] 缓冲完成，文件数: %s, 缓冲文本数: %s",
            len(files),
            buffered_text_count,
        )

        user_instruction = " ".join(texts).strip() if texts else ""
        if not user_instruction:
            logger.debug("[消息缓冲] 已缓存上传文件，等待显式命令继续处理")
            return

        await self.requeue_buffered_upload_request(
            event,
            upload_infos=upload_infos,
            user_instruction=user_instruction,
            file_components=files,
            reentry_count=reentry_count,
        )
