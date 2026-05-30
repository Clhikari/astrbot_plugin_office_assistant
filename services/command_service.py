from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..constants import DOC_COMMAND_TRIGGER_EVENT_KEY
from ..constants import ALL_OFFICE_SUFFIXES
from ..domain.document.render_backends import NodeDocumentRenderBackend
from ..utils import format_file_size

if TYPE_CHECKING:
    from .image_asset_service import ImageAssetService


class CommandService:
    def __init__(
        self,
        *,
        workspace_service,
        pdf_converter,
        plugin_data_path: Path,
        auto_delete: bool,
        allow_external_input_files: bool,
        enable_features_in_group: bool,
        auto_block_execution_tools: bool,
        allow_local_excel_script: bool,
        reply_to_user: bool,
        upload_session_service,
        image_asset_service: ImageAssetService,
        is_group_feature_enabled,
        is_all_users_allowed,
        check_permission,
        group_feature_disabled_error,
        node_renderer_entry: str = "",
    ) -> None:
        self._workspace_service = workspace_service
        self._pdf_converter = pdf_converter
        self._plugin_data_path = plugin_data_path
        self._auto_delete = auto_delete
        self._allow_external_input_files = allow_external_input_files
        self._enable_features_in_group = enable_features_in_group
        self._auto_block_execution_tools = auto_block_execution_tools
        self._allow_local_excel_script = allow_local_excel_script
        self._reply_to_user = reply_to_user
        self._upload_session_service = upload_session_service
        self._image_asset_service = image_asset_service
        self._is_group_feature_enabled = is_group_feature_enabled
        self._is_all_users_allowed = is_all_users_allowed
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error
        self._node_renderer_entry = node_renderer_entry

    def delete_file(self, event, message_text: str) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        parts = message_text.strip().split(maxsplit=1)
        if len(parts) < 2:
            return "❌ 用法: /delete_file 文件名"

        filename = parts[1].strip()
        display_name = self._workspace_service.display_name(filename)
        valid, file_path, error = self._workspace_service.validate_path(filename)
        if not valid:
            return f"❌ {error}"

        if not file_path.exists():
            return f"错误：找不到文件 '{display_name}'"

        try:
            file_path.unlink(missing_ok=True)
            return f"成功：文件 '{display_name}' 已删除。"
        except IsADirectoryError:
            return f"'{display_name}'是目录,拒绝删除"
        except PermissionError:
            return "❌ 权限不足，无法删除文件"
        except Exception as exc:
            return f"删除文件时发生错误{exc}"

    def fileinfo(self, event) -> str:
        if not self._is_group_feature_enabled(event):
            return "❌ " + self._group_feature_disabled_error()

        storage_mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"
        pdf_caps = self._pdf_converter.capabilities
        pdf_status = []
        if pdf_caps.get("office_to_pdf"):
            pdf_status.append("Office→PDF ✓")
        else:
            pdf_status.append("Office→PDF ✗ (需要LibreOffice)")
        if pdf_caps.get("pdf_to_word"):
            pdf_status.append("PDF→Word ✓")
        else:
            pdf_status.append("PDF→Word ✗ (需要pdf2docx)")
        if pdf_caps.get("pdf_to_excel"):
            pdf_status.append("PDF→Excel ✓")
        else:
            pdf_status.append("PDF→Excel ✗ (需要tabula-py)")
        word_toolchain_status = self._build_word_toolchain_status()

        return (
            "📂 AstrBot 文件操作工具\n"
            f"存储模式: {storage_mode}\n"
            f"工作目录: {self._plugin_data_path}\n"
            f"外部路径读取: {'开启' if self._allow_external_input_files else '关闭'}\n"
            f"群聊启用插件功能: {'开启' if self._enable_features_in_group else '关闭'}\n"
            f"允许所有用户使用: {'开启' if self._is_all_users_allowed() else '关闭'}\n"
            f"自动屏蔽 shell/python 工具: {'开启' if self._auto_block_execution_tools else '关闭'}\n"
            f"允许本地 Excel 脚本工具: {'开启' if self._allow_local_excel_script else '关闭'}\n"
            f"回复模式: {'开启' if self._reply_to_user else '关闭'}\n"
            f"Word工具链: {word_toolchain_status}\n"
            f"PDF转换: {', '.join(pdf_status)}"
        )

    def _build_word_toolchain_status(self) -> str:
        backend = NodeDocumentRenderBackend(self._node_renderer_entry)
        if backend.is_available():
            return f"✅ Node 渲染可用 ({backend.entry_path})"
        return f"❌ Node 渲染不可用 ({backend.entry_path}，需要 node 和渲染器入口)"

    def list_files(self, event) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        files = [
            file_path
            for file_path in self._plugin_data_path.glob("*")
            if file_path.is_file() and file_path.suffix.lower() in ALL_OFFICE_SUFFIXES
        ]
        if not files:
            result = "文件库当前没有 Office 文件"
            if self._auto_delete:
                result += "（自动删除模式已开启，文件发送后会自动清理）"
            return result

        files.sort(key=lambda item: item.stat().st_mtime, reverse=True)
        lines = ["📂 机器人工作区 Office 文件列表："]
        if self._auto_delete:
            lines.append("⚠️ 自动删除模式已开启")
        for file_path in files:
            lines.append(
                f"- {file_path.name} ({format_file_size(file_path.stat().st_size)})"
            )
        return "\n".join(lines)

    def pdf_status(self, event) -> str:
        if not self._is_group_feature_enabled(event):
            return "❌ " + self._group_feature_disabled_error()

        status = self._pdf_converter.get_detailed_status()
        caps = status["capabilities"]
        missing = self._pdf_converter.get_missing_dependencies()
        lines = ["📄 PDF 转换功能状态\n"]

        lines.append("【功能可用性】")
        office_status = "✅ 可用" if caps["office_to_pdf"] else "❌ 不可用"
        if status["office_to_pdf_backend"]:
            office_status += f" ({status['office_to_pdf_backend']})"
        lines.append(f"  Office→PDF: {office_status}")

        word_status = "✅ 可用" if caps["pdf_to_word"] else "❌ 不可用"
        if status["word_backend"]:
            word_status += f" ({status['word_backend']})"
        lines.append(f"  PDF→Word:   {word_status}")

        excel_status = "✅ 可用" if caps["pdf_to_excel"] else "❌ 不可用"
        if status["excel_backend"]:
            excel_status += f" ({status['excel_backend']})"
        lines.append(f"  PDF→Excel:  {excel_status}")

        lines.append("\n【环境检测】")
        lines.append(f"  平台: {'Windows' if status['is_windows'] else 'Linux/macOS'}")
        lines.append(
            f"  Java: {'✅ 可用' if status['java_available'] else '❌ 不可用'}"
        )
        if status["libreoffice_path"]:
            lines.append(f"  LibreOffice: {status['libreoffice_path']}")

        libs = status["libs"]
        installed = [name for name, is_installed in libs.items() if is_installed]
        if installed:
            lines.append(f"\n【已安装库】 {', '.join(installed)}")

        if missing:
            lines.append("\n【缺失依赖】")
            for dependency in missing:
                lines.append(f"  • {dependency}")
        else:
            lines.append("\n✅ 所有依赖已安装")

        return "\n".join(lines)

    def doc_list(self, event) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        available_files = self._upload_session_service.list_session_upload_infos(event)
        return self._format_doc_list(available_files)

    def doc_clear(self, event, file_id: str = "") -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        available_files = self._upload_session_service.list_session_upload_infos(event)
        if not available_files:
            return "❌ 当前没有可处理的上传文件。"

        cleared_count = self._upload_session_service.clear_session_upload_infos(
            event,
            file_id=file_id.strip() or None,
        )
        if cleared_count == 0:
            return "❌ 当前没有匹配的上传文件可清除。"
        if file_id.strip():
            return f"✅ 已清除文件 {file_id.strip()}。"
        return f"✅ 已清除 {cleared_count} 个待处理文件。"

    async def doc_use(self, event, selection: str) -> str | None:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        available_files = self._upload_session_service.list_session_upload_infos(event)
        if not available_files:
            return "❌ 当前没有可处理的上传文件，请先上传文件。"

        selected_infos, normalized_ids, normalized_instruction = (
            self._parse_doc_use_selection(available_files, selection.strip())
        )

        if selected_infos is None:
            return self._format_doc_selection_help(available_files)

        if not normalized_instruction:
            selected_label = " ".join(normalized_ids) if normalized_ids else "文件ID"
            return f"❌ 用法: /doc use {selected_label} 你的要求"

        await self._requeue_doc_request(
            event,
            upload_infos=selected_infos,
            user_instruction=normalized_instruction,
        )
        return None

    def _parse_doc_use_selection(
        self,
        available_files: list[dict],
        raw_selection: str,
    ) -> tuple[list[dict] | None, list[str], str]:
        available_by_id = {
            str(info.get("file_id", "")).lower(): info for info in available_files
        }
        tokens = [token for token in raw_selection.split() if token]
        selected_ids: list[str] = []
        instruction_tokens: list[str] = []

        for index, token in enumerate(tokens):
            normalized = token.strip().lower()
            if normalized in available_by_id:
                if normalized not in selected_ids:
                    selected_ids.append(normalized)
                continue

            instruction_tokens = tokens[index:]
            break

        if not selected_ids:
            if len(available_files) == 1:
                return [available_files[0]], [], raw_selection
            return None, [], ""

        selected_infos = [available_by_id[file_id] for file_id in selected_ids]
        instruction = " ".join(instruction_tokens).strip()
        return selected_infos, selected_ids, instruction

    def _format_doc_list(self, upload_infos: list[dict]) -> str:
        if not upload_infos:
            return "当前没有可处理的上传文件。"
        lines = ["当前可用文件："]
        for info in upload_infos:
            file_id = info.get("file_id", "unknown")
            original_name = info.get("original_name", "未命名文件")
            lines.append(f"- [{file_id}] {original_name}")
        return "\n".join(lines)

    def _format_doc_selection_help(self, upload_infos: list[dict]) -> str:
        lines = [self._format_doc_list(upload_infos)]
        lines.append("")
        lines.append("请使用 `/doc use 文件ID 你的要求` 指定要处理的文件。")
        lines.append("需要多个文件时，可以连续写多个文件ID。")
        lines.append("例如：/doc use f2 根据这份文件整理成正式汇报")
        lines.append("例如：/doc use f1 f2 根据这些文件整理成正式汇报")
        lines.append("也可以使用 `/doc list` 查看文件，或 `/doc clear` 清空当前缓存。")
        return "\n".join(lines)

    def _require_access(self, event) -> str | None:
        if not self._is_group_feature_enabled(event):
            return "❌ " + self._group_feature_disabled_error()
        if not self._check_permission(event):
            return "❌ 权限不足"
        return None

    async def _requeue_doc_request(
        self, event, *, upload_infos: list[dict], user_instruction: str
    ) -> None:
        event.set_extra(DOC_COMMAND_TRIGGER_EVENT_KEY, True)
        await self._upload_session_service.requeue_upload_request(
            event,
            upload_infos=upload_infos,
            user_instruction=user_instruction,
        )

    # ── /img commands ──

    def img_add(
        self,
        event,
        *,
        source_items: list[tuple[Path, str]],
        selection: str = "",
    ) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        if not source_items:
            return "当前没有待注册的图片。请在消息中附带图片并发送 /img add，或先上传图片再使用 /img add。"

        session_key = self._upload_session_service.get_attachment_session_key(event)

        parts = selection.strip().split(maxsplit=1) if selection.strip() else []
        selector = parts[0] if parts else "all"
        note = parts[1] if len(parts) > 1 else ""

        if len(source_items) == 1 and selection.strip() and not note:
            note = selection.strip()
            targets = list(enumerate(source_items, 1))
        elif selector == "all":
            targets = list(enumerate(source_items, 1))
        else:
            try:
                idx = int(selector)
            except ValueError:
                note = selection.strip()
                targets = list(enumerate(source_items, 1))
            else:
                if idx < 1 or idx > len(source_items):
                    return f"序号超出范围。当前有 {len(source_items)} 张待注册图片。"
                targets = [(idx, source_items[idx - 1])]

        registered = []
        errors = []
        for idx, (path, original_name) in targets:
            try:
                info = self._image_asset_service.register_image(
                    path,
                    session_key=session_key,
                    note=note,
                    original_name=original_name,
                )
                registered.append(info)
            except (ValueError, FileNotFoundError) as exc:
                errors.append(f"图片 {idx}: {exc}")

        lines = []
        if registered:
            lines.append(f"已注册 {len(registered)} 张图片：")
            for info in registered:
                size_str = format_file_size(info["size_bytes"])
                lines.append(
                    f"  {info['ref']} ({info['width']}x{info['height']}, {size_str})"
                )
        if errors:
            lines.append("注册失败：")
            lines.extend(f"  {e}" for e in errors)
        return "\n".join(lines)

    def img_list(self, event) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        session_key = self._upload_session_service.get_attachment_session_key(event)
        images = self._image_asset_service.list_images(session_key)

        if not images:
            return "当前会话没有已注册的图片。使用 /img add 注册上传的图片。"

        lines = [f"📷 当前会话图片（共 {len(images)} 张）："]
        for i, info in enumerate(images, 1):
            size_str = format_file_size(info["size_bytes"])
            note_str = f" — {info['note']}" if info["note"] else ""
            detail_parts = []
            if info.get("original_name"):
                detail_parts.append(info["original_name"])
            detail_parts.append(f"{info['width']}x{info['height']}")
            detail_parts.append(size_str)
            lines.append(
                f"  {i}. {info['ref']}\n     {' | '.join(detail_parts)}{note_str}"
            )
        return "\n".join(lines)

    def img_note(self, event, ref: str, note: str) -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        ref = ref.strip()
        note = note.strip()
        if not ref or not note:
            return "用法: /img note <引用或序号> <新备注>"

        session_key = self._upload_session_service.get_attachment_session_key(event)
        images = self._image_asset_service.list_images(session_key)

        resolved = self._resolve_img_ref(ref, images)
        if resolved is None:
            return f"未找到图片: {ref}"

        if self._image_asset_service.update_note(
            resolved, note, session_key=session_key
        ):
            return f"已更新备注: {resolved} — {note}"
        return f"更新失败: {resolved}"

    def img_clear(self, event, target: str = "") -> str:
        access_error = self._require_access(event)
        if access_error:
            return access_error

        session_key = self._upload_session_service.get_attachment_session_key(event)
        target = target.strip()

        if not target or target == "all":
            count = self._image_asset_service.clear_images(session_key=session_key)
            return f"已清理 {count} 张图片。" if count else "当前会话没有图片可清理。"

        images = self._image_asset_service.list_images(session_key)
        ref = self._resolve_img_ref(target, images)
        if ref is None:
            return f"未找到图片: {target}"

        count = self._image_asset_service.clear_images(session_key=session_key, ref=ref)
        return f"已清理: {ref}" if count else f"清理失败: {ref}"

    def _resolve_img_ref(self, ref_or_idx: str, images: list[dict]) -> str | None:
        if ref_or_idx.startswith("images/"):
            return (
                ref_or_idx if any(img["ref"] == ref_or_idx for img in images) else None
            )
        try:
            idx = int(ref_or_idx)
            if 1 <= idx <= len(images):
                return images[idx - 1]["ref"]
        except ValueError:
            pass
        return None
