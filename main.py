from email import message
from pathlib import Path
from datetime import datetime
import base64
from typing import Optional

from astrbot.api.event import filter, AstrMessageEvent, MessageEventResult
from astrbot.api.star import Context, Star
from astrbot.api import logger, llm_tool, AstrBotConfig
from astrbot.api.provider import ProviderRequest
from astrbot.core.message.message_event_result import MessageChain
from astrbot.core.platform.message_type import MessageType
from astrbot.core.message.components import At, Reply
import astrbot.api.message_components as Comp
from astrbot.core.utils.astrbot_path import get_astrbot_data_path

# 导入底层生成器
from .file_generator import FileGenerator
from .office_generator import OfficeGenerator
from .utils import format_file_size


class FileOperationPlugin(Star):
    """基于工具调用的智能文件管理插件"""

    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config
        self.plugin_data_path = (
            Path(get_astrbot_data_path()) / "plugins_file_operation_tool"
        )

        Store_files = strself.plugin_data_path + "Store files"

        self.file_gen = FileGenerator(self.plugin_data_path)
        self.office_gen = OfficeGenerator(self.plugin_data_path)

        self.FILE_TOOLS = ["list_files", "read_file", "write_file", "delete_file"]
        logger.info(f"[文件管理] 插件加载完成。数据目录: {self.plugin_data_path}")

    def _check_permission(self, event: AstrMessageEvent) -> bool:
        """检查用户权限"""
        logger.info("正在检查用户权限")
        perm_cfg = self.config.get("permission_settings", {})

        # 管理员检查
        if perm_cfg.get("require_admin", False) and not event.is_admin():
            return False

        # 白名单检查
        whitelist = perm_cfg.get("whitelist_users", [])
        if whitelist:
            user_id = str(event.get_sender_id())
            if user_id not in [str(u) for u in whitelist]:
                return False

        return True

    def _is_bot_mentioned(self, event: AstrMessageEvent) -> bool:
        """检查是否被@/回复"""
        try:
            bot_id = str(event.message_obj.self_id)
            for segment in event.message_obj.message:
                if isinstance(segment, At) or isinstance(segment, Reply):
                    target_id = getattr(segment, "qq", None) or getattr(
                        segment, "target", None
                    )
                    if target_id and str(target_id) == bot_id:
                        return True
            return False
        except Exception as e:
            logger.error(f"未知错误{e}")
            return False

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        """动态控制工具可见性"""
        trigger_cfg = self.config.get("trigger_settings", {})

        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        should_expose = True

        # 权限拦截
        if not self._check_permission(event):
            should_expose = False
        # 群聊@/回复拦截
        elif (
            is_group
            and trigger_cfg.get("require_at_in_group", True)
            and not self._is_bot_mentioned(event)
        ):
            should_expose = False

        if not should_expose and req.func_tool:
            for tool_name in self.FILE_TOOLS:
                req.func_tool.remove_tool(tool_name)

    @llm_tool(name="list_files")
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人文件库中的所有文件。"""
        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return ""
        try:
            files = [f for f in self.plugin_data_path.glob("*") if f.is_file()]
            if not files:
                await event.send(MessageChain().message("文件库当前为空"))
                return ""
            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            res = ["📂 机器人工作区文件列表："]
            for f in files:
                res.append(f"- {f.name} ({format_file_size(f.stat().st_size)})")
            return "\n".join(res)
        except Exception as e:
            logger.error(f"获取列表失败: {e}")
            await event.send(MessageChain().message("获取列表失败喵"))

    @llm_tool(name="read_file")
    async def read_file(self, event: AstrMessageEvent, filename: str) -> str:
        """读取并查看文件内容。"""
        if not self._check_permission(event):
            return "拒绝访问：权限不足。"
        file_path = self.plugin_data_path / filename
        if not file_path.exists():
            await event.send(MessageChain().message("文件不存在，请检车"))
            return f"错误：文件 {filename} 不存在。"

        try:
            suffix = file_path.suffix.lower()
            text_suffixes = {
                ".txt",
                ".md",
                ".py",
                ".js",
                ".ts",
                ".json",
                ".csv",
                ".html",
                ".css",
                ".yaml",
                ".yml",
                ".sql",
                ".sh",
                ".bat",
                ".c",
                ".cpp",
                ".java",
            }
            if suffix in text_suffixes:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    return f"内容:\n{f.read()}"
            return "该文件为二进制格式，无法直接读取。"
        except Exception as e:
            return f"读取失败: {e}"

    @llm_tool(name="write_file")
    async def write_file(
        self,
        event: AstrMessageEvent,
        filename: str,
        content: str,
        file_type: str = "text",
    ) -> str:
        """在机器人工作区中创建或更新文件。"""
        if not self._check_permission(event):
            return "拒绝访问：权限不足。"

        if file_type.lower() in ["word", "excel", "powerpoint"] and not self.config.get(
            "feature_settings", {}
        ).get("enable_office_files", True):
            return "错误：当前配置禁用了 Office 文件生成功能。"

        file_info = {
            "type": file_type.lower(),
            "filename": filename,
            "content": content,
        }
        try:
            if file_info["type"] in ["word", "excel", "powerpoint"]:
                file_path = await self.office_gen.generate(file_info["type"], file_info)
            else:
                file_path = await self.file_gen.generate(file_info)

            if file_path and file_path.exists():
                with open(file_path, "rb") as f:
                    b64_str = base64.b64encode(f.read()).decode("utf-8")

                chain = [
                    Comp.Plain(f"✅ 文件已处理成功：{file_path.name}"),
                    Comp.File(file=f"base64://{b64_str}", name=file_path.name),
                ]
                use_reply = self.config.get("trigger_settings", {}).get(
                    "reply_to_user", True
                )
                await event.send(event.chain_result(chain) if use_reply else chain)
                return f"成功：文件 '{file_path.name}' 已发送。"
            return "生成文件失败。"
        except Exception as e:
            return f"文件操作异常: {e}"

    @llm_tool(name="delete_file")
    async def delete_file(self, event: AstrMessageEvent, filename: str) -> str:
        """从工作区中永久删除指定文件。"""
        if not self._check_permission(event):
            return "拒绝访问：权限不足。"
        file_path = self.plugin_data_path / filename
        if file_path.exists():
            try:
                file_path.unlink()
                return f"成功：文件 '{filename}' 已删除。"
            except Exception as e:
                return f"删除失败: {str(e)}"
        return f"错误：找不到文件 '{filename}'。"

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        yield event.plain_result(
            "📂 AstrBot 文件操作工具\n"
            f"工作目录: {self.plugin_data_path}\n"
            f"回复模式: {'开启' if self.config.get('trigger_settings', {}).get('reply_to_user') else '关闭'}"
        )
