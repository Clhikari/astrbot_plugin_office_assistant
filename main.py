from pathlib import Path
from datetime import datetime
from typing import Optional

from astrbot.api.event import filter, AstrMessageEvent, MessageEventResult
from astrbot.api.star import Context, Star, register
from astrbot.api import logger
import astrbot.api.message_components as Comp
from astrbot.core.message.components import At
from astrbot.core.utils.astrbot_path import get_astrbot_data_path
from astrbot.core.platform.message_type import MessageType

# 导入子模块
from .config_manager import ConfigManager
from .file_analyzer import FileAnalyzer
from .file_generator import FileGenerator
from .office_generator import OfficeGenerator
from .utils import format_file_size


@register(
    "file_generator",
    "AI Assistant",
    "智能文件生成器 - 支持Office三件套及多种文件格式",
    "1.0.0",
    "https://github.com/Clhikari/astrbot_plugin_file_generator",
)
class FileGeneratorPlugin(Star):
    """智能文件生成器插件主类"""

    def __init__(self, context: Context):
        super().__init__(context)

        # 初始化插件数据目录
        self.plugin_data_path = (
            Path(get_astrbot_data_path()) / "plugin_data" / "file_generator"
        )
        self.plugin_data_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"[文件生成器] 插件数据目录: {self.plugin_data_path}")

        # 初始化子模块
        self.config_mgr = ConfigManager(self)
        self.analyzer = FileAnalyzer(self.context, self.config_mgr)
        self.file_gen = FileGenerator(self.plugin_data_path)
        self.office_gen = OfficeGenerator(self.plugin_data_path)

        logger.info(f"[文件生成器] 插件加载完成")
        logger.info(
            f"[文件生成器] Office支持: Word={self.office_gen.support['word']}, "
            f"Excel={self.office_gen.support['excel']}, PPT={self.office_gen.support['powerpoint']}"
        )

    def _should_process_message(self, event: AstrMessageEvent) -> bool:
        """判断是否应该处理该消息"""
        message_text = event.message_str.strip()

        # 消息长度检查
        min_length = self.config_mgr.get("min_message_length", 15)
        if len(message_text) < min_length:
            return False

        # 忽略指令消息
        if message_text.startswith("/"):
            return False

        # 权限检查
        if not self._check_permission(event):
            return False

        # 判断消息类型
        is_private = event.message_obj.type == MessageType.FRIEND_MESSAGE
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE

        # 私聊消息
        if is_private:
            return self.config_mgr.get("auto_detect_in_private", True)

        # 群聊消息
        if is_group:
            if not self.config_mgr.get("auto_detect_in_group", False):
                if self.config_mgr.get("require_at_in_group", True):
                    return self._is_bot_mentioned(event)
                return False

            if self.config_mgr.get("require_at_in_group", True):
                return self._is_bot_mentioned(event)

            return True

        return False

    def _check_permission(self, event: AstrMessageEvent) -> bool:
        """检查用户是否有权限使用插件"""
        # 获取配置
        require_admin = self.config_mgr.get("require_admin", False)
        whitelist_config = self.config_mgr.get("whitelist_users", "")

        # 处理白名单数据
        if isinstance(whitelist_config, list):
            # 兼容旧配置或直接列表格式
            whitelist = [str(u) for u in whitelist_config]
        elif isinstance(whitelist_config, str) and whitelist_config.strip():
            # 处理 WebUI 传来的多行文本 (每行一个ID)
            whitelist = [
                line.strip()
                for line in whitelist_config.replace("\r\n", "\n").split("\n")
                if line.strip()
            ]

        # 如果不需要管理员且白名单为空，则所有人可用
        if not require_admin and not whitelist:
            return True

        user_id = str(event.get_sender_id())

        # 检查白名单
        if user_id in whitelist:
            return True

        # 检查管理员权限
        if require_admin and event.is_admin():
            return True

        logger.info(f"[文件生成器] 用户 {user_id} 无权限使用 (不在白名单且非管理员)")
        return False

    def _is_bot_mentioned(self, event: AstrMessageEvent) -> bool:
        """检查消息是否@了机器人"""
        try:
            bot_id = str(event.message_obj.self_id)
            for segment in event.message_obj.message:
                # 检查是否是At类型的消息段
                if isinstance(segment, At):
                    # 使用 getattr 安全地获取属性
                    target_id = getattr(segment, "qq", None) or getattr(
                        segment, "target", None
                    )
                    if target_id and str(target_id) == bot_id:
                        return True
            return False
        except Exception as e:
            logger.error(f"[文件生成器] 检查@失败: {e}")
            return False

    @filter.event_message_type(filter.EventMessageType.ALL)
    async def handle_file_generation(self, event: AstrMessageEvent):
        """处理消息并判断是否需要生成文件"""

        if not self._should_process_message(event):
            return

        try:
            message_text = event.message_str.strip()

            # 使用 AI 分析用户消息
            analysis_result = await self.analyzer.analyze_message(event, message_text)

            if not analysis_result or not analysis_result.get("needs_file", False):
                return

            file_info = analysis_result.get("file_info", {})
            logger.info(f"[文件生成器] 检测到文件生成需求: {file_info}")

            # 生成文件
            file_path = await self._generate_file(event, file_info)

            if file_path and file_path.exists():
                yield await self._send_file(event, file_path, file_info)
            else:
                yield event.plain_result("❌ 文件生成失败，请稍后重试")

        except Exception as e:
            logger.error(f"[文件生成器] 处理消息时出错: {e}", exc_info=True)

    async def _generate_file(
        self, event: AstrMessageEvent, file_info: dict
    ) -> Optional[Path]:
        """生成文件（路由到对应的生成器）"""
        file_type = file_info.get("type", "text").lower()

        # Office文件
        if file_type in ["word", "excel", "powerpoint"]:
            return await self.office_gen.generate(file_type, file_info)

        # 普通文件
        return await self.file_gen.generate(file_info)

    async def _send_file(
        self, event: AstrMessageEvent, file_path: Path, file_info: dict
    ) -> MessageEventResult:
        """发送文件给用户"""
        try:
            description = file_info.get("description", "文件已生成")
            filename = file_path.name
            file_type = file_info.get("type", "未知")

            chain = [
                Comp.Plain(
                    f"✅ {description}\n📄 文件名: {filename}\n📋 类型: {file_type}\n"
                ),
                Comp.File(file=str(file_path), name=filename),
            ]

            return event.chain_result(chain)
        except Exception as e:
            logger.error(f"[文件生成器] 发送文件失败: {e}", exc_info=True)
            return event.plain_result(f"❌ 文件发送失败: {str(e)}")

    @filter.command("genfile", alias={"生成文件", "gf"})
    async def manual_generate_file(
        self, event: AstrMessageEvent, file_type: str = "text", *content_words
    ):
        """手动生成文件指令"""
        # 权限检查
        if not self._check_permission(event):
            yield event.plain_result("❌ 无权限使用此功能")
            return

        if not content_words:
            office_hint = ""
            if self.office_gen.support["word"]:
                office_hint += "word, "
            if self.office_gen.support["excel"]:
                office_hint += "excel, "
            if self.office_gen.support["powerpoint"]:
                office_hint += "powerpoint, "

            yield event.plain_result(
                "📝 使用方法：\n"
                "/genfile <类型> <内容>\n\n"
                "支持的文件类型：\n"
                f"代码: python, javascript, java, cpp, html, css\n"
                f"数据: json, csv, xml, yaml\n"
                f"文档: markdown, text\n"
                f"Office: {office_hint if office_hint else '(未安装相关库)'}\n\n"
                "示例：\n"
                "/genfile python 快速排序算法\n"
                "/genfile word 项目进度报告\n"
                "/genfile excel 销售数据统计表"
            )
            return

        content = " ".join(content_words)

        # 使用 AI 生成完整内容
        full_content = await self.analyzer.generate_content(event, file_type, content)

        # 生成文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_info = {
            "type": file_type,
            "filename": f"manual_{file_type}_{timestamp}",
            "content": full_content,
            "description": f"手动生成的{file_type}文件",
        }

        file_path = await self._generate_file(event, file_info)

        if file_path and file_path.exists():
            yield await self._send_file(event, file_path, file_info)
        else:
            yield event.plain_result("❌ 文件生成失败")

    @filter.command("fileconfig", alias={"文件配置"})
    async def config_command(
        self,
        event: AstrMessageEvent,
        action: str = "show",
        key: str = "",
        value: str = "",
    ):
        """配置插件行为"""
        if action == "show":
            config_text = "⚙️ 当前配置：\n\n"
            config_text += f"私聊自动检测: {'✅' if self.config_mgr.get('auto_detect_in_private') else '❌'}\n"
            config_text += f"群聊自动检测: {'✅' if self.config_mgr.get('auto_detect_in_group') else '❌'}\n"
            config_text += f"群聊需要@: {'✅' if self.config_mgr.get('require_at_in_group') else '❌'}\n"
            config_text += (
                f"最小消息长度: {self.config_mgr.get('min_message_length')}\n"
            )
            config_text += f"启用Office文件: {'✅' if self.config_mgr.get('enable_office_files') else '❌'}\n"
            config_text += f"需要管理员权限: {'✅' if self.config_mgr.get('require_admin') else '❌'}\n"

            whitelist = self.config_mgr.get("whitelist_users", [])
            if whitelist:
                config_text += f"白名单用户: {', '.join(whitelist)}\n"
            else:
                config_text += f"白名单用户: (未设置，所有人可用)\n"

            config_text += "\nOffice支持状态：\n"
            config_text += (
                f"Word: {'✅' if self.office_gen.support['word'] else '❌'}\n"
            )
            config_text += (
                f"Excel: {'✅' if self.office_gen.support['excel'] else '❌'}\n"
            )
            config_text += f"PowerPoint: {'✅' if self.office_gen.support['powerpoint'] else '❌'}\n"

            yield event.plain_result(config_text)

        elif action == "set" and key:
            result = await self.config_mgr.set(key, value)
            if result:
                yield event.plain_result(
                    f"✅ 配置已更新: {key} = {self.config_mgr.get(key)}"
                )
            else:
                yield event.plain_result(f"❌ 无效的配置项或值")
        else:
            yield event.plain_result(
                "⚙️ 配置管理\n\n"
                "查看配置: /fileconfig show\n"
                "修改配置: /fileconfig set <配置项> <值>\n\n"
                "可用配置项:\n"
                "- auto_detect_in_private (true/false)\n"
                "- auto_detect_in_group (true/false)\n"
                "- require_at_in_group (true/false)\n"
                "- min_message_length (数字)\n"
                "- enable_office_files (true/false)\n"
                "- require_admin (true/false) 🔒\n"
                "- whitelist_users (逗号分隔的用户ID) 🔒\n\n"
                "示例:\n"
                "/fileconfig set require_admin true\n"
                "/fileconfig set whitelist_users 123456,789012\n"
                "/fileconfig set min_message_length 20"
            )

    @filter.command("listfiles", alias={"文件列表", "lf"})
    async def list_files(self, event: AstrMessageEvent):
        """列出已生成的文件"""
        try:
            files = list(self.plugin_data_path.glob("*"))

            if not files:
                yield event.plain_result("📂 暂无已生成的文件")
                return

            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)

            file_list = ["📂 已生成的文件列表：\n"]
            for i, file in enumerate(files[:20], 1):
                size = file.stat().st_size
                size_str = format_file_size(size)
                mtime = datetime.fromtimestamp(file.stat().st_mtime).strftime(
                    "%Y-%m-%d %H:%M"
                )
                file_list.append(f"{i}. {file.name} ({size_str}) - {mtime}")

            if len(files) > 20:
                file_list.append(f"\n... 还有 {len(files) - 20} 个文件未显示")

            yield event.plain_result("\n".join(file_list))
        except Exception as e:
            logger.error(f"[文件生成器] 列出文件失败: {e}")
            yield event.plain_result(f"❌ 列出文件失败: {str(e)}")

    @filter.command("clearfiles", alias={"清空文件"})
    async def clear_files(self, event: AstrMessageEvent):
        """清空所有已生成的文件"""
        try:
            files = list(self.plugin_data_path.glob("*"))
            count = len(files)

            for file in files:
                if file.is_file():
                    file.unlink()

            yield event.plain_result(f"✅ 已清空 {count} 个文件")
        except Exception as e:
            logger.error(f"[文件生成器] 清空文件失败: {e}")
            yield event.plain_result(f"❌ 清空文件失败: {str(e)}")

    async def terminate(self):
        """插件卸载时调用"""
        logger.info("[文件生成器] 插件已卸载")
