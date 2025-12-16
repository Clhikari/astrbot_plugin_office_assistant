import asyncio
from typing import Any
from astrbot.api import logger


class ConfigManager:
    """配置管理器"""

    VALID_KEYS = {
        "auto_detect_in_private": bool,
        "auto_detect_in_group": bool,
        "require_at_in_group": bool,
        "min_message_length": int,
        "enable_office_files": bool,
    }

    DEFAULT_CONFIG = {
        "auto_detect_in_private": True,
        "auto_detect_in_group": False,
        "require_at_in_group": True,
        "min_message_length": 15,
        "enable_office_files": True,
    }

    def __init__(self, plugin_instance):
        self.plugin = plugin_instance
        self.config = self._load_config()

    def _load_config(self) -> dict[str, Any]:
        """加载配置"""
        if hasattr(self.plugin, "config") and self.plugin.config:
            return self.plugin.config

        try:
            config_data = asyncio.run(
                self.plugin.get_kv_data("config", self.DEFAULT_CONFIG)
            )
            return config_data if config_data else self.DEFAULT_CONFIG
        except Exception as e:
            logger.warning(f"[文件生成器] 加载配置失败: {e}, 使用默认配置")
            return self.DEFAULT_CONFIG.copy()

    def get(self, key: str, default=None) -> Any:
        """获取配置项"""
        return self.config.get(key, default)

    async def set(self, key: str, value: str) -> bool:
        """设置配置项"""
        if key not in self.VALID_KEYS:
            logger.warning(f"[文件生成器] 无效的配置项: {key}")
            return False

        try:
            # 类型转换
            expected_type = self.VALID_KEYS[key]
            if expected_type == bool:
                new_value = value.lower() in ["true", "1", "yes", "on", "是", "开启"]
            elif expected_type == int:
                new_value = int(value)
            else:
                new_value = value

            self.config[key] = new_value
            await self.plugin.put_kv_data("config", self.config)

            logger.info(f"[文件生成器] 配置已更新: {key} = {new_value}")
            return True
        except Exception as e:
            logger.error(f"[文件生成器] 设置配置失败: {e}")
            return False

    def get_all(self) -> dict[str, Any]:
        """获取所有配置"""
        return self.config.copy()
