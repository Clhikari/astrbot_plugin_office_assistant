from collections.abc import Callable

from astrbot.api import logger
from astrbot.core.message.components import At, Reply
from astrbot.core.platform.message_type import MessageType


class AccessPolicyService:
    def __init__(
        self,
        *,
        whitelist_users: list[str] | None = None,
        admin_users: list[str] | None = None,
        get_admin_users: Callable[[], set[str] | list[str] | None] | None = None,
        enable_features_in_group: bool,
    ) -> None:
        self._whitelist_users = {str(user_id) for user_id in whitelist_users or []}
        self._admin_users = {str(user_id) for user_id in admin_users or []}
        self._get_admin_users = get_admin_users
        self._enable_features_in_group = bool(enable_features_in_group)

    def check_permission(self, event) -> bool:
        logger.debug("正在检查用户权限")
        if event.is_admin():
            return True
        user_id = str(event.get_sender_id())
        admin_users = self._admin_users
        if self._get_admin_users is not None:
            try:
                dynamic_admin_users = self._get_admin_users() or set()
                admin_users = {
                    str(admin_id) for admin_id in dynamic_admin_users
                }
            except Exception as exc:
                logger.warning(
                    f"读取框架管理员配置失败: {exc}",
                    exc_info=True,
                )
        if user_id in admin_users:
            return True
        if not self._whitelist_users:
            return False
        return user_id in self._whitelist_users

    def is_group_message(self, event) -> bool:
        return event.message_obj.type == MessageType.GROUP_MESSAGE

    def is_group_feature_enabled(self, event) -> bool:
        if not self.is_group_message(event):
            return True
        return self._enable_features_in_group

    def group_feature_disabled_error(self) -> str:
        return (
            "错误：群聊中已禁用本插件功能，请私聊使用，或在配置中开启“群聊启用插件功能”"
        )

    def is_bot_mentioned(self, event) -> bool:
        try:
            platform_level_mention = getattr(event, "is_mentioned", None)
            if callable(platform_level_mention) and platform_level_mention():
                return True
            bot_id = str(event.message_obj.self_id)
            for segment in event.message_obj.message:
                if isinstance(segment, (At, Reply)):
                    target_id = getattr(segment, "qq", None) or getattr(
                        segment, "target", None
                    )
                    if target_id and str(target_id) == bot_id:
                        return True
            return False
        except Exception as exc:
            logger.error(f"未知错误{exc}")
            return False
