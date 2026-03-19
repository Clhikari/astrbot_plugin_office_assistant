from astrbot.api import logger
from astrbot.core.message.components import At, Reply
from astrbot.core.platform.message_type import MessageType


class AccessPolicyService:
    def __init__(
        self,
        *,
        whitelist_users: list[str] | None = None,
        enable_features_in_group: bool,
    ) -> None:
        self._whitelist_users = [str(user_id) for user_id in whitelist_users or []]
        self._enable_features_in_group = bool(enable_features_in_group)

    def check_permission(self, event) -> bool:
        logger.debug("正在检查用户权限")
        if event.is_admin():
            return True
        if not self._whitelist_users:
            return False
        user_id = str(event.get_sender_id())
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
            bot_id = str(event.message_obj.self_id)
            for segment in event.message_obj.message:
                if isinstance(segment, At) or isinstance(segment, Reply):
                    target_id = getattr(segment, "qq", None) or getattr(
                        segment, "target", None
                    )
                    if target_id and str(target_id) == bot_id:
                        return True
            return False
        except Exception as exc:
            logger.error(f"未知错误{exc}")
            return False
