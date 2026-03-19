from astrbot.api import logger
from astrbot.core.message.message_event_result import MessageChain


class ErrorHookService:
    def __init__(self, *, context, config, plugin_name: str) -> None:
        self._context = context
        self._config = config
        self._plugin_name = plugin_name

    async def handle_plugin_error(
        self,
        event,
        plugin_name: str,
        handler_name: str,
        error: Exception,
        traceback_text: str,
    ) -> None:
        if plugin_name != self._plugin_name:
            return

        debug_settings = self._config.get("debug_settings", {})
        target_session = debug_settings.get(
            "debug_error_hook_target_session",
            self._config.get("debug_error_hook_target_session"),
        )
        if target_session is None:
            target_session = ""
        else:
            target_session = str(target_session).strip()
        if not target_session:
            target_session = event.unified_msg_origin

        trace_lines = traceback_text.splitlines() if traceback_text else []
        trace_tail = "\n".join(trace_lines[-3:]) if trace_lines else ""
        if len(trace_tail) > 800:
            trace_tail = trace_tail[-800:]

        sent = await self._context.send_message(
            target_session,
            MessageChain().message(
                "[plugin-error-hook]\n"
                f"plugin={plugin_name}\n"
                f"handler={handler_name}\n"
                f"error={error}\n"
                f"trace_tail={trace_tail}",
            ),
        )
        if not sent:
            logger.warning(
                f"[plugin-error-hook] target session not found: {target_session}"
            )

        event.stop_event()
