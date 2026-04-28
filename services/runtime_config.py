import inspect

from astrbot.api.event import AstrMessageEvent

SUPPORTED_COMPUTER_RUNTIME_MODES = frozenset({"local", "sandbox", "none"})
_SESSION_CONFIG_KEYWORD_NAMES = (
    "umo",
    "session_id",
    "unified_msg_origin",
)


def _is_call_shape_type_error(exc: TypeError) -> bool:
    traceback_obj = exc.__traceback__
    return traceback_obj is not None and traceback_obj.tb_next is None


def get_session_config(get_config, session_id: str):
    try:
        signature = inspect.signature(get_config)
    except (TypeError, ValueError):
        try:
            return get_config(session_id)
        except TypeError as exc:
            if not _is_call_shape_type_error(exc):
                raise
            try:
                return get_config(umo=session_id)
            except TypeError as exc:
                if not _is_call_shape_type_error(exc):
                    raise
                return get_config()

    parameters = tuple(signature.parameters.values())
    if any(
        parameter.kind
        in (
            inspect.Parameter.POSITIONAL_ONLY,
            inspect.Parameter.POSITIONAL_OR_KEYWORD,
            inspect.Parameter.VAR_POSITIONAL,
        )
        for parameter in parameters
    ):
        return get_config(session_id)

    if any(
        parameter.kind is inspect.Parameter.VAR_KEYWORD for parameter in parameters
    ):
        return get_config(umo=session_id)

    for parameter_name in _SESSION_CONFIG_KEYWORD_NAMES:
        if parameter_name in signature.parameters:
            return get_config(**{parameter_name: session_id})

    return get_config()


def resolve_computer_runtime_mode(
    astrbot_context,
    event: AstrMessageEvent,
    *,
    default: str = "local",
) -> str:
    if astrbot_context is None:
        return default
    config = None
    get_config = getattr(astrbot_context, "get_config", None)
    if callable(get_config):
        session_id = str(getattr(event, "unified_msg_origin", "") or "")
        config = get_session_config(get_config, session_id)
    if not isinstance(config, dict):
        legacy_config = getattr(astrbot_context, "astrbot_config", None)
        config = legacy_config if isinstance(legacy_config, dict) else None
    if not isinstance(config, dict):
        return default
    provider_settings = config.get("provider_settings", {})
    if not isinstance(provider_settings, dict):
        return default
    if "computer_use_runtime" not in provider_settings:
        return default
    runtime = provider_settings.get("computer_use_runtime")
    if isinstance(runtime, str) and runtime.strip():
        return runtime.strip().lower()
    if runtime is None:
        return "null"
    if isinstance(runtime, bool):
        return str(runtime).lower()
    if isinstance(runtime, str):
        return "<empty>"
    return str(runtime)
