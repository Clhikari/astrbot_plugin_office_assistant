"""Bind draft ownership to trusted AstrBot context, never model arguments."""

from typing import Protocol

OwnerKey = tuple[str, str, str]


class OwnedDraft(Protocol):
    _owner_key: OwnerKey | None


def owner_from_context(context: object | None) -> OwnerKey | None:
    # Context-free callers are supported for standalone/local tool integrations.
    # They can only access context-free drafts, not AstrBot-owned drafts.
    if context is None:
        return None
    event = getattr(getattr(context, "context", None), "event", None)
    try:
        values = (
            event.get_platform_id(),
            event.get_sender_id(),
            event.unified_msg_origin,
        )
    except (AttributeError, TypeError) as exc:
        raise ValueError("权限不足：无法确认当前用户会话") from exc
    if any(type(value) not in (str, int) or not str(value).strip() for value in values):
        raise ValueError("权限不足：无法确认当前用户会话")
    return tuple(str(value) for value in values)


def require_draft_owner(draft: OwnedDraft, context: object | None) -> None:
    if draft._owner_key != owner_from_context(context):
        raise ValueError("权限不足：该文档或工作簿不属于当前用户会话")
