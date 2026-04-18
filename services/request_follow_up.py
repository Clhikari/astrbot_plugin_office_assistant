from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class IdentifierDescriptor:
    token_name: str
    group_name: str
    follow_up_re: re.Pattern[str]
    explicit_capture_re: re.Pattern[str]
    bare_capture_re: re.Pattern[str]
    bare_validator_name: str


@dataclass(frozen=True, slots=True)
class FollowUpNoticeStrategy:
    identifier: IdentifierDescriptor
    lookup_attr_name: str
    section_builder_name: str
    missing_section_builder_name: str
    payload_builder_name: str
    log_label: str


def build_identifier_token_pattern(token_name: str) -> str:
    return rf"(?<![A-Za-z0-9_]){re.escape(token_name)}(?![A-Za-z0-9_])"


def compile_identifier_token_regex(token_name: str) -> re.Pattern[str]:
    return re.compile(build_identifier_token_pattern(token_name), flags=re.IGNORECASE)


def compile_identifier_explicit_capture_regex(
    token_name: str,
    group_name: str,
) -> re.Pattern[str]:
    return re.compile(
        build_identifier_token_pattern(token_name)
        + rf"(?:\s*[:=：]\s*|\s*(?:为|是)\s*|\s+is\s+)[`\"']?(?P<{group_name}>[A-Za-z0-9_-]+)[`\"']?",
        flags=re.IGNORECASE,
    )


def compile_identifier_bare_capture_regex(
    token_name: str,
    group_name: str,
) -> re.Pattern[str]:
    return re.compile(
        build_identifier_token_pattern(token_name)
        + rf"\s+[`\"']?(?P<{group_name}>[A-Za-z0-9_-]*[\d_-][A-Za-z0-9_-]*)[`\"']?",
        flags=re.IGNORECASE,
    )


def build_identifier_descriptor(
    *,
    token_name: str,
    bare_validator_name: str,
) -> IdentifierDescriptor:
    return IdentifierDescriptor(
        token_name=token_name,
        group_name=token_name,
        follow_up_re=compile_identifier_token_regex(token_name),
        explicit_capture_re=compile_identifier_explicit_capture_regex(
            token_name,
            token_name,
        ),
        bare_capture_re=compile_identifier_bare_capture_regex(
            token_name,
            token_name,
        ),
        bare_validator_name=bare_validator_name,
    )


def extract_identifier_from_text(
    *,
    request_text: str,
    explicit_capture_re: re.Pattern[str],
    bare_capture_re: re.Pattern[str],
    group_name: str,
    is_valid_bare_id: Callable[[str], bool],
) -> str:
    if not request_text:
        return ""

    explicit_match = explicit_capture_re.search(request_text)
    if explicit_match:
        return str(explicit_match.group(group_name) or "").strip()

    bare_match = bare_capture_re.search(request_text)
    if not bare_match:
        return ""

    candidate = str(bare_match.group(group_name) or "").strip()
    if is_valid_bare_id(candidate):
        return candidate
    return ""


__all__ = [
    "FollowUpNoticeStrategy",
    "IdentifierDescriptor",
    "build_identifier_descriptor",
    "build_identifier_token_pattern",
    "extract_identifier_from_text",
]
