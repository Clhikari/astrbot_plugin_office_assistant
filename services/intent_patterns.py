import re
from collections.abc import Iterable


NEGATIVE_TOOL_PREFIX_RE = re.compile(
    r"(?:不要|别|勿|不用|无需|do\s+not|don't|not)\s*(?:调用|使用|call|use|invoke)?\s*$",
    flags=re.IGNORECASE,
)

DOCUMENT_ID_RE = re.compile(
    r'document_id["`\']?\s*[:=]\s*["`\']?(?P<document_id>[A-Za-z0-9_-]+)',
    flags=re.IGNORECASE,
)

DOCUMENT_INTENT_RE = re.compile(
    r"(create_document|add_blocks|finalize_document|export_document|"
    r"document_id\b|整理成\s*(?:文档|报告|汇报|word|docx)|"
    r"导出成\s*word|导出为\s*word|生成\s*(?:文档|报告|汇报|word|docx)|"
    r"生成\s*(?:excel|xlsx|表格|powerpoint|ppt|幻灯片)|"
    r"\bword\b|\bdocx\b|\bexcel\b|\bxlsx\b|\bpowerpoint\b|\bppt\b|"
    r"文档|报告|汇报|表格|幻灯片)",
    flags=re.IGNORECASE,
)

FILE_INTENT_RE = re.compile(
    r"(read_file|convert_to_pdf|convert_from_pdf|"
    r"读取.*文件|查看.*文件|看看.*文件|读取内容|读取这个|"
    r"\bpdf\b|转成\s*pdf|转换成\s*pdf|导出成\s*pdf|导出为\s*pdf|"
    r"pdf\s*转\s*word|pdf\s*转\s*excel|word\s*转\s*pdf|excel\s*转\s*pdf)",
    flags=re.IGNORECASE,
)

PDF_CONVERSION_INTENT_RE = re.compile(
    r"(convert_to_pdf|convert_from_pdf|"
    r"pdf\s*转\s*word|pdf\s*转\s*excel|word\s*转\s*pdf|excel\s*转\s*pdf|"
    r"转成\s*pdf|转换成\s*pdf|导出成\s*pdf|导出为\s*pdf|"
    r"pdf.*(?:转\s*word|转\s*excel|导出|转换)|"
    r"(?:转\s*word|转\s*excel).*pdf)",
    flags=re.IGNORECASE,
)

DOCUMENT_FOLLOWUP_RE = re.compile(
    r"(继续|接着|再加|加一章|加一节|补充|完善|导出|发给我)",
    flags=re.IGNORECASE,
)

DOCUMENT_FOLLOWUP_ACTION_RE = re.compile(
    r"(导出|发给我|加一章|加一节|补充|再加)",
    flags=re.IGNORECASE,
)

DOCUMENT_FOLLOWUP_SHORT_RE = re.compile(
    r"^(继续|接着|继续写|继续补充|继续完善|补充|完善|导出|发给我)$",
    flags=re.IGNORECASE,
)

DOCUMENT_FOLLOWUP_TOPICAL_RE = re.compile(
    r"(文档|报告|汇报|正文|内容|段落|章节|小节|标题|表格|草稿|这一章|这一节|上一段|下一段|上一版)",
    flags=re.IGNORECASE,
)

DOCUMENT_DETAIL_HINT_RE = re.compile(
    r"(create_document|正式汇报|正式报告|导出成\s*word|导出为\s*word|"
    r"\bword\b|\bdocx\b|汇报|报告|"
    r"生成\s*(?:word|docx|报告|汇报)|整理成\s*(?:word|docx|报告|汇报)|"
    r"business_report|project_review|executive_brief|accent_color|document_style)",
    flags=re.IGNORECASE,
)


def extract_document_id(text: str) -> str | None:
    if not text:
        return None
    match = DOCUMENT_ID_RE.search(text)
    if not match:
        return None
    return match.group("document_id")


def detect_explicit_file_tool(text: str, file_tools: Iterable[str]) -> str | None:
    if not text:
        return None

    explicit_matches: set[str] = set()
    tool_invocation_prefix = (
        r"(?:调用|使用|请求(?:调用|使用)?|请(?!问)(?:调用|使用)?|"
        r"\b(?:call|use|invoke)\b|\bplease\s+(?:call|use|invoke)\b)"
    )

    for tool_name in sorted(file_tools, key=len, reverse=True):
        patterns = (
            rf"(?P<tool>{tool_invocation_prefix}\s*`?{re.escape(tool_name)}`?)",
            rf"(?P<tool>`{re.escape(tool_name)}`)",
            rf"(?P<tool>{re.escape(tool_name)}\s*\()",
            rf"(?P<tool>{re.escape(tool_name)}\s*[,，]\s*[a-zA-Z_]\w*\s*=)",
        )
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                tool_start = match.start("tool")
                prefix = text[max(0, tool_start - 20) : tool_start]
                if NEGATIVE_TOOL_PREFIX_RE.search(prefix):
                    continue
                explicit_matches.add(tool_name)
                break

    if len(explicit_matches) == 1:
        return next(iter(explicit_matches))
    return None


def has_document_intent(text: str) -> bool:
    return bool(extract_document_id(text) or DOCUMENT_INTENT_RE.search(text or ""))


def has_file_intent(text: str) -> bool:
    return bool(FILE_INTENT_RE.search(text or ""))


def has_pdf_conversion_intent(text: str) -> bool:
    return bool(PDF_CONVERSION_INTENT_RE.search(text or ""))


def looks_like_document_followup(text: str) -> bool:
    normalized_text = str(text or "").strip()
    if not normalized_text:
        return False
    if DOCUMENT_FOLLOWUP_SHORT_RE.fullmatch(normalized_text):
        return True
    if not DOCUMENT_FOLLOWUP_RE.search(normalized_text):
        return False
    if DOCUMENT_FOLLOWUP_TOPICAL_RE.search(normalized_text):
        return True
    return bool(DOCUMENT_FOLLOWUP_ACTION_RE.search(normalized_text))


def should_use_active_document_summary(text: str) -> bool:
    normalized_text = str(text or "").strip()
    if not normalized_text:
        return False
    if DOCUMENT_FOLLOWUP_SHORT_RE.fullmatch(normalized_text):
        return True
    return bool(
        DOCUMENT_FOLLOWUP_RE.search(normalized_text)
        and (
            DOCUMENT_FOLLOWUP_TOPICAL_RE.search(normalized_text)
            or DOCUMENT_FOLLOWUP_ACTION_RE.search(normalized_text)
        )
    )


def should_inject_document_tool_detail(
    *,
    request_text: str,
    document_id: str | None,
) -> bool:
    if not request_text:
        return False
    if document_id and not DOCUMENT_DETAIL_HINT_RE.search(request_text):
        return False
    return bool(DOCUMENT_DETAIL_HINT_RE.search(request_text))
