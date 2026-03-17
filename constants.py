"""
静态常量
"""

from enum import Enum, auto


class OfficeType(Enum):
    WORD = auto()
    EXCEL = auto()
    POWERPOINT = auto()


OFFICE_EXTENSIONS = {
    OfficeType.WORD: ".docx",
    OfficeType.EXCEL: ".xlsx",
    OfficeType.POWERPOINT: ".pptx",
}

OFFICE_LIBS = {
    OfficeType.WORD: ("docx", "python-docx"),
    OfficeType.EXCEL: ("openpyxl", "openpyxl"),
    OfficeType.POWERPOINT: ("pptx", "python-pptx"),
}

# 字符串映射
OFFICE_TYPE_MAP = {
    "word": OfficeType.WORD,
    "excel": OfficeType.EXCEL,
    "powerpoint": OfficeType.POWERPOINT,
}

TEXT_SUFFIXES = frozenset(
    {
        ".txt",
        ".md",
        ".log",
        ".rst",
        ".py",
        ".js",
        ".ts",
        ".jsx",
        ".tsx",
        ".c",
        ".cpp",
        ".h",
        ".java",
        ".go",
        ".rs",
        ".json",
        ".yaml",
        ".yml",
        ".toml",
        ".xml",
        ".csv",
        ".html",
        ".css",
        ".sql",
        ".sh",
        ".bat",
    }
)

# 默认值
OFFICE_SUFFIXES = frozenset(OFFICE_EXTENSIONS.values())
SUFFIX_TO_OFFICE_TYPE = {v: k for k, v in OFFICE_EXTENSIONS.items()}

# 旧版 Office 格式也映射到对应类型（用于读取）
LEGACY_OFFICE_SUFFIXES = {
    ".doc": OfficeType.WORD,
    ".xls": OfficeType.EXCEL,
    ".ppt": OfficeType.POWERPOINT,
}
# 合并新旧格式映射
SUFFIX_TO_OFFICE_TYPE.update(LEGACY_OFFICE_SUFFIXES)

# 所有可读取的 Office 格式（新旧）
ALL_OFFICE_SUFFIXES = frozenset(SUFFIX_TO_OFFICE_TYPE.keys())

DEFAULT_MAX_FILE_SIZE_MB = 20
DEFAULT_CHUNK_SIZE = 64 * 1024  # 64 KB
# 所有文件类工具名称，用于 before_llm_chat 中的可见性控制。
# 同步点：@llm_tool(name=...) 定义在 main.py，
#         document tool 名称定义在 agent_tools/document_tools.py 的 name 字段。
FILE_TOOLS = [
    "read_file",
    "create_office_file",
    "create_document",
    "add_blocks",
    "finalize_document",
    "export_document",
    "convert_to_pdf",
    "convert_from_pdf",
]

EXECUTION_TOOLS = [
    "astrbot_execute_shell",
    "astrbot_execute_python",
    "astrbot_execute_ipython",
]

NOTICE_TOOLS_DENIED = (
    "\n[System Notice] File/Office/PDF actions are unavailable in this chat."
    " Say so and suggest private chat or admin enablement."
    " Do not call file tools or use `astrbot_execute_python`, `astrbot_execute_shell`,"
    " or `astrbot_execute_ipython` to bypass this restriction."
)

NOTICE_DOCUMENT_TOOLS_GUIDE = (
    "\n[System Notice] For Word documents, use the stateful document tools:"
    " `create_document(theme_name=..., table_template=..., density=..., accent_color=...)`"
    " -> `add_blocks(blocks=[...])`"
    " -> `finalize_document` -> `export_document`."
    " Prefer one `add_blocks` call per section or logical chunk,"
    " and only call it again when appending more content."
    " Prefer theme presets such as `business_report`, `project_review`, or `executive_brief`."
    " Prefer table presets such as `report_grid`, `metrics_compact`, or `minimal`."
    " Use `density=compact` for tighter layouts and `accent_color=RRGGBB` for brand accents."
    " Use block-level `style={align, emphasis, font_scale, table_grid, cell_align}`"
    " and `layout={spacing_before, spacing_after}` tokens when presets are not enough."
    " Once any document tool has been used, do not stop with a normal assistant reply"
    " while the document is still draft or finalized."
    " Continue calling document tools until `export_document` succeeds."
    " `export_document` sends the file."
    " If the user request depends on uploaded readable files,"
    " call `read_file` before `create_document` or `create_office_file`."
    " Use `create_office_file` only for simple one-shot output (Excel/PPT)."
)

NOTICE_UPLOADED_FILE_TEMPLATE = (
    "\n[System Notice] Received an uploaded {type_desc}: {original_name} (suffix: {file_suffix})."
    " Stored in workspace."
    " If the user request depends on this uploaded file,"
    " call `read_file` before `create_document` or `create_office_file`."
    " Do not create a new document before reading the uploaded source at least once."
    " Ask the user what they want before calling tools only when the task is still unclear."
    " For complex Word output, prefer the stateful document tools over `create_office_file`."
)

MSG_DOCUMENT_EXPORTED = "✅ 文档已导出"

# PDF 相关常量
PDF_SUFFIX = ".pdf"
PDF_LIBS = {
    "pdf2docx": ("pdf2docx", "pdf2docx"),
    "tabula": ("tabula", "tabula-py"),
    "pdfplumber": ("pdfplumber", "pdfplumber"),
}

# 支持转换为 PDF 的格式（包括旧版 Office 格式）
CONVERTIBLE_TO_PDF = frozenset(
    {
        ".docx",
        ".doc",  # Word
        ".xlsx",
        ".xls",  # Excel
        ".pptx",
        ".ppt",  # PowerPoint
    }
)

# 支持从 PDF 转换的目标格式
PDF_TARGET_FORMATS = {
    "word": (".docx", "Word 文档"),
    "excel": (".xlsx", "Excel 表格"),
}
