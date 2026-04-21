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
SUFFIX_TO_OFFICE_TYPE = {v: k for k, v in OFFICE_EXTENSIONS.items()} | {
    ".doc": OfficeType.WORD,
    ".xls": OfficeType.EXCEL,
    ".ppt": OfficeType.POWERPOINT,
}

EXCEL_SUFFIXES = frozenset(
    suffix
    for suffix, office_type in SUFFIX_TO_OFFICE_TYPE.items()
    if office_type is OfficeType.EXCEL
)

# 所有可读取的 Office 格式（新旧）
ALL_OFFICE_SUFFIXES = frozenset(SUFFIX_TO_OFFICE_TYPE.keys())

DEFAULT_MAX_FILE_SIZE_MB = 20
DEFAULT_CHUNK_SIZE = 64 * 1024  # 64 KB
DEFAULT_MAX_INLINE_DOCX_IMAGE_MB = 2
DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT = 3
DOCUMENT_BLOCK_FONT_SCALE_MIN = 0.75
DOCUMENT_BLOCK_FONT_SCALE_MAX = 2.0
DOCUMENT_BLOCK_SPACING_MIN = 0.0
DOCUMENT_BLOCK_SPACING_MAX = 72.0
# 所有文件类工具名称，用于 before_llm_chat 中的可见性控制。
# 同步点：@llm_tool(name=...) 定义在 main.py，
#         document tool 名称定义在 agent_tools/document_tools.py 的 name 字段。
FILE_TOOLS = (
    "read_file",
    "read_workbook",
    "create_office_file",
    "create_document",
    "add_blocks",
    "finalize_document",
    "export_document",
    "create_workbook",
    "write_rows",
    "export_workbook",
    "execute_excel_script",
    "convert_to_pdf",
    "convert_from_pdf",
)

EXPLICIT_FILE_TOOL_EVENT_KEY = "office_assistant_explicit_file_tool_name"
DOC_COMMAND_TRIGGER_EVENT_KEY = "office_assistant_doc_command_trigger"

EXECUTION_TOOLS = (
    "astrbot_execute_shell",
    "astrbot_execute_python",
    "astrbot_execute_ipython",
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
