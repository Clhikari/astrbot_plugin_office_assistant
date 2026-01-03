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
FILE_TOOLS = [
    "read_file",
    "create_office_file",
    "convert_to_pdf",
    "convert_from_pdf",
]

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
