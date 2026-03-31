"""平台兼容性检测

将 _IS_WINDOWS、_WIN32COM_AVAILABLE、_ANTIWORD_AVAILABLE 等跨模块共享的
平台检测逻辑集中于此，避免 utils.py 和 pdf_converter.py 各自重复定义。
"""

import platform
import shutil

_IS_WINDOWS: bool = platform.system() == "Windows"
_ANTIWORD_AVAILABLE: bool = shutil.which("antiword") is not None

_WIN32COM_AVAILABLE: bool = False

if _IS_WINDOWS:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401

        _WIN32COM_AVAILABLE = True
    except ImportError:
        pass
