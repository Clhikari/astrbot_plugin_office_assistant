def format_file_size(size: int | float) -> str:
    """格式化文件大小"""
    if size < 0:
        return "0 B"
    if size == 0:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} TB"
