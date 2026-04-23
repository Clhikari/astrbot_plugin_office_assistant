def build_excel_routing_notice() -> str:
    return (
        "\n[System Notice] Excel 路由规则\n"
        "- 新建 + 无图表无公式 -> `create_workbook` / `write_rows` / `export_workbook`\n"
        "- 新建 + 有图表或公式 -> `execute_excel_script`\n"
        "- 读取已有 `.xlsx/.xls` -> `read_workbook`\n"
        "- 修改已有 `.xlsx` -> `execute_excel_script`\n"
        "- 修改已有 `.xls` 不走 `execute_excel_script`，先用 `read_workbook`\n"
        "- 读取完成后，如目标是基于读取结果继续生成新文件，可以继续调用新建路径\n"
    )


def build_excel_read_notice() -> str:
    return (
        "\n[System Notice] Excel 读取工具指南\n"
        "- 读取已有 `.xlsx/.xls` 时，优先使用 `read_workbook`\n"
        "- `read_workbook` 会返回文件名、Sheet 列表和每个 Sheet 的文本内容\n"
        "- 读取场景不要调用 `create_workbook`、`write_rows`、`export_workbook`\n"
        "- 如果读取后还要继续生成新的 Excel，可以在读取完成后再进入新建路径\n"
    )


def build_excel_script_notice() -> str:
    return (
        "\n[System Notice] Excel 脚本工具指南\n"
        "- 复杂新建或修改已有 `.xlsx` 时，使用 `execute_excel_script`\n"
        "- 现有 `.xls` 不支持 `execute_excel_script` 修改\n"
        "- 优先使用 `sandbox` runtime；`local` runtime 会在宿主机直接执行脚本，仅适合受信任环境\n"
        "- 脚本运行环境提供 `openpyxl`、`Workbook`、`load_workbook`、`Path`\n"
        "- 已校验的输入工作簿通过 `input_files` 提供；需要导出文件时写到 `output_path`\n"
        "- 成功时只能二选一：设置 `result_text` 返回文本，或写出 `output_path` 返回文件\n"
        "- 失败时会返回错误信息、完整 traceback 和本次脚本原文\n"
        "- 脚本最多重试 2 次；超过后停止重试，直接说明失败原因，并建议缩小需求或改走原语路径\n"
    )
