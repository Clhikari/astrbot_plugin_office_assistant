def build_excel_routing_notice() -> str:
    return (
        "\n[System Notice] Excel 路径选择规则\n"
        "- 你需要根据用户目标、上传文件、可用工具和本轮动态上下文选择路径；不要只按关键词机械选择\n"
        "- 只查看、解释、提取、确认已有工作簿内容 -> `read_workbook`\n"
        "- 新建普通录入表、简单汇总表、简单筛选结果，且不需要公式/图表/条件格式/数据验证 -> `create_workbook` / `write_rows` / `export_workbook`\n"
        "- 新建或修改时涉及公式、图表、条件格式、数据验证、多 Sheet 公式联动、旧表清洗后导出 -> `execute_excel_script`\n"
        "- 用户上传了 Excel 且目标不清楚时，先读取内容；如果读取后能确定要生成新文件，应继续调用合适的生成路径\n"
        "- 上传 Excel 后做简单汇总、筛选、排序、拆分结果表，且不需要公式/图表/条件格式/数据验证时，先 `read_workbook`；读取内容足够时用原语生成结果，不要直接用脚本\n"
        "- `.xls` 是旧版格式；修改前先读取，只有用户明确要求导出 `.xlsx` 且说明改动时才生成新文件\n"
        "- 工具不可用时不要假装完成；改用当前可用路径，或说明当前限制\n"
    )


def build_excel_read_notice() -> str:
    return (
        "\n[System Notice] Excel 读取工具指南\n"
        "- 读取已有 `.xlsx/.xls` 时，优先使用 `read_workbook`\n"
        "- `read_workbook` 会返回文件名、Sheet 列表和每个 Sheet 的文本内容\n"
        "- 如果用户最终目标是生成、修改或导出新文件，读取结果只是准备步骤，不是最终交付\n"
        "- `.xls` 是旧版格式；用户只说更新、保存但未说明具体改动时，读取后说明需要转换为 `.xlsx`，并询问要改哪些数据\n"
        "- 读取场景不要调用 `create_workbook`、`write_rows`、`export_workbook`\n"
        "- 如果读取后还要继续生成新的 Excel，可以在读取完成后再进入新建路径\n"
    )


def build_excel_script_notice() -> str:
    return (
        "\n[System Notice] Excel 脚本工具指南\n"
        "- 需要公式、图表、条件格式、数据验证、多 Sheet 公式联动、清洗后导出或修改已有 Excel 时，使用 `execute_excel_script`\n"
        "- 脚本运行环境提供 `openpyxl`、`Workbook`、`load_workbook`、`Path`\n"
        "- 已校验的输入工作簿通过脚本变量 `input_files` 提供；其中每个路径都相对于脚本执行目录，按顺序使用 `input_files[0]`、`input_files[1]`，不要硬写临时文件名\n"
        "- 需要导出文件时，必须保存到脚本变量 `output_path`，不要自己另取输出文件名\n"
        "- 不要把 `.xls` 当成可原格式保存的工作簿；只有用户明确要求转换/导出为 `.xlsx` 且说明具体改动时，才用脚本生成 `.xlsx`\n"
        "- 用户只说更新、处理、保存，但没有说明要改哪些单元格、列、公式或备注时，不要自行编造改动规则；先询问用户\n"
        "- 当你调用 `execute_excel_script` 时，assistant content 必须为空；不要在工具执行前输出说明、脚本正文或完成结论\n"
        "- 脚本正文只能放在 tool arguments 的 `script` 字段，不要以 Markdown 代码块发给用户\n"
        "- 图表报告的 Dashboard 不能只是空白图表容器；需要写入源数据、关键指标区或明确可核对的数据表\n"
        "- 写公式时必须处理空值、0 值和无法匹配目标的数据；例如分母可能为空时用 `IF`/`IFERROR`，状态列应显示 `Needs Target` 等明确结果\n"
        "- 数据清洗脚本里写正则时优先用 raw string，并在脚本中对关键样本做最小校验；不要在 raw string 中多写一层反斜杠导致匹配字面量反斜杠\n"
        "- 成功时只能二选一：设置 `result_text` 返回文本，或写出 `output_path` 返回文件\n"
        "- 失败时会返回错误信息、完整 traceback 和本次脚本原文\n"
        "- 工具失败后，直接根据错误修正脚本或说明失败原因；不要先发送角色表演或与错误无关的解释\n"
        "- 如果工具结果包含 `requires_review: true` 或 `quality_warnings`，必须逐条列出全部警告，不要遗漏，不要表述为完全完成\n"
        "- 脚本最多重试 3 次；超过后停止重试，直接说明失败原因，并建议缩小需求或改走原语路径\n"
    )


def build_excel_script_unavailable_notice() -> str:
    return (
        "\n[System Notice] Excel 脚本工具当前不可用\n"
        "- 当前未暴露 Excel 脚本执行工具，不能新增公式、图表、样式，也不能导出脚本生成的新版本\n"
        "- 如果用户上传了 Excel 或指定了已有 Excel，先调用 `read_workbook` 读取内容\n"
        "- 读取后直接说明当前环境未启用 Excel 脚本运行能力，无法完成新增公式或导出新版本\n"
        "- 不要承诺稍后可以导出文件；不要只要求用户补充公式位置或公式内容\n"
        "- 可以建议管理员启用 sandbox runtime，或让用户改成只读取、总结表格内容\n"
    )
