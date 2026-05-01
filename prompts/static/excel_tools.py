def build_excel_routing_notice() -> str:
    return (
        "\n[System Notice] Excel 路径选择规则\n"
        "- 你需要根据用户目标、上传文件、可用工具和本轮动态上下文选择路径；不要只按关键词机械选择\n"
        "- 只查看、解释、提取、确认已有工作簿内容 -> `read_workbook`\n"
        "- 新建普通录入表、简单汇总表、简单筛选结果，且不需要公式/图表/条件格式/数据验证 -> `create_workbook` / `write_rows` / `export_workbook`\n"
        "- 新建或修改时涉及公式、图表、条件格式、数据验证、多 Sheet 公式联动、旧表清洗后导出 -> `execute_excel_script`\n"
        "- 用户上传了 Excel 且目标不清楚时，先读取内容；如果读取后能确定要生成新文件，应继续调用合适的生成路径\n"
        "- 上传 Excel 后做简单汇总、筛选、排序、拆分结果表，且不需要公式/图表/条件格式/数据验证时，先 `read_workbook`；读取内容足够时用原语生成结果，不要直接用脚本\n"
        "- 如果生成结果需要引用、保留或修改源 Sheet，不属于原语路径，应走 `execute_excel_script`\n"
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
        "- 如果读取后还要继续生成新的 Excel，可以在读取完成后再进入新建路径\n"
    )


def build_excel_script_notice() -> str:
    return (
        "\n[System Notice] Excel 脚本工具指南\n"
        "- 需要公式、图表、条件格式、数据验证、多 Sheet 公式联动、清洗后导出或修改已有 Excel 时，使用 `execute_excel_script`\n"
        "- 脚本运行环境提供 `openpyxl`、`Workbook`、`load_workbook`、`Path`\n"
        "- 脚本还提供辅助函数：`load_input_workbook()` 保留输入工作簿样式读取，`save_output_workbook(workbook)` 保存到正确输出路径并自动美化新建表格 Sheet，`ensure_readable_sheet(sheet)` 为多行文本调整自动换行、列宽和行高，`format_table_sheet(sheet)` 美化普通表，`format_course_list_sheet(sheet)` 美化课表清单，`auto_format_workbook(workbook)` 可手动美化整本工作簿，`preserve_input_sheet_layout(workbook)` 可把输入 Sheet 的列宽、行高、冻结窗格、合并单元格和样式同步到同名输出 Sheet\n"
        "- 已校验的输入工作簿通过脚本变量 `input_files` 提供；其中每个路径都相对于脚本执行目录，按顺序使用 `input_files[0]`、`input_files[1]`，不要硬写临时文件名\n"
        "- 需要导出文件时，必须保存到脚本变量 `output_path`，不要自己另取输出文件名\n"
        "- 不要把 `.xls` 当成可原格式保存的工作簿；只有用户明确要求转换/导出为 `.xlsx` 且说明具体改动时，才用脚本生成 `.xlsx`\n"
        "- 用户只说更新、处理、保存，但没有说明要改哪些单元格、列、公式或备注时，不要自行编造改动规则；先询问用户\n"
        "- 当你调用 `execute_excel_script` 时，assistant content 必须为空；不要在工具执行前输出说明、脚本正文或完成结论\n"
        "- 脚本正文只能放在 tool arguments 的 `script` 字段，不要以 Markdown 代码块发给用户\n"
        "- Excel 公式里的文本常量必须用双引号，例如 `\"Keyboard\"`；单引号只用于 Sheet 名引用，不能写成 `'Keyboard'`\n"
        "- 当需求要求保留或导入原表、跨表引用、基于源表生成新表时，必须保留用户可见的原 Sheet 名；不要把 `Catalog` 改成 `Catalog_Ref` 这类内部名，如需辅助引用表，应额外创建辅助表\n"
        "- 基于输入工作簿新增明细表时，优先写 `workbook = load_input_workbook()`，新增 Sheet 后调用对应 `format_*_sheet`，最后 `save_output_workbook(workbook)`；不要用 `Workbook()` 重建原表\n"
        "- 新建 Raw、List、Summary、Issues 等表格 Sheet 时，至少用 `format_table_sheet(sheet)`；如果不确定每个 Sheet 的样式，保存前调用 `auto_format_workbook(workbook)`\n"
        "- 说明、README、空白、占位 Sheet 不应复制成 Raw 数据表；如果需要保留处理记录，写入 IgnoredSheets 或 Notes\n"
        "- Raw Sheet 只在需要保留原始读取结果时创建；如果清洗后的明细与 Raw 内容完全相同，优先只保留明细表、Summary 和 Issues\n"
        "- 如果已经用 `load_input_workbook()` 读取了原工作簿，不要再调用 `preserve_input_sheet_layout(workbook)`；原表样式已经在 workbook 里。`preserve_input_sheet_layout` 只用于你不得不新建 `Workbook()` 并复制原表值的情况\n"
        "- 包含课表、课程块、合并单元格、单元格内换行的 Sheet，保存前必须对该 Sheet 调用 `ensure_readable_sheet(sheet)`\n"
        "- 写公式时必须处理空值、0 值和无法匹配目标的数据；例如分母可能为空时用 `IF`/`IFERROR`，状态列应显示 `Needs Target` 等明确结果\n"
        "- 数据清洗脚本里写正则时优先用 raw string，并在脚本中对关键样本做最小校验；不要在 raw string 中多写一层反斜杠导致匹配字面量反斜杠\n"
        "- 成功时只能二选一：设置 `result_text` 返回文本，或写出 `output_path` 返回文件\n"
        "- 失败时会返回错误信息、完整 traceback 和本次脚本原文\n"
        "- 工具失败后，直接根据错误修正脚本或说明失败原因；不要先发送角色表演或与错误无关的解释\n"
        "- 如果工具结果包含 `requires_review: true` 或 `quality_warnings`，能根据警告修正时继续调用 `execute_excel_script`；无法修正时必须逐条列出全部警告，不要遗漏，不要表述为完全完成\n"
        "- 脚本最多重试 3 次；超过后停止重试，直接说明失败原因，并建议缩小需求或改走原语路径\n"
    )


def build_excel_domain_hints(scenario: str) -> str:
    scenario = (scenario or "").strip().lower()
    if not scenario:
        return ""

    lines = ["\n[System Notice] Excel 场景补充规则\n"]
    if scenario == "schedule":
        lines.extend(
            [
                "- 课表转明细时，备注、说明、空白占位不能作为数据行；必填字段缺失时应跳过或写入 Issues，不能伪装成课程记录\n",
                "- 课表源 Sheet 通常包含星期行、班级行、节次列和合并课程块；解析合并单元格时只处理左上角一次，并按覆盖到的班级列分别写入明细\n",
                "- 生成 CourseList 后调用 `format_course_list_sheet(course_list_sheet)`；保存前对原课表 Sheet 调用 `ensure_readable_sheet(source_sheet)`，同一个合并课程块覆盖到的行高要保持一致\n",
            ]
        )
    elif scenario == "dashboard":
        lines.append(
            "- Dashboard 必须写入源数据、关键指标区或明确可核对的数据表，不能只放图表对象\n"
        )
    elif scenario == "chart":
        lines.append(
            "- 用 openpyxl 生成饼图时，每个饼图应只有一个数据系列；横向区域要使用 `add_data(..., from_rows=True)`\n"
        )
    elif scenario == "pivot":
        lines.append(
            "- 未生成真实数据透视表而创建等价 `PivotSummary` 时，优先用 `SUMIFS` 等公式引用源 Sheet，避免用 Python 计算后写死静态值导致后续数据变化无法联动\n"
        )
    else:
        return ""
    return "".join(lines)


def build_excel_script_unavailable_notice() -> str:
    return (
        "\n[System Notice] Excel 脚本工具当前不可用\n"
        "- 当前未暴露 Excel 脚本执行工具，不能新增公式、图表、样式，也不能导出脚本生成的新版本\n"
        "- 如果用户上传了 Excel 或指定了已有 Excel，先调用 `read_workbook` 读取内容\n"
        "- 读取后直接说明当前环境未启用 Excel 脚本运行能力，无法完成新增公式或导出新版本\n"
        "- 不要承诺稍后可以导出文件；不要只要求用户补充公式位置或公式内容\n"
        "- 可以建议管理员启用 sandbox runtime，或让用户改成只读取、总结表格内容\n"
    )
