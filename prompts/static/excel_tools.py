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
        "\n"
        "[硬约束] 禁止滥用 execute_excel_script\n"
        "- 如果最终输出不包含公式、图表、条件格式或数据验证，且不涉及编辑已有工作簿，禁止使用 `execute_excel_script`\n"
        "- 即使需要随机数据、大量行或复杂计算逻辑，只要输出是纯值表格，必须用 `write_rows`\n"
        "- `execute_excel_script` 仅用于输出文件本身需要 Excel 高级特性的场景，或需要编辑已有工作簿（保留原有布局/样式）的场景\n"
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
        "\n"
        "[关键规则] 保存方式\n"
        "- 必须使用 `save_output_workbook(workbook)` 保存文件，这是唯一正确的保存方式\n"
        "- 禁止使用 `workbook.save(...)` 或任何其他保存方式，否则脚本会报错\n"
        "- `save_output_workbook` 会自动保存到正确路径并美化新建 Sheet\n"
        "\n"
        "[辅助函数]\n"
        "- `load_input_workbook()`: 读取输入工作簿（保留样式）\n"
        "- `save_output_workbook(workbook)`: 保存到输出路径（唯一保存方式）\n"
        "- `format_table_sheet(sheet)`: 美化普通数据表（表头加粗、冻结首行、启用筛选）\n"
        "- `ensure_readable_sheet(sheet)`: 调整列宽行高和自动换行\n"
        "- `auto_format_workbook(workbook)`: 美化整本工作簿所有 Sheet\n"
        "- `preserve_input_sheet_layout(workbook)`: 把输入 Sheet 的样式同步到同名输出 Sheet（仅用于 Workbook() 新建场景）\n"
        "\n"
        "[输入输出]\n"
        "- 输入文件通过 `input_files` 变量提供（相对路径列表），按顺序使用 `input_files[0]`\n"
        "- 输出路径通过 `output_path` 变量提供，不要自己另取文件名\n"
        "- 成功时只能二选一：设置 `result_text` 返回文本，或写出 `output_path` 返回文件\n"
        "\n"
        "[常见错误]\n"
        "- 禁止 `wb.save('xxx.xlsx')` -> 必须 `save_output_workbook(wb)`\n"
        "- 禁止硬写输入文件名 -> 必须用 `input_files[0]`\n"
        "- 禁止用静态 PatternFill 逐单元格着色来实现条件高亮 -> 必须用 `CellIsRule` 或 `FormulaRule` 条件格式\n"
        "- Excel 公式文本常量必须用双引号：`\"Keyboard\"` 不是 `'Keyboard'`\n"
        "- 写公式时必须处理空值和 0 值：用 `IF`/`IFERROR` 包裹\n"
        "\n"
        "[行为约束]\n"
        "- 调用 `execute_excel_script` 时，assistant content 必须为空\n"
        "- 脚本正文只能放在 tool arguments 的 `script` 字段\n"
        "- 保留用户可见的原 Sheet 名，不要改成内部名\n"
        "- 基于输入工作簿时用 `load_input_workbook()`，不要用 `Workbook()` 重建\n"
        "- 新建表格 Sheet 时至少用 `format_table_sheet(sheet)`\n"
        "- 失败后根据错误修正脚本，不要先发角色表演\n"
        "- 如果结果包含 `requires_review: true`，能修正时继续调用；无法修正时逐条列出警告\n"
        "- 脚本最多重试 3 次；超过后停止重试，说明失败原因\n"
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
    elif scenario == "conditional_format":
        lines.extend(
            [
                "- 必须使用 openpyxl 的条件格式 API，禁止用静态 PatternFill 逐单元格着色\n",
                "- 静态填充在数据修改后不会自动更新，条件格式会随数据变化动态生效\n",
                "- 使用 `from openpyxl.formatting.rule import CellIsRule, FormulaRule` 和 `from openpyxl.styles import PatternFill`\n",
                "- 示例：`sheet.conditional_formatting.add('B2:B11', CellIsRule(operator='lessThan', formula=['60'], fill=PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')))`\n",
                "- 条件格式的范围必须精确到数据区域（如 B2:B11），不要用整列引用如 B:B\n",
            ]
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
