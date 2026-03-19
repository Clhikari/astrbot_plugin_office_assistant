# Changelog

本文件记录 `astrbot_plugin_office_assistant` 的重要更新。

格式参考 Keep a Changelog，版本号遵循语义化版本（SemVer）。

## [v1.3.1] - 2026-03-19

这一轮更新主要是把插件入口进一步收敛成薄编排层，把运行时、策略和文件处理职责系统化地下沉到 `services/`，同时补齐回归测试，并更新项目图示以匹配当前实现。

### Added

- 新增 `services/` 运行时服务层，补齐 `runtime_builder`、访问策略、LLM 请求策略、上传会话、消息入口、工作区、文件工具、交付发送、命令处理、错误钩子等服务模块。
- 新增 `tests/test_office_assistant_services.py`，集中覆盖 service 层行为和重构后的关键回归路径。

### Changed

- `main.py` 进一步收敛为薄入口层，原先堆在入口中的编排、策略和文件处理逻辑改为委托给 `services/`。
- 复杂 Word 工作流继续保持 `create_document -> add_blocks -> finalize_document -> export_document` 四步公开 contract，不重新放大工具面。
- 上传文件场景的引导、recent text 缓存和工具调用顺序约束进一步收紧，减少先建文档再读文件或乱跳工具的情况。
- 系统提示和相关工具说明调整为更短的中文版本，降低中英混杂提示泄漏到用户侧的概率。
- 更新项目架构图和贡献流程图，使文档展示与当前 `main.py + services/ + core` 分层保持一致。

### Fixed

- 终止型文件工具在成功交付结果后会直接结束 tool loop，避免模型在导出或发文件后继续重复调用其他工具。

## [v1.3.0] - 2026-03-17

这一版主要把复杂 Word 的生成方式改成了分步工作流，同时把相关模块拆开，补上了测试，也顺手把一些实现细节理顺了。

### Added

- 增加复杂 Word 四步工具链：`create_document -> add_blocks -> finalize_document -> export_document`。
- 增加统一块模型，`add_blocks` 可按顺序追加 `heading`、`paragraph`、`list`、`table`、`summary_card`、`page_break`、`group`、`columns` 八类内容块。
- 增加文档会话存储 `DocumentSessionStore`，支持草稿创建、块追加、定稿、导出和导出状态管理。
- 增加 `document_core/`，把文档模型、Word 渲染器和卡片宏单独拆出来。
- 增加 Word 主题和样式参数，支持 `business_report`、`project_review`、`executive_brief` 三套主题，以及表格模板、密度、强调色等配置。
- 增加 MCP 文档工具层和 Agent 工具层，复杂 Word 工作流可以通过统一协议暴露给上层调用。
- 增加针对文档工具链、`before_llm_chat` 工具注入逻辑、Office 生成器兼容路径的专项测试。

### Changed

- 原先分散的 Word 工具调用方式收敛成四步链路，`add_heading`、`add_paragraph`、`add_table`、`add_summary_card` 等能力统一归入 `add_blocks`。
- `before_llm_chat` 会按权限和上下文动态注入文档工具、补系统提示，并在插件实际生效时自动隐藏执行类工具。
- `create_office_file` 还保留，但已经标记为 deprecated；复杂 Word 场景改为优先走四步工具链，简单一次性输出继续兼容旧入口。
- `office_generator.py` 现在优先走新的 `DocumentModel + WordDocumentBuilder` 路径，失败时再回退旧版生成逻辑。
- 项目目录从偏扁平调整为分层结构，新增 `agent_tools/`、`document_core/`、`mcp_server/`、`tests/` 等目录。
- `README.md`、`CHANGELOG.md` 和复杂 Word 相关说明同步重写，文档表述与新工具链保持一致。

### Improved

- 调整复杂 Word 导出路径与沙箱路径处理逻辑，减少导出后文件不可访问或图片路径异常的情况。
- 调整摘要卡片宏展开逻辑，避免卡片内容结构异常。
- 提高文档块输入的容错性，遇到无效块时不容易把整条链路带崩。
- 优化文件消息与后续文本消息的衔接处理，改善“先发文件再发指令”的识别稳定性。
- 调整事件处理优先级，减少插件冲突和文件拦截时序问题。

### Notes

- `create_office_file` 目前仍可用于简单 Word/Excel/PPT 生成，复杂 Word 场景建议改走四步工具链。

## [v1.2.4] - 2026-02-19

### Added

- 新增报错移动到指定会话功能，便于统一处理错误信息，提升用户体验（此功能需要框架版本到v4.17.5以上）。

## [v1.2.3] - 2026-02-14

### Fixed

- 修复LLM对话被意外禁用的问题，保持正常的 LLM 对话功能可用

## [v1.2.2] - 2026-02-10

### Fixed

- 修复上传文件名路径穿越风险：上传时仅使用 basename 并进行路径校验，避免越界写入。
- 补充 PDF -> Excel 依赖声明：新增 `pandas` 与 `tabula-py`（可选，需 Java）。

## [v1.2.1] - 2026-02-04

### Changed

- 重构 PDF 转换器中的 COM 资源管理，统一使用 `com_application`。
- 优化配置预加载与条件逻辑，减少重复读取配置。
- 代码格式化与 import 顺序整理。

## [v1.2.0] - 2026-01-13

### Added

- 新增文件预览图生成器。
- 增加预览图相关配置项。
- 集成 PDF 文本提取能力。

### Changed

- 更新 README 文档。

### Fixed

- 修复 `docx2pdf` 的 COM 初始化问题。

## [v1.1.2] - 2026-01-06

### Added

- 增加旧格式 Office 跨平台支持（`.doc` / `.xls` / `.ppt` 相关处理增强）。
- 增补依赖与工具函数完善。

### Changed

- 完善依赖安装说明与 README。

## [v1.1.1] - 2026-01-04

### Changed

- 引入 `requirements.txt`，简化依赖安装。
- 代码风格与 import 顺序整理。

### Fixed

- 优化文件类型检测，不支持格式不再触发插件读取流程。

## [v1.1.0] - 2026-01-03

### Added

- 新增 PDF 转换能力（Office <-> PDF 的基础支持）。
- 新增 PDF 转换配置项。
- 增强错误信息安全处理与文件名清理逻辑。

### Changed

- 统一线程池与部分模块结构优化。

## [v1.0.0] - 2025-12-30

### Added

- 首个稳定版本发布。
- 提供 Office 文件读取与生成基础能力。
- 提供消息缓冲与权限控制基础机制。

---

注：`Unreleased` 记录的是已进入仓库但尚未打 tag 的变更。
