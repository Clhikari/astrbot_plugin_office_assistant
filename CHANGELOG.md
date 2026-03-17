# Changelog

本文件记录 `astrbot_plugin_office_assistant` 的重要更新。

格式参考 Keep a Changelog，版本号遵循语义化版本（SemVer）。

## [v1.3.0] - 2026-03-17

这个版本的核心目标不是再补几个零散接口，而是把复杂 Word 生成从“一次性输出”升级为“可分步构建、可持续追加、可控导出”的状态化工作流。

### Added

- 新增复杂 Word 四步工具链：`create_document -> add_blocks -> finalize_document -> export_document`。
- 新增统一块模型，`add_blocks` 可按顺序追加 `heading`、`paragraph`、`list`、`table`、`summary_card`、`page_break`、`group`、`columns` 八类内容块。
- 新增文档会话存储 `DocumentSessionStore`，支持草稿创建、块追加、定稿、导出和导出状态管理。
- 新增 Word 文档核心模块 `document_core/`，将文档模型、渲染器和卡片宏从主流程中拆分出来。
- 新增 Word 主题和样式能力，支持 `business_report`、`project_review`、`executive_brief` 三套主题，以及表格模板、密度、强调色等参数。
- 新增 MCP 文档工具层与 Agent 工具层，复杂 Word 工作流可通过统一协议暴露给上层调用。
- 新增针对文档工具链、`before_llm_chat` 工具注入逻辑、Office 生成器兼容路径的专项测试。

### Changed

- 原先的多工具 Word 流程被收敛为四步链路，`add_heading`、`add_paragraph`、`add_table`、`add_summary_card` 等分散能力统一归入 `add_blocks`。
- `before_llm_chat` 现在会按权限和上下文动态注入文档工具、补充系统提示，并在插件实际生效时自动收敛执行类工具。
- `create_office_file` 仍保留，但已明确标记为 deprecated；复杂 Word 场景优先走四步工具链，简单一次性输出继续兼容旧入口。
- `office_generator.py` 现在优先走新的 `DocumentModel + WordDocumentBuilder` 路径，失败时再回退旧版生成逻辑，兼顾新能力和兼容性。
- 项目目录结构从偏扁平改为分层设计，新增 `agent_tools/`、`document_core/`、`mcp_server/`、`tests/` 等目录，主流程、协议层、文档核心层和测试层职责更清晰。
- `README.md`、`CHANGELOG.md` 和复杂 Word 相关说明整体重写，文档表述与新工具链保持一致。

### Fixed

- 修复复杂 Word 导出路径与沙箱路径处理问题，减少导出后文件不可访问或图片路径异常的情况。
- 修复摘要卡片宏展开错误，避免卡片内容结构异常。
- 修复文档块容错问题，遇到无效块输入时不再轻易导致整条生成链路失败。
- 修复文件消息与后续文本消息的衔接处理，改善“先发文件再发指令”的识别稳定性。
- 修复事件处理优先级带来的插件冲突和文件拦截时序问题。

### Upgrade Notes

- `create_office_file` 目前仍可用于简单 Word/Excel/PPT 生成，但后续复杂 Word 能力将以四步链为主继续演进。
- 本次版本已经具备明显的能力升级和结构升级，建议作为 `v1.3.0` 发布，而不是继续沿用 `v1.2.x` 补丁版本号。

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
- 代码格式化与清理。

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
