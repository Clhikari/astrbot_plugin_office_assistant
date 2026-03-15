# Changelog

本文件记录 `astrbot_plugin_office_assistant` 的重要更新。

格式参考 Keep a Changelog，版本号遵循语义化版本（SemVer）。

## [Unreleased]

### Added

- 新增群聊总开关 `enable_features_in_group`，可在群聊中一键禁用插件能力（工具、命令、文件拦截）。
- 新增执行工具屏蔽配置 `auto_block_execution_tools`，支持在插件能力生效时自动隐藏 `astrbot_execute_shell` / `astrbot_execute_python` / `astrbot_execute_ipython`。
- 新增路径访问配置 `allow_external_input_files`，允许按配置读取/转换工作区外绝对路径输入文件（删除仍限制在工作区内）。
- 增加会话级最近文本缓存与文件消息聚合补全逻辑，优化“先发文件后发文本”场景下的指令衔接。
- 新增复杂 Word 状态化工具链：`create_document`、`add_heading`、`add_paragraph`、`add_table`、`add_summary_card`、`finalize_document`、`export_document`。
- 新增复杂 Word 的主题、表格模板、密度和强调色参数支持。
- 新增复杂 Word 导出后自动回传文件的插件内发送链路。

### Changed

- 命令名统一为 `/list_files` 与 `/delete_file`，避免与部分平台通用命令冲突。
- 文件上传入库逻辑重构为安全命名+重名递增，减少覆盖风险。
- `README.md` 重构：新增目录跳转、配置速查、迁移说明、FAQ、社群答疑入口。
- 复杂 Word 的对外公开路径收敛为精细控制工具链，不再默认暴露章节级快捷写入工具。
- `before_llm_chat` 现在会在合适场景下注入复杂 Word 的状态化使用提示，并将简单文档与复杂 Word 路径区分开。

### Fixed

- 修复 `EXECUTION_TOOLS` 未导入导致的运行期异常风险。
- 修复执行工具屏蔽作用范围：仅在插件功能实际生效时屏蔽，避免误伤无权限会话。
- 修复 `auto_block_execution_tools` 在文档、配置 schema 与代码默认值不一致的问题。

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
