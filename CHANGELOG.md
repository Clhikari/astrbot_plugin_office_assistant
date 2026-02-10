# Changelog

本文件记录 `astrbot_plugin_office_assistant` 的重要更新。

格式参考 Keep a Changelog，版本号遵循语义化版本（SemVer）。

## [Unreleased]

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
