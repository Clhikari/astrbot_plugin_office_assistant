<p align="center">
  <img src="https://count.getloli.com/@Clhikariproject2?name=Clhikariproject2&theme=booru-touhoulat&padding=7&offset=0&align=center&scale=0.8&pixelated=1&darkmode=1" alt="Moe Counter">
</p>

# 📁 Office 助手（astrbot_plugin_office_assistant）

一个给 AstrBot 用的「文件魔法工房」插件。目标很直接：让机器人更稳地读文件、做文档、跑转换。

它主要做三件事：

1. 让模型读取并分析文件内容（文本 / 代码 / Office / PDF）。
2. 让模型按结构化内容生成 Word、Excel、PPT。
3. 提供 Office 与 PDF 的双向转换，并在发送时可附带预览图。

---

## 目录

- [🌟 快速了解](#快速了解)
- [🚀 快速开始](#快速开始)
- [⚙️ 配置说明](#配置说明)
- [🛠️ 工具与命令](#工具与命令)
- [📚 支持的文件格式](#支持的文件格式)
- [🧱 系统依赖安装](#系统依赖安装)
- [🐳 Docker 使用建议](#docker-使用建议)
- [🛡️ 安全与行为说明](#安全与行为说明)
- [❓ 常见问题（Q/A）](#常见问题qa)
- [💬 社群答疑与新想法](#社群答疑与新想法)
- [🧭 升级与迁移说明](#升级与迁移说明)
- [📄 许可证](#许可证)
- [🤝 贡献](#贡献)

---

## 快速了解

当前版本的核心行为：

- 群聊默认不启用插件能力（`enable_features_in_group=false`）。
- 若群聊启用后，默认仍要求 `@` / 引用机器人才会暴露文件工具（`require_at_in_group=true`）。
- 默认会在插件能力生效时隐藏执行类工具：
  - `astrbot_execute_shell`
  - `astrbot_execute_python`
  - `astrbot_execute_ipython`
- 默认禁止读取工作区外路径；可通过配置开启外部绝对路径读取（仅对 `read_file`/PDF 转换生效）。

---

## 快速开始

### 1 📦 安装插件（开箱）

通过 AstrBot 插件管理器安装即可。Python 依赖会随插件自动安装。

### 2 ✅ 启动后先做两条检查（验机）

- `/fileinfo`：查看插件当前运行状态与配置生效情况。
- `/pdf_status`（或 `/pdf状态`）：查看 PDF 转换可用性与缺失依赖。

### 3 🧪 最小可用验证（跑一遍就安心）

- 发一个 `.txt` 或 `.md` 给机器人，要求它读取并总结。
- 让机器人生成一个 `.xlsx`。
- 如果安装了转换依赖，再试一次 Office -> PDF。

---

## 配置说明

以下配置均在 AstrBot 管理面板中设置。

### 🔔 触发设置（`trigger_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 群聊需要@/引用机器人 (`require_at_in_group`) | bool | true | 群聊中仅在 `@` / 引用机器人时暴露文件工具。 |
| 群聊启用插件功能 (`enable_features_in_group`) | bool | false | 关闭时，群聊里本插件完全不生效（工具/命令/文件拦截全部关闭）。 |
| 自动屏蔽 shell/python 工具 (`auto_block_execution_tools`) | bool | true | 开启后，在插件功能生效时自动隐藏 `astrbot_execute_*` 三个执行工具。 |
| 发送文件时@用户 (`reply_to_user`) | bool | true | 机器人发送文件时是否 `@` 发起人。 |

### 🔐 权限管理（`permission_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 用户白名单 (`whitelist_users`) | list | [] | 允许使用插件的用户 ID；留空时仅管理员可用。 |

### 🧩 功能开关（`feature_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 启用 Office 文件生成 (`enable_office_files`) | bool | true | 是否允许 `create_office_file`。 |
| 启用 PDF 转换 (`enable_pdf_conversion`) | bool | true | 是否允许 Office <-> PDF 转换（仍需系统依赖）。 |

### 📏 文件限制（`file_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 最大文件大小MB (`max_file_size_mb`) | int | 20 | 读取/发送文件大小上限。 |
| 发送后自动删除文件 (`auto_delete_files`) | bool | true | 发送后删除生成文件；关闭则持久化到插件工作区。 |
| 文件消息缓冲时间秒 (`message_buffer_seconds`) | float | 4 | 用于聚合“先发文件后发文本”的场景。 |

### 🛣️ 路径访问（`path_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 允许外部绝对路径 (`allow_external_input_files`) | bool | false | 开启后，`read_file` / `convert_to_pdf` / `convert_from_pdf` 可访问工作区外绝对路径；`delete_file` 仍只允许工作区内删除。 |

### 🖼️ 预览图（`preview_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 启用预览图 (`enable`) | bool | true | 发送 Office/PDF 时尝试发送第一页预览图。 |
| 预览图分辨率 (`dpi`) | int | 150 | 推荐 100~200。 |

---

## 工具与命令

### 🤖 LLM 工具

| 工具名 | 作用 |
| --- | --- |
| `read_file` | 读取文本、代码、Office、PDF 内容。 |
| `create_office_file` | 生成 Word / Excel / PowerPoint 文件。 |
| `convert_to_pdf` | Office -> PDF。 |
| `convert_from_pdf` | PDF -> Word 或 Excel。 |

### ⌨️ 插件命令

| 命令 | 别名 | 作用 |
| --- | --- | --- |
| `/list_files` | `/file_ls`, `/文件列表` | 查看工作区中的 Office 文件。 |
| `/delete_file <文件名>` | `/file_rm`, `/删除文件` | 删除工作区内文件。 |
| `/fileinfo` | 无 | 查看运行状态、开关状态、工作目录等信息。 |
| `/pdf_status` | `/pdf状态` | 查看 PDF 转换可用性与缺失依赖。 |

---

## 支持的文件格式

### 📖 可读取分析

- Office：`.docx` `.xlsx` `.pptx` `.doc` `.xls` `.ppt`
- PDF：`.pdf`
- 文本/代码：
  - `.txt` `.md` `.log` `.rst`
  - `.json` `.yaml` `.yml` `.toml` `.xml` `.csv`
  - `.html` `.css` `.sql` `.sh` `.bat`
  - `.py` `.js` `.ts` `.jsx` `.tsx` `.c` `.cpp` `.h` `.java` `.go` `.rs`

### 📝 可生成

- Word：`.docx`
- Excel：`.xlsx`
- PowerPoint：`.pptx`

说明：当前只支持基础内容生成，不覆盖复杂排版、图表、动画等高级能力。

### 🔄 PDF 转换

| 转换方向 | 依赖 | 说明 |
| --- | --- | --- |
| Office -> PDF | Windows: `docx2pdf`/`pywin32`（需 Office）；或 LibreOffice | 支持 Word/Excel/PPT 转 PDF |
| PDF -> Word | `pdf2docx` | 文本型 PDF 效果更好 |
| PDF -> Excel | `tabula-py`（需 Java）或 `pdfplumber` | 以表格提取为主，复杂版面会有损失 |

---

## 系统依赖安装

> Python 包大多会自动安装。下面是系统层依赖（需手动）。

### 🪟 Windows

```bash
# 旧格式读取 + win32com 转换
pip install pywin32

# Word 优先的 Office->PDF
pip install docx2pdf
```

也可以安装 LibreOffice 作为跨平台转换后端：

- <https://www.libreoffice.org/download/>

### 🐧 Linux（Debian/Ubuntu）

```bash
# 读取 .doc
apt install antiword

# Office -> PDF
apt install libreoffice-writer libreoffice-calc libreoffice-impress
```

### 🍎 macOS

```bash
# 读取 .doc
brew install antiword

# Office -> PDF
brew install --cask libreoffice
```

安装系统依赖后，请重启 AstrBot。

---

## Docker 使用建议

### 方案 A：⚡ 进容器安装（快，但不持久）

```bash
docker exec -it <容器名> bash
apt-get update
apt-get install -y antiword
apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress
```

容器删除重建后需要重装。

### 方案 B：🧰 写入 Dockerfile（推荐）

```dockerfile
RUN apt-get update && apt-get install -y \
    antiword \
    libreoffice-writer libreoffice-calc libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*
```

---

## 安全与行为说明

1. 默认只允许访问插件工作区。
2. 即使开启“允许外部绝对路径”，也只放开读取/转换输入文件，不放开删除。
3. 群聊默认关闭插件；建议按需启用，避免误触发。
4. 若白名单为空，默认只有管理员可用。
5. 大文件会按块读取，超过上限会拒绝并返回提示。

---

## 常见问题（Q/A）

### Q1：为什么在群聊里插件完全没反应？

A：先检查 `enable_features_in_group`。默认是 `false`，群聊里会直接禁用插件能力。

### Q2：我已经在群里开启插件，为什么还要 `@` 才能用？

A：`require_at_in_group` 默认是 `true`。如果不想强制 `@`，把它关掉。

### Q3：为什么我在私聊看不到文件工具？

A：检查白名单。白名单为空时，默认只有管理员可用。

### Q4：为什么 `/rm`、`/lsf` 不能用了？

A：这两个命令名容易和平台或其他插件冲突。当前统一使用：

- `/delete_file`
- `/list_files`

### Q5：为什么没有预览图？

A：常见原因有三类：

- 关闭了 `preview_settings.enable`
- 缺少 `pymupdf`
- Office 预览目前仅支持 Windows 上的 Word（`.doc/.docx`）

### Q6：为什么 PDF->Excel 结果是空表或结构很乱？

A：PDF->Excel 本质是“表格提取”，对扫描件、跨页表格、复杂版面不稳定。可先尝试：

- 更干净的原始 PDF
- 安装 Java + `tabula-py`
- 或改走 PDF->Word 再人工整理

### Q7：为什么 Office->PDF 失败？

A：先看 `/pdf_status` 输出。Windows 建议安装 Office + `pywin32`/`docx2pdf`；Linux/macOS 建议安装 LibreOffice。

### Q8：开启“允许外部绝对路径”后，为什么还是删不了工作区外文件？

A：这是刻意限制。外部路径只允许读取和转换输入，不允许删除。

### Q9：这个插件现在还有哪些不足？

A：有，而且是已知问题，先说清楚：

- Office 文件生成目前偏基础，复杂排版、图表、动画不是强项。
- PDF -> Excel 是表格提取路线，扫描件/跨页/复杂版式容易失真。
- Office 预览图对平台有差异：目前 Word 预览在 Windows 上更稳。
- 环境依赖对转换能力影响很大，尤其是 LibreOffice / Java / Office 组件是否齐全。

---

## 社群答疑与新想法

如果你在群聊里遇到奇怪行为、想提新点子，欢迎直接来群里聊：

- QQ 群：`1072198212`

---

## 升级与迁移说明

如果你从旧版本迁移，优先检查两件事：

1. 命令名是否已改为 `/list_files`、`/delete_file`。
2. 群聊默认行为是否符合预期（默认关闭插件能力）。

---

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。若是行为变更，请附上复现步骤和预期结果，便于快速定位。
