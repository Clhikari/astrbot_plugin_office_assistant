<p align="center">
  <img src="https://count.getloli.com/@Clhikariproject2?name=Clhikariproject2&theme=booru-touhoulat&padding=7&offset=0&align=center&scale=0.8&pixelated=1&darkmode=1" alt="Moe Counter">
</p>

# 📁 Office 助手

这是一个为 AstrBot 设计的 Office 助手插件。它赋予大语言模型（LLM）直接操作文件的能力，支持读取并分析多种格式文件，以及生成 Office 文档。

## 🌟 核心特性

- **智能文件分析**：LLM 可读取 Python, JS, HTML, Markdown, JSON 等文本文件，并对内容进行总结、分析，提供帮助或建议。
- **Office 文档生成**：原生支持生成 Word (.docx), Excel (.xlsx) 和 PowerPoint (.pptx) 文件。
- **PDF 转换**：支持 Office⇄PDF 双向转换（需额外依赖）。
- **流式处理**：大文件采用分块读取，避免内存溢出。
- **安全防护**：内置路径验证机制，防止路径遍历攻击。
- **精细权限控制**：
  - **白名单机制**：支持指定特定用户调用，留空则仅管理员可用。
  - **群聊安全逻辑**：支持"必须 @ 或引用机器人"才暴露工具，防止群聊误触发及 Token 浪费。

## 🛠️ 安装与配置

### Python 依赖

通过 AstrBot 插件管理器安装时，Python 依赖会自动安装

如需手动安装：

```bash
pip install python-docx openpyxl python-pptx pdfplumber pdf2docx docx2pdf pywin32
```

### Office→PDF 转换（系统依赖）

Office→PDF 转换需要额外的系统依赖：

**Windows（推荐）：** 需要已安装 Microsoft Office
- `docx2pdf` 支持 Word→PDF
- `pywin32` 支持 Word/Excel/PPT→PDF

**Windows（备选）：** 从 [LibreOffice 官网](https://www.libreoffice.org/download/) 下载安装（无需 MS Office）

**Linux/Docker：**

```bash
apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress
```

**macOS：**

```bash
brew install --cask libreoffice
```

### Docker 环境

```dockerfile
# 系统依赖（Office→PDF 转换）
RUN apt-get update && apt-get install -y \
    libreoffice-writer libreoffice-calc libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*
```

> 💡 **提示**：使用 `/pdf_status` 或 `/pdf状态` 命令可查看当前 PDF 转换功能的可用性和缺失依赖。

### 配置项说明

在 AstrBot 管理面板中，你可以配置以下内容：

#### 触发设置

| 配置项               | 类型 | 默认值 | 说明                                                     |
| -------------------- | ---- | ------ | -------------------------------------------------------- |
| 群聊需要@/引用机器人 | bool | true   | 开启后，群聊中仅在@/引用机器人时模型才会获得文件操作能力 |
| 发送文件时@用户      | bool | true   | 发送生成的文件时，是否@用户                              |

#### 权限管理

| 配置项     | 类型 | 默认值 | 说明                                      |
| ---------- | ---- | ------ | ----------------------------------------- |
| 用户白名单 | list | []     | 允许使用插件的用户 ID，留空则仅管理员可用 |

#### 功能开关

| 配置项               | 类型 | 默认值 | 说明                                             |
| -------------------- | ---- | ------ | ------------------------------------------------ |
| 启用 Office 文件生成 | bool | true   | 是否允许生成 Word/Excel/PPT 文件                 |
| 启用 PDF 转换        | bool | true   | 是否允许 Office⇄PDF 互转（需安装对应依赖才可用） |

#### 文件限制

| 配置项             | 类型  | 默认值 | 说明                                                               |
| ------------------ | ----- | ------ | ------------------------------------------------------------------ |
| 最大文件大小(MB)   | int   | 20     | 允许读取/发送的最大文件大小                                        |
| 发送后自动删除文件 | bool  | true   | 生成的文件发送后自动删除，关闭则持久化存储                         |
| 文件消息缓冲时间   | float | 4      | 收到文件后等待后续消息的时间(秒)，用于聚合分离发送的文件和文本消息 |

## 📖 提供的工具 (LLM Tools)

| 工具名称             | 功能描述                                                    |
| -------------------- | ----------------------------------------------------------- |
| `read_file`          | 读取文本文件内容，供 LLM 分析并提供总结、建议或帮助         |
| `create_office_file` | 生成 Office 文档（Word/Excel/PPT）                          |
| `convert_to_pdf`     | 将 Office 文件转换为 PDF（Windows: docx2pdf/pywin32，其他: LibreOffice） |
| `convert_from_pdf`   | 将 PDF 转换为 Word 或 Excel（需对应依赖）                   |

### 命令

| 命令          | 功能描述                        |
| ------------- | ------------------------------- |
| `/lsf`        | 查看工作区的 Office 文件        |
| `/rm`         | 永久删除指定文件                |
| `/fileinfo`   | 显示插件运行信息                |
| `/pdf_status` | 查看 PDF 转换功能状态和缺失依赖 |

### 支持的文件类型

#### 📖 可读取分析的文件格式

`.txt`, `.md`, `.json`, `.csv`, `.log`, `.py`, `.js`, `.html`, `.css`, `.xml`, `.yaml`, `.yml` 等

> LLM 会读取这些文件的内容，并根据用户需求进行：代码审查、内容总结、问题排查、优化建议等。

#### 📝 可生成的文件格式（仅 Office）

| 格式       | 扩展名  | 类型参数     | 说明               |
| ---------- | ------- | ------------ | ------------------ |
| Word       | `.docx` | `word`       | 支持多段落文档生成 |
| Excel      | `.xlsx` | `excel`      | 支持表格数据生成   |
| PowerPoint | `.pptx` | `powerpoint` | 支持多页幻灯片生成 |

> ⚠️ **注意**：本插件仅支持生成 Office 文件，不支持生成普通文本文件（如 .txt, .py 等）。

> ⚠️ **功能限制**：目前仅支持生成**简单的** Office 文件（基本表格、纯文本文档、简单幻灯片）。不支持复杂样式、图表等高级功能。

#### 🔄 PDF 转换支持

| 转换方向   | 依赖                                        | 说明                                 |
| ---------- | ------------------------------------------- | ------------------------------------ |
| Office→PDF | docx2pdf/pywin32 (Windows) 或 LibreOffice   | 支持 Word/Excel/PPT 转 PDF           |
| PDF→Word   | pdf2docx                                    | 适用于文本为主的 PDF                 |
| PDF→Excel  | pdfplumber 或 tabula-py                     | 仅提取表格数据，非表格会丢失         |

> Windows 用户推荐使用 docx2pdf（需已安装 MS Office），体积小、转换质量好。

## 💬 使用示例

与机器人对话即可触发文件操作：

```
用户：看看工作区有什么文件
机器人：📂 工作区文件列表：
       1. main.py (2.1 KB) - 2025-12-25 10:30
       2. config.json (256 B) - 2025-12-25 09:15
       共 2 个文件

用户：帮我看看 main.py 有没有什么问题
机器人：我来分析一下这个文件...

       📋 代码审查结果：
       1. 第15行：建议添加异常处理
       2. 第23行：变量命名可以更具描述性
       ...

用户：帮我做一个Excel表格，包含本月销售数据
机器人：好的，我来为你生成 Excel 文件。
       ✅ 文件已处理成功：sales_report.xlsx
       [sales_report.xlsx 文件]

用户：把这个 Excel 转成 PDF
机器人：✅ 已将 sales_report.xlsx 转换为 PDF
       [sales_report.pdf 文件]

用户：把这个 PDF 转成 Word
机器人：✅ 已将 document.pdf 转换为 Word 文档
       [document.docx 文件]
```

## ⚠️ 注意事项

1. **安全性**：所有文件操作均限制在插件工作区目录内，无法访问系统其他目录。
2. **写入限制**：本插件仅支持生成 Office 文档，不支持写入/生成普通文本文件。
3. **文件大小**：建议根据服务器内存合理配置最大文件大小，过大可能导致发送失败。
4. **Office 依赖**：如未安装 Office 相关库，对应格式的生成功能将自动禁用，插件会提示所需的包名。
5. **权限配置**：建议在公开群聊中启用"群聊需要@机器人"选项，避免意外触发。
6. **PDF 转换限制**：
   - PDF→Word：复杂布局的 PDF 转换后可能有偏差
   - PDF→Excel：仅能提取 PDF 中的表格，非表格内容会丢失
   - Office→PDF：Windows 推荐 docx2pdf（需 MS Office）；Linux/Docker 需安装 LibreOffice

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！
