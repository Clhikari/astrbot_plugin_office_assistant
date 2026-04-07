<p align="center">
  <img src="https://count.getloli.com/@Clhikariproject2?name=Clhikariproject2&theme=booru-touhoulat&padding=7&offset=0&align=center&scale=0.8&pixelated=1&darkmode=1" alt="Moe Counter">
</p>

# Office 助手（astrbot_plugin_office_assistant）

AstrBot 的文件处理插件。读文件、生成 Office 文档、做格式转换，发送时可带预览图。

---

## 目录

- [快速了解](#快速了解)
- [快速开始](#快速开始)
- [配置说明](#配置说明)
- [工具与命令](#工具与命令)
- [复杂 Word 工作流](#复杂-word-工作流)
- [后续规划](#后续规划)
- [支持的文件格式](#支持的文件格式)
- [系统依赖安装](#系统依赖安装)
- [Docker 部署](#docker-部署)
- [安全说明](#安全说明)
- [常见问题](#常见问题)
- [交流](#交流)
- [升级与迁移](#升级与迁移)
- [许可证](#许可证)
- [贡献](#贡献)

---

## 快速了解

当前版本的默认行为：

- 群聊默认不启用插件（`enable_features_in_group=false`）。
- 群聊启用后，需要 `@` 或引用机器人才暴露文件工具（`require_at_in_group=true`）。
- 群聊里如果先上传文件、后面再处理，建议直接用 `/doc` 这组命令。
- 插件生效时自动隐藏 `astrbot_execute_shell`、`astrbot_execute_python`、`astrbot_execute_ipython`。
- 默认禁止访问工作区外路径，可配置放开（仅对 `read_file` 和 PDF 转换生效）。
- 复杂 Word 走四步工具链：`create_document → add_blocks → finalize_document → export_document`。`add_blocks` 支持标题、正文、列表、表格、卡片、分页、分节、目录、分组、分栏。
- `create_office_file` 还可以继续用；它的 Word 导出现在也走同一套 JS 渲染链。
- Word 导出现在只走 JS 渲染器，不再带 Python 回退。

---

## 快速开始

### 1. 安装

通过 AstrBot 插件管理器安装。Python 依赖会自动装好。

### 2. 检查状态

- `/fileinfo`：看插件运行状态和配置。
- `/pdf_status`（或 `/pdf状态`）：看 PDF 转换是否可用、缺哪些依赖。

### 3. 试一下

- 发一个 `.txt` 或 `.md` 给机器人，让它读取并总结。
- 让机器人生成一个 `.xlsx`。
- 装了转换依赖的话，试一次 Office → PDF。
- 群聊里想先传文件再下指令，可以试试 `/doc list` 和 `/doc use f1 你的要求`。

---

## 配置说明

在 AstrBot 管理面板中设置。

常用的几项：

| 配置项 | 默认值 | 什么情况下改 |
| --- | --- | --- |
| `enable_features_in_group` | `false` | 要在群聊里用插件 |
| `require_at_in_group` | `true` | 群聊里不想强制 `@` |
| `enable_docx_image_review` | `true` | 不需要模型读 Word 里的图片 |
| `auto_block_execution_tools` | `true` | 不想自动隐藏执行类工具 |
| `allow_external_input_files` | `false` | 要读工作区外的文件 |
| `enable_pdf_conversion` | `true` | 不需要 Office/PDF 转换 |
| `auto_delete_files` | `true` | 想保留生成的文件 |
| `word_style_settings.default_font_name` | 空 | 想统一改 Word 默认字体，比如 Arial |

> [!TIP]
> 刚装好的话，一般只需要看 `enable_features_in_group`、`require_at_in_group`、`allow_external_input_files`、`enable_docx_image_review` 这四个。

<details>
<summary>完整配置表</summary>

### 触发设置（`trigger_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 群聊需要@/引用机器人 (`require_at_in_group`) | bool | true | 群聊中 `@` 或引用机器人时才暴露文件工具 |
| 群聊启用插件功能 (`enable_features_in_group`) | bool | false | 关了的话群聊里插件完全不生效 |
| 自动屏蔽 shell/python 工具 (`auto_block_execution_tools`) | bool | true | 插件生效时隐藏 `astrbot_execute_*` 系列工具 |
| 发送文件时@用户 (`reply_to_user`) | bool | true | 发文件时是否 `@` 发起人 |

### 权限管理（`permission_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 用户白名单 (`whitelist_users`) | list | [] | 允许使用插件的用户 ID，留空则仅管理员可用 |

### 功能开关（`feature_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 启用 Office 文件生成 (`enable_office_files`) | bool | true | 是否允许 `create_office_file` |
| 启用 PDF 转换 (`enable_pdf_conversion`) | bool | true | 是否允许 Office ↔ PDF 转换（需系统依赖） |

### 文件限制（`file_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 最大文件大小MB (`max_file_size_mb`) | int | 20 | 读取和发送的大小上限 |
| 启用 Word 图片理解 (`enable_docx_image_review`) | bool | true | 读 `.docx` 时把嵌入图片注入上下文，关了就按纯文本读 |
| Word图片注入大小上限MB (`max_inline_docx_image_mb`) | int | 2 | 单张图超过这个大小就跳过 |
| Word图片最多注入张数 (`max_inline_docx_image_count`) | int | 3 | 最多注入几张图到上下文 |
| 发送后自动删除文件 (`auto_delete_files`) | bool | true | 发完就删，关了就留在工作区 |
| 文件合并等待时间秒 (`message_buffer_seconds`) | float | 4 | 上传文件后，系统会先等一会儿，把同一波发来的文件合在一起 |
| 旧流程文本缓存时间秒 (`recent_text_ttl_seconds`) | int | 20 | 兼容旧流程用。现在主要影响“文件和文字一起进来”的老用法，平时一般不用动 |
| 上传文件保留时间秒 (`upload_session_ttl_seconds`) | int | 600 | 群聊里上传完文件后，这些文件还能在当前会话里保留多久，供 `/doc list`、`/doc use`、`/doc clear` 使用 |

### 路径访问（`path_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 允许外部绝对路径 (`allow_external_input_files`) | bool | false | 放开后 `read_file` 和转换工具可以访问工作区外路径，`delete_file` 不受影响 |

### 预览图（`preview_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 启用预览图 (`enable`) | bool | true | 发 Office/PDF 时带首页预览图 |
| 预览图分辨率 (`dpi`) | int | 150 | 推荐 100~200 |

### Word 默认样式（`word_style_settings`）

| 配置项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| 正文默认字体 (`default_font_name`) | string | 空 | 模型没显式指定字体时使用，比如 `Arial` |
| 标题默认字体 (`default_heading_font_name`) | string | 空 | 留空时跟随正文默认字体 |
| 表格默认字体 (`default_table_font_name`) | string | 空 | 留空时跟随正文默认字体 |
| 代码默认字体 (`default_code_font_name`) | string | 空 | 留空时继续使用内置默认值 |

</details>

---

## 工具与命令

### LLM 工具

| 工具名 | 干什么 |
| --- | --- |
| `read_file` | 读文本、代码、Office、PDF 的内容 |
| `create_office_file` | 生成 Word / Excel / PPT（deprecated，Word 建议走四步链） |
| `create_document` | 新建 Word 草稿，定主题、表格模板、密度、强调色和文档级页眉页脚 |
| `add_blocks` | 往草稿里加内容块（标题、正文、列表、表格、卡片、分页、分节、目录、分组、分栏） |
| `finalize_document` | 锁定草稿 |
| `export_document` | 导出 .docx，自动发给用户 |
| `convert_to_pdf` | Office → PDF |
| `convert_from_pdf` | PDF → Word 或 Excel |

### 插件命令

| 命令 | 别名 | 干什么 |
| --- | --- | --- |
| `/list_files` | `/file_ls`, `/文件列表` | 看工作区里的 Office 文件 |
| `/delete_file <文件名>` | `/file_rm`, `/删除文件` | 删工作区里的文件 |
| `/fileinfo` | 无 | 看运行状态和配置 |
| `/pdf_status` | `/pdf状态` | 看 PDF 转换是否可用 |
| `/doc list` | 无 | 看当前会话里还能继续用的上传文件 |
| `/doc clear [文件ID]` | 无 | 清空当前会话文件，或只删一个文件 |
| `/doc use [文件ID...] 你的要求` | 无 | 选一个或多个已上传文件，继续往下处理 |

### `/doc` 怎么用

这组命令主要给群聊用。先把文件传上来，再决定要拿哪几个文件继续处理，会顺手很多。

例如：

- `/doc list`
- `/doc clear`
- `/doc clear f1`
- `/doc use f1 根据这份文件整理成正式汇报`
- `/doc use f1 f2 根据这些文件整理成正式汇报`

插件会按“平台 + 会话 + 用户”分开记文件，所以群里其他人的文件不会混进来。

如果当前会话里只有一个文件，也可以直接写：

- `/doc use 根据这份文件整理成正式汇报`

---

## 复杂 Word 工作流

这套流程更适合分步走：先建草稿，再一块一块填内容，最后锁定和导出。这样文档会稳很多，也更容易补改。

适合做这些东西：

- 管理层汇报、经营复盘
- 周报、月报
- 带标题、表格、列表、摘要卡片的报告

### 流程

1. `create_document`：建草稿，定 `theme_name`、`table_template`、`density`、`accent_color`、文档级页眉页脚
2. `add_blocks`：塞内容块进去，可以调多次
3. `finalize_document`：锁定
4. `export_document`：导出 `.docx`，自动发文件和预览图

### 当前渲染后端

复杂 Word 四步链现在只走 JS 渲染器：

- JS 渲染器：位于 `word_renderer_js/`，由 Python 在导出阶段调用 `node dist/cli.js input.json output.docx`

`create_office_file` 生成 `.docx` 时，也会走这一套 JS 渲染链。

如果本地还没构建 `word_renderer_js/dist/cli.js`，导出会直接报错，需要先构建 JS 渲染器。

### JS 渲染器构建

如果要验证或开发 JS 渲染器，需要先构建：

```powershell
cd word_renderer_js
npm install
npm run build
```

构建完成后，默认入口文件是：

```text
word_renderer_js/dist/cli.js
```

```
┌───────────────────────────┐
│     create_document       │
│                           │
│  theme_name               │
│  table_template           │
│  density                  │
│  accent_color             │
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│       add_blocks          │◄──┐
│                           │   │
│  heading    paragraph     │   │
│  list       table         │   │ 可多次调用
│  summary_card             │   │
│  page_break  section_break│   │
│  toc        group columns │   │
└─────────────┬─────────────┘   │
              │ 内容完成        │
              │─────────────────┘
              ▼
┌───────────────────────────┐
│   finalize_document       │
│       锁定草稿            │
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│    export_document        │
│  导出 .docx + 自动发送    │
└───────────────────────────┘
```

### 可用的选项

当前常用块类型包括：`hero_banner`、`heading`、`paragraph`、`list`、`table`、`summary_card`、`accent_box`、`metric_cards`、`page_break`、`section_break`、`toc`、`group`、`columns`

主题 3 套：`business_report`、`project_review`、`executive_brief`

表格样式 3 套：`report_grid`、`metrics_compact`、`minimal`

排版密度：`comfortable`（宽松）或 `compact`（紧凑）

强调色：`accent_color=RRGGBB`

卡片变体：`summary` 或 `conclusion`

页眉页脚：

- 文档级可设置常规页眉、页脚、页码
- 支持首页不同页眉页脚
- 支持奇偶页不同页眉页脚

分节：

- `section_break` 支持 `new_page`、`continuous`、`odd_page`、`even_page`、`new_column`
- 节级可覆盖页面方向、页边距、页码起始值和页眉页脚

目录：

- `toc` 支持目录标题、目录层级
- 可控制目录是否单独起页

### 注意

- 中途出错别续旧草稿，重新来一份更稳
- 想参考旧版内容，先上传旧文档让模型提取，再基于提取结果生成

> [!WARNING]
> Gemini 预览模型（如 `gemini-3-flash-preview`）在这条工具链上偶尔会返回空响应（`candidate.content.parts` 为空）。这类情况通常出在模型侧。
>
> 如果频繁遇到，换个稳定的非 Gemini 模型，或者减少单次请求的量。

> [!TIP]
> 这套工作流能让模型按步骤构建文档，减少一次硬生成整篇的失控概率。但提示写得太模糊的话，结果照样一般。
>
> 写提示的时候把这些说清楚会好很多：文档给谁看、要哪些章节、哪些内容用表格或卡片、语气偏正式还是简洁、排版偏宽松还是紧凑。

### 当前已支持

- 文档级页眉页脚
- 节级页眉页脚覆盖
- 首页不同页眉页脚
- 奇偶页不同页眉页脚
- 节级页面方向、页边距和页码重置
- 目录字段插入和目录单独起页

### 还没做的

- 合并单元格、复杂跨列布局
- 更自由的局部样式编辑
- 图文混排（图片目前不在主线范围内）

---

## 后续规划

Word 先做稳，再考虑 Excel 和 PPT。

### Word

- [ ] 合并单元格和简单的跨列表头
- [ ] 局部字体/颜色/边框自由编辑
- [ ] 脚注/尾注
- [ ] 超链接
- [ ] 页码格式（罗马数字等）

### Excel

- [ ] 独立的工作簿/工作表模型，不复用 Word 的块结构
- [ ] 多工作表、数字格式、条件格式

### PPT

- [ ] 独立的演示文稿模型
- [ ] 标题页、内容页、结论页的基础版式

---

## 支持的文件格式

### 可读取

- Office：`.docx` `.xlsx` `.pptx` `.doc` `.xls` `.ppt`
- PDF：`.pdf`
- 文本和代码：
  - `.txt` `.md` `.log` `.rst`
  - `.json` `.yaml` `.yml` `.toml` `.xml` `.csv`
  - `.html` `.css` `.sql` `.sh` `.bat`
  - `.py` `.js` `.ts` `.jsx` `.tsx` `.c` `.cpp` `.h` `.java` `.go` `.rs`

### 可生成

- Word：`.docx`
- Excel：`.xlsx`
- PPT：`.pptx`

目前只做基础内容生成，复杂排版、图表、动画还没覆盖。

### PDF 转换

| 方向 | 依赖 | 说明 |
| --- | --- | --- |
| Office → PDF | Windows: `docx2pdf` / `pywin32`（要装 Office）；或 LibreOffice | Word / Excel / PPT 都能转 |
| PDF → Word | `pdf2docx` | 文本型 PDF 效果好一些 |
| PDF → Excel | `tabula-py`（要 Java）或 `pdfplumber` | 抽表格为主，复杂版面会丢东西 |

---

## 系统依赖安装

> [!NOTE]
> Python 包基本自动装。下面说的是系统层面需要你自己装的东西。

| 平台 | 读旧 `.doc` | Office → PDF | 怎么选 |
| --- | --- | --- | --- |
| Windows | `pywin32` | Word 用 `docx2pdf`，Excel/PPT 用 `pywin32` 或 LibreOffice | 有 Office 就用 `docx2pdf` + `pywin32`，没有就装 LibreOffice |
| Linux | `antiword` | LibreOffice | 装 LibreOffice 就行 |
| macOS | `antiword` | LibreOffice | `brew install` |

> [!TIP]
> 不确定缺什么，跑一下 `/pdf_status` 就知道了。

<details>
<summary>各平台安装命令</summary>

### Windows

```bash
# 旧格式读取 + win32com 转换
pip install pywin32

# Word 转 PDF
pip install docx2pdf
```

也可以装 LibreOffice 当跨平台转换后端：<https://www.libreoffice.org/download/>

### Linux（Debian/Ubuntu）

```bash
# 读 .doc
apt install antiword

# Office -> PDF
apt install libreoffice-writer libreoffice-calc libreoffice-impress
```

### macOS

```bash
# 读 .doc
brew install antiword

# Office -> PDF
brew install --cask libreoffice
```

</details>

> [!TIP]
> 装完记得重启 AstrBot。

---

## Docker 部署

写进 Dockerfile，镜像重建也不丢：

```dockerfile
RUN apt-get update && apt-get install -y \
    antiword \
    libreoffice-writer libreoffice-calc libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*
```

> [!TIP]
> 只是临时排查环境可以进容器手装，正式部署还是写进镜像。

<details>
<summary>进容器手装（临时）</summary>

```bash
docker exec -it <容器名> bash
apt-get update
apt-get install -y antiword
apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress
```

容器删了重建得重装。

</details>

---

## 安全说明

- 默认只能访问插件工作区。
- 开了外部路径也只放开读取和转换，删除不受影响。
- 群聊默认插件关闭，按需开。
- 白名单留空 = 仅管理员可用。
- 大文件超限直接拒绝；纯文本超出分块上限会截断并提示。

---

## 常见问题

### 群聊里插件没反应？

`enable_features_in_group` 默认 `false`，打开它。

### 群聊开了，为什么还得 `@`？

`require_at_in_group` 默认 `true`，不想每次都 `@` 就关掉。

### 群聊里想先传文件，后面再说需求怎么办？

用 `/doc`。

先上传文件，再发：

- `/doc list` 看文件编号
- `/doc use f1 你的要求`

如果要一起用多个文件，可以写：

- `/doc use f1 f2 你的要求`

### 私聊看不到文件工具？

查白名单。留空的话只有管理员能用。

### `/doc list` 里为什么会看到 `f1`、`f2` 这种编号？

这是当前会话里的临时文件编号，方便你明确告诉插件要用哪个文件。编号只在这一轮会话缓存里有意义。

### `/doc clear` 说当前没有可处理的上传文件？

常见情况有两个：

- 文件已经超过 `upload_session_ttl_seconds`，自动过期了
- 你之前已经清掉过这一轮文件

### `/rm`、`/lsf` 不能用了？

换名了：`/delete_file` 和 `/list_files`。

### 没有预览图？

三种可能：`preview_settings.enable` 关了；缺 `pymupdf`；Office 预览目前只支持 Windows 上的 Word。

### PDF → Excel 出来是空表或者乱的？

PDF → Excel 是抽表格，扫描件、跨页表格、复杂版面都容易出问题。可以试试换一份更干净的 PDF 源文件，装 Java + `tabula-py`，或者先转 Word 再手动整理。

### Office → PDF 失败？

跑 `/pdf_status` 看缺什么。Windows 上 Word 要 Office + `docx2pdf`，Excel/PPT 要 Office + `pywin32`；Linux/macOS 装 LibreOffice。

### 开了外部路径还是删不了工作区外的文件？

这是故意的。外部路径只开放读取和转换，不开放删除。

### 目前有哪些不足？

- 文件生成偏基础，复杂排版、图表、动画还没做
- PDF → Excel 走表格提取，扫描件和复杂版式容易失真
- 预览图各平台表现不一样，Windows 上 Word 预览最稳
- 转换能力很依赖系统环境（LibreOffice / Java / Office 装没装齐）

---

## 交流

遇到问题或者有想法，来群里聊：

- QQ 群：`1072198212`

---

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Clhikari/astrbot_plugin_office_assistant&type=Date&theme=dark)](https://star-history.com/#Clhikari/astrbot_plugin_office_assistant&Date)

---

## 升级与迁移

从旧版本过来的话，看两件事：

1. 命令名改成 `/list_files` 和 `/delete_file` 了没。
2. 群聊默认行为现在是关的。

---

## 许可证

MIT License

## 贡献

| 项目架构图 | 贡献指南 |
| --- | --- |
| ![Project architecture diagram](./.github/images/architecture-flow.jpg) | ![Contribution flow diagram](./.github/images/contribution-flow.jpg) |

欢迎提 Issue 和 PR。涉及行为变更的话附上复现步骤和预期结果。
