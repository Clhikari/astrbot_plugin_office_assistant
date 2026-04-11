# 贡献指南

欢迎提交 Bug 报告、功能建议和代码 PR。

## 目录

- [快速开始](#快速开始)
- [项目结构](#项目结构)
- [开发环境](#开发环境)
- [提交规范](#提交规范)
- [Pull Request 流程](#pull-request-流程)
- [代码风格](#代码风格)
- [测试](#测试)
- [Issue 指南](#issue-指南)
- [许可证](#许可证)

---

## 快速开始

```bash
# 1. Fork 并克隆仓库
git clone https://github.com/<your-username>/astrbot_plugin_office_assistant.git
cd astrbot_plugin_office_assistant

# 2. 安装 Python 依赖
pip install -r requirements.txt

# 3. 如果修改了 word_renderer_js，再安装并构建 Node 侧依赖
cd word_renderer_js
npm install
npm run build
cd ..

# 4. 运行测试
pytest
```

## 项目结构

```text
astrbot_plugin_office_assistant/
├── main.py                  # 插件入口
├── app/                     # 运行时对象与配置
│   ├── runtime.py
│   └── settings.py
├── services/                # 服务层
│   ├── runtime_builder.py
│   ├── command_service.py
│   ├── llm_request_policy.py
│   ├── request_hook_service.py
│   ├── prompt_context_service.py
│   ├── upload_session_service.py
│   ├── incoming_message_service.py
│   ├── file_tool_service.py
│   └── ...
├── message_buffer.py        # 上传文件缓冲
├── agent_tools/             # AstrBot 工具暴露层
├── mcp_server/              # MCP 暴露层
├── tools/                   # 共享工具注册与适配
│   ├── registry.py
│   ├── astrbot_adapter.py
│   └── mcp_adapter.py
├── domain/document/         # 文档契约、session store、导出编排、渲染后端配置
├── document_core/           # 文档块模型、宏和 Word builder
├── prompts/                 # 提示词模板
├── word_renderer_js/        # TypeScript Word 渲染器
├── tests/                   # 测试套件
└── .github/                 # CI、Issue 模板等
```

## 开发环境

### 必需

| 工具 | 版本要求 |
|------|----------|
| Python | >= 3.10 |
| Node.js | >= 18（仅修改 `word_renderer_js` 时需要） |

### Python 依赖

```bash
pip install -r requirements.txt
```

常用依赖包括 `python-docx`、`openpyxl`、`python-pptx`。完整列表见 [requirements.txt](requirements.txt)。

### Word 渲染器（TypeScript）

如果你的改动涉及 `word_renderer_js/`：

```bash
cd word_renderer_js
npm install
npm run build
```

修改 `.ts` 文件后要重新执行 `npm run build`，否则 Python 端调用的还是旧的 `dist/cli.js`。

## 提交规范

使用 [Conventional Commits](https://www.conventionalcommits.org/) 格式：

```text
<type>(<scope>): <description>

[optional body]
```

### 常用 type

| Type | 用途 |
|------|------|
| `feat` | 新功能 |
| `fix` | Bug 修复 |
| `refactor` | 重构，不改变外部行为 |
| `perf` | 性能优化 |
| `test` | 测试相关 |
| `docs` | 文档 |
| `chore` | 构建、CI、依赖等杂项 |

### 常用 scope

`services`、`domain`、`renderer`、`agent-tools`、`mcp`、`docs`、`tests`

### 示例

```text
feat(renderer): add section_break support for landscape pages
fix(services): stabilize document follow-up notice parsing
docs(readme): replace static architecture image with mermaid diagram
```

## Pull Request 流程

1. 从最新的 `master` 创建分支

   ```bash
   git checkout -b feat/your-feature master
   ```

2. 一个 commit 尽量只做一类改动

3. 跑相关测试

   ```bash
   pytest
   ```

4. 推送并创建 PR
   如果仓库里启用了 PR 模板，就按模板填写；如果没有，就把变更内容、验证方式和影响范围写清楚

5. 跟进 review 评论，更新代码或说明

### PR 检查清单

- [ ] 分支基于最新的 `master`
- [ ] 相关测试通过
- [ ] 新功能或修复有对应测试
- [ ] 修改 TypeScript 后重新执行了 `npm run build`
- [ ] 涉及 E2E 能力或端上可见行为的改动，已经在实际端上流程中验证通过
- [ ] 共享逻辑继续放在 `services/`、`tools/registry.py`、`domain/document/`，没有在适配层重复实现
- [ ] 提交信息符合 Conventional Commits
- [ ] 没有引入不必要的依赖

## 代码风格

### Python

- 遵循 PEP 8
- 使用 type hints
- 函数参数优先使用 keyword-only（`*`）
- 保持现有注释和 docstring 风格
- 日志统一使用 `logger.info/debug/error("[模块名] 消息", ...)` 这种形式，优先 `%s` 占位

### TypeScript

- 使用 `strict` 模式
- 优先使用 `const`
- 避免 `any`
- 渲染器导出的函数签名保持一致

### 通用

- 不要留下 `TODO`、临时调试输出或无用注释
- 新增配置项要同步更新 `_conf_schema.json`
- 修改工具 schema 时，同步更新对应测试
- `main.py` 保持薄入口，优先把逻辑放在合适的 service 或 domain 层
- `agent_tools/` 和 `mcp_server/` 负责暴露能力，公共行为优先收敛到 `tools/registry.py`

## 测试

```bash
# 运行全部测试
pytest

# 运行单个测试文件
pytest tests/test_office_assistant_services.py

# 带详细输出
pytest -v

# 按关键字筛选
pytest -k "upload_session"
```

### 主要测试文件

| 文件 | 覆盖范围 |
|------|----------|
| `test_office_assistant_services.py` | 服务层，包括 request hook、prompt context、upload session、command service 等 |
| `test_office_assistant_before_llm_chat.py` | LLM 请求策略 |
| `test_office_assistant_agent_tools.py` | agent tool 定义与调用 |
| `test_office_assistant_contracts.py` | 文档契约与 schema |
| `test_office_assistant_session_store.py` | 文档 session store 与块归一化 |
| `test_office_assistant_node_renderer.py` | Node 渲染器与复杂 Word 导出 |
| `test_office_assistant_office_generator.py` | Office 文件生成 |

### 测试建议

- 改 `request_hook_service.py`、`prompt_context_service.py`、`llm_request_policy.py` 时，优先补 `test_office_assistant_services.py` 或 `test_office_assistant_before_llm_chat.py`
- 改 `domain/document/` 下的 contract、session store、export pipeline 时，至少覆盖 `test_office_assistant_contracts.py` 或 `test_office_assistant_session_store.py`
- 改 `word_renderer_js/`、`render_backends.py` 或复杂 Word 导出链时，记得跑 `test_office_assistant_node_renderer.py`
- 改动如果会影响真实对话流程、上传文件流程、导出回传、命令交互或其他 E2E 行为，除了自动化测试，还要在实际端上完整走一遍并确认结果正确
- 辅助函数放在 `conftest.py` 或 `_docx_test_helpers.py`

## 许可证

本项目使用 [AGPL-3.0](LICENSE) 许可证。提交贡献即表示你同意以相同许可证发布你的代码。
