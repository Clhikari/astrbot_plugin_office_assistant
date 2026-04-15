# 贡献指南

欢迎参与这个项目。下面是开发环境搭建和提交代码的基本流程。

## 开发环境

1. 克隆仓库，确保 Python 3.12+ 和 Node 18+ 可用
2. 安装 Python 依赖：`pip install -e .` 或让 AstrBot 自动安装
3. 构建 Word JS 渲染器：

```bash
cd word_renderer_js
npm install
npm run build
```

## 跑测试

项目用 pytest，测试文件在 `tests/` 目录下：

```bash
pytest tests/ -v
```

测试用的临时文件会自动清理，不用手动管。

## 提交 PR

1. 从 `master` 分支拉新分支
2. 改完代码跑一遍测试，确认没挂
3. 提 PR 时写清楚改了什么、为什么改。涉及行为变更的话附
上复现步骤和预期结果
4. 如果不知道pr怎么写，可参考项目.github/pull_request_template.md里的模板

commit message 没有强制格式，说清楚就行。

## 项目架构

架构图和模块关系见 [docs/architecture.md](docs/architecture.md)。

几个关键路径：

- `main.py` — 插件入口，注册工具和命令
- `services/` — 请求管道、权限、文件处理等服务
- `domain/document/` — Word 文档领域模型和导出管道
- `tools/` — 工具注册和适配器
- `word_renderer_js/` — JS 渲染器，负责生成 .docx
- `tests/` — pytest 测试

## 报 Bug

开 Issue，写上：

- 你做了什么操作
- 期望的结果
- 实际的结果
- 跑 `/fileinfo` 和 `/pdf_status` 的输出（如果相关的话）
