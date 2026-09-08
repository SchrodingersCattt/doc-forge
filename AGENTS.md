# Agent guide for `docforge`

docforge 是项目无关的通用 Python 包：把 Markdown、LaTeX 和 AI 生成图整理为可审计的 Word 文档。本仓库是独立 git 仓库（GitHub: SchrodingersCattt/docforge），作为 submodule 挂载到使用它的研究仓库（如 pap-h2 的 `algorithm/docforge`）。

## 包边界

- 只放通用逻辑：Markdown 解析/渲染、DOCX 跟踪修订 diff、TeX→DOCX、图像生成 sidecar、source ZIP 打包、绘图样式。
- 不放项目特定内容：NSFC 官方标题/模板/variant 目录、PAP-H2 渠道策略、INVAR 稿件承诺等都留在消费方仓库。若从消费方脚本迁移代码，先剥离硬编码路径与变体常量。
- 新需求优先作为通用 API 实现；若确实属于某一项目的专属逻辑，不要进本包。

## 构建与测试

```bash
pip install -e ".[tests]"
pytest
python -m py_compile src/docforge/**/*.py   # 快速语法检查
```

源码在 `src/docforge/`，遵循 src layout；入口 `docforge` 由 pyproject `[project.scripts]` 提供。可选依赖分组：`tex` / `plot` / `aigc` / `tests`。

## 修改约定

- 保持薄 CLI：`docforge.cli` 里的子命令只做参数解析+调用库函数，逻辑在子包内。
- 保持可审计性：生成件必须带 sha256 sidecar 或验证步骤；不要静默吞掉验证异常。
- 保持类型标注 `from __future__ import annotations`，dataclass 建模（`Block`、`Span`、`PromptSpec` 等）。
- 改动后跑 `pytest`；涉及 DOCX/跟踪修订的行为变更需补测试。
- 提交信息用英语，一句话说明做了什么；默认直接在 `main` 上维护。

## 不在这里的事情

- 期刊/基金模板细节、具体项目文案、科学结论；这些应留在使用方的脚本或文档中。