---
plan_id: template-filler-mcp-v1
generated_at: "2026-07-25"
git_commit: TBD
scope_mode: HOLD
status: draft
---

# Plan: template-filler-mcp — MCP Server Wrapper

## 背景与目标

- **问题/需求描述**：将 `template-filler-skill` 的 8 个 Python 脚本封装成标准 MCP server，使任何 MCP 兼容 Agent 都能直接调用其 extract→apply→verify 管线，而不必通过 shell 命令运行脚本。
- **目标**：产出可安装的 Python 包 `template-filler-mcp`，提供 8 个 MCP tool，与原始脚本行为完全一致。
- **非目标**：
  - 不添加新功能（如图片替换、母版文本编辑）
  - 不修改 `scripts/` 目录下的原始脚本
  - 不设置 CI/CD

## 修改方案

- **方案概述**：在现有 repo 内新增 `template_filler_mcp/` Python 包，核心逻辑从 `scripts/` 中提取重构为带类型签名的函数，通过 FastMCP `@mcp.tool()` 暴露为 MCP tools。遵循生态现有规范（setuptools + dual entry + pytest + ruff）。
- **关键设计决策**：
  - **独立包不修改 scripts/**：保持向后兼容。`scripts/` 下脚本可继续作为 CLI 独立使用。
  - **模块拆分**：按功能拆分为 `_extract.py` / `_apply.py` / `_verify.py` / `_parity.py` / `_render.py` 五个内部模块，`server.py` 统一导入。
  - **Tool 设计 1:1 映射**：每个原始脚本对应一个 MCP tool，签名保留原名语义。
  - **changes 参数**：`apply_*` 和 `verify_parity` 接受 `list[dict]` 而非文件路径——MCP 原生传参，Agent 无需写临时文件。
  - **render_preview 容错**：LibreOffice 不可用时返回 SKIPPED 状态而非报错。
- **影响范围**：
  - 新增：`template_filler_mcp/` 目录（~8 文件）
  - 新增：`pyproject.toml`（根目录）
  - 新增：`tests/` 目录
  - 新增：`.gitignore` 条目
  - 修改：`SKILL.md`（追加 MCP 使用说明）
  - 不修改：`scripts/`、`docs/`、`requirements.txt`、`README.md`

## 执行计划

### Phase 1: 项目脚手架

#### Task 1.1: 创建 Python 包结构与 pyproject.toml
- **目标**：创建可安装的 Python 包骨架
- **依赖**：无
- **执行者**：Plan Architect → Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/__init__.py`：`"""Template filler MCP server."""` + `__version__ = "0.1.0"`
  - 新建 `pyproject.toml`：遵循生态规范，包含 `[build-system]`、`[project]`、`[project.scripts]`、`[tool.setuptools.packages.find]`、`[tool.pytest.ini_options]`、`[tool.ruff]`
  - 新建 `uv.lock`（运行 `uv lock`）
  - 追加 `.gitignore`：`__pycache__/`、`*.egg-info/`、`.venv/`、`.ruff_cache/`、`.pytest_cache/`
- **修改边界**：不修改 `scripts/`、`SKILL.md`、`README.md`
- **质量检查方式**：
  - `pip install -e .` 成功
  - `python -c "import template_filler_mcp; print(template_filler_mcp.__version__)"` 输出 `0.1.0`
- **验收标准**：
  - ✅ `pip install -e .` exit 0
  - ✅ `uv lock` exit 0
  - ✅ `ruff check template_filler_mcp/` 无 error（空包）
- **潜在风险**：无

> **前置条件**：需 Python ≥3.12 和 `uv` 已安装。`uv` 未安装时用 `pip install uv`。

#### Task 1.2: 生成 uv.lock 并验证依赖解析
- **目标**：锁定依赖版本，确保可重现构建
- **依赖**：T1.1
- **执行者**：Task Executor
- **修改内容**：
  - 运行 `uv lock` 生成 `uv.lock`
  - 运行 `uv pip install -e .` 验证所有依赖可安装
- **修改边界**：仅生成 `uv.lock`，不修改其他文件
- **质量检查方式**：
  - `uv lock --check` exit 0
- **验收标准**：
  - ✅ `uv lock` exit 0
  - ✅ `uv pip install -e .` 成功安装 python-pptx, python-docx, lxml, pymupdf, mcp, pydantic, typer
- **潜在风险**：`uv` 未安装时需先 `pip install uv`；若网络受限可用 `pip-tools` 替代

### Phase 2: 核心管线模块重构

#### Task 2.1: 创建 `_extract.py` — Extractor 核心模块
- **目标**：从 `scripts/extract_pptx.py` 和 `scripts/extract_docx.py` 提取核心逻辑，重构为带类型签名的函数
- **依赖**：T1.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/_extract.py`
  - 提取公共函数 `_table_origin_cells(table)` （合并单元格去重逻辑，共享于 PPTX/DOCX）
  - 定义 `extract_pptx_content(template_path: str) -> dict[str, Any]`：整合 `walk_shapes` / `walk_table` / `walk_text_frame` 逻辑，返回 content_map dict
  - 定义 `extract_docx_content(template_path: str) -> dict[str, Any]`：整合 `emit_paragraph_runs` / `emit_table` / `emit_header_footer` 逻辑，返回 content_map dict
  - 两个函数均直接返回 dict（不写文件），保留 `paragraph_text`、`field`、`empty` 等辅助字段
- **修改边界**：不修改 `scripts/extract_pptx.py` 和 `scripts/extract_docx.py`
- **质量检查方式**：
  - 用测试 fixture `.pptx` / `.docx` 文件验证输出结构与原始脚本一致
  - 对比 `extract_pptx_content(fixture.pptx)` 输出与 `python scripts/extract_pptx.py fixture.pptx /tmp/map.json && cat /tmp/map.json` 输出
- **验收标准**：
  - ✅ `extract_pptx_content()` 返回的 dict 包含 `format`, `source`, `slide_count`, `slides` 字段
  - ✅ `extract_docx_content()` 返回的 dict 包含 `format`, `source`, `section_count`, `body_runs`, `header_footer_runs` 字段
  - ✅ 与原始脚本输出 JSON 逐字段一致（slide 数量、run 数量、每个 run 的 text 值）
- **潜在风险**：python-pptx 版本差异导致 shape 遍历顺序不同——使用与 `requirements.txt` 相同的最低版本约束

#### Task 2.2: 创建 `_apply.py` — Applicator 核心模块
- **目标**：从 `scripts/apply_pptx.py` 和 `scripts/apply_docx.py` 提取核心逻辑
- **依赖**：T1.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/_apply.py`
  - 定义 `apply_pptx_changes(template_path: str, changes: list[dict], output_path: str) -> dict[str, Any]`
  - 定义 `apply_docx_changes(template_path: str, changes: list[dict], output_path: str) -> dict[str, Any]`
  - 返回 `{"applied": N, "total": M, "failed": [...], "output_path": "..."}`
  - 保留 paragraph-level 回退（ID 无尾部 run 索引时替换整个段落）
  - 保留分组形状穿透（`find_shape_by_path`）
  - 注意：若多个 changes 条目指向同一 run ID，最后一条生效（覆盖式）
- **修改边界**：不修改 `scripts/apply_pptx.py` 和 `scripts/apply_docx.py`
- **质量检查方式**：
  - 对 fixture 文件：extract → 构造 changes → apply → 重新 extract，验证 changes 中的 run text 已更新
- **验收标准**：
  - ✅ `apply_pptx_changes()` 修改指定 run 后，未被 changes 涉及的 run text 不变
  - ✅ `apply_docx_changes()` 同样行为
  - ✅ paragraph-level 替换（如 `"2/5/0"`）正确工作：首 run 获新文本，其余 run 清空
- **潜在风险**：changes 中 ID 不存在时应 warn 而非 crash——参考原始脚本的 try/except 模式

#### Task 2.3: 创建 `_verify.py` 与 `_parity.py` — Verifier 核心模块
- **目标**：从 `scripts/verify_pptx.py`、`scripts/verify_docx.py`、`scripts/verify_parity.py` 提取核心逻辑
- **依赖**：T2.1, T2.2
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/_verify.py`
    - `verify_pptx_structure(filepath: str) -> dict[str, Any]`：返回 `{"ok": bool, "problems": [...]}`
    - `verify_docx_structure(filepath: str) -> dict[str, Any]`：同上
  - 新建 `template_filler_mcp/_parity.py`
    - `verify_parity(original_path: str, output_path: str, changes: list[dict]) -> dict[str, Any]`
    - 直接调用 `_extract.extract_pptx_content()` / `_extract.extract_docx_content()` 而非 subprocess
    - 返回 `{"ok": bool, "problems": [...]}`
- **修改边界**：不修改 `scripts/verify_*.py`
- **质量检查方式**：
  - 对有效文件运行 verify，预期 `ok: true`
  - 构造一个损坏的 PPTX（如 ZIP 中删除一个 XML 文件），预期 `ok: false`
  - 构造 changes 验证 parity：apply 后 verify_parity 应 PASS
- **验收标准**：
  - ✅ `verify_pptx_structure()` 对有效 .pptx 返回 `{"ok": True, "problems": []}`
  - ✅ `verify_parity()` 对仅修改 changes 中 run 的文件返回 `{"ok": True}`
  - ✅ `verify_parity()` 对多修改了额外 run 的文件返回 `{"ok": False}` 并列出差异
- **潜在风险**：verify_parity 中 `covered` 集合的 paragraph-level 匹配逻辑需与 `_apply.py` 保持一致

#### Task 2.4: 创建 `_render.py` — Renderer 核心模块
- **目标**：从 `scripts/render_preview.py` 提取核心逻辑
- **依赖**：T1.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/_render.py`
  - `render_preview_pages(filepath: str, output_dir: str, pages: list[int] | None = None) -> dict[str, Any]`
  - 返回 `{"ok": bool, "rendered": [...], "skipped_reason": str | None}`
  - LibreOffice 不可用时返回 `{"ok": False, "skipped_reason": "soffice not found"}`
- **修改边界**：不修改 `scripts/render_preview.py`
- **质量检查方式**：
  - 在有 LibreOffice 的环境中运行，验证输出 PNG 文件存在
  - 在无 LibreOffice 的环境中运行，验证返回 `skipped_reason`
- **验收标准**：
  - ✅ LibreOffice 可用：返回 `ok: true`，`rendered` 列表非空
  - ✅ LibreOffice 不可用：返回 `ok: false`，`skipped_reason` 非空，不抛异常
- **潜在风险**：soffice headless 转换大文件可能超时——设置 120s timeout

### Phase 3: MCP Server 与 CLI

#### Task 3.1: 创建 `server.py` — FastMCP Server
- **目标**：将 5 个核心模块的 8 个函数通过 `@mcp.tool()` 暴露为 MCP tools
- **依赖**：T2.1, T2.2, T2.3, T2.4
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/server.py`
  - 初始化 `mcp = FastMCP("template-filler", instructions="Fill existing PPTX/DOCX templates with new content while preserving formatting...")`
  - 注册 8 个 tool：
    1. `extract_pptx(template_path: str) -> dict[str, Any]`
    2. `extract_docx(template_path: str) -> dict[str, Any]`
    3. `apply_pptx(template_path: str, changes: list[dict], output_path: str) -> dict[str, Any]`
    4. `apply_docx(template_path: str, changes: list[dict], output_path: str) -> dict[str, Any]`
    5. `verify_pptx(filepath: str) -> dict[str, Any]`
    6. `verify_docx(filepath: str) -> dict[str, Any]`
    7. `verify_parity(original_path: str, output_path: str, changes: list[dict]) -> dict[str, Any]`
    8. `render_preview(filepath: str, output_dir: str, pages: list[int] | None = None) -> dict[str, Any]`
  - 定义 `def main() -> None: mcp.run()`
- **修改边界**：不修改 `scripts/`，不修改 `template_filler_mcp/` 下其他模块
- **质量检查方式**：
  - `python -c "from template_filler_mcp.server import mcp; print(mcp.name)"` 输出 `template-filler`
  - `python -c "from template_filler_mcp.server import mcp; print(len(mcp._tool_manager._tools))"` 输出 `8`
- **验收标准**：
  - ✅ `template-filler-mcp` 命令可启动（`python -m template_filler_mcp.server`）
  - ✅ `mcp._tool_manager._tools` 包含 8 个已注册 tool
  - ✅ 每个 tool 的 docstring 包含功能说明和参数文档
- **潜在风险**：FastMCP tool 注册的 `list[dict]` 参数类型序列化——需验证 MCP stdio transport 对嵌套 dict 类型的支持

#### Task 3.2: 创建 `cli.py` — Typer CLI
- **目标**：提供独立命令行入口，方便非 MCP 场景使用
- **依赖**：T2.1, T2.2, T2.3, T2.4
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `template_filler_mcp/cli.py`
  - 使用 Typer 暴露 5 个子命令：`extract`、`apply`、`verify`、`parity`、`render`
  - 每个子命令映射到对应的核心模块函数
  - `apply` 和 `parity` 子命令的 `--changes` 参数接受 JSON 文件路径（兼容原 workflow）
- **修改边界**：不修改 `server.py`，不修改 `scripts/`
- **质量检查方式**：
  - `template-filler extract pptx fixture.pptx` 输出 content_map JSON
  - `template-filler verify pptx fixture.pptx` 输出结构验证结果
- **验收标准**：
  - ✅ `template-filler --help` 显示 5 个子命令
  - ✅ `template-filler extract pptx <path>` exit 0
  - ✅ `template-filler verify pptx <path>` exit 0 对有效文件
- **潜在风险**：CLI 的 JSON 输出与原始脚本输出格式需一致——用户可能已有脚本依赖

### Phase 4: 测试

#### Task 4.1: 创建测试 fixtures
- **目标**：生成最小但合法的 PPTX 和 DOCX 测试文件
- **依赖**：T1.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/__init__.py`（空文件）
  - 新建 `tests/fixtures/` 目录
  - 创建 `tests/fixtures/minimal.pptx`：1 slide，含 1 个文本框（"Hello World"）、1 个表格（2×2）。使用 python-pptx 编程创建（`from pptx import Presentation; prs = Presentation(); ...`），非手工制作
  - 创建 `tests/fixtures/minimal.docx`：1 段落（"Hello DOCX"）、1 个表格（2×2）、默认页眉。使用 python-docx 编程创建
- **修改边界**：fixture 文件必须合法可被 PowerPoint/Word 正常打开
- **质量检查方式**：
  - `python -c "from pptx import Presentation; Presentation('tests/fixtures/minimal.pptx')"` 不报错
  - `python -c "from docx import Document; Document('tests/fixtures/minimal.docx')"` 不报错
- **验收标准**：
  - ✅ `verify_pptx_structure('tests/fixtures/minimal.pptx')` 返回 `ok: true`
  - ✅ `extract_pptx_content('tests/fixtures/minimal.pptx')['slides'][0]['runs']` 非空
- **潜在风险**：生成的 fixture 可能过于简单，无法覆盖合并单元格等边界场景——Phase 4.5 的 parity 测试会覆蓋

#### Task 4.2: `test_extract.py` — Extractor 测试
- **目标**：验证 extract 函数输出格式和内容正确
- **依赖**：T2.1, T4.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/test_extract.py`
  - 测试用例：
    1. `test_extract_pptx_returns_valid_structure`：验证返回 dict 的顶层 key
    2. `test_extract_pptx_run_ids_match_expected`：验证 ID 格式 `{slide}/{shape_id}/{para}/{run}`
    3. `test_extract_docx_returns_body_and_headers`：验证 `body_runs` 和 `header_footer_runs`
    4. `test_extract_pptx_matches_legacy_script`：对比与原始脚本输出一致
    5. `test_extract_docx_paragraph_text_field`：验证 `paragraph_text` 拼接完整
- **修改边界**：只读测试，不修改 fixture 文件
- **质量检查方式**：`pytest tests/test_extract.py -v`
- **验收标准**：✅ 全部 5 条测试通过

#### Task 4.3: `test_apply.py` — Applicator 测试
- **目标**：验证 apply 函数正确修改指定 run，不触及其他 run
- **依赖**：T2.2, T4.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/test_apply.py`
  - 测试用例：
    1. `test_apply_single_run_change`：修改一个 run，验证该 run 文本已更新
    2. `test_apply_paragraph_level_change`：使用段落级 ID，验证整个段落替换
    3. `test_apply_untouched_runs_preserved`：确认其他 run 文本不变
    4. `test_apply_invalid_id_warns`：无效 ID 不抛异常但返回失败信息
    5. `test_apply_table_cell_change`：修改表格单元格文本
- **修改边界**：每次测试用 fixture 副本，不修改原 fixture
- **质量检查方式**：`pytest tests/test_apply.py -v`
- **验收标准**：✅ 全部 5 条测试通过

#### Task 4.4: `test_verify.py` — Verifier 测试
- **目标**：验证结构校驗正确检测问题
- **依赖**：T2.3, T4.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/test_verify.py`
  - 测试用例：
    1. `test_verify_pptx_valid`：有效文件返回 ok
    2. `test_verify_docx_valid`：同上
    3. `test_verify_pptx_corrupt`：损坏 ZIP 返回 ok=false
    4. `test_verify_docx_corrupt`：同上
- **修改边界**：损坏文件在测试中动态构造，不预存
- **质量检查方式**：`pytest tests/test_verify.py -v`
- **验收标准**：✅ 全部 4 条测试通过

#### Task 4.5: `test_parity.py` 与 `test_render.py`
- **目标**：验证奇偶校验和预览渲染
- **依赖**：T2.3, T2.4, T4.1
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/test_parity.py`
    - `test_parity_clean_apply`：apply 后 parity 应 PASS
    - `test_parity_detects_unexpected_change`：手动修改额外 run 后 parity 应 FAIL
    - `test_parity_detects_missing_run`：删除 run 后 parity 应 FAIL
  - 新建 `tests/test_render.py`
    - `test_render_pptx_creates_png`：在有 soffice 的环境中验证 PNG 输出
    - `test_render_no_soffice_graceful`：无 soffice 时返回 skipped_reason
- **修改边界**：同 4.3
- **质量检查方式**：`pytest tests/test_parity.py tests/test_render.py -v`
- **验收标准**：
  - ✅ parity 全部 3 条测试通过
  - ✅ render 测试：有 soffice 时通过，无 soffice 时 SKIP

#### Task 4.6: `test_server.py` — MCP Server 集成测试
- **目标**：验证 MCP server 可启动且所有 tool 可通过函数调用访问
- **依赖**：T3.1, T4.1–T4.5
- **执行者**：Task Executor
- **修改内容**：
  - 新建 `tests/test_server.py`
  - 测试用例：
    1. `test_server_name`：mcp.name == "template-filler"
    2. `test_all_tools_registered`：8 个 tool 已注册
    3. `test_each_tool_callable`：每个 tool 函数可调用（用 fixture 文件）
- **修改边界**：测试调用 tool 函数本身（非 MCP transport），复用核心模块逻辑
- **质量检查方式**：`pytest tests/test_server.py -v`
- **验收标准**：✅ 全部 3 条测试通过

### Phase 5: 文档与收尾

#### Task 5.1: 更新 SKILL.md
- **目标**：在现有 skill 文档中追加 MCP 使用模式
- **依赖**：T3.1
- **执行者**：Task Executor
- **修改内容**：
  - 在 `SKILL.md` 末尾追加 "## MCP Server Usage" 节
  - 说明：安装方式、MCP 配置示例（`mcpServers` JSON）、tool 列表、与原始脚本工作流的对应关系
  - 注明 `render_preview` 依赖 LibreOffice
- **修改边界**：仅追加，不修改现有 SKILL.md 内容
- **质量检查方式**：人工阅读确认
- **验收标准**：
  - ✅ SKILL.md 存在 "MCP Server Usage" 节
  - ✅ 包含完整的 mcpServers 配置示例
- **潜在风险**：无

#### Task 5.2: 端到端验证
- **目标**：模拟真实工作流完整验证 MCP server
- **依赖**：T4.1–T4.6
- **执行者**：Task Executor
- **修改内容**：
  - 运行完整测试套件：`pytest tests/ -v`
  - 运行 ruff 检查：`ruff check template_filler_mcp/ tests/`
  - 验证 scripts 向后兼容：`python scripts/extract_pptx.py tests/fixtures/minimal.pptx /tmp/test.json`
  - 端到端流程：
    1. extract_pptx → 读取 content_map
    2. 构造 changes（修改 fixture 中的 "Hello World" → "Hello MCP"）
    3. apply_pptx → 写入新文件
    4. verify_pptx → 结构验证 PASS
    5. verify_parity → 奇偶校验 PASS
    6. extract_pptx(output) → 确认 run text 已更新
- **修改边界**：不修改任何文件，全流程只读+临时输出
- **质量检查方式**：所有命令 exit 0
- **验收标准**：
  - ✅ `pytest tests/ -v` 全部通过
  - ✅ `ruff check` 0 errors
  - ✅ `scripts/extract_pptx.py` 仍可独立工作
  - ✅ 端到端流程 6 步全部 PASS

## Execution Wave（并行执行波次）

| Wave | 可并行 Task | 依赖已完成 |
|------|------------|------------|
| W1 | T1.1, T1.2 | — |
| W2 | T2.1, T2.2, T2.3, T2.4 | W1 |
| W3 | T3.1, T3.2, T4.1 | W2 (T2.x) |
| W4 | T4.2, T4.3, T4.4, T4.5 | W3 (T4.1) |
| W5 | T4.6, T5.1 | W3 (T3.1), W4 |
| W6 | T5.2 | W5 |

## Post-Execution Verification

### Automated Verification（Task Executor 自动执行）

| ID | Description | Command | Expected |
|----|-------------|---------|----------|
| V1 | 包可安装 | `cd /home/gw/opt/template-filler-skill && pip install -e .` | exit 0 |
| V2 | 导入无错误 | `python -c "from template_filler_mcp.server import mcp; print(mcp.name)"` | stdout: `template-filler` |
| V3 | 全量测试 | `pytest tests/ -v` | exit 0, 所有测试 PASS |
| V4 | ruff 无错误 | `ruff check template_filler_mcp/ tests/` | exit 0 |
| V5 | 原始脚本向后兼容 | `python scripts/extract_pptx.py tests/fixtures/minimal.pptx /tmp/test_legacy.json` | exit 0 |

### Probe (best-effort, run if available)

| ID | Description | Command | Expected |
|----|-------------|---------|----------|
| P1 | render_preview 渲染测试 | `python -c "from template_filler_mcp._render import render_preview_pages; r = render_preview_pages('tests/fixtures/minimal.pptx', '/tmp/preview_test'); print(r['ok'])"` | `True` 或 `False`（取决于 LibreOffice 可用性） |

### Manual（真正需要人工判断）

| ID | Description |
|----|-------------|
| M1 | 人工阅读 SKILL.md 的 "MCP Server Usage" 节，确认内容完整准确 |
| M2 | 在 MCP client（如 Claude Desktop）中配置 `template-filler-mcp`，实际调用 1-2 个 tool 验证端到端可用 |

## 审查日志

| 轮次 | 聚焦 | 发现问题数 | 已修正 | 剩余 |
|------|------|-----------|--------|------|
| R1 | 结构完整性 | 0 | 0 | 0 |
| R1.5 | 外部引用事实核查 | 0 | 0 | 0 |
| R2 | 可执行性（含依赖验证） | 1 (T4.1 fixture 创建方法不明确) | 1 | 0 |
| R2.8 | LLM 可执行性审查 | 1 (T1.2 缺失 uv 前置说明) | 1 | 0 |
| R3 | 风险与边缘 | 2 (changes 重复 ID 行为、前置条件未标注) | 2 | 0 |
| **终止** | **R3 PASS — 全部 issue 清零** | | | **0** |

## Execution Log

（Task Executor 执行后回写）
