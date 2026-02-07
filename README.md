# qms-monitor

用于监控药企质量管理体系运行状态的小型 Python 项目。程序从 `config.xlsx` 读取各类质量台账配置，逐个解析台账内容，统计超期与责任分布，并结合 OpenAI 兼容接口输出质量体系运行报告。

## 功能概览

- 读取配置文件 `config.xlsx`（使用项目内置 Excel 读取器）
- 按模块读取台账（变更/偏差/OOS/OOT/投诉等）
- 支持两种输入模式：
  - `excel`：直接读取 Excel（Windows + Excel COM）
  - `csv`：读取 Windows 预导出的 CSV 缓存（适合 macOS 调试）
- 自动统计：
  - 各年度总起数
  - 超期起数与占比
  - 超期清单
  - 超期按分管 QA、分管 QA 中层降序统计
- 调用 OpenAI 兼容接口进行模块分析（失败时自动回退本地统计）
- 输出 Markdown、PDF 报告与 JSON 明细

## 环境要求

- Python >= 3.12
- 运行 `excel` 模式或导出 CSV 缓存时：
  - Windows + Microsoft Excel（通过 COM 自动化读取 Excel）
  - `pywin32`（用于 COM 调用）
- 运行 `csv` 模式时：
  - 任意系统（包括 macOS）

安装依赖：

```bash
uv sync
```

## 配置文件说明

默认读取项目根目录下 `config.xlsx`，列定义如下：

- A: 序号
- B: 主题
- C: 质量模块
- D: 年份
- E: Excel 文件路径
- F: sheet 表名称
- G: 编号所在列
- H: 内容所在列
- I: 发起日期列
- J: 计划完成日期列
- K: 状态列
- L: 责任部门列
- M: 责任人列
- N: 分管 QA 列
- O: 分管 QA 中层列
- P: 未完成状态值（必填；用于状态判定）
- Q: 数据起始行（可选；默认 2）

实现规则：

- 若 `J` 列缺失或具体行计划日期为空，默认计划日期 = 发起日期 + 1 个月。
- 若路径不可读或文件读取失败，记录告警并跳过该文件，继续处理后续文件。
- 若缺失分管 QA 中层列，相关统计自动跳过。
- `未完成状态值` 列按“模块”生效，多个相同模块应保持一致。
- `数据起始行` 允许按模块单独设置（例如 2、3）；未填写时默认从第 2 行开始解析。

判定逻辑：

- 若模块配置了 `未完成状态值`：仅当状态值完全相等时视为 `open`，其他状态均视为完成。
- 若模块未配置或同模块存在冲突配置：程序会直接报错并终止执行。

## 运行方式

### 1) Windows 导出 CSV 缓存（用于 macOS 调试）

```bash
uv run python export_csv_cache.py \
  --config config.xlsx \
  --output-dir artifacts/csv_cache
```

如果提示找不到 `config.xlsx`：

- 请确认当前目录确实是项目根目录。
- Windows 资源管理器默认隐藏扩展名，实际文件可能是 `config.xlsx.xlsx`。
- 可直接用绝对路径，例如 `--config D:\\qms-monitor\\config.xlsx`。
- 脚本会打印“已尝试路径”和“相似文件”用于排查。

导出后会生成：

- `artifacts/csv_cache/rows/*.csv`
- `artifacts/csv_cache/manifest.json`

`manifest.json` 会保存每条配置（含 `open_status_value`），保证在 macOS 的 `csv` 模式下无需再读取 Excel 也能复用同样的状态判定规则。

### 2) 常规 Excel 模式（Windows）

```bash
uv run python main.py \
  --config config.xlsx \
  --input-mode excel \
  --output-dir outputs \
  --report-date 2026-02-07
```

### 3) CSV 模式（macOS / 任意系统）

```bash
uv run python main.py \
  --input-mode csv \
  --csv-manifest artifacts/csv_cache/manifest.json \
  --output-dir outputs \
  --report-date 2026-02-07
```

说明：若 `manifest.json` 含有导出时保存的完整配置，`csv` 模式可不依赖本地可读的 `config.xlsx`。

可选参数：

- `--input-mode`：`excel` 或 `csv`
- `--csv-manifest`：`csv` 模式下使用的 manifest 路径
- `--skip-llm`：跳过 LLM 调用，仅做本地统计

## LLM 配置

LLM 配置**必须通过 `.env` 文件设置**，不再支持命令行参数。

程序启动时会自动读取项目根目录下的 `.env` 文件。
你可以先复制模板：

```bash
cp .env.example .env
```

然后在 `.env` 中配置：

```bash
QMS_LLM_BASE_URL=https://api.openai.com/v1
QMS_LLM_MODEL=gpt-4o-mini
QMS_LLM_API_KEY=<YOUR_API_KEY>
QMS_LLM_TIMEOUT=120
QMS_LLM_PROGRESS_INTERVAL=15
QMS_INPUT_MODE=excel
QMS_CSV_MANIFEST=
```

支持字段：

- `QMS_LLM_BASE_URL`
- `QMS_LLM_MODEL`
- `QMS_LLM_API_KEY`
- `QMS_LLM_TIMEOUT`
- `QMS_LLM_PROGRESS_INTERVAL`
- `QMS_INPUT_MODE`
- `QMS_CSV_MANIFEST`

## 输出文件

程序在 `outputs/` 下生成：

- `qms_report_YYYYMMDD_HHMMSS.md`：质量体系运行报告
- `qms_report_YYYYMMDD_HHMMSS.pdf`：由 Markdown 报告导出的 PDF 版本
- `qms_report_YYYYMMDD_HHMMSS.json`：结构化明细（含告警）
- `qms_overdue_events_YYYYMMDD_HHMMSS.xlsx`：全部模块的超期事件汇总（单 Sheet，含“质量模块”列）

## 注意事项

- 项目内置了本地 Excel 读取模块 `qms_monitor/excel_reader.py`，不再依赖外部 Excel 读取包。
- `csv` 模式会按 `config.xlsx` 的 `row_no` 关联到 `manifest.json` 中的导出文件。
- 本仓库不在设计阶段主动读取真实质量台账文件；实际运行时按配置执行。

## PDF 排版说明（LaTeX）

程序会优先使用 `pandoc + xelatex` 将 Markdown 报告导出为排版版 PDF；若失败会自动回退到内置 reportlab 渲染。

- 需要本机安装：
  - `pandoc`
  - `xelatex`（可通过 TinyTeX / MacTeX 提供）
- 可选环境变量（`.env`）：
  - `QMS_PDF_ENGINE`：`latex`（默认）或 `reportlab`
  - `QMS_LATEX_MAINFONT`
  - `QMS_LATEX_SANSFONT`
  - `QMS_LATEX_MONOFONT`
