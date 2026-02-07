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
- 输出 Markdown 报告与 JSON 明细

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
- B: 质量模块
- C: 年份
- D: Excel 文件路径
- E: sheet 表名称
- F: 编号所在列
- G: 内容所在列
- H: 发起日期列
- I: 计划完成日期列
- J: 状态列
- K: 责任部门列
- L: 责任人列
- M: 分管 QA 列
- N: 分管 QA 中层列

实现规则：

- 若 `I` 列缺失或具体行计划日期为空，默认计划日期 = 发起日期 + 1 个月。
- 若路径不可读或文件读取失败，记录告警并跳过该文件，继续处理后续文件。
- 若缺失分管 QA 中层列，相关统计自动跳过。

## 运行方式

### 1) Windows 导出 CSV 缓存（用于 macOS 调试）

```bash
uv run python export_csv_cache.py \
  --config config.xlsx \
  --output-dir artifacts/csv_cache
```

导出后会生成：

- `artifacts/csv_cache/rows/*.csv`
- `artifacts/csv_cache/manifest.json`

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
- `--llm-base-url`：OpenAI 兼容接口地址（默认 `https://api.openai.com/v1`）
- `--llm-model`：模型名（也可通过环境变量设置）
- `--llm-api-key`：API Key（也可通过环境变量设置）
- `--llm-timeout`：请求超时秒数

## LLM 环境变量文件（推荐）

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
QMS_INPUT_MODE=excel
QMS_CSV_MANIFEST=
```

支持字段：

- `QMS_LLM_BASE_URL`
- `QMS_LLM_MODEL`
- `QMS_LLM_API_KEY`
- `QMS_LLM_TIMEOUT`
- `QMS_INPUT_MODE`
- `QMS_CSV_MANIFEST`

## 输出文件

程序在 `outputs/` 下生成：

- `qms_report_YYYYMMDD_HHMMSS.md`：质量体系运行报告
- `qms_report_YYYYMMDD_HHMMSS.json`：结构化明细（含告警）

## 注意事项

- 项目内置了本地 Excel 读取模块 `qms_monitor/excel_reader.py`，不再依赖外部 Excel 读取包。
- `csv` 模式会按 `config.xlsx` 的 `row_no` 关联到 `manifest.json` 中的导出文件。
- 本仓库不在设计阶段主动读取真实质量台账文件；实际运行时按配置执行。
