from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any


def safe_md_cell(value: Any) -> str:
    s = str(value or "")
    s = s.replace("|", "\\|").replace("\n", " ").strip()
    return s


def render_markdown_report(
    report_date: date,
    config_path: Path,
    module_results: dict[str, dict[str, Any]],
    warnings: list[str],
    processed_files: int,
    skipped_files: int,
) -> str:
    lines: list[str] = []
    lines.append("# 质量体系运行报告")
    lines.append("")
    lines.append(f"- 报告日期: {report_date.isoformat()}")
    lines.append(f"- 配置文件: {config_path}")
    lines.append(f"- 成功读取台账: {processed_files}")
    lines.append(f"- 跳过台账: {skipped_files}")
    lines.append("")

    if warnings:
        lines.append("## 处理告警")
        for warning in warnings:
            lines.append(f"- {warning}")
        lines.append("")

    for module in sorted(module_results.keys()):
        item = module_results[module]
        lines.append(f"## 模块：{module}")

        yearly_totals = item.get("yearly_totals", [])
        lines.append("### 各年度总起数")
        if yearly_totals:
            lines.append("| 年份 | 起数 |")
            lines.append("|---|---:|")
            for row in yearly_totals:
                lines.append(f"| {row.get('year', '')} | {row.get('count', 0)} |")
        else:
            lines.append("无数据")
        lines.append("")

        overdue = item.get("overdue", {})
        lines.append("### 超期情况")
        lines.append(f"- 超期起数: {overdue.get('count', 0)}")
        lines.append(f"- 超期占比: {overdue.get('ratio', 0)}%")
        lines.append("")

        overdue_items = overdue.get("items", [])
        if overdue_items:
            lines.append("#### 超期清单")
            lines.append("| 年份 | 编号 | 内容 | 计划完成日期 | 状态 | 分管QA | 分管QA中层 | 来源 |")
            lines.append("|---|---|---|---|---|---|---|---|")
            for row in overdue_items:
                lines.append(
                    "| {year} | {event_id} | {content} | {planned_date} | {status} | {qa} | {qa_manager} | {source} |".format(
                        year=row.get("year", ""),
                        event_id=safe_md_cell(row.get("event_id", "")),
                        content=safe_md_cell(row.get("content", "")),
                        planned_date=row.get("planned_date", ""),
                        status=safe_md_cell(row.get("status", "")),
                        qa=safe_md_cell(row.get("qa", "")),
                        qa_manager=safe_md_cell(row.get("qa_manager", "")),
                        source=safe_md_cell(row.get("source", "")),
                    )
                )
        else:
            lines.append("#### 超期清单")
            lines.append("无超期记录")
        lines.append("")

        lines.append("### 超期按分管QA统计（降序）")
        qa_rank = item.get("overdue_by_qa", [])
        if qa_rank:
            lines.append("| 分管QA | 起数 |")
            lines.append("|---|---:|")
            for row in qa_rank:
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} |")
        else:
            lines.append("无可统计数据")
        lines.append("")

        lines.append("### 超期按分管QA中层统计（降序）")
        qa_manager_rank = item.get("overdue_by_qa_manager", [])
        if qa_manager_rank:
            lines.append("| 分管QA中层 | 起数 |")
            lines.append("|---|---:|")
            for row in qa_manager_rank:
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} |")
        else:
            lines.append("无可统计数据（可能配置中缺失分管QA中层列）")
        lines.append("")

        summary = (item.get("summary") or "").strip()
        if summary:
            lines.append("### LLM总结")
            lines.append(summary)
            lines.append("")

    return "\n".join(lines).strip() + "\n"
