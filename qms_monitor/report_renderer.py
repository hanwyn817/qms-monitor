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
    topic_results: dict[str, dict[str, Any]],
    warnings: list[str],
    processed_files: int,
    skipped_files: int,
) -> str:
    lines: list[str] = []
    lines.append("")
    lines.append(f"- 报告日期: {report_date.isoformat()}")
    lines.append(f"- 配置文件: {config_path}")
    lines.append(f"- 成功读取台账: {processed_files}")
    lines.append(f"- 跳过台账: {skipped_files}")
    lines.append("")

    if warnings:
        lines.append("# 处理告警")
        for warning in warnings:
            lines.append(f"- {warning}")
        lines.append("")

    for topic in sorted(topic_results.keys()):
        item = topic_results[topic]
        lines.append(f"# 主题：{topic}")

        yearly_totals = item.get("yearly_totals", [])

        yearly_overdue = item.get("yearly_overdue", [])
        lines.append("## 各年度超期情况")
        if yearly_overdue:
            lines.append("| 年份 | 起数 | 超期起数 | 超期占比 |")
            lines.append("|---|---:|---:|---:|")
            for row in yearly_overdue:
                lines.append(
                    f"| {row.get('year', '')} | {row.get('count', 0)} | {row.get('overdue_count', 0)} | {row.get('overdue_ratio', 0)}% |"
                )
        else:
            lines.append("无可统计数据")
        lines.append("")

        total = item.get("total", {})
        total_count = total.get("count")
        if total_count is None:
            total_count = sum(int(r.get("count", 0) or 0) for r in yearly_totals)
        lines.append("## 总起数和超期情况")
        lines.append(f"- 总起数: {total_count}")

        overdue = item.get("overdue", {})
        lines.append(f"- 总超期起数: {overdue.get('count', 0)}")
        lines.append(f"- 总超期占比: {overdue.get('ratio', 0)}%")
        lines.append("")

        lines.append("## 超期按分管QA统计（降序）")
        qa_rank = item.get("overdue_by_qa", [])
        if qa_rank:
            lines.append("| 分管QA | 起数 |")
            lines.append("|---|---:|")
            for row in qa_rank:
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} |")
        else:
            lines.append("无可统计数据")
        lines.append("")

        lines.append("## 超期按分管QA中层统计（降序）")
        qa_manager_rank = item.get("overdue_by_qa_manager", [])
        if qa_manager_rank:
            lines.append("| 分管QA中层 | 起数 |")
            lines.append("|---|---:|")
            for row in qa_manager_rank:
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} |")
        else:
            lines.append("无可统计数据（可能配置中缺失分管QA中层列）")
        lines.append("")

        lines.append("## 总结")
        summary = (item.get("summary") or "").strip()
        lines.append(summary if summary else "无")
        lines.append("")

    return "\n".join(lines).strip() + "\n"
