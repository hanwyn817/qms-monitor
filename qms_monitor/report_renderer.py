from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

MAX_RANK_TABLE_ROWS = 15


def safe_md_cell(value: Any) -> str:
    s = str(value or "")
    s = s.replace("|", "\\|").replace("\n", " ").strip()
    return s


def _chunk_text_for_table(text: str, max_len: int = 24) -> list[str]:
    chunks: list[str] = []
    buf = ""
    punctuation = set("。！？；;!?")

    for ch in text:
        buf += ch
        if (ch in punctuation and len(buf) >= max(8, max_len // 2)) or len(buf) >= max_len:
            part = buf.strip()
            if part:
                chunks.append(part)
            buf = ""

    tail = buf.strip()
    if tail:
        chunks.append(tail)
    return chunks


def format_summary_cell(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "-"

    lines: list[str] = []
    for line in raw.splitlines():
        normalized = line.strip()
        if not normalized:
            continue
        lines.extend(_chunk_text_for_table(normalized))

    if not lines:
        return "-"
    return "<br/>".join(safe_md_cell(line) for line in lines)


def split_rank_rows(rows: list[dict[str, Any]], max_rows: int = MAX_RANK_TABLE_ROWS) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if len(rows) <= max_rows:
        return rows, []
    return rows[:max_rows], rows[max_rows:]


def format_overflow_note(rows: list[dict[str, Any]], label: str) -> str:
    if not rows:
        return ""
    parts: list[str] = []
    for row in rows:
        name = str(row.get("name", "") or "").strip()
        if not name:
            continue
        parts.append(f"{name}({row.get('count', 0)})")
    if not parts:
        return ""
    return f"其他涉及{label}有：{'、'.join(parts)}。"


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
            qa_table_rows, qa_overflow_rows = split_rank_rows(qa_rank)
            lines.append("| 分管QA | 起数 | 超期内容概括 |")
            lines.append("|---|---:|---|")
            for row in qa_table_rows:
                summary = format_summary_cell(row.get("summary", ""))
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} | {summary} |")
            overflow_note = format_overflow_note(qa_overflow_rows, "人员")
            if overflow_note:
                lines.append("")
                lines.append(overflow_note)
        else:
            lines.append("无可统计数据")
        lines.append("")

        lines.append("## 超期按分管QA中层统计（降序）")
        qa_manager_rank = item.get("overdue_by_qa_manager", [])
        if qa_manager_rank:
            qa_manager_table_rows, qa_manager_overflow_rows = split_rank_rows(qa_manager_rank)
            lines.append("| 分管QA中层 | 起数 | 超期内容概括 |")
            lines.append("|---|---:|---|")
            for row in qa_manager_table_rows:
                summary = format_summary_cell(row.get("summary", ""))
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} | {summary} |")
            overflow_note = format_overflow_note(qa_manager_overflow_rows, "人员")
            if overflow_note:
                lines.append("")
                lines.append(overflow_note)
        else:
            lines.append("无可统计数据（可能配置中缺失分管QA中层列）")
        lines.append("")

        lines.append("## 超期按责任部门统计（降序）")
        owner_dept_rank = item.get("overdue_by_owner_dept", [])
        if owner_dept_rank:
            owner_dept_table_rows, owner_dept_overflow_rows = split_rank_rows(owner_dept_rank)
            lines.append("| 责任部门 | 起数 |")
            lines.append("|---|---:|")
            for row in owner_dept_table_rows:
                lines.append(f"| {safe_md_cell(row.get('name', ''))} | {row.get('count', 0)} |")
            overflow_note = format_overflow_note(owner_dept_overflow_rows, "部门")
            if overflow_note:
                lines.append("")
                lines.append(overflow_note)
        else:
            lines.append("无可统计数据")
        lines.append("")

        lines.append("## 总结")
        summary = (item.get("summary") or "").strip()
        lines.append(summary if summary else "无")
        lines.append("")

    return "\n".join(lines).strip() + "\n"
