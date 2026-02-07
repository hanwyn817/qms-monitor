from __future__ import annotations

from collections import Counter
from datetime import date
from typing import Any

from .models import QmsEvent


__all__ = [
    "build_local_stats",
    "build_topic_stats",
    "build_event_records",
    "build_overdue_event_records",
    "is_open_status",
]


def is_open_status(module: str, status: str, open_status_rules: dict[str, str]) -> bool:
    status_value = (status or "").strip()
    open_status = open_status_rules.get(module)
    if open_status is None:
        raise ValueError(f"模块[{module}]未配置未完成状态值")
    return status_value == open_status


def build_local_stats(
    module: str,
    events: list[QmsEvent],
    report_date: date,
    open_status_rules: dict[str, str] | None = None,
) -> dict[str, Any]:
    rules = open_status_rules or {}
    yearly_counter: Counter[str] = Counter()
    overdue_items: list[dict[str, Any]] = []

    for event in events:
        yearly_counter[event.year] += 1

        if event.planned_date and event.planned_date < report_date and is_open_status(module, event.status, rules):
            overdue_items.append(
                {
                    "topic": event.topic,
                    "module": event.module,
                    "year": event.year,
                    "event_id": event.event_id,
                    "content": event.content,
                    "initiated_date": event.initiated_date_str,
                    "planned_date": event.planned_date_str,
                    "status": event.status,
                    "owner_dept": event.owner_dept,
                    "owner": event.owner,
                    "qa": event.qa,
                    "qa_manager": event.qa_manager,
                    "source": f"{event.source_file} | {event.source_sheet} | row {event.row_index}",
                }
            )

    overdue_items.sort(key=lambda x: (x.get("planned_date") or "9999-12-31", x.get("event_id") or ""))
    total = len(events)
    overdue_count = len(overdue_items)
    ratio = round((overdue_count / total) * 100, 2) if total else 0.0

    qa_counter = Counter(item["qa"] for item in overdue_items if item.get("qa"))
    qa_manager_counter = Counter(item["qa_manager"] for item in overdue_items if item.get("qa_manager"))

    yearly_totals = [{"year": y, "count": c} for y, c in sorted(yearly_counter.items(), key=lambda x: x[0])]
    overdue_by_qa = [{"name": name, "count": count} for name, count in sorted(qa_counter.items(), key=lambda x: (-x[1], x[0]))]
    overdue_by_qa_manager = [
        {"name": name, "count": count}
        for name, count in sorted(qa_manager_counter.items(), key=lambda x: (-x[1], x[0]))
    ]

    return {
        "module": module,
        "yearly_totals": yearly_totals,
        "overdue": {
            "count": overdue_count,
            "ratio": ratio,
            "items": overdue_items,
        },
        "overdue_by_qa": overdue_by_qa,
        "overdue_by_qa_manager": overdue_by_qa_manager,
    }


def build_topic_stats(
    topic: str,
    events: list[QmsEvent],
    report_date: date,
    open_status_rules: dict[str, str] | None = None,
) -> dict[str, Any]:
    rules = open_status_rules or {}
    yearly_total_counter: Counter[str] = Counter()
    yearly_overdue_counter: Counter[str] = Counter()
    module_total_counter: Counter[str] = Counter()
    module_overdue_counter: Counter[str] = Counter()
    overdue_items: list[dict[str, Any]] = []

    for event in events:
        yearly_total_counter[event.year] += 1
        module_total_counter[event.module] += 1

        is_overdue = bool(
            event.planned_date
            and event.planned_date < report_date
            and is_open_status(event.module, event.status, rules)
        )
        if not is_overdue:
            continue

        yearly_overdue_counter[event.year] += 1
        module_overdue_counter[event.module] += 1
        overdue_items.append(
            {
                "topic": event.topic,
                "module": event.module,
                "year": event.year,
                "event_id": event.event_id,
                "content": event.content,
                "initiated_date": event.initiated_date_str,
                "planned_date": event.planned_date_str,
                "status": event.status,
                "owner_dept": event.owner_dept,
                "owner": event.owner,
                "qa": event.qa,
                "qa_manager": event.qa_manager,
                "source": f"{event.source_file} | {event.source_sheet} | row {event.row_index}",
            }
        )

    overdue_items.sort(key=lambda x: (x.get("planned_date") or "9999-12-31", x.get("event_id") or ""))
    total_count = len(events)
    overdue_count = len(overdue_items)
    overdue_ratio = round((overdue_count / total_count) * 100, 2) if total_count else 0.0

    yearly_totals = [{"year": y, "count": c} for y, c in sorted(yearly_total_counter.items(), key=lambda x: x[0])]
    yearly_overdue = []
    for year, total in sorted(yearly_total_counter.items(), key=lambda x: x[0]):
        overdue = yearly_overdue_counter.get(year, 0)
        ratio = round((overdue / total) * 100, 2) if total else 0.0
        yearly_overdue.append(
            {
                "year": year,
                "count": total,
                "overdue_count": overdue,
                "overdue_ratio": ratio,
            }
        )

    by_module = []
    for module, total in sorted(module_total_counter.items(), key=lambda x: x[0]):
        overdue = module_overdue_counter.get(module, 0)
        ratio = round((overdue / total) * 100, 2) if total else 0.0
        by_module.append(
            {
                "module": module,
                "count": total,
                "overdue_count": overdue,
                "overdue_ratio": ratio,
            }
        )

    qa_counter = Counter(item["qa"] for item in overdue_items if item.get("qa"))
    qa_manager_counter = Counter(item["qa_manager"] for item in overdue_items if item.get("qa_manager"))
    overdue_by_qa = [{"name": name, "count": count} for name, count in sorted(qa_counter.items(), key=lambda x: (-x[1], x[0]))]
    overdue_by_qa_manager = [
        {"name": name, "count": count}
        for name, count in sorted(qa_manager_counter.items(), key=lambda x: (-x[1], x[0]))
    ]

    return {
        "topic": topic,
        "yearly_totals": yearly_totals,
        "yearly_overdue": yearly_overdue,
        "total": {"count": total_count},
        "overdue": {
            "count": overdue_count,
            "ratio": overdue_ratio,
            "items": overdue_items,
        },
        "by_module": by_module,
        "overdue_by_qa": overdue_by_qa,
        "overdue_by_qa_manager": overdue_by_qa_manager,
    }


def build_event_records(
    events: list[QmsEvent],
    open_status_rules: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    rules = open_status_rules or {}
    records: list[dict[str, Any]] = []
    for event in events:
        is_open = is_open_status(event.module, event.status, rules)
        records.append(
            {
                "topic": event.topic,
                "module": event.module,
                "year": event.year,
                "event_id": event.event_id,
                "content": event.content,
                "initiated_date": event.initiated_date_str,
                "planned_date": event.planned_date_str,
                "status": event.status,
                "owner_dept": event.owner_dept,
                "owner": event.owner,
                "qa": event.qa,
                "qa_manager": event.qa_manager,
                "status_semantic": "open" if is_open else "completed",
                "source_file": event.source_file,
                "source_sheet": event.source_sheet,
                "source_row": event.row_index,
            }
        )
    return records


def build_overdue_event_records(
    events: list[QmsEvent],
    report_date: date,
    open_status_rules: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    rules = open_status_rules or {}
    records: list[dict[str, Any]] = []
    for event in events:
        is_open = is_open_status(event.module, event.status, rules)
        is_overdue = bool(
            event.planned_date
            and event.planned_date < report_date
            and is_open
        )
        if not is_overdue:
            continue
        records.append(
            {
                "topic": event.topic,
                "module": event.module,
                "year": event.year,
                "event_id": event.event_id,
                "content": event.content,
                "initiated_date": event.initiated_date_str,
                "planned_date": event.planned_date_str,
                "status": event.status,
                "owner_dept": event.owner_dept,
                "owner": event.owner,
                "qa": event.qa,
                "qa_manager": event.qa_manager,
                "status_semantic": "open" if is_open else "completed",
                "source_file": event.source_file,
                "source_sheet": event.source_sheet,
                "source_row": event.row_index,
            }
        )
    return records
