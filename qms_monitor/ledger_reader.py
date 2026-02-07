from __future__ import annotations

from typing import Any

from .excel_reader import ExcelBatchReader, read_excel_document
from .models import LedgerConfig, QmsEvent
from .parsers import add_one_month, get_cell, parse_date_cell, parse_tabular_text


HEADER_HINTS = ("申请时间", "发起日期", "计划完成日期", "完成日期", "状态", "编号", "内容", "责任人", "责任部门", "分管")


def _safe_cell_str(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _values_to_rows(values: Any) -> list[list[str]]:
    if values is None:
        return []

    if not isinstance(values, tuple):
        return [[_safe_cell_str(values)]]

    rows: list[list[str]] = []
    for row in values:
        if isinstance(row, tuple):
            rows.append([_safe_cell_str(cell) for cell in row])
        else:
            rows.append([_safe_cell_str(row)])
    return rows


def _is_header_like_row(event_id: str, content: str, initiated_raw: str) -> bool:
    event_id_v = (event_id or "").strip().lower()
    content_v = (content or "").strip().lower()
    initiated_v = (initiated_raw or "").strip()
    if not initiated_v:
        return False

    header_hit = any(hint in initiated_v for hint in HEADER_HINTS)
    if not header_hit:
        return False

    id_like = ("编号" in event_id) or ("序号" in event_id) or (event_id_v in {"id", "编号", "序号"})
    content_like = ("内容" in content) or (content_v in {"content", "事项"})
    return id_like or content_like


def read_ledger_events(
    cfg: LedgerConfig,
    batch_reader: ExcelBatchReader | None = None,
    source_rows: list[list[str]] | None = None,
) -> tuple[list[QmsEvent], list[str]]:
    warnings: list[str] = []
    events: list[QmsEvent] = []

    sheet: str | int = int(cfg.sheet_name) if cfg.sheet_name.isdigit() else cfg.sheet_name
    rows: list[list[str]]

    if source_rows is not None:
        rows = source_rows
    elif batch_reader is not None:
        ok, values, err, _, _, _, _ = batch_reader.read_cells_sheet(
            cfg.file_path,
            sheet=sheet,
            auto_bounds=True,
            look_in="formulas",
        )
        if not ok:
            warnings.append(f"模块[{cfg.module}] 文件读取失败，已跳过: {cfg.file_path} ({err})")
            return events, warnings
        rows = _values_to_rows(values)
    else:
        result = read_excel_document(cfg.file_path, sheet=sheet)
        if not result.ok:
            warnings.append(
                f"模块[{cfg.module}] 文件读取失败，已跳过: {cfg.file_path} ({result.error_type}: {result.error_message})"
            )
            return events, warnings
        rows = parse_tabular_text(result.text)

    if len(rows) <= 1:
        warnings.append(f"模块[{cfg.module}] 表内容为空或只有表头: {cfg.file_path} / {cfg.sheet_name}")
        return events, warnings

    start_idx = max(2, cfg.data_start_row) - 1
    if start_idx >= len(rows):
        warnings.append(
            f"模块[{cfg.module}] 数据起始行[{cfg.data_start_row}]超出表格范围: {cfg.file_path} / {cfg.sheet_name}"
        )
        return events, warnings

    for row_idx, row in enumerate(rows[start_idx:], start=start_idx + 1):
        event_id = get_cell(row, cfg.id_col)
        content = get_cell(row, cfg.content_col)
        initiated_raw = get_cell(row, cfg.initiated_col)

        if not event_id and not content and not initiated_raw:
            continue

        initiated_date = parse_date_cell(initiated_raw)
        planned_raw = get_cell(row, cfg.planned_col)
        planned_date = parse_date_cell(planned_raw)

        if planned_date is None and initiated_date is not None:
            planned_date = add_one_month(initiated_date)

        status = get_cell(row, cfg.status_col)
        owner_dept = get_cell(row, cfg.owner_dept_col)
        owner = get_cell(row, cfg.owner_col)
        qa = get_cell(row, cfg.qa_col)
        qa_manager = get_cell(row, cfg.qa_manager_col)

        if initiated_date is None and _is_header_like_row(event_id, content, initiated_raw):
            continue

        if initiated_date is None and initiated_raw:
            warnings.append(
                f"模块[{cfg.module}] 行{row_idx}发起日期解析失败: '{initiated_raw}' ({cfg.file_path}/{cfg.sheet_name})"
            )

        events.append(
            QmsEvent(
                module=cfg.module,
                year=cfg.year,
                event_id=event_id,
                content=content,
                initiated_date=initiated_date,
                planned_date=planned_date,
                status=status,
                owner_dept=owner_dept,
                owner=owner,
                qa=qa,
                qa_manager=qa_manager,
                source_file=cfg.file_path,
                source_sheet=cfg.sheet_name,
                row_index=row_idx,
            )
        )

    return events, warnings
