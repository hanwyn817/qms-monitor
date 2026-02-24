from __future__ import annotations

import re
from pathlib import Path

from .constants import HEADER_LEN
from .excel_reader import read_excel_document
from .models import LedgerConfig
from .parsers import col_to_index, normalize_sheet_name, parse_tabular_text, parse_year


def _parse_data_start_row(raw: str, row_no: int, module: str, warnings: list[str]) -> int:
    value = (raw or "").strip()
    if not value:
        return 2
    try:
        n = int(float(value))
    except ValueError:
        warnings.append(f"config第{row_no}行 模块[{module}]数据起始行非法[{value}]，已回退为2")
        return 2
    if n < 2:
        warnings.append(f"config第{row_no}行 模块[{module}]数据起始行[{n}]小于2，已回退为2")
        return 2
    return n


def _parse_planned_rule(raw: str, row_no: int, module: str, warnings: list[str]) -> tuple[int | None, int | None]:
    value = (raw or "").strip()
    if not value:
        return None, None

    if re.fullmatch(r"\d+(\.0+)?", value):
        return None, int(float(value))

    if re.fullmatch(r"[A-Za-z]+", value):
        planned_col = col_to_index(value)
        if planned_col is not None:
            return planned_col, None

    warnings.append(
        f"config第{row_no}行 模块[{module}]计划完成规则非法[{value}]，应为列标字母(如J/AA)或数字天数"
    )
    return None, None


def load_config(config_path: Path) -> tuple[list[LedgerConfig], list[str]]:
    warnings: list[str] = []
    result = read_excel_document(str(config_path), sheet=1)
    if not result.ok:
        raise RuntimeError(f"读取配置失败: {result.error_type} - {result.error_message}")

    rows = parse_tabular_text(result.text)
    if not rows:
        raise RuntimeError("配置文件为空")

    configs: list[LedgerConfig] = []
    
    header_row = [c.strip() for c in rows[0]]
    def get_col_idx(*keywords: str, default: int) -> int:
        for kw in keywords:
            for idx, col in enumerate(header_row):
                if kw in col:
                    return idx
        return default

    idx_topic = get_col_idx("主题", default=1)
    idx_module = get_col_idx("质量模块", "模块", default=2)
    idx_year = get_col_idx("年份", default=3)
    idx_file_path = get_col_idx("路径", "文件", default=4)
    idx_sheet_name = get_col_idx("sheet", "表名称", default=5)
    idx_id_col = get_col_idx("编号", default=6)
    idx_content_col = get_col_idx("内容", default=7)
    idx_initiated_col = get_col_idx("发起日期", "时间", default=8)
    idx_planned_rule = get_col_idx("计划规则", "完成时限", default=9)
    idx_status_col = get_col_idx("状态", default=10)
    idx_owner_dept_col = get_col_idx("责任部门", "部门", default=11)
    idx_owner_col = get_col_idx("责任人", default=12)
    idx_qa_col = get_col_idx("分管 QA 列", "分管 QA", "分管QA", default=13)
    idx_qa_manager_col = get_col_idx("分管 QA 中层列", "分管 QA 中层", "分管QA中层", default=14)
    idx_open_status = get_col_idx("未完成状态", "未完成", default=15)
    idx_data_start_row = get_col_idx("数据起始行", "起始行", default=16)

    def extract_val(r: list[str], i: int) -> str:
        return r[i] if i < len(r) else ""

    for i, row in enumerate(rows[1:], start=2):
        topic = extract_val(row, idx_topic).strip()
        module = extract_val(row, idx_module).strip()
        year = parse_year(extract_val(row, idx_year))
        file_path = extract_val(row, idx_file_path).strip()
        sheet_name = extract_val(row, idx_sheet_name).strip()

        if not module and not file_path:
            continue

        id_col = col_to_index(extract_val(row, idx_id_col))
        content_col = col_to_index(extract_val(row, idx_content_col))
        initiated_col = col_to_index(extract_val(row, idx_initiated_col))

        if not module:
            warnings.append(f"config第{i}行缺失质量模块，已跳过")
            continue
        if not file_path:
            warnings.append(f"config第{i}行缺失文件路径，已跳过: 模块={module}")
            continue
        if id_col is None or content_col is None or initiated_col is None:
            warnings.append(f"config第{i}行核心列(编号/内容/发起日期)缺失或非法，已跳过: 模块={module}")
            continue
        planned_col, planned_due_days = _parse_planned_rule(extract_val(row, idx_planned_rule), i, module, warnings)

        configs.append(
            LedgerConfig(
                row_no=i,
                topic=topic,
                module=module,
                year=year,
                file_path=file_path,
                sheet_name=str(normalize_sheet_name(sheet_name)),
                id_col=id_col,
                content_col=content_col,
                initiated_col=initiated_col,
                planned_col=planned_col,
                planned_due_days=planned_due_days,
                status_col=col_to_index(extract_val(row, idx_status_col)),
                owner_dept_col=col_to_index(extract_val(row, idx_owner_dept_col)),
                owner_col=col_to_index(extract_val(row, idx_owner_col)),
                qa_col=col_to_index(extract_val(row, idx_qa_col)),
                qa_manager_col=col_to_index(extract_val(row, idx_qa_manager_col)),
                open_status_value=extract_val(row, idx_open_status).strip(),
                data_start_row=_parse_data_start_row(extract_val(row, idx_data_start_row), i, module, warnings),
            )
        )

    if not configs:
        raise RuntimeError("配置文件中没有可用配置")
    return configs, warnings


def build_open_status_rules(configs: list[LedgerConfig]) -> dict[str, str]:
    errors: list[str] = []
    rules: dict[str, str] = {}

    for cfg in configs:
        module = cfg.module.strip()
        open_status = cfg.open_status_value.strip()
        if not module:
            continue
        if not open_status:
            errors.append(f"config第{cfg.row_no}行 模块[{module}]缺少未完成状态值")
            continue
        existing = rules.get(module)
        if existing is not None and existing != open_status:
            errors.append(
                f"模块[{module}]存在多个未完成状态值: [{existing}] 与 [{open_status}]"
            )
            continue
        rules[module] = open_status

    if errors:
        details = "; ".join(errors)
        raise RuntimeError(f"未完成状态值配置错误: {details}")

    return rules
