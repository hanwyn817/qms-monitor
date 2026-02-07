from __future__ import annotations

from pathlib import Path

from .constants import HEADER_LEN
from .excel_reader import read_excel_document
from .models import LedgerConfig
from .parsers import col_to_index, normalize_sheet_name, parse_tabular_text, parse_year


def load_config(config_path: Path) -> tuple[list[LedgerConfig], list[str]]:
    warnings: list[str] = []
    result = read_excel_document(str(config_path), sheet=1)
    if not result.ok:
        raise RuntimeError(f"读取配置失败: {result.error_type} - {result.error_message}")

    rows = parse_tabular_text(result.text)
    if not rows:
        raise RuntimeError("配置文件为空")

    configs: list[LedgerConfig] = []
    for i, row in enumerate(rows[1:], start=2):
        row_padded = row + [""] * max(0, HEADER_LEN - len(row))

        module = row_padded[1].strip()
        year = parse_year(row_padded[2])
        file_path = row_padded[3].strip()
        sheet_name = row_padded[4].strip()

        if not module and not file_path:
            continue

        id_col = col_to_index(row_padded[5])
        content_col = col_to_index(row_padded[6])
        initiated_col = col_to_index(row_padded[7])

        if not module:
            warnings.append(f"config第{i}行缺失质量模块，已跳过")
            continue
        if not file_path:
            warnings.append(f"config第{i}行缺失文件路径，已跳过: 模块={module}")
            continue
        if id_col is None or content_col is None or initiated_col is None:
            warnings.append(f"config第{i}行核心列(F/G/H)缺失或非法，已跳过: 模块={module}")
            continue

        configs.append(
            LedgerConfig(
                row_no=i,
                module=module,
                year=year,
                file_path=file_path,
                sheet_name=str(normalize_sheet_name(sheet_name)),
                id_col=id_col,
                content_col=content_col,
                initiated_col=initiated_col,
                planned_col=col_to_index(row_padded[8]),
                status_col=col_to_index(row_padded[9]),
                owner_dept_col=col_to_index(row_padded[10]),
                owner_col=col_to_index(row_padded[11]),
                qa_col=col_to_index(row_padded[12]),
                qa_manager_col=col_to_index(row_padded[13]),
            )
        )

    if not configs:
        raise RuntimeError("配置文件中没有可用配置")
    return configs, warnings
