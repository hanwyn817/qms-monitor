from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from .models import LedgerConfig


def read_csv_rows(path: Path) -> tuple[list[list[str]], str | None]:
    encodings = ["utf-8-sig", "utf-8", "gb18030"]
    for encoding in encodings:
        try:
            with path.open("r", encoding=encoding, newline="") as file:
                rows = [[cell.strip() for cell in row] for row in csv.reader(file)]
            return rows, None
        except UnicodeDecodeError:
            continue
        except OSError as exc:
            return [], str(exc)

    return [], f"无法解码CSV文件: {path}"


def write_csv_rows(path: Path, rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(rows)


def load_csv_manifest(path: Path) -> tuple[dict[int, Path], list[str]]:
    warnings: list[str] = []

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise RuntimeError(f"读取manifest失败: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"manifest不是有效JSON: {exc}") from exc

    items = payload.get("items")
    if not isinstance(items, list):
        raise RuntimeError("manifest缺少items数组")

    mapping: dict[int, Path] = {}
    for item in items:
        if not isinstance(item, dict):
            continue

        row_no = item.get("row_no")
        ok = bool(item.get("ok", True))
        csv_path = item.get("csv_path")

        if not isinstance(row_no, int):
            warnings.append(f"manifest项缺少有效row_no: {item}")
            continue

        if not ok:
            error = str(item.get("error", "未知错误"))
            warnings.append(f"manifest标记失败 row_no={row_no}: {error}")
            continue

        if not isinstance(csv_path, str) or not csv_path.strip():
            warnings.append(f"manifest项缺少csv_path row_no={row_no}")
            continue

        p = Path(csv_path)
        resolved = p if p.is_absolute() else (path.parent / p)
        mapping[row_no] = resolved

    return mapping, warnings


def dump_csv_manifest(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _to_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _to_int_optional(value: Any) -> int | None:
    if value is None or value == "":
        return None
    return _to_int(value)


def _config_from_dict(raw: dict[str, Any]) -> LedgerConfig | None:
    row_no = _to_int(raw.get("row_no"))
    id_col = _to_int(raw.get("id_col"))
    content_col = _to_int(raw.get("content_col"))
    initiated_col = _to_int(raw.get("initiated_col"))
    if row_no is None or id_col is None or content_col is None or initiated_col is None:
        return None

    return LedgerConfig(
        row_no=row_no,
        module=str(raw.get("module", "")).strip(),
        year=str(raw.get("year", "")).strip(),
        file_path=str(raw.get("file_path", "")).strip(),
        sheet_name=str(raw.get("sheet_name", "")).strip() or "1",
        id_col=id_col,
        content_col=content_col,
        initiated_col=initiated_col,
        planned_col=_to_int_optional(raw.get("planned_col")),
        status_col=_to_int_optional(raw.get("status_col")),
        owner_dept_col=_to_int_optional(raw.get("owner_dept_col")),
        owner_col=_to_int_optional(raw.get("owner_col")),
        qa_col=_to_int_optional(raw.get("qa_col")),
        qa_manager_col=_to_int_optional(raw.get("qa_manager_col")),
    )


def _parse_open_status_rules(raw: Any, warnings: list[str]) -> dict[str, str]:
    rules: dict[str, str] = {}

    if isinstance(raw, dict):
        for module, status in raw.items():
            module_key = str(module).strip()
            status_value = str(status).strip()
            if module_key and status_value:
                rules[module_key] = status_value
        return rules

    if isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                continue
            module = str(item.get("module", "")).strip()
            open_status = str(item.get("open_status", "")).strip()
            if module and open_status:
                rules[module] = open_status
        return rules

    if raw is not None:
        warnings.append("manifest中的open_status_rules格式无效，已忽略")
    return rules


def load_csv_manifest_bundle(path: Path) -> tuple[list[LedgerConfig], dict[int, Path], dict[str, str], list[str]]:
    warnings: list[str] = []

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise RuntimeError(f"读取manifest失败: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"manifest不是有效JSON: {exc}") from exc

    items = payload.get("items")
    if not isinstance(items, list):
        raise RuntimeError("manifest缺少items数组")

    open_status_rules = _parse_open_status_rules(payload.get("open_status_rules"), warnings)
    config_map: dict[int, LedgerConfig] = {}
    csv_map: dict[int, Path] = {}
    for item in items:
        if not isinstance(item, dict):
            continue

        row_no = _to_int(item.get("row_no"))
        if row_no is None:
            warnings.append(f"manifest项缺少有效row_no: {item}")
            continue

        config_raw = item.get("config")
        cfg: LedgerConfig | None = None
        if isinstance(config_raw, dict):
            cfg = _config_from_dict(config_raw)
        if cfg is None:
            warnings.append(f"manifest项缺少有效config row_no={row_no}")
        else:
            config_map[row_no] = cfg

        ok = bool(item.get("ok", True))
        if not ok:
            error = str(item.get("error", "未知错误"))
            warnings.append(f"manifest标记失败 row_no={row_no}: {error}")
            continue

        csv_path = item.get("csv_path")
        if not isinstance(csv_path, str) or not csv_path.strip():
            warnings.append(f"manifest项缺少csv_path row_no={row_no}")
            continue

        p = Path(csv_path)
        resolved = p if p.is_absolute() else (path.parent / p)
        csv_map[row_no] = resolved

    configs = [config_map[row_no] for row_no in sorted(config_map.keys())]
    return configs, csv_map, open_status_rules, warnings
