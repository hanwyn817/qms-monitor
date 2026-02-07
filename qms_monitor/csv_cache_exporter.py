from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .config_loader import load_config
from .csv_io import dump_csv_manifest, write_csv_rows
from .excel_reader import ExcelBatchReader


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


def export_csv_cache(config_path: Path, output_dir: Path) -> tuple[Path, list[str]]:
    configs, warnings = load_config(config_path)

    rows_dir = output_dir / "rows"
    manifest_path = output_dir / "manifest.json"

    items: list[dict[str, Any]] = []
    batch_reader: ExcelBatchReader | None = None
    try:
        batch_reader = ExcelBatchReader(visible=False).open()

        for cfg in configs:
            sheet: str | int = int(cfg.sheet_name) if cfg.sheet_name.isdigit() else cfg.sheet_name
            ok, values, err, _, last_row, last_col, sheet_name = batch_reader.read_cells_sheet(
                cfg.file_path,
                sheet=sheet,
                auto_bounds=True,
                look_in="formulas",
            )

            rel_csv = Path("rows") / f"row_{cfg.row_no:04d}.csv"
            csv_path = output_dir / rel_csv

            if not ok:
                items.append(
                    {
                        "row_no": cfg.row_no,
                        "config": asdict(cfg),
                        "module": cfg.module,
                        "year": cfg.year,
                        "source_file": cfg.file_path,
                        "source_sheet": cfg.sheet_name,
                        "ok": False,
                        "error": err,
                    }
                )
                continue

            rows = _values_to_rows(values)
            write_csv_rows(csv_path, rows)
            items.append(
                {
                    "row_no": cfg.row_no,
                    "config": asdict(cfg),
                    "module": cfg.module,
                    "year": cfg.year,
                    "source_file": cfg.file_path,
                    "source_sheet": sheet_name or cfg.sheet_name,
                    "ok": True,
                    "csv_path": rel_csv.as_posix(),
                    "last_row": last_row,
                    "last_col": last_col,
                }
            )
    finally:
        if batch_reader is not None:
            try:
                batch_reader.close()
            except Exception:
                pass

    manifest = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "config_path": str(config_path),
        "output_dir": str(output_dir),
        "items": items,
    }
    dump_csv_manifest(manifest_path, manifest)
    return manifest_path, warnings
