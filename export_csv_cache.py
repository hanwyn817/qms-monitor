from __future__ import annotations

import argparse
import sys
from pathlib import Path

from qms_monitor.csv_cache_exporter import export_csv_cache


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export Excel ledgers to CSV cache")
    parser.add_argument("--config", default="config.xlsx", help="配置文件路径")
    parser.add_argument(
        "--output-dir",
        default="artifacts/csv_cache",
        help="CSV缓存输出目录（会生成rows/和manifest.json）",
    )
    return parser.parse_args()


def resolve_config_path(raw_path: str) -> tuple[Path | None, list[Path]]:
    candidates: list[Path] = []
    entered = Path(raw_path).expanduser()

    if entered.is_absolute():
        candidates.append(entered)
    else:
        cwd = Path.cwd()
        project_root = Path(__file__).resolve().parent
        candidates.append(cwd / entered)
        if (project_root / entered) not in candidates:
            candidates.append(project_root / entered)

    seen: set[str] = set()
    deduped: list[Path] = []
    for candidate in candidates:
        key = str(candidate.resolve(strict=False))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)

    for candidate in deduped:
        if candidate.exists() and candidate.is_file():
            return candidate, deduped
    return None, deduped


def main() -> int:
    args = parse_args()
    config_path, tried_paths = resolve_config_path(args.config)
    output_dir = Path(args.output_dir)

    if config_path is None:
        print("配置文件不存在。", file=sys.stderr)
        print(f"当前工作目录: {Path.cwd()}", file=sys.stderr)
        print("已尝试路径:", file=sys.stderr)
        for path in tried_paths:
            print(f"- {path}", file=sys.stderr)

        project_root = Path(__file__).resolve().parent
        suggestions = sorted(project_root.glob("config*.xls*"))
        if suggestions:
            print("项目目录下发现以下相似文件:", file=sys.stderr)
            for file in suggestions:
                print(f"- {file}", file=sys.stderr)
        return 1

    try:
        manifest_path, warnings = export_csv_cache(config_path, output_dir)
    except Exception as exc:
        print(f"导出CSV缓存失败: {exc}", file=sys.stderr)
        return 1

    print(f"CSV缓存导出完成: {manifest_path}")
    if warnings:
        print("导出告警:")
        for warning in warnings:
            print(f"- {warning}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
