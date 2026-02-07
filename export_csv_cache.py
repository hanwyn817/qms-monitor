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


def main() -> int:
    args = parse_args()
    config_path = Path(args.config)
    output_dir = Path(args.output_dir)

    if not config_path.exists():
        print(f"配置文件不存在: {config_path}", file=sys.stderr)
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
