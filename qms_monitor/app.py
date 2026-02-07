from __future__ import annotations

import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .cli import parse_args
from .config_loader import load_config
from .csv_io import load_csv_manifest_bundle, read_csv_rows
from .excel_reader import ExcelBatchReader
from .ledger_reader import read_ledger_events
from .llm_client import call_llm
from .models import QmsEvent
from .report_renderer import render_markdown_report
from .stats import build_event_records, build_local_stats


def main() -> int:
    args = parse_args()

    config_path = Path(args.config)

    try:
        report_date = datetime.strptime(args.report_date, "%Y-%m-%d").date()
    except ValueError:
        print("--report-date 格式必须是 YYYY-MM-DD", file=sys.stderr)
        return 1

    warnings: list[str] = []
    configs = []
    csv_map: dict[int, Path] = {}

    if args.input_mode == "csv":
        if not args.csv_manifest:
            print("CSV模式需要提供 --csv-manifest", file=sys.stderr)
            return 1

        manifest_path = Path(args.csv_manifest)
        if not manifest_path.exists():
            print(f"CSV manifest不存在: {manifest_path}", file=sys.stderr)
            return 1

        try:
            manifest_configs, csv_map, csv_warnings = load_csv_manifest_bundle(manifest_path)
            warnings.extend(csv_warnings)
        except Exception as exc:
            print(f"读取CSV manifest失败: {exc}", file=sys.stderr)
            return 1

        if manifest_configs:
            configs = manifest_configs
        else:
            if not config_path.exists():
                print("manifest未包含有效config，且--config文件不存在", file=sys.stderr)
                return 1
            try:
                configs, config_warnings = load_config(config_path)
                warnings.extend(config_warnings)
            except Exception as exc:
                print(f"读取配置失败: {exc}", file=sys.stderr)
                return 1
    else:
        if not config_path.exists():
            print(f"配置文件不存在: {config_path}", file=sys.stderr)
            return 1
        try:
            configs, config_warnings = load_config(config_path)
            warnings.extend(config_warnings)
        except Exception as exc:
            print(f"读取配置失败: {exc}", file=sys.stderr)
            return 1

    grouped: dict[str, list[QmsEvent]] = defaultdict(list)
    processed_files = 0
    skipped_files = 0
    if args.input_mode == "csv":
        for cfg in configs:
            csv_path = csv_map.get(cfg.row_no)
            if csv_path is None:
                warnings.append(f"模块[{cfg.module}] row_no={cfg.row_no} 在manifest中未找到CSV，已跳过")
                skipped_files += 1
                continue

            rows, err = read_csv_rows(csv_path)
            if err:
                warnings.append(f"模块[{cfg.module}] CSV读取失败，已跳过: {csv_path} ({err})")
                skipped_files += 1
                continue

            events, ledger_warnings = read_ledger_events(cfg, source_rows=rows)
            warnings.extend(ledger_warnings)
            if ledger_warnings and not events:
                skipped_files += 1
            else:
                processed_files += 1
            grouped[cfg.module].extend(events)
    else:
        batch_reader: ExcelBatchReader | None = None
        try:
            try:
                batch_reader = ExcelBatchReader(visible=False).open()
            except Exception as exc:
                warnings.append(f"批量读取初始化失败，已回退单文件读取: {exc}")
                batch_reader = None

            for cfg in configs:
                events, ledger_warnings = read_ledger_events(cfg, batch_reader=batch_reader)
                warnings.extend(ledger_warnings)
                if ledger_warnings and not events:
                    skipped_files += 1
                else:
                    processed_files += 1
                grouped[cfg.module].extend(events)
        finally:
            if batch_reader is not None:
                try:
                    batch_reader.close()
                except Exception:
                    pass

    module_results: dict[str, dict[str, Any]] = {}
    for module, events in grouped.items():
        local_stats = build_local_stats(module, events, report_date)
        records = build_event_records(events)

        if args.skip_llm:
            module_results[module] = local_stats
            continue

        try:
            llm_stats = call_llm(
                module=module,
                report_date=report_date,
                local_stats=local_stats,
                event_records=records,
                base_url=args.llm_base_url,
                model=args.llm_model,
                api_key=args.llm_api_key,
                timeout_seconds=args.llm_timeout,
            )
            module_results[module] = llm_stats
        except Exception as exc:
            warnings.append(f"模块[{module}] LLM分析失败，已回退本地统计: {exc}")
            module_results[module] = local_stats

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = output_dir / f"qms_report_{timestamp}.md"
    detail_path = output_dir / f"qms_report_{timestamp}.json"

    report_text = render_markdown_report(
        report_date=report_date,
        config_path=config_path,
        module_results=module_results,
        warnings=warnings,
        processed_files=processed_files,
        skipped_files=skipped_files,
    )
    report_path.write_text(report_text, encoding="utf-8")

    detail_payload = {
        "report_date": report_date.isoformat(),
        "config": str(config_path),
        "processed_files": processed_files,
        "skipped_files": skipped_files,
        "warnings": warnings,
        "modules": module_results,
    }
    detail_path.write_text(json.dumps(detail_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"报告已生成: {report_path}")
    print(f"明细已生成: {detail_path}")
    return 0
