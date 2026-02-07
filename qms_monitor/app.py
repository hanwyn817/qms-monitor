from __future__ import annotations

import json
import os
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .cli import parse_args
from .config_loader import build_open_status_rules, load_config
from .csv_io import load_csv_manifest_bundle, read_csv_rows
from .excel_reader import ExcelBatchReader
from .ledger_reader import read_ledger_events
from .llm_client import call_llm
from .models import QmsEvent
from .overdue_excel_exporter import export_overdue_events_excel
from .pdf_exporter import export_markdown_file_to_pdf
from .pdf_exporter_latex import export_markdown_file_to_pdf_latex
from .report_renderer import render_markdown_report
from .stats import build_event_records, build_local_stats, build_overdue_event_records, build_topic_stats


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
    open_status_rules: dict[str, str] = {}

    if args.input_mode == "csv":
        if not args.csv_manifest:
            print("CSV模式需要提供 --csv-manifest", file=sys.stderr)
            return 1

        manifest_path = Path(args.csv_manifest)
        if not manifest_path.exists():
            print(f"CSV manifest不存在: {manifest_path}", file=sys.stderr)
            return 1

        try:
            manifest_configs, csv_map, open_status_rules, csv_warnings = load_csv_manifest_bundle(manifest_path)
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
                open_status_rules = build_open_status_rules(configs)
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
            open_status_rules = build_open_status_rules(configs)
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

    module_local_results: dict[str, dict[str, Any]] = {}
    for module, events in grouped.items():
        module_local_results[module] = build_local_stats(module, events, report_date, open_status_rules)

    topic_grouped: dict[str, list[QmsEvent]] = defaultdict(list)
    for events in grouped.values():
        for event in events:
            topic_grouped[(event.topic or "").strip() or "未分类"].append(event)

    topic_results: dict[str, dict[str, Any]] = {}
    for topic, events in topic_grouped.items():
        local_stats = build_topic_stats(topic, events, report_date, open_status_rules)
        overdue_records = build_overdue_event_records(events, report_date, open_status_rules)

        if args.skip_llm:
            topic_results[topic] = local_stats
            continue

        try:
            llm_start = time.time()
            print(f"[LLM] 开始分析主题[{topic}] ...", file=sys.stderr, flush=True)
            llm_stats = call_llm(
                topic=topic,
                report_date=report_date,
                local_stats=local_stats,
                overdue_records=overdue_records,
                base_url=os.getenv("QMS_LLM_BASE_URL", "https://api.openai.com/v1"),
                model=os.getenv("QMS_LLM_MODEL", ""),
                api_key=os.getenv("QMS_LLM_API_KEY", ""),
                timeout_seconds=int(os.getenv("QMS_LLM_TIMEOUT", "120")),
                progress_interval_seconds=int(os.getenv("QMS_LLM_PROGRESS_INTERVAL", "15")),
            )
            elapsed = time.time() - llm_start
            print(f"[LLM] 主题[{topic}] 分析完成，用时 {elapsed:.1f}s", file=sys.stderr, flush=True)
            topic_results[topic] = llm_stats
        except Exception as exc:
            warnings.append(f"主题[{topic}] LLM分析失败，已回退本地统计: {exc}")
            print(f"[LLM] 主题[{topic}] 分析失败: {exc}", file=sys.stderr, flush=True)
            topic_results[topic] = local_stats

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = output_dir / f"qms_report_{timestamp}.md"
    pdf_path = output_dir / f"qms_report_{timestamp}.pdf"
    detail_path = output_dir / f"qms_report_{timestamp}.json"
    overdue_excel_path = output_dir / f"qms_overdue_events_{timestamp}.xlsx"

    overdue_event_count = 0
    overdue_excel_exported = False
    try:
        overdue_event_count = export_overdue_events_excel(overdue_excel_path, module_local_results)
        overdue_excel_exported = True
    except Exception as exc:
        warnings.append(f"超期事件Excel导出失败: {exc}")
        print(f"[EXPORT] 超期事件Excel导出失败: {exc}", file=sys.stderr, flush=True)

    report_text = render_markdown_report(
        report_date=report_date,
        config_path=config_path,
        topic_results=topic_results,
        warnings=warnings,
        processed_files=processed_files,
        skipped_files=skipped_files,
    )
    report_path.write_text(report_text, encoding="utf-8")
    pdf_exported = False
    pdf_engine = os.getenv("QMS_PDF_ENGINE", "latex").strip().lower() or "latex"
    try:
        if pdf_engine == "reportlab":
            export_markdown_file_to_pdf(report_path, pdf_path)
        else:
            export_markdown_file_to_pdf_latex(report_path, pdf_path)
        pdf_exported = True
    except Exception as exc:
        try:
            export_markdown_file_to_pdf(report_path, pdf_path)
            pdf_exported = True
            fallback_msg = f"PDF导出已回退到reportlab: {exc}"
            warnings.append(fallback_msg)
            print(f"[EXPORT] {fallback_msg}", file=sys.stderr, flush=True)
        except Exception as fallback_exc:
            warnings.append(f"PDF导出失败: {exc}; 回退失败: {fallback_exc}")
            print(f"[EXPORT] PDF导出失败: {exc}; 回退失败: {fallback_exc}", file=sys.stderr, flush=True)

    detail_payload = {
        "report_date": report_date.isoformat(),
        "config": str(config_path),
        "processed_files": processed_files,
        "skipped_files": skipped_files,
        "warnings": warnings,
        "pdf_report": str(pdf_path) if pdf_exported else "",
        "overdue_excel": str(overdue_excel_path) if overdue_excel_exported else "",
        "overdue_event_count": overdue_event_count,
        "topics": topic_results,
    }
    detail_path.write_text(json.dumps(detail_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"报告已生成: {report_path}")
    if pdf_exported:
        print(f"PDF已生成: {pdf_path}")
    print(f"明细已生成: {detail_path}")
    if overdue_excel_exported:
        print(f"超期事件Excel已生成: {overdue_excel_path} (共 {overdue_event_count} 条)")
    return 0
