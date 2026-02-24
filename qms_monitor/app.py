from __future__ import annotations

import json
import os
import sys
import time
import argparse
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any

from .cli import parse_args
from .config_loader import build_open_status_rules, load_config
from .csv_io import load_csv_manifest_bundle, read_csv_rows
from .excel_reader import ExcelBatchReader
from .ledger_reader import read_ledger_events
from .llm_client import call_llm_person_summaries, call_llm_topic_summary
from .models import QmsEvent
from .overdue_excel_exporter import export_overdue_events_excel
from .pdf_exporter import export_markdown_file_to_pdf
from .pdf_exporter_latex import export_markdown_file_to_pdf_latex
from .report_renderer import render_markdown_report
from .stats import build_event_records, build_local_stats, build_overdue_event_records, build_topic_stats

class ReportOrchestrator:
    def __init__(self, args: argparse.Namespace):
        self.args = args
        self.config_path = Path(args.config)
        self.output_dir = Path(args.output_dir)
        try:
            self.report_date = datetime.strptime(args.report_date, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError("--report-date 格式必须是 YYYY-MM-DD") from exc

        self.warnings: list[str] = []
        self.configs = []
        self.csv_map: dict[int, Path] = {}
        self.open_status_rules: dict[str, str] = {}
        
        self.processed_files = 0
        self.skipped_files = 0

    def run(self) -> int:
        try:
            self.load_configuration()
            grouped = self.ingest_data()
            topic_results, module_local_results = self.analyze_data(grouped)
            self.export_reports(topic_results, module_local_results)
            return 0
        except ValueError as exc:
            print(f"数据校验失败: {exc}", file=sys.stderr)
            return 1
        except Exception as exc:
            import traceback
            traceback.print_exc()
            print(f"执行发生系统错误: {exc}", file=sys.stderr)
            return 1

    def load_configuration(self) -> None:
        if self.args.input_mode == "csv":
            if not self.args.csv_manifest:
                raise ValueError("CSV模式需要提供 --csv-manifest")
            manifest_path = Path(self.args.csv_manifest)
            if not manifest_path.exists():
                raise ValueError(f"CSV manifest不存在: {manifest_path}")
            
            manifest_configs, self.csv_map, self.open_status_rules, csv_warnings = load_csv_manifest_bundle(manifest_path)
            self.warnings.extend(csv_warnings)
            
            if manifest_configs:
                self.configs = manifest_configs
            else:
                if not self.config_path.exists():
                    raise ValueError("manifest未包含有效config，且--config文件不存在")
                self.configs, config_warnings = load_config(self.config_path)
                self.warnings.extend(config_warnings)
                self.open_status_rules = build_open_status_rules(self.configs)
        else:
            if not self.config_path.exists():
                raise ValueError(f"配置文件不存在: {self.config_path}")
            self.configs, config_warnings = load_config(self.config_path)
            self.warnings.extend(config_warnings)
            self.open_status_rules = build_open_status_rules(self.configs)

    def ingest_data(self) -> dict[str, list[QmsEvent]]:
        grouped: dict[str, list[QmsEvent]] = defaultdict(list)
        if self.args.input_mode == "csv":
            for cfg in self.configs:
                csv_path = self.csv_map.get(cfg.row_no)
                if csv_path is None:
                    self.warnings.append(f"模块[{cfg.module}] row_no={cfg.row_no} 在manifest中未找到CSV，已跳过")
                    self.skipped_files += 1
                    continue

                rows, err = read_csv_rows(csv_path)
                if err:
                    self.warnings.append(f"模块[{cfg.module}] CSV读取失败，已跳过: {csv_path} ({err})")
                    self.skipped_files += 1
                    continue

                events, ledger_warnings = read_ledger_events(cfg, source_rows=rows)
                self.warnings.extend(ledger_warnings)
                if ledger_warnings and not events:
                    self.skipped_files += 1
                else:
                    self.processed_files += 1
                grouped[cfg.module].extend(events)
        else:
            batch_reader: ExcelBatchReader | None = None
            try:
                try:
                    batch_reader = ExcelBatchReader(visible=False).open()
                except Exception as exc:
                    self.warnings.append(f"批量读取初始化失败，已回退单文件读取: {exc}")
                    batch_reader = None

                for cfg in self.configs:
                    events, ledger_warnings = read_ledger_events(cfg, batch_reader=batch_reader)
                    self.warnings.extend(ledger_warnings)
                    if ledger_warnings and not events:
                        self.skipped_files += 1
                    else:
                        self.processed_files += 1
                    grouped[cfg.module].extend(events)
            finally:
                if batch_reader is not None:
                    try:
                        batch_reader.close()
                    except Exception:
                        pass
        return grouped

    def _process_single_topic_llm(self, topic: str, local_stats: dict[str, Any], overdue_records: list[dict[str, Any]]) -> dict[str, Any]:
        base_url = os.getenv("QMS_LLM_BASE_URL", "https://api.openai.com/v1")
        model = os.getenv("QMS_LLM_MODEL", "")
        api_key = os.getenv("QMS_LLM_API_KEY", "")
        timeout_seconds = int(os.getenv("QMS_LLM_TIMEOUT", "120"))
        progress_interval_seconds = int(os.getenv("QMS_LLM_PROGRESS_INTERVAL", "15"))

        merged_stats = dict(local_stats)
        
        try:
            llm_start = time.time()
            print(f"[LLM] 开始主题总结[{topic}] ...", file=sys.stderr, flush=True)
            summary = call_llm_topic_summary(
                topic=topic,
                report_date=self.report_date,
                local_stats=local_stats,
                overdue_records=overdue_records,
                base_url=base_url,
                model=model,
                api_key=api_key,
                timeout_seconds=timeout_seconds,
                progress_interval_seconds=progress_interval_seconds,
            )
            merged_stats["summary"] = summary
            elapsed = time.time() - llm_start
            print(f"[LLM] 主题总结[{topic}] 完成，用时 {elapsed:.1f}s", file=sys.stderr, flush=True)
        except Exception as exc:
            self.warnings.append(f"主题[{topic}] LLM主题总结失败，已回退本地统计: {exc}")
            print(f"[LLM] 主题总结[{topic}] 失败: {exc}", file=sys.stderr, flush=True)
            merged_stats.setdefault("summary", local_stats.get("summary", ""))

        try:
            llm_start = time.time()
            print(f"[LLM] 开始人员概括[{topic}] ...", file=sys.stderr, flush=True)
            merged_stats = call_llm_person_summaries(
                topic=topic,
                report_date=self.report_date,
                local_stats=merged_stats,
                base_url=base_url,
                model=model,
                api_key=api_key,
                timeout_seconds=timeout_seconds,
                progress_interval_seconds=progress_interval_seconds,
            )
            elapsed = time.time() - llm_start
            print(f"[LLM] 人员概括[{topic}] 完成，用时 {elapsed:.1f}s", file=sys.stderr, flush=True)
        except Exception as exc:
            self.warnings.append(f"主题[{topic}] LLM人员概括失败，已保留现有统计: {exc}")
            print(f"[LLM] 人员概括[{topic}] 失败: {exc}", file=sys.stderr, flush=True)

        return merged_stats

    def analyze_data(self, grouped: dict[str, list[QmsEvent]]) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
        module_local_results: dict[str, dict[str, Any]] = {}
        for module, events in grouped.items():
            module_local_results[module] = build_local_stats(module, events, self.report_date, self.open_status_rules)

        topic_grouped: dict[str, list[QmsEvent]] = defaultdict(list)
        for events in grouped.values():
            for event in events:
                topic_grouped[(event.topic or "").strip() or "未分类"].append(event)

        topic_results: dict[str, dict[str, Any]] = {}

        if self.args.skip_llm:
            for topic, events in topic_grouped.items():
                local_stats = build_topic_stats(topic, events, self.report_date, self.open_status_rules)
                topic_results[topic] = local_stats
            return topic_results, module_local_results

        topic_payloads = []
        for topic, events in topic_grouped.items():
            local_stats = build_topic_stats(topic, events, self.report_date, self.open_status_rules)
            overdue_records = build_overdue_event_records(events, self.report_date, self.open_status_rules)
            topic_payloads.append((topic, local_stats, overdue_records))

        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {
                executor.submit(self._process_single_topic_llm, t, ls, ov): t
                for t, ls, ov in topic_payloads
            }
            for future in as_completed(futures):
                topic = futures[future]
                try:
                    merged_stats = future.result()
                    topic_results[topic] = merged_stats
                except Exception as exc:
                    self.warnings.append(f"主题[{topic}]并发处理失败: {exc}")
                    
        return topic_results, module_local_results

    def export_reports(self, topic_results: dict[str, dict[str, Any]], module_local_results: dict[str, dict[str, Any]]) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = self.output_dir / f"qms_report_{timestamp}.md"
        pdf_path = self.output_dir / f"qms_report_{timestamp}.pdf"
        detail_path = self.output_dir / f"qms_report_{timestamp}.json"
        overdue_excel_path = self.output_dir / f"qms_overdue_events_{timestamp}.xlsx"

        overdue_event_count = 0
        overdue_excel_exported = False
        try:
            overdue_event_count = export_overdue_events_excel(overdue_excel_path, module_local_results)
            overdue_excel_exported = True
        except Exception as exc:
            self.warnings.append(f"超期事件Excel导出失败: {exc}")
            print(f"[EXPORT] 超期事件Excel导出失败: {exc}", file=sys.stderr, flush=True)

        report_text = render_markdown_report(
            report_date=self.report_date,
            config_path=self.config_path,
            topic_results=topic_results,
            warnings=self.warnings,
            processed_files=self.processed_files,
            skipped_files=self.skipped_files,
        )
        report_path.write_text(report_text, encoding="utf-8")
        
        pdf_exported = False
        pdf_engine = os.getenv("QMS_PDF_ENGINE", "latex").strip().lower() or "latex"
        try:
            if pdf_engine == "reportlab":
                export_markdown_file_to_pdf(report_path, pdf_path)
            else:
                latex_result = export_markdown_file_to_pdf_latex(report_path, pdf_path)
                if latex_result.mode == "plain":
                    reason = latex_result.fallback_reason or "增强样式导出失败"
                    fallback_msg = f"PDF已降级为基础LaTeX样式（未应用pandoc_header.tex）: {reason}"
                    self.warnings.append(fallback_msg)
                    print(f"[EXPORT] {fallback_msg}", file=sys.stderr, flush=True)
            pdf_exported = True
        except Exception as exc:
            try:
                export_markdown_file_to_pdf(report_path, pdf_path)
                pdf_exported = True
                fallback_msg = f"PDF导出已回退到reportlab: {exc}"
                self.warnings.append(fallback_msg)
                print(f"[EXPORT] {fallback_msg}", file=sys.stderr, flush=True)
            except Exception as fallback_exc:
                self.warnings.append(f"PDF导出失败: {exc}; 回退失败: {fallback_exc}")
                print(f"[EXPORT] PDF导出失败: {exc}; 回退失败: {fallback_exc}", file=sys.stderr, flush=True)

        detail_payload = {
            "report_date": self.report_date.isoformat(),
            "config": str(self.config_path),
            "processed_files": self.processed_files,
            "skipped_files": self.skipped_files,
            "warnings": self.warnings,
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

def main() -> int:
    args = parse_args()
    orchestrator = ReportOrchestrator(args)
    return orchestrator.run()
