from __future__ import annotations

import json
import re
import sys
import threading
import time
from datetime import date
from typing import Any

from openai import OpenAI


SOURCE_KEYS = {"source", "source_file", "source_sheet", "source_row"}


def strip_source_fields(payload: Any) -> Any:
    if isinstance(payload, dict):
        return {k: strip_source_fields(v) for k, v in payload.items() if k not in SOURCE_KEYS}
    if isinstance(payload, list):
        return [strip_source_fields(v) for v in payload]
    return payload


def extract_json_object(text: str) -> dict[str, Any]:
    content = text.strip()
    if content.startswith("```"):
        content = re.sub(r"^```[a-zA-Z]*\n", "", content)
        content = re.sub(r"\n```$", "", content)

    try:
        parsed = json.loads(content)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    start = content.find("{")
    end = content.rfind("}")
    if start >= 0 and end > start:
        candidate = content[start : end + 1]
        parsed = json.loads(candidate)
        if isinstance(parsed, dict):
            return parsed

    raise ValueError("LLM输出不是有效JSON对象")


def call_llm(
    module: str,
    report_date: date,
    local_stats: dict[str, Any],
    event_records: list[dict[str, Any]],
    base_url: str,
    model: str,
    api_key: str,
    timeout_seconds: int,
    progress_interval_seconds: int = 15,
) -> dict[str, Any]:
    if not model or not api_key:
        raise RuntimeError("缺少LLM配置: model/api_key")

    client = OpenAI(
        api_key=api_key,
        base_url=base_url.rstrip("/"),
        timeout=timeout_seconds,
    )

    messages = [
        {
            "role": "system",
            "content": (
                "你是药企质量管理体系分析助手。"
                "你必须严格输出JSON对象，不要输出任何额外文本。"
            ),
        },
        {
            "role": "user",
            "content": json.dumps(
                {
                    "task": "根据输入生成模块分析结果",
                    "report_date": report_date.isoformat(),
                    "module": module,
                    "records": strip_source_fields(event_records),
                    "input_stats": strip_source_fields(local_stats),
                    "output_schema": {
                        "module": "string",
                        "yearly_totals": [{"year": "string", "count": "number"}],
                        "overdue": {
                            "count": "number",
                            "ratio": "number",
                        },
                        "overdue_by_qa": [{"name": "string", "count": "number"}],
                        "overdue_by_qa_manager": [{"name": "string", "count": "number"}],
                        "summary": "string",
                    },
                    "requirements": [
                        "yearly_totals必须覆盖所有出现的年份",
                        "overdue.ratio使用百分比数值，例如12.5表示12.5%",
                        "overdue_by_qa和overdue_by_qa_manager按count降序",
                        "如果缺失分管QA中层数据，overdue_by_qa_manager可返回空数组",
                    ],
                },
                ensure_ascii=False,
            ),
        },
    ]

    completion = None

    def _request_once(use_response_format: bool):
        kwargs: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "temperature": 0.1,
        }
        if use_response_format:
            kwargs["response_format"] = {"type": "json_object"}
        return client.chat.completions.create(**kwargs)

    def _execute_with_heartbeat(use_response_format: bool):
        if progress_interval_seconds <= 0:
            return _request_once(use_response_format=use_response_format)

        start_ts = time.time()
        stop_event = threading.Event()

        def heartbeat() -> None:
            while not stop_event.wait(progress_interval_seconds):
                waited = int(time.time() - start_ts)
                print(f"[LLM] 模块[{module}] 调用中，已等待 {waited}s ...", file=sys.stderr, flush=True)

        thread = threading.Thread(target=heartbeat, daemon=True)
        thread.start()
        try:
            return _request_once(use_response_format=use_response_format)
        finally:
            stop_event.set()

    try:
        completion = _execute_with_heartbeat(use_response_format=True)
    except Exception as exc:
        if "response_format" in str(exc):
            try:
                completion = _execute_with_heartbeat(use_response_format=False)
            except Exception as fallback_exc:
                raise RuntimeError(f"LLM请求失败: {fallback_exc}") from fallback_exc
        else:
            raise RuntimeError(f"LLM请求失败: {exc}") from exc

    content = completion.choices[0].message.content or ""
    result = extract_json_object(content)

    result.setdefault("module", module)
    result.setdefault("yearly_totals", local_stats.get("yearly_totals", []))
    local_overdue = local_stats.get("overdue", {})
    result_overdue = result.get("overdue")
    if not isinstance(result_overdue, dict):
        result_overdue = {}
    result_overdue.setdefault("count", local_overdue.get("count", 0))
    result_overdue.setdefault("ratio", local_overdue.get("ratio", 0.0))
    # 超期清单始终以本地统计为准，不依赖LLM输出。
    result_overdue["items"] = local_overdue.get("items", [])
    result["overdue"] = result_overdue
    result.setdefault("overdue_by_qa", local_stats.get("overdue_by_qa", []))
    result.setdefault("overdue_by_qa_manager", local_stats.get("overdue_by_qa_manager", []))
    return result
