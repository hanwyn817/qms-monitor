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
    topic: str,
    report_date: date,
    local_stats: dict[str, Any],
    overdue_records: list[dict[str, Any]],
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
                    "task": "根据超期项目统计生成主题分析总结",
                    "report_date": report_date.isoformat(),
                    "topic": topic,
                    "overdue_records": strip_source_fields(overdue_records),
                    "input_stats": strip_source_fields(local_stats),
                    "output_schema": {"summary": "string"},
                    "requirements": [
                        "总结需覆盖超期项目的总体态势、主要风险、重点关注项",
                        "基于超期项目明细进行分析，聚焦问题根源和改进方向",
                        "优先以input_stats中的超期统计值作为结论依据",
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
                print(f"[LLM] 主题[{topic}] 调用中，已等待 {waited}s ...", file=sys.stderr, flush=True)

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
    merged = dict(local_stats)
    summary = str(result.get("summary", "")).strip()
    if summary:
        merged["summary"] = summary
    else:
        merged.setdefault("summary", "")
    return merged
