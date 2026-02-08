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
PERSON_SUMMARY_MAX_CHARS = 30


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


def _parse_named_summary_map(payload: Any) -> dict[str, str]:
    result: dict[str, str] = {}

    if isinstance(payload, dict):
        for k, v in payload.items():
            name = str(k or "").strip()
            summary = _normalize_person_summary(v)
            if name and summary:
                result[name] = summary
        return result

    if not isinstance(payload, list):
        return result

    for item in payload:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "") or "").strip()
        summary = _normalize_person_summary(item.get("summary", ""))
        if name and summary:
            result[name] = summary
    return result


def _normalize_person_summary(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    if len(text) > PERSON_SUMMARY_MAX_CHARS:
        text = text[:PERSON_SUMMARY_MAX_CHARS].rstrip()
    return text


def _merge_rank_summaries(stats: dict[str, Any], rank_key: str, summaries: dict[str, str]) -> None:
    rank_rows = stats.get(rank_key, [])
    if not isinstance(rank_rows, list):
        return
    for row in rank_rows:
        if not isinstance(row, dict):
            continue
        name = str(row.get("name", "") or "").strip()
        if not name:
            row.setdefault("summary", "")
            continue
        summary = str(summaries.get(name, "") or "").strip()
        if summary:
            row["summary"] = summary
        else:
            row.setdefault("summary", "")


def _extract_top20_names(local_stats: dict[str, Any], key: str) -> set[str]:
    rows = local_stats.get(key, [])
    if not isinstance(rows, list):
        return set()
    names: set[str] = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        name = str(row.get("name", "") or "").strip()
        if name:
            names.add(name)
    return names


def _filter_summaries_by_names(summaries: dict[str, str], allowed_names: set[str]) -> dict[str, str]:
    if not allowed_names:
        return {}
    return {name: text for name, text in summaries.items() if name in allowed_names}


def _request_llm_json(
    topic: str,
    stage: str,
    payload: dict[str, Any],
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
            "content": json.dumps(payload, ensure_ascii=False),
        },
    ]

    completion = None

    def _request_once(use_response_format: bool):
        kwargs: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "temperature": 0.3,
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
                print(f"[LLM] 主题[{topic}] {stage}调用中，已等待 {waited}s ...", file=sys.stderr, flush=True)

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
    return extract_json_object(content)


def call_llm_topic_summary(
    topic: str,
    report_date: date,
    local_stats: dict[str, Any],
    overdue_records: list[dict[str, Any]],
    base_url: str,
    model: str,
    api_key: str,
    timeout_seconds: int,
    progress_interval_seconds: int = 15,
) -> str:
    payload = {
        "task": "根据超期项目统计生成主题分析总结",
        "report_date": report_date.isoformat(),
        "topic": topic,
        "overdue_records": strip_source_fields(overdue_records),
        "input_stats": strip_source_fields(local_stats),
        "output_schema": {"summary": "string"},
        "requirements": [
            "仅输出summary字段",
            "简要总结需覆盖超期项目的总体态势、主要风险、重点关注项",
            "基于超期项目明细进行分析，聚焦问题根源和改进方向",
            "优先以input_stats中的超期统计值作为结论依据",
            "允许分析原因并给出建议",
        ],
    }
    result = _request_llm_json(
        topic=topic,
        stage="主题总结",
        payload=payload,
        base_url=base_url,
        model=model,
        api_key=api_key,
        timeout_seconds=timeout_seconds,
        progress_interval_seconds=progress_interval_seconds,
    )
    return str(result.get("summary", "") or "").strip()


def call_llm_person_summaries(
    topic: str,
    report_date: date,
    local_stats: dict[str, Any],
    base_url: str,
    model: str,
    api_key: str,
    timeout_seconds: int,
    progress_interval_seconds: int = 15,
) -> dict[str, Any]:
    payload = {
        "task": "根据超期项目统计生成人员超期内容概括",
        "report_date": report_date.isoformat(),
        "topic": topic,
        "input_stats": strip_source_fields(local_stats),
        "output_schema": {
            "qa_top20_summaries": [{"name": "string", "summary": "string"}],
            "qa_manager_top20_summaries": [{"name": "string", "summary": "string"}],
        },
        "requirements": [
            "只输出qa_top20_summaries和qa_manager_top20_summaries，不要输出summary字段",
            "对input_stats.overdue_by_qa_top20中的每个人仅输出2～3句简短概括，如：主要为……",
            "对input_stats.overdue_by_qa_manager_top20中的每个人仅输出2～3句简短概括，如：主要为……",
            "个人概括必须基于overdue_items中的content字段提炼，不要只按module字段名称概括",
            "若content有值，个人概括中不要直接罗列“变更/偏差/OOS/OOT/投诉”等模块名作为主体",
            "个人概括不要分析原因，不要提出建议，不要出现“建议”“需”“应”等措辞",
            "每条个人概括尽量控制在30个汉字以内",
            "如果对应列表为空，输出空数组",
        ],
    }
    result = _request_llm_json(
        topic=topic,
        stage="人员概括",
        payload=payload,
        base_url=base_url,
        model=model,
        api_key=api_key,
        timeout_seconds=timeout_seconds,
        progress_interval_seconds=progress_interval_seconds,
    )

    merged = dict(local_stats)

    qa_top20_names = _extract_top20_names(local_stats, "overdue_by_qa_top20")
    qa_manager_top20_names = _extract_top20_names(local_stats, "overdue_by_qa_manager_top20")
    qa_summaries = _filter_summaries_by_names(
        _parse_named_summary_map(result.get("qa_top20_summaries")),
        qa_top20_names,
    )
    qa_manager_summaries = _filter_summaries_by_names(
        _parse_named_summary_map(result.get("qa_manager_top20_summaries")),
        qa_manager_top20_names,
    )
    _merge_rank_summaries(merged, "overdue_by_qa", qa_summaries)
    _merge_rank_summaries(merged, "overdue_by_qa_manager", qa_manager_summaries)
    return merged
