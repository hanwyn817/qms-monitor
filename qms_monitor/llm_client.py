from __future__ import annotations

import json
import re
from datetime import date
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


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
) -> dict[str, Any]:
    if not model or not api_key:
        raise RuntimeError("缺少LLM配置: model/api_key")

    url = base_url.rstrip("/") + "/chat/completions"
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
                    "records": event_records,
                    "input_stats": local_stats,
                    "output_schema": {
                        "module": "string",
                        "yearly_totals": [{"year": "string", "count": "number"}],
                        "overdue": {
                            "count": "number",
                            "ratio": "number",
                            "items": [
                                {
                                    "year": "string",
                                    "event_id": "string",
                                    "content": "string",
                                    "planned_date": "YYYY-MM-DD",
                                    "status": "string",
                                    "qa": "string",
                                    "qa_manager": "string",
                                    "source": "string",
                                }
                            ],
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

    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }

    req = Request(
        url=url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )

    try:
        with urlopen(req, timeout=timeout_seconds) as resp:
            raw = resp.read().decode("utf-8")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"网络错误: {exc}") from exc

    parsed = json.loads(raw)
    content = parsed["choices"][0]["message"]["content"]
    result = extract_json_object(content)

    result.setdefault("module", module)
    result.setdefault("yearly_totals", local_stats.get("yearly_totals", []))
    result.setdefault("overdue", local_stats.get("overdue", {}))
    result.setdefault("overdue_by_qa", local_stats.get("overdue_by_qa", []))
    result.setdefault("overdue_by_qa_manager", local_stats.get("overdue_by_qa_manager", []))
    return result
