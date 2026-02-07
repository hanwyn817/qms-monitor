from __future__ import annotations

import argparse
import os
from datetime import date
from pathlib import Path

from .constants import ENV_FILE_DEFAULT


def load_env_file(path: str | Path = ENV_FILE_DEFAULT) -> None:
    env_path = Path(path)
    if not env_path.exists() or not env_path.is_file():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        if not key:
            continue

        value = value.strip()
        if (value.startswith('"') and value.endswith('"')) or (
            value.startswith("'") and value.endswith("'")
        ):
            value = value[1:-1]

        os.environ.setdefault(key, value)


def parse_args() -> argparse.Namespace:
    load_env_file()

    parser = argparse.ArgumentParser(description="QMS monitor report generator")
    parser.add_argument("--config", default="config.xlsx", help="配置文件路径")
    parser.add_argument("--output-dir", default="outputs", help="报告输出目录")
    parser.add_argument(
        "--input-mode",
        choices=["excel", "csv"],
        default=os.getenv("QMS_INPUT_MODE", "excel"),
        help="数据输入模式：excel(默认) 或 csv",
    )
    parser.add_argument(
        "--csv-manifest",
        default=os.getenv("QMS_CSV_MANIFEST", ""),
        help="CSV模式下使用的manifest.json路径",
    )
    parser.add_argument(
        "--report-date",
        default=date.today().isoformat(),
        help="统计基准日期，格式 YYYY-MM-DD，默认当天",
    )
    parser.add_argument(
        "--skip-llm",
        action="store_true",
        help="跳过LLM调用，仅使用本地统计",
    )
    parser.add_argument(
        "--llm-base-url",
        default=os.getenv("QMS_LLM_BASE_URL", "https://api.openai.com/v1"),
        help="OpenAI兼容接口Base URL",
    )
    parser.add_argument(
        "--llm-model",
        default=os.getenv("QMS_LLM_MODEL", ""),
        help="模型名称，默认为环境变量QMS_LLM_MODEL",
    )
    parser.add_argument(
        "--llm-api-key",
        default=os.getenv("QMS_LLM_API_KEY", ""),
        help="API Key，默认为环境变量QMS_LLM_API_KEY",
    )
    parser.add_argument(
        "--llm-timeout",
        type=int,
        default=int(os.getenv("QMS_LLM_TIMEOUT", "120")),
        help="LLM请求超时时间（秒）",
    )
    return parser.parse_args()
