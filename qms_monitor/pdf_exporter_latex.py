from __future__ import annotations

import os
import subprocess
from pathlib import Path


def _pick_existing_font(*candidates: str) -> str:
    for name in candidates:
        if name.strip():
            return name
    return ""


def export_markdown_file_to_pdf_latex(markdown_path: Path, output_path: Path) -> None:
    if not markdown_path.exists():
        raise RuntimeError(f"Markdown文件不存在: {markdown_path}")

    header_path = Path(__file__).resolve().parent / "resources" / "pandoc_header.tex"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    mainfont = _pick_existing_font(
        os.getenv("QMS_LATEX_MAINFONT", "").strip(),
        "PingFang SC",
        "Songti SC",
    )
    sansfont = _pick_existing_font(
        os.getenv("QMS_LATEX_SANSFONT", "").strip(),
        "Helvetica Neue",
        "PingFang SC",
    )
    monofont = _pick_existing_font(
        os.getenv("QMS_LATEX_MONOFONT", "").strip(),
        "Menlo",
        "Courier New",
    )
    base_cmd = [
        "pandoc",
        str(markdown_path),
        "--from",
        "gfm+pipe_tables+task_lists+smart",
        "--pdf-engine=xelatex",
        "--output",
        str(output_path),
        "--toc",
        "--toc-depth=2",
        "--number-sections",
        "--metadata=title:质量体系运行报告",
        "--variable=documentclass:article",
        "--variable=geometry:margin=22mm",
        "--variable=colorlinks:true",
        "--variable=linkcolor:blue",
        "--variable=urlcolor:blue",
        "--variable=mainfont:" + mainfont,
        "--variable=sansfont:" + sansfont,
        "--variable=monofont:" + monofont,
    ]
    styled_cmd = [*base_cmd, "--include-in-header", str(header_path)]

    proc = subprocess.run(styled_cmd, capture_output=True, text=True)
    if proc.returncode == 0:
        return

    fallback_proc = subprocess.run(base_cmd, capture_output=True, text=True)
    if fallback_proc.returncode == 0:
        return

    detail = (fallback_proc.stderr or fallback_proc.stdout or proc.stderr or proc.stdout or "").strip()
    if detail:
        raise RuntimeError(f"pandoc/xelatex 导出失败: {detail}")
    raise RuntimeError("pandoc/xelatex 导出失败")
