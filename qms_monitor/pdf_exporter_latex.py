from __future__ import annotations

import os
import subprocess
from dataclasses import dataclass
from pathlib import Path


def _list_available_fonts() -> set[str]:
    try:
        proc = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        return set()

    if proc.returncode != 0:
        return set()

    names: set[str] = set()
    for raw_line in proc.stdout.splitlines():
        for part in raw_line.split(","):
            name = part.strip()
            if name:
                names.add(name)
    return names


def _pick_existing_font(available_fonts: set[str], *candidates: str) -> str:
    # If font discovery is unavailable, still allow user-provided font name.
    if not available_fonts:
        for name in candidates:
            if name.strip():
                return name.strip()
        return ""

    for name in candidates:
        normalized = name.strip()
        if normalized and normalized in available_fonts:
            return normalized
    return ""


@dataclass(frozen=True)
class LatexExportResult:
    mode: str
    fallback_reason: str = ""


def _compact_error_text(stderr: str, stdout: str, *, limit: int = 300) -> str:
    text = (stderr or stdout or "").strip()
    if not text:
        return ""
    single_line = " ".join(text.split())
    if len(single_line) <= limit:
        return single_line
    return single_line[: limit - 3] + "..."


def export_markdown_file_to_pdf_latex(markdown_path: Path, output_path: Path) -> LatexExportResult:
    if not markdown_path.exists():
        raise RuntimeError(f"Markdown文件不存在: {markdown_path}")

    header_path = Path(__file__).resolve().parent / "resources" / "pandoc_header.tex"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    available_fonts = _list_available_fonts()

    mainfont = _pick_existing_font(
        available_fonts,
        os.getenv("QMS_LATEX_MAINFONT", "").strip(),
        "PingFang SC",
        "Songti SC",
        "Heiti SC",
        "Arial Unicode MS",
    )
    sansfont = _pick_existing_font(
        available_fonts,
        os.getenv("QMS_LATEX_SANSFONT", "").strip(),
        "Helvetica Neue",
        "PingFang SC",
        "Helvetica",
    )
    monofont = _pick_existing_font(
        available_fonts,
        os.getenv("QMS_LATEX_MONOFONT", "").strip(),
        "Menlo",
        "Courier New",
        "Courier",
    )
    base_cmd = [
        "pandoc",
        str(markdown_path),
        "--from",
        "gfm+pipe_tables+task_lists+smart",
        "--syntax-highlighting=none",
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
    ]
    if mainfont:
        base_cmd.append("--variable=mainfont:" + mainfont)
    if sansfont:
        base_cmd.append("--variable=sansfont:" + sansfont)
    if monofont:
        base_cmd.append("--variable=monofont:" + monofont)
    styled_cmd = [*base_cmd, "--include-in-header", str(header_path)]

    proc = subprocess.run(styled_cmd, capture_output=True, text=True)
    if proc.returncode == 0:
        return LatexExportResult(mode="styled")

    fallback_proc = subprocess.run(base_cmd, capture_output=True, text=True)
    if fallback_proc.returncode == 0:
        reason = _compact_error_text(proc.stderr, proc.stdout)
        return LatexExportResult(mode="plain", fallback_reason=reason)

    detail = (fallback_proc.stderr or fallback_proc.stdout or proc.stderr or proc.stdout or "").strip()
    if detail:
        raise RuntimeError(f"pandoc/xelatex 导出失败: {detail}")
    raise RuntimeError("pandoc/xelatex 导出失败")
