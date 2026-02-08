from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Iterable


def _weighted_text_len(text: str) -> int:
    length = 0
    for ch in text:
        length += 2 if ord(ch) > 127 else 1
    return max(length, 1)


def _adaptive_table_total_width(
    plain_rows: list[list[str]],
    usable_width: float,
    *,
    min_ratio: float,
    max_ratio: float,
) -> float:
    if not plain_rows:
        return usable_width * min_ratio

    col_count = max(len(row) for row in plain_rows)
    col_max_units: list[int] = [1] * col_count
    for idx in range(col_count):
        for row in plain_rows:
            cell_text = (row[idx] if idx < len(row) else "").strip()
            col_max_units[idx] = max(col_max_units[idx], min(_weighted_text_len(cell_text), 40))

    # Estimate natural table width from content units and per-column paddings.
    natural_width = sum(col_max_units) * 4.6 + col_count * 14
    min_width = usable_width * min_ratio
    max_width = usable_width * max_ratio
    return max(min_width, min(natural_width, max_width))


def _estimate_col_widths(plain_rows: list[list[str]], total_width: float) -> list[float]:
    if not plain_rows:
        return []

    col_count = max(len(row) for row in plain_rows)
    weights: list[float] = [1.0] * col_count
    for idx in range(col_count):
        col_max = 1
        for row in plain_rows:
            cell_text = (row[idx] if idx < len(row) else "").strip()
            col_max = max(col_max, min(_weighted_text_len(cell_text), 40))
        weights[idx] = float(col_max)

    total = sum(weights) or float(col_count)
    return [total_width * (w / total) for w in weights]


def _to_para_markup(nodes: Iterable[object], *, code_font: str) -> str:
    from bs4 import NavigableString, Tag

    parts: list[str] = []
    for node in nodes:
        if isinstance(node, NavigableString):
            parts.append(escape(str(node)))
            continue

        if not isinstance(node, Tag):
            continue

        inner = _to_para_markup(node.contents, code_font=code_font)
        tag_name = node.name.lower()

        if tag_name == "br":
            parts.append("<br/>")
            continue
        if tag_name in {"strong", "b"}:
            parts.append(f"<b>{inner}</b>")
            continue
        if tag_name in {"em", "i"}:
            parts.append(f"<i>{inner}</i>")
            continue
        if tag_name == "code":
            code_text = escape(node.get_text())
            parts.append(f"<font name='{code_font}'>{code_text}</font>")
            continue
        if tag_name == "a":
            href = escape(node.get("href", ""), quote=True)
            if href:
                parts.append(f"<a href='{href}'>{inner or href}</a>")
            else:
                parts.append(inner)
            continue

        parts.append(inner)

    return "".join(parts).strip()


def export_markdown_text_to_pdf(markdown_text: str, output_path: Path) -> None:
    try:
        import markdown
        from bs4 import BeautifulSoup, NavigableString, Tag
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import (
            HRFlowable,
            ListFlowable,
            ListItem,
            Paragraph,
            Preformatted,
            SimpleDocTemplate,
            Spacer,
            Table,
            TableStyle,
        )
    except ModuleNotFoundError as exc:
        raise RuntimeError("缺少依赖 markdown/beautifulsoup4/reportlab，请先运行 `uv sync`") from exc

    preferred_font = "STSong-Light"
    fallback_font = "Helvetica"
    code_font = "Courier"

    font_name = fallback_font
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(preferred_font))
        font_name = preferred_font
    except Exception:
        font_name = fallback_font

    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        "QMSBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10.5,
        leading=15,
        spaceBefore=2,
        spaceAfter=6,
    )
    heading_styles = {
        "h1": ParagraphStyle(
            "QMSH1",
            parent=styles["Heading1"],
            fontName=font_name,
            fontSize=19,
            leading=24,
            spaceBefore=8,
            spaceAfter=10,
        ),
        "h2": ParagraphStyle(
            "QMSH2",
            parent=styles["Heading2"],
            fontName=font_name,
            fontSize=16,
            leading=20,
            spaceBefore=8,
            spaceAfter=8,
        ),
        "h3": ParagraphStyle(
            "QMSH3",
            parent=styles["Heading3"],
            fontName=font_name,
            fontSize=13.5,
            leading=17,
            spaceBefore=6,
            spaceAfter=6,
        ),
    }
    quote_style = ParagraphStyle(
        "QMSQuote",
        parent=body_style,
        leftIndent=14,
        textColor=colors.HexColor("#374151"),
    )
    code_block_style = ParagraphStyle(
        "QMSCode",
        parent=styles["Code"],
        fontName=code_font,
        fontSize=9,
        leading=12,
        leftIndent=8,
        rightIndent=8,
        spaceBefore=4,
        spaceAfter=8,
    )
    table_cell_style = ParagraphStyle(
        "QMSTableCell",
        parent=body_style,
        fontSize=9.5,
        leading=13,
        spaceAfter=0,
        spaceBefore=0,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    html = markdown.markdown(
        markdown_text,
        extensions=["tables", "fenced_code", "sane_lists", "nl2br"],
        output_format="html5",
    )
    soup = BeautifulSoup(html, "html.parser")

    story: list[object] = []
    heading_counters = {"h1": 0, "h2": 0, "h3": 0}

    for block in soup.contents:
        if isinstance(block, NavigableString):
            text = str(block).strip()
            if text:
                story.append(Paragraph(escape(text), body_style))
            continue

        if not isinstance(block, Tag):
            continue

        tag_name = block.name.lower()

        if tag_name in {"h1", "h2", "h3"}:
            if tag_name == "h1":
                heading_counters["h1"] += 1
                heading_counters["h2"] = 0
                heading_counters["h3"] = 0
                number_prefix = f"{heading_counters['h1']}、"
            elif tag_name == "h2":
                if heading_counters["h1"] == 0:
                    heading_counters["h1"] = 1
                heading_counters["h2"] += 1
                heading_counters["h3"] = 0
                number_prefix = f"{heading_counters['h1']}.{heading_counters['h2']}"
            else:
                if heading_counters["h1"] == 0:
                    heading_counters["h1"] = 1
                if heading_counters["h2"] == 0:
                    heading_counters["h2"] = 1
                heading_counters["h3"] += 1
                number_prefix = (
                    f"{heading_counters['h1']}."
                    f"{heading_counters['h2']}."
                    f"{heading_counters['h3']}"
                )

            heading_markup = _to_para_markup(block.contents, code_font=code_font)
            story.append(Paragraph(f"{number_prefix} {heading_markup}", heading_styles[tag_name]))
            continue

        if tag_name == "p":
            story.append(Paragraph(_to_para_markup(block.contents, code_font=code_font), body_style))
            continue

        if tag_name in {"ul", "ol"}:
            is_ordered = tag_name == "ol"
            list_items: list[ListItem] = []
            for li in block.find_all("li", recursive=False):
                inline_nodes = [
                    node
                    for node in li.contents
                    if not isinstance(node, Tag) or node.name not in {"ul", "ol"}
                ]
                item_markup = _to_para_markup(inline_nodes, code_font=code_font) or " "
                list_items.append(ListItem(Paragraph(item_markup, body_style)))

                for nested in li.find_all(["ul", "ol"], recursive=False):
                    nested_items = []
                    for nested_li in nested.find_all("li", recursive=False):
                        nested_markup = _to_para_markup(nested_li.contents, code_font=code_font) or " "
                        nested_items.append(ListItem(Paragraph(nested_markup, body_style)))
                    if nested_items:
                        nested_kwargs = {
                            "bulletType": "1" if nested.name == "ol" else "bullet",
                            "leftIndent": 16,
                        }
                        if nested.name == "ol":
                            nested_kwargs["start"] = "1"
                        list_items.append(
                            ListItem(
                                ListFlowable(
                                    nested_items,
                                    **nested_kwargs,
                                )
                            )
                        )

            if list_items:
                list_kwargs = {
                    "bulletType": "1" if is_ordered else "bullet",
                    "leftIndent": 12,
                }
                if is_ordered:
                    list_kwargs["start"] = "1"
                story.append(
                    ListFlowable(
                        list_items,
                        **list_kwargs,
                    )
                )
                story.append(Spacer(1, 4))
            continue

        if tag_name == "table":
            rows: list[list[Paragraph]] = []
            plain_rows: list[list[str]] = []
            header_count = 0
            has_summary_column = False
            for tr in block.find_all("tr", recursive=True):
                row: list[Paragraph] = []
                is_header_row = tr.find_parent("thead") is not None
                plain_cells: list[str] = []

                for cell in tr.find_all(["th", "td"], recursive=False):
                    plain_cells.append(cell.get_text(" ", strip=True))
                    cell_markup = _to_para_markup(cell.contents, code_font=code_font)
                    if cell.name.lower() == "th":
                        cell_markup = f"<b>{cell_markup}</b>"
                        is_header_row = True
                    row.append(Paragraph(cell_markup or " ", table_cell_style))

                if row:
                    rows.append(row)
                    plain_rows.append(plain_cells)
                    if is_header_row:
                        header_count += 1
                        if any((text or "").strip() == "超期内容概括" for text in plain_cells):
                            has_summary_column = True

            if rows:
                col_count = max(len(row) for row in rows)
                normalized_rows: list[list[Paragraph]] = []
                normalized_plain_rows: list[list[str]] = []
                for row, plain_row in zip(rows, plain_rows):
                    if len(row) < col_count:
                        row = row + [Paragraph(" ", table_cell_style)] * (col_count - len(row))
                    if len(plain_row) < col_count:
                        plain_row = plain_row + [""] * (col_count - len(plain_row))
                    normalized_rows.append(row)
                    normalized_plain_rows.append(plain_row)

                if has_summary_column and col_count == 3:
                    usable_width = A4[0] - 80
                    table_total_width = _adaptive_table_total_width(
                        normalized_plain_rows,
                        usable_width,
                        min_ratio=0.62,
                        max_ratio=0.86,
                    )
                    table = Table(
                        normalized_rows,
                        colWidths=[
                            table_total_width * (10 / 72),
                            table_total_width * (7 / 72),
                            table_total_width * (55 / 72),
                        ],
                        repeatRows=header_count if header_count > 0 else 0,
                        hAlign="CENTER",
                    )
                else:
                    usable_width = A4[0] - 80
                    min_ratio = min(0.78, 0.34 + 0.09 * col_count)
                    max_ratio = 0.90 if col_count <= 3 else 0.95
                    table_total_width = _adaptive_table_total_width(
                        normalized_plain_rows,
                        usable_width,
                        min_ratio=min_ratio,
                        max_ratio=max_ratio,
                    )
                    col_widths = _estimate_col_widths(normalized_plain_rows, table_total_width)
                    table = Table(
                        normalized_rows,
                        colWidths=col_widths if col_widths else None,
                        repeatRows=header_count if header_count > 0 else 0,
                        hAlign="CENTER",
                    )
                table.setStyle(
                    TableStyle(
                        [
                            ("FONTNAME", (0, 0), (-1, -1), font_name),
                            ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                            ("LEADING", (0, 0), (-1, -1), 12),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c5cbd3")),
                            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#98a2b3")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    )
                )
                if header_count > 0:
                    table.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, header_count - 1), colors.HexColor("#eef2f7")),
                                ("TEXTCOLOR", (0, 0), (-1, header_count - 1), colors.HexColor("#111827")),
                            ]
                        )
                    )
                story.append(table)
                story.append(Spacer(1, 8))
            continue

        if tag_name == "pre":
            code_tag = block.find("code")
            code_text = code_tag.get_text() if code_tag else block.get_text()
            story.append(Preformatted(code_text, code_block_style))
            continue

        if tag_name == "blockquote":
            quote_markup = _to_para_markup(block.contents, code_font=code_font) or escape(block.get_text())
            story.append(Paragraph(quote_markup, quote_style))
            story.append(Spacer(1, 4))
            continue

        if tag_name in {"hr"}:
            story.append(
                HRFlowable(
                    thickness=0.7,
                    color=colors.HexColor("#98a2b3"),
                    spaceBefore=4,
                    spaceAfter=8,
                )
            )
            continue

        fallback_text = block.get_text(" ", strip=True)
        if fallback_text:
            story.append(Paragraph(escape(fallback_text), body_style))

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40,
        title="QMS Report",
    )

    def _draw_page_footer(canvas_obj, doc_obj) -> None:
        canvas_obj.saveState()
        canvas_obj.setFont(font_name, 9)
        canvas_obj.setFillColor(colors.HexColor("#6b7280"))
        page_text = str(canvas_obj.getPageNumber())
        canvas_obj.drawCentredString(doc_obj.pagesize[0] / 2, 18, page_text)
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=_draw_page_footer, onLaterPages=_draw_page_footer)


def export_markdown_file_to_pdf(markdown_path: Path, output_path: Path) -> None:
    markdown_text = markdown_path.read_text(encoding="utf-8")
    export_markdown_text_to_pdf(markdown_text, output_path)
