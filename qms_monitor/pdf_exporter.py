from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Iterable


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
            story.append(Paragraph(_to_para_markup(block.contents, code_font=code_font), heading_styles[tag_name]))
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
                        list_items.append(
                            ListItem(
                                ListFlowable(
                                    nested_items,
                                    bulletType="1" if nested.name == "ol" else "bullet",
                                    start="1",
                                    leftIndent=16,
                                )
                            )
                        )

            if list_items:
                story.append(
                    ListFlowable(
                        list_items,
                        bulletType="1" if is_ordered else "bullet",
                        start="1",
                        leftIndent=12,
                    )
                )
                story.append(Spacer(1, 4))
            continue

        if tag_name == "table":
            rows: list[list[Paragraph]] = []
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
                    if is_header_row:
                        header_count += 1
                        if any((text or "").strip() == "超期内容概括" for text in plain_cells):
                            has_summary_column = True

            if rows:
                col_count = max(len(row) for row in rows)
                normalized_rows: list[list[Paragraph]] = []
                for row in rows:
                    if len(row) < col_count:
                        row = row + [Paragraph(" ", table_cell_style)] * (col_count - len(row))
                    normalized_rows.append(row)

                if has_summary_column and col_count == 3:
                    usable_width = A4[0] - 80
                    table_total_width = usable_width * 0.72
                    table = Table(
                        normalized_rows,
                        colWidths=[
                            table_total_width * (0.10 / 0.72),
                            table_total_width * (0.07 / 0.72),
                            table_total_width * (0.55 / 0.72),
                        ],
                        repeatRows=header_count if header_count > 0 else 0,
                        hAlign="CENTER",
                    )
                else:
                    table = Table(normalized_rows, repeatRows=header_count if header_count > 0 else 0)
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
    doc.build(story)


def export_markdown_file_to_pdf(markdown_path: Path, output_path: Path) -> None:
    markdown_text = markdown_path.read_text(encoding="utf-8")
    export_markdown_text_to_pdf(markdown_text, output_path)
