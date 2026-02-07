from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile


HEADERS = [
    "主题",
    "质量模块",
    "年份",
    "编号",
    "内容",
    "发起日期",
    "计划完成日期",
    "状态",
    "责任部门",
    "责任人",
    "分管QA",
    "分管QA中层",
]


def _a1_col(col_index: int) -> str:
    if col_index <= 0:
        return "A"
    chars = ""
    n = col_index
    while n:
        n, remainder = divmod(n - 1, 26)
        chars = chr(65 + remainder) + chars
    return chars


def _is_valid_xml_char(codepoint: int) -> bool:
    return (
        codepoint == 0x9
        or codepoint == 0xA
        or codepoint == 0xD
        or 0x20 <= codepoint <= 0xD7FF
        or 0xE000 <= codepoint <= 0xFFFD
        or 0x10000 <= codepoint <= 0x10FFFF
    )


def _sanitize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    return "".join(ch for ch in text if _is_valid_xml_char(ord(ch)))


def _escape_xml_text(value: Any) -> str:
    text = _sanitize_text(value)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _cell_xml(row_no: int, col_no: int, value: Any) -> str:
    ref = f"{_a1_col(col_no)}{row_no}"
    text = _sanitize_text(value)
    escaped = _escape_xml_text(text)
    preserve = ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") else ""
    return f'<c r="{ref}" t="inlineStr"><is><t{preserve}>{escaped}</t></is></c>'


def _collect_overdue_rows(module_results: dict[str, dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for module in sorted(module_results.keys()):
        module_data = module_results.get(module, {})
        if not isinstance(module_data, dict):
            continue
        overdue = module_data.get("overdue", {})
        if not isinstance(overdue, dict):
            continue
        items = overdue.get("items", [])
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            rows.append(
                {
                    "主题": str(item.get("topic", "") or ""),
                    "质量模块": module,
                    "年份": str(item.get("year", "") or ""),
                    "编号": str(item.get("event_id", "") or ""),
                    "内容": str(item.get("content", "") or ""),
                    "发起日期": str(item.get("initiated_date", "") or ""),
                    "计划完成日期": str(item.get("planned_date", "") or ""),
                    "状态": str(item.get("status", "") or ""),
                    "责任部门": str(item.get("owner_dept", "") or ""),
                    "责任人": str(item.get("owner", "") or ""),
                    "分管QA": str(item.get("qa", "") or ""),
                    "分管QA中层": str(item.get("qa_manager", "") or ""),
                }
            )
    rows.sort(key=lambda row: (row["质量模块"], row["计划完成日期"], row["编号"]))
    return rows


def _build_sheet_xml(rows: list[dict[str, str]]) -> str:
    xml_parts: list[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]

    row_no = 1
    header_cells = "".join(_cell_xml(row_no, col_no, header) for col_no, header in enumerate(HEADERS, start=1))
    xml_parts.append(f'<row r="{row_no}">{header_cells}</row>')

    for item in rows:
        row_no += 1
        row_cells = "".join(_cell_xml(row_no, col_no, item.get(header, "")) for col_no, header in enumerate(HEADERS, start=1))
        xml_parts.append(f'<row r="{row_no}">{row_cells}</row>')

    xml_parts.extend(["</sheetData>", "</worksheet>"])
    return "".join(xml_parts)


def export_overdue_events_excel(path: Path, module_results: dict[str, dict[str, Any]]) -> int:
    rows = _collect_overdue_rows(module_results)
    sheet_xml = _build_sheet_xml(rows)

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="超期事件" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, mode="w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return len(rows)
