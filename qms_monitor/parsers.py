from __future__ import annotations

import re
from datetime import date, datetime, timedelta


def parse_tabular_text(text: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for raw_line in text.splitlines():
        rows.append([cell.strip() for cell in raw_line.split("\t")])
    return rows


def get_cell(row: list[str], idx: int | None) -> str:
    if idx is None:
        return ""
    if idx < 0 or idx >= len(row):
        return ""
    return row[idx].strip()


def col_to_index(raw: str) -> int | None:
    value = (raw or "").strip()
    if not value:
        return None

    if value.isdigit():
        num = int(value)
        return num - 1 if num > 0 else None

    letters = "".join(ch for ch in value.upper() if "A" <= ch <= "Z")
    if not letters:
        return None

    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def parse_year(raw: str) -> str:
    value = (raw or "").strip()
    if not value:
        return "未知"
    if re.fullmatch(r"\d+(\.0+)?", value):
        return str(int(float(value)))
    return value


def parse_date_cell(raw: str) -> date | None:
    value = (raw or "").strip()
    if not value:
        return None

    num_match = re.fullmatch(r"\d+(\.\d+)?", value)
    if num_match:
        serial = float(value)
        if serial > 0:
            origin = datetime(1899, 12, 30)
            return (origin + timedelta(days=serial)).date()

    normalized = (
        value.replace("年", "-")
        .replace("月", "-")
        .replace("日", "")
        .replace("/", "-")
        .replace(".", "-")
    )
    normalized = re.sub(r"\s+", " ", normalized)

    candidates = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H",
        "%Y-%m",
    ]
    for fmt in candidates:
        try:
            dt = datetime.strptime(normalized, fmt)
            if fmt == "%Y-%m":
                return date(dt.year, dt.month, 1)
            return dt.date()
        except ValueError:
            continue

    match = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", normalized)
    if match:
        y, mon, day = map(int, match.groups())
        try:
            return date(y, mon, day)
        except ValueError:
            return None
    return None


def is_leap(year: int) -> bool:
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)


def add_one_month(d: date) -> date:
    if d.month == 12:
        target_year = d.year + 1
        target_month = 1
    else:
        target_year = d.year
        target_month = d.month + 1

    month_days = [
        31,
        29 if is_leap(target_year) else 28,
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ]
    target_day = min(d.day, month_days[target_month - 1])
    return date(target_year, target_month, target_day)


def normalize_sheet_name(raw: str) -> str | int:
    value = (raw or "").strip()
    if not value:
        return 1
    if re.fullmatch(r"\d+", value):
        return int(value)
    return value
