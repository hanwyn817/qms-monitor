from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass
class LedgerConfig:
    row_no: int
    topic: str
    module: str
    year: str
    file_path: str
    sheet_name: str
    id_col: int
    content_col: int
    initiated_col: int
    planned_col: int | None = None
    status_col: int | None = None
    owner_dept_col: int | None = None
    owner_col: int | None = None
    qa_col: int | None = None
    qa_manager_col: int | None = None
    open_status_value: str = ""
    data_start_row: int = 2


@dataclass
class QmsEvent:
    topic: str
    module: str
    year: str
    event_id: str
    content: str
    initiated_date: date | None
    planned_date: date | None
    status: str
    owner_dept: str
    owner: str
    qa: str
    qa_manager: str
    source_file: str
    source_sheet: str
    row_index: int

    @property
    def initiated_date_str(self) -> str:
        return self.initiated_date.isoformat() if self.initiated_date else ""

    @property
    def planned_date_str(self) -> str:
        return self.planned_date.isoformat() if self.planned_date else ""
