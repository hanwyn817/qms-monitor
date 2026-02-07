from __future__ import annotations

import platform
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Tuple, Union


# Excel constants
XL_BY_ROWS = 1
XL_BY_COLUMNS = 2
XL_PREVIOUS = 2
XL_FORMULAS = -4123
XL_VALUES = -4163
XL_CALC_MANUAL = -4135


def _get_com_modules() -> Tuple[Any, Any]:
    if platform.system() != "Windows":
        raise RuntimeError("This project requires Windows + Microsoft Excel for COM automation.")

    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except ModuleNotFoundError as exc:
        raise RuntimeError("pywin32 is required. Install with: pip install pywin32") from exc

    return pythoncom, win32


def _safe_str(value: Any) -> str:
    try:
        return "" if value is None else str(value)
    except Exception:
        return repr(value)


def _normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _a1_col(n: int) -> str:
    if n <= 0:
        return "A"
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _a1_addr(row: int, col: int) -> str:
    return f"{_a1_col(col)}{row}"


def _values_to_delimited_text(values: Any, sep: str = "\t", row_sep: str = "\n") -> str:
    if values is None:
        return ""
    if not isinstance(values, tuple):
        return _safe_str(values)

    lines: list[str] = []
    for row in values:
        if isinstance(row, tuple):
            lines.append(sep.join(_safe_str(c) for c in row))
        else:
            lines.append(_safe_str(row))
    return row_sep.join(lines)


def _normalize_excel_path(path: str) -> str:
    raw = Path(path).expanduser()
    try:
        return str(raw.resolve(strict=False))
    except Exception:
        return str(raw)


@dataclass
class OfficeReadResult:
    ok: bool
    app: str
    path: str
    elapsed_ms: int = 0

    text: str = ""
    char_count: int = 0
    preview: str = ""

    excel_last_row: Optional[int] = None
    excel_last_col: Optional[int] = None
    excel_sheet: Optional[Union[int, str]] = None
    excel_sheet_name: Optional[str] = None
    excel_range_a1: Optional[str] = None

    error_type: str = ""
    error_message: str = ""


class ExcelSession:
    def __init__(self, visible: bool = False):
        self.visible = visible
        self.excel = None
        self._pythoncom = None

    def __enter__(self):
        pythoncom, win32 = _get_com_modules()
        self._pythoncom = pythoncom
        pythoncom.CoInitialize()

        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = bool(self.visible)
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        self.excel.EnableEvents = False
        self.excel.ScreenUpdating = False
        try:
            self.excel.Calculation = XL_CALC_MANUAL
        except Exception:
            pass
        try:
            self.excel.AutomationSecurity = 3
        except Exception:
            pass
        return self.excel

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.excel is not None:
                self.excel.Quit()
        except Exception:
            pass
        finally:
            self.excel = None

        try:
            if self._pythoncom is not None:
                self._pythoncom.CoUninitialize()
        except Exception:
            pass


def find_last_cell(ws, *, look_in: int) -> tuple[int, int]:
    last_row_cell = ws.Cells.Find(
        What="*",
        After=ws.Cells(1, 1),
        LookIn=look_in,
        LookAt=1,
        SearchOrder=XL_BY_ROWS,
        SearchDirection=XL_PREVIOUS,
        MatchCase=False,
    )
    last_col_cell = ws.Cells.Find(
        What="*",
        After=ws.Cells(1, 1),
        LookIn=look_in,
        LookAt=1,
        SearchOrder=XL_BY_COLUMNS,
        SearchDirection=XL_PREVIOUS,
        MatchCase=False,
    )

    if last_row_cell is None or last_col_cell is None:
        return 1, 1
    return int(last_row_cell.Row), int(last_col_cell.Column)


def read_excel_document(
    path: str,
    sheet: Union[int, str] = 1,
    range_a1: Optional[str] = None,
    *,
    auto_bounds: bool = True,
    look_in: str = "formulas",
    max_rows: Optional[int] = None,
    max_cols: Optional[int] = None,
    sep: str = "\t",
    row_sep: str = "\n",
    preview_chars: int = 800,
    visible: bool = False,
    password: Optional[str] = None,
) -> OfficeReadResult:
    t0 = time.time()
    workbook = None
    excel_path = _normalize_excel_path(path)

    try:
        with ExcelSession(visible=visible) as excel:
            open_kwargs = dict(
                Filename=excel_path,
                ReadOnly=True,
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            if password:
                open_kwargs["Password"] = password

            workbook = excel.Workbooks.Open(**open_kwargs)
            worksheet = workbook.Worksheets(sheet)
            sheet_name = worksheet.Name

            last_row = last_col = None
            effective_a1 = None

            if range_a1:
                read_range = worksheet.Range(range_a1)
                effective_a1 = range_a1
            elif auto_bounds:
                lookin_const = XL_FORMULAS if look_in.lower() == "formulas" else XL_VALUES
                lr, lc = find_last_cell(worksheet, look_in=lookin_const)
                if max_rows is not None:
                    lr = min(lr, int(max_rows))
                if max_cols is not None:
                    lc = min(lc, int(max_cols))
                read_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lr, lc))
                last_row, last_col = lr, lc
                effective_a1 = f"A1:{_a1_addr(lr, lc)}"
            else:
                read_range = worksheet.Range("A1")
                last_row, last_col = 1, 1
                effective_a1 = "A1"

            values = read_range.Value
            text = _normalize_newlines(_values_to_delimited_text(values, sep=sep, row_sep=row_sep))
            elapsed_ms = int((time.time() - t0) * 1000)

            return OfficeReadResult(
                ok=True,
                app="excel",
                path=excel_path,
                elapsed_ms=elapsed_ms,
                text=text,
                char_count=len(text),
                preview=text[:preview_chars],
                excel_last_row=last_row,
                excel_last_col=last_col,
                excel_sheet=sheet,
                excel_sheet_name=sheet_name,
                excel_range_a1=effective_a1,
            )
    except Exception as exc:
        elapsed_ms = int((time.time() - t0) * 1000)
        return OfficeReadResult(
            ok=False,
            app="excel",
            path=excel_path,
            elapsed_ms=elapsed_ms,
            error_type=type(exc).__name__,
            error_message=_safe_str(exc),
        )
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass


class ExcelBatchReader:
    def __init__(self, *, visible: bool = False, disable_macros: bool = True):
        self.visible = visible
        self.disable_macros = disable_macros
        self.excel = None
        self._pythoncom = None

    def open(self) -> "ExcelBatchReader":
        pythoncom, win32 = _get_com_modules()
        self._pythoncom = pythoncom
        pythoncom.CoInitialize()

        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = bool(self.visible)
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        self.excel.EnableEvents = False
        self.excel.ScreenUpdating = False

        try:
            self.excel.Calculation = XL_CALC_MANUAL
        except Exception:
            pass

        if self.disable_macros:
            try:
                self.excel.AutomationSecurity = 3
            except Exception:
                pass

        return self

    def close(self) -> None:
        try:
            if self.excel is not None:
                self.excel.Quit()
        except Exception:
            pass
        finally:
            self.excel = None

        try:
            if self._pythoncom is not None:
                self._pythoncom.CoUninitialize()
        except Exception:
            pass
        finally:
            self._pythoncom = None

    def _require_open(self) -> None:
        if self.excel is None:
            raise RuntimeError("ExcelBatchReader is not opened. Call .open() first.")

    def read_cells_sheet(
        self,
        path: str,
        *,
        sheet: Union[int, str],
        auto_bounds: bool = True,
        look_in: str = "formulas",
        max_rows: Optional[int] = None,
        max_cols: Optional[int] = None,
    ) -> tuple[bool, Any, str, Optional[str], Optional[int], Optional[int], Optional[str]]:
        self._require_open()
        workbook = None
        excel_path = _normalize_excel_path(path)

        try:
            workbook = self.excel.Workbooks.Open(
                Filename=excel_path,
                ReadOnly=True,
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            worksheet = workbook.Worksheets(sheet)
            sheet_name = worksheet.Name

            if auto_bounds:
                lookin_const = XL_FORMULAS if look_in.lower() == "formulas" else XL_VALUES
                last_row, last_col = find_last_cell(worksheet, look_in=lookin_const)
                if max_rows is not None:
                    last_row = min(last_row, int(max_rows))
                if max_cols is not None:
                    last_col = min(last_col, int(max_cols))
                read_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(last_row, last_col))
                effective_a1 = f"A1:{_a1_addr(last_row, last_col)}"
            else:
                read_range = worksheet.Range("A1")
                effective_a1 = "A1"
                last_row, last_col = 1, 1

            values = read_range.Value
            return True, values, "", effective_a1, last_row, last_col, sheet_name
        except Exception as exc:
            return False, None, _safe_str(exc), None, None, None, None
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            except Exception:
                pass
