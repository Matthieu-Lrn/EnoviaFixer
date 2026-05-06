from __future__ import annotations

import time
from typing import Any

DEFAULT_WORKBOOK_PATH = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\LangflowEnoviaExtraction.xlsm"
)
DEFAULT_INPUT_SHEET = "Inputs"

INPUT_LABELS = [
    ("A1", "Top Assembly"),
    ("A2", "Revision"),
    ("A3", "Project Number"),
    ("A4", "Export Path"),
    ("A5", "Sync From BSF"),
    ("A6", "CI Option"),
    ("A7", "Non-CI Option"),
]

DEFAULT_VALUES = [
    ("B1", ""),
    ("B2", ""),
    ("B3", ""),
    ("B4", r"C:\Temp\EnoviaExports"),
    ("B5", "FALSE"),
    ("B6", "DisplayNever"),
    ("B7", "LatestRel"),
]


def ensure_input_sheet(workbook):
    try:
        sheet = workbook.Worksheets(DEFAULT_INPUT_SHEET)
    except Exception:
        sheet = workbook.Worksheets.Add()
        sheet.Name = DEFAULT_INPUT_SHEET

    for address, label in INPUT_LABELS:
        sheet.Range(address).Value = label

    for address, value in DEFAULT_VALUES:
        if sheet.Range(address).Value in (None, ""):
            sheet.Range(address).Value = value

    sheet.Columns("A:B").EntireColumn.AutoFit()
    return sheet


def run_excel_macro(
    macro_name: str,
    args: list[Any] | None = None,
    *,
    workbook_path: str = DEFAULT_WORKBOOK_PATH,
    excel_visible: bool = False,
    save_workbook: bool = True,
    close_delay_seconds: int = 0,
    ensure_inputs_sheet: bool = False,
    sheet_values: dict[str, Any] | None = None,
) -> dict[str, Any]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required on the Langflow/Excel host.") from exc

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = None

    try:
        excel.Visible = bool(excel_visible)
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(workbook_path)
        sheet = None

        if ensure_inputs_sheet or sheet_values:
            sheet = ensure_input_sheet(workbook)

        if sheet is not None and sheet_values:
            for cell, value in sheet_values.items():
                sheet.Range(cell).Value = value

        qualified_macro = f"{workbook.Name}!{macro_name}"
        run_result = excel.Run(qualified_macro, *(args or []))

        if save_workbook:
            workbook.Save()

        if int(close_delay_seconds) > 0:
            time.sleep(int(close_delay_seconds))

        return {
            "workbook_path": workbook_path,
            "workbook_name": workbook.Name,
            "macro_name": qualified_macro,
            "arguments": list(args or []),
            "run_result": run_result,
        }
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=bool(save_workbook))
        excel.Quit()


def get_catia():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required on the Langflow/CATIA host.") from exc

    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        try:
            return win32com.client.Dispatch("CATIA.Application")
        except Exception as exc:
            raise RuntimeError(
                "Could not connect to CATIA.Application with GetActiveObject or Dispatch."
            ) from exc


def get_active_catia_document_name() -> str:
    catia = get_catia()
    try:
        return catia.ActiveDocument.Name
    except Exception:
        return ""
