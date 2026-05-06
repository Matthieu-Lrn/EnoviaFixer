from __future__ import annotations

from langflow.custom import Component
from langflow.io import BoolInput, IntInput, Output, StrInput
from langflow.schema import Data

DEFAULT_WORKBOOK_PATH = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\LangflowEnoviaExtraction.xlsm"
)
DEFAULT_MACRO_NAME = "RunFullExtraction"
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


class ExcelVbaExtractionRunner(Component):
    display_name = "LF0 Excel VBA Extraction Runner"
    description = "Opens the hardcoded server workbook, writes inputs into it, and runs the VBA orchestrator."
    icon = "FileSpreadsheet"
    name = "LF0ExcelVbaExtractionRunner"

    inputs = [
        StrInput(name="top_assembly", display_name="Top Assembly", value="", required=True),
        StrInput(name="revision", display_name="Revision", value="", required=True),
        StrInput(name="project_number", display_name="Project Number", value=""),
        StrInput(name="export_path", display_name="Export Path", value=r"C:\Temp\EnoviaExports"),
        BoolInput(name="sync_from_bsf", display_name="Sync From BSF", value=False),
        StrInput(name="ci_option", display_name="CI Option", value="DisplayNever"),
        StrInput(name="non_ci_option", display_name="Non-CI Option", value="LatestRel"),
        StrInput(name="cell_top_assembly", display_name="Top Assembly Cell", value="B1"),
        StrInput(name="cell_revision", display_name="Revision Cell", value="B2"),
        StrInput(name="cell_project_number", display_name="Project Cell", value="B3"),
        StrInput(name="cell_export_path", display_name="Export Path Cell", value="B4"),
        StrInput(name="cell_sync_from_bsf", display_name="Sync From BSF Cell", value="B5"),
        StrInput(name="cell_ci_option", display_name="CI Option Cell", value="B6"),
        StrInput(name="cell_non_ci_option", display_name="Non-CI Option Cell", value="B7"),
        BoolInput(name="excel_visible", display_name="Excel Visible", value=False),
        BoolInput(name="save_workbook", display_name="Save Workbook", value=True),
        IntInput(name="close_delay_seconds", display_name="Close Delay Seconds", value=0),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    @staticmethod
    def _ensure_input_sheet(workbook):
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

    def build_result(self) -> Data:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError("pywin32 is required on the Langflow/Excel host.") from exc

        import time

        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = None
        try:
            excel.Visible = bool(self.excel_visible)
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(DEFAULT_WORKBOOK_PATH)
            sheet = self._ensure_input_sheet(workbook)

            sheet.Range(self.cell_top_assembly).Value = self.top_assembly
            sheet.Range(self.cell_revision).Value = self.revision
            sheet.Range(self.cell_project_number).Value = self.project_number
            sheet.Range(self.cell_export_path).Value = self.export_path
            sheet.Range(self.cell_sync_from_bsf).Value = "TRUE" if self.sync_from_bsf else "FALSE"
            sheet.Range(self.cell_ci_option).Value = self.ci_option
            sheet.Range(self.cell_non_ci_option).Value = self.non_ci_option

            qualified_macro = f"{workbook.Name}!{DEFAULT_MACRO_NAME}"
            run_result = excel.Run(qualified_macro)

            if self.save_workbook:
                workbook.Save()

            if int(self.close_delay_seconds) > 0:
                time.sleep(int(self.close_delay_seconds))

            payload = {
                "workbook_path": DEFAULT_WORKBOOK_PATH,
                "workbook_name": workbook.Name,
                "macro_name": qualified_macro,
                "run_result": run_result,
                "top_assembly": self.top_assembly,
                "revision": self.revision,
                "project_number": self.project_number,
                "export_path": self.export_path,
                "sync_from_bsf": bool(self.sync_from_bsf),
                "ci_option": self.ci_option,
                "non_ci_option": self.non_ci_option,
            }
            return Data(data=payload)
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=bool(self.save_workbook))
            excel.Quit()
