from __future__ import annotations

from langflow.custom import Component
from langflow.io import BoolInput, DataInput, Output, StrInput
from langflow.schema import Data

DEFAULT_WORKBOOK_PATH = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\LangflowEnoviaExtraction.xlsm"
)


def _run_excel_macro(
    macro_name: str,
    *,
    args: list | None = None,
    workbook_path: str = DEFAULT_WORKBOOK_PATH,
    excel_visible: bool = False,
    save_workbook: bool = True,
    close_delay_seconds: int = 0,
):
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
        qualified_macro = f"{workbook.Name}!{macro_name}"
        run_result = excel.Run(qualified_macro, *(args or []))

        if save_workbook:
            workbook.Save()

        if int(close_delay_seconds) > 0:
            import time

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


def _get_catia():
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


def _get_active_catia_document_name() -> str:
    catia = _get_catia()
    try:
        return catia.ActiveDocument.Name
    except Exception:
        return ""


class CatiaExtracterPvrSync(Component):
    display_name = "LF5 CATIA Extracter (PVRSync)"
    description = "Runs the in-workbook export step for the active synced ENOVIA PVR."
    icon = "Download"
    name = "LF5CatiaExtracterPvrSync"

    inputs = [
        DataInput(
            name="export_path_data",
            display_name="Export Path Data",
            info="Output from LF4 PVRSync Export Path.",
            required=True,
        ),
        DataInput(
            name="previous_step_data",
            display_name="Previous Step Data",
            info="Optional chain input from LF3 CATIA TopAssy Syncer.",
            required=False,
        ),
        StrInput(
            name="workbook_path",
            display_name="Workbook Path",
            value=DEFAULT_WORKBOOK_PATH,
        ),
        StrInput(
            name="macro_name",
            display_name="Macro Name",
            value="RunPVRExportStep",
        ),
        StrInput(
            name="kbe_path_file",
            display_name="KBE Path File",
            value=r"\\aero\mtlplm\catia\V5_KBE_Tools\Production\00_KBE_Env\06_KBE_CATScript\01_MACROS\02_BA_GCC\08_PATH_FILES\BA_COMMON_KBE_PATH.txt",
        ),
        StrInput(name="toolbar_path", display_name="Toolbar Path", value=""),
        BoolInput(name="excel_visible", display_name="Excel Visible", value=False),
        BoolInput(name="save_workbook", display_name="Save Workbook", value=True),
        StrInput(name="close_delay_seconds", display_name="Close Delay Seconds", value="0"),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    def build_result(self) -> Data:
        export_path = self.export_path_data.data.get("export_path")
        if not export_path:
            raise ValueError("Export path input did not contain 'export_path'.")

        payload = _run_excel_macro(
            macro_name=self.macro_name,
            args=[export_path, self.kbe_path_file, self.toolbar_path],
            workbook_path=self.workbook_path,
            excel_visible=bool(self.excel_visible),
            save_workbook=bool(self.save_workbook),
            close_delay_seconds=int(self.close_delay_seconds),
        )
        payload.update(
            {
                "export_path": export_path,
                "previous_step_data": self.previous_step_data.data if self.previous_step_data else None,
                "active_document": _get_active_catia_document_name(),
            }
        )
        return Data(data=payload)
