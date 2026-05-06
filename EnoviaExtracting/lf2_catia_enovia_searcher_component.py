from __future__ import annotations

import time

from langflow.custom import Component
from langflow.io import BoolInput, DataInput, IntInput, Output, StrInput
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


class CatiaEnoviaSearcher(Component):
    display_name = "LF2 CATIA ENOVIA Searcher"
    description = "Runs the workbook ENOVIA search macro to open the selected top assembly in CATIA."
    icon = "Search"
    name = "LF2CatiaEnoviaSearcher"

    inputs = [
        DataInput(
            name="selection_data",
            display_name="Selection Data",
            info="Output from LF1 TopAssy Selector.",
            required=False,
        ),
        StrInput(
            name="manual_top_assembly",
            display_name="Manual Top Assembly",
            value="",
            info="Optional override if you want to bypass the selector.",
        ),
        StrInput(
            name="manual_revision",
            display_name="Manual Revision",
            value="",
            info="Revision paired with Manual Top Assembly.",
        ),
        StrInput(
            name="workbook_path",
            display_name="Workbook Path",
            value=DEFAULT_WORKBOOK_PATH,
        ),
        StrInput(
            name="macro_name",
            display_name="Macro Name",
            value="RunTopAssySearch",
        ),
        BoolInput(
            name="excel_visible",
            display_name="Excel Visible",
            value=False,
        ),
        BoolInput(
            name="save_workbook",
            display_name="Save Workbook",
            value=True,
        ),
        BoolInput(
            name="wait_active_document",
            display_name="Wait Active Document",
            value=True,
        ),
        IntInput(
            name="wait_timeout_sec",
            display_name="Wait Timeout Sec",
            value=900,
        ),
        IntInput(
            name="close_delay_seconds",
            display_name="Close Delay Seconds",
            value=0,
        ),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    def build_result(self) -> Data:
        selection = self.selection_data.data if self.selection_data else {}
        top_assy = (self.manual_top_assembly or selection.get("top_assembly") or "").strip()
        revision = (self.manual_revision or selection.get("revision") or "").strip()

        if not top_assy:
            raise ValueError("Provide a top assembly through LF1 TopAssy Selector or Manual Top Assembly.")

        before_doc = _get_active_catia_document_name()
        payload = _run_excel_macro(
            macro_name=self.macro_name,
            args=[top_assy, revision],
            workbook_path=self.workbook_path,
            excel_visible=bool(self.excel_visible),
            save_workbook=bool(self.save_workbook),
            close_delay_seconds=int(self.close_delay_seconds),
        )

        active_doc = _get_active_catia_document_name()
        if bool(self.wait_active_document):
            deadline = time.time() + int(self.wait_timeout_sec)
            while time.time() < deadline:
                active_doc = _get_active_catia_document_name()
                if active_doc and active_doc != "CATImmSearchDoc" and active_doc != before_doc:
                    break
                time.sleep(1)

        payload.update(
            {
                "top_assembly": top_assy,
                "revision": revision,
                "active_document_before": before_doc,
                "active_document_after": active_doc,
            }
        )
        return Data(data=payload)
