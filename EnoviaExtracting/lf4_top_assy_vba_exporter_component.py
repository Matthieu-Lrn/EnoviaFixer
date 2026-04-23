from __future__ import annotations

from langflow.custom import Component
from langflow.io import DataInput, Output, StrInput
from langflow.schema import Data

DEFAULT_VBA_ROOT = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu"
)


def _get_catia():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required on the Langflow/CATIA host.") from exc

    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise RuntimeError("Could not connect to a running CATIA.Application session.") from exc


def _run_catia_macro(
    macro_path: str,
    module_name: str,
    procedure_name: str,
    args: list[str],
    timeout_sec: int = 3600,
) -> dict[str, object]:
    catia = _get_catia()
    result = catia.SystemService.ExecuteScript(
        macro_path,
        2,
        module_name,
        procedure_name,
        args,
    )
    active_doc = ""
    try:
        active_doc = catia.ActiveDocument.Name
    except Exception:
        active_doc = ""
    return {
        "macro_path": macro_path,
        "module_name": module_name,
        "procedure_name": procedure_name,
        "arguments": args,
        "result": result,
        "active_document": active_doc,
        "timeout_sec": timeout_sec,
    }


class TopAssyVbaExporter(Component):
    display_name = "LF4 TopAssy VBA Exporter"
    description = "Runs SaveFileBase extraction for the active synced ENOVIA PVR."
    icon = "Download"
    name = "LF4TopAssyVbaExporter"

    inputs = [
        DataInput(
            name="export_path_data",
            display_name="Export Path Data",
            info="Output from Export Path component.",
            required=True,
        ),
        StrInput(name="vba_root", display_name="VBA Root", value=DEFAULT_VBA_ROOT),
        StrInput(name="macro_file", display_name="Macro File", value="LangflowSaveFileBase.catvba"),
        StrInput(name="module_name", display_name="Module", value="BA_KBE_GCC_SAVEFILEBASE"),
        StrInput(name="procedure_name", display_name="Procedure", value="ExportActivePVR"),
        StrInput(
            name="kbe_path_file",
            display_name="KBE Path File",
            value=r"\\aero\mtlplm\catia\V5_KBE_Tools\Production\00_KBE_Env\06_KBE_CATScript\01_MACROS\02_BA_GCC\08_PATH_FILES\BA_COMMON_KBE_PATH.txt",
        ),
        StrInput(name="toolbar_path", display_name="Toolbar Path", value=""),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    def build_result(self) -> Data:
        export_path = self.export_path_data.data.get("export_path")
        if not export_path:
            raise ValueError("Export path input did not contain 'export_path'.")

        macro_path = str((self.vba_root.rstrip("\\/") + "\\" + self.macro_file))
        payload = _run_catia_macro(
            macro_path=macro_path,
            module_name=self.module_name,
            procedure_name=self.procedure_name,
            args=[export_path, self.kbe_path_file, self.toolbar_path],
        )
        payload["export_path"] = export_path
        return Data(data=payload)
