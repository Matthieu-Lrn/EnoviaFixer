from __future__ import annotations

from typing import Any

from langflow.custom import Component
from langflow.io import BoolInput, DropdownInput, Output, StrInput
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
    timeout_sec: int = 1800,
) -> dict[str, Any]:
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


class TopAssyVbaSyncer(Component):
    display_name = "LF2 TopAssy VBA Syncer"
    description = "Runs the automated DDP PVR Sync macro against the active ENOVIA PVR."
    icon = "RefreshCw"
    name = "LF2TopAssyVbaSyncer"

    inputs = [
        StrInput(name="vba_root", display_name="VBA Root", value=DEFAULT_VBA_ROOT),
        StrInput(name="macro_file", display_name="Macro File", value="LangflowPVRSync.catvba"),
        StrInput(name="module_name", display_name="Module", value="BA_KBE_GCC_DDP"),
        StrInput(name="procedure_name", display_name="Procedure", value="SyncActivePVR"),
        StrInput(
            name="kbe_path_file",
            display_name="KBE Path File",
            value=r"\\aero\mtlplm\catia\V5_KBE_Tools\Production\00_KBE_Env\06_KBE_CATScript\01_MACROS\02_BA_GCC\08_PATH_FILES\BA_COMMON_KBE_PATH.txt",
        ),
        StrInput(name="toolbar_path", display_name="Toolbar Path", value=""),
        StrInput(
            name="project_number",
            display_name="Project/Tail",
            value="",
            info="Use S#### when syncing by aircraft/tail. Leave blank for BSF mode.",
        ),
        BoolInput(
            name="sync_from_bsf",
            display_name="Sync From BSF",
            value=False,
            info="False uses project/tail mode and revision options.",
        ),
        DropdownInput(
            name="ci_option",
            display_name="CI Option",
            options=["DisplayNever", "DisplayNonRel", "DisplayAlways"],
            value="DisplayNever",
        ),
        DropdownInput(
            name="non_ci_option",
            display_name="Non-CI Option",
            options=["LatestRel", "BSF"],
            value="LatestRel",
        ),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    def build_result(self) -> Data:
        macro_path = str((self.vba_root.rstrip("\\/") + "\\" + self.macro_file))
        payload = _run_catia_macro(
            macro_path=macro_path,
            module_name=self.module_name,
            procedure_name=self.procedure_name,
            args=[
                self.project_number,
                "True" if self.sync_from_bsf else "False",
                self.ci_option,
                self.non_ci_option,
                self.kbe_path_file,
                self.toolbar_path,
            ],
        )
        return Data(data=payload)
