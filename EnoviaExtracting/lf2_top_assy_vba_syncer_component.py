from __future__ import annotations

import sys
from pathlib import Path

from langflow.custom import Component
from langflow.io import BoolInput, DropdownInput, Output, StrInput
from langflow.schema import Data

sys.path.append(str(Path(__file__).parent))

from lf_macro_runner_base import DEFAULT_VBA_ROOT, run_catia_macro


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
        payload = run_catia_macro(
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
            timeout_sec=1800,
        )
        return Data(data=payload)
