from __future__ import annotations

import sys
from pathlib import Path

from langflow.custom import Component
from langflow.io import DataInput, Output, StrInput
from langflow.schema import Data

sys.path.append(str(Path(__file__).parent))

from lf_macro_runner_base import DEFAULT_VBA_ROOT, run_catia_macro


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
        payload = run_catia_macro(
            macro_path=macro_path,
            module_name=self.module_name,
            procedure_name=self.procedure_name,
            args=[export_path, self.kbe_path_file, self.toolbar_path],
            timeout_sec=3600,
        )
        payload["export_path"] = export_path
        return Data(data=payload)
