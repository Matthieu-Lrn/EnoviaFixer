from __future__ import annotations

import sys
from pathlib import Path

from langflow.custom import Component
from langflow.io import BoolInput, DataFrameInput, IntInput, Output, StrInput
from langflow.schema import Data

sys.path.append(str(Path(__file__).parent))

from lf_macro_runner_base import DEFAULT_VBA_ROOT, first_value, run_catia_macro


class TopAssyEnoviaSearcher(Component):
    display_name = "LF1 TopAssy Enovia Searcher"
    description = "Opens one top assembly in CATIA through the ENOVIA search CATVBA."
    icon = "Search"
    name = "LF1TopAssyEnoviaSearcher"

    inputs = [
        DataFrameInput(
            name="top_assemblies",
            display_name="Top Assemblies",
            info="DataFrame from TopAssyListReader.",
            required=False,
        ),
        IntInput(
            name="row_index",
            display_name="Row Index",
            value=0,
            info="Top assembly row to open.",
        ),
        StrInput(
            name="manual_top_assembly",
            display_name="Manual Top Assembly",
            value="",
            info="Optional test override. If filled, the dataframe is ignored.",
        ),
        StrInput(
            name="manual_revision",
            display_name="Manual Revision",
            value="",
            info="Optional test revision paired with Manual Top Assembly.",
        ),
        StrInput(
            name="vba_root",
            display_name="VBA Root",
            value=DEFAULT_VBA_ROOT,
        ),
        StrInput(
            name="macro_file",
            display_name="Macro File",
            value="EnoviaTopAssemblySearch.catvba",
        ),
        StrInput(
            name="module_name",
            display_name="Module",
            value="EnoviaSearching",
        ),
        StrInput(
            name="procedure_name",
            display_name="Procedure",
            value="OpenTopAssemblyFromEnovia",
        ),
        BoolInput(
            name="wait_active_document",
            display_name="Wait Active Document",
            value=True,
        ),
    ]

    outputs = [
        Output(display_name="Result", name="result", method="build_result"),
    ]

    def build_result(self) -> Data:
        top_assy = (self.manual_top_assembly or "").strip()
        revision = (self.manual_revision or "").strip()

        if not top_assy:
            df = self.top_assemblies
            if df is None:
                raise ValueError("Provide either Manual Top Assembly or a Top Assemblies dataframe.")
            row = df.iloc[int(self.row_index)]
            top_assy = first_value(row, "top_assembly", "TopAssembly", "Top Assembly", "part_number", "Part Number")
            revision = first_value(row, "revision", "Revision", "rev", "Rev", "expected_revision", "Expected Revision")

        if not top_assy:
            raise ValueError("Could not find a top assembly value in the selected row.")

        macro_path = str((self.vba_root.rstrip("\\/") + "\\" + self.macro_file))
        payload = run_catia_macro(
            macro_path=macro_path,
            module_name=self.module_name,
            procedure_name=self.procedure_name,
            args=[top_assy, revision],
            wait_active_document=bool(self.wait_active_document),
        )
        payload.update({"top_assembly": top_assy, "revision": revision, "row_index": int(self.row_index)})
        return Data(data=payload)
