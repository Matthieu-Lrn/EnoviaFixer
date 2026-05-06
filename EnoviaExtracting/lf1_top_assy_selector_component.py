from __future__ import annotations

from typing import Any

from langflow.custom import Component
from langflow.io import DataFrameInput, IntInput, Output, StrInput
from langflow.schema import Data


def _first_value(row: Any, *names: str, default: str = "") -> str:
    if row is None:
        return default
    if hasattr(row, "to_dict"):
        row = row.to_dict()
    for name in names:
        if isinstance(row, dict) and name in row and row[name] not in (None, ""):
            return str(row[name])
    return default


class TopAssySelector(Component):
    display_name = "LF1 TopAssy Selector"
    description = "Selects one top assembly and revision for the downstream CATIA/Excel steps."
    icon = "ListChecks"
    name = "LF1TopAssySelector"

    inputs = [
        DataFrameInput(
            name="top_assemblies",
            display_name="Top Assemblies",
            info="Optional DataFrame source for batch-style selection.",
            required=False,
        ),
        IntInput(
            name="row_index",
            display_name="Row Index",
            value=0,
            info="Row to select when using a dataframe.",
        ),
        StrInput(
            name="manual_top_assembly",
            display_name="Manual Top Assembly",
            value="",
            info="If filled, this overrides the dataframe.",
        ),
        StrInput(
            name="manual_revision",
            display_name="Manual Revision",
            value="",
            info="Revision paired with Manual Top Assembly.",
        ),
    ]

    outputs = [
        Output(display_name="Selection", name="selection", method="build_selection"),
    ]

    def build_selection(self) -> Data:
        top_assy = (self.manual_top_assembly or "").strip()
        revision = (self.manual_revision or "").strip()

        if not top_assy:
            df = self.top_assemblies
            if df is None:
                raise ValueError("Provide either Manual Top Assembly or a Top Assemblies dataframe.")
            row = df.iloc[int(self.row_index)]
            top_assy = _first_value(row, "top_assembly", "TopAssembly", "Top Assembly", "part_number", "Part Number")
            revision = _first_value(row, "revision", "Revision", "rev", "Rev", "expected_revision", "Expected Revision")

        if not top_assy:
            raise ValueError("Could not find a top assembly value in the selected row.")

        return Data(
            data={
                "top_assembly": top_assy,
                "revision": revision,
                "row_index": int(self.row_index),
            }
        )
