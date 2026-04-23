from __future__ import annotations

import time
from typing import Any

from langflow.custom import Component
from langflow.io import BoolInput, DataFrameInput, IntInput, Output, StrInput
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
    wait_active_document: bool = False,
    timeout_sec: int = 900,
) -> dict[str, Any]:
    catia = _get_catia()

    before_doc = ""
    try:
        before_doc = catia.ActiveDocument.Name
    except Exception:
        before_doc = ""

    result = catia.SystemService.ExecuteScript(
        macro_path,
        2,
        module_name,
        procedure_name,
        args,
    )

    if wait_active_document:
        deadline = time.time() + timeout_sec
        while time.time() < deadline:
            try:
                active_name = catia.ActiveDocument.Name
                if active_name and active_name != "CATImmSearchDoc" and active_name != before_doc:
                    break
            except Exception:
                pass
            time.sleep(1)

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
    }


def _first_value(row: Any, *names: str, default: str = "") -> str:
    if row is None:
        return default
    if hasattr(row, "to_dict"):
        row = row.to_dict()
    for name in names:
        if isinstance(row, dict) and name in row and row[name] not in (None, ""):
            return str(row[name])
    return default


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
            top_assy = _first_value(row, "top_assembly", "TopAssembly", "Top Assembly", "part_number", "Part Number")
            revision = _first_value(row, "revision", "Revision", "rev", "Rev", "expected_revision", "Expected Revision")

        if not top_assy:
            raise ValueError("Could not find a top assembly value in the selected row.")

        macro_path = str((self.vba_root.rstrip("\\/") + "\\" + self.macro_file))
        payload = _run_catia_macro(
            macro_path=macro_path,
            module_name=self.module_name,
            procedure_name=self.procedure_name,
            args=[top_assy, revision],
            wait_active_document=bool(self.wait_active_document),
        )
        payload.update({"top_assembly": top_assy, "revision": revision, "row_index": int(self.row_index)})
        return Data(data=payload)
