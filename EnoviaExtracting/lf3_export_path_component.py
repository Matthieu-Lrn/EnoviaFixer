from __future__ import annotations

from pathlib import Path

from langflow.custom import Component
from langflow.io import BoolInput, Output, StrInput
from langflow.schema import Data


class ExportPathComponent(Component):
    display_name = "LF3 Export Path"
    description = "Provides and optionally creates the CATIA extraction output folder."
    icon = "Folder"
    name = "LF3ExportPathComponent"

    inputs = [
        StrInput(
            name="export_path",
            display_name="Export Path",
            value=r"C:\Temp\EnoviaExports",
            required=True,
        ),
        BoolInput(
            name="create_folder",
            display_name="Create Folder",
            value=True,
        ),
    ]

    outputs = [
        Output(display_name="Path", name="path", method="build_path"),
    ]

    def build_path(self) -> Data:
        path = Path(self.export_path)
        if self.create_folder:
            path.mkdir(parents=True, exist_ok=True)
        if not path.exists():
            raise FileNotFoundError(f"Export path does not exist: {path}")
        return Data(data={"export_path": str(path)})
