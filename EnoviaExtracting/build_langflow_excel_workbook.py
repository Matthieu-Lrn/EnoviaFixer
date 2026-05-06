from __future__ import annotations

import argparse
from pathlib import Path


DEFAULT_WORKBOOK_PATH = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\LangflowEnoviaExtraction.xlsm"
)

WORKSHEET_NAME = "Inputs"
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52
DEFAULT_VBA_DIR = "excel_vba_native"

INPUT_LABELS = [
    ("A1", "Top Assembly"),
    ("A2", "Revision"),
    ("A3", "Project Number"),
    ("A4", "Export Path"),
    ("A5", "Sync From BSF"),
    ("A6", "CI Option"),
    ("A7", "Non-CI Option"),
]

DEFAULT_VALUES = [
    ("B1", ""),
    ("B2", ""),
    ("B3", ""),
    ("B4", r"C:\Temp\EnoviaExports"),
    ("B5", "FALSE"),
    ("B6", "DisplayNever"),
    ("B7", "LatestRel"),
]


def _load_excel():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required to build the Excel workbook.") from exc
    return pythoncom, win32com.client


def build_workbook(
    workbook_path: str,
    force: bool = False,
    visible: bool = False,
    vba_dir_name: str = DEFAULT_VBA_DIR,
) -> Path:
    pythoncom, win32 = _load_excel()
    pythoncom.CoInitialize()

    repo_root = Path(__file__).resolve().parent
    merged_vba_root = repo_root / vba_dir_name
    module_paths = sorted(
        path
        for path in merged_vba_root.iterdir()
        if path.is_file() and path.suffix.lower() in {".bas", ".cls", ".frm"}
    )

    missing_modules = [str(path) for path in module_paths if not path.exists()]
    if missing_modules:
        raise FileNotFoundError(f"Missing VBA source files: {missing_modules}")
    if not module_paths:
        raise FileNotFoundError(f"No VBA import files were found in: {merged_vba_root}")

    target = Path(workbook_path)
    if target.exists() and not force:
        raise FileExistsError(
            f"Workbook already exists: {target}. Re-run with --force to overwrite it."
        )

    target.parent.mkdir(parents=True, exist_ok=True)

    excel = win32.Dispatch("Excel.Application")
    workbook = None
    try:
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Add()

        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()

        sheet = workbook.Worksheets(1)
        sheet.Name = WORKSHEET_NAME

        for address, label in INPUT_LABELS:
            sheet.Range(address).Value = label

        for address, value in DEFAULT_VALUES:
            sheet.Range(address).Value = value

        sheet.Columns("A:B").EntireColumn.AutoFit()

        if target.exists():
            target.unlink()

        workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)

        try:
            vb_project = workbook.VBProject
        except Exception as exc:
            raise RuntimeError(
                "Excel blocked VBA project access. Enable 'Trust access to the VBA project object model' "
                "in Excel Macro Settings, then run the builder again."
            ) from exc

        for module_path in module_paths:
            vb_project.VBComponents.Import(str(module_path))

        workbook.Save()
        return target
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create the Langflow Excel workbook and import the VBA modules."
    )
    parser.add_argument("--workbook-path", default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--force", action="store_true", help="Overwrite the workbook if it already exists.")
    parser.add_argument("--visible", action="store_true", help="Show Excel while building the workbook.")
    parser.add_argument(
        "--vba-dir-name",
        default=DEFAULT_VBA_DIR,
        help="Folder under EnoviaExtracting that contains the VBA import bundle.",
    )
    args = parser.parse_args()

    workbook_path = build_workbook(
        workbook_path=args.workbook_path,
        force=args.force,
        visible=args.visible,
        vba_dir_name=args.vba_dir_name,
    )
    print(f"Created workbook: {workbook_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
