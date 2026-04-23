from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any


LIBRARY_TYPES = {
    "document": 0,
    "directory": 1,
    "vba_project": 2,
    "catvba": 2,
}


def _load_win32com():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required on the CATIA/Langflow host to automate CATIA."
        ) from exc
    return pythoncom, win32com.client


def get_catia():
    pythoncom, win32 = _load_win32com()
    pythoncom.CoInitialize()
    try:
        return win32.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise RuntimeError("Could not connect to a running CATIA.Application session.") from exc


def execute_macro(
    macro_path: str,
    module_name: str,
    procedure_name: str,
    args: list[str],
    library_type: str = "catvba",
    wait_active_document: bool = False,
    timeout_sec: int = 600,
) -> dict[str, Any]:
    catia = get_catia()
    macro_file = Path(macro_path)
    if not macro_file.exists():
        raise FileNotFoundError(f"Macro file does not exist: {macro_file}")

    library_type_value = LIBRARY_TYPES.get(library_type.lower())
    if library_type_value is None:
        raise ValueError(f"Unsupported library type: {library_type}")

    before_doc = ""
    try:
        before_doc = catia.ActiveDocument.Name
    except Exception:
        before_doc = ""

    result = catia.SystemService.ExecuteScript(
        str(macro_file),
        library_type_value,
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
        "macro_path": str(macro_file),
        "module_name": module_name,
        "procedure_name": procedure_name,
        "arguments": args,
        "result": result,
        "active_document": active_doc,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Run a CATIA CATVBA/CATScript macro.")
    parser.add_argument("--macro-path", required=True)
    parser.add_argument("--module", required=True)
    parser.add_argument("--procedure", required=True)
    parser.add_argument("--library-type", default="catvba", choices=sorted(LIBRARY_TYPES))
    parser.add_argument("--arg", action="append", default=[])
    parser.add_argument("--wait-active-document", action="store_true")
    parser.add_argument("--timeout-sec", type=int, default=600)
    args = parser.parse_args()

    try:
        payload = execute_macro(
            macro_path=args.macro_path,
            module_name=args.module,
            procedure_name=args.procedure,
            args=args.arg,
            library_type=args.library_type,
            wait_active_document=args.wait_active_document,
            timeout_sec=args.timeout_sec,
        )
        print(json.dumps(payload, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, indent=2), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
