from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from typing import Any


DEFAULT_VBA_ROOT = (
    r"\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument"
    r"\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu"
)


def run_catia_macro(
    macro_path: str,
    module_name: str,
    procedure_name: str,
    args: list[str],
    timeout_sec: int = 900,
    wait_active_document: bool = False,
) -> dict[str, Any]:
    runner = Path(__file__).with_name("catia_macro_runner.py")
    command = [
        sys.executable,
        str(runner),
        "--macro-path",
        macro_path,
        "--module",
        module_name,
        "--procedure",
        procedure_name,
        "--timeout-sec",
        str(timeout_sec),
    ]
    if wait_active_document:
        command.append("--wait-active-document")
    for arg in args:
        command.extend(["--arg", str(arg)])

    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=timeout_sec + 30,
        check=False,
    )
    stdout = completed.stdout.strip()
    stderr = completed.stderr.strip()
    if completed.returncode != 0:
        raise RuntimeError(stderr or stdout or f"CATIA macro failed with code {completed.returncode}")
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        return {"stdout": stdout, "stderr": stderr}


def first_value(row: Any, *names: str, default: str = "") -> str:
    if row is None:
        return default
    if hasattr(row, "to_dict"):
        row = row.to_dict()
    for name in names:
        if isinstance(row, dict) and name in row and row[name] not in (None, ""):
            return str(row[name])
    return default
