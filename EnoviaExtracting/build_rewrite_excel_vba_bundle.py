from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "excel_vba_rewrite"

REWRITE_FILES = [
    "ExcelRewriteBootstrap.bas",
    "EnoviaSearching.bas",
    "ExcelRewriteExtraction.bas",
    "ExcelRewriteXml.bas",
    "ExcelRewriteReports.bas",
    "ExcelRewriteLangflow.bas",
]


def _write_text(path: Path, content: str) -> None:
    normalized = "\r\n".join(content.replace("\r\n", "\n").replace("\r", "\n").splitlines())
    if content.endswith(("\r\n", "\r", "\n")):
        normalized += "\r\n"
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(normalized)


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def build_bundle() -> Path:
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for file_name in REWRITE_FILES:
        _write_text(OUTPUT_DIR / file_name, _read_text(ROOT / file_name))

    _write_text(
        OUTPUT_DIR / "README.txt",
        "Clean rewrite bundle.\r\n"
        "This folder contains only workbook-owned .bas modules.\r\n"
        "No legacy CATVBA helper classes or imported library copies are included.\r\n"
        "Current scope: search + extraction rewrite. PVRSync rewrite is still pending.\r\n",
    )
    return OUTPUT_DIR


def main() -> int:
    output_dir = build_bundle()
    print(f"Created clean rewrite VBA bundle: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
