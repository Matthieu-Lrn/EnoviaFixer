from __future__ import annotations

from pathlib import Path
import re
import shutil


ROOT = Path(__file__).resolve().parent
SOURCE_ROOT = ROOT / "vba_source"
MERGED_ROOT = ROOT / "excel_vba_merged"
OUTPUT_DIR = ROOT / "excel_vba_native"

PVR_ROOT = SOURCE_ROOT / "LangflowPVRSync"
SFB_ROOT = SOURCE_ROOT / "LangflowSaveFileBase"

CORE_FILES = [
    "ExcelCatiaBootstrap.bas",
    "ExcelCatiaBridge.bas",
    "EnoviaSearching.bas",
    "ExcelLangflowSteps.bas",
    "ExcelOrchestrator.bas",
    "ExcelNativeExtraction.bas",
]


def _write_text(path: Path, content: str) -> None:
    normalized = "\r\n".join(content.splitlines())
    if content.endswith(("\r\n", "\r", "\n")):
        normalized += "\r\n"
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(normalized)


def _normalize_text(text: str) -> str:
    text = text.replace("\ufeff", "")
    if "Attribute VB_Name" in text:
        text = text[text.index("Attribute VB_Name") :]
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"(_)\n(?:[ \t]*\n)+([ \t]+)", r"\1\n\2", text)
    return text


def _read_vba_text(path: Path) -> str:
    raw = path.read_bytes()
    text = raw.decode("utf-8", errors="ignore")
    return _normalize_text(text)


def _copy_core_files() -> None:
    for file_name in CORE_FILES:
        _write_text(OUTPUT_DIR / file_name, _read_vba_text(ROOT / file_name))


def _copy_non_prefixed_merged_files() -> None:
    for path in sorted(MERGED_ROOT.iterdir()):
        if not path.is_file() or path.suffix.lower() not in {".bas", ".cls", ".frm"}:
            continue
        if path.name.startswith("SFB_"):
            continue
        if path.name in CORE_FILES:
            continue
        _write_text(OUTPUT_DIR / path.name, _read_vba_text(path))


def _extract_block(text: str, start_pattern: str, end_pattern: str) -> str:
    match = re.search(start_pattern, text, flags=re.I | re.S)
    if not match:
        raise ValueError(f"Could not find block matching: {start_pattern}")
    start = match.start()
    end_match = re.search(end_pattern, text[match.end() :], flags=re.I | re.S)
    if not end_match:
        raise ValueError(f"Could not find end block matching: {end_pattern}")
    end = match.end() + end_match.end()
    return text[start:end].strip("\n")


def _merge_common_module() -> str:
    sfb_common = _read_vba_text(SFB_ROOT / "BA_KBE_GCC_COMMON.bas")
    pvr_common = _read_vba_text(PVR_ROOT / "BA_KBE_GCC_COMMON.bas")

    if "Public Function GetAttributesOfDocRevision" not in sfb_common:
        function_block = _extract_block(
            pvr_common,
            r"Public Function GetAttributesOfDocRevision\b.*?\n",
            r"\nEnd Function",
        )
        sfb_common = sfb_common.rstrip("\n") + "\n\n" + function_block + "\n"

    return sfb_common


def _merge_bdifunctions_module() -> str:
    pvr_bdi = _read_vba_text(PVR_ROOT / "BDIfunctions.bas")
    sfb_bdi = _read_vba_text(SFB_ROOT / "BDIfunctions.bas")

    if "Public Function GetPrimaryDocFromPart" not in pvr_bdi:
        function_block = _extract_block(
            sfb_bdi,
            r"Public Function GetPrimaryDocFromPart\b.*?\n",
            r"\nEnd Function",
        )
        pvr_bdi = pvr_bdi.rstrip("\n") + "\n\n" + function_block + "\n"

    return pvr_bdi


def _write_manifest() -> None:
    manifest = """Workbook-native Excel VBA bundle.

This folder is the single-project replacement for the old CATVBA split:
- keep the ENOVIA search modules already used by the workbook
- keep one shared VBA helper set
- add the native PVRSync runner
- add the native export/download runners

Import every .bas/.cls/.frm file in this folder into one macro-enabled workbook.
No external .catvba files are required at runtime.
"""
    _write_text(OUTPUT_DIR / "README.txt", manifest)


def build_bundle() -> Path:
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    _copy_core_files()
    _copy_non_prefixed_merged_files()

    _write_text(OUTPUT_DIR / "BA_KBE_GCC_COMMON.bas", _merge_common_module())
    _write_text(OUTPUT_DIR / "BDIfunctions.bas", _merge_bdifunctions_module())
    _write_manifest()
    return OUTPUT_DIR


def main() -> int:
    output_dir = build_bundle()
    print(f"Created workbook-native VBA bundle: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
