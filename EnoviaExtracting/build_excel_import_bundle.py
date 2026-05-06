from __future__ import annotations

import re
import shutil
from pathlib import Path

from oletools.olevba import VBA_Parser


ROOT = Path(__file__).resolve().parent
REPO_ROOT = ROOT.parent
OUTPUT_DIR = ROOT / "excel_vba_merged"
SFB_PREFIX = "SFB_"

DDP_CATVBA = REPO_ROOT / "LaunchScript" / "LaunchScript" / "BA_KBE_GCC_DDP_OCT14_2020d.catvba"
SFB_CATVBA = REPO_ROOT / "LaunchScript" / "LaunchScript" / "BA_KBE_GCC_SAVEFILEBASE_OCT14_2020.catvba"

CORE_FILES = [
    "ExcelCatiaBootstrap.bas",
    "ExcelCatiaBridge.bas",
    "EnoviaSearching.bas",
    "ExcelLangflowSteps.bas",
    "ExcelOrchestrator.bas",
]

DECLARATION_PATTERNS = [
    re.compile(r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
    re.compile(r"^\s*(?:Public|Private)\s+Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
    re.compile(r"^\s*(?:Public|Private)\s+Const\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
    re.compile(r"^\s*(?:Public|Private)\s+Enum\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
    re.compile(r"^\s*(?:Public|Private)\s+Type\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
    re.compile(r"^\s*(?:Public|Private|Dim|Global|Static)\s+(?!Const\b|Sub\b|Function\b|Enum\b|Type\b|Declare\b|Property\b)([A-Za-z_][A-Za-z0-9_]*)\b", re.I),
]


def _write_text(path: Path, content: str) -> None:
    normalized = "\r\n".join(content.splitlines())
    if content.endswith(("\r\n", "\r", "\n")):
        normalized += "\r\n"
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(normalized)


def _module_name_from_text(text: str, fallback: str) -> str:
    match = re.search(r'Attribute VB_Name = "([^"]+)"', text)
    return match.group(1) if match else fallback


def _extract_catvba_modules(catvba_path: Path) -> dict[str, str]:
    if not catvba_path.exists():
        raise FileNotFoundError(f"Missing CATVBA file: {catvba_path}")

    parser = VBA_Parser(str(catvba_path))
    try:
        if not parser.detect_vba_macros():
            raise RuntimeError(f"No VBA macros detected in: {catvba_path}")

        modules: dict[str, str] = {}
        for (_, _, vba_filename, vba_code) in parser.extract_macros():
            modules[vba_filename] = vba_code
        return modules
    finally:
        parser.close()


def _collect_sfb_identifier_map(source_files: dict[str, str]) -> dict[str, str]:
    mapping: dict[str, str] = {}

    for file_name, text in source_files.items():
        module_name = _module_name_from_text(text, Path(file_name).stem)
        mapping[module_name] = f"{SFB_PREFIX}{module_name}"

    for file_name, text in source_files.items():
        if not file_name.lower().endswith((".bas", ".frm")):
            continue
        for line in text.splitlines():
            for pattern in DECLARATION_PATTERNS:
                match = pattern.match(line)
                if match:
                    identifier = match.group(1)
                    mapping.setdefault(identifier, f"{SFB_PREFIX}{identifier}")
                    break

    return mapping


def _replace_identifier_tokens(text: str, mapping: dict[str, str]) -> str:
    for old, new in sorted(mapping.items(), key=lambda item: len(item[0]), reverse=True):
        text = re.sub(rf"\b{re.escape(old)}\b", new, text)
    return text


def _dedupe_shell_execute_declares(text: str, function_name: str) -> str:
    lines = text.splitlines()
    new_lines: list[str] = []
    seen_in_vba7 = False
    seen_in_else = False
    in_vba7 = False
    in_else = False

    pattern = re.compile(rf"^\s*Public\s+Declare(?:\s+PtrSafe)?\s+Function\s+{re.escape(function_name)}\b", re.I)

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("#If VBA7 Then"):
            in_vba7 = True
            in_else = False
        elif stripped.startswith("#Else"):
            in_vba7 = False
            in_else = True
        elif stripped.startswith("#End If"):
            in_vba7 = False
            in_else = False

        if pattern.match(line):
            if in_vba7:
                if seen_in_vba7:
                    continue
                seen_in_vba7 = True
            elif in_else:
                if seen_in_else:
                    continue
                seen_in_else = True

        new_lines.append(line)

    return "\r\n".join(new_lines)


def _fix_broken_line_continuations(text: str) -> str:
    # Some extracted CATVBA modules contain a blank line immediately after
    # a VBA line-continuation underscore, which causes "Expected: identifier"
    # during compile in Excel VBA.
    return re.sub(r"(_)\r?\n(?:[ \t]*\r?\n)+([ \t]+)", r"\1\r\n\2", text)


def _copy_core_files() -> None:
    for file_name in CORE_FILES:
        shutil.copy2(ROOT / file_name, OUTPUT_DIR / file_name)


def _write_modules(modules: dict[str, str], prefix: str = "") -> None:
    for file_name, text in modules.items():
        text = _fix_broken_line_continuations(text)
        if file_name == "BA_KBE_GCC_COMMON.bas" and not prefix:
            text = _dedupe_shell_execute_declares(text, "ShellExecute")
        if file_name == "BA_KBE_GCC_COMMON.bas" and prefix == SFB_PREFIX:
            text = _dedupe_shell_execute_declares(text, f"{SFB_PREFIX}ShellExecute")
        _write_text(OUTPUT_DIR / f"{prefix}{file_name}", text)


def _write_manifest() -> None:
    manifest = f"""Import this folder into the shared Excel workbook.

Recommended import set:
1. Core Excel modules:
   - ExcelCatiaBootstrap.bas
   - ExcelCatiaBridge.bas
   - EnoviaSearching.bas
   - ExcelLangflowSteps.bas
   - ExcelOrchestrator.bas
2. All non-prefixed DDP/PVRSync files in this folder
3. All {SFB_PREFIX} prefixed SaveFileBase files in this folder

This bundle is generated directly from:
- {DDP_CATVBA}
- {SFB_CATVBA}
"""
    _write_text(OUTPUT_DIR / "README.txt", manifest)


def build_bundle() -> Path:
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    _copy_core_files()

    ddp_modules = _extract_catvba_modules(DDP_CATVBA)
    _write_modules(ddp_modules)

    sfb_modules = _extract_catvba_modules(SFB_CATVBA)
    sfb_mapping = _collect_sfb_identifier_map(sfb_modules)
    transformed_sfb_modules = {
        file_name: _replace_identifier_tokens(text, sfb_mapping)
        for file_name, text in sfb_modules.items()
    }
    _write_modules(transformed_sfb_modules, prefix=SFB_PREFIX)

    _write_manifest()
    return OUTPUT_DIR


def main() -> int:
    output_dir = build_bundle()
    print(f"Created Excel VBA import bundle: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
