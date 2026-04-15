from __future__ import annotations

import argparse
import re
from pathlib import Path


ITERATION_SUFFIX_RE = re.compile(r"_\d+$")


def build_new_stem(stem: str) -> str | None:
    """Return the renamed stem, or None when no change is needed."""
    if ITERATION_SUFFIX_RE.search(stem):
        return None

    if stem.endswith(" ---"):
        return f"{stem[:-1]}_1"

    if stem.endswith("--"):
        return f"{stem}_1"

    return None


def rename_files(target_dir: Path, dry_run: bool = False) -> list[tuple[Path, Path]]:
    renamed: list[tuple[Path, Path]] = []

    for path in sorted(target_dir.iterdir()):
        if not path.is_file():
            continue

        new_stem = build_new_stem(path.stem)
        if new_stem is None:
            continue

        new_path = path.with_name(f"{new_stem}{path.suffix}")
        if new_path.exists():
            raise FileExistsError(
                f"Cannot rename '{path.name}' to '{new_path.name}': destination already exists."
            )

        renamed.append((path, new_path))
        if not dry_run:
            path.rename(new_path)

    return renamed


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Add missing Enovia iteration suffixes to exported files."
    )
    parser.add_argument(
        "directory",
        nargs="?",
        default=".",
        help="Directory containing the exported files. Defaults to the current folder.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show the planned renames without modifying any files.",
    )
    args = parser.parse_args()

    target_dir = Path(args.directory).resolve()
    if not target_dir.is_dir():
        raise NotADirectoryError(f"Directory not found: {target_dir}")

    renamed = rename_files(target_dir, dry_run=args.dry_run)

    if not renamed:
        print("No files needed renaming.")
        return 0

    for old_path, new_path in renamed:
        print(f"{old_path.name} -> {new_path.name}")

    if args.dry_run:
        print(f"\nDry run only: {len(renamed)} rename(s) planned.")
    else:
        print(f"\nDone: {len(renamed)} file(s) renamed.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
