# /// script
# requires-python = ">=3.10"
# ///
"""Materialize sources/<electiondate>/ from git history.

Source spreadsheets for the 2000-2012 elections were removed from this repo
in commit 635725929 ("moved source files to sources repo", Nov 2020). This
script extracts them from the parent of that commit into ./sources/<date>/.

Quattro Pro `.wb3` files are skipped: the 2000 general has `.xls` twins.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

DELETION_COMMIT = "635725929a9b495dd251a8b1c609217ac63b1678"
REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = REPO_ROOT / "sources"
SKIP_SUFFIXES = (".wb3",)


def list_deleted_files() -> list[str]:
    out = subprocess.check_output(
        ["git", "show", f"{DELETION_COMMIT}", "--name-only", "--diff-filter=D", "--no-renames"],
        cwd=REPO_ROOT,
        text=True,
    )
    files: list[str] = []
    for line in out.splitlines():
        line = line.strip()
        if line and "/" in line and not line.startswith("commit "):
            files.append(line)
    return files


def extract(path: str) -> bytes:
    return subprocess.check_output(
        ["git", "show", f"{DELETION_COMMIT}^:{path}"],
        cwd=REPO_ROOT,
    )


def main() -> int:
    files = list_deleted_files()
    written = 0
    skipped = 0
    for rel in files:
        if any(rel.lower().endswith(s) for s in SKIP_SUFFIXES):
            skipped += 1
            continue
        dest = SOURCES_DIR / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists():
            continue
        data = extract(rel)
        dest.write_bytes(data)
        written += 1
    print(f"wrote {written} files, skipped {skipped} (.wb3), into {SOURCES_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())