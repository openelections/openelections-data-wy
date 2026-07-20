# /// script
# requires-python = ">=3.10"
# dependencies = ["xlrd>=2.0", "openpyxl>=3.1"]
# ///
"""Parse Wyoming election sources into OpenElections CSVs.

Usage:
    uv run parsers/parse_all.py --election 20121106
    uv run parsers/parse_all.py --all
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = REPO_ROOT / "sources"
OUTPUT_ROOT = REPO_ROOT

# Allow `uv run parsers/parse_all.py` to import sibling modules.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from crosstab import Row  # noqa: E402
from emit import write_election  # noqa: E402
from families import ELECTION_DATES, ElectionSpec, discover_elections, parse_election  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--election", help="Election date YYYYMMDD")
    ap.add_argument("--all", action="store_true", help="Parse all available elections")
    args = ap.parse_args()

    if not args.election and not args.all:
        ap.error("specify --election YYYYMMDD or --all")

    specs = discover_elections(SOURCES_DIR)
    if args.election:
        specs = [s for s in specs if s.date == args.election]
        if not specs:
            print(f"no election found for date {args.election}", file=sys.stderr)
            return 2

    for spec in specs:
        print(f"parsing {spec.date} {spec.election_type} (family {spec.family})...", flush=True)
        per_county = parse_election(spec)
        if not per_county:
            print(f"  no rows for {spec.date}", file=sys.stderr)
            continue
        n_rows = sum(len(v) for v in per_county.values())
        write_election(OUTPUT_ROOT, spec.date, spec.election_type, per_county)
        print(f"  wrote {len(per_county)} counties, {n_rows} precinct rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())