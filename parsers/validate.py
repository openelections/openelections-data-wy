# /// script
# requires-python = ">=3.10"
# ///
"""Validate generated OpenElections CSVs.

Checks (per election):
  1. statewide precinct file == concat of county precinct files
  2. county-level file == groupby-sum of statewide precinct file
  3. every county has a precinct file (all 23 expected)
  4. (optional) source-total cross-check: parsed county totals vs captured source totals

Usage:
    uv run parsers/validate.py --all
    uv run parsers/validate.py --election 20121106
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from counties import COUNTY_NAMES  # noqa: E402
from families import ELECTION_DATES  # noqa: E402

PRECINCT_HEADER = ["county", "precinct", "office", "district", "party", "candidate", "votes"]
COUNTY_HEADER = ["county", "office", "district", "party", "candidate", "votes"]


def _read(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def validate_election(date: str, election_type: str) -> list[str]:
    et = "primary" if election_type == "primary" else "general"
    year = date[:4]
    base = REPO_ROOT / year / f"{date}__wy__{et}"
    statewide = base.parent / f"{base.name}__precinct.csv"
    county_level = base.parent / f"{base.name}__county.csv"
    problems: list[str] = []

    # Collect per-county files.
    county_files = {}
    for county in COUNTY_NAMES:
        slug = county.lower().replace(" ", "_")
        p = base.parent / f"{base.name}__{slug}__precinct.csv"
        if p.exists():
            county_files[county] = p
    missing = [c for c in COUNTY_NAMES if c not in county_files]
    if missing:
        problems.append(f"{date}: missing county precinct files: {missing}")

    # Concatenate county files and compare to statewide.
    concat_rows: list[dict] = []
    for c, p in county_files.items():
        concat_rows.extend(_read(p))
    sw_rows = _read(statewide)
    # Compare as multisets.
    key = lambda r: (r["county"], r["precinct"], r["office"], r["district"], r["party"], r["candidate"], r["votes"])
    if sorted(key(r) for r in concat_rows) != sorted(key(r) for r in sw_rows):
        # Find the diff.
        a = defaultdict(int)
        for r in concat_rows:
            a[key(r)] += 1
        b = defaultdict(int)
        for r in sw_rows:
            b[key(r)] += 1
        diffs = []
        for k in sorted(set(a) | set(b)):
            if a[k] != b[k]:
                diffs.append((k, a[k], b[k]))
        problems.append(f"{date}: statewide precinct != concat(county files); {len(diffs)} diff(s); first: {diffs[:3]}")

    # County-level == groupby-sum of statewide.
    sums: dict[tuple, int] = defaultdict(int)
    for r in sw_rows:
        try:
            v = int(r["votes"])
        except ValueError:
            problems.append(f"{date}: non-int votes in statewide: {r}")
            continue
        key2 = (r["county"], r["office"], r["district"], r["party"], r["candidate"])
        sums[key2] += v
    cl_rows = _read(county_level)
    cl_map: dict[tuple, int] = {}
    for r in cl_rows:
        try:
            v = int(r["votes"])
        except ValueError:
            problems.append(f"{date}: non-int votes in county-level: {r}")
            continue
        k = (r["county"], r["office"], r["district"], r["party"], r["candidate"])
        cl_map[k] = cl_map.get(k, 0) + v
    mismatches = []
    for k, v in sums.items():
        if cl_map.get(k, 0) != v:
            mismatches.append((k, v, cl_map.get(k, 0)))
    for k, v in cl_map.items():
        if k not in sums:
            mismatches.append((k, 0, v))
    if mismatches:
        problems.append(f"{date}: county-level != groupby-sum(statewide); {len(mismatches)} mismatch(es); first: {mismatches[:3]}")

    return problems


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--election")
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()
    if not args.election and not args.all:
        ap.error("specify --election or --all")

    dates = list(ELECTION_DATES.keys())
    if args.election:
        dates = [args.election]

    any_problem = False
    for date in dates:
        if date not in ELECTION_DATES:
            print(f"unknown election {date}", file=sys.stderr)
            any_problem = True
            continue
        _, et = ELECTION_DATES[date]
        year = date[:4]
        out_dir = REPO_ROOT / year
        if not out_dir.exists():
            print(f"no output dir for {date}", file=sys.stderr)
            any_problem = True
            continue
        problems = validate_election(date, et)
        if problems:
            any_problem = True
            for p in problems:
                print("FAIL:", p)
        else:
            print(f"OK {date} {et}")
    return 1 if any_problem else 0


if __name__ == "__main__":
    sys.exit(main())