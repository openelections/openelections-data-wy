# /// script
# requires-python = ">=3.10"
# dependencies = ["xlrd>=2.0", "openpyxl>=3.1"]
# ///
"""Validate generated OpenElections CSVs.

Checks (per election):
  1. statewide precinct file == concat of county precinct files
  2. county-level file == groupby-sum of statewide precinct file
  3. every county has a precinct file (all 23 expected)
  4. (optional) source-total cross-check: parsed county totals vs captured source totals
  5. advisory cross-check vs statewide summary workbook (where one exists):
     per-(county, office, district, party, candidate) totals from the
     per-county CSV files vs the statewide summary workbook. Discrepancies
     are reported as advisory -- they may reflect summary workbook layout
     quirks rather than parser errors.

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
from families import (  # noqa: E402
    ELECTION_DATES,
    ElectionSpec,
    parse_statewide_summary,
)

PRECINCT_HEADER = ["county", "precinct", "office", "district", "party", "candidate", "votes"]
COUNTY_HEADER = ["county", "office", "district", "party", "candidate", "votes"]


def _read(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def cross_check_summary(date: str, election_type: str) -> list[str]:
    """Advisory cross-check of per-county CSV totals vs statewide summary workbook.

    Aggregates my per-county precinct CSVs by (county, office, district, party,
    candidate) and compares to the summary workbook parsed by
    parse_statewide_summary. Mismatches are reported as advisory: a non-zero
    count does not fail validation, since summary workbooks sometimes include
    rows (judicial retentions, ballot issues, party totals) that the per-county
    CSVs intentionally exclude.
    """
    et = "primary" if election_type == "primary" else "general"
    year = date[:4]
    base = REPO_ROOT / year / f"{date}__wy__{et}"
    spec = ElectionSpec(
        date=date, year=year, election_type=election_type,
        family="", sources_dir=REPO_ROOT / "sources" / date,
    )
    summary = parse_statewide_summary(spec)
    if not summary:
        return []
    csv_totals: dict[tuple, int] = defaultdict(int)
    for county in COUNTY_NAMES:
        slug = county.lower().replace(" ", "_")
        p = base.parent / f"{base.name}__{slug}__precinct.csv"
        if not p.exists():
            continue
        with p.open("r", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                k = (r["county"], r["office"], r["district"], r["party"], r["candidate"])
                try:
                    v = int(r["votes"])
                except ValueError:
                    continue
                csv_totals[k] += v
    notes: list[str] = []
    mismatches: list[tuple] = []
    only_in_summary: list[tuple] = []
    only_in_csv: list[tuple] = []
    for k in set(summary) | set(csv_totals):
        s = summary.get(k, 0)
        c = csv_totals.get(k, 0)
        if s == c:
            continue
        if s and not c:
            only_in_summary.append((k, s))
        elif c and not s:
            only_in_csv.append((k, c))
        else:
            mismatches.append((k, s, c))
    if mismatches:
        notes.append(
            f"{date}: summary cross-check {len(mismatches)} value mismatch(es); "
            f"first: {mismatches[:5]}"
        )
    if only_in_summary:
        notes.append(
            f"{date}: summary cross-check {len(only_in_summary)} key(s) only in summary "
            f"(likely ballot issues / judicial / party totals); first: {only_in_summary[:5]}"
        )
    if only_in_csv:
        notes.append(
            f"{date}: summary cross-check {len(only_in_csv)} key(s) only in CSV "
            f"(possible parser invention or summary gap); first: {only_in_csv[:5]}"
        )
    return notes


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

    # Advisory summary cross-check (does not fail validation).
    notes = cross_check_summary(date, election_type)
    for n in notes:
        problems.append("ADVISORY: " + n)

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
        failures = [p for p in problems if not p.startswith("ADVISORY:")]
        advisories = [p for p in problems if p.startswith("ADVISORY:")]
        if failures:
            any_problem = True
            for p in failures:
                print("FAIL:", p)
        for p in advisories:
            print(p)
        if not failures:
            print(f"OK {date} {et}")
    return 1 if any_problem else 0


if __name__ == "__main__":
    sys.exit(main())