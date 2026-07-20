"""Output: Row sort order and CSV writers."""
from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from typing import Iterable

from crosstab import Row
from counties import county_slug

# Canonical office ordering for deterministic CSV output. Offices not listed
# sort by name after the listed ones.
OFFICE_ORDER = {
    "President": 0,
    "U.S. Senate": 1,
    "U.S. Senate (Unexpired Term)": 2,
    "U.S. House": 3,
    "Governor": 4,
    "Secretary of State": 5,
    "State Auditor": 6,
    "State Treasurer": 7,
    "Superintendent of Public Instruction": 8,
    "State Senate": 9,
    "State House": 10,
}

# Pseudo-candidates sort after real candidates within a race, in this order.
PSEUDO_ORDER = {
    "Ballots Cast": 0,
    "Write-ins": 1,
    "Over Votes": 2,
    "Under Votes": 3,
}

# Party sort order (matches the 2020 file's grouping).
PARTY_ORDER = {
    "REP": 0,
    "DEM": 1,
    "LIB": 2,
    "IND": 3,
    "NL": 4,
    "REF": 5,
    "CTRY": 6,
    "NP": 7,
    "": 8,
}


def _office_key(office: str) -> tuple:
    return (OFFICE_ORDER.get(office, 99), office)


def _district_key(d: str) -> tuple:
    if not d:
        return (0, 0)
    try:
        return (1, int(d))
    except ValueError:
        return (1, 0)


def _party_key(p: str) -> tuple:
    return (PARTY_ORDER.get(p, 99), p)


def _candidate_key(cand: str) -> tuple:
    if cand in PSEUDO_ORDER:
        return (1, PSEUDO_ORDER[cand], cand)
    return (0, 0, cand)


def sort_rows(rows: Iterable[Row]) -> list[Row]:
    return sorted(rows, key=lambda r: (
        r.county,
        r.precinct,
        _office_key(r.office),
        _district_key(r.district),
        _party_key(r.party),
        _candidate_key(r.candidate),
        r.candidate,
    ))


def sort_county_rows(rows: Iterable[Row]) -> list[Row]:
    """Sort rows for the county-level file (no precinct)."""
    return sorted(rows, key=lambda r: (
        r.county,
        _office_key(r.office),
        _district_key(r.district),
        _party_key(r.party),
        _candidate_key(r.candidate),
        r.candidate,
    ))


PRECINCT_HEADER = ["county", "precinct", "office", "district", "party", "candidate", "votes"]
COUNTY_HEADER = ["county", "office", "district", "party", "candidate", "votes"]


def _write_csv(path: Path, header: list[str], rows: Iterable[Row], precinct: bool) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        w = csv.writer(f, quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
        w.writerow(header)
        for r in rows:
            if precinct:
                w.writerow([r.county, r.precinct, r.office, r.district, r.party, r.candidate, r.votes])
            else:
                w.writerow([r.county, r.office, r.district, r.party, r.candidate, r.votes])


def write_precinct_csv(path: Path, rows: Iterable[Row]) -> None:
    _write_csv(path, PRECINCT_HEADER, rows, precinct=True)


def write_county_csv(path: Path, rows: Iterable[Row]) -> None:
    _write_csv(path, COUNTY_HEADER, rows, precinct=False)


def aggregate_county(rows: Iterable[Row]) -> list[Row]:
    """Sum precinct rows into county-level rows. Input must be precinct rows."""
    sums: dict[tuple, int] = defaultdict(int)
    for r in rows:
        key = (r.county, r.office, r.district, r.party, r.candidate)
        sums[key] += r.votes
    out: list[Row] = []
    for (county, office, district, party, candidate), v in sums.items():
        out.append(Row(
            county=county,
            precinct="",
            office=office,
            district=district,
            party=party,
            candidate=candidate,
            votes=v,
        ))
    return sort_county_rows(out)


def election_paths(root: Path, date: str, election_type: str, county: str | None) -> dict[str, Path]:
    """Return the output file paths for an election.

    If `county` is given, returns the per-county precinct path under key 'county'.
    Always returns 'statewide_precinct' and 'county_level'.
    """
    et = "primary" if election_type == "primary" else "general"
    base = root / date[:4] / f"{date}__wy__{et}"
    return {
        "statewide_precinct": base.parent / f"{base.name}__precinct.csv",
        "county_level": base.parent / f"{base.name}__county.csv",
        "county": base.parent / f"{base.name}__{county_slug(county)}__precinct.csv" if county else None,
    }


def write_election(root: Path, date: str, election_type: str,
                   per_county: dict[str, list[Row]]) -> None:
    """Write all CSVs for one election.

    `per_county` maps county name -> list of precinct rows.
    """
    et = "primary" if election_type == "primary" else "general"
    base = root / date[:4] / f"{date}__wy__{et}"
    base.parent.mkdir(parents=True, exist_ok=True)

    all_precinct: list[Row] = []
    for c, rows in per_county.items():
        rows = sort_rows(rows)
        write_precinct_csv(
            base.parent / f"{base.name}__{county_slug(c)}__precinct.csv",
            rows,
        )
        all_precinct.extend(rows)

    write_precinct_csv(base.parent / f"{base.name}__precinct.csv", sort_rows(all_precinct))
    write_county_csv(base.parent / f"{base.name}__county.csv", aggregate_county(all_precinct))