"""Per-election layout configs and file classification."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from crosstab import LayoutCfg, Row, SheetResult, parse_sheet, read_any
from counties import (
    COUNTY_NAMES_SET,
    county_from_abbr,
    county_from_long_slug,
    county_slug,
)
from offices import normalize_party

ELECTION_DATES = {
    "20000822": ("2000", "primary"),
    "20001107": ("2000", "general"),
    "20020820": ("2002", "primary"),
    "20021105": ("2002", "general"),
    "20040817": ("2004", "primary"),
    "20041102": ("2004", "general"),
    "20080819": ("2008", "primary"),
    "20081104": ("2008", "general"),
    "20100817": ("2010", "primary"),
    "20101102": ("2010", "general"),
    "20120821": ("2012", "primary"),
    "20121106": ("2012", "general"),
}


@dataclass
class ElectionSpec:
    date: str
    year: str
    election_type: str  # "primary" or "general"
    family: str         # "A", "B", or "C"
    sources_dir: Path


# ---------------------------------------------------------------------------
# File classification
# ---------------------------------------------------------------------------

# Family C: 20XX_{County}_County_{Primary|General}_PbP.xlsx + ..._PbP_TBC.xlsx
FAMC_PBP_RE = re.compile(r"^\d{4}_(.+)_County_(Primary|General)_PbP\.xlsx$", re.IGNORECASE)
FAMC_TBC_RE = re.compile(r"^\d{4}_(.+?)_(Primary|General)_PbP_TBC\.xlsx$", re.IGNORECASE)


def classify_family_c(name: str) -> dict | None:
    n = name.lower()
    if n.endswith(".xls"):
        return None
    m = FAMC_TBC_RE.match(name) or FAMC_TBC_RE.match(name.lower())
    if m:
        county_slug_raw = m.group(1)
        et = m.group(2).lower()
        return {"role": "tbc", "county": county_from_long_slug(county_slug_raw), "election_type": et}
    m = FAMC_PBP_RE.match(name)
    if m:
        county_slug_raw = m.group(1)
        et = m.group(2).lower()
        return {"role": "pbp", "county": county_from_long_slug(county_slug_raw), "election_type": et}
    return None


# Family B: 2008 primary uses {CO}-{P}-PbP.xls + {CO}-{P}-PbP_U-O-WI.xls + {CO}-TBC_PbP.xlsx
# General: {CO}-PbP.xls + {CO}-TBC_PbP.xlsx (or .XLSX)
FAMB_PRIMARY_PBP_RE = re.compile(r"^([A-Z]{2})-([A-Z])-PbP(?:_U-O-WI)?\.(xlsx|xls)$", re.IGNORECASE)
FAMB_PRIMARY_TBC_RE = re.compile(r"^([A-Z]{2})-TBC_PbP\.(xlsx|xls)$", re.IGNORECASE)
FAMB_GENERAL_PBP_RE = re.compile(r"^([A-Z]{2})-PbP\.(xlsx|xls)$", re.IGNORECASE)
FAMB_GENERAL_TBC_RE = re.compile(r"^([A-Z]{2})-TBC_PbP\.(xlsx|xls)$", re.IGNORECASE)


def classify_family_b(name: str, election_type: str) -> dict | None:
    if election_type == "primary":
        m = FAMB_PRIMARY_TBC_RE.match(name)
        if m:
            return {"role": "tbc", "county": county_from_abbr(m.group(1))}
        m = FAMB_PRIMARY_PBP_RE.match(name)
        if m:
            county = county_from_abbr(m.group(1))
            party = normalize_party(m.group(2))
            role = "u_o_wi" if "_U-O-WI" in name.upper() else "pbp"
            return {"role": role, "county": county, "party": party}
        return None
    # general
    m = FAMB_GENERAL_TBC_RE.match(name)
    if m:
        return {"role": "tbc", "county": county_from_abbr(m.group(1))}
    m = FAMB_GENERAL_PBP_RE.match(name)
    if m:
        return {"role": "pbp", "county": county_from_abbr(m.group(1))}
    return None


# Family A: 2000 primary {CO}-{D|R}-2000.xls, 2002/2004 primaries {CO}.xls (per-party sheets),
# 2002/2004/2000 generals {CO}.xls (single sheet, possibly with extras).
FAMA_2000_PRIMARY_RE = re.compile(r"^([A-Z]{2})-([A-Z])-2000\.xls$", re.IGNORECASE)
FAMA_2000_GENERAL_RE = re.compile(r"^([A-Za-z ]+)\.xls$", re.IGNORECASE)
FAMA_2002_2004_RE = re.compile(r"^([A-Za-z ]+)\.xls$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Family C: 2010, 2012
# ---------------------------------------------------------------------------

def family_c_cfg(role: str, election_type: str, county: str) -> LayoutCfg:
    if role == "tbc":
        return LayoutCfg(
            office_rows=[0],
            candidate_rows=[1],
            data_start_row=2,
            na_markers=frozenset({"-", "`"}),
            total_row_labels=frozenset({"total", "totals", "official totals"}),
            county=county,
            election_type=election_type,
            is_tbc=True,
        )
    if election_type == "primary":
        return LayoutCfg(
            office_rows=[0],
            party_row=1,
            candidate_rows=[2],
            data_start_row=3,
            na_markers=frozenset({"-", "`"}),
            total_row_labels=frozenset({"total", "totals", "official totals"}),
            county=county,
            election_type=election_type,
        )
    return LayoutCfg(
        office_rows=[0],
        candidate_rows=[1],
        data_start_row=2,
        na_markers=frozenset({"-", "`"}),
        total_row_labels=frozenset({"total", "totals", "official totals"}),
        county=county,
        election_type=election_type,
    )


def parse_family_c(spec: ElectionSpec) -> dict[str, list[Row]]:
    """Returns per-county list of precinct rows.

    PbP and TBC files for the same county are merged. TBC files (Total Ballots
    Cast) use shorter precinct names that drop the trailing precinct code, so
    we align TBC rows to PbP rows by index and use the more-specific PbP name.
    """
    per_county: dict[str, list[Row]] = {}
    pbp_res_by_county: dict[str, SheetResult] = {}
    tbc_res_by_county: dict[str, SheetResult] = {}
    for path in sorted(spec.sources_dir.iterdir()):
        if not path.is_file():
            continue
        info = classify_family_c(path.name)
        if info is None:
            continue
        if info["election_type"] != spec.election_type:
            continue
        county = info["county"]
        cfg = family_c_cfg(info["role"], spec.election_type, county)
        # Source typo: 2012 Fremont general labels col 1 as "United States Senator"
        # but the candidates are the presidential ticket. Override to President.
        if spec.date == "20121106" and county == "Fremont" and info["role"] == "pbp":
            cfg.office_overrides = {1: "United States President"}
        sheets = read_any(str(path))
        for sheet in sheets:
            if not sheet or all(all(v is None for v in r) for r in sheet):
                continue
            res = parse_sheet(sheet, cfg)
            if res.rows or res.precincts:
                if info["role"] == "tbc":
                    tbc_res_by_county[county] = res
                else:
                    pbp_res_by_county[county] = res
                break

    for county, pbp in pbp_res_by_county.items():
        rows = list(pbp.rows)
        tbc = tbc_res_by_county.get(county)
        if tbc:
            # Align TBC rows to PbP rows by index; use PbP precinct names.
            pbp_precincts = pbp.precincts
            tbc_precincts = tbc.precincts
            if len(pbp_precincts) != len(tbc_precincts):
                # Mismatch: emit TBC rows with their own names and flag for review.
                # This is rare; fall back to emitting TBC rows as-is.
                rows.extend(tbc.rows)
            else:
                # Rebuild TBC rows with PbP precinct names by walking the TBC
                # precincts list and re-emitting each precinct's Ballots Cast rows.
                tbc_rows_by_precinct: dict[str, list[Row]] = {}
                for r in tbc.rows:
                    tbc_rows_by_precinct.setdefault(r.precinct, []).append(r)
                for i, tname in enumerate(tbc_precincts):
                    pname = pbp_precincts[i]
                    for r in tbc_rows_by_precinct.get(tname, []):
                        rows.append(Row(
                            county=r.county, precinct=pname, office=r.office,
                            district=r.district, party=r.party, candidate=r.candidate,
                            votes=r.votes,
                        ))
        per_county[county] = rows
    return per_county


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Family B: 2008
# ---------------------------------------------------------------------------

def family_b_cfg(role: str, election_type: str, county: str, party: str = "") -> LayoutCfg:
    if role == "tbc":
        return LayoutCfg(
            office_rows=[0, 1],  # Either row may carry the label; forward-fill handles it.
            candidate_rows=[1],
            data_start_row=2,
            na_markers=frozenset({"-", "`"}),
            total_row_labels=frozenset({"total", "totals", "official totals"}),
            county=county,
            election_type=election_type,
            is_tbc=True,
        )
    # PbP and U-O-WI share the same layout: office row 1, candidate row 2, data row 3+.
    return LayoutCfg(
        office_rows=[1],
        candidate_rows=[2],
        data_start_row=3,
        na_markers=frozenset({"-", "`"}),
        total_row_labels=frozenset({"total", "totals", "official totals"}),
        county=county,
        election_type=election_type,
        file_party=party,
    )


def _align_tbc_to_pbp(rows: list[Row], tbc: SheetResult, pbp_precincts: list[str]) -> list[Row]:
    """Re-key TBC rows onto PbP precinct names by index."""
    if len(tbc.precincts) != len(pbp_precincts):
        return list(tbc.rows)
    by_tname: dict[str, list[Row]] = {}
    for r in tbc.rows:
        by_tname.setdefault(r.precinct, []).append(r)
    out: list[Row] = []
    for i, tname in enumerate(tbc.precincts):
        pname = pbp_precincts[i]
        for r in by_tname.get(tname, []):
            out.append(Row(
                county=r.county, precinct=pname, office=r.office,
                district=r.district, party=r.party, candidate=r.candidate,
                votes=r.votes,
            ))
    return out


def parse_family_b(spec: ElectionSpec) -> dict[str, list[Row]]:
    per_county: dict[str, list[Row]] = {}
    pbp_precincts_by_county: dict[str, list[str]] = {}
    tbc_res_by_county: dict[str, SheetResult] = {}

    def first_real_sheet(sheets):
        for sh in sheets:
            if not sh or all(all(v is None for v in r) for r in sh):
                continue
            return sh
        return None

    for path in sorted(spec.sources_dir.iterdir()):
        if not path.is_file():
            continue
        info = classify_family_b(path.name, spec.election_type)
        if info is None:
            continue
        county = info["county"]
        role = info["role"]
        party = info.get("party", "")
        cfg = family_b_cfg(role, spec.election_type, county, party)
        sheets = read_any(str(path))
        sheet = first_real_sheet(sheets)
        if sheet is None:
            continue
        res = parse_sheet(sheet, cfg)
        if role == "tbc":
            tbc_res_by_county[county] = res
            continue
        if not res.rows and not res.precincts:
            continue
        per_county.setdefault(county, []).extend(res.rows)
        if county not in pbp_precincts_by_county or len(res.precincts) > len(pbp_precincts_by_county[county]):
            pbp_precincts_by_county[county] = res.precincts

    # Merge TBC rows onto PbP precinct names.
    for county, tbc in tbc_res_by_county.items():
        pbp_p = pbp_precincts_by_county.get(county, tbc.precincts)
        aligned = _align_tbc_to_pbp(per_county.get(county, []), tbc, pbp_p)
        per_county.setdefault(county, []).extend(aligned)
    return per_county


# ---------------------------------------------------------------------------
# Family A: 2000, 2002, 2004
# ---------------------------------------------------------------------------

# Map sheet-name party token -> canonical party code.
SHEET_PARTY_TOKENS = {
    "democratic": "DEM",
    "republican": "REP",
    "libertarian": "LIB",
    "non-partisan": "NP",
    "nonpartisan": "NP",
    "independent": "IND",
    "constitution": "CTRY",
    "reform": "REF",
}


def sheet_party(sheet_name: str) -> str:
    s = sheet_name.lower()
    for token, code in SHEET_PARTY_TOKENS.items():
        if token in s:
            return code
    return ""


def family_a_cfg(spec: ElectionSpec, sheet_name: str, county: str, file_party: str = "") -> LayoutCfg:
    """Family A layouts vary by year and election type.

    2000 primary:   2 office rows (0-1) + 1 candidate row (3). data_start_row=4.
                    Per-county-party files (AL-D-2000.xls); party from filename.
    2000 general:   2 office rows (0-1) + 1 candidate row (4). data_start_row=5.
    2002 primary:   3 office rows (0-2) + 1 candidate row (3). data_start_row=4.
    2002 general:   3 office rows (0-1, with row 2 frag) + 1 candidate row (3). data_start_row=4.
    2004 primary:   2 office rows (0-1) + 2 candidate rows (3-4). data_start_row=4.
    2004 general:   2 office rows (0-1) + 2 candidate rows (3-4). data_start_row=4.
    """
    if spec.year == "2000" and spec.election_type == "general":
        return LayoutCfg(
            office_rows=[0, 1],
            candidate_rows=[4],
            data_start_row=5,
            na_markers=frozenset({"NA", "N", "-", "`"}),
            total_row_labels=frozenset({"official totals", "official results", "total", "totals"}),
            county=county,
            election_type=spec.election_type,
            file_party="",
        )
    party = file_party or (sheet_party(sheet_name) if spec.election_type == "primary" else "")
    if spec.year == "2004":
        office_rows = [0, 1]
        candidate_rows = [3, 4]
    elif spec.election_type == "primary":
        office_rows = [0, 1, 2]
        candidate_rows = [3]
    else:
        # 2002 general.
        office_rows = [0, 1]
        candidate_rows = [3]
    return LayoutCfg(
        office_rows=office_rows,
        candidate_rows=candidate_rows,
        data_start_row=4,
        na_markers=frozenset({"NA", "N", "-", "`"}),
        total_row_labels=frozenset({"official totals", "official results", "total", "totals"}),
        county=county,
        election_type=spec.election_type,
        file_party=party,
    )


# 2002 primary Albany col 11 (State Treasurer) has a missing office label.
# We inject the correct label per-file.
FAMILY_A_OFFICE_OVERRIDES: dict[tuple[str, str], dict[int, str]] = {
    ("20020820", "Albany"): {11: "State Treasurer"},
}

# 2002 general: every county file has 4 Governor candidates (Freudenthal, Dawson,
# Bebout, Brossman) at cols 7-10, but the source's "Governor" merged-cell range
# spans only cols 7-9, so Brossman (c10) is mislabeled as "Secretary of State".
# Override cols 7-10 to "Governor" and c11 to "Secretary of State" uniformly.
FAMILY_2002_GENERAL_OVERRIDES = {
    7: "Governor", 8: "Governor", 9: "Governor", 10: "Governor",
    11: "Secretary of State",
}


def parse_family_a(spec: ElectionSpec) -> dict[str, list[Row]]:
    per_county: dict[str, list[Row]] = {}
    for path in sorted(spec.sources_dir.iterdir()):
        if not path.is_file():
            continue
        base = path.stem.lower()
        if "summary" in base or "recount" in base:
            continue
        # 2002 general: skip natrona2.xls (identical + recount flags) -- use natrona.xls.
        if spec.date == "20021105" and path.name.lower() == "natrona2.xls":
            continue
        county, file_party = _family_a_county_from_path(path, spec)
        if county is None:
            continue
        overrides = FAMILY_A_OFFICE_OVERRIDES.get((spec.date, county))
        if spec.date == "20021105":
            # 2002 general: every county file has the same column layout for the
            # Governor group; apply the uniform override.
            overrides = dict(FAMILY_2002_GENERAL_OVERRIDES)
            if FAMILY_A_OFFICE_OVERRIDES.get((spec.date, county)):
                overrides.update(FAMILY_A_OFFICE_OVERRIDES[(spec.date, county)])
        sheets = read_any_with_names(str(path))
        for sheet_name, sheet in sheets:
            if not sheet or all(all(v is None for v in r) for r in sheet):
                continue
            cfg = family_a_cfg(spec, sheet_name, county, file_party=file_party)
            if overrides:
                cfg.office_overrides = overrides
            try:
                res = parse_sheet(sheet, cfg)
            except ValueError as e:
                raise ValueError(f"{path.name} sheet {sheet_name!r}: {e}") from e
            if res.rows or res.precincts:
                per_county.setdefault(county, []).extend(res.rows)
    return per_county


def _family_a_county_from_path(path: Path, spec: ElectionSpec) -> tuple[str | None, str]:
    """Return (county, file_party) for a Family A source path.

    Family A county filenames are like 'Albany.xls', 'Big Horn.xls', 'natrona.xls'.
    2000 primary uses per-county-party files like 'AL-D-2000.xls' where the
    single-letter token carries the party.
    """
    name = path.stem
    nlow = name.lower().strip()
    # 2000 primary: per-county-party files like 'AL-D-2000.xls'.
    if spec.year == "2000" and spec.election_type == "primary":
        m = re.match(r"^([A-Z]{2})-([A-Z])-\d{4}$", path.stem, re.IGNORECASE)
        if m:
            county = county_from_abbr(m.group(1))
            party_token = m.group(2).upper()
            try:
                file_party = normalize_party(party_token)
            except ValueError:
                file_party = ""
            return county, file_party
        return None, ""
    for c in COUNTY_NAMES_SET:
        if nlow == c.lower():
            return c, ""
    return None, ""


def _sheet_name(path: Path, idx: int) -> str:
    p = str(path).lower()
    if p.endswith(".xls"):
        import xlrd
        b = xlrd.open_workbook(str(path))
        return b.sheet_names()[idx] if idx < b.nsheets else ""
    if p.endswith(".xlsx") or p.endswith(".xlsm"):
        from openpyxl import load_workbook
        b = load_workbook(str(path), read_only=True)
        names = b.sheetnames
        return names[idx] if idx < len(names) else ""
    return ""


def read_any_with_names(path: str | Path) -> list[tuple[str, list[list]]]:
    p = str(path).lower()
    if p.endswith(".xls"):
        import xlrd
        book = xlrd.open_workbook(str(path))
        out = []
        for sname in book.sheet_names():
            sh = book.sheet_by_name(sname)
            rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
            out.append((sname, rows))
        return out
    if p.endswith(".xlsx") or p.endswith(".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(str(path), data_only=True, read_only=True)
        out = []
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            out.append((ws.title, rows))
        return out
    raise ValueError(f"unsupported file type: {path}")


def parse_election(spec: ElectionSpec) -> dict[str, list[Row]]:
    if spec.family == "A":
        return parse_family_a(spec)
    if spec.family == "B":
        return parse_family_b(spec)
    if spec.family == "C":
        return parse_family_c(spec)
    raise NotImplementedError(spec.family)


def discover_elections(sources_root: Path) -> list[ElectionSpec]:
    specs: list[ElectionSpec] = []
    for date, (year, et) in ELECTION_DATES.items():
        d = sources_root / date
        if not d.is_dir():
            continue
        if year in ("2010", "2012"):
            family = "C"
        elif year == "2008":
            family = "B"
        else:
            family = "A"
        specs.append(ElectionSpec(date=date, year=year, election_type=et, family=family, sources_dir=d))
    return specs