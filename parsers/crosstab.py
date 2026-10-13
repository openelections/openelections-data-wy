"""Generic crosstab spreadsheet reader for Wyoming election results."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, Iterator

from offices import (
    PSEUDO_BALLOTS,
    PSEUDO_OVER,
    PSEUDO_UNDER,
    PSEUDO_WRITEINS,
    BALLOTS_ALIASES,
    OVER_ALIASES,
    UNDER_ALIASES,
    WRITEINS_ALIASES,
    normalize_office,
    normalize_party,
    split_candidate_suffix,
    PRESIDENT,
)


@dataclass(frozen=True)
class Row:
    county: str
    precinct: str
    office: str
    district: str
    party: str
    candidate: str
    votes: int


@dataclass
class LayoutCfg:
    office_rows: list[int]
    candidate_rows: list[int]
    party_row: int | None = None
    data_start_row: int = 2
    na_markers: frozenset[str] = frozenset()
    total_row_labels: frozenset[str] = frozenset({"official totals", "total", "totals"})
    precinct_col: int = 0
    file_party: str = ""        # fixed party for entire file/sheet (primaries 2000-2008)
    election_type: str = "general"  # "primary" or "general"
    include_ballot_issues: bool = False
    is_tbc: bool = False        # Total Ballots Cast workbook layout
    county: str = ""
    # If True, candidate names are last-name only (2000 primary) -- we still emit them as-is.
    last_name_only: bool = False
    # Per-column office label overrides applied BEFORE forward-fill.
    office_overrides: dict[int, str] | None = None
    # Per-column party label overrides applied AFTER forward-fill (overwriting
    # both the forward-filled label and the explicit/party-raw label). Used for
    # source files where a party group's label cell is missing (e.g. 2026 Platte
    # primary's Democratic House 4 write-in block sits unlabeled after the
    # Republican block).
    party_overrides: dict[int, str] | None = None
    # Back-fill office labels rightward into candidate columns that lack one.
    # Used for .xls statewide summary workbooks (2000/2002 primary/general)
    # where the first candidate of each office group sits one column LEFT of
    # the office label, so forward-fill alone misses it.
    office_back_fill: bool = False
    # Per-cell vote-value overrides for known source data errors, keyed by
    # (precinct, office, district, party, candidate). Precinct "" matches the
    # source Total row (whose key is used in source_totals). Cells matching a
    # key emit the override value even when the source cell is non-numeric;
    # fractional source values (e.g. a stray "56.37") with no matching override
    # still raise, so new typos surface instead of being silently floored.
    vote_overrides: dict[tuple, int] | None = None


@dataclass
class SheetResult:
    rows: list[Row] = field(default_factory=list)
    # source totals captured: {(office, district, party, candidate): votes}
    source_totals: dict[tuple[str, str, str, str], int] = field(default_factory=dict)
    precincts: list[str] = field(default_factory=list)


def norm_cell(v) -> str | int | None:
    """Normalize a raw cell value to a stripped string, an int, or None."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int,)) and not isinstance(v, bool):
        return int(v)
    if isinstance(v, float):
        if v != v:  # NaN
            return None
        if v.is_integer():
            return int(v)
        return v  # leave non-integral floats for caller to error on
    if isinstance(v, str):
        s = v.replace("\n", " ").replace("\r", " ")
        s = re.sub(r"\s+", " ", s).strip()
        return s if s else None
    # Dates, datetimes, etc. -- stringify.
    s = str(v).strip()
    return s if s else None


def _is_int_like(v) -> bool:
    return isinstance(v, int) or (isinstance(v, float) and v.is_integer())


def _as_int(v) -> int:
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    raise ValueError(f"expected integer vote, got {v!r}")


# Ballot-issue indicators: candidate labels OR office labels that mark non-candidate races.
BALLOT_ISSUE_CANDIDATES = {"yes", "no", "for", "against"}
BALLOT_ISSUE_OFFICE_KEYWORDS = (
    "constitutional amendment",
    "judge",
    "retention",
    "judicial",
    "shall",
    "district court commissioners",
    "court of appeals",
    "supreme court justice",
    "circuit court",
)


def _is_ballot_issue(office: str, candidate: str) -> bool:
    cand_l = candidate.lower().strip()
    if cand_l in BALLOT_ISSUE_CANDIDATES:
        return True
    off_l = office.lower()
    if any(k in off_l for k in BALLOT_ISSUE_OFFICE_KEYWORDS):
        return True
    return False


def _is_pseudo_candidate(cand: str) -> str | None:
    """Return canonical pseudo label if matched, else None."""
    c = cand.lower().strip()
    if c in WRITEINS_ALIASES:
        return PSEUDO_WRITEINS
    if c in OVER_ALIASES:
        return PSEUDO_OVER
    if c in UNDER_ALIASES:
        return PSEUDO_UNDER
    if c in BALLOTS_ALIASES:
        return PSEUDO_BALLOTS
    return None


def _cleanup_precinct(name: str) -> str:
    n = re.sub(r"\s+", " ", name).strip()
    # Trailing '*' recount flags (e.g. 2002 natrona2.xls) -- strip.
    n = n.rstrip("*").strip()
    return n


def _drop_presidential_running_mate(name: str) -> str:
    """For presidential tickets like 'Mitt Romney and Paul Ryan', keep only the
    presidential candidate (matches the repo's 2020 file convention)."""
    # Split on " and " -- but only if there's a clearmate. Be conservative: only
    # apply when there's exactly one " and " and the result has at least 2 tokens
    # on each side of length > 2.
    parts = re.split(r"\s+and\s+", name, flags=re.I)
    if len(parts) >= 2:
        return parts[0].strip()
    return name


def parse_sheet(grid: list[list], cfg: LayoutCfg) -> SheetResult:
    """Parse a single sheet (already loaded as a list-of-lists grid)."""
    if cfg.is_tbc:
        return _parse_tbc(grid, cfg)
    res = SheetResult()

    nrows = len(grid)
    if nrows == 0:
        return res
    ncols = max((len(r) for r in grid), default=0)

    # Normalize grid into a rectangular matrix of norm_cell outputs.
    mat: list[list] = []
    for r in grid:
        row_cells = [norm_cell(r[i]) if i < len(r) else None for i in range(ncols)]
        mat.append(row_cells)

    # Build per-column office/party/candidate labels.
    def vjoin(rows: list[int], col: int) -> str:
        parts: list[str] = []
        for r in rows:
            if r >= nrows:
                continue
            v = mat[r][col] if col < len(mat[r]) else None
            if isinstance(v, str) and v:
                parts.append(v)
        return " ".join(parts).strip()

    def ffill_row(row_idx: int) -> list:
        out: list = [None] * ncols
        cur = None
        for c in range(ncols):
            v = mat[row_idx][c] if row_idx < nrows and c < len(mat[row_idx]) else None
            if isinstance(v, str) and v:
                cur = v
            out[c] = cur
        return out

    office_labels_raw = [vjoin(cfg.office_rows, c) for c in range(ncols)]
    # Apply per-column office overrides before forward-fill.
    if cfg.office_overrides:
        for c, label in cfg.office_overrides.items():
            if c < ncols:
                office_labels_raw[c] = label
    # Forward-fill office labels horizontally.
    office_labels: list[str] = []
    cur_office = ""
    for c in range(ncols):
        v = office_labels_raw[c]
        if v:
            cur_office = v
        office_labels.append(cur_office)

    party_labels: list[str] = [""] * ncols
    party_raw: list[str] = [""] * ncols
    if cfg.party_row is not None and cfg.party_row < nrows:
        party_labels = [("" if v is None else (v if isinstance(v, str) else str(v)))
                        for v in ffill_row(cfg.party_row)]
        for c in range(ncols):
            v = mat[cfg.party_row][c] if c < len(mat[cfg.party_row]) else None
            party_raw[c] = (v if isinstance(v, str) and v else "")
    if cfg.party_overrides:
        for c, p in cfg.party_overrides.items():
            if c < ncols:
                party_labels[c] = p
                party_raw[c] = p

    candidate_labels: list[str] = [vjoin(cfg.candidate_rows, c) for c in range(ncols)]

    # Optional back-fill mode (for .xls statewide summary workbooks where the
    # office label is positioned mid-group, not at the first candidate of the
    # group). Forward-fill alone mis-extends the previous office across spacer
    # columns into the next group's first candidates. Instead, identify runs
    # of contiguous candidate columns and assign each run a single office
    # label: the explicit label found inside the run, else the nearest
    # explicit office label on either side (ties prefer left, since labels
    # conventionally start groups). Non-candidate columns with their own
    # explicit label (e.g. "Total Ballots Cast") keep that label.
    if cfg.office_back_fill:
        office_labels = list(office_labels_raw)
        # Build per-column nearest explicit-office info on each side.
        left_lbl: list[str] = [""] * ncols
        left_dist: list[int | None] = [None] * ncols
        cur = ""
        cur_d: int | None = None
        for c in range(ncols):
            if office_labels_raw[c]:
                cur = office_labels_raw[c]
                cur_d = 0
            elif cur_d is not None:
                cur_d += 1
            left_lbl[c] = cur
            left_dist[c] = cur_d
        right_lbl: list[str] = [""] * ncols
        right_dist: list[int | None] = [None] * ncols
        cur = ""
        cur_d = None
        for c in range(ncols - 1, -1, -1):
            if office_labels_raw[c]:
                cur = office_labels_raw[c]
                cur_d = 0
            elif cur_d is not None:
                cur_d += 1
            right_lbl[c] = cur
            right_dist[c] = cur_d
        # Walk runs of contiguous candidate columns.
        c = 0
        while c < ncols:
            if not candidate_labels[c]:
                c += 1
                continue
            run_start = c
            while c < ncols and candidate_labels[c]:
                c += 1
            run_end = c  # exclusive
            # Find the first explicit office label inside the run.
            first_explicit = -1
            for j in range(run_start, run_end):
                if office_labels_raw[j]:
                    first_explicit = j
                    break
            if first_explicit < 0:
                # No explicit label inside the run: use nearest outside.
                ld = left_dist[run_start]
                rd = right_dist[run_start]
                if ld is None and rd is None:
                    label = ""
                elif ld is None:
                    label = right_lbl[run_start]
                elif rd is None:
                    label = left_lbl[run_start]
                elif rd < ld:
                    label = right_lbl[run_start]
                else:
                    label = left_lbl[run_start]
                for j in range(run_start, run_end):
                    office_labels[j] = label
            else:
                # Back-fill the prefix (run_start..first_explicit) from the
                # first explicit label, then forward-fill the rest.
                first_label = office_labels_raw[first_explicit]
                for j in range(run_start, first_explicit):
                    office_labels[j] = first_label
                cur = first_label
                for j in range(first_explicit, run_end):
                    if office_labels_raw[j]:
                        cur = office_labels_raw[j]
                    office_labels[j] = cur

    # If a column's joined office text doesn't normalize to a known office,
    # inherit the previous column's normalized office (handles multi-column
    # merged labels where the continuation cell has a fragment like
    # "of Public Instruction").
    last_norm_office: tuple[str, str] = ("", "")
    for c in range(ncols):
        if not office_labels[c]:
            continue
        canon, dist = normalize_office(office_labels[c])
        if not canon:
            # Doesn't normalize; inherit previous normalized office if there is
            # a candidate in this column.
            if candidate_labels[c] and last_norm_office[0]:
                office_labels[c] = last_norm_office[0]  # raw text becomes canonical
        else:
            last_norm_office = (canon, dist)

    # Classify columns and pre-compute (office_canon, district, party, candidate_clean, is_pseudo, is_skip).
    col_meta: list[dict] = []
    for c in range(ncols):
        if c == cfg.precinct_col:
            col_meta.append({"skip": True})
            continue
        cand = candidate_labels[c]
        office_raw = office_labels[c]
        if not cand and not office_raw:
            col_meta.append({"skip": True})
            continue
        # Detect "Total Votes Cast"/"Total Ballots Cast" office columns where the
        # candidate cell holds a party name (Family A primaries column B).
        office_l = office_raw.lower().strip()
        if office_l in BALLOTS_ALIASES or office_l == "total":
            party_code = ""
            if cand:
                try:
                    party_code = normalize_party(cand)
                    cand_for_emit = PSEUDO_BALLOTS
                except ValueError:
                    cand_for_emit = cand
                    party_code = cfg.file_party or ""
            else:
                cand_for_emit = PSEUDO_BALLOTS
                party_code = cfg.file_party or ""
            col_meta.append({
                "skip": False,
                "office": PSEUDO_BALLOTS,
                "district": "",
                "party": party_code,
                "candidate": cand_for_emit,
                "is_pseudo": True,
                "is_ballot_issue": False,
            })
            continue
        # Blank candidate with non-blank office: a ballots-cast column for the office group
        # (Family A primaries column B). Treat as Ballots Cast pseudo with that office's party.
        # Only fire when the office label is explicit in the source -- a forward-filled
        # label in a spacer column (e.g. 2004 Teton primary unlabeled cols after a
        # single-candidate House district) is not a real ballots-cast column.
        if not cand and office_raw:
            if not office_labels_raw[c]:
                col_meta.append({"skip": True})
                continue
            party = party_labels[c] if party_labels else (cfg.file_party or "")
            party_code = normalize_party(party) if party else (cfg.file_party or "")
            col_meta.append({
                "skip": False,
                "office": office_raw,
                "district": "",
                "party": party_code,
                "candidate": PSEUDO_BALLOTS,
                "is_pseudo": True,
                "is_ballot_issue": False,
            })
            continue
        # Has candidate label.
        pseudo = _is_pseudo_candidate(cand)
        office_canon, district = normalize_office(office_raw) if office_raw else ("", "")
        if pseudo:
            # Skip orphan pseudo-candidate columns. Some 2012 source files have
            # stray Write-Ins/Over Votes/Under Votes columns wedged between office
            # groups with no associated candidate; they would otherwise inherit
            # the preceding office/party via forward-fill and emit duplicates.
            # Walk backwards through contiguous pseudo columns; if the first
            # non-pseudo column we hit is blank (no explicit office, no
            # candidate), this pseudo is orphan -- UNLESS the current column
            # carries an explicit party label (e.g. 2010 DEM-only primaries
            # where a statewide office only had write-ins filed).
            is_orphan = True
            if party_raw and party_raw[c]:
                is_orphan = False
            else:
                for j in range(c - 1, -1, -1):
                    cj = candidate_labels[j]
                    if _is_pseudo_candidate(cj):
                        continue
                    if not office_labels_raw[j] and not cj:
                        is_orphan = True
                    else:
                        is_orphan = False
                    break
            # A pseudo run whose first column carries an explicit party label
            # (e.g. a write-in-only race's Write-Ins column labeled "Democratic",
            # followed by party-less Over/Under columns) belongs to that party's
            # candidate block, even though the run sits after a spacer. Only an
            # explicitly-labeled party that matches this column's forward-filled
            # party anchor counts; unrelated strays stay orphaned.
            if is_orphan:
                # Find the contiguous pseudo run ending at this column.
                run_start = c
                while run_start > 0 and candidate_labels[run_start - 1] and _is_pseudo_candidate(candidate_labels[run_start - 1]):
                    run_start -= 1
                lbl = party_labels[c] if party_labels else ""
                want = ""
                try:
                    want = normalize_party(lbl) if lbl else ""
                except ValueError:
                    want = ""
                if want and party_raw and party_raw[run_start]:
                    try:
                        if normalize_party(party_raw[run_start]) == want:
                            is_orphan = False
                    except ValueError:
                        pass
            if is_orphan:
                col_meta.append({"skip": True})
                continue
            raw_party = party_labels[c] if party_labels else ""
            try:
                pcode = normalize_party(raw_party) if raw_party else ""
            except ValueError:
                pcode = ""
            if not pcode:
                pcode = cfg.file_party or ""
            col_meta.append({
                "skip": False,
                "office": office_canon or office_raw,
                "district": district,
                "party": pcode,
                "candidate": pseudo,
                "is_pseudo": True,
                "is_ballot_issue": False,
            })
            continue
        # Real candidate (or ballot-issue candidate).
        clean_name, suffix_party = split_candidate_suffix(cand)
        # Office may itself be a retention/amendment (ballot issue).
        is_bi = _is_ballot_issue(office_raw, clean_name)
        if is_bi and not cfg.include_ballot_issues:
            col_meta.append({"skip": True, "ballot_issue": True})
            continue
        # President: drop running mate.
        if office_canon == PRESIDENT:
            clean_name = _drop_presidential_running_mate(clean_name)
        # Party precedence: explicit suffix (general) > party-row label (2010/12 primary) > file_party.
        if suffix_party:
            party_code = suffix_party
        elif party_labels and party_labels[c]:
            party_code = normalize_party(party_labels[c])
        elif cfg.file_party:
            party_code = cfg.file_party
        else:
            party_code = ""
        col_meta.append({
            "skip": False,
            "office": office_canon or office_raw,
            "district": district,
            "party": party_code,
            "candidate": clean_name,
            "is_pseudo": False,
            "is_ballot_issue": is_bi,
        })

    # Iterate data rows.
    seen_precincts: list[str] = []
    seen_precinct_set: set[str] = set()
    # Track emitted (precinct, office, district, party, candidate) keys to drop
    # duplicates created by source-layout quirks (e.g. 2020 Uinta primary has a
    # second Democratic Write-Ins/Overvotes/Undervotes block wedged after the
    # main candidate block for House District 49 — same key, different value).
    emitted_keys: set[tuple] = set()
    # Labels in the precinct column that should be skipped as block separators
    # (2000 primary uses "::" to delimit race groups within a single sheet).
    SKIP_PRECINCT_LABELS = {"::", ":"}

    # Pre-scan: detect repeated precinct-name columns. 2000 source layouts
    # split the sheet into side-by-side blocks and repeat the precinct name
    # in spacer columns between blocks. Those columns have no candidate label
    # and would otherwise be misclassified as ballots-cast columns. Mark any
    # column whose value matches the row's precinct label as skip.
    for r in range(cfg.data_start_row, nrows):
        if r >= nrows:
            break
        row = mat[r]
        if not row:
            continue
        label = row[cfg.precinct_col] if cfg.precinct_col < len(row) else None
        if not isinstance(label, str) or not label or label in SKIP_PRECINCT_LABELS:
            continue
        for c, meta in enumerate(col_meta):
            if meta.get("skip"):
                continue
            if c >= len(row):
                continue
            v = row[c]
            if isinstance(v, str) and v == label:
                meta["skip"] = True

    for r in range(cfg.data_start_row, nrows):
        row = mat[r]
        if not row:
            continue
        label = row[cfg.precinct_col] if cfg.precinct_col < len(row) else None
        # Detect totals row.
        if isinstance(label, str) and label.lower().strip() in cfg.total_row_labels:
            # Capture per-column source totals.
            for c, meta in enumerate(col_meta):
                if meta.get("skip"):
                    continue
                if c >= len(row):
                    continue
                v = row[c]
                if v is None or (isinstance(v, str) and (not v or v in cfg.na_markers)):
                    continue
                key = (meta["office"], meta["district"], meta["party"], meta["candidate"])
                tot_key = ("", *key)
                if cfg.vote_overrides and tot_key in cfg.vote_overrides:
                    res.source_totals[key] = cfg.vote_overrides[tot_key]
                elif not _is_int_like(v):
                    continue
                else:
                    res.source_totals[key] = _as_int(v)
            continue
        # Skip fully-blank rows.
        if all((v is None or (isinstance(v, str) and not v)) for v in row):
            continue
        if not isinstance(label, str) or not label:
            # No precinct label; treat as spacer/section break.
            continue
        if label in SKIP_PRECINCT_LABELS:
            continue
        precinct = _cleanup_precinct(label)
        if not precinct:
            continue
        if precinct in seen_precinct_set:
            # Source files occasionally repeat a row (e.g. 2002 Laramie primary).
            # Skip the duplicate rather than aborting.
            continue
        seen_precinct_set.add(precinct)
        seen_precincts.append(precinct)

        for c, meta in enumerate(col_meta):
            if meta.get("skip"):
                continue
            if c >= len(row):
                continue
            v = row[c]
            emit_key = (precinct, meta["office"], meta["district"], meta["party"], meta["candidate"])
            override = cfg.vote_overrides.get(emit_key) if cfg.vote_overrides else None
            if override is not None and not _is_int_like(v):
                # Known-bad cell (e.g. fractional source value); emit the override.
                votes = override
            else:
                if v is None:
                    continue
                if isinstance(v, str):
                    if not v:
                        continue
                    if v in cfg.na_markers:
                        continue
                    # Tolerate text-formatted integers like "0" or "12".
                    if re.fullmatch(r"-?\d+", v.strip()):
                        v = int(v.strip())
                    else:
                        raise ValueError(
                            f"non-numeric vote {v!r} at precinct {precinct!r} col {c} "
                            f"office={meta['office']!r} candidate={meta['candidate']!r} in {cfg.county}"
                        )
                if not _is_int_like(v):
                    raise ValueError(
                        f"non-integer vote {v!r} at precinct {precinct!r} col {c} "
                        f"office={meta['office']!r} candidate={meta['candidate']!r} in {cfg.county}"
                    )
                votes = _as_int(v)
            if emit_key in emitted_keys:
                # Source-layout duplicate (see comment above). Skip the later
                # occurrence rather than emitting a duplicate row.
                continue
            emitted_keys.add(emit_key)
            res.rows.append(Row(
                county=cfg.county,
                precinct=precinct,
                office=meta["office"],
                district=meta["district"],
                party=meta["party"],
                candidate=meta["candidate"],
                votes=votes,
            ))

    res.precincts = seen_precincts
    return res


def _parse_tbc(grid: list[list], cfg: LayoutCfg) -> SheetResult:
    """Parse a Total Ballots Cast sheet.

    Layouts:
      - General (single column): row 0 = ['', 'Total Ballots Cast'], data from row 1.
      - Primary (per-party):     row 0 = ['', 'Total Ballots Cast', '', ...],
                                 row 1 = ['', 'Republican', 'Democratic', 'Non-Partisan', 'Total Ballots'],
                                 data from row 2.
    Each non-precinct column emits a `Ballots Cast` row per precinct, with party
    determined by the column header (or no party for the totals column).
    """
    res = SheetResult()
    nrows = len(grid)
    if nrows == 0:
        return res
    ncols = max((len(r) for r in grid), default=0)
    mat = [[norm_cell(r[i]) if i < len(r) else None for i in range(ncols)] for r in grid]

    # Find header row: last row with blank precinct col that has a string in any other col.
    # Treat a row whose precinct-col cell is the literal "Precinct" (case-insensitive)
    # as a column-header row — Family D (2022+) TBC sheets include a "Precinct" label
    # above the data, which would otherwise be parsed as a precinct name.
    header_idx = -1
    for i in range(nrows):
        row = mat[i]
        label = row[cfg.precinct_col] if cfg.precinct_col < len(row) else None
        if isinstance(label, str) and label:
            if label.lower().strip() == "precinct":
                header_idx = i
                continue
            break
        if any(isinstance(v, str) and v for j, v in enumerate(row) if j != cfg.precinct_col):
            header_idx = i
    if header_idx < 0:
        return res
    header_row = mat[header_idx]

    # Build (col, party_code, is_totals) list.
    cols: list[tuple[int, str, bool]] = []
    last_label = None
    for c in range(ncols):
        if c == cfg.precinct_col:
            continue
        v = header_row[c] if c < len(header_row) else None
        if isinstance(v, str) and v:
            last_label = v
        if last_label is None:
            continue
        lbl_low = last_label.lower().strip()
        if lbl_low in BALLOTS_ALIASES or "total ballot" in lbl_low:
            cols.append((c, "", True))
        else:
            try:
                party = normalize_party(last_label)
            except ValueError:
                party = cfg.file_party or ""
            cols.append((c, party, False))

    data_start = header_idx + 1
    seen: set[str] = set()
    for r in range(data_start, nrows):
        row = mat[r]
        label = row[cfg.precinct_col] if cfg.precinct_col < len(row) else None
        if not isinstance(label, str) or not label:
            continue
        if label.lower().strip() in cfg.total_row_labels:
            continue
        precinct = _cleanup_precinct(label)
        if not precinct or precinct in seen:
            continue
        seen.add(precinct)
        res.precincts.append(precinct)
        for c, party, _is_total in cols:
            v = row[c] if c < len(row) else None
            if v is None or (isinstance(v, str) and (not v or v in cfg.na_markers)):
                continue
            if not _is_int_like(v):
                raise ValueError(
                    f"non-numeric TBC value {v!r} at precinct {precinct!r} col {c} in {cfg.county}"
                )
            res.rows.append(Row(
                county=cfg.county,
                precinct=precinct,
                office=PSEUDO_BALLOTS,
                district="",
                party=party,
                candidate=PSEUDO_BALLOTS,
                votes=_as_int(v),
            ))
    return res


def read_xlsx(path: str) -> list[list]:
    """Read all sheets from an .xlsx file as list-of-lists of raw values."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append([list(r) for r in ws.iter_rows(values_only=True)])
    return out


def read_xls(path: str) -> list[list]:
    """Read all sheets from an .xls file as list-of-lists of raw values."""
    import xlrd
    book = xlrd.open_workbook(path, formatting_info=False)
    out = []
    for sname in book.sheet_names():
        sh = book.sheet_by_name(sname)
        rows = []
        for r in range(sh.nrows):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
        out.append(rows)
    return out


def read_any(path: str) -> list[list]:
    p = path.lower()
    if p.endswith(".xlsx") or p.endswith(".xlsm"):
        return read_xlsx(path)
    if p.endswith(".xls"):
        return read_xls(path)
    raise ValueError(f"unsupported file type: {path}")