# Wyoming election results parsers

These scripts parse Wyoming Secretary of State precinct-level results spreadsheets
(2000–2012) into OpenElections CSVs.

## Layout

```
parsers/
  fetch_sources.py   materialize sources/<date>/ from git history (commit 635725929^)
  counties.py         23 county names + abbreviation maps
  offices.py          office / party / pseudo-candidate normalization
  crosstab.py         generic crosstab reader (the core)
  families.py         per-election layout configs + file classification
  emit.py             CSV writers and county aggregation
  parse_all.py        CLI entry point
  validate.py         internal invariants + cross-checks
```

## Running

```bash
uv run parsers/fetch_sources.py        # one-time: extract sources/ from git history
uv run parsers/parse_all.py --all       # write 2000/…2012/ CSVs
uv run parsers/validate.py --all        # invariants + cross-checks
```

## Conventions (flagged for maintainer review)

1. **Single primary file per election** with a `party` column (not one file per
   party). e.g. `20120821__wy__primary__albany__precinct.csv`.
2. **Year subdirectories** (`2000/`…`2012/`) instead of flat layout. Existing
   root-level files (1978–1988, 2020, specials) are untouched.
3. **`U.S. Senate (Unexpired Term)` office label** for the 2008 second Senate
   race (Barrasso's special); the 2008 primary files distinguish the two
   Senate races only by column order.

## Output shape

- Per-county precinct CSVs: `<YYYYMMDD>__wy__<primary|general>__<county>__precinct.csv`
- Statewide precinct CSV: `<YYYYMMDD>__wy__<primary|general>__precinct.csv`
- Statewide county CSV: `<YYYYMMDD>__wy__<primary|general>__county.csv`

Columns (precinct): `county,precinct,office,district,party,candidate,votes`
Columns (county): `county,office,district,party,candidate,votes`

Pseudo-candidate rows: `Write-ins`, `Over Votes`, `Under Votes`, `Ballots Cast`.
County-level rows are computed sums of precinct rows (the source totals rows
are used only for validation).

## Scope & caveats

- 2000–2012 only. 2014–2024 sources live on `sos.wyo.gov`, which is blocked
  in this environment; those elections are follow-up.
- Ballot issues (constitutional amendments, judicial retentions) are
  excluded by default. Pass `include_ballot_issues=True` on the `LayoutCfg`
  to emit them.
- **2002 HD36 recount**: certified totals (Gentile 1466 / Tanner 1463) differ
  from the precinct-level source file (1461 / 1460). The 2002 general parser
  uses `20021105/natrona.xls`; `natrona2.xls` (identical + `*` recount flags)
  and `NA-HD36-Recount.xls` (totals-only) are excluded. The 20021126 special
  election CSV is already in the repo.
- **2008 dual U.S. Senate races**: the general file labels one race
  `(Unexpired Term: 4 years)`; the primary files distinguish them only by
  column order (first = regular/Enzi, second = unexpired/Barrasso). Disambiguation
  is asserted against the statewide summary workbooks.
- **2012 Fremont general source typo**: col 1 office label says "United States
  Senator" but the candidates are the presidential ticket; the parser overrides
  cols 1–7 to `President`.

## Dependencies

- Python 3.10+
- `xlrd` (for `.xls` files, 2000–2008)
- `openpyxl` (for `.xlsx` files, 2008–2012)

Each script declares its dependencies via PEP 723 inline metadata, so `uv run`
installs them automatically.