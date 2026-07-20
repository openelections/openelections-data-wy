"""Wyoming county names and abbreviation maps."""
from __future__ import annotations

# Canonical 23 Wyoming counties (alphabetical).
COUNTY_NAMES: list[str] = [
    "Albany",
    "Big Horn",
    "Campbell",
    "Carbon",
    "Converse",
    "Crook",
    "Fremont",
    "Goshen",
    "Hot Springs",
    "Johnson",
    "Laramie",
    "Lincoln",
    "Natrona",
    "Niobrara",
    "Park",
    "Platte",
    "Sheridan",
    "Sublette",
    "Sweetwater",
    "Teton",
    "Uinta",
    "Washakie",
    "Weston",
]
COUNTY_NAMES_SET = frozenset(COUNTY_NAMES)

# Default 2-letter abbreviations used by 2002/2004/2008 county files.
DEFAULT_ABBR = {
    "AL": "Albany",
    "BH": "Big Horn",
    "CA": "Campbell",
    "CB": "Carbon",
    "CK": "Crook",
    "CV": "Converse",
    "FR": "Fremont",
    "GO": "Goshen",
    "HT": "Hot Springs",
    "HS": "Hot Springs",
    "JO": "Johnson",
    "LA": "Laramie",
    "LI": "Lincoln",
    "NA": "Natrona",
    "NI": "Niobrara",
    "PK": "Park",
    "PL": "Platte",
    "SB": "Sublette",
    "SN": "Sheridan",
    "SH": "Sheridan",  # 2000 primary used SH for Sheridan
    "SW": "Sweetwater",
    "TE": "Teton",
    "UI": "Uinta",
    "WA": "Washakie",
    "WE": "Weston",
}

# Filename-friendly lowercase slug used in OpenElections CSV names.
def county_slug(name: str) -> str:
    return name.lower().replace(" ", "_")


def county_from_abbr(abbr: str) -> str:
    key = abbr.upper()
    if key not in DEFAULT_ABBR:
        raise KeyError(f"unknown county abbreviation: {abbr!r}")
    return DEFAULT_ABBR[key]


# 2010/2012 files use long-form county names in filenames; map slug -> canonical.
LONG_NAME_SLUGS = {county_slug(c): c for c in COUNTY_NAMES}


def county_from_long_slug(slug: str) -> str:
    s = slug.lower()
    # Filenames look like "2012_Albany_County_General_PbP.xlsx" -> token between year and "County".
    if s.endswith("_county"):
        s = s[: -len("_county")]
    # Some 2010/2012 files use "Big_Horn", others "Big Horn".
    if s not in LONG_NAME_SLUGS:
        raise KeyError(f"unknown long county slug: {slug!r}")
    return LONG_NAME_SLUGS[s]