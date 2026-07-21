"""Office, party, and pseudo-candidate normalization."""
from __future__ import annotations

import re

# Canonical office names used in output CSVs.
PRESIDENT = "President"
US_SENATE = "U.S. Senate"
US_SENATE_UNEXPIRED = "U.S. Senate (Unexpired Term)"
US_HOUSE = "U.S. House"
GOVERNOR = "Governor"
SEC_STATE = "Secretary of State"
STATE_AUDITOR = "State Auditor"
STATE_TREASURER = "State Treasurer"
SPI = "Superintendent of Public Instruction"
STATE_SENATE = "State Senate"
STATE_HOUSE = "State House"

# Pseudo-candidate row labels (canonical).
PSEUDO_WRITEINS = "Write-ins"
PSEUDO_OVER = "Over Votes"
PSEUDO_UNDER = "Under Votes"
PSEUDO_BALLOTS = "Ballots Cast"

# Source labels that map to pseudo-candidates.
WRITEINS_ALIASES = {"write-ins", "write ins", "write-ins:", "write-in", "write in", "write-ins / totals"}
OVER_ALIASES = {"over votes", "over vote", "overvotes", "over"}
UNDER_ALIASES = {"under votes", "under vote", "undervotes", "under"}
BALLOTS_ALIASES = {
    "total ballots cast",
    "total ballots",
    "ballots cast",
    "total votes cast",
    "votes cast",
    "ballot cast",
    "total ballot cast",
}

# Party canonical codes (OpenElections style used by the 2020 file in this repo).
PARTY_DEM = "DEM"
PARTY_REP = "REP"
PARTY_LIB = "LIB"
PARTY_IND = "IND"
PARTY_NL = "NL"  # Nonpartisan / NL
PARTY_REF = "REF"
PARTY_CTRY = "CTRY"  # Constitutional Party (Wyoming 2012 general)
PARTY_NP = "NP"  # Non-Partisan

PARTY_ALIASES = {
    "dem": PARTY_DEM,
    "democrat": PARTY_DEM,
    "democratic": PARTY_DEM,
    "d": PARTY_DEM,
    "(d)": PARTY_DEM,
    "rep": PARTY_REP,
    "republican": PARTY_REP,
    "r": PARTY_REP,
    "(r)": PARTY_REP,
    "lib": PARTY_LIB,
    "libertarian": PARTY_LIB,
    "l": PARTY_LIB,
    "(l)": PARTY_LIB,
    "ind": PARTY_IND,
    "independent": PARTY_IND,
    "i": PARTY_IND,
    "(i)": PARTY_IND,
    "np": PARTY_NP,
    "nonpartisan": PARTY_NP,
    "non-partisan": PARTY_NP,
    "non partisan": PARTY_NP,
    "nl": PARTY_NL,
    "(nl)": PARTY_NL,
    "n": PARTY_NL,  # 2000 general used "(N)" for Natural Law (Hagelin)
    "ref": PARTY_REF,
    "reform": PARTY_REF,
    "(ref)": PARTY_REF,
    "ctry": PARTY_CTRY,
    "constitutional": PARTY_CTRY,
    "constitution": PARTY_CTRY,
    "(ctry)": PARTY_CTRY,
    "const": PARTY_CTRY,
    "c": PARTY_CTRY,  # 2000 used "(C)" for Constitution; conflict-free in practice (we always know context)
}

# Suffixes that appear in candidate cells for general elections, e.g. "John Barrasso (R)".
SUFFIX_PARTY_RE = re.compile(r"\s*\((D|R|L|I|N|Ref|NL|CTRY|NP|Ind|CONST|C)\)\s*$")
# 2002 general used " - D" style suffixes.
DASH_SUFFIX_RE = re.compile(r"\s+-\s+(D|R|L|I|N|Ref|NL|CTRY|NP|Ind|CONST|C)\s*$")


def normalize_party(token: str | None) -> str:
    if token is None:
        return ""
    t = token.strip()
    if not t:
        return ""
    key = t.lower().strip("()")
    if key in PARTY_ALIASES:
        return PARTY_ALIASES[key]
    # Try wrapping like "(R)"
    alt = t.lower()
    if alt in PARTY_ALIASES:
        return PARTY_ALIASES[alt]
    raise ValueError(f"unknown party token: {token!r}")


def split_candidate_suffix(name: str) -> tuple[str, str]:
    """Return (clean_name, party_code) where party_code is "" if no suffix."""
    n = name.strip()
    m = SUFFIX_PARTY_RE.search(n)
    if m:
        party = normalize_party(m.group(1))
        return SUFFIX_PARTY_RE.sub("", n).strip(), party
    m = DASH_SUFFIX_RE.search(n)
    if m:
        party = normalize_party(m.group(1))
        return DASH_SUFFIX_RE.sub("", n).strip(), party
    return n, ""


OFFICE_PATTERNS = [
    (re.compile(r"president", re.I), PRESIDENT, ""),
    (re.compile(r"u\.?\s*s\.?\s+senator|united states senator|united states sen(?!ate)", re.I), US_SENATE, ""),
    (re.compile(r"u\.?\s*s\.?\s+representative|u\.?\s*s\.?\s+rep\b|united states rep", re.I), US_HOUSE, ""),
    (re.compile(r"governor", re.I), GOVERNOR, ""),
    (re.compile(r"secretary of state", re.I), SEC_STATE, ""),
    (re.compile(r"state auditor|auditor", re.I), STATE_AUDITOR, ""),
    (re.compile(r"state treasurer|treasurer", re.I), STATE_TREASURER, ""),
    (re.compile(r"superintendent of public instruction|superintendent", re.I), SPI, ""),
]

SENATE_DISTRICT_RE = re.compile(r"sen(?:ate)?\s*(?:district)?\s*(\d+)", re.I)
HOUSE_DISTRICT_RE = re.compile(r"house\s*(?:district)?\s*(\d+)", re.I)
# Some sheets label legislative races as "State Senate District 10" or "Senate District 10" or "Sen 10".
# Allow trailing content (e.g. "(Unexpired Term: 2 years)") after the district number.
LEG_SENATE_RE = re.compile(r"^state\s+sen(?:ate)?(\s+district\s*\d+)?", re.I)
LEG_HOUSE_RE = re.compile(r"^state\s+house(\s+district\s*\d+)?", re.I)
SEN_DISTRICT_ONLY_RE = re.compile(r"^sen(?:ate)?\s+(?:district\s*)?(\d+)\b", re.I)
HOUSE_DISTRICT_ONLY_RE = re.compile(r"^house\s+(?:district\s*)?(\d+)\b", re.I)


def normalize_office(raw: str) -> tuple[str, str]:
    """Return (canonical_office, district_str) where district_str is "" if none."""
    if raw is None:
        return "", ""
    s = re.sub(r"\s+", " ", raw).strip()
    if not s:
        return "", ""
    sl = s.lower()
    # Unexpired U.S. Senate term (2008). Must be U.S. Senate (not State Senate
    # District N (Unexpired Term: M years), which is a State Senate race).
    if "unexpired" in sl and ("united states sen" in sl or "u.s. sen" in sl or "u s sen" in sl):
        return US_SENATE_UNEXPIRED, ""
    # State legislative.
    m = SENATE_DISTRICT_RE.search(s)
    # Avoid matching "U.S. Senate" via "sen" - require "state senate" or "senate district" or "sen <digit>"
    if LEG_SENATE_RE.match(s) or SEN_DISTRICT_ONLY_RE.match(s):
        d_match = re.search(r"(\d+)", s)
        return STATE_SENATE, d_match.group(1) if d_match else ""
    if LEG_HOUSE_RE.match(s) or HOUSE_DISTRICT_ONLY_RE.match(s):
        d_match = re.search(r"(\d+)", s)
        return STATE_HOUSE, d_match.group(1) if d_match else ""
    for pat, canon, dist in OFFICE_PATTERNS:
        if pat.search(s):
            return canon, dist
    return s, ""