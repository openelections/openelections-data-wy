# /// script
# requires-python = ">=3.10"
# ///
"""Download 2014-2024 Wyoming election result zips from sos.wyo.gov and extract
to sources/<date>/.

Each election ships as a single zip containing either per-county .xlsx files
(2014-2020, Family C) or consolidated workbooks with per-county sheets
(2022-2024, Family D).
"""
from __future__ import annotations

import io
import sys
import urllib.request
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = REPO_ROOT / "sources"

# (date, year, election_type, zip_url)
DOWNLOADS = [
    ("20140819", "2014", "primary", "https://sos.wyo.gov/Elections/Docs/2014/Results/Primary/2014_Wyoming_Primary_Results.zip"),
    ("20141104", "2014", "general", "https://sos.wyo.gov/Elections/Docs/2014/Results/General/2014_Wyoming_General_Election_Results.zip"),
    ("20160816", "2016", "primary", "https://sos.wyo.gov/Elections/Docs/2016/Results/Primary/2016_Wyoming_Primary_Results.zip"),
    ("20161108", "2016", "general", "https://sos.wyo.gov/Elections/Docs/2016/Results/General/2016_Wyoming_General_Election_Results.zip"),
    ("20180821", "2018", "primary", "https://sos.wyo.gov/Elections/Docs/2018/Results/Primary/2018_Wyoming_Primary_Results.zip"),
    ("20181106", "2018", "general", "https://sos.wyo.gov/Elections/Docs/2018/Results/General/2018_Wyoming_General_Results.zip"),
    ("20200818", "2020", "primary", "https://sos.wyo.gov/Elections/Docs/2020/Results/Primary/2020_Wyoming_Primary_Results.zip"),
    ("20201103", "2020", "general", "https://sos.wyo.gov/Elections/Docs/2020/Results/General/2020_Wyoming_General_Results.zip"),
    ("20220816", "2022", "primary", "https://sos.wyo.gov/Elections/Docs/2022/Results/Primary/2022_Wyoming_Primary_Results.zip"),
    ("20221108", "2022", "general", "https://sos.wyo.gov/Elections/Docs/2022/Results/General/2022_Wyoming_General_Results.zip"),
    ("20240820", "2024", "primary", "https://sos.wyo.gov/Elections/Docs/2024/Results/Primary/2024_Wyoming_Primary_Results.zip"),
    ("20241105", "2024", "general", "https://sos.wyo.gov/Elections/Docs/2024/Results/General/2024_Wyoming_General_Results.zip"),
]


def main() -> int:
    for date, year, et, url in DOWNLOADS:
        dest_dir = SOURCES_DIR / date
        if dest_dir.exists() and any(dest_dir.iterdir()):
            print(f"skip {date} (already extracted)")
            continue
        print(f"fetching {date} {et} from {url}")
        req = urllib.request.Request(url, headers={"User-Agent": "openelections-data-wy"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        dest_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            z.extractall(dest_dir)
        n = sum(1 for _ in dest_dir.iterdir())
        print(f"  wrote {n} files to {dest_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())