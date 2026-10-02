"""
olympiad_scraper.py

Scrapes participant/medalist lists from several international science
olympiads and writes them to one combined CSV: `olympiad_participants.csv`.

CURRENTLY IMPLEMENTED (verified against the live sites while writing this):
  - IMO   (Mathematics)  -> imo-official.org
  - IPhO  (Physics)      -> ipho-unofficial.org   (community-run, but the
                             most complete public archive of IPhO results)
  - IChO  (Chemistry)    -> icho-official.org/results  (official)
  - IOI   (Informatics)  -> cphof.org/standings/ioi/{year}  (community-run
                             "Competitive Programming Hall of Fame" archive;
                             used because the OFFICIAL IOI archive,
                             stats.ioinformatics.org, disallows automated
                             scraping in its robots.txt — cphof.org does not,
                             so this respects that instead of overriding it)

NOT IMPLEMENTED, ON PURPOSE, WITH REASONS:
  - IBO  (Biology)       There is no single central results archive; every
                         year's results live on that year's own host-
                         country website (ibo2025.org, ibo2024...org, ...),
                         each with a different page structure.
  - IOAA (Astronomy) and regional olympiads (APhO, EGMO, Asian Physics
                         Olympiad, Balkan MO, etc.) have the same problem
                         as IBO: no stable, uniform archive to point a
                         generic scraper at.

For IBO / IOAA / regional olympiads, use the GenericTableSource class below:
give it the specific year-by-year URLs plus a small column mapping, and it
will scrape any results page that is a plain HTML <table>. You will need to
inspect each site yourself and fill in the config, because their table
layouts are not uniform.

Install:
    pip install requests beautifulsoup4 cloudscraper

Run:
    python olympiad_scraper.py

Resumable: skips (year, subject) pairs already present in the output CSV.
"""

from __future__ import annotations

import csv
import os
import re
import time
import random
from dataclasses import dataclass, field
from urllib.parse import urljoin, quote

import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------

CONTACT_EMAIL = "tvuong03@gmail.com"
# A self-identifying bot UA gets blocked/empty-paged by several sites we hit here
# (notably cphof.org), so use an ordinary browser UA instead.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}

# cphof.org sits behind a Cloudflare JS bot-challenge: plain `requests` gets the
# challenge page (no table, no error) instead of the real content regardless of
# headers. `cloudscraper` is a requests-compatible client that solves that
# challenge automatically. pip install cloudscraper if you don't have it.
try:
    import cloudscraper
    _SESSION = cloudscraper.create_scraper()
except ImportError:
    print("WARNING: cloudscraper not installed (pip install cloudscraper). "
          "IOI (cphof.org) will very likely return 0 rows without it, since "
          "the site sits behind a Cloudflare JS challenge that plain "
          "`requests` cannot pass.")
    _SESSION = requests
REQUEST_TIMEOUT = 30
MIN_DELAY, MAX_DELAY = 1.5, 3.0   # be polite between page fetches

OUTPUT_CSV = "olympiad_participants.csv"
FIELDS = ["subject", "olympiad", "year", "name", "country", "award",
          "rank", "profile_url", "source_url"]

# Award normalization: every source maps its raw label to one of these.
AWARD_GOLD, AWARD_SILVER, AWARD_BRONZE, AWARD_HM, AWARD_NONE = (
    "Gold", "Silver", "Bronze", "Honourable Mention", "Participant")


def throttle() -> None:
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def fetch(url: str) -> BeautifulSoup:
    resp = _SESSION.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return BeautifulSoup(resp.text, "html.parser")


# ----------------------------------------------------------------------------
# Base class
# ----------------------------------------------------------------------------

@dataclass
class OlympiadSource:
    # start_year/end_year have no default and must stay first: subclasses
    # below re-declare subject/olympiad WITH a default value, and dataclass
    # field order is inherited-position-based, not redeclaration-based. If
    # subject/olympiad came first here, a subclass giving them defaults
    # would push the still-default-less start_year/end_year after a
    # defaulted field, which dataclass rejects at class-creation time.
    start_year: int
    end_year: int
    subject: str = ""     # "Mathematics", "Physics", ... (set by subclass)
    olympiad: str = ""    # "IMO", "IPhO", ...            (set by subclass)

    def year_url(self, year: int) -> str:
        raise NotImplementedError

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        raise NotImplementedError

    def scrape_all(self) -> list[dict]:
        rows: list[dict] = []
        for year in range(self.start_year, self.end_year + 1):
            url = self.year_url(year)
            print(f"[{self.olympiad}] scraping {year} ...")
            try:
                soup = fetch(url)
                year_rows = self.parse_year(year, soup, url)
            except Exception as exc:
                print(f"  failed for {year}: {exc}")
                continue
            print(f"  {len(year_rows)} rows")
            rows.extend(year_rows)
            throttle()
        return rows


# ----------------------------------------------------------------------------
# IMO — Mathematics (imo-official.org)
# ----------------------------------------------------------------------------

@dataclass
class IMOSource(OlympiadSource):
    subject: str = "Mathematics"
    olympiad: str = "IMO"
    keep_awards: set = field(default_factory=lambda: {"G", "S", "B", "H"})

    def year_url(self, year: int) -> str:
        return f"https://www.imo-official.org/results/individual/year/{year}/"

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        award_name = {"G": AWARD_GOLD, "S": AWARD_SILVER, "B": AWARD_BRONZE,
                      "H": AWARD_HM}
        rows = []
        for tr in soup.find_all("tr"):
            name_a = tr.find("a", href=re.compile(r"/results/contestant/\d+"))
            if not name_a:
                continue
            name = name_a.get_text(" ", strip=True)
            profile_url = urljoin(url, name_a["href"])

            country_a = tr.find("a", href=re.compile(r"/individual/country/[A-Z]{3}"))
            country = ""
            if country_a:
                m = re.search(r"/country/([A-Z]{3})", country_a["href"])
                country = m.group(1) if m else ""

            award = ""
            for td in tr.find_all("td"):
                t = td.get_text(strip=True)
                if t in award_name:
                    award = t
                    break
            if award not in self.keep_awards:
                continue

            rows.append({
                "subject": self.subject, "olympiad": self.olympiad, "year": year,
                "name": name, "country": country, "award": award_name[award],
                "rank": "", "profile_url": profile_url, "source_url": url,
            })
        return rows


# ----------------------------------------------------------------------------
# IPhO — Physics (ipho-unofficial.org)
# ----------------------------------------------------------------------------

@dataclass
class IPhOSource(OlympiadSource):
    subject: str = "Physics"
    olympiad: str = "IPhO"

    def year_url(self, year: int) -> str:
        return f"https://ipho-unofficial.org/timeline/{year}/individual"

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        table = soup.find("table")
        rows = []
        if not table:
            return rows
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue
            name_cell, country_cell, rank_cell, award_cell = tds[0], tds[1], tds[2], tds[3]

            name_a = name_cell.find("a")
            name = (name_a.get_text(" ", strip=True) if name_a
                    else name_cell.get_text(" ", strip=True))
            profile_url = urljoin(url, name_a["href"]) if name_a else ""

            country_a = country_cell.find("a")
            country = country_a.get_text(" ", strip=True) if country_a else \
                country_cell.get_text(" ", strip=True)

            rank = rank_cell.get_text(" ", strip=True)

            award_text = award_cell.get_text(" ", strip=True).lower()
            if "gold" in award_text:
                award = AWARD_GOLD
            elif "silver" in award_text:
                award = AWARD_SILVER
            elif "bronze" in award_text:
                award = AWARD_BRONZE
            elif "honourable" in award_text or "honorable" in award_text:
                award = AWARD_HM
            else:
                continue  # header row or unrecognized row

            rows.append({
                "subject": self.subject, "olympiad": self.olympiad, "year": year,
                "name": name, "country": country, "award": award,
                "rank": rank, "profile_url": profile_url, "source_url": url,
            })
        return rows


# ----------------------------------------------------------------------------
# IChO — Chemistry (icho-official.org/results)
# ----------------------------------------------------------------------------
# The IChO site indexes olympiads by a numeric `id`, not directly by year,
# and the mapping (id <-> year) has to be resolved from the index page since
# id 1 = 1968 but three years (e.g. 1971) were skipped historically.

@dataclass
class IChOSource(OlympiadSource):
    subject: str = "Chemistry"
    olympiad: str = "IChO"
    _id_for_year: dict = field(default_factory=dict, init=False, repr=False)

    INDEX_URL = "https://www.icho-official.org/results/"

    def _load_year_ids(self) -> None:
        if self._id_for_year:
            return
        soup = fetch(self.INDEX_URL)
        for a in soup.find_all("a", href=re.compile(r"results\.php\?id=\d+&year=\d+")):
            m = re.search(r"id=(\d+)&year=(\d+)", a["href"])
            if m:
                self._id_for_year[int(m.group(2))] = int(m.group(1))

    def year_url(self, year: int) -> str:
        self._load_year_ids()
        icho_id = self._id_for_year.get(year)
        if icho_id is None:
            raise ValueError(f"no IChO edition found for year {year}")
        return f"https://www.icho-official.org/results/results.php?id={icho_id}&year={year}"

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        rows = []
        table = soup.find("table")
        if not table:
            return rows
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 5:
                continue
            name_cell, _script_cell, country_cell, rank_cell, award_cell = tds[:5]

            name = name_cell.get_text(" ", strip=True)
            if not name or name.lower() == "contestant":
                continue

            country_a = country_cell.find("a")
            country = country_a.get_text(" ", strip=True) if country_a else \
                country_cell.get_text(" ", strip=True)

            rank = rank_cell.get_text(" ", strip=True)

            award_text = award_cell.get_text(" ", strip=True).lower()
            if "gold" in award_text:
                award = AWARD_GOLD
            elif "silver" in award_text:
                award = AWARD_SILVER
            elif "bronze" in award_text:
                award = AWARD_BRONZE
            elif "honorable" in award_text or "honourable" in award_text:
                award = AWARD_HM
            elif "participant" in award_text:
                award = AWARD_NONE
            else:
                continue

            rows.append({
                "subject": self.subject, "olympiad": self.olympiad, "year": year,
                "name": name, "country": country, "award": award,
                "rank": rank, "profile_url": "", "source_url": url,
            })
        return rows


# ----------------------------------------------------------------------------
# IOI — Informatics (cphof.org — Competitive Programming Hall of Fame)
# ----------------------------------------------------------------------------
# The OFFICIAL IOI archive (stats.ioinformatics.org) disallows automated
# scraping via robots.txt, so this uses cphof.org instead: a fetchable,
# community-run standings archive with the same underlying results
# (rank / country / name / score / medal), one page per year at
# https://cphof.org/standings/ioi/{year}.

@dataclass
class IOISource(OlympiadSource):
    subject: str = "Informatics"
    olympiad: str = "IOI"

    def year_url(self, year: int) -> str:
        return f"https://cphof.org/standings/ioi/{year}"

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        rows = []
        table = soup.find("table")
        if not table:
            return rows
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue
            rank_cell, country_cell, name_cell, score_cell = tds[:4]

            name_a = name_cell.find("a", href=re.compile(r"/profile/ioi:\d+"))
            if not name_a:
                continue  # header row or malformed row
            name = name_a.get_text(" ", strip=True)
            profile_url = urljoin(url, name_a["href"])

            country_a = country_cell.find("a", href=re.compile(r"/country/"))
            country = country_a.get_text(" ", strip=True) if country_a else \
                country_cell.get_text(" ", strip=True)

            rank_text = rank_cell.get_text(" ", strip=True).lower()
            img = rank_cell.find("img")
            medal_alt = (img.get("alt", "") if img else "").lower()
            if "gold" in medal_alt:
                award = AWARD_GOLD
            elif "silver" in medal_alt:
                award = AWARD_SILVER
            elif "bronze" in medal_alt:
                award = AWARD_BRONZE
            else:
                award = AWARD_NONE  # "finalist" — no medal

            score = score_cell.get_text(" ", strip=True)

            rows.append({
                "subject": self.subject, "olympiad": self.olympiad, "year": year,
                "name": name, "country": country, "award": award,
                "rank": rank_text, "profile_url": profile_url, "source_url": url,
            })
        return rows


# ----------------------------------------------------------------------------
# Generic fallback for olympiads without a uniform archive
# (IBO, IOAA, regional olympiads: APhO, EGMO, Balkan MO, Asian Physics
# Olympiad, etc.) You supply the per-year URL and which table-column index
# holds each field; this scraper does the rest.
# ----------------------------------------------------------------------------

@dataclass
class GenericTableSource(OlympiadSource):
    """
    year_urls: {year: url, ...}           -- you must find these yourself
    columns:   {"name": 0, "country": 1, "rank": 2, "award": 3}
               -- 0-based <td> index within each results row; omit a key
               if that field isn't on the page.
    award_keywords: mapping of lowercase substring -> normalized award,
               e.g. {"gold": AWARD_GOLD, "silver": AWARD_SILVER, ...}
    """
    year_urls: dict = field(default_factory=dict)
    columns: dict = field(default_factory=dict)
    award_keywords: dict = field(default_factory=dict)

    def year_url(self, year: int) -> str:
        if year not in self.year_urls:
            raise ValueError(f"no URL configured for {self.olympiad} {year}")
        return self.year_urls[year]

    def parse_year(self, year: int, soup: BeautifulSoup, url: str) -> list[dict]:
        rows = []
        table = soup.find("table")
        if not table:
            return rows
        name_i = self.columns.get("name")
        country_i = self.columns.get("country")
        rank_i = self.columns.get("rank")
        award_i = self.columns.get("award")
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if name_i is None or len(tds) <= name_i:
                continue
            name = tds[name_i].get_text(" ", strip=True)
            if not name:
                continue
            country = tds[country_i].get_text(" ", strip=True) if country_i is not None and len(tds) > country_i else ""
            rank = tds[rank_i].get_text(" ", strip=True) if rank_i is not None and len(tds) > rank_i else ""
            award = AWARD_NONE
            if award_i is not None and len(tds) > award_i:
                text = tds[award_i].get_text(" ", strip=True).lower()
                for kw, norm in self.award_keywords.items():
                    if kw in text:
                        award = norm
                        break
            rows.append({
                "subject": self.subject, "olympiad": self.olympiad, "year": year,
                "name": name, "country": country, "award": award,
                "rank": rank, "profile_url": "", "source_url": url,
            })
        return rows

    def scrape_all(self) -> list[dict]:
        rows = []
        for year in sorted(self.year_urls):
            url = self.year_url(year)
            print(f"[{self.olympiad}] scraping {year} ...")
            try:
                soup = fetch(url)
                year_rows = self.parse_year(year, soup, url)
            except Exception as exc:
                print(f"  failed for {year}: {exc}")
                continue
            print(f"  {len(year_rows)} rows")
            rows.extend(year_rows)
            throttle()
        return rows


# ----------------------------------------------------------------------------
# Which sources to actually run
# ----------------------------------------------------------------------------

START_YEAR = 2008
END_YEAR = 2016

SOURCES: list[OlympiadSource] = [

    IOISource(start_year=START_YEAR, end_year=END_YEAR),

    # Example of how to bolt on an olympiad with no uniform archive.
    # Fill in real URLs/columns yourself; disabled (empty year_urls) by default.
    # GenericTableSource(
    #     subject="Biology", olympiad="IBO",
    #     start_year=START_YEAR, end_year=END_YEAR,
    #     year_urls={
    #         2016: "https://example-ibo-2016-results-page/",
    #     },
    #     columns={"name": 0, "country": 1, "rank": 2, "award": 3},
    #     award_keywords={"gold": AWARD_GOLD, "silver": AWARD_SILVER,
    #                      "bronze": AWARD_BRONZE, "honou": AWARD_HM},
    # ),
]


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def load_done_keys() -> set[tuple]:
    done = set()
    if os.path.exists(OUTPUT_CSV):
        with open(OUTPUT_CSV, encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                done.add((row["olympiad"], row["year"], row["name"], row["country"]))
    return done


def main() -> None:
    done = load_done_keys()
    write_header = not os.path.exists(OUTPUT_CSV)
    total_written = 0

    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, extrasaction="ignore")
        if write_header:
            w.writeheader()

        for source in SOURCES:
            rows = source.scrape_all()
            for r in rows:
                key = (r["olympiad"], str(r["year"]), r["name"], r["country"])
                if key in done:
                    continue
                w.writerow(r)
                done.add(key)
                total_written += 1
            fh.flush()

    print(f"done. wrote {total_written} new rows to {OUTPUT_CSV}")


if __name__ == "__main__":
    main()