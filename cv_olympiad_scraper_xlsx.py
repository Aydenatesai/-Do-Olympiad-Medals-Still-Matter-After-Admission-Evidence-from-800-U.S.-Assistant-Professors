#!/usr/bin/env python3
"""
cv_olympiad_scraper_xlsx.py  (v4 — RAW education snippet, not parsed institution/year)

WHAT CHANGED IN v4 (per request: stop trying to parse out a clean
"institution + year" — it's slow to debug and error-prone. Instead, just
find the "Education" section of the CV/page and copy the RAW TEXT of that
section straight into the spreadsheet. A human can read three lines of
raw CV text far faster than I can write a regex that never breaks.)
------------------------------------------------------------------------
  - REMOVED the whole institution-name-parsing pipeline: INSTITUTION_RE,
    _extract_institution_from_window(), _clean_institution_name(),
    _ATTRIBUTION_TRIGGER_RE / _strip_attribution_names(),
    _ACRONYM_SCHOOL_RE, _find_undergrad_in_education_section(),
    _find_undergrad_by_degree_keyword(), find_undergrad(). All of that
    complexity existed only to produce a clean "Institution (Year)"
    string, which was the single biggest source of wrong/messy results
    and the slowest thing to debug. Gone.
  - NEW: find_education_snippet(text) — finds the Education section (or,
    if there's no clean header, a window around the first Bachelor's-
    level degree keyword it can find) and returns the RAW joined text of
    that section, capped at a sane length. This is what gets written to
    the spreadsheet now — read it yourself, no parsing required.
  - A lightweight year-scrape (just "smallest 4-digit year mentioned in
    that raw snippet") is still done, ONLY so the existing "olympiad
    before college" comparison still has something to compare against.
    This is NOT trying to identify which year is "the" start year in any
    smart way — it's a rough signal, on purpose, in keeping with "just
    grab the snippet and let a human look at it."
  - NEW: has_education_signal(text) — quick yes/no check (does this text
    contain an Education header OR a Bachelor's-level degree keyword
    anywhere?). Used specifically for **website**-sourced fetches (see
    below) — NOT for local cv_download files, which are trusted to
    contain a real CV already.
  - Per your instruction, "website" type entries now follow this rule:
    fetch the page, run has_education_signal() on it. If there's no
    education signal at all, DON'T try to dig further — mark the row
    "No education signal on page — skipped" and move on. This avoids
    burning time/requests trying to squeeze education info out of pages
    that clearly don't have any (e.g. a bare publications list, a
    seminar-talk abstract page, etc.).
  - "cv_download" (local file) entries are NOT subject to that skip rule
    — those are actual downloaded CVs, so we always try to extract a
    snippet from them, even if the Education header detection is a bit
    fuzzy on that particular file.
  - Still covers ALL FIVE subject tabs automatically: --direct-mode
    processes every sheet in the workbook whose name starts with
    "CV or Website" (i.e. "CV or Website - Math", "... - CS",
    "... - Physics", "... - Engineer", "... - Biochem") — this was
    already true in v3 and is unchanged; nothing extra needed for that
    part, it "just works" across all 5 tabs in one run.
  - Output columns changed to match the new approach:
       "Education Section (Raw)"     -- the raw snippet, verbatim
       "Earliest Year In Snippet"    -- rough year signal (see above)
       "Olympiad Before College"     -- unchanged logic, just now
                                          compared against the rough
                                          year instead of a parsed one
  - Olympiad detection itself (find_olympiads) is UNCHANGED — that part
    wasn't the problem, so it's left exactly as it was.
  - Everything else (local-folder indexing/matching, retries, network
    fetch, autosave, --verbose-log, CLI args) is UNCHANGED from v3.

Requirements
------------
    pip install openpyxl requests beautifulsoup4 pdfplumber lxml
    # Optional, only needed if any of your local CVs are .docx files:
    pip install python-docx

Usage (same as v3)
-------------------
Process the cv_or_website_results.xlsx workbook directly — every
"CV or Website - <Subject>" sheet (Math, CS, Physics, Engineer, Biochem)
gets processed in one run:

    python cv_olympiad_scraper_xlsx.py --direct-mode

    # test on a handful of rows per sheet first:
    python cv_olympiad_scraper_xlsx.py --direct-mode --limit 5

    # write a companion CSV with full detail for manual review:
    python cv_olympiad_scraper_xlsx.py --direct-mode --verbose-log matches.csv

Python version note
--------------------
This script targets Python 3.8+ . The `from __future__ import annotations`
line right below is required for that — without it, type hints like
`list[str]` (which only work natively as of Python 3.9) would raise
`TypeError: 'type' object is not subscriptable` on 3.8. Do not remove it.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import csv
import os
import unicodedata
import traceback
import difflib
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Optional
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup
import openpyxl

try:
    import pdfplumber
    HAVE_PDF = True
except ImportError:
    HAVE_PDF = False

try:
    import docx  # python-docx
    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

SLEEP_SECONDS = 3.0
REQUEST_TIMEOUT = 20

# EXPERIMENTAL: pdfplumber's default word-spacing tolerance sometimes loses
# spaces between words on certain PDF fonts/layouts (producing text like
# "ofScienceinMathematics..."). normalize_extracted_text() below tries to
# repair that after the fact. Left as None to use pdfplumber's default
# unless you set it.
PDF_X_TOLERANCE = None
MAX_RETRIES = 3
RETRY_BACKOFF_BASE = 5
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
BROWSER_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

LOCAL_CV_EXTENSIONS = {".pdf", ".docx", ".txt", ".md", ".html", ".htm"}

# Words to ignore when comparing a person's name against filenames.
FILENAME_STOPWORDS = {
    "cv", "resume", "curriculum", "vitae", "phd", "dr", "prof", "professor",
    "final", "updated", "latest", "new", "copy", "v1", "v2", "v3",
}

OLYMPIAD_KEYWORDS = [
    "International Mathematical Olympiad", "IMO",
    "International Olympiad in Informatics", "IOI",
    "International Physics Olympiad", "IPhO",
    "International Chemistry Olympiad", "IChO",
    "International Biology Olympiad", "IBO",
    "International Linguistics Olympiad", "IOL",
    "International Astronomy Olympiad", "IAO",
    "USA Mathematical Olympiad", "USAMO",
    "USA Physics Olympiad", "USAPhO",
    "USA Biology Olympiad", "USABO",
    "US National Chemistry Olympiad", "USNCO",
    "American Regions Mathematics League", "ARML",
    "American Invitational Mathematics Examination", "AIME",
    "Science Olympiad",
    "Math Olympiad",
]

# Short acronyms (no spaces, e.g. "IMO", "IChO", "ARML") are matched with
# EXACT case only — case-insensitive matching on a 3-5 letter acronym risks
# false positives (e.g. lowercase "imo" as in "in my opinion"). Full
# descriptive names are safe to match case-insensitively.
_ACRONYM_KEYWORDS = [k for k in OLYMPIAD_KEYWORDS if " " not in k]
_FULLNAME_KEYWORDS = [k for k in OLYMPIAD_KEYWORDS if " " in k]

_ACRONYM_RE = re.compile(
    r"\b(" + "|".join(re.escape(k) for k in _ACRONYM_KEYWORDS) + r")\b"
)  # no re.IGNORECASE -- exact case required
_FULLNAME_RE = re.compile(
    r"\b(" + "|".join(re.escape(k) for k in _FULLNAME_KEYWORDS) + r")\b",
    re.IGNORECASE,
)

# Kept ONLY as the fallback trigger for "is there a Bachelor's-level degree
# keyword anywhere in this text" (used by has_education_signal() and by
# find_education_snippet()'s no-header fallback). No longer used to try to
# extract an institution name — see module docstring.
UNDERGRAD_DEGREE_PATTERNS = [
    r"\bB\.?S\.?(?![a-z])", r"\bB\.?A\.?(?![a-z])",
    r"\bB\.?Eng\.?(?![a-z])", r"\bB\.?Tech\.?(?![a-z])",
    r"\bB\.?E\.?(?![a-z])",    # bare "B.E." (Bachelor of Engineering short
                                 # form) -- DISTINCT from "B.Eng." above.
                                 # Confirmed missed on a real faculty bio:
                                 # "my M.S. and B.E. from Fudan University
                                 # and Nanjing University, respectively."
    r"\bS\.?B\.?(?![a-z])",
    r"\bSc\.?B\.?(?![a-z])",   # Brown / Cornell style ("Sc.B.")
    r"\bA\.?B\.?(?![a-z])",    # Harvard style ("A.B.")
    r"\bB\.?Sc\.?(?![a-z])",   # some non-US CVs
    r"\bB\.?A\.?Sc\.?(?![a-z])",  # BASc (Bachelor of Applied Science,
                                    # standard in Canadian engineering
                                    # programs) -- confirmed missed on a
                                    # real CV ("2011 BASc in Engineering
                                    # Science, University of Toronto").
                                    # This list had drifted out of sync
                                    # with extract_college's own DEGREE_RE,
                                    # which already had this pattern --
                                    # the two lists must stay in sync,
                                    # since THIS one gates whether a fetched
                                    # page's text ever gets saved anywhere
                                    # at all (see has_education_signal()
                                    # below): a pattern missing here means
                                    # the row gets marked "no education
                                    # signal" and the real fetched text is
                                    # discarded outright, with nothing left
                                    # to even retry from later.
    r"\bB\.?Math\.?(?![a-z])",     # Waterloo-style "Bachelor of Mathematics"
    r"\bBachelor(?:'s)?(?:\s+of\s+\w+)?",
    r"\bUndergraduate\b",
    # -------------------------------------------------------------------
    # INFORMAL, PROSE-STYLE mentions -- confirmed missed on real faculty
    # "About me" / "Biography" / "Personal" pages that have NO "Education"
    # header at all and never write a formal degree abbreviation, e.g.:
    #   "I loved spending my college years at MIT..."
    #   "I am a graduate of NYU and Columbia University."
    # These are lower-precision than the formal abbreviations above (e.g.
    # "graduate of" can occasionally refer to a grad program rather than
    # undergrad) -- but this tool's whole point is "surface a snippet for
    # a human to read", not assert a parsed answer, so an occasional
    # slightly-off snippet is a far smaller cost than silently missing
    # real content, which is what was happening before these were added.
    r"\bcollege\s+years\b",
    r"\bas\s+an\s+undergraduate\b",
    r"\bundergraduate\s+(?:degree|years|study|studies)\b",
    r"\bundergraduate\s+at\b",
    r"\bgraduate\s+of\b",
    r"\battended\b[^.]{0,40}\bfor\s+(?:my|his|her|their)\s+undergraduate\b",
]
_DEGREE_KEYWORD_RE = re.compile(
    r"(" + "|".join(UNDERGRAD_DEGREE_PATTERNS) + r")", re.IGNORECASE
)
# Glued ALL-CAPS format companion check, e.g. "BSCHEMISTRY", "BAECONOMICS"
# -- mirrors DEGREE_GLUED_CAPS_RE in the ported college-extraction module
# below, kept as a SEPARATE, case-sensitive pattern for the same reason
# explained there (a case-insensitive version would false-positive on
# ordinary words like "Beckman"). Needed here too, not just in the
# College-extraction logic -- without it, a page whose ONLY degree mention
# is in this glued-caps style would still get marked "no education
# signal" and discarded before extract_college() ever gets a chance to
# run on it.
_DEGREE_GLUED_CAPS_SIGNAL_RE = re.compile(
    r"\bB\.?S\.?(?=[A-Z]{6,})|\bB\.?A\.?(?=[A-Z]{6,})|\bB\.?E\.?(?=[A-Z]{6,})"
)


YEAR_RE = re.compile(r"\b(19[5-9]\d|20[0-4]\d)\b")

SHEET_NAME = "research info"
COL_NAME = "Name"
COL_SCHOOL = "School"
COL_CV_URL = "CV URL"
NEW_COL_EDU_SNIPPET = "Education Section (Raw)"
NEW_COL_EARLIEST_YEAR = "Earliest Year In Snippet"
NEW_COL_OLYMPIAD = "Olympiad Before College"
NEW_COL_COLLEGE = "College"

# --------------------------------------------------------------------------
# CONFIG -- edit these paths directly so you can just run this script with
# no command-line arguments at all.
# --------------------------------------------------------------------------
DEFAULT_INPUT_XLSX = r"D:\Research 2026\cvs\Math.xlsx"
DEFAULT_OUTPUT_XLSX = r"D:\Research 2026\cvs\Math - enriched.xlsx"
DEFAULT_SHEET = SHEET_NAME
DEFAULT_CV_FOLDER = r"D:\Research 2026\cvs\Math"

DEFAULT_FALLBACK_XLSX = r"D:\Research 2026\cvs\cv_or_website_results.xlsx"

# Used by --direct-mode: process cv_or_website_results.xlsx ITSELF, across
# EVERY sheet starting with "CV or Website" -- i.e. all 5 subject tabs
# (Math, CS, Physics, Engineer, Biochem) in one run.
DEFAULT_DIRECT_MODE_XLSX = r"D:\Research 2026\cvs\cv_or_website_results.xlsx"
FALLBACK_SHEET_PREFIX = "CV or Website"
FALLBACK_COL_NAME = "name"
FALLBACK_COL_RESULT_TYPE = "result_type"
FALLBACK_COL_VALUE = "value"
FALLBACK_COL_SOURCE_URL = "source_url"

DEBUG = False


def log_error(prefix: str, e: Exception):
    if DEBUG:
        print(f"{prefix}: {e}", file=sys.stderr)
        traceback.print_exc()
    else:
        print(f"{prefix}: {e} (run with --debug for full traceback)", file=sys.stderr)


# --------------------------------------------------------------------------
# Local CV folder indexing + name matching (unchanged from v3)
# --------------------------------------------------------------------------

def normalize_text(s: str) -> str:
    """Lowercase, strip accents, keep only letters/spaces."""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z\s]", " ", s)
    return s.lower()


def name_tokens(s: str) -> list[str]:
    return [t for t in normalize_text(s).split() if t and t not in FILENAME_STOPWORDS]


def full_name_similarity(person_name: str, filename_stem: str) -> float:
    a = normalize_text(person_name).replace(" ", "")
    b = normalize_text(filename_stem).replace(" ", "")
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _looks_like_url(s: str) -> bool:
    return bool(s) and (s.startswith("http://") or s.startswith("https://"))


def _looks_like_file_uri(s: str) -> bool:
    return bool(s) and s.lower().startswith("file://")


def _file_uri_to_path(uri: str) -> Path:
    """Converts a 'file:///C:/Users/HP/Desktop/cv.pdf' or
    'file://d:/Research 2026/cvs/x.pdf' style URI (confirmed present in
    real data -- a browser 'copy link' on a local file that got pasted
    into the spreadsheet) into an actual local Path. `urllib.parse` +
    manual Windows-drive-letter handling, since Python's own
    `pathlib.Path.from_uri` (3.13+) can't be relied on for an unknown
    local Python version."""
    from urllib.parse import urlparse, unquote
    parsed = urlparse(uri)
    raw_path = unquote(parsed.path)
    # A file URI with only TWO slashes after the scheme ("file://d:/...",
    # missing the standard third slash that "file:///d:/..." would have)
    # gets misparsed by urlparse: "d:" lands in `netloc`, not `path`,
    # silently DROPPING the drive letter entirely if left unhandled.
    # Confirmed as a real case in actual spreadsheet data. Detect a
    # drive-letter-shaped netloc ("d:", "C:", etc.) and reattach it to
    # the front of the path.
    if parsed.netloc and len(parsed.netloc) == 2 and parsed.netloc[1] == ":":
        raw_path = f"{parsed.netloc}{raw_path}"
    # A Windows-style file URI often parses with a leading "/" before the
    # drive letter (e.g. "/C:/Users/..." or "/d:/Research 2026/..."),
    # which isn't a valid Windows path -- strip it when the next 2 chars
    # look like a drive letter + colon.
    if len(raw_path) > 2 and raw_path[0] == "/" and raw_path[2] == ":":
        raw_path = raw_path[1:]
    return Path(raw_path)


def _find_file_by_basename_under_root(missing_path: Path, search_root: Path) -> Optional[Path]:
    """
    Fallback for a "cv_download" row whose recorded path doesn't exist on
    disk. Confirmed as a real, common cause -- the spreadsheet's `value`
    column records paths like "d:\\Research 2026\\cvs\\Computer
    Science\\Aayush_Jain.pdf", but the actual folder on disk may use a
    different naming convention for the subject subfolder (e.g. "CS"
    instead of the spelled-out "Computer Science") -- a folder rename or
    reorganization that happened after these paths were originally
    recorded. The FILENAME itself is much more likely to still be
    accurate than the exact subfolder path, so: search recursively for a
    file with the exact same basename anywhere under a shared root
    (2 levels up from the missing file, e.g. "d:\\Research 2026\\cvs\\"),
    and use it if found. Returns None if nothing matches, or if
    `search_root` doesn't exist / isn't a real ancestor -- never guesses
    across an unrelated directory tree.
    """
    if not search_root.exists() or not search_root.is_dir():
        return None
    basename = missing_path.name
    try:
        matches = list(search_root.rglob(basename))
    except Exception:
        return None
    if len(matches) == 1:
        return matches[0]
    # Multiple files with the same exact filename found -- too risky to
    # guess which one is right (could be a genuine name collision across
    # different subjects); leave it as a real "file not found" instead of
    # silently picking one.
    return None


def read_cv_download_value(value: str) -> str:
    """
    Handles a "cv_download" row's value -- which is SUPPOSED to always be
    a local file path, but real data confirmed some rows are mislabeled:
    result_type says "cv_download" while `value` actually contains a URL
    (e.g. "https://allen.physics.ucsd.edu/people/Allen_CV_v10.pdf")
    instead of a genuine local path like "d:\\Research 2026\\cvs\\....pdf".

    Blindly doing Path(value) on a URL string produces exactly the
    confirmed real-world failure: on Windows, Path() treats "https:" like
    a drive-ish prefix and turns the rest into backslash-separated path
    segments (e.g. "https:\\allen.physics.u..."), which obviously doesn't
    exist on disk -> "[Errno 22] Invalid argument" for every single one of
    these mislabeled rows, even though the CV is perfectly fetchable as a
    normal URL.

    Fix: check whether `value` actually LOOKS like a URL before deciding
    how to read it, regardless of what result_type claims. This function
    dispatches to the right reader either way and returns the extracted
    text, so callers don't need to duplicate this check.

    Also handles a "file://" URI value (a browser 'copy link' on a local
    file) by converting it to a real local path first, and falls back to
    _find_file_by_basename_under_root() when the recorded exact path
    doesn't exist -- see that function's docstring for why.
    """
    if _looks_like_url(value):
        text = fetch_text_from_url(value)
        time.sleep(SLEEP_SECONDS)  # only sleep when we actually hit the network
        return text

    path = _file_uri_to_path(value) if _looks_like_file_uri(value) else Path(value)

    if path.exists():
        return read_local_file(path)

    # Exact path doesn't exist -- try the basename-anywhere-under-root
    # fallback before giving up. `search_root` is 2 levels up from the
    # file itself (skips the immediate, possibly-renamed subject
    # subfolder, e.g. "...\\cvs\\Computer Science\\x.pdf" ->
    # "...\\cvs\\"). Falls back to 1 level up if the path is too shallow
    # to have 2 parents above the file.
    search_root = path.parent.parent if path.parent.parent != path.parent else path.parent
    fallback = _find_file_by_basename_under_root(path, search_root)
    if fallback is not None:
        print(f"    [fallback] recorded path missing, found by filename instead: {fallback}")
        return read_local_file(fallback)

    # Genuinely not found anywhere -- raise the same kind of error the
    # direct Path().open() attempt would have, so callers' existing
    # FileNotFoundError handling still works unchanged.
    raise FileNotFoundError(f"[Errno 2] No such file or directory (also not found by "
                             f"filename under {search_root}): '{value}'")


def _resolve_cell_value(cell) -> str:
    """
    Return the best usable string for a cell: if the cell's plain text
    value doesn't look like a usable URL (e.g. it's a hyperlink DISPLAY
    LABEL like "webb-CV.pdf - Google Drive" rather than the actual link),
    but the cell has a real Excel hyperlink attached, use the hyperlink's
    TARGET instead. Falls back to the plain text value (even if it's not
    URL-shaped, e.g. a local file path for cv_download rows) if there's no
    hyperlink or the hyperlink's target isn't URL-shaped either.
    """
    raw = str(cell.value).strip() if cell.value else ""
    if _looks_like_url(raw):
        return raw
    if cell.hyperlink and cell.hyperlink.target and _looks_like_url(cell.hyperlink.target):
        return cell.hyperlink.target.strip()
    return raw


def load_fallback_results(fallback_xlsx_path: str) -> dict:
    result = {}
    if not fallback_xlsx_path or not os.path.exists(fallback_xlsx_path):
        if fallback_xlsx_path:
            print(f"[fallback-results] file not found, skipping: {fallback_xlsx_path}")
        return result

    try:
        wb = openpyxl.load_workbook(fallback_xlsx_path, data_only=True)
    except Exception as e:
        print(f"[fallback-results] could not open {fallback_xlsx_path}: {e}")
        return result

    matching_sheets = [s for s in wb.sheetnames if s.startswith(FALLBACK_SHEET_PREFIX)]
    if not matching_sheets:
        print(f"[fallback-results] no sheets starting with '{FALLBACK_SHEET_PREFIX}' found in {fallback_xlsx_path}")
        return result

    for sheet_name in matching_sheets:
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        try:
            name_col = headers.index(FALLBACK_COL_NAME) + 1
            type_col = headers.index(FALLBACK_COL_RESULT_TYPE) + 1
            value_col = headers.index(FALLBACK_COL_VALUE) + 1
        except ValueError:
            print(f"[fallback-results] sheet '{sheet_name}' missing expected columns; skipping this sheet")
            continue
        source_col = headers.index(FALLBACK_COL_SOURCE_URL) + 1 if FALLBACK_COL_SOURCE_URL in headers else None

        for row in ws.iter_rows(min_row=2):
            name_cell = row[name_col - 1]
            if not name_cell.value:
                continue
            rtype = row[type_col - 1].value
            rtype_norm = str(rtype).strip().lower() if rtype else ""
            if rtype_norm == "website":
                value = _resolve_cell_value(row[value_col - 1])
            else:
                value = str(row[value_col - 1].value).strip() if row[value_col - 1].value else ""
            source_url = _resolve_cell_value(row[source_col - 1]) if source_col else None
            key = normalize_text(str(name_cell.value)).strip()
            if key:
                result[key] = {
                    "result_type": rtype_norm,
                    "value": value,
                    "source_url": source_url or "",
                }

    print(f"[fallback-results] loaded {len(result)} people from {len(matching_sheets)} sheet(s): {matching_sheets}")
    return result


def build_local_cv_index(folder: str) -> list[tuple[Path, list[str]]]:
    index = []
    root = Path(folder)
    if not root.exists():
        raise FileNotFoundError(f"--cv-folder path does not exist: {folder}")
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in LOCAL_CV_EXTENSIONS:
            tokens = name_tokens(path.stem)
            index.append((path, tokens))
    return index


def find_local_cv(
    person_name: str,
    index: list[tuple[Path, list[str]]],
    surname_counts: Optional[dict] = None,
) -> tuple[Optional[Path], str]:
    """Returns (path_or_None, status) where status is one of:
    'matched', 'no_local_match', 'ambiguous'. See v3 docstring history for
    full reasoning -- unchanged in v4."""
    p_tokens = name_tokens(person_name)
    if not p_tokens:
        return None, "no_local_match"

    surname = p_tokens[-1]
    other_tokens = set(p_tokens[:-1])

    candidates = []
    for path, f_tokens in index:
        f_token_set = set(f_tokens)
        if surname not in f_token_set:
            continue
        score = 0
        for t in other_tokens:
            if t in f_token_set:
                score += 2
            elif any(ft.startswith(t[0]) for ft in f_token_set if len(t) > 0):
                score += 1
        similarity = full_name_similarity(person_name, path.stem)
        candidates.append((score, similarity, path))

    if not candidates:
        return None, "no_local_match"

    surname_is_shared = bool(surname_counts) and surname_counts.get(surname, 1) > 1

    if len(candidates) == 1:
        if candidates[0][0] > 0 or not surname_is_shared:
            return candidates[0][2], "matched"
        return None, "ambiguous"

    positive = [c for c in candidates if c[0] > 0]
    if not positive:
        return None, "no_local_match"

    positive.sort(key=lambda c: (c[0], c[1]), reverse=True)
    top_score = positive[0][0]
    top_matches = [c for c in positive if c[0] == top_score]

    if len(top_matches) > 1:
        top_matches.sort(key=lambda c: c[1], reverse=True)
        best_sim = top_matches[0][1]
        second_sim = top_matches[1][1]
        if best_sim - second_sim >= 0.15:
            return top_matches[0][2], "matched"
        return None, "ambiguous"

    return top_matches[0][2], "matched"


# --------------------------------------------------------------------------
# Reading local files (PDF / DOCX / TXT / HTML) -- unchanged from v3
# --------------------------------------------------------------------------

def _extract_pdf_page_text(page) -> str:
    if PDF_X_TOLERANCE is not None:
        return page.extract_text(x_tolerance=PDF_X_TOLERANCE) or ""
    return page.extract_text() or ""


def normalize_extracted_text(text: str) -> str:
    """Repairs PDFs whose extraction drops spaces between words on
    certain fonts/layouts. See v3 docstring for the full explanation and
    known limitation (can't fix two glued lowercase words with no
    internal capital)."""
    out_lines = []
    for line in text.splitlines():
        words = line.split()
        if words:
            avg_word_len = sum(len(w) for w in words) / len(words)
            if avg_word_len >= 12:
                line = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", line)
        out_lines.append(line)
    return "\n".join(out_lines)


def read_local_file(path: Path) -> str:
    suffix = path.suffix.lower()

    if suffix == ".pdf":
        if not HAVE_PDF:
            raise RuntimeError("pdfplumber not installed — run: pip install pdfplumber")
        text_chunks = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text_chunks.append(_extract_pdf_page_text(page))
        return normalize_extracted_text("\n".join(text_chunks))

    if suffix == ".docx":
        if not HAVE_DOCX:
            raise RuntimeError("python-docx not installed — run: pip install python-docx")
        d = docx.Document(str(path))
        return normalize_extracted_text("\n".join(p.text for p in d.paragraphs))

    if suffix in (".html", ".htm"):
        raw = path.read_text(encoding="utf-8", errors="ignore")
        soup = BeautifulSoup(raw, "lxml")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        return soup.get_text(separator="\n")

    return path.read_text(encoding="utf-8", errors="ignore")


# --------------------------------------------------------------------------
# Networking helpers -- unchanged from v3
# --------------------------------------------------------------------------

def request_with_retries(method: str, url: str, **kwargs) -> requests.Response:
    kwargs.setdefault("timeout", REQUEST_TIMEOUT)
    headers = kwargs.setdefault("headers", {})
    for k, v in BROWSER_HEADERS.items():
        headers.setdefault(k, v)

    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.request(method, url, **kwargs)
            if resp.status_code == 429 or resp.status_code >= 500:
                raise requests.HTTPError(f"HTTP {resp.status_code}")
            resp.raise_for_status()
            return resp
        except Exception as e:
            last_exc = e
            if attempt < MAX_RETRIES:
                wait = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                print(f"    [retry {attempt}/{MAX_RETRIES}] {url[:70]}... "
                      f"failed ({e}); waiting {wait}s", file=sys.stderr)
                time.sleep(wait)
    raise last_exc


# --------------------------------------------------------------------------
# Google Drive share-link handling
# --------------------------------------------------------------------------
#
# THE BUG THIS SECTION FIXES: a normal Google Drive "share" URL --
# https://drive.google.com/file/d/<FILE_ID>/view?usp=sharing -- does NOT
# serve the file's content over a plain GET request. It serves a mostly-
# empty HTML shell page that uses JavaScript to render the actual viewer
# client-side. `requests` never executes JavaScript, so fetching that URL
# always returns that same empty "Loading..." shell, no matter what file
# it points to -- which is exactly why has_education_signal() found
# nothing on virtually every Drive-linked row (confirmed against a real
# run: dozens of clearly-legitimate personal/CV pages all showing "No
# education signal" purely because their `value` was a Drive share link,
# not because the underlying CV genuinely lacked an Education section).
#
# THE FIX: Drive exposes a separate, plain-HTTP-fetchable "direct
# download" endpoint for any file whose sharing is set to "Anyone with
# the link can view": https://drive.google.com/uc?export=download&id=
# <FILE_ID> . That endpoint returns the raw file bytes directly (no JS
# needed) for small files. For LARGER files (Drive's threshold, roughly
# 25-100MB+ depending on file type -- CVs are never this big, but some
# people share a whole portfolio/zip this way), Drive instead returns an
# HTML "Google Drive can't scan this file for viruses" warning page with
# a hidden confirmation token embedded in it; you have to resubmit the
# request with that token to get the real file. Both cases are handled
# below. Files that are NOT shared "Anyone with the link" (still
# restricted to specific people) can't be fetched by any of this --
# those will still come back empty, but that's a genuine permissions
# wall, not a bug.
_GDRIVE_FILE_ID_RE = re.compile(
    r"drive\.google\.com/(?:file/d/|open\?id=|uc\?.*?[?&]id=)([a-zA-Z0-9_-]{10,})"
)
_GDRIVE_CONFIRM_TOKEN_RE = re.compile(r'confirm=([0-9A-Za-z_-]+)')


def _is_google_drive_url(url: str) -> bool:
    return "drive.google.com" in url.lower()


def _resolve_google_drive_download_url(url: str) -> Optional[str]:
    """Extract the file ID from any recognizable Drive share-link shape
    and build the direct-download URL. Returns None if no file ID could
    be found (not a recognizable Drive file URL -- e.g. a folder link)."""
    m = _GDRIVE_FILE_ID_RE.search(url)
    if not m:
        return None
    file_id = m.group(1)
    return f"https://drive.google.com/uc?export=download&id={file_id}"


def _fetch_google_drive_file(url: str) -> requests.Response:
    """
    Fetch a Google Drive file's actual bytes, working around both the
    JS-only share-link viewer AND the large-file virus-scan-warning
    interstitial page. Raises the same exceptions request_with_retries
    would on genuine network/HTTP failure. Returns the final response
    (the real file content) on success.
    """
    download_url = _resolve_google_drive_download_url(url)
    if download_url is None:
        # Not a file-shaped Drive URL (e.g. a folder) -- nothing we can do,
        # fetch it as-is and let the caller's normal handling take over
        # (will likely yield an empty/JS-shell page, same as before).
        return request_with_retries("GET", url)

    session = requests.Session()
    resp = session.get(download_url, headers=BROWSER_HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "").lower()
    if "text/html" in content_type and len(resp.content) < 200_000:
        # Small HTML response after requesting a "download" -- almost
        # certainly the virus-scan-warning interstitial rather than the
        # actual file. Look for the confirm token and retry with it.
        text_preview = resp.text
        token_match = _GDRIVE_CONFIRM_TOKEN_RE.search(text_preview)
        if token_match:
            confirmed_url = f"{download_url}&confirm={token_match.group(1)}"
            resp = session.get(confirmed_url, headers=BROWSER_HEADERS, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
    return resp


def fetch_text_from_url(url: str) -> str:
    if _looks_like_file_uri(url):
        # A "file://" URI (a browser 'copy link' on a local file) landed
        # in a "website"-type row's value -- confirmed as a real, distinct
        # failure mode from the "cv_download" mislabeling handled
        # elsewhere: `requests` has no adapter registered for the "file"
        # scheme at all, so passing this straight to request_with_retries
        # always failed with "No connection adapters were found for
        # 'file:///...'" no matter how many retries. Route it through the
        # local-file path instead of a network request.
        return read_cv_download_value(url)

    if _is_google_drive_url(url):
        resp = _fetch_google_drive_file(url)
    else:
        resp = request_with_retries("GET", url)

    content_type = resp.headers.get("Content-Type", "").lower()
    is_pdf = "pdf" in content_type or url.lower().endswith(".pdf")
    # Google's download endpoint often doesn't set a helpful Content-Type
    # for PDFs the way a normal web server would -- fall back to sniffing
    # the actual file signature ("%PDF-" magic bytes) when we fetched via
    # the Drive path and the header-based check above didn't already
    # decide it's a PDF.
    if not is_pdf and _is_google_drive_url(url):
        is_pdf = resp.content[:5] == b"%PDF-"

    if is_pdf:
        if not HAVE_PDF:
            raise RuntimeError("PDF detected but pdfplumber isn't installed.")
        text_chunks = []
        with pdfplumber.open(BytesIO(resp.content)) as pdf:
            for page in pdf.pages:
                text_chunks.append(_extract_pdf_page_text(page))
        return normalize_extracted_text("\n".join(text_chunks))

    soup = BeautifulSoup(resp.text, "lxml")
    for tag in soup(["script", "style", "nav", "footer"]):
        tag.decompose()
    return soup.get_text(separator="\n")


def search_for_cv_url(name: str, university: str) -> Optional[str]:
    query = f'{name} {university} curriculum vitae CV'
    url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
    try:
        resp = request_with_retries("GET", url, headers={"User-Agent": USER_AGENT})
    except Exception as e:
        log_error(f"  [search error for {name}]", e)
        return None

    soup = BeautifulSoup(resp.text, "lxml")
    links = [a.get("href") for a in soup.select("a.result__a") if a.get("href")]
    if not links:
        return None
    for href in links[:8]:
        low = href.lower()
        if ".edu" in low and ("cv" in low or "~" in low or "people" in low or "faculty" in low):
            return href
    return links[0]


# --------------------------------------------------------------------------
# Education section detection (header-scoping logic kept from v3 -- this
# part was fine and isn't the thing that was slow/wrong) + NEW v4 raw-
# snippet extraction (replaces all the institution/year parsing)
# --------------------------------------------------------------------------

_EDUCATION_HEADER_RE = re.compile(
    r"^(education(?:\s*(?:and|&|/)\s*(?:experience|training|employment))?"
    r"|academic\s+background|academic\s+training)\s*:?\s*$",
    re.IGNORECASE,
)
_NEXT_SECTION_HEADER_KEYWORDS = [
    "experience", "employment", "appointment", "position", "publication",
    "research interest", "teaching", "service", "award", "honor", "honour",
    "grant", "skill", "affiliation", "achievement", "recognition", "press",
    "media coverage", "mentoring", "invited talk", "presentation",
    "professional activit", "students mentored", "advising",
]


def _is_next_section_header(line: str) -> bool:
    line = line.strip()
    if not line or len(line) > 50:
        return False
    low = line.lower().rstrip(":").strip()
    words = re.findall(r"[a-z]+", low)
    for kw in _NEXT_SECTION_HEADER_KEYWORDS:
        kw_words = kw.split()
        if len(kw_words) == 1:
            if any(w.startswith(kw_words[0]) for w in words):
                return True
        else:
            if kw in low:
                return True
    return False


def find_education_section(lines: list[str]) -> Optional[tuple[int, int]]:
    """Find the (start, end) line-index bounds of an "Education" section,
    if a recognizable header exists. Returns None if not found -- caller
    falls back to a degree-keyword window instead. Unchanged from v3."""
    for i, line in enumerate(lines):
        if _EDUCATION_HEADER_RE.match(line.strip()):
            start = i + 1
            end = min(start + 25, len(lines))
            for j in range(start, end):
                if _is_next_section_header(lines[j]):
                    end = j
                    break
            return start, end
    return None


_NOISE_LINE_RE = re.compile(r"^[\s•\-–—*◦▪·+]*$")

# How much raw text we're willing to dump into one spreadsheet cell.
# Generous enough to capture a full Education block, capped so one
# freakishly long section can't blow out the whole sheet.
MAX_SNIPPET_LINES = 30
MAX_SNIPPET_CHARS = 1500


def find_education_snippet(text: str) -> Optional[str]:
    """
    v4 replacement for the old institution/year parser. Just finds the
    Education section and returns its RAW TEXT, verbatim, for a human to
    read -- no attempt to identify "the" institution name or "the" start
    year out of it.

    Path 1: a recognizable "Education" header exists -> return the raw
    text of that section (capped at MAX_SNIPPET_LINES / MAX_SNIPPET_CHARS).

    Path 2 (fallback, no clean header found): grab a window of text
    (3 lines before, 10 lines after) around the FIRST line anywhere in the
    document that matches a Bachelor's-level degree keyword (B.S., B.A.,
    Bachelor, etc.) -- covers CVs that mention their degree in a sentence
    rather than a clean bulleted "Education" section.

    Returns None if neither path finds anything -- i.e. no education
    signal of any kind anywhere in the text.
    """
    lines = [
        ln.strip() for ln in text.splitlines()
        if ln.strip() and not _NOISE_LINE_RE.match(ln.strip())
    ]
    if not lines:
        return None

    edu_bounds = find_education_section(lines)
    if edu_bounds:
        start, end = edu_bounds
        end = min(end, start + MAX_SNIPPET_LINES)
        snippet = "\n".join(lines[start:end]).strip()
        if snippet:
            return snippet[:MAX_SNIPPET_CHARS]

    for i, line in enumerate(lines):
        if _DEGREE_KEYWORD_RE.search(line):
            window = lines[max(0, i - 3): i + 10]
            snippet = "\n".join(window).strip()
            if snippet:
                return snippet[:MAX_SNIPPET_CHARS]

    return None


def earliest_year_in(snippet: str) -> Optional[int]:
    """Rough signal only (see module docstring) -- just the smallest
    4-digit year anywhere in the given text, used solely so the
    olympiad-before-college comparison still has something to compare
    against. Not a claim about which year is 'the' start year."""
    years = [int(y) for y in YEAR_RE.findall(snippet)]
    return min(years) if years else None


def has_education_signal(text: str) -> bool:
    """
    Quick yes/no: does this text contain ANY education signal at all
    (an Education header, or a Bachelor's-level degree keyword anywhere,
    including the glued-ALL-CAPS format)?

    Used specifically for **website**-sourced fetches: per instruction, if
    a fetched page has no education signal whatsoever, we don't try to dig
    further -- we mark it skipped and move on, rather than burning effort
    trying to extract something that isn't there.
    """
    return find_education_snippet(text) is not None or bool(_DEGREE_GLUED_CAPS_SIGNAL_RE.search(text))


def find_olympiads(text: str) -> list[tuple[str, Optional[int], str]]:
    """Unchanged from v3 -- this part wasn't the problem."""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    hits = []
    for i, line in enumerate(lines):
        m = _ACRONYM_RE.search(line) or _FULLNAME_RE.search(line)
        if not m:
            continue

        same_line_years = [int(y) for y in YEAR_RE.findall(line)]
        if same_line_years:
            year = min(same_line_years)
        else:
            window = " ".join(lines[max(0, i - 1): i + 2])
            window_years = [int(y) for y in YEAR_RE.findall(window)]
            year = min(window_years) if window_years else None

        hits.append((m.group(1), year, line[:150]))

    return hits


# --------------------------------------------------------------------------
# Per-person orchestration
# --------------------------------------------------------------------------

@dataclass
class Result:
    name: str
    university: str
    source: str = ""            # "cv_url_column" / "local_file" / "web_search" / none
    source_detail: str = ""     # the URL, or local file path
    fetch_status: str = ""      # ok / no_source / ambiguous_local_match / no_education_signal / fetch_error: ...
    education_snippet: str = ""
    earliest_year: Optional[int] = None
    olympiad_mentions: list = field(default_factory=list)
    olympiad_before_college: bool = False
    olympiad_summary: str = ""
    college: Optional[str] = None

    @property
    def edu_cell_value(self) -> str:
        if self.fetch_status == "no_source":
            return "No CV found"
        if self.fetch_status == "ambiguous_local_match":
            return "Ambiguous local match — check manually"
        if self.fetch_status == "no_education_signal":
            return "No education signal on page — skipped"
        if self.fetch_status != "ok":
            return f"Error: {self.fetch_status}"
        return self.education_snippet or "Education section not found"

    @property
    def olympiad_cell_value(self) -> str:
        if self.fetch_status == "no_source":
            return "No CV found"
        if self.fetch_status == "ambiguous_local_match":
            return "Ambiguous local match — check manually"
        if self.fetch_status == "no_education_signal":
            return "Skipped (no education signal)"
        if self.fetch_status != "ok":
            return f"Error: {self.fetch_status}"
        return self.olympiad_summary if self.olympiad_summary else "None found"


# ============================================================================
# College (undergraduate institution) extraction -- ported in from a
# separately-maintained extract_college.py module, after extensive real-world
# bug-fixing against actual scraped CV/bio text (mixed formatting, ALL-CAPS
# pages, glued PDF-extraction artifacts, informal prose mentions, non-English
# institution names, etc. -- see inline comments throughout for the specific
# confirmed bugs each piece fixes). This runs ENTIRELY on already-fetched
# text -- no network calls of its own -- which is what makes it possible to
# re-run cheaply against rows that already have an Education Section (Raw)
# cached from a previous run (see process_cv_or_website_workbook() below).
# ============================================================================

# Institution name patterns. Extended beyond English-only "University of X"
# / "X University" to also catch the equivalent constructions in Spanish/
# Portuguese ("Universidad(e) de X"), French ("Université de X"), Italian
# ("Università di/degli Studi di X"), German ("Universität X") -- confirmed
# missed on a real CV (Joaquín Moraga's undergrad, "Universidad de
# Concepción", was silently dropped because the old pattern only recognized
# the English word "University").
#
# Digits are deliberately EXCLUDED from every character class here (using
# a-zA-Z instead of \w) -- real institution names never contain digits, and
# allowing them let date/glued-text artifacts get swept in (confirmed:
# "University of Michigan,August2017" was extracting as "University of
# Michigan, August" before this).
INSTITUTION_RE = re.compile(
    r"\b("
    r"University\s*of\s*[A-Z][a-zA-Z.'-]*(?:[\s,]+[A-Z][a-zA-Z.'-]*){0,2}"
    r"|Universidad(?:e)?\s*(?:de|do|del)?\s*[A-Z][a-zA-Záéíóúñãõ'-]*(?:\s+[A-Z][a-zA-Záéíóúñãõ'-]*){0,2}"
    r"|Universit[éeà]\s*(?:de|di|degli\s+Studi\s+di)?\s*[A-Z][a-zA-Zàèéìòù'-]*(?:\s+[A-Z][a-zA-Zàèéìòù'-]*){0,2}"
    r"|Universit[äa]t\s*[A-Z]?[a-zA-Zäöüß'-]*(?:\s+[A-Z][a-zA-Zäöüß'-]*){0,2}"
    r"|(?:[A-Z][a-zA-Z.&'-]*\s*){1,5}(?:University|College|Institute(?:\s*of\s*Technology)?|Polytechnic(?:\s*Institute)?)"
    r"(?:\s*(?:London|Dublin|Cork|Galway))?"
    # "UC <Campus>" informal short form -- e.g. "UC Berkeley", "UC San
    # Diego", "UC Santa Barbara" -- confirmed missed on a real CV
    # ("UC Berkeley Ph.D. in Mathematics..."). This is an extremely
    # common way University of California campuses are referred to,
    # distinct from the formal "University of California, Berkeley"
    # already handled by the first branch above.
    r"|UC\s+[A-Z][a-zA-Z]*(?:\s+[A-Z][a-zA-Z]*)?"
    r")\b"
    # Deliberately CASE-SENSITIVE (no re.IGNORECASE) -- the "[A-Z]" classes
    # throughout this pattern are load-bearing: they're what stops a
    # lowercase filler word ("at", "of", "in", "the") from being swept in
    # as if it were a proper-noun continuation of the institution's name.
    # An earlier attempt added re.IGNORECASE here to handle ALL-CAPS source
    # text (e.g. "NORTHWESTERN UNIVERSITY") -- confirmed as a real
    # regression: it also made "[A-Z]" match lowercase letters everywhere
    # else in the pattern, so "University of North Carolina at Chapel
    # Hill" started extending straight through "at" into "Chapel Hill" too
    # eagerly, and unrelated lowercase words got swept in generally. The
    # ALL-CAPS case is instead handled by pre-normalizing an all-uppercase
    # WINDOW to Title Case before this regex ever runs -- see
    # extract_institution_from_window() -- which fixes the real problem
    # (matching literal "University" against literal "UNIVERSITY") without
    # this collateral damage.
)
_KNOWN_ACRONYM_SCHOOLS = [
    "MIT", "UCLA", "UCSD", "UCSB", "UCI", "UCB", "UCD", "NYU", "USC", "CMU",
    "UPenn", "UVA", "UNC", "UMD", "UIUC", "WashU", "JHU", "RPI", "UW",
    "OSU", "BU", "GT", "UT", "ETH", "EPFL", "ENS",
    # Well-known institution names that are commonly written WITHOUT any
    # "University"/"College"/"Institute" suffix at all -- confirmed missed
    # on real CVs: "Virginia Tech, Blacksburg, VA" (no "Institute" in
    # sight -- its formal name is "Virginia Polytechnic Institute and
    # State University", but virtually nobody writes that out) and
    # "Dartmouth 2014" (its formal name IS "Dartmouth College", but often
    # just "Dartmouth" alone on a CV). "Caltech" is a single fused word
    # with no separate "Institute"/"Technology" token for the ordinary
    # suffix-branch to latch onto.
    "Virginia Tech", "Dartmouth", "Caltech",
]
_ACRONYM_SCHOOL_RE = re.compile(
    r"\b(" + "|".join(re.escape(a) for a in _KNOWN_ACRONYM_SCHOOLS) + r")\b"
)

# Used ONLY to try extending an already-found match further right, to catch
# compound names like "Chinese University of Hong Kong". See comment on
# _extend_with_trailing_of() below.
_TRAILING_OF_RE = re.compile(r"^\s*of\s+([A-Z][a-zA-Z.'-]*(?:\s+[A-Z][a-zA-Z.'-]*){0,2})")

# STRICT degree patterns -- deliberately EXCLUDES bare "Undergraduate" by
# itself (too generic -- triggers on "Undergraduate Students", "Undergraduate
# algorithms" course names, "Undergraduate Program" dept sections, none of
# which are the CV subject's own degree). Only genuine degree abbreviations
# survive here.
DEGREE_RE = re.compile(
    r"\bB\.?\s?S\.?(?![a-z])|\bB\.?\s?A\.?(?![a-z])|\bB\.?\s?Eng\.?(?![a-z])|"
    r"\bB\.?\s?Tech\.?(?![a-z])|\bB\.?\s?E\.?(?![a-z])|\bS\.?\s?B\.?(?![a-z])|"
    r"\bSc\.?\s?B\.?(?![a-z])|\bA\.?\s?B\.?(?![a-z])|\bB\.?\s?Sc\.?(?![a-z])|"
    # BASc (Bachelor of Applied Science, standard in Canadian engineering
    # programs) -- confirmed missed on a real CV (University of Toronto,
    # "2011 BASc in Engineering Science"). Distinct from plain "B.Sc" above
    # (has an extra "A"), so needs its own alternative rather than being
    # caught by an existing pattern. BMath similarly covers Waterloo-style
    # "Bachelor of Mathematics".
    r"\bB\.?A\.?Sc\.?(?![a-z])|\bB\.?Math\.?(?![a-z])|"
    r"\bBachelor(?:'s)?(?:\s+of\s+\w+)?|"
    # Handles "BAin Mathematics" / "BSin Physics" -- confirmed on a real
    # CV where "B.A. in Mathematics" got extracted with zero space
    # between the degree abbreviation and the word "in" ("BAin"). The
    # ordinary patterns above all reject this via their "(?![a-z])"
    # guard, since a lowercase "i" immediately follows -- correctly, for
    # avoiding a false match mid-word in general, but this specific
    # "immediately followed by lowercase 'in' then a capital or space"
    # shape is distinctive enough to safely carve out as an exception.
    r"\bB\.?A\.?(?=in[A-Z\s])|\bB\.?S\.?(?=in[A-Z\s])|\bB\.?E\.?(?=in[A-Z\s])|"
    # -------------------------------------------------------------------
    # INFORMAL, PROSE-STYLE mentions -- confirmed missed on real faculty
    # "About me"/"Biography" pages with NO formal degree abbreviation at
    # all, e.g. "Before that, I was an undergraduate at the University of
    # Waterloo", "As an undergraduate, Matt worked in the lab of...",
    # "completed his undergraduate studies at RWTH Aachen University".
    # (These patterns were already added to the separate scraper script's
    # own degree-keyword list months earlier -- this was a real gap: they
    # were never carried over into THIS module, which does the actual
    # college-name extraction, so has_degree_signal() never fired on any
    # of these pages even though the scraper's own has_education_signal()
    # correctly flagged the page as having SOME education content.)
    r"\bcollege\s+years\b|\bas\s+an\s+undergraduate\b|"
    r"\bundergraduate\s+(?:degree|years|studies|study)\b|\bundergraduate\s+at\b|"
    r"\bgraduate\s+of\b",
    re.IGNORECASE,
)

# Glued ALL-CAPS format, e.g. "BSCHEMISTRY", "BAECONOMICS", "BECHEMICALENG"
# -- confirmed on a real page where degree + subject were extracted with
# zero separator, entirely uppercase. Requires 6+ more uppercase letters
# immediately after (not 4 -- confirmed too low: "BECKMAN" has 5 letters
# after "BE" and was matching as a false positive at that threshold; real
# subject names glued on ("CHEMISTRY", "ENGINEERING", "PHYSICS") are
# reliably 6+ letters, while short common name endings after BA/BS/BE
# generally aren't).
#
# Deliberately COMPILED WITHOUT re.IGNORECASE (unlike DEGREE_RE above) --
# this is not an oversight. The whole point of "{A-Z}{6,}" is to verify
# the text is GENUINELY uppercase; under IGNORECASE, "[A-Z]" stops meaning
# "capital letter" and starts meaning "any letter at all", which would
# silently degrade the check into "followed by 6+ letters of any case" --
# true of almost every word in English.
DEGREE_GLUED_CAPS_RE = re.compile(
    r"\bB\.?S\.?(?=[A-Z]{6,})|\bB\.?A\.?(?=[A-Z]{6,})|\bB\.?E\.?(?=[A-Z]{6,})"
)


def has_degree_signal(line):
    return bool(DEGREE_RE.search(line) or DEGREE_GLUED_CAPS_RE.search(line))

# Attribution guard -- strip "Advisor(s): Name, Name, Name" spans before
# institution matching. Handles a COMMA-SEPARATED LIST of names, and
# tolerates missing whitespace after commas (a common PDF-extraction
# artifact, e.g. "Thomas Y.Hou,Houman Owhadi,Andrew M.Stuart").
ATTRIB_RE = re.compile(
    r"\b(?i:advisors?|advised\s+by|co-?advisors?|supervisors?|supervised\s+by)\s*:?\s*"
    r"(?:(?:Dr|Prof|Professor)\.?\s*)?"
    r"(?:[A-Z][\w.'-]*\s*)+"
    r"(?:,\s*(?:(?:Dr|Prof|Professor)\.?\s*)?(?:[A-Z][\w.'-]*\s*)+)*"
    r"(?:and\s+(?:(?:Dr|Prof|Professor)\.?\s*)?(?:[A-Z][\w.'-]*\s*)+)?"
)

# Degree-abbreviation tokens that must NEVER be swept into an institution
# name match (e.g. "A.B. Harvard College" was confirmed extracting with the
# "A.B." still attached).
_DEGREE_TOKEN_STRIP_RE = re.compile(
    r"\b(?:A\.?B\.?|B\.?S\.?|B\.?A\.?|B\.?Eng\.?|B\.?Tech\.?|B\.?E\.?|S\.?B\.?|"
    r"Sc\.?B\.?|B\.?Sc\.?|Ph\.?D\.?|D\.?Phil\.?|M\.?S\.?|M\.?A\.?|M\.?Eng\.?|"
    r"M\.?Phil\.?|M\.?Math\.?|M\.?Sc\.?)\b\.?,?"
)

# Month names glued next to a year with no separating space are a common
# CV-formatting artifact (e.g. "University of Florida May 2012") -- without
# stripping these, "May" (a valid-looking capitalized word) gets swept in
# as if it were part of a campus/city name.
_MONTH_STRIP_RE = re.compile(
    r"\b(?:January|February|March|April|May|June|July|August|September|"
    r"October|November|December)\b"
)

# Common CV section-header words AND academic subject/major names that
# must never be swept in as if they were part of an institution's name
# (e.g. "Employment Northeastern University" -- the true undergrad was
# several lines above; "Physics. Massachusetts Institute of Technology" --
# "Physics" is the subject/major line right before the school name, not
# part of the school's name).
_SECTION_HEADER_STRIP_RE = re.compile(
    r"\b(?:Employment|Experience|Positions?|Appointments?|Education|"
    r"Awards?|Honors?|Honours?|Grants?|Publications?|Service|Teaching|"
    r"Research|Fields?|Physics|Mathematics|Chemistry|Biology|Biochemistry|"
    r"Astronomy|Astrophysics|Engineering|Statistics)\b(?=\s*[:,.]?\s*[A-Z])"
)

# A STRONGER version of the same idea, used to exclude an entire line from
# window-building (not just strip one leading word from it). Needed
# specifically because window search order now tries the CLOSEST lines
# first regardless of direction (see extract_college() below) -- if a
# job-listing line like "Employment Northeastern University, Assistant
# Professor..." merely had "Employment" stripped but "Northeastern
# University" left behind, it would still get matched as if it were an
# institution, just without the word "Employment" attached. Confirmed as a
# real risk once forward-context is tried early: excluding the WHOLE line
# is what correctly makes the search skip past it to find the real
# institution elsewhere, rather than settling for a job employer.
_NON_EDUCATION_LINE_START_RE = re.compile(
    r"^\s*(?:Employment|Experience|Positions?|Appointments?|"
    r"Awards?|Honors?|Honours?|Grants?|Publications?|Service|Teaching|"
    r"Research\s+(?:Interests?|Positions?|Experience)|Professional\s+"
    r"Experience|Current\s+Position)\b",
    re.IGNORECASE,
)


def _filter_non_education_lines(lines):
    return [ln for ln in lines if not _NON_EDUCATION_LINE_START_RE.match(ln)]

# Reuses the script's existing _NOISE_LINE_RE (defined earlier in this file)
_TRAILING_JUNK_RE = re.compile(
    r"\s*(?:19[5-9]\d|20[0-4]\d|\d+|B\.?S\.?|B\.?A\.?|Ph\.?D\.?|M\.?S\.?|M\.?A\.?|M\.?Math\.?).*$"
)
_TRAILING_ABBREV_RE = re.compile(r",?\s*[A-Z]{2,4}$")
# Strips a trailing ". <PronounOrArticle>" -- confirmed on a real case
# ("University of Waterloo. My research is supported..." extracted as
# "University of Waterloo. My" because the period got absorbed as part of
# "Waterloo." by the word-character class, and then the continuation
# pattern picked up "My" as if it were one more word of the institution's
# name). This signals a NEW SENTENCE starting, not part of the name --
# real institution names don't end mid-match with ". <common pronoun>".
_TRAILING_SENTENCE_RE = re.compile(
    r"\.\s*(?:My|His|Her|He|She|I|We|They|The|This|That|It|Our|You|Your)\b.*$"
)


def clean_institution(name):
    name = name.replace("\x01", " ")
    # Normalize ALL-CAPS matches (e.g. "NORTHWESTERN UNIVERSITY", from a
    # real source page formatted entirely in uppercase) to Title Case for
    # a readable spreadsheet -- but skip anything that's already mixed-
    # case (leaves acronym-style names like "MIT" or genuinely mixed
    # names untouched, since .title()'ing something already correct risks
    # mangling internal capitals like "McGill" -> "Mcgill").
    if name.isupper():
        name = name.title()
        # str.title() incorrectly capitalizes the letter after an
        # apostrophe (e.g. "St. Xavier's" -> "St. Xavier'S") -- fix that
        # specific, common gotcha back down.
        name = re.sub(r"'([A-Z])\b", lambda mo: "'" + mo.group(1).lower(), name)
    name = re.sub(r"(?<=[a-z])of", " of", name)
    name = re.sub(r"of(?=[A-Z])", "of ", name)
    name = re.sub(r"(?<=Universidad)(de|do|del)(?=\s*[A-Z])", r" \1 ", name)
    name = re.sub(r"(?<=Universidade)(de|do|del)(?=\s*[A-Z])", r" \1 ", name)
    name = re.sub(r"(?<=[a-z])(University|College|Institute|Polytechnic)", r" \1", name)
    name = re.sub(r"\s+", " ", name).strip(" .,-")
    name = _TRAILING_JUNK_RE.sub("", name).strip(" .,-")
    name = name.rstrip("´`'")
    return name


def _extend_with_trailing_of(matched_text, window, match_end_pos):
    """
    Handles compound institution names like "Chinese University of Hong
    Kong" or "City University of New York" -- these contain the word
    "University" in the MIDDLE of their formal name, followed by "of
    <Place>". Without this step, the main regex's leftmost-match behavior
    picks the SHORTER "Chinese University" (via the generic "leading-words
    + University" branch, which starts earlier in the string and wins
    under Python's leftmost-match search) instead of the full name.

    Fix: after getting a match, check what comes right after it in the
    original window. If it starts with "of <Capitalized words>", that's
    almost certainly the rest of the same compound name -- append it.
    """
    remainder = window[match_end_pos:]
    m = _TRAILING_OF_RE.match(remainder)
    if m:
        return f"{matched_text} of {m.group(1)}"
    return matched_text


def _normalize_if_all_caps(window):
    """
    If the window is (almost) entirely uppercase -- e.g. a real source
    page formatted as "BA, CHEMISTRY..., NORTHWESTERN UNIVERSITY" -- the
    case-sensitive INSTITUTION_RE below would never match the literal
    title-case string "University"/"College" against literal "UNIVERSITY".
    Converting to Title Case here (checked BEFORE any of the other
    stripping steps, on a copy, so degree/section-header word stripping
    downstream still recognizes their all-caps forms too since those ARE
    compiled with re.IGNORECASE) fixes that without touching
    INSTITUTION_RE's case-sensitivity, which is load-bearing elsewhere
    (see the comment on INSTITUTION_RE itself for why).

    Uses a letters-only ratio check rather than str.isupper() directly,
    since isupper() is thrown off by the digits/punctuation ubiquitous in
    these windows (years, commas, periods) which aren't letters at all.
    """
    letters = [c for c in window if c.isalpha()]
    if len(letters) < 6:
        return window
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio > 0.9:
        titled = window.title()
        # str.title() incorrectly capitalizes the letter right after an
        # apostrophe or a period-abbreviation ("O'Brien" -> "O'Brien" is
        # fine, but "M.S." -> "M.S." can become "M.S." fine too -- the
        # real gotcha is names like "ST. XAVIER'S" -> "St. Xavier'S").
        titled = re.sub(r"'([A-Z])\b", lambda mo: "'" + mo.group(1).lower(), titled)
        # str.title() ALSO capitalizes every connector word, including
        # "of" -- confirmed as a real bug: "UNIVERSITY OF NORTH CAROLINA"
        # became "University Of North Carolina", which then silently
        # failed to match INSTITUTION_RE's "University\s*of\s*..." branch
        # (that pattern requires a literal LOWERCASE "of" -- deliberately
        # case-sensitive, see the comment on INSTITUTION_RE for why the
        # whole pattern can't just be made case-insensitive instead).
        # Lower-case these small connector words back down so the
        # downstream match still succeeds.
        titled = re.sub(r"\b(Of|The|At|And|De|Do|Del|Di|Da|Von|Van)\b",
                         lambda mo: mo.group(1).lower(), titled)
        return titled
    return window


def extract_institution_from_window(window):
    window = _normalize_if_all_caps(window)
    window = ATTRIB_RE.sub(" ", window)
    window = _DEGREE_TOKEN_STRIP_RE.sub(" ", window)
    window = _MONTH_STRIP_RE.sub(" ", window)
    window = _SECTION_HEADER_STRIP_RE.sub(" ", window)
    window = re.sub(r"\bThe\s+(?=University|College|Institute|Polytechnic)", "", window)

    m = INSTITUTION_RE.search(window)
    if m:
        extended = _extend_with_trailing_of(m.group(1), window, m.end(1))
        cleaned = clean_institution(extended)
        cleaned = _TRAILING_SENTENCE_RE.sub("", cleaned).strip(" .,-")
        cleaned = _TRAILING_ABBREV_RE.sub("", cleaned).strip(" .,-")
        if cleaned and len(cleaned) >= 3:
            return cleaned[:100]
    m2 = _ACRONYM_SCHOOL_RE.search(window)
    if m2:
        return m2.group(1)
    return None


def extract_college(raw_text):
    """
    Given the raw 'Education Section' text for one person, try to find
    their undergraduate (Bachelor's-level) institution specifically.
    Returns None if no confident match -- caller leaves the cell blank
    in that case rather than guessing.

    Window search order tries IMMEDIATE FORWARD context first, then
    immediate backward, then progressively wider windows in both
    directions, ending with a deep-backward pass. This ordering reflects
    two real, competing CV conventions, confirmed on real data:

      (a) "[Degree], [Year]" on one line, "[Institution], [City]" on the
          VERY NEXT line -- confirmed as the dominant pattern (e.g.
          Danielle Mai's CV lists three separate degrees, each with its
          own institution line immediately following it. A pure
          backward-first search was matching the WRONG institution here
          -- the one belonging to the degree listed just above, not the
          one actually paired with the matched Bachelor's line).
      (b) One institution stated ONCE, followed by several degree lines
          below it (Ph.D., then M.S., then B.A., no repeated institution
          line) -- where the correct institution is several lines ABOVE
          the matched degree line, not after it.

    Trying close-forward before close-backward correctly resolves (a).
    Falling through to progressively wider windows (including deep
    backward) still correctly resolves (b) once the close attempts find
    nothing -- AS LONG AS whatever immediately follows the degree line in
    case (b) isn't itself something that looks like a plausible (but
    wrong) institution. That's what _filter_non_education_lines() below
    is for: it removes job/employment-listing lines from consideration
    entirely (not just strips a leading word from them), so trying
    forward-first can't accidentally match an employer's name instead of
    the real school -- confirmed as a real risk otherwise (Pablo Boixeda
    Alvarez's CV has "University of Cambridge" stated once above two
    degree lines, followed immediately by an unrelated "Employment
    Northeastern University..." line; without filtering, forward-first
    search would have matched Northeastern instead of Cambridge).
    """
    if not raw_text:
        return None
    lines = [ln.strip() for ln in str(raw_text).splitlines() if ln.strip() and not _NOISE_LINE_RE.match(ln.strip())]
    if not lines:
        return None
    lines = _filter_non_education_lines(lines)
    if not lines:
        return None

    # Build a SEPARATE list of institution-search lines, each individually
    # normalized to Title Case if THAT LINE ALONE is overwhelmingly
    # uppercase -- kept separate from `lines` (used for degree-signal
    # detection) for the same reason explained above (glued-caps degree
    # detection needs original casing). Critically, this normalization
    # decision is made PER LINE, not on the combined multi-line window
    # text -- confirmed as a real bug otherwise: a CV formatted with an
    # all-caps institution line ("UNIVERSITY OF NORTH CAROLINA ASHEVILLE,
    # ASHEVILLE, NC") immediately followed by a normal-case degree line
    # ("Bachelor of Science, Chemistry and minor in Biology") were, once
    # joined into one window for matching, only ~50% uppercase overall --
    # well under the 90% ratio threshold that triggers normalization -- so
    # the all-caps institution text was left untouched and never matched
    # INSTITUTION_RE's case-sensitive "University" literal. Checking each
    # line independently means the all-caps institution line still gets
    # normalized correctly regardless of what case its neighboring lines
    # happen to be in.
    institution_lines = [_normalize_if_all_caps(ln) for ln in lines]

    for i, line in enumerate(lines):
        if not has_degree_signal(line):
            continue
        # Lines are joined with "\x01" (a control character, not matched
        # by \s or any character class used in INSTITUTION_RE) rather than
        # a plain space. This lets an institution match still extend
        # WITHIN one original line (e.g. "University of California, Los
        # Angeles" on a single line), while preventing it from bleeding
        # ACROSS a line boundary into unrelated comma-separated content on
        # an adjacent line.
        for window_lines in (
            [institution_lines[i]],                              # same line
            institution_lines[i: i + 2],                          # line + immediate next (close FORWARD)
            institution_lines[max(0, i - 1): i + 1],              # immediate prev + line (close BACKWARD)
            institution_lines[max(0, i - 1): i + 2],              # prev + line + next
            institution_lines[i: i + 3],                          # line + next 2 (wider forward)
            institution_lines[max(0, i - 2): i + 1],              # 2 back + line (wider backward)
            institution_lines[max(0, i - 2): i + 3],
            institution_lines[max(0, i - 4): i + 1],              # deep backward
            institution_lines[max(0, i - 4): i + 4],              # widest catch-all
        ):
            inst = extract_institution_from_window(" \x01 ".join(window_lines))
            if inst:
                return inst

    return None


def _looks_like_terminal_status_marker(edu_cell_text: str) -> bool:
    """True if a cached Education Section (Raw) cell holds one of the
    terminal STATUS markers this script itself writes (fetch error, no
    education signal, no CV found, etc.) rather than real scraped text.
    Used by the tier-2 resume path in process_cv_or_website_workbook():
    only cells with REAL cached text are worth re-running college
    extraction against for free -- a status marker has no education
    content to extract anything from."""
    t = edu_cell_text.strip()
    return (
        t.startswith("Error:")
        or t == "No CV found"
        or t == "No education signal on page — skipped"
        or t == "Ambiguous local match — check manually"
        or t == "Education section not found"
    )


def _is_retryable_terminal_marker(edu_cell_text: str) -> bool:
    """
    A SUBSET of terminal markers worth automatically retrying with a
    fresh re-fetch, rather than skipping forever. Specifically: "No
    education signal on page — skipped".

    Why this one specifically: that verdict comes from
    has_education_signal(), which is gated on UNDERGRAD_DEGREE_PATTERNS
    (plus the glued-ALL-CAPS check). Both were extended with new patterns
    (BASc, B.Math, informal prose phrasings, glued-caps degree+subject)
    AFTER many rows had already been fetched and marked this way under
    the older, narrower version. Confirmed as a real, unrecoverable gap:
    when a row is marked "no education signal", the actual fetched page
    text is discarded entirely -- NOTHING is saved to fall back on -- so
    a row stuck with this specific marker can only ever be fixed by a
    fresh re-fetch, never by re-analyzing a cache (there isn't one).

    Other terminal markers are NOT retried automatically: fetch errors,
    "No CV found", and "Ambiguous local match" all reflect network/source
    problems that a smarter degree-pattern list cannot fix, so retrying
    them here would just waste time re-hitting the same dead ends for no
    benefit.
    """
    return edu_cell_text.strip() == "No education signal on page — skipped"


def _run_extraction(text: str) -> tuple[str, Optional[int], list, str, Optional[str]]:
    """Shared by both processing paths below: given fetched/read text,
    return (education_snippet, earliest_year, olympiad_hits,
    olympiad_summary, college). `college` is computed by extract_college()
    against the FULL fetched text (not just the snippet) -- it does its
    own independent line-scoping/windowing internally, same as it always
    has when run standalone."""
    snippet = find_education_snippet(text) or ""
    year = earliest_year_in(snippet) if snippet else None

    olympiad_hits = find_olympiads(text)
    qualifying = []
    for kw, oy, snip in olympiad_hits:
        if year is not None and oy is not None and oy <= year:
            qualifying.append(f"{kw} ({oy})")
    olympiad_summary = "; ".join(qualifying)

    college = extract_college(text)

    return snippet, year, olympiad_hits, olympiad_summary, college


def _college_from_snippet_only(snippet: str) -> Optional[str]:
    """Cheaper alternative to _run_extraction() used for the resume/re-run
    path (see process_cv_or_website_workbook()): when a row ALREADY has a
    cached Education Section (Raw) snippet from a previous run, college
    extraction can run directly against that cached snippet -- no network
    re-fetch needed at all. This is what makes it fast to re-run this
    script against a TODO-only export (rows that have real education text
    but were missing a College value under an OLDER version of
    extract_college()): every fix since then gets applied retroactively
    just by re-running locally against text that's already on disk."""
    return extract_college(snippet)


def process_person(
    name: str,
    university: str,
    cv_url: str,
    local_index: Optional[list],
    do_search: bool,
    surname_counts: Optional[dict] = None,
    fallback_results: Optional[dict] = None,
) -> Result:
    r = Result(name=name, university=university)

    text = None
    was_ambiguous_locally = False

    if cv_url:
        r.source = "cv_url_column"
        r.source_detail = cv_url
        try:
            text = fetch_text_from_url(cv_url)
        except Exception as e:
            log_error(f"  [fetch error for {name} @ {cv_url}]", e)
            r.fetch_status = f"fetch_error: {str(e)[:100]}"
            return r

    if text is None and local_index is not None:
        local_path, match_status = find_local_cv(name, local_index, surname_counts)
        if match_status == "matched":
            r.source = "local_file"
            r.source_detail = str(local_path)
            try:
                text = read_local_file(local_path)
            except Exception as e:
                log_error(f"  [local read error for {name} @ {local_path}]", e)
                r.fetch_status = f"fetch_error: {str(e)[:100]}"
                return r
        elif match_status == "ambiguous":
            was_ambiguous_locally = True

    if text is None and fallback_results:
        key = normalize_text(name).strip()
        entry = fallback_results.get(key)
        if entry:
            rtype = entry["result_type"]
            value = entry["value"]
            source_url = entry["source_url"]
            if rtype == "cv_download" and value:
                r.source = "fallback_cv_download"
                r.source_detail = value
                try:
                    text = read_cv_download_value(value)
                except Exception as e:
                    log_error(f"  [fallback cv_download read error for {name} @ {value}]", e)
                    r.fetch_status = f"fetch_error: {str(e)[:100]}"
                    return r
                # If this "cv_download" row turned out to actually be a
                # URL (mislabeled), it just hit the network like a normal
                # website fetch -- apply the same skip-if-no-signal rule.
                if _looks_like_url(value) and text is not None and not has_education_signal(text):
                    r.fetch_status = "no_education_signal"
                    return r
            elif rtype == "website":
                url = value or source_url
                if url:
                    r.source = "fallback_website"
                    r.source_detail = url
                    try:
                        text = fetch_text_from_url(url)
                    except Exception as e:
                        log_error(f"  [fallback website fetch error for {name} @ {url}]", e)
                        r.fetch_status = f"fetch_error: {str(e)[:100]}"
                        return r
                    # Website source -> apply the skip-if-no-signal rule.
                    if text is not None and not has_education_signal(text):
                        r.fetch_status = "no_education_signal"
                        return r

    if text is None and do_search:
        try:
            found = search_for_cv_url(name, university)
        except Exception as e:
            log_error(f"  [unexpected search error for {name}]", e)
            found = None
        time.sleep(SLEEP_SECONDS)
        if found:
            r.source = "web_search"
            r.source_detail = found
            try:
                text = fetch_text_from_url(found)
            except Exception as e:
                log_error(f"  [fetch error for {name} @ {found}]", e)
                r.fetch_status = f"fetch_error: {str(e)[:100]}"
                return r
            if text is not None and not has_education_signal(text):
                r.fetch_status = "no_education_signal"
                return r

    if text is None and was_ambiguous_locally:
        r.fetch_status = "ambiguous_local_match"
        r.source = "local_file"
        return r

    if text is None:
        r.fetch_status = "no_source"
        return r

    r.fetch_status = "ok"
    snippet, year, olympiad_hits, olympiad_summary, college = _run_extraction(text)
    r.education_snippet = snippet
    r.earliest_year = year
    r.olympiad_mentions = olympiad_hits
    r.olympiad_summary = olympiad_summary
    r.olympiad_before_college = bool(olympiad_summary)
    r.college = college

    return r


# --------------------------------------------------------------------------
# Workbook I/O
# --------------------------------------------------------------------------

def col_letter_to_index(ws, header_name: str) -> Optional[int]:
    for cell in ws[1]:
        if cell.value == header_name:
            return cell.column
    return None


def process_cv_or_website_workbook(
    input_path: str,
    output_path: Optional[str] = None,
    limit: Optional[int] = None,
    verbose_log: Optional[str] = None,
    force: bool = False,
) -> None:
    """
    --direct-mode: processes EVERY sheet in `input_path` whose name starts
    with FALLBACK_SHEET_PREFIX ("CV or Website") in one run -- i.e. all
    five subject tabs (Math, CS, Physics, Engineer, Biochem) if your
    workbook has all five, no extra setup needed per tab.

    For each person:
      - result_type == "cv_download": read the local file at `value`,
        then ALWAYS attempt extraction (these are real downloaded CVs, so
        we trust there's something worth reading even if the Education
        header detection is fuzzy on that particular file).
      - result_type == "website": fetch the URL at `value` (or
        `source_url`). Then check has_education_signal() on the fetched
        text FIRST -- if there's no education signal on the page at all,
        mark the row "No education signal on page — skipped" and move on
        without further processing, per your instruction. Only if a
        signal IS present do we run the actual snippet/olympiad
        extraction.

    Writes 3 columns onto each sheet: NEW_COL_EDU_SNIPPET (raw text),
    NEW_COL_EARLIEST_YEAR (rough year signal), NEW_COL_OLYMPIAD.
    """
    wb = openpyxl.load_workbook(input_path)
    sheets = [s for s in wb.sheetnames if s.startswith(FALLBACK_SHEET_PREFIX)]
    if not sheets:
        print(f"ERROR: no sheets starting with '{FALLBACK_SHEET_PREFIX}' found in {input_path}. "
              f"Sheets: {wb.sheetnames}", file=sys.stderr)
        sys.exit(1)

    out_path = output_path or input_path
    in_place = (os.path.abspath(out_path) == os.path.abspath(input_path))

    # SAFETY: this mode writes results directly onto every "CV or Website
    # - <Subject>" sheet in your workbook and, by default, saves back over
    # the SAME file (in_place=True) -- that's the point (you asked for
    # results written straight into cv_or_website_results.xlsx, across all
    # 5 tabs, with no separate output file to merge back later). But since
    # this file also holds hand-verified columns from prior work (e.g.
    # "College", "Olympiad Medal" on the Math tab), an unattended run that
    # errors out partway through is exactly the kind of thing you want a
    # safety net for. So: before touching anything, copy the untouched
    # input file to a timestamped backup sitting right next to it. This
    # never gets read by the script again -- it's purely an undo point for
    # you. Skipped only if in_place is False (you passed a different
    # --output, so your original file is never written to anyway).
    if in_place:
        import shutil
        from datetime import datetime
        backup_path = f"{input_path}.backup-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
        try:
            shutil.copy2(input_path, backup_path)
            print(f"[safety] backed up untouched file to: {backup_path}")
        except Exception as e:
            log_error("[safety] backup failed -- proceeding anyway, but you have no undo point", e)

    print(f"[direct-mode] writing IN PLACE onto: {input_path}" if in_place
          else f"[direct-mode] writing to a SEPARATE file: {out_path} (input file untouched)")
    print(f"[direct-mode] sheets to process ({len(sheets)}): {sheets}")

    verbose_rows = []
    totals = {"ok": 0, "no_source": 0, "no_signal": 0, "error": 0}
    row_counter = 0

    sheets_completed = []
    sheets_crashed = []

    for sheet_num, sheet_name in enumerate(sheets, start=1):
        print(f"\n{'='*70}\n=== STARTING SHEET {sheet_num}/{len(sheets)}: {sheet_name} ===\n{'='*70}")

        # Everything for ONE sheet lives inside this try/except. This is
        # the fix for "the run stopped after one sheet and never touched
        # the others": previously, any exception that slipped past the
        # per-fetch try/except (e.g. a corrupt cell, an unexpected type,
        # anything in the extraction/write path itself) would propagate
        # all the way up and kill the ENTIRE run -- every sheet after the
        # one that crashed was simply never reached, silently. Now: a
        # sheet-level crash is caught, logged, the workbook is saved with
        # whatever progress THAT sheet made so far, and the loop moves on
        # to the next sheet regardless.
        try:
            ws = wb[sheet_name]
            headers = [c.value for c in ws[1]]
            try:
                name_col = headers.index(FALLBACK_COL_NAME) + 1
                type_col = headers.index(FALLBACK_COL_RESULT_TYPE) + 1
                value_col = headers.index(FALLBACK_COL_VALUE) + 1
            except ValueError:
                print(f"[skip] sheet '{sheet_name}' missing one of "
                      f"'{FALLBACK_COL_NAME}'/'{FALLBACK_COL_RESULT_TYPE}'/'{FALLBACK_COL_VALUE}' "
                      f"in row 1 (found: {headers}); skipping this sheet")
                continue
            source_col = headers.index(FALLBACK_COL_SOURCE_URL) + 1 if FALLBACK_COL_SOURCE_URL in headers else None

            edu_col = col_letter_to_index(ws, NEW_COL_EDU_SNIPPET)
            if edu_col is None:
                edu_col = ws.max_column + 1
                ws.cell(row=1, column=edu_col, value=NEW_COL_EDU_SNIPPET)
            year_col = col_letter_to_index(ws, NEW_COL_EARLIEST_YEAR)
            if year_col is None:
                year_col = ws.max_column + 1
                ws.cell(row=1, column=year_col, value=NEW_COL_EARLIEST_YEAR)
            olympiad_col = col_letter_to_index(ws, NEW_COL_OLYMPIAD)
            if olympiad_col is None:
                olympiad_col = ws.max_column + 1
                ws.cell(row=1, column=olympiad_col, value=NEW_COL_OLYMPIAD)
            college_col = col_letter_to_index(ws, NEW_COL_COLLEGE)
            if college_col is None:
                college_col = ws.max_column + 1
                ws.cell(row=1, column=college_col, value=NEW_COL_COLLEGE)

            total_rows = ws.max_row - 1
            n_to_process = min(limit, total_rows) if limit else total_rows
            sheet_ok = sheet_skipped_already_done = sheet_no_signal = sheet_no_source = sheet_error = 0
            sheet_college_from_cache = 0
            sheet_escalated_to_refetch = 0
            print(f"({n_to_process} rows to check)")

            for i in range(2, 2 + n_to_process):
                name = ws.cell(row=i, column=name_col).value
                if not name:
                    continue
                name = str(name).strip()

                existing_college = ws.cell(row=i, column=college_col).value
                existing_edu = ws.cell(row=i, column=edu_col).value

                # RESUME support, tier 1: College already filled -> this
                # row is fully done, nothing left to do (unless --force).
                if existing_college and not force:
                    sheet_skipped_already_done += 1
                    continue

                # RESUME support, tier 2: if Education Section (Raw) is
                # ALREADY cached from a previous run AND it looks like real
                # content (not one of the terminal "Error:"/"No CV
                # found"/"skipped" markers), try college extraction
                # directly against that cached text FIRST, with ZERO
                # network calls -- this is what makes it cheap to re-run
                # this script after extract_college() itself has been
                # improved (new degree patterns, new institution formats,
                # bug fixes, etc.).
                #
                # If that STILL finds nothing, don't just give up on the
                # row: the cached snippet may have been TRUNCATED by an
                # older/shorter capture (find_education_snippet() caps
                # snippet length), so a real Bachelor's mention could exist
                # on the page just outside what got cached -- there's no
                # way to tell from the cache alone. Confirmed as a real,
                # recurring situation: many rows kept coming back "still no
                # college found in cached snippet" on every re-run, with no
                # path to progress without ever trying a live re-fetch. So
                # a cache MISS here escalates into a real fresh re-fetch
                # below (falls through, does NOT skip) -- this still only
                # re-fetches the rows that actually need it, not the whole
                # sheet, unlike a blanket --force.
                cache_miss_escalated = False
                if (not force) and existing_edu and not _looks_like_terminal_status_marker(str(existing_edu)):
                    college = _college_from_snippet_only(str(existing_edu))
                    if college:
                        ws.cell(row=i, column=college_col, value=college)
                        sheet_college_from_cache += 1
                        print(f"[{sheet_name} row {i}] {name} -> college from CACHED snippet (no fetch): {college!r}")
                        continue
                    else:
                        print(f"[{sheet_name} row {i}] {name} -> no college in cached snippet; "
                              f"escalating to a fresh re-fetch (cache may be truncated/incomplete)")
                        sheet_escalated_to_refetch += 1
                        cache_miss_escalated = True

                if existing_edu and not force and not cache_miss_escalated:
                    if _is_retryable_terminal_marker(str(existing_edu)):
                        # Was marked "no education signal" under the OLDER,
                        # narrower detection -- worth a fresh re-fetch now
                        # that has_education_signal() recognizes more
                        # patterns (BASc, B.Math, glued-caps, informal
                        # prose). No cache to fall back on here (the
                        # original fetched text was discarded), so this
                        # can only be recovered by re-fetching for real.
                        print(f"[{sheet_name} row {i}] {name} -> retrying "
                              f"(previously 'no education signal' under an older, narrower detection)")
                        sheet_escalated_to_refetch += 1
                    else:
                        # Other terminal markers (fetch error / no CV
                        # found / ambiguous local match) reflect
                        # network/source problems a smarter degree-pattern
                        # list can't fix -- still skipped, same as always.
                        sheet_skipped_already_done += 1
                        continue

                # Falls through to here either because there was no
                # usable cache at all, OR because tier-2 just escalated a
                # cache-miss into a real re-fetch -- both cases proceed to
                # the normal fetch-from-source path below.

                # Per-row processing is ALSO wrapped, separately from the
                # fetch-specific try/except below -- so if something odd
                # happens in extraction/cell-writing for one specific row
                # (not just the network fetch), it can't take down the
                # whole sheet either. Worst case: that one row gets an
                # "Error: unexpected_error: ..." and processing continues.
                try:
                    rtype_raw = ws.cell(row=i, column=type_col).value
                    rtype = str(rtype_raw).strip().lower() if rtype_raw else ""

                    value_cell = ws.cell(row=i, column=value_col)
                    if rtype == "website":
                        value = _resolve_cell_value(value_cell)
                    else:
                        value = str(value_cell.value).strip() if value_cell.value else ""

                    source_url = ""
                    if source_col:
                        source_url = _resolve_cell_value(ws.cell(row=i, column=source_col))

                    source_detail = value or source_url
                    print(f"[{sheet_name} row {i}] {name} ({rtype or 'unknown'}) ...", flush=True)

                    text = None
                    fetch_status = "no_source"

                    try:
                        if rtype == "cv_download" and value:
                            text = read_cv_download_value(value)
                            if _looks_like_url(value):
                                # Mislabeled row: "cv_download" but value
                                # was actually a URL -- just fetched it
                                # like a website, so apply the same
                                # skip-if-no-signal rule.
                                fetch_status = "ok" if has_education_signal(text) else "no_education_signal"
                            else:
                                fetch_status = "ok"
                        elif rtype == "website":
                            url = value or source_url
                            if url:
                                text = fetch_text_from_url(url)
                                time.sleep(SLEEP_SECONDS)
                                if has_education_signal(text):
                                    fetch_status = "ok"
                                else:
                                    fetch_status = "no_education_signal"
                    except Exception as e:
                        log_error(f"  [fetch error for {name} @ {source_detail}]", e)
                        fetch_status = f"fetch_error: {str(e)[:100]}"

                    year = None
                    college = None
                    if fetch_status == "ok" and text is not None:
                        snippet, year, olympiad_hits, olympiad_summary, college = _run_extraction(text)
                        edu_cell = snippet or "Education section not found"
                        olympiad_cell = olympiad_summary if olympiad_summary else "None found"
                        totals["ok"] += 1
                        sheet_ok += 1
                    elif fetch_status == "no_education_signal":
                        edu_cell = "No education signal on page — skipped"
                        olympiad_cell = "Skipped (no education signal)"
                        totals["no_signal"] += 1
                        sheet_no_signal += 1
                    elif fetch_status.startswith("fetch_error"):
                        edu_cell = f"Error: {fetch_status}"
                        olympiad_cell = f"Error: {fetch_status}"
                        totals["error"] += 1
                        sheet_error += 1
                    else:
                        edu_cell = "No CV found"
                        olympiad_cell = "No CV found"
                        totals["no_source"] += 1
                        sheet_no_source += 1

                    ws.cell(row=i, column=edu_col, value=edu_cell)
                    ws.cell(row=i, column=year_col, value=year)
                    ws.cell(row=i, column=olympiad_col, value=olympiad_cell)
                    if college:
                        ws.cell(row=i, column=college_col, value=college)

                    print(f"    -> status={fetch_status} | edu_snippet={edu_cell[:80]!r} | "
                          f"college={college!r} | olympiad={olympiad_cell!r}")

                    if verbose_log:
                        verbose_rows.append({
                            "sheet": sheet_name, "name": name, "result_type": rtype,
                            "source": source_detail, "fetch_status": fetch_status,
                            "education_snippet": edu_cell,
                            "earliest_year": year, "olympiad_before_college": olympiad_cell,
                        })
                except Exception as e:
                    log_error(f"  [UNEXPECTED per-row error for {name} in {sheet_name}]", e)
                    ws.cell(row=i, column=edu_col, value=f"Error: unexpected_error: {str(e)[:100]}")
                    totals["error"] += 1
                    sheet_error += 1

                row_counter += 1
                if row_counter % 25 == 0:
                    try:
                        wb.save(out_path)
                        print(f"    [autosaved progress to {out_path}]")
                    except Exception as e:
                        log_error("    [autosave failed]", e)

            print(f"=== FINISHED SHEET {sheet_num}/{len(sheets)}: {sheet_name} -- "
                  f"ok={sheet_ok} no_signal={sheet_no_signal} no_source={sheet_no_source} "
                  f"error={sheet_error} college_from_cache={sheet_college_from_cache} "
                  f"escalated_to_refetch={sheet_escalated_to_refetch} "
                  f"already_done_skipped={sheet_skipped_already_done} ===")
            sheets_completed.append(sheet_name)

            # Save after EVERY sheet finishes, not just every 25 rows --
            # so even if the NEXT sheet crashes hard, this sheet's results
            # are guaranteed to already be on disk.
            try:
                wb.save(out_path)
                print(f"    [saved after finishing sheet '{sheet_name}']")
            except Exception as e:
                log_error(f"    [save after sheet '{sheet_name}' failed]", e)

        except Exception as e:
            # Sheet-level crash: log it clearly, save whatever progress
            # this sheet made before it died, and move on to the NEXT
            # sheet instead of aborting the whole run.
            log_error(f"[SHEET CRASHED] '{sheet_name}' hit an unrecoverable error and was "
                      f"aborted partway through -- moving on to the next sheet", e)
            sheets_crashed.append(sheet_name)
            try:
                wb.save(out_path)
                print(f"    [saved partial progress for '{sheet_name}' before moving on]")
            except Exception as e2:
                log_error("    [save after sheet crash failed]", e2)

    wb.save(out_path)
    print(f"\n{'='*70}\nDone across {len(sheets)} sheet(s).")
    print(f"  Completed cleanly: {sheets_completed}")
    if sheets_crashed:
        print(f"  CRASHED partway through (re-run the same command to pick up where each left off "
              f"-- already-done rows are skipped automatically): {sheets_crashed}")
    print(f"  Totals: ok={totals['ok']} no_source={totals['no_source']} "
          f"no_education_signal={totals['no_signal']} errors={totals['error']}")
    print(f"Saved to {out_path}")

    if verbose_log and verbose_rows:
        with open(verbose_log, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(verbose_rows[0].keys()))
            writer.writeheader()
            writer.writerows(verbose_rows)
        print(f"Wrote verbose log to {verbose_log}")

    print("Remember: every result still needs a human eyeball-check.")


def main():
    global DEBUG

    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--input", default=None,
                     help=f'Path to the .xlsx workbook (default: {DEFAULT_INPUT_XLSX!r})')
    ap.add_argument("--output", default=None,
                     help=f"Where to save results (default: {DEFAULT_OUTPUT_XLSX!r})")
    ap.add_argument("--sheet", default=None, help=f'Sheet name (default: "{DEFAULT_SHEET}")')
    ap.add_argument("--cv-folder", default=None,
                     help=f"Path to a local folder of CV files, searched recursively "
                          f"(default: {DEFAULT_CV_FOLDER!r})")
    ap.add_argument("--fallback-xlsx", default=None,
                     help=f"Path to the 'CV or Website' results workbook used when a person has "
                          f"no local CV (default: {DEFAULT_FALLBACK_XLSX!r}). Pass an empty string "
                          f'("") to disable this fallback entirely.')
    ap.add_argument("--direct-mode", action="store_true",
                     help="(This is now the DEFAULT behavior -- this flag is optional, kept only "
                          "for backward compatibility / being explicit.) Process a "
                          "cv_or_website_results.xlsx-style workbook DIRECTLY -- every "
                          "'CV or Website - <Subject>' sheet in it (all 5 tabs) is processed in "
                          f"one run. Defaults --input to {DEFAULT_DIRECT_MODE_XLSX!r}.")
    ap.add_argument("--single-sheet-mode", action="store_true",
                     help="Opt IN to the OLDER, single-file/single-sheet mode instead (one "
                          f'"research info" sheet in a subject-specific workbook, default input '
                          f"{DEFAULT_INPUT_XLSX!r}). This used to be the default when no flags "
                          "were given at all -- that caused real confusion (running the script "
                          "with no flags silently processed the wrong file/sheet instead of "
                          "cv_or_website_results.xlsx). Direct-mode is now the default instead; "
                          "you only need this flag if you specifically want the old behavior back.")
    ap.add_argument("--search", action="store_true",
                     help="Fall back to a web search for anyone still unresolved after the local "
                          "folder and fallback-xlsx steps")
    ap.add_argument("--limit", type=int, default=None, help="Only process the first N rows per sheet")
    ap.add_argument("--verbose-log", default=None,
                     help="Optional path to write a companion CSV with full match/evidence detail")
    ap.add_argument("--force", action="store_true",
                     help="(direct-mode only) Re-process rows that already have a result in the "
                          "Education Section (Raw) column. Without this flag, re-running the same "
                          "command SKIPS already-completed rows -- this is what makes it safe to just "
                          "re-run after an interruption/crash without losing progress or re-fetching "
                          "everything from scratch.")
    ap.add_argument("--debug", action="store_true", help="Print full tracebacks on error")
    args = ap.parse_args()
    DEBUG = args.debug

    # DEFAULT (no flags needed at all): direct-mode -- process every
    # "CV or Website - <Subject>" sheet in cv_or_website_results.xlsx, all
    # 5 tabs, in one run. --single-sheet-mode is the ONLY way to get the
    # old single-file/single-sheet behavior back; forgetting a flag can no
    # longer silently run the wrong thing the way it did before.
    if not args.single_sheet_mode:
        input_path = args.input if args.input is not None else DEFAULT_DIRECT_MODE_XLSX
        output_path = args.output if args.output is not None else input_path
        mode_note = "explicit --direct-mode" if args.direct_mode else "default -- no flag needed"
        print(f"[direct-mode ({mode_note})] input:  {input_path}")
        print(f"[direct-mode ({mode_note})] output: {output_path}")
        print()
        process_cv_or_website_workbook(
            input_path, output_path=output_path, limit=args.limit,
            verbose_log=args.verbose_log, force=args.force,
        )
        return

    input_path = args.input if args.input is not None else DEFAULT_INPUT_XLSX
    output_path = args.output if args.output is not None else DEFAULT_OUTPUT_XLSX
    sheet_name = args.sheet if args.sheet is not None else DEFAULT_SHEET
    cv_folder = args.cv_folder if args.cv_folder is not None else DEFAULT_CV_FOLDER
    fallback_xlsx_path = args.fallback_xlsx if args.fallback_xlsx is not None else DEFAULT_FALLBACK_XLSX

    print(f"input:          {input_path}")
    print(f"output:         {output_path}")
    print(f"sheet:          {sheet_name}")
    print(f"cv_folder:      {cv_folder}")
    print(f"fallback_xlsx:  {fallback_xlsx_path or '(disabled)'}")
    print()

    local_index = None
    if cv_folder:
        local_index = build_local_cv_index(cv_folder)
        print(f"Indexed {len(local_index)} local CV file(s) in {cv_folder}")

    fallback_results = load_fallback_results(fallback_xlsx_path) if fallback_xlsx_path else {}

    wb = openpyxl.load_workbook(input_path)
    if sheet_name not in wb.sheetnames:
        print(f"ERROR: sheet '{sheet_name}' not found. Sheets: {wb.sheetnames}", file=sys.stderr)
        sys.exit(1)
    ws = wb[sheet_name]

    name_col = col_letter_to_index(ws, COL_NAME)
    school_col = col_letter_to_index(ws, COL_SCHOOL)
    cv_url_col = col_letter_to_index(ws, COL_CV_URL)

    if name_col is None or school_col is None:
        print(f"ERROR: expected columns '{COL_NAME}' and '{COL_SCHOOL}' in row 1. "
              f"Found: {[c.value for c in ws[1]]}", file=sys.stderr)
        sys.exit(1)

    edu_col = col_letter_to_index(ws, NEW_COL_EDU_SNIPPET)
    if edu_col is None:
        edu_col = ws.max_column + 1
        ws.cell(row=1, column=edu_col, value=NEW_COL_EDU_SNIPPET)

    year_col = col_letter_to_index(ws, NEW_COL_EARLIEST_YEAR)
    if year_col is None:
        year_col = ws.max_column + 1
        ws.cell(row=1, column=year_col, value=NEW_COL_EARLIEST_YEAR)

    olympiad_col = col_letter_to_index(ws, NEW_COL_OLYMPIAD)
    if olympiad_col is None:
        olympiad_col = ws.max_column + 1
        ws.cell(row=1, column=olympiad_col, value=NEW_COL_OLYMPIAD)

    total_rows = ws.max_row - 1
    n_to_process = min(args.limit, total_rows) if args.limit else total_rows

    surname_to_names = {}
    for i in range(2, ws.max_row + 1):
        raw_name = ws.cell(row=i, column=name_col).value
        if not raw_name:
            continue
        toks = name_tokens(str(raw_name).strip())
        if not toks:
            continue
        sn = toks[-1]
        surname_to_names.setdefault(sn, set()).add(str(raw_name).strip())
    surname_counts = {sn: len(names) for sn, names in surname_to_names.items()}

    verbose_rows = []
    counts = {"ok": 0, "local_file": 0, "cv_url_column": 0, "web_search": 0,
              "fallback_cv_download": 0, "fallback_website": 0,
              "no_source": 0, "ambiguous": 0, "no_education_signal": 0, "error": 0}

    for i in range(2, 2 + n_to_process):
        name = ws.cell(row=i, column=name_col).value
        university = ws.cell(row=i, column=school_col).value
        if not name or not university:
            continue
        name = str(name).strip()
        university = str(university).strip()

        cv_url = ""
        if cv_url_col:
            v = ws.cell(row=i, column=cv_url_col).value
            cv_url = str(v).strip() if v else ""

        print(f"[row {i}] {name} ({university}) ...", flush=True)

        try:
            r = process_person(name, university, cv_url, local_index, do_search=args.search,
                                surname_counts=surname_counts, fallback_results=fallback_results)
        except Exception as e:
            log_error(f"  [UNEXPECTED error processing {name}]", e)
            r = Result(name=name, university=university, fetch_status=f"unexpected_error: {e}")

        ws.cell(row=i, column=edu_col, value=r.edu_cell_value)
        ws.cell(row=i, column=year_col, value=r.earliest_year if r.fetch_status == "ok" else None)
        ws.cell(row=i, column=olympiad_col, value=r.olympiad_cell_value)

        if r.fetch_status == "ok":
            counts["ok"] += 1
            counts[r.source] = counts.get(r.source, 0) + 1
        elif r.fetch_status == "no_source":
            counts["no_source"] += 1
        elif r.fetch_status == "ambiguous_local_match":
            counts["ambiguous"] += 1
        elif r.fetch_status == "no_education_signal":
            counts["no_education_signal"] += 1
        else:
            counts["error"] += 1

        print(f"    -> status={r.fetch_status} source={r.source or '-'} "
              f"({r.source_detail[:60] if r.source_detail else ''}) | "
              f"edu_snippet={r.edu_cell_value[:80]!r} | olympiad={r.olympiad_cell_value!r}")

        if args.verbose_log:
            raw = " | ".join(
                f"{kw}({year if year is not None else '?'}): {snip}"
                for kw, year, snip in r.olympiad_mentions
            )
            verbose_rows.append({
                "name": r.name, "university": r.university,
                "source": r.source, "source_detail": r.source_detail,
                "fetch_status": r.fetch_status,
                "education_snippet": r.education_snippet,
                "earliest_year": r.earliest_year or "",
                "olympiad_before_college": r.olympiad_before_college,
                "olympiad_summary": r.olympiad_summary,
                "all_olympiad_mentions_raw": raw,
            })

        if (i - 1) % 25 == 0:
            try:
                wb.save(output_path)
                print(f"    [autosaved progress to {output_path}]")
            except Exception as e:
                log_error("    [autosave failed]", e)

        if r.source in ("cv_url_column", "web_search"):
            time.sleep(SLEEP_SECONDS)

    wb.save(output_path)
    print(f"\nDone. Processed {n_to_process} rows.")
    print(f"  ok: {counts['ok']}  "
          f"(local_file={counts.get('local_file', 0)}, "
          f"cv_url_column={counts.get('cv_url_column', 0)}, "
          f"fallback_cv_download={counts.get('fallback_cv_download', 0)}, "
          f"fallback_website={counts.get('fallback_website', 0)}, "
          f"web_search={counts.get('web_search', 0)})")
    print(f"  no_source (no local/URL/fallback/search match): {counts['no_source']}")
    print(f"  no_education_signal (website fetched, but nothing education-shaped on it): {counts['no_education_signal']}")
    print(f"  ambiguous local matches (needs manual check): {counts['ambiguous']}")
    print(f"  errors: {counts['error']}")
    print(f"Updated '{sheet_name}' with columns '{NEW_COL_EDU_SNIPPET}', '{NEW_COL_EARLIEST_YEAR}', "
          f"and '{NEW_COL_OLYMPIAD}'. Saved to {output_path}")

    if args.verbose_log:
        with open(args.verbose_log, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(verbose_rows[0].keys()) if verbose_rows else [])
            writer.writeheader()
            writer.writerows(verbose_rows)
        print(f"Also wrote full match/evidence log to {args.verbose_log}")

    print("Remember: every result needs a human eyeball-check before you trust it — "
          "especially anything flagged 'Ambiguous local match'.")


if __name__ == "__main__":
    main()