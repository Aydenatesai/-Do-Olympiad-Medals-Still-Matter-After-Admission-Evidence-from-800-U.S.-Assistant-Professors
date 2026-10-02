#!/usr/bin/env python3
"""
check_position_orcid.py

A BETTER-FIT alternative to check_position_tavily.py for the specific
need of finding someone's current JOB TITLE (assistant/associate/full
professor, postdoc, etc.) -- something neither OpenAlex nor Semantic
Scholar can provide (they simply have no title/rank field at all).

WHY ORCID INSTEAD OF A GENERAL SEARCH API:
ORCID (orcid.org) is a free, structured researcher-identity registry.
Each ORCID record can have an "Employment" section -- self-reported by
the researcher themselves -- with fields for organization name, role
title, department, and start/end dates. When a researcher keeps this
updated, it's a DIRECT, STRUCTURED answer to "what is their current
position", rather than something inferred/synthesized from free-text web
search results (which is what Tavily-based lookup has to do instead).

TRADE-OFFS TO KNOW ABOUT (being upfront, not oversold):
  - Coverage gap: not every academic has an ORCID iD, and even those who
    do don't always keep the Employment section updated -- self-reported
    data quality varies. A "not found" or empty result here does NOT
    mean the person isn't in academia, only that ORCID doesn't have (or
    doesn't have current) data for them.
  - Disambiguation is WEAKER than the OpenAlex-based script for common
    names: ORCID's expanded-search returns institution-name per
    candidate (used here as a tie-break, same idea as before), but has
    no subject/topic field to check against, and no publication-timing
    signal either (ORCID doesn't expose per-year publication counts the
    way OpenAlex does). For very common names with several candidates,
    this is genuinely less reliable -- cross-check manually when it
    matters.
  - Requires a one-time FREE setup step (see below) -- unlike OpenAlex,
    ORCID's public API needs an OAuth client_id/client_secret pair (not
    just a single API key) to get an access token.

ONE-TIME SETUP (free, ~2 minutes):
  1. Create a free ORCID account at https://orcid.org/register (if you
     don't already have one -- this is a personal researcher account,
     distinct from the "client" credentials below).
  2. Go to https://orcid.org/developer-tools and register a new
     application to get a `client_id` and `client_secret` pair. This is
     what authenticates your SCRIPT (not you personally) to the public
     API -- you only need to do this once, and the credentials don't
     expire.

Usage:
    pip install requests openpyxl --break-system-packages
    python check_position_orcid.py --input olympiad_selected_us_or_matching_school.xlsx \\
        --output position_results_orcid.xlsx \\
        --client-id YOUR_CLIENT_ID --client-secret YOUR_CLIENT_SECRET --limit 20
"""
from __future__ import annotations

import argparse
import sys
import time
import re
import unicodedata
from dataclasses import dataclass
from typing import Optional

import requests
import openpyxl

ORCID_TOKEN_URL = "https://orcid.org/oauth/token"
ORCID_SEARCH_BASE = "https://pub.orcid.org/v3.0/expanded-search/"
ORCID_EMPLOYMENTS_BASE = "https://pub.orcid.org/v3.0"  # + /{orcid_id}/employments
REQUEST_TIMEOUT = 20
MAX_RETRIES = 4
RETRY_BACKOFF_BASE = 5
SLEEP_BETWEEN_REQUESTS = 0.3


def normalize(s) -> str:
    if not s:
        return ""
    if isinstance(s, list):
        # Defensive: some ORCID fields (e.g. institution-name) return a
        # list rather than a plain string -- confirmed as a real crash in
        # practice. Join rather than raise, here too, in case another
        # field surprises us the same way somewhere this isn't yet
        # explicitly handled at the call site.
        s = " ; ".join(str(x) for x in s if x)
        if not s:
            return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[.\-'\u2019]", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def fuzzy_name_score(a: str, b: str) -> float:
    ta, tb = set(normalize(a).split()), set(normalize(b).split())
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))


def request_with_retries(method: str, url: str, max_retries: int = MAX_RETRIES, **kwargs) -> Optional[requests.Response]:
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            if resp.status_code == 429 or resp.status_code >= 500:
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after else RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                last_exc = requests.HTTPError(f"HTTP {resp.status_code}")
                if attempt < max_retries:
                    print(f"    [retry {attempt}/{max_retries}, waiting {wait:.0f}s: HTTP {resp.status_code}]", file=sys.stderr)
                    time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp
        except Exception as e:
            last_exc = e
            if attempt < max_retries:
                wait = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                print(f"    [retry {attempt}/{max_retries}, waiting {wait:.0f}s: {e}]", file=sys.stderr)
                time.sleep(wait)
    print(f"    [request failed after {max_retries} attempts: {last_exc}]", file=sys.stderr)
    return None


def get_access_token(client_id: str, client_secret: str) -> Optional[str]:
    """One long-lived token per run -- ORCID's client_credentials tokens
    are typically valid for ~20 years, so there's no need to refresh
    mid-run. Fetched once in main() and passed to every subsequent call."""
    resp = request_with_retries(
        "POST", ORCID_TOKEN_URL, max_retries=3,
        headers={"Accept": "application/json"},
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "/read-public",
            "grant_type": "client_credentials",
        },
    )
    if not resp:
        return None
    return resp.json().get("access_token")


def search_orcid(name: str, token: str) -> list[dict]:
    """Uses expanded-search, which returns richer per-candidate info
    (institution-name included) in one call, rather than the basic
    /search endpoint which only returns bare ORCID iDs and would need a
    separate lookup per candidate just to get a name/institution to
    disambiguate with."""
    resp = request_with_retries(
        "GET", ORCID_SEARCH_BASE, max_retries=2,
        headers={"Accept": "application/vnd.orcid+json", "Authorization": f"Bearer {token}"},
        params={"q": name},
    )
    if not resp:
        return []
    try:
        data = resp.json()
    except Exception:
        return []
    return data.get("expanded-result") or []


def get_current_employment(orcid_id: str, token: str) -> tuple[str, str, str]:
    """Returns (organization, role_title, department) for whichever
    employment record has NO end-date (still ongoing) -- if several
    qualify, the one with the latest start-date wins. Returns empty
    strings if the person has no employment records on file at all
    (common -- many researchers never fill this section in)."""
    resp = request_with_retries(
        "GET", f"{ORCID_EMPLOYMENTS_BASE}/{orcid_id}/employments", max_retries=2,
        headers={"Accept": "application/vnd.orcid+json", "Authorization": f"Bearer {token}"},
    )
    if not resp:
        return "", "", ""
    try:
        data = resp.json()
    except Exception:
        return "", "", ""

    candidates = []
    for group in data.get("employment-summary", []) or data.get("affiliation-group", []):
        # ORCID's JSON nests summaries slightly differently across
        # sub-versions -- handle both a flat list and the grouped form.
        summaries = group.get("summaries", [group]) if isinstance(group, dict) and "summaries" in group else [group]
        for s in summaries:
            entry = s.get("employment-summary", s) if isinstance(s, dict) else s
            if not isinstance(entry, dict):
                continue
            end_date = entry.get("end-date")
            org = ((entry.get("organization") or {}).get("name")) or ""
            role = entry.get("role-title") or ""
            dept = entry.get("department-name") or ""
            start = entry.get("start-date") or {}
            start_year = (start.get("year") or {}).get("value") if start else None
            candidates.append((end_date is None, start_year or 0, org, role, dept))

    if not candidates:
        return "", "", ""
    # Prefer no end-date (still current) first, then latest start year.
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    _, _, org, role, dept = candidates[0]
    return org, role, dept


def get_undergrad_education(orcid_id: str, token: str) -> tuple[str, str, str]:
    """Returns (organization, role_title, start_year) for the EARLIEST
    entry in the person's ORCID "Education" affiliations (a distinct
    section from Employment -- covers "participation in an academic
    higher education program to receive an undergraduate, graduate, or
    other degree"). The earliest-starting education entry is used as the
    undergrad guess, same reasoning as the publication-window heuristic
    in check_academia_status.py: undergrad comes chronologically first,
    before any Master's/PhD entries. `role_title` often (not always)
    contains the degree type the person entered (e.g. "BSc", "Bachelor
    of Science") when they bothered to fill it in -- included as-is,
    unvalidated, since ORCID doesn't constrain this field's format.
    Returns empty strings if no Education entries exist on file at all."""
    resp = request_with_retries(
        "GET", f"{ORCID_EMPLOYMENTS_BASE}/{orcid_id}/educations", max_retries=2,
        headers={"Accept": "application/vnd.orcid+json", "Authorization": f"Bearer {token}"},
    )
    if not resp:
        return "", "", ""
    try:
        data = resp.json()
    except Exception:
        return "", "", ""

    candidates = []
    for group in data.get("education-summary", []) or data.get("affiliation-group", []):
        summaries = group.get("summaries", [group]) if isinstance(group, dict) and "summaries" in group else [group]
        for s in summaries:
            entry = s.get("education-summary", s) if isinstance(s, dict) else s
            if not isinstance(entry, dict):
                continue
            org = ((entry.get("organization") or {}).get("name")) or ""
            role = entry.get("role-title") or ""
            start = entry.get("start-date") or {}
            start_year = (start.get("year") or {}).get("value") if start else None
            if org:
                candidates.append((int(start_year) if start_year else 9999, org, role, start_year or ""))

    if not candidates:
        return "", "", ""
    candidates.sort(key=lambda c: c[0])  # earliest first -- likely undergrad
    _, org, role, start_year = candidates[0]
    return org, role, start_year


@dataclass
class OrcidResult:
    name: str
    hint_institution: str
    orcid_id: str = ""
    matched_name: str = ""
    current_organization: str = ""
    current_role: str = ""
    current_department: str = ""
    orcid_undergrad_org: str = ""
    orcid_undergrad_role: str = ""
    orcid_undergrad_start_year: str = ""
    found: str = "No"
    note: str = ""


def check_one_person_orcid(name: str, hint_institution: str, token: str) -> OrcidResult:
    r = OrcidResult(name=name, hint_institution=hint_institution)

    candidates = search_orcid(name, token)
    if not candidates:
        r.note = "No ORCID profile found matching this name."
        return r

    scored = []
    for c in candidates:
        given = c.get("given-names", "")
        family = c.get("family-names", "")
        full = f"{given} {family}".strip()
        name_score = fuzzy_name_score(name, full)
        if name_score < 0.4:
            continue
        inst_bonus = 0.0
        # ORCID's expanded-search returns institution-name as a LIST
        # (a researcher can have several affiliations on record), not a
        # single string -- confirmed as a real crash in practice
        # (TypeError: normalize() argument 2 must be str, not list).
        # Join them into one string so the substring-match below can
        # check against ANY of the listed institutions.
        raw_inst = c.get("institution-name")
        if isinstance(raw_inst, list):
            inst_text = " ; ".join(str(x) for x in raw_inst if x)
        else:
            inst_text = raw_inst or ""
        cand_inst = normalize(inst_text)
        if hint_institution and cand_inst and normalize(hint_institution) in cand_inst:
            inst_bonus = 1.0
        scored.append((name_score * 3 + inst_bonus, c, full))

    if not scored:
        r.note = "ORCID candidates found, but none matched the name closely enough."
        return r

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best, best_full = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else -1.0

    r.orcid_id = best.get("orcid-id", "")
    r.matched_name = best_full

    if not r.orcid_id:
        r.note = "Matched a candidate but no ORCID iD was returned -- unexpected, skipping."
        return r

    org, role, dept = get_current_employment(r.orcid_id, token)
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not org and not role:
        r.note = "ORCID profile found, but no Employment record on file (common -- many researchers never fill this in)."
    else:
        r.found = "Yes"
        r.current_organization = org
        r.current_role = role
        r.current_department = dept

    ug_org, ug_role, ug_year = get_undergrad_education(r.orcid_id, token)
    time.sleep(SLEEP_BETWEEN_REQUESTS)
    if ug_org:
        r.orcid_undergrad_org = ug_org
        r.orcid_undergrad_role = ug_role
        r.orcid_undergrad_start_year = ug_year

    if len(scored) > 1 and (best_score - second_score) < 1.0:
        r.note = f"{r.note} [Low-confidence: multiple similarly-named ORCID profiles found, verify manually.]".strip()

    return r


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", default="position_results_orcid.xlsx")
    ap.add_argument("--client-id", required=True, help="ORCID API client_id (see setup instructions in this file's docstring)")
    ap.add_argument("--client-secret", required=True, help="ORCID API client_secret")
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--start-row", type=int, default=2)
    args = ap.parse_args()

    print("Fetching ORCID access token...")
    token = get_access_token(args.client_id, args.client_secret)
    if not token:
        print("FATAL: could not obtain an ORCID access token -- check your client_id/client_secret.", file=sys.stderr)
        sys.exit(1)
    print("Token obtained.\n")

    wb_in = openpyxl.load_workbook(args.input, data_only=True)
    ws_in = wb_in.active
    headers = [c.value for c in ws_in[1]]
    name_col = headers.index("name")
    inst_col = headers.index("institution") if "institution" in headers else None
    guess_col = headers.index("guessed_undergrad") if "guessed_undergrad" in headers else None

    total_rows = ws_in.max_row - 1
    n_to_process = total_rows if args.limit is None or args.limit < 0 else min(args.limit, total_rows)

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "ORCID Position Check"
    ws_out.append([
        "name", "hint_institution", "found", "orcid_id", "matched_name",
        "current_organization", "current_role", "current_department",
        "orcid_undergrad_org", "orcid_undergrad_role", "orcid_undergrad_start_year", "note",
    ])

    processed = 0
    for i in range(args.start_row, args.start_row + n_to_process):
        row = ws_in[i]
        name = row[name_col].value
        if not name:
            continue
        hint = ""
        if inst_col is not None and row[inst_col].value:
            hint = str(row[inst_col].value)
        elif guess_col is not None and row[guess_col].value:
            hint = str(row[guess_col].value)

        print(f"[{i}] {name} (hint: {hint or '-'}) ...", flush=True)
        r = check_one_person_orcid(str(name), hint, token)
        print(f"    -> found={r.found} | role={r.current_role!r} | org={r.current_organization!r}")

        ws_out.append([
            r.name, r.hint_institution, r.found, r.orcid_id, r.matched_name,
            r.current_organization, r.current_role, r.current_department,
            r.orcid_undergrad_org, r.orcid_undergrad_role, r.orcid_undergrad_start_year, r.note,
        ])

        processed += 1
        if processed % 25 == 0:
            wb_out.save(args.output)
            print(f"    [autosaved after {processed} rows]")

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    wb_out.save(args.output)
    print(f"\nDone. Processed {processed} rows. Saved to {args.output}")
    print("Remember: an empty `current_role` doesn't mean 'not in academia' -- it often just means")
    print("that person never filled in ORCID's Employment section. Cross-check against")
    print("check_academia_status.py's OpenAlex-based results for better coverage.")


if __name__ == "__main__":
    main()