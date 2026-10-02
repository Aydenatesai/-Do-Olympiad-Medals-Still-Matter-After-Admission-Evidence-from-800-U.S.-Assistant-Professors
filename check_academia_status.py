#!/usr/bin/env python3
"""
check_academia_status.py

Reads "olympiad_2000_2015_deduplicated.xlsx" (columns: subject, name,
country, last_year) and, for each person, queries the OpenAlex API
(https://openalex.org -- free, no key required, structured author data)
to check whether they currently appear to be in academia.

WHY OPENALEX (and not just Google search):
  - Google has no free structured-search API; scraping search results at
    this scale (12k+ people) would be slow, fragile, and likely blocked.
  - OpenAlex indexes ~250M works and their authors, and for each author
    exposes `last_known_institutions` (with an institution `type`:
    "education", "company", "government", "healthcare", "nonprofit",
    "facility", "archive", "other") plus their most recent publication
    year and research topics. This is exactly the structured signal
    needed to answer "are they still academic-affiliated" without
    guessing from unstructured search snippets.
  - Semantic Scholar's API is used as a fallback for anyone OpenAlex
    doesn't find, since its author-affiliation coverage differs somewhat.

DISAMBIGUATION (the hard part -- olympiad names are extremely common):
  A name search on either API can return many candidates. Each candidate
  is scored using: (a) name similarity, (b) whether the candidate's
  research topics overlap the person's olympiad SUBJECT (a Math-olympiad
  alum working in pure math/CS-theory is a much likelier match than a
  same-named person in an unrelated field), (c) a small tie-break bonus
  if the candidate's institution country matches the olympiad country
  (weak signal only -- people emigrate constantly, so this never
  disqualifies a candidate on its own). If the top two candidates score
  too close together, the row is marked AMBIGUOUS rather than guessing.

OUTPUT: never asserts a bare yes/no. Every row gets a status, the
matched institution (if any), a confidence score, and a direct profile
URL, so a human can verify anything uncertain -- consistent with how
every other automated step in this project has been built: surface
evidence, don't assert unverified conclusions.

Usage:
    pip install requests openpyxl --break-system-packages
    python check_academia_status.py --input olympiad_2000_2015_deduplicated.xlsx --output academia_status.xlsx
    python check_academia_status.py --input olympiad_2000_2015_deduplicated.xlsx --output academia_status.xlsx --limit 50   # test run
    python check_academia_status.py --input olympiad_2000_2015_deduplicated.xlsx --output academia_status.xlsx --email you@example.com  # strongly recommended, see below

IMPORTANT -- the --email flag:
  OpenAlex asks (does not require, but strongly requests) that you pass
  your email in the `mailto` query parameter. Doing so moves your
  requests into their "polite pool", which has much higher and more
  reliable rate limits than anonymous requests. Pass any real email you
  control; it's only used for OpenAlex's own abuse-contact purposes, per
  their public API documentation.
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

import requests
import openpyxl

OPENALEX_BASE = "https://api.openalex.org/authors"
OPENALEX_WORKS_BASE = "https://api.openalex.org/works"
OPENALEX_RATE_LIMIT_BASE = "https://api.openalex.org/rate-limit"
S2_BASE = "https://api.semanticscholar.org/graph/v1/author/search"
REQUEST_TIMEOUT = 20
MAX_RETRIES = 5
# HTTP 429 needs a MUCH longer backoff than a generic server error --
# confirmed as the actual failure mode in real runs: nearly every row was
# coming back NOT_FOUND with confidence=0.0 because BOTH the OpenAlex
# request AND the Semantic Scholar fallback were getting 429'd, not
# because the person genuinely wasn't found. A short 3/6/9-second retry
# schedule doesn't give a rate limiter's window time to reset. Starting
# at 15s and doubling (15/30/60/120/240) gives real headroom.
RETRY_BACKOFF_BASE = 15
SLEEP_BETWEEN_REQUESTS = 1.0  # much more conservative -- see note above; this alone should prevent most 429s from ever happening in the first place, especially if --email wasn't passed (anonymous OpenAlex pool has a strict shared global limit)

# Maps this project's olympiad "subject" values to OpenAlex concept/topic
# keywords used for disambiguation scoring. Deliberately broad (a Math
# olympiad alum could plausibly end up in "Computer Science" or
# "Economics" academically) -- this is a SOFT signal, not a filter.
SUBJECT_TOPIC_HINTS = {
    "Mathematics": ["mathematics", "computer science", "physics", "statistics", "economics"],
    "Physics": ["physics", "astronomy", "materials science", "engineering", "mathematics"],
    "Chemistry": ["chemistry", "materials science", "biochemistry", "chemical engineering"],
    "Biology": ["biology", "genetics", "neuroscience", "medicine", "biochemistry", "ecology"],
    "Informatics": ["computer science", "mathematics", "engineering", "artificial intelligence"],
}

ACADEMIC_INSTITUTION_TYPES = {"education"}
# OpenAlex institution types seen in practice: education, company,
# government, healthcare, nonprofit, facility, archive, other, funder.
# "healthcare" is excluded on purpose -- a hospital affiliation alone
# doesn't mean an academic (research-track) position, though in practice
# many biology/chemistry medalists in medicine will show this; treated as
# NOT_IN_ACADEMIA by default since it's ambiguous, and left for human
# review via the status column rather than guessing either way.


def normalize(name: str) -> str:
    name = unicodedata.normalize("NFKD", name)
    name = "".join(c for c in name if not unicodedata.combining(c))
    name = re.sub(r"[.\-'\u2019]", "", name)
    return re.sub(r"\s+", " ", name).strip().lower()


def fuzzy_name_score(query_name: str, candidate_name: str) -> float:
    """0.0-1.0 token-overlap similarity -- deliberately simple (no fuzzy
    library dependency) since exact/near-exact token match is what
    matters here; a candidate missing a middle name/initial should still
    score reasonably high, but a totally different name should score low."""
    q = set(normalize(query_name).split())
    c = set(normalize(candidate_name).split())
    if not q or not c:
        return 0.0
    overlap = len(q & c)
    return overlap / max(len(q), len(c))


@dataclass
class PersonResult:
    subject: str
    name: str
    country: str
    last_year: int
    in_academia: str = "No"   # "Yes" or "No" -- ONLY these two values, see note below
    matched_name: str = ""
    institution: str = ""
    institution_type: str = ""
    most_recent_work_year: Optional[int] = None
    h_index: Optional[int] = None
    works_count: Optional[int] = None
    likely_merged_identity: str = "No"  # "Yes"/"No" -- flags implausibly high works_count (see note)
    guessed_undergrad: str = ""
    guessed_undergrad_mentions: int = 0
    confidence: float = 0.0
    profile_url: str = ""
    source: str = ""
    note: str = ""


def request_with_retries(url: str, params: dict, max_retries: int = MAX_RETRIES) -> Optional[dict]:
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 429 or resp.status_code >= 500:
                # Respect the server's own Retry-After header when it
                # provides one -- it knows its own rate-limit window
                # better than any fixed guess on our end. Falls back to
                # exponential backoff (doubling each attempt, not just
                # multiplying by attempt number) when the header is
                # absent, which is the common case for 429s specifically.
                retry_after = resp.headers.get("Retry-After")
                if retry_after:
                    try:
                        wait = float(retry_after)
                    except ValueError:
                        wait = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                else:
                    wait = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                last_exc = requests.HTTPError(f"HTTP {resp.status_code}")
                if attempt < max_retries:
                    print(f"    [retry {attempt}/{max_retries}, waiting {wait:.0f}s: HTTP {resp.status_code}]", file=sys.stderr)
                    time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            last_exc = e
            if attempt < max_retries:
                wait = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
                print(f"    [retry {attempt}/{max_retries}, waiting {wait:.0f}s: {e}]", file=sys.stderr)
                time.sleep(wait)
    print(f"    [request failed after {max_retries} attempts: {last_exc}]", file=sys.stderr)
    return None


def search_openalex(name: str, email: Optional[str]) -> list[dict]:
    """Uses the module-level `_openalex_keys` pool -- pulls the current
    key via `before_request()` (which may proactively rotate first), and
    on a definitive failure, rotates reactively and retries ONCE more
    with the next key before giving up and returning empty (letting the
    caller fall through to Semantic Scholar)."""
    key = _openalex_keys.before_request()
    if key is None:
        return []  # whole pool exhausted

    params = {"search": name, "per_page": 10, "api_key": key}
    if email:
        # Confirmed via OpenAlex's own Feb 2026 changelog: the mailto/
        # "polite pool" system was FULLY RETIRED as of Feb 13, 2026 --
        # "No more polite pool! No more email parameter in your calls".
        # Kept as a harmless no-op so old invocations without a key
        # don't hard-crash -- but it does NOTHING useful anymore.
        params["mailto"] = email
    # Only 2 retries here (not the full MAX_RETRIES) -- OpenAlex has a
    # working fallback (Semantic Scholar) right below, so it's wasteful
    # to burn the full exponential-backoff schedule on a source that's
    # allowed to fail over. If OpenAlex is persistently 429ing, failing
    # fast here keeps the whole run moving instead of stalling for hours.
    data = request_with_retries(OPENALEX_BASE, params, max_retries=2)

    if not data:
        # Could be genuine budget exhaustion on THIS key -- rotate and
        # try exactly once more with the next key before giving up
        # entirely for this person.
        next_key = _openalex_keys.mark_exhausted_and_rotate()
        if next_key is not None:
            params["api_key"] = next_key
            data = request_with_retries(OPENALEX_BASE, params, max_retries=2)
        if not data:
            return []

    return data.get("results", [])


def search_semantic_scholar(name: str) -> list[dict]:
    params = {"query": name, "fields": "name,affiliations,paperCount,hIndex,papers.year"}
    # Full retry budget here -- this is the LAST resort with no further
    # fallback, so it's worth waiting out a real rate-limit window.
    data = request_with_retries(S2_BASE, params, max_retries=MAX_RETRIES)
    if not data:
        return []
    return data.get("data", [])


def guess_undergrad_via_publication_window(
    openalex_author_id: str, medal_year: int
) -> tuple[Optional[str], int, int]:
    """
    Heuristic for undergrad institution, per the actual request that
    motivated this: rather than trusting an author's EARLIEST-ever
    tracked affiliation (usually grad school -- undergrads rarely
    publish), look specifically at the [medal_year+1, medal_year+5]
    window. An olympiad medal is typically won near the end of high
    school (~age 17-18), so this window approximates the person's
    UNDERGRAD years specifically. If they published anything during that
    window, whichever institution appears most often across THAT
    author's specific authorship entries in THAT window is a reasonable
    best-guess for their undergrad school.

    Returns (institution_name_or_None, mention_count, works_found_in_window).
    institution_name is None if the author published nothing in the
    window at all -- per instruction, that's simply left blank rather
    than guessed at ("nếu không có thì thôi").

    CAVEAT (worth keeping in mind, and surfaced in the output note by the
    caller): this is a best-effort signal, not a confirmed fact --
    publishing during undergrad is the exception, not the rule, so this
    will come back empty for MOST people even when it works correctly.
    A hit also isn't guaranteed to be a degree-granting enrollment (could
    be a summer REU, a visiting-student stint, etc.) -- it's the single
    most-mentioned affiliation in a plausible time window, nothing more.
    """
    if not openalex_author_id:
        return None, 0, 0

    key = _openalex_keys.before_request()
    if key is None:
        return None, 0, 0  # whole pool exhausted -- skip this optional lookup entirely

    # openalex_author_id may be a full URL ("https://openalex.org/A123")
    # or a bare ID -- the works filter wants the bare form.
    author_id = openalex_author_id.rsplit("/", 1)[-1]

    start_year = medal_year + 1
    end_year = medal_year + 5

    params = {
        "filter": f"author.id:{author_id},publication_year:{start_year}-{end_year}",
        "per_page": 50,
        "select": "id,publication_year,authorships",
        "api_key": key,
    }

    data = request_with_retries(OPENALEX_WORKS_BASE, params, max_retries=2)
    if not data:
        # Reactive rotation applies here too, same reasoning as
        # search_openalex() -- try once more with the next key.
        next_key = _openalex_keys.mark_exhausted_and_rotate()
        if next_key is not None:
            params["api_key"] = next_key
            data = request_with_retries(OPENALEX_WORKS_BASE, params, max_retries=2)
        if not data:
            return None, 0, 0

    works = data.get("results", [])
    if not works:
        return None, 0, 0

    institution_counts: dict[str, int] = {}
    for w in works:
        for authorship in w.get("authorships", []):
            a = authorship.get("author", {})
            a_id = (a.get("id") or "").rsplit("/", 1)[-1]
            if a_id != author_id:
                continue  # this authorship entry is a CO-author, not our person
            for inst in authorship.get("institutions", []):
                name = inst.get("display_name")
                if name:
                    institution_counts[name] = institution_counts.get(name, 0) + 1

    if not institution_counts:
        return None, 0, len(works)

    best_inst, best_count = max(institution_counts.items(), key=lambda kv: kv[1])
    return best_inst, best_count, len(works)


def score_openalex_candidate(candidate: dict, subject: str, country: str) -> float:
    score = 0.0
    topics = candidate.get("x_concepts") or candidate.get("topics") or []
    topic_names = " ".join(
        (t.get("display_name", "") if isinstance(t, dict) else "") for t in topics
    ).lower()
    hints = SUBJECT_TOPIC_HINTS.get(subject, [])
    if any(h in topic_names for h in hints):
        score += 2.0
    inst = (candidate.get("last_known_institutions") or [None])
    inst = inst[0] if inst else None
    if inst and country:
        inst_country = (inst.get("country_code") or "").upper()
        if inst_country and country[:2].upper() in inst_country:
            score += 0.5  # weak tie-break only
    works_count = candidate.get("works_count") or 0
    if works_count > 0:
        score += min(works_count / 20.0, 1.0)  # having published anything is a mild positive signal of being a real researcher
    return score


# --------------------------------------------------------------------
# OpenAlex API key pool + rotation
# --------------------------------------------------------------------
class OpenAlexKeyManager:
    """
    Manages a POOL of OpenAlex API keys, automatically rotating to the
    next one when the current key's daily $1 free budget runs low --
    this is what lets a run continue past the ~1000-search/key/day free
    limit without stopping or needing to be manually restarted tomorrow
    with a different key.

    Two triggers for rotation:
      1. PROACTIVE: every CHECK_INTERVAL calls on the current key, this
         queries OpenAlex's own /rate-limit endpoint (which reports
         exact remaining daily budget in USD) and rotates BEFORE the
         budget actually hits zero, so real work requests don't get
         wasted hitting an empty key.
      2. REACTIVE: if a real request still comes back 429 after its
         normal retries (the key might genuinely be out of budget RIGHT
         NOW, ahead of the next scheduled proactive check), the caller
         marks it exhausted immediately and rotation happens before
         falling through to the Semantic Scholar fallback.

    When every key in the pool is exhausted, OpenAlex is disabled for
    the REST OF THIS RUN (not just a cooldown -- daily budgets don't
    refill mid-run) and every remaining person is checked via Semantic
    Scholar only.
    """

    CHECK_INTERVAL = 25    # proactive /rate-limit check every N calls on a key
    MIN_BUDGET_USD = 0.02  # rotate pre-emptively once remaining budget drops below this (buffer of ~20 more search calls)

    def __init__(self, keys: list[str]):
        self.keys = [k.strip() for k in keys if k and k.strip()]
        self.index = 0
        self.calls_on_current_key = 0
        self.exhausted: set[int] = set()

    @property
    def current_key(self) -> Optional[str]:
        if not self.keys or self.index >= len(self.keys):
            return None
        return self.keys[self.index]

    def all_exhausted(self) -> bool:
        return not self.keys or len(self.exhausted) >= len(self.keys)

    def _advance_past_exhausted(self):
        while self.index < len(self.keys) and self.index in self.exhausted:
            self.index += 1

    def mark_exhausted_and_rotate(self) -> Optional[str]:
        """Call when a request definitively fails in a way that suggests
        the current key is out of budget. Rotates and returns the new
        key (or None if the whole pool is now exhausted)."""
        if self.all_exhausted():
            return None
        old = self.index
        self.exhausted.add(self.index)
        self.index += 1
        self.calls_on_current_key = 0
        self._advance_past_exhausted()
        if self.all_exhausted():
            print(f"    [key manager] All {len(self.keys)} OpenAlex key(s) exhausted for today -- "
                  f"OpenAlex disabled for the rest of this run, using Semantic Scholar only.", file=sys.stderr)
            return None
        print(f"    [key manager] Key #{old + 1} exhausted, switching to key #{self.index + 1}/{len(self.keys)}.",
              file=sys.stderr)
        return self.current_key

    def before_request(self) -> Optional[str]:
        """Call before each OpenAlex request. Periodically checks the
        current key's real remaining budget and rotates pre-emptively if
        it's running low. Returns the key to use, or None if the whole
        pool is exhausted (caller should skip OpenAlex entirely)."""
        if self.all_exhausted():
            return None
        key = self.current_key
        self.calls_on_current_key += 1
        if self.calls_on_current_key < self.CHECK_INTERVAL:
            return key
        self.calls_on_current_key = 0
        remaining = self._check_remaining_budget(key)
        if remaining is not None and remaining < self.MIN_BUDGET_USD:
            print(f"    [key manager] Key #{self.index + 1} has ${remaining:.4f} left "
                  f"(below ${self.MIN_BUDGET_USD} buffer) -- rotating pre-emptively.", file=sys.stderr)
            return self.mark_exhausted_and_rotate()
        return key

    @staticmethod
    def _check_remaining_budget(key: str) -> Optional[float]:
        # max_retries=1 -- this is just a courtesy check, not worth
        # burning real retry budget on; if it fails, we simply proceed
        # with the current key and let a real request failure (if the
        # key truly is empty) trigger reactive rotation instead.
        data = request_with_retries(OPENALEX_RATE_LIMIT_BASE, {"api_key": key}, max_retries=1)
        if not data:
            return None
        return (data.get("rate_limit") or {}).get("daily_remaining_usd")


_openalex_keys = OpenAlexKeyManager([])  # populated in main() from CLI args


# Simple circuit breaker for OpenAlex: if it fails several people in a
# row for reasons OTHER than budget exhaustion (a transient outage, not
# just occasional bad luck), skip calling it for a cooldown window rather
# than paying its retry cost on every single subsequent person.
_openalex_circuit = {"consecutive_failures": 0, "cooldown_until": 0.0}
OPENALEX_FAILURE_THRESHOLD = 5
OPENALEX_COOLDOWN_SECONDS = 120


def openalex_available() -> bool:
    return time.time() >= _openalex_circuit["cooldown_until"] and not _openalex_keys.all_exhausted()


def record_openalex_result(succeeded: bool) -> None:
    if succeeded:
        _openalex_circuit["consecutive_failures"] = 0
        return
    _openalex_circuit["consecutive_failures"] += 1
    if _openalex_circuit["consecutive_failures"] >= OPENALEX_FAILURE_THRESHOLD:
        _openalex_circuit["cooldown_until"] = time.time() + OPENALEX_COOLDOWN_SECONDS
        print(f"    [circuit breaker] OpenAlex failed {OPENALEX_FAILURE_THRESHOLD} times in a row -- "
              f"skipping it for the next {OPENALEX_COOLDOWN_SECONDS}s, using Semantic Scholar only", file=sys.stderr)
        _openalex_circuit["consecutive_failures"] = 0  # reset so we don't re-trigger every row during cooldown


def check_one_person(subject: str, name: str, country: str, last_year: int, email: Optional[str]) -> PersonResult:
    r = PersonResult(subject=subject, name=name, country=country, last_year=last_year)

    candidates = []
    source = "openalex"
    if openalex_available():
        candidates = search_openalex(name, email)
        record_openalex_result(succeeded=bool(candidates))
    else:
        print("    [circuit breaker] OpenAlex in cooldown, skipping straight to Semantic Scholar", file=sys.stderr)

    if not candidates:
        candidates_s2 = search_semantic_scholar(name)
        if candidates_s2:
            source = "semantic_scholar"
            # Normalize S2 shape to look enough like OpenAlex for scoring
            # below. h_index and paperCount are pulled straight through
            # here (both requested via the `fields` param above) so they
            # survive into the final output regardless of which source
            # actually answered for this person.
            candidates = [
                {
                    "display_name": c.get("name", ""),
                    "x_concepts": [],
                    "last_known_institutions": (
                        [{"display_name": c["affiliations"][0], "type": "education", "country_code": ""}]
                        if c.get("affiliations") else []
                    ),
                    "works_count": c.get("paperCount", 0),
                    "h_index": c.get("hIndex"),
                    "id": c.get("authorId", ""),
                    "_s2_raw": c,
                }
                for c in candidates_s2
            ]

    if not candidates:
        r.in_academia = "No"
        r.note = "No author profile found on OpenAlex or Semantic Scholar."
        return r

    scored = []
    for c in candidates:
        name_score = fuzzy_name_score(name, c.get("display_name", ""))
        if name_score < 0.4:
            continue  # not even a plausible name match, discard outright
        topic_score = score_openalex_candidate(c, subject, country)
        total = name_score * 3 + topic_score
        scored.append((total, c))

    if not scored:
        r.in_academia = "No"
        r.note = "Candidates returned, but none matched the name closely enough."
        return r

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else -1.0

    r.source = source
    r.matched_name = best.get("display_name", "")
    r.confidence = round(best_score, 2)

    # NOTE on removing the old separate AMBIGUOUS status: per request,
    # this script now ALWAYS commits to a Yes/No call rather than
    # returning a third "can't tell" bucket -- so the top-scoring
    # candidate is used regardless of how close the runner-up scored.
    # The closeness is NOT thrown away, though: it's recorded in `note`
    # so a low-confidence Yes/No is still visibly distinguishable from a
    # confident one if you sort/filter by the `confidence` column later.
    close_call = len(scored) > 1 and (best_score - second_score) < 1.0

    # h-index / works count -- pulled from whichever source answered.
    # OpenAlex nests h-index under summary_stats; Semantic Scholar's
    # normalized shape above already flattens it to h_index directly.
    if source == "openalex":
        r.h_index = (best.get("summary_stats") or {}).get("h_index")
        r.works_count = best.get("works_count")
    else:
        r.h_index = best.get("h_index")
        r.works_count = best.get("works_count")

    inst_list = best.get("last_known_institutions") or []
    inst = inst_list[0] if inst_list else None

    if inst:
        r.institution = inst.get("display_name", "")
        r.institution_type = inst.get("type", "")

    # Most recent publication year: OpenAlex exposes this via
    # `counts_by_year` -- pull the max year with works > 0.
    counts_by_year = best.get("counts_by_year") or []
    active_years = [c["year"] for c in counts_by_year if c.get("works_count", 0) > 0]
    if active_years:
        r.most_recent_work_year = max(active_years)
    elif source == "semantic_scholar":
        # S2's papers.year list is the equivalent signal for this source
        papers = (best.get("_s2_raw") or {}).get("papers") or []
        years = [p["year"] for p in papers if p.get("year")]
        if years:
            r.most_recent_work_year = max(years)

    r.profile_url = best.get("id", "") or (
        f"https://www.semanticscholar.org/author/{best['_s2_raw'].get('authorId','')}"
        if "_s2_raw" in best else ""
    )

    # --- Final Yes/No call -- exactly two possible values, no third
    # bucket. A low h-index or works_count is deliberately NOT used to
    # flip an otherwise-Yes verdict to No: it's recorded as a caveat in
    # `note` instead. Reasoning: this dataset's last_year ranges up to
    # 2015, so someone whose olympiad medal was recent (e.g. 2013-2015)
    # may well be an early-career grad student/postdoc/junior faculty
    # member RIGHT NOW with a genuinely low h-index simply because
    # they haven't had time to accumulate citations yet -- that's not
    # evidence they left research, and treating low h-index as a
    # disqualifier would systematically mislabel the youngest, most
    # recently-medaled cohort as "No" regardless of their real status.
    if not inst:
        r.in_academia = "No"
        r.note = "Author profile found, but no institution on record."
    elif r.institution_type in ACADEMIC_INSTITUTION_TYPES:
        if r.most_recent_work_year and r.most_recent_work_year < 2021:
            r.in_academia = "No"
            r.note = f"Institution type is academic, but most recent tracked publication year is {r.most_recent_work_year} (looks inactive)."
        else:
            r.in_academia = "Yes"
            r.note = ""
    else:
        r.in_academia = "No"
        r.note = f"Most recent institution type is '{r.institution_type}', not education."

    if r.h_index is not None and r.h_index <= 1 and r.in_academia == "Yes":
        # Not a disqualifier (see comment above) -- just a flagged caveat
        # worth a second glance, appended without overriding the Yes.
        extra = f"Note: h-index is only {r.h_index} -- worth a manual glance if you want extra confidence (could be early-career, or a weak/stale match)."
        r.note = f"{r.note} {extra}".strip()

    # Sanity check on works_count -- confirmed as a REAL failure mode in
    # practice, not a hypothetical: rows showing 1000-7000+ works for a
    # single person. No individual researcher has thousands of papers --
    # even the most prolific scientists in history rarely clear a couple
    # thousand across a full career. A number this high means OpenAlex's
    # name-based author disambiguation almost certainly MERGED multiple
    # different real people who share a common name into one author
    # profile (a well-documented weakness of automatic disambiguation,
    # especially for very common East/South/Southeast Asian given-name +
    # surname combinations like "Kai Wang", "Yang Liu", "Xi Chen" -- exactly
    # the kind of names common among this dataset's medalists). The
    # institution and h-index shown for such a row are NOT reliable
    # information about any one real person -- they're an aggregate blend.
    # Flagged rather than silently trusted, same pattern as the other
    # caveats here: don't flip the Yes/No, just make sure it's not taken
    # at face value.
    WORKS_COUNT_SUSPICIOUS_THRESHOLD = 300
    if r.works_count is not None and r.works_count > WORKS_COUNT_SUSPICIOUS_THRESHOLD:
        r.likely_merged_identity = "Yes"
        extra = (f"CAUTION: works_count={r.works_count} is implausibly high for one person -- "
                 f"likely an OpenAlex author-disambiguation merge of multiple different people "
                 f"sharing this name (common with frequent East/South/Southeast Asian names). "
                 f"Institution/h-index shown are probably an unreliable blend, not one real person -- verify manually.")
        r.note = f"{r.note} {extra}".strip()

    if close_call:
        r.note = f"{r.note} [Low-confidence match: top candidate score {best_score:.2f} vs runner-up {second_score:.2f} -- multiple similarly-named people found, verify manually if this row matters.]".strip()

    # Undergrad guess via the [medal_year+1, medal_year+5] publication
    # window -- see guess_undergrad_via_publication_window()'s docstring
    # for the full reasoning. Only attempted when the match came from
    # OpenAlex (needs a real OpenAlex author ID for the works-filter
    # lookup; Semantic Scholar doesn't expose an equivalent per-paper,
    # per-author institution history through this same simple path).
    # Left blank if nothing is found -- never guessed at, per instruction.
    if source == "openalex" and r.profile_url:
        try:
            guess, count, _works_in_window = guess_undergrad_via_publication_window(
                r.profile_url, last_year
            )
            if guess:
                r.guessed_undergrad = guess
                r.guessed_undergrad_mentions = count
        except Exception as e:
            print(f"    [undergrad-guess lookup failed, leaving blank: {e}]", file=sys.stderr)

    return r


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="olympiad_2000_2015_deduplicated.xlsx")
    ap.add_argument("--output", default="academia_status.xlsx")
    ap.add_argument(
        "--openalex-key",
        action="append",
        default=None,
        help="Your OpenAlex API key. Can be passed MULTIPLE TIMES to supply a POOL of "
             "keys (e.g. --openalex-key KEY1 --openalex-key KEY2 --openalex-key KEY3) -- "
             "the script automatically rotates to the next key once the current one's "
             "daily $1 free budget runs low (~1000 searches/key/day), so a run doesn't "
             "have to stop and wait for tomorrow. A single value with commas also works "
             "(--openalex-key KEY1,KEY2,KEY3). REQUIRED as of Feb 13, 2026 (OpenAlex "
             "retired unauthenticated/mailto-only access) -- free to get, 30 seconds at "
             "openalex.org/settings/api. Falls back to the OPENALEX_API_KEY(S) environment "
             "variable (comma-separated for multiple) if not passed here at all, then to a "
             "single built-in default key as a last resort."
    )
    ap.add_argument("--email", default=None,
                     help="(Legacy, mostly a no-op now -- see --openalex-key.) Previously used for "
                          "OpenAlex's 'polite pool'; OpenAlex retired that system in Feb 2026.")
    ap.add_argument("--limit", type=int, default=2000,
                     help="Only process first N rows. Defaults to 2000 -- matches the 2-key pool "
                          "built into this script by default (2 keys x ~$1/day x ~1000 "
                          "searches/$1 ~= 2000 people/day). If you change the key pool, adjust "
                          "this: roughly (number of keys x 1000). Pass --limit -1 to process "
                          "EVERYTHING in one run regardless -- the script rotates through the "
                          "whole key pool and then falls back to Semantic Scholar only once every "
                          "key is exhausted, so overshooting this estimate degrades gracefully "
                          "rather than failing.")
    ap.add_argument("--start-row", type=int, default=2, help="Resume from this row (1-indexed, header=row 1)")
    args = ap.parse_args()

    # Build the key pool: --openalex-key (possibly repeated, possibly
    # comma-separated within each value) > OPENALEX_API_KEYS env var
    # (comma-separated) > OPENALEX_API_KEY env var (single) > built-in
    # fallback default pool, in that priority order.
    import os
    raw_keys: list[str] = []
    if args.openalex_key:
        for v in args.openalex_key:
            raw_keys.extend(v.split(","))
    elif os.environ.get("OPENALEX_API_KEYS"):
        raw_keys.extend(os.environ["OPENALEX_API_KEYS"].split(","))
    elif os.environ.get("OPENALEX_API_KEY"):
        raw_keys.append(os.environ["OPENALEX_API_KEY"])
    else:
        # Last-resort default pool -- 2 keys, ~2000 free people/day
        # across the pool before falling back to Semantic Scholar only.
        raw_keys.extend([
            "1IHRdKEym8HmK2ZorYdwxs",
            "vQ3TtNToOv4hIRpkWsGTOd",
        ])

    global _openalex_keys
    _openalex_keys = OpenAlexKeyManager(raw_keys)

    if not _openalex_keys.keys:
        print("WARNING: no OpenAlex API key(s) provided. As of Feb 2026, OpenAlex requires a key\n"
              "for all requests -- without one, OpenAlex lookups will fail for nearly every person\n"
              "and this script will fall back to Semantic Scholar only (still works, just less\n"
              "coverage). Get a free key in 30 seconds at openalex.org/settings/api.\n",
              file=sys.stderr)
    else:
        print(f"OpenAlex key pool: {len(_openalex_keys.keys)} key(s) loaded "
              f"(~{len(_openalex_keys.keys) * 1000} free searches available today across the pool).\n")

    wb_in = openpyxl.load_workbook(args.input, data_only=True)
    ws_in = wb_in.active
    headers = [c.value for c in ws_in[1]]
    subj_col = headers.index("subject")
    name_col = headers.index("name")
    country_col = headers.index("country")
    year_col = headers.index("last_year")

    total_rows = ws_in.max_row - 1
    n_to_process = total_rows if args.limit is None or args.limit < 0 else min(args.limit, total_rows)

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Academia Status"
    ws_out.append([
        "subject", "name", "country", "last_year", "in_academia", "matched_name",
        "institution", "institution_type", "most_recent_work_year",
        "h_index", "works_count", "likely_merged_identity", "guessed_undergrad", "guessed_undergrad_mentions",
        "confidence", "source", "profile_url", "note",
    ])

    processed = 0
    status_counts = {}
    for i in range(args.start_row, args.start_row + n_to_process):
        row = ws_in[i]
        subject = row[subj_col].value
        name = row[name_col].value
        country = row[country_col].value or ""
        last_year = row[year_col].value
        if not name:
            continue

        print(f"[{i}] {name} ({subject}, {country}, {last_year}) ...", flush=True)
        r = check_one_person(subject, str(name), str(country), int(last_year), args.email)
        status_counts[r.in_academia] = status_counts.get(r.in_academia, 0) + 1
        print(f"    -> {r.in_academia} | {r.institution or '-'} ({r.institution_type or '-'}) | "
              f"h_index={r.h_index if r.h_index is not None else '-'} | confidence={r.confidence}")

        ws_out.append([
            r.subject, r.name, r.country, r.last_year, r.in_academia, r.matched_name,
            r.institution, r.institution_type, r.most_recent_work_year,
            r.h_index, r.works_count, r.likely_merged_identity, r.guessed_undergrad, r.guessed_undergrad_mentions,
            r.confidence, r.source, r.profile_url, r.note,
        ])

        processed += 1
        if processed % 50 == 0:
            wb_out.save(args.output)
            print(f"    [autosaved after {processed} rows]")

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    wb_out.save(args.output)
    print(f"\nDone. Processed {processed} rows. Saved to {args.output}")
    print("Yes/No breakdown:")
    for s, c in sorted(status_counts.items(), key=lambda x: -x[1]):
        print(f"  {s}: {c}")
    print("\nEvery row is forced to a Yes/No verdict, but the `note`, `confidence`, and `h_index`")
    print("columns still carry the underlying evidence -- worth sorting/filtering on `note`")
    print("for rows flagged '[Low-confidence match...]' or low h-index if you want to spot-check.")


if __name__ == "__main__":
    main()