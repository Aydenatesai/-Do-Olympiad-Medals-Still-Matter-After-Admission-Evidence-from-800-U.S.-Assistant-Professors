"""
Faculty CV / Website Finder (verified) -- profile-URL-first, Tavily for dept sites
====================================================================================

SEARCH PRIORITY (mới):
  1. Extract personal website link từ profile page → trả về ngay
  2. Tìm CV trên profile page → tải về
  3. Tìm CV trên derived/Tavily pages → tải về
  4. Fallback: trả về profile URL

Unlike the earlier "Physics batch" version, this assumes your input sheet
already has an INDIVIDUAL profile page per person (e.g. a
"Department Profile URL" column like
https://www.mcb.harvard.edu/directory/ryan-nett/) rather than one shared
department homepage for many people. So there's no need to crawl a
directory page and guess which link matches which name -- we already know
exactly which page belongs to each person.

For each person:
  1. Profile page -> Website link (NEW: ưu tiên nhất)
       Fetch their individual profile page. Look for a "Website" link
       and return it immediately.
  2. Profile page -> CV
       Look for a CV (PDF or Google Drive link) directly on that page.
  3. Tavily (once per department) -> faculty listing page -> derived
     personal page -> CV
  4. Nothing found anywhere -> record the profile URL itself as fallback.
  5. Results are written as a NEW TAB in the existing results workbook.

SETUP:
  pip install openpyxl requests beautifulsoup4 pypdf
  Set TAVILY_API_KEY below (get one at https://tavily.com).

USAGE:
  1. Set INPUT_XLSX to your people-list file (must have a
     "Department Profile URL" column).
  2. Set RESULTS_XLSX to your existing cv_or_website_results.xlsx.
  3. Set NEW_SHEET_NAME / CV_FOLDER for this batch (e.g. "Biochem").
  4. python find_cv_or_website.py
"""

from __future__ import annotations

import io
import os
import re
import sys
import traceback
import unicodedata
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font
from pypdf import PdfReader

# Force unbuffered/line-buffered stdout
try:
    sys.stdout.reconfigure(line_buffering=True)
except AttributeError:
    pass

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

REQUEST_TIMEOUT = 20

INPUT_XLSX = "Computer Science.xlsx"
RESULTS_XLSX = "cv_or_website_results.xlsx"
NEW_SHEET_NAME = "CV or Website - CS"
CV_FOLDER = "Computer Sci"

# --- Tavily -----------------------------------------------------------------
TAVILY_API_KEY = "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"
TAVILY_API_URL = "https://api.tavily.com/search"
TAVILY_MAX_RESULTS = 5

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CV_FOLDER_ABS = os.path.join(_SCRIPT_DIR, CV_FOLDER)

CONTACT_EMAIL = "hoangndsy7@gmail.com"
HEADERS = {"User-Agent": f"faculty-cv-finder (mailto:{CONTACT_EMAIL})"}

PDF_URL_RE = re.compile(r"\.pdf($|\?)", re.IGNORECASE)
CV_HINT_RE = re.compile(r"(?:^|[^a-z])(cv|resume|r[eé]sum[eé]|curriculum[\s_-]?vitae)(?:[^a-z]|$)", re.IGNORECASE)
EMPTY_VALUES = {"", "na", "n/a", "none", "-"}

NON_WEBSITE_DOMAIN_SIGNALS = [
    "linkedin.com", "researchgate.net", "scholar.google", "facebook.com",
    "twitter.com", "x.com", "instagram.com", "wikipedia.org",
]

GDRIVE_ID_RES = [
    re.compile(r"drive\.google\.com/file/d/([a-zA-Z0-9_-]+)"),
    re.compile(r"drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)"),
    re.compile(r"drive\.google\.com/uc\?.*[?&]id=([a-zA-Z0-9_-]+)"),
    re.compile(r"[?&]id=([a-zA-Z0-9_-]+)"),
]

CV_SECTION_KEYWORDS = [
    "education", "employment", "positions held", "appointments",
    "publications", "research interests", "teaching experience",
    "awards", "honors", "honours", "professional experience",
    "work experience", "grants", "curriculum vitae",
]
PAPER_SIGNAL_RE = re.compile(r"^\s*abstract\b", re.IGNORECASE)
TIMETABLE_SIGNAL_RE = re.compile(r"\b(MWF|TTh|MW|Lecture|Office Hours|Room \d)\b")

# URL patterns strongly suggesting research papers (not CVs)
PAPER_URL_PATTERNS = re.compile(
    r"/(paper|papers|proceedings|pubs|publications|preprint|preprints|"
    r"article|articles|pdf/\d{4}|abs/\d{4})/|"
    r"(arxiv|biorxiv|medrxiv|neurips|nips\.cc|proceedings\.|ecva\.net|"
    r"aclanthology|openreview|jmlr|ieee\.org|acm\.org|springer\.com|"
    r"sciencedirect|nature\.com|science\.org|cell\.com|wiley\.com|"
    r"pnas\.org|plos\.org|frontiersin|mdpi\.com)",
    re.IGNORECASE,
)

# Additional research paper signals in text
PAPER_STRUCTURE_RE = re.compile(
    r"\b(abstract|introduction|related work|methodology|methods|"
    r"experiments?|results|conclusion|references|acknowledgments?|"
    r"we propose|we present|in this paper|our approach|our method|"
    r"figure \d|table \d|equation \d|section \d|"
    r"arxiv:|doi:|proceedings of|conference on)\b",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Basic helpers
# ---------------------------------------------------------------------------

def safe_filename(name: str) -> str:
    normalized = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    cleaned = re.sub(r"[^\w\s-]", "", normalized).strip().replace(" ", "_")
    return cleaned or "unknown"


def is_empty(value: str) -> bool:
    return (value or "").strip().lower() in EMPTY_VALUES


def is_gdrive_link(url: str) -> bool:
    return "drive.google.com" in url.lower()


def extract_gdrive_id(url: str) -> str | None:
    for pattern in GDRIVE_ID_RES:
        m = pattern.search(url)
        if m:
            return m.group(1)
    return None


def looks_like_pdf_bytes(content: bytes) -> bool:
    return content[:5] == b"%PDF-"


def fetch_page(url: str) -> BeautifulSoup | None:
    try:
        resp = requests.get(url, timeout=REQUEST_TIMEOUT, headers=HEADERS)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")
    except requests.RequestException as e:
        print(f"    [!] couldn't fetch {url}: {e}")
        return None
    except Exception as e:
        print(f"    [!] couldn't parse {url}: {e}")
        return None


# ---------------------------------------------------------------------------
# NEW: Detect department homepage & navigate to individual faculty page
# ---------------------------------------------------------------------------

def is_department_homepage(url: str) -> bool:
    """Detect if URL is a DEPARTMENT page (not an individual profile
    and not a personal website).
    
    Signals of DEPARTMENT page:
      - Path is "/" or very short (e.g., cs.cornell.edu/, chem.umd.edu/)
      - Path contains department indicators like "/departments/", "/schools/"
      - URL ends with department page (e.g., .../computer-science.html)
      - Domain contains department keywords (cs, eng, chem, bio, phys, math)
      - Domain is .edu (institutional)
    
    Signals of PERSONAL website (NOT department):
      - Domain is github.io, personal name domain
      - Domain contains person-like name (e.g., donglaiw.github.io)
    
    Signals of INDIVIDUAL profile (NOT homepage):
      - Path has clear identifier like /people/john-doe, /faculty/anjalie-field
      - Path ends with a person's name/identifier
    """
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    path = parsed.path.strip("/").lower()
    
    # Personal sites are NEVER department pages
    if "github.io" in domain:
        return False
    if "wordpress.com" in domain or "wix.com" in domain or "sites.google.com" in domain:
        return False
    if "vercel.app" in domain or "netlify.app" in domain or "gitlab.io" in domain:
        return False
    
    # Non-.edu (and non-.ac.) domains are usually personal
    if ".edu" not in domain and ".ac." not in domain:
        return False
    
    # SHORT PATH: definitely department homepage
    if len(path) < 3:
        return True
    
    # LONG PATH with department indicators
    # e.g., /bc-web/schools/morrissey/departments/computer-science.html
    # e.g., /schools/engineering/departments/cs
    # e.g., /academics/departments/computer-science
    DEPT_PATH_INDICATORS = re.compile(
        r"/(departments?|schools?|colleges?|academics?)/|"
        r"/(computer[\s_-]?science|engineering|mathematics|physics|"
        r"chemistry|biology|biochemistry)(\.html|\.htm|/?)$|"
        r"(department|school)[\s_-]?of[\s_-]",
        re.IGNORECASE,
    )
    
    if DEPT_PATH_INDICATORS.search("/" + path):
        # Additional check: must NOT have a person's name after "/faculty/" or "/people/"
        # If path has /faculty/name or /people/name, it's an INDIVIDUAL profile
        INDIVIDUAL_PROFILE_PATTERN = re.compile(
            r"/(faculty|people|directory|staff|profile|member|researchers?|display)/[a-z0-9_-]+/?$",
            re.IGNORECASE,
        )
        if INDIVIDUAL_PROFILE_PATTERN.search("/" + path):
            return False  # It's an individual profile, not dept homepage
        return True
    
    return False


FACULTY_LINK_TEXT_RE = re.compile(
    r"\b(faculty|people|directory|professors?|instructors?)\b",
    re.IGNORECASE,
)


def find_faculty_listing_on_homepage(homepage_url: str) -> str | None:
    """Find 'Faculty' or 'People' link on department homepage.
    Returns URL of the faculty listing page, or None if not found.
    
    Ranks candidates by URL quality:
      - Higher score for URLs ending in /people, /faculty, /directory
      - Lower score for URLs with /news/, /about/, /advising/, etc.
    """
    soup = fetch_page(homepage_url)
    if soup is None:
        return None
    
    homepage_domain = urlparse(homepage_url).netloc
    scored_candidates = []  # (score, url)
    
    # Bad URL patterns - subpages that aren't listings
    BAD_URL_PATTERNS = re.compile(
        r"/(news|about|advising|recruitment|hiring|jobs|events|calendar|"
        r"admissions|apply|awards|prize|honor|story|blog)/",
        re.IGNORECASE,
    )
    
    for a in soup.find_all("a", href=True):
        text = (a.get_text(strip=True) or "").lower()
        href = (a["href"] or "").strip()
        
        if not href or href.startswith("#"):
            continue
        
        # Look for text matching "faculty", "people", etc.
        if not FACULTY_LINK_TEXT_RE.search(text):
            continue
        
        full_url = urljoin(homepage_url, href)
        
        # Only same-domain links
        if urlparse(full_url).netloc != homepage_domain:
            continue
        
        # Skip PDFs and Google Drive
        if PDF_URL_RE.search(full_url) or is_gdrive_link(full_url):
            continue
        
        # Score URL quality
        url_path = urlparse(full_url).path.lower().rstrip("/")
        score = 0
        
        # High score: URL ends with clear listing keyword
        if url_path.endswith(("/people", "/faculty", "/directory", "/professors")):
            score = 10
        # Medium: contains but doesn't end with
        elif re.search(r"/(people|faculty|directory)/?$", url_path):
            score = 8
        elif re.search(r"/(people|faculty|directory)", url_path):
            score = 5
        else:
            score = 2
        
        # Penalty for bad patterns
        if BAD_URL_PATTERNS.search(url_path):
            score -= 15
        
        # Bonus if link text is short/exact ("Faculty" > "Meet Our Faculty & Staff")
        if len(text) < 15:
            score += 2
        
        scored_candidates.append((score, full_url))
    
    # Sort by score, remove duplicates
    scored_candidates.sort(key=lambda x: -x[0])
    seen = set()
    ranked = []
    for score, url in scored_candidates:
        if url not in seen and score > 0:
            seen.add(url)
            ranked.append((score, url))
    
    if ranked:
        print(f"    [homepage] found {len(ranked)} faculty listing candidate(s), ranked by score:")
        for score, url in ranked[:5]:
            print(f"      [score={score}] {url}")
        return ranked[0][1]  # Return highest scored
    
    return None


def navigate_to_individual_page(homepage_url: str, name: str,
                                  university: str, department: str) -> str | None:
    """Navigate from department homepage to an individual faculty page.
    
    Strategy:
      1. Find "Faculty"/"People" link on homepage
      2. Crawl that listing page for individual faculty links
      3. Match person's name against links
      4. Verify candidate is real individual profile
    
    Falls back to Tavily search for faculty listing if homepage doesn't
    have a clear link."""
    
    # Step 1: Find faculty listing link on homepage
    print(f"    [navigate] looking for faculty listing link on {homepage_url}")
    faculty_listing = find_faculty_listing_on_homepage(homepage_url)
    
    # Step 2: Fallback to Tavily if not found on homepage
    if not faculty_listing:
        print(f"    [navigate] no faculty listing found on homepage, using Tavily")
        faculty_listing = tavily_find_faculty_listing_page(university, department)
    
    if not faculty_listing:
        print(f"    [navigate] couldn't find faculty listing for {department} @ {university}")
        return None
    
    # Step 3: Crawl listing page and match name
    print(f"    [navigate] crawling faculty listing: {faculty_listing}")
    link_map, dept_pattern = build_faculty_link_map(faculty_listing)
    candidates = find_person_url_candidates_in_link_map(link_map, name)
    
    if not candidates:
        print(f"    [navigate] no candidates matching name '{name}' in listing")
        return None
    
    # Step 4: Verify each candidate
    for cand in candidates:
        ok, reason = check_website(cand, name)
        if ok:
            print(f"    [navigate] ✅ verified individual page: {cand}")
            return cand
        print(f"    [navigate] [x] rejected {cand}: {reason}")
    
    return None


# ---------------------------------------------------------------------------
# NEW: Extract personal website links from profile page
# ---------------------------------------------------------------------------

def find_website_links_on_profile(profile_url: str, full_name: str = "") -> list[str]:
    """Extract personal website links from profile page.
    Looks for links labeled 'website', 'homepage', 'personal site', etc.
    Filters out social media and other non-personal-website domains.
    
    Returns list of URLs, RANKED by likelihood of being real PW:
      - Domain contains last name (e.g., sotiraki.com) → highest score
      - github.io / netlify / vercel → highest score
      - Different domain (external) → higher score (real PW)
      - Same domain (department subpage) → lower score (probably not PW)
    """
    soup = fetch_page(profile_url)
    if soup is None:
        return []

    profile_domain = urlparse(profile_url).netloc.lower()
    
    # Extract last name for dynamic domain matching
    last_name_lower = ""
    if full_name:
        parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
        if parts:
            last_name_lower = re.sub(r"[^\w-]", "", parts[-1]).lower()
    
    # STRICT keywords - only these count as "website" markers
    WEBSITE_KEYWORDS = ["website", "homepage", "home page", "personal site", 
                        "personal website", "personal page", "webpage"]
    
    # Personal site domain indicators (highest score)
    PERSONAL_DOMAIN_PATTERNS = [
        "github.io", "netlify.app", "vercel.app", "gitlab.io",
        "wordpress.com", "sites.google.com", "wixsite.com",
    ]
    
    scored_candidates = []  # (score, url)
    
    for a in soup.find_all("a", href=True):
        href = (a["href"] or "").strip()
        text = (a.get_text(strip=True) or "").lower().strip()

        # STRICT: text must be EXACTLY or MOSTLY the keyword, not just contain it
        is_website_link = False
        for kw in WEBSITE_KEYWORDS:
            if text == kw or re.match(rf"^{re.escape(kw)}\b", text) or re.search(rf"\b{re.escape(kw)}$", text):
                is_website_link = True
                break
        
        if not is_website_link:
            continue

        # Skip empty or anchor links
        if not href or href.startswith("#"):
            continue

        # Skip social media
        if any(sig in href.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
            continue

        # Skip PDFs and Google Drive
        if PDF_URL_RE.search(href) or is_gdrive_link(href):
            continue

        # Resolve relative URLs
        full_url = urljoin(profile_url, href)
        candidate_domain = urlparse(full_url).netloc.lower()
        clean_domain = candidate_domain.replace("www.", "").split(".")[0]
        
        # Score
        score = 0
        
        # HIGHEST: domain contains person's last name (sotiraki.com, anuragkhandelwal.com)
        if last_name_lower and len(last_name_lower) >= 3 and last_name_lower in clean_domain:
            score = 25
        # HIGHEST: personal domain patterns (github.io, sites.google.com, etc.)
        elif any(p in candidate_domain for p in PERSONAL_DOMAIN_PATTERNS):
            score = 20
        # HIGH: external domain (not the same as profile domain)
        elif candidate_domain != profile_domain:
            score = 10
        # LOW: same domain (probably a department subpage)
        else:
            score = 2
        
        # Bonus: URL path is very short (root of personal site)
        path = urlparse(full_url).path.strip("/")
        if len(path) < 3:
            score += 3
        
        # Penalty: URL has bad path indicating not-a-PW
        BAD_PATH_PATTERNS = re.compile(
            r"/(alumni|professors|faculty|graduate|undergraduate|news|"
            r"about|resources|committee|admissions|apply|jobs|hiring)/",
            re.IGNORECASE,
        )
        if BAD_PATH_PATTERNS.search(full_url):
            score -= 15
        
        scored_candidates.append((score, full_url))

    # Sort by score descending, dedupe
    scored_candidates.sort(key=lambda x: -x[0])
    seen = set()
    ranked = []
    for score, url in scored_candidates:
        if url not in seen and score > 0:
            seen.add(url)
            ranked.append((score, url))
    
    if ranked:
        print(f"    [website] found {len(ranked)} personal website candidate(s), ranked:")
        for score, url in ranked[:5]:
            print(f"      [score={score}] {url}")
    
    return [url for _, url in ranked]


# ---------------------------------------------------------------------------
# Tavily: find the department's faculty LISTING page (once per department),
# then derive each person's individual page from the shared URL structure
# ---------------------------------------------------------------------------

_DEPARTMENT_LINK_MAP_CACHE: dict[tuple[str, str], tuple[dict[str, str], dict | None]] = {}

NON_PERSON_LINK_TEXT_RE = re.compile(
    r"^\s*(home|about|contact|news|events|research|admissions|apply|login|"
    r"search|resources|giving|alumni|calendar|directory|faculty|people|"
    r"staff|students|graduate|undergraduate|courses|prev(ious)?|next|"
    r"page\s*\d*|\d+)\s*$",
    re.IGNORECASE,
)
LISTING_URL_HINT_RE = re.compile(r"\b(people|faculty|directory|staff)\b", re.IGNORECASE)


def tavily_raw_search(query: str) -> list[dict]:
    """Shared low-level Tavily call -- returns the raw `results` list."""
    if not TAVILY_API_KEY or TAVILY_API_KEY == "PASTE_YOUR_TAVILY_API_KEY_HERE":
        print("    [!] TAVILY_API_KEY not set -- skipping Tavily search")
        return []

    try:
        resp = requests.post(
            TAVILY_API_URL,
            json={
                "api_key": TAVILY_API_KEY,
                "query": query,
                "search_depth": "basic",
                "max_results": TAVILY_MAX_RESULTS,
            },
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        print(f"    [!] Tavily search failed for {query!r}: {e}")
        return []
    except ValueError as e:
        print(f"    [!] Tavily returned unparseable response for {query!r}: {e}")
        return []

    return data.get("results") or []


def _tavily_url_usable(url: str) -> bool:
    # Must have http/https scheme (reject relative URLs, /goto?url=..., etc.)
    if not url or not (url.startswith("http://") or url.startswith("https://")):
        return False
    
    # Reject Tavily's own redirect tracker URLs
    if "/goto?url=" in url:
        return False
    
    # Reject social media and other non-website domains
    if any(sig in url.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
        return False
    
    # Reject PDFs and Google Drive
    if PDF_URL_RE.search(url) or is_gdrive_link(url):
        return False
    
    return True


def tavily_find_faculty_listing_page(university: str, department: str) -> str | None:
    """One Tavily search per DEPARTMENT for the faculty listing/people page."""
    query = f"{department} {university} faculty directory people".strip()
    results = tavily_raw_search(query)
    if not results:
        print(f"    [tavily] no results for {query!r}")
        return None

    for r in results:
        url = (r.get("url") or "").strip()
        if _tavily_url_usable(url) and LISTING_URL_HINT_RE.search(url):
            print(f"    [tavily] '{query}' -> listing page: {url}")
            return url

    for r in results:
        url = (r.get("url") or "").strip()
        if _tavily_url_usable(url):
            print(f"    [tavily] '{query}' -> best-guess listing page: {url}")
            return url

    return None


def tavily_search_person_urls(full_name: str, university: str, department: str,
                                profile_url: str = "") -> list[str]:
    """Search Tavily for this person's individual DEPARTMENT PROFILE page with STRICT filters.
    
    Query strategy:
      - Use FULL NAME (not just last name — helps disambiguate common last names)
      - Add site:<root>.edu filter (root domain from profile_url)
      - Include department to disambiguate
    
    URL filters:
      1. Path must contain profile-like tag (/bio/, /profile/, /faculty/, /people/, /user/)
      2. After the tag, must have another segment (not just /faculty/ homepage)
      3. Page title must contain BOTH first name AND last name of the person
         (avoids matching different people with same last name in other departments)
    """
    # Build site filter from profile_url
    site_filter = ""
    if profile_url:
        try:
            domain = urlparse(profile_url).netloc
            if domain:
                root = get_root_domain(domain)
                site_filter = f" site:{root}"
        except Exception:
            pass
    
    # Extract first name + last name for title check
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return []
    last_name = re.sub(r"[^\w-]", "", parts[-1])
    first_name = re.sub(r"[^\w-]", "", parts[0])
    last_name_lower = last_name.lower()
    first_name_lower = first_name.lower()
    
    # Query: FULL NAME + department + site filter
    query = f"{full_name} {department}{site_filter}".strip()
    results = tavily_raw_search(query)
    
    # STRICT filter: URL must have profile-like path tag with extra segment
    PROFILE_TAG_RE = re.compile(
        r"/(bio|profile|faculty|faculty-directory|people|user|homes|members|"
        r"person|display|directory)/([^/?#]+)",
        re.IGNORECASE,
    )
    
    filtered_urls = []
    all_urls = []
    
    for r in results:
        url = (r.get("url") or "").strip()
        title = (r.get("title") or "").lower()
        
        if not _tavily_url_usable(url):
            continue
        all_urls.append(url)
        
        # Filter 1: URL must have profile tag + extra segment
        tag_match = PROFILE_TAG_RE.search(urlparse(url).path)
        if not tag_match:
            continue
        segment = tag_match.group(2).strip()
        if not segment or segment.lower() in ["index", "home", "list"]:
            continue
        
        # Filter 2: Title must contain BOTH first name AND last name
        # (avoids matching different people with same last name)
        title_has_last = last_name_lower and last_name_lower in title
        title_has_first = first_name_lower and first_name_lower in title
        
        if not title_has_last:
            continue
        # If first name is a single letter or very short, be lenient
        # (some titles use "F. Lastname" or "Fitzgerald, T.")
        if len(first_name_lower) >= 3 and not title_has_first:
            continue
        
        filtered_urls.append(url)
    
    # Log results
    if not filtered_urls and all_urls:
        print(f"    [tavily] '{query}' -> {len(all_urls)} raw result(s), ALL rejected by strict filter:")
        for u in all_urls[:5]:
            print(f"      ✗ {u}")
    elif filtered_urls:
        print(f"    [tavily] '{query}' -> {len(filtered_urls)} strict-filtered candidate(s):")
        for u in filtered_urls:
            print(f"      ✓ {u}")
    else:
        print(f"    [tavily] no results for {query!r}")
    
    return filtered_urls


def _dirname(url: str) -> str:
    """Parent path of a URL."""
    path = urlparse(url).path.rstrip("/")
    return path.rsplit("/", 1)[0] if "/" in path else path


def build_faculty_link_map(listing_url: str) -> tuple[dict[str, str], dict | None]:
    """Crawls faculty listing page and returns:
      - {link text (lowercased): url} map of individual faculty pages
      - department_pattern dict: {domain, dirname} used to verify future URLs
    
    Example pattern: {"domain": "engineering.jhu.edu", "dirname": "/faculty"}
    This lets us verify that a Tavily result URL like 
    "engineering.jhu.edu/faculty/anjalie-field" IS a real faculty page
    (matches pattern), while rejecting "cs.jhu.edu/news/xxx" (different path).
    """
    soup = fetch_page(listing_url)
    if soup is None:
        return {}, None

    listing_domain = urlparse(listing_url).netloc
    raw_links = []
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        if not text or NON_PERSON_LINK_TEXT_RE.match(text):
            continue
        full_url = urljoin(listing_url, a["href"])
        if any(sig in full_url.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
            continue
        if PDF_URL_RE.search(full_url) or is_gdrive_link(full_url):
            continue
        if urlparse(full_url).netloc != listing_domain:
            continue
        raw_links.append((text, full_url, _dirname(full_url)))

    if not raw_links:
        print(f"    [listing] no candidate links found on {listing_url}")
        return {}, None

    dirname_counts: dict[str, int] = {}
    for _, _, d in raw_links:
        dirname_counts[d] = dirname_counts.get(d, 0) + 1
    majority_dirname, majority_count = max(dirname_counts.items(), key=lambda kv: kv[1])

    if majority_count < 2:
        print(f"    [listing] no repeated URL structure found on {listing_url}")
        return {}, None

    dept_pattern = {"domain": listing_domain, "dirname": majority_dirname}
    print(f"    [listing] 🎯 department pattern detected: '{listing_domain}{majority_dirname}/<person>' "
          f"({majority_count} matching link(s))")

    link_map = {text.lower(): url for text, url, dirname in raw_links if dirname == majority_dirname}
    return link_map, dept_pattern


def get_root_domain(domain: str) -> str:
    """Extract root domain from a subdomain.
    
    Examples:
      engineering.vanderbilt.edu → vanderbilt.edu
      computing.vanderbilt.edu → vanderbilt.edu
      cse.wustl.edu → wustl.edu
      engineering.washu.edu → washu.edu (different root!)
      cs.yale.edu → yale.edu
    """
    # Take last 2 parts (e.g., yale.edu, vanderbilt.edu)
    parts = domain.lower().split(".")
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return domain.lower()


def matches_department_pattern(url: str, dept_pattern: dict) -> bool:
    """Check if URL matches the shared department faculty page pattern.
    
    RELAXED: accepts URLs from the same root domain even if subdomain differs.
    This is because universities often have multiple valid faculty domains
    (e.g., cs.yale.edu AND cpsc.yale.edu, engineering.vanderbilt.edu AND 
    computing.vanderbilt.edu).
    
    Example pattern: {"domain": "engineering.jhu.edu", "dirname": "/faculty"}
    ✅ engineering.jhu.edu/faculty/anjalie-field  (exact match)
    ✅ cs.jhu.edu/people/anjalie-field           (same root domain, path has faculty-like structure)
    ❌ schmidtsciences.org/grantee/anjalie-field (different root domain)
    ❌ engineering.jhu.edu/news/xxx              (same domain but news path)
    """
    if not dept_pattern:
        return False
    
    parsed = urlparse(url)
    url_domain = parsed.netloc
    pattern_domain = dept_pattern["domain"]
    
    # Same root domain?
    if get_root_domain(url_domain) != get_root_domain(pattern_domain):
        return False
    
    # Now check path - must NOT be a news/story/blog path
    url_path = parsed.path.lower()
    if NON_PROFILE_URL_PATTERNS.search(url_path):
        return False
    
    # Reject if URL ends with listing path (e.g., ends with /faculty, /people)
    stripped = url_path.rstrip("/")
    LISTING_PATH_ENDINGS = ["/faculty", "/people", "/directory", "/staff", 
                            "/members", "/researchers", "/professors"]
    if any(stripped.endswith(ending) for ending in LISTING_PATH_ENDINGS):
        return False
    
    # For exact pattern match (same subdomain), also require same dirname
    # For root-domain match (different subdomain), accept if path looks profile-like
    if url_domain == pattern_domain:
        # Same subdomain: require exact dirname match OR unix-style path (~name)
        url_dirname = _dirname(url)
        if url_dirname == dept_pattern["dirname"]:
            return True
        # Also accept /~lastname pattern
        if re.search(r"/~[a-z0-9_-]+", url_path, re.IGNORECASE):
            return True
        return False
    else:
        # Different subdomain but same root domain: accept if path looks like a profile
        # Must have /faculty/, /people/, /profile/, /bio/, /~ etc.
        PROFILE_PATH_INDICATORS = re.compile(
            r"/(faculty|people|profile|bio|homes?|members?|directory|staff|"
            r"researchers?|professors?|display)/[a-z0-9_-]+|/~[a-z0-9_-]+",
            re.IGNORECASE,
        )
        return bool(PROFILE_PATH_INDICATORS.search(url_path))


def get_department_link_map(university: str, department: str) -> tuple[dict[str, str], dict | None]:
    """Cached per (university, department). Returns (link_map, dept_pattern)."""
    key = (university.strip().lower(), department.strip().lower())
    if key in _DEPARTMENT_LINK_MAP_CACHE:
        return _DEPARTMENT_LINK_MAP_CACHE[key]

    listing_url = tavily_find_faculty_listing_page(university, department)
    if listing_url:
        link_map, dept_pattern = build_faculty_link_map(listing_url)
    else:
        link_map, dept_pattern = {}, None
    
    _DEPARTMENT_LINK_MAP_CACHE[key] = (link_map, dept_pattern)
    return link_map, dept_pattern


def find_person_url_candidates_in_link_map(link_map: dict[str, str], full_name: str) -> list[str]:
    """Matches a person's name against department's link map.
    
    Handles multiple name formats:
      - "Beidi Chen" (First Last)
      - "Chen, Beidi" (Last, First)
      - "B. Chen" (Initial Last)
      - "Chen" (Last only)
    """
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return []
    last_name = re.sub(r"[^\w-]", "", parts[-1]).lower()
    first_name = re.sub(r"[^\w-]", "", parts[0]).lower()
    if not last_name:
        return []

    scored: list[tuple[int, str]] = []
    for text, url in link_map.items():
        text_clean = text.lower().strip()
        
        # Skip if last name not present at all
        if last_name not in text_clean:
            continue
        
        # Also check that the URL contains something name-related
        # (avoid matching random text that mentions the name)
        url_lower = url.lower()
        url_has_name = last_name in url_lower or (first_name and first_name in url_lower)
        
        # Score by match quality
        score = 0
        
        # Best: both first and last name in text
        if first_name and first_name in text_clean:
            score = 4
        # Good: initial + last name (e.g., "B. Chen" for "Beidi Chen")
        elif first_name and re.search(rf"\b{re.escape(first_name[0])}\.?\s+{re.escape(last_name)}\b", text_clean):
            score = 3
        # OK: last name only
        else:
            score = 1
        
        # Bonus if URL also has name
        if url_has_name:
            score += 2
        
        scored.append((score, url))
    
    scored.sort(key=lambda pair: -pair[0])

    seen: set[str] = set()
    ordered: list[str] = []
    for _, url in scored:
        if url not in seen:
            seen.add(url)
            ordered.append(url)
    return ordered


NAME_LIKE_LINK_RE = re.compile(r"^[A-Z][a-zA-Z.'-]+(?:\s+[A-Z][a-zA-Z.'-]+){1,2}$")

# URLs that are NEWS / STORY / BLOG pages, NOT profile pages
NON_PROFILE_URL_PATTERNS = re.compile(
    r"/(news|story|stories|blog|press|announcements?|articles?|"
    r"post|posts|events?|awards?|prize|honor|welcome|grantee|"
    r"faculty-qa|q-and-a|interview|spotlight|profile-story)/",
    re.IGNORECASE,
)

# Title/heading signals for news articles (NOT profile pages)
NEWS_TITLE_SIGNALS = re.compile(
    r"\b(new faculty|welcomes?|joins?|announces?|q\s*&?a|"
    r"interview|spotlight|congratulations|awarded|"
    r"receives?\s+award|news:|press release)\b",
    re.IGNORECASE,
)


def check_website(url: str, full_name: str) -> tuple[bool, str]:
    """Fact-check that a candidate URL is THIS person's faculty profile page."""
    
    # --- Check 0: URL pattern (early exit for news/story/blog pages) ---
    if NON_PROFILE_URL_PATTERNS.search(url):
        return False, f"URL pattern suggests news/story/blog page, not a profile"
    
    # --- Check 0b: URL ends with listing-like path ---
    # e.g. /faculty, /people, /directory, /members WITHOUT anything after
    url_path = urlparse(url).path.rstrip("/").lower()
    LISTING_PATH_ENDINGS = ["/faculty", "/people", "/directory", "/staff", 
                            "/members", "/researchers", "/professors"]
    if any(url_path.endswith(ending) for ending in LISTING_PATH_ENDINGS):
        return False, f"URL ends with listing path ({url_path}), not an individual profile"
    
    soup = fetch_page(url)
    if soup is None:
        return False, "couldn't fetch page"

    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    last_name = re.sub(r"[^\w-]", "", parts[-1].lower()) if parts else ""
    first_name = re.sub(r"[^\w-]", "", parts[0].lower()) if parts else ""
    if not last_name:
        return False, "no usable name to check against"

    page_text = soup.get_text(" ", strip=True).lower()
    if last_name not in page_text:
        return False, f"person's name ('{last_name}') not found anywhere on the page"

    title_text = soup.title.get_text(strip=True) if soup.title else ""
    heading_text = " ".join(h.get_text(strip=True) for h in soup.find_all(["h1", "h2"])[:3])
    combined_heading = f"{title_text} {heading_text}".lower()
    name_in_heading = last_name in combined_heading or (first_name and first_name in combined_heading)

    # --- Check title/heading for news-article signals ---
    if NEWS_TITLE_SIGNALS.search(title_text) or NEWS_TITLE_SIGNALS.search(heading_text):
        return False, (f"title/heading has news-article signals "
                       f"(title: '{title_text[:80]}'), not a profile page")
    
    # --- Check title for listing-page signals ---
    # Only check TITLE, not headings (headings often include nav text like "Faculty | People | Research")
    LISTING_TITLE_SIGNALS = re.compile(
        r"^\s*(faculty|people|directory|staff|members|researchers|"
        r"professors|our (faculty|team|people)|meet (our|the))\b",
        re.IGNORECASE,
    )
    if LISTING_TITLE_SIGNALS.search(title_text):
        return False, (f"title looks like a listing page "
                       f"(title: '{title_text[:80]}'), not an individual profile")

    other_name_like_links = sum(
        1 for a in soup.find_all("a", href=True)
        if NAME_LIKE_LINK_RE.match(a.get_text(strip=True))
    )
    # ⭐ STRICTER: reduced from 8 to 5 (listing pages have many names)
    looks_like_listing = other_name_like_links >= 5

    if looks_like_listing and not name_in_heading:
        return False, (f"page has {other_name_like_links} other name-shaped links and "
                        f"'{last_name}' isn't in the title/heading -- looks like a listing")

    if not name_in_heading:
        return False, (f"'{last_name}' appears on page but not in title/heading -- "
                        f"likely not {full_name}'s own profile")

    return True, "ok"


# ---------------------------------------------------------------------------
# CV discovery on a given page
# ---------------------------------------------------------------------------

def find_pdf_or_drive_candidates(url: str) -> tuple[list[str], list[str]]:
    soup = fetch_page(url)
    if soup is None:
        return [], []

    links = soup.find_all("a", href=True)
    direct_labeled, direct_any, subpages = [], [], []

    for a in links:
        full_url = urljoin(url, a["href"])
        label = f'{a.get_text(strip=True)} {a["href"]}'
        is_direct = bool(PDF_URL_RE.search(full_url)) or is_gdrive_link(full_url)
        is_cv_labeled = bool(CV_HINT_RE.search(label))

        if is_direct and is_cv_labeled:
            direct_labeled.append(full_url)
        elif is_direct:
            direct_any.append(full_url)
        elif is_cv_labeled:
            subpages.append(full_url)

    return direct_labeled + direct_any, subpages


def fetch_bytes(url: str) -> bytes | None:
    try:
        resp = requests.get(url, timeout=REQUEST_TIMEOUT, headers=HEADERS)
        resp.raise_for_status()
        return resp.content
    except requests.RequestException as e:
        print(f"    [!] fetch failed for {url}: {e}")
        return None


def fetch_gdrive_bytes(url: str) -> bytes | None:
    file_id = extract_gdrive_id(url)
    if not file_id:
        return None
    session = requests.Session()
    endpoint = "https://drive.google.com/uc"
    try:
        resp = session.get(endpoint, params={"id": file_id, "export": "download"},
                            timeout=REQUEST_TIMEOUT, headers=HEADERS, stream=True)
        resp.raise_for_status()

        token = None
        for key, value in resp.cookies.items():
            if key.startswith("download_warning"):
                token = value
        if not token and b"confirm=" in resp.content[:5000]:
            m = re.search(rb"confirm=([0-9A-Za-z_-]+)", resp.content[:5000])
            if m:
                token = m.group(1).decode()

        if token:
            resp = session.get(endpoint, params={"id": file_id, "export": "download", "confirm": token},
                                timeout=REQUEST_TIMEOUT, headers=HEADERS, stream=True)
            resp.raise_for_status()

        return resp.content
    except requests.RequestException as e:
        print(f"    [!] Drive fetch failed for {url}: {e}")
        return None


def get_candidate_bytes(url: str) -> bytes | None:
    content = fetch_gdrive_bytes(url) if is_gdrive_link(url) else fetch_bytes(url)
    if content and looks_like_pdf_bytes(content):
        return content
    return None


def verify_pdf_is_cv(pdf_bytes: bytes, full_name: str, source_url: str = "") -> tuple[bool, str]:
    # --- Check 1: URL patterns (early exit for obvious papers) ---
    if source_url and PAPER_URL_PATTERNS.search(source_url):
        return False, f"URL pattern suggests research paper, not CV: {source_url}"
    
    # --- Check 2: PDF parsing ---
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "\n".join(p.extract_text() or "" for p in reader.pages[:3])
    except Exception as e:
        return False, f"couldn't parse PDF: {e}"

    if not text.strip():
        return False, "no extractable text (likely a scanned image)"

    text_lower = text.lower()
    
    # --- Check 3: Name in document ---
    last_name = re.sub(r"[^\w-]", "", full_name.strip().split()[-1].lower())
    if last_name and last_name not in text_lower:
        return False, f"person's name ('{last_name}') not found in document"

    # --- Check 4: URL label ---
    # If URL contains "cv", "resume", or "vitae" as a clear label, TRUST it
    # (academic CVs with long publication lists trigger paper-detection false positives)
    url_lower = source_url.lower() if source_url else ""
    URL_CV_LABEL_RE = re.compile(
        r"[/_-](cv|resume|vitae|curriculum[_-]?vitae)([/_.-]|$)",
        re.IGNORECASE,
    )
    url_labeled_as_cv = bool(URL_CV_LABEL_RE.search(url_lower))
    
    # --- Check 5: CV section keywords ---
    keyword_hits = sum(1 for kw in CV_SECTION_KEYWORDS if kw in text_lower)
    
    # --- Check 6: Research paper structure detection ---
    paper_signals = len(PAPER_STRUCTURE_RE.findall(text_lower))
    
    # --- Check 7: First 500 chars for paper signals ---
    first_500 = text.strip()[:500].lower()
    starts_with_abstract = bool(PAPER_SIGNAL_RE.match(first_500))
    
    # If URL clearly labeled as CV → RELAX checks (trust the URL)
    if url_labeled_as_cv:
        # Still reject if the document STARTS with "Abstract" (definitely a paper)
        if starts_with_abstract:
            return False, "URL says CV but document starts with 'Abstract' (likely a paper)"
        # Still require at least 1 CV keyword to make sure it's not empty
        if keyword_hits == 0 and paper_signals >= 8:
            return False, f"URL says CV but content looks like a paper ({paper_signals} paper signals, 0 CV keywords)"
        return True, f"URL labeled as CV ({keyword_hits} CV keywords, {paper_signals} paper signals)"
    
    # --- STRICT checks for non-URL-labeled documents ---
    
    # If document looks like a paper (multiple signals), reject
    if paper_signals >= 5:
        return False, (f"document has {paper_signals} research paper signals "
                       f"(abstract/introduction/methodology/etc.), not a CV")
    
    # If starts with "Abstract" AND has paper signals, reject
    if starts_with_abstract and paper_signals >= 2:
        return False, f"document starts with 'Abstract' and has paper structure"
    
    # Check timetable/syllabus
    timetable_hits = len(TIMETABLE_SIGNAL_RE.findall(text))
    if timetable_hits >= 3 and keyword_hits < 3:
        return False, "document looks like a timetable/syllabus, not a CV"

    # Require at least 3 CV keywords
    if keyword_hits < 3:
        return False, (f"only {keyword_hits} CV keyword(s) found -- "
                       f"real CVs have multiple sections (education, publications, awards, etc.)")
    
    # If paper signals ≥ keyword hits, it's probably a paper
    if paper_signals >= keyword_hits:
        return False, (f"{paper_signals} paper signals vs only {keyword_hits} CV keywords "
                       f"-- likely a research paper")

    return True, f"ok ({keyword_hits} CV keywords, {paper_signals} paper signals)"


def guess_cv_subpages(url: str) -> list[str]:
    """Guess common CV subpaths for a website.
    
    Example: base = https://sites.google.com/site/tapomayukh
    Returns: [
        https://sites.google.com/site/tapomayukh/cv,
        https://sites.google.com/site/tapomayukh/resume,
        ...
    ]
    """
    # Common CV subpaths
    CV_SUBPATHS = ["cv", "resume", "vitae", "curriculum-vitae", "curriculum_vitae",
                   "about/cv", "about/resume", "bio/cv"]
    
    # Normalize base URL (no trailing slash)
    base = url.rstrip("/")
    
    return [f"{base}/{sub}" for sub in CV_SUBPATHS]


def try_site_for_cv(url: str, full_name: str, depth: int = 0) -> tuple[bytes | None, str | None]:
    direct_candidates, subpages = find_pdf_or_drive_candidates(url)

    for candidate_url in direct_candidates:
        # Early URL check before downloading
        if PAPER_URL_PATTERNS.search(candidate_url):
            print(f"    [x] skipped {candidate_url}: URL pattern suggests research paper")
            continue
        
        content = get_candidate_bytes(candidate_url)
        if not content:
            continue
        ok, reason = verify_pdf_is_cv(content, full_name, source_url=candidate_url)
        if ok:
            return content, candidate_url
        print(f"    [x] rejected {candidate_url}: {reason}")

    if depth == 0:
        # Try labeled subpages first
        for subpage_url in subpages[:2]:
            content, source = try_site_for_cv(subpage_url, full_name, depth=1)
            if content:
                return content, source
        
        # NEW: Try guessed CV subpaths (e.g., /cv, /resume)
        # Only if we haven't already tried them via subpages
        already_tried = set(subpages[:2])
        for guessed_url in guess_cv_subpages(url):
            if guessed_url in already_tried:
                continue
            
            # Check if page exists (fetch might fail if no such page)
            print(f"    [guess] trying {guessed_url}")
            content, source = try_site_for_cv(guessed_url, full_name, depth=1)
            if content:
                return content, source

    return None, None


# ---------------------------------------------------------------------------
# Per-person pipeline (NEW SEARCH PRIORITY)
# ---------------------------------------------------------------------------

def tavily_search_personal_website(full_name: str, university: str, department: str) -> str | None:
    """Fallback search when department profile page can't be fetched.
    Searches for personal website matching one of these patterns:
      - Domain contains last name (sotiraki.com, anuragkhandelwal.com)
      - sites.google.com/site/<something>
      - <name>.github.io
    
    Query: "{last_name} {university} {department}"
    """
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return None
    last_name = re.sub(r"[^\w-]", "", parts[-1])
    last_name_lower = last_name.lower()
    
    # Use FULL NAME for better disambiguation
    query = f"{full_name} {university} {department}".strip()
    results = tavily_raw_search(query)
    
    print(f"    [tavily-pw] '{query}' -> checking for personal website")
    
    all_urls = []
    accepted = []
    
    for r in results:
        url = (r.get("url") or "").strip()
        title = (r.get("title") or "").lower()
        
        if not _tavily_url_usable(url):
            continue
        all_urls.append(url)
        
        # Check title contains last name (avoid unrelated people)
        if last_name_lower not in title:
            continue
        
        domain = urlparse(url).netloc.lower()
        
        # Accept if domain matches personal website patterns
        is_personal = False
        reason = ""
        
        # Pattern 1: Domain contains last name (sotiraki.com, anuragkhandelwal.com)
        # Remove common subdomains like "www."
        clean_domain = domain.replace("www.", "")
        if last_name_lower in clean_domain.split(".")[0]:
            is_personal = True
            reason = f"domain contains last name '{last_name_lower}'"
        
        # Pattern 2: github.io
        elif "github.io" in domain:
            is_personal = True
            reason = "github.io site"
        
        # Pattern 3: sites.google.com/site/...
        elif "sites.google.com" in domain and "/site/" in urlparse(url).path:
            is_personal = True
            reason = "sites.google.com site"
        
        if is_personal:
            print(f"      ✓ [{reason}] {url}")
            accepted.append(url)
        else:
            print(f"      ✗ {url}")
    
    if accepted:
        return accepted[0]
    
    if all_urls:
        print(f"    [tavily-pw] no personal website found among {len(all_urls)} result(s)")
    else:
        print(f"    [tavily-pw] no results")
    return None


def process_person(record: dict) -> dict | None:
    name = record["name"].strip()
    university = (record.get("university") or "").strip()
    department = (record.get("department") or "").strip()
    profile_url = (record.get("profile_url") or "").strip()

    if is_empty(profile_url):
        print(f"  -> no profile URL in sheet for {name!r}, skipping")
        return None

    department_website = profile_url  # fallback to department site
    
    print(f"\n  ========== SEARCH PRIORITY ==========")
    print(f"  1️⃣  Tavily → Individual Profile Page")
    print(f"  2️⃣  CV on Profile Page")
    print(f"  3️⃣  Personal Website (link on Profile) → CV on PW")
    print(f"  4️⃣  DW fallback")
    print(f"  =====================================\n")

    # --- STEP 1: Tavily search for INDIVIDUAL DEPARTMENT PROFILE PAGE ---
    print(f"  [1/4] Searching Tavily for {name}'s individual profile page")
    tavily_urls = tavily_search_person_urls(name, university, department, profile_url)
    
    verified_profile = None
    profile_fetch_failed = False
    
    for cand in tavily_urls:
        ok, reason = check_website(cand, name)
        if ok:
            verified_profile = cand
            print(f"  [1/4] ✅ VERIFIED individual profile: {cand}")
            break
        # Track if we saw a fetch failure (403, etc.) — indicator that Tavily found the right URL
        # but we can't access it (need to look for personal website instead)
        if "couldn't fetch" in reason.lower():
            profile_fetch_failed = True
        print(f"    [x] rejected {cand}: {reason}")
    
    # If profile page couldn't be verified for ANY reason (fetch fail, verification fail, 
    # or no Tavily results at all) → try personal website search
    if not verified_profile:
        if profile_fetch_failed:
            print(f"  [1/4] ⚠️  Profile page(s) found but couldn't fetch")
        elif tavily_urls:
            print(f"  [1/4] ⚠️  Tavily URLs found but none verified")
        else:
            print(f"  [1/4] ⚠️  Tavily didn't return any usable URLs")
        
        print(f"  [1/4] 🔄 Falling back to personal website search")
        personal_pw = tavily_search_personal_website(name, university, department)
        
        if personal_pw:
            print(f"  [1/4] ✅ Found personal website: {personal_pw}")
            # Try CV on personal website
            content, source = try_site_for_cv(personal_pw, name)
            if content:
                os.makedirs(CV_FOLDER_ABS, exist_ok=True)
                dest = os.path.join(CV_FOLDER_ABS, f"{safe_filename(name)}.pdf")
                with open(dest, "wb") as f:
                    f.write(content)
                print(f"  ✅ FOUND CV on personal website: {dest}")
                return {"university": university, "department": department, "name": name,
                        "result_type": "cv_download", "value": dest, "source_url": source}
            # No CV, return the personal website
            print(f"  ✅ No CV on PW, returning personal website: {personal_pw}")
            return {"university": university, "department": department, "name": name,
                    "result_type": "website", "value": personal_pw, "source_url": personal_pw}
    
    # If Tavily didn't find anything AND no personal website, use department URL as fallback
    if not verified_profile:
        print(f"  [1/4] ⚠️  No individual profile or personal website found")
        print(f"  [4/4] FALLBACK: returning department website")
        print(f"  ✅ FALLBACK: {department_website}")
        return {"university": university, "department": department, "name": name,
                "result_type": "website", "value": department_website, "source_url": department_website}

    # --- STEP 2: Look for CV on the verified profile page ---
    print(f"  [2/4] Looking for CV on profile page: {verified_profile}")
    content, source = try_site_for_cv(verified_profile, name)
    if content:
        os.makedirs(CV_FOLDER_ABS, exist_ok=True)
        dest = os.path.join(CV_FOLDER_ABS, f"{safe_filename(name)}.pdf")
        with open(dest, "wb") as f:
            f.write(content)
        print(f"  ✅ FOUND CV on profile: {dest}")
        return {"university": university, "department": department, "name": name,
                "result_type": "cv_download", "value": dest, "source_url": source}

    # --- STEP 3: Extract "Website" link from profile page, check CV on PW ---
    print(f"  [3/4] Looking for 'Website' link on profile page")
    website_links = find_website_links_on_profile(verified_profile, name)
    
    if website_links:
        for website_url in website_links:
            print(f"  [3/4] Checking CV on personal website: {website_url}")
            content, source = try_site_for_cv(website_url, name)
            if content:
                os.makedirs(CV_FOLDER_ABS, exist_ok=True)
                dest = os.path.join(CV_FOLDER_ABS, f"{safe_filename(name)}.pdf")
                with open(dest, "wb") as f:
                    f.write(content)
                print(f"  ✅ FOUND CV on PW: {dest}")
                return {"university": university, "department": department, "name": name,
                        "result_type": "cv_download", "value": dest, "source_url": source}
        
        # No CV on PW, return the PW (personal website is better than dept profile)
        best_website = website_links[0]
        print(f"  ✅ No CV on PW, returning personal website: {best_website}")
        return {"university": university, "department": department, "name": name,
                "result_type": "website", "value": best_website, "source_url": best_website}
    
    # No website link either, return the verified profile page
    print(f"  ✅ No CV/PW, returning verified profile page: {verified_profile}")
    return {"university": university, "department": department, "name": name,
            "result_type": "website", "value": verified_profile, "source_url": verified_profile}


# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------


def _normalize_header(h) -> str:
    return re.sub(r"\s+", " ", (h or "").strip()).lower()


HEADER_ALIASES = {
    "name": "name",
    "university": "university",
    "department": "department",
    "title": "title",
    "email": "email",
    "department profile url": "profile_url",
    "note": "note",
}


def load_people(path: str) -> list[dict]:
    wb = load_workbook(path)
    ws = wb.active
    raw_headers = [c.value for c in ws[1]]
    normalized_headers = [HEADER_ALIASES.get(_normalize_header(h), h) for h in raw_headers]

    print(f"[headers] raw columns: {raw_headers}")
    print(f"[headers] mapped to: {normalized_headers}")
    if "profile_url" not in normalized_headers:
        print("[headers] !!! WARNING: no 'profile_url' column -- every person will be skipped")

    people = []
    skipped_blank_name = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        record = dict(zip(normalized_headers, row))
        name = (record.get("name") or "").strip()
        if name:
            people.append(record)
        else:
            skipped_blank_name += 1

    print(f"[loaded] {len(people)} row(s) with name, {skipped_blank_name} skipped (blank name)")
    return people


def write_new_tab(rows: list[dict]) -> None:
    results_path = resolve_input_path(RESULTS_XLSX)
    if results_path is not None:
        wb = load_workbook(results_path)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        results_path = os.path.join(script_dir, RESULTS_XLSX)
        from openpyxl import Workbook
        wb = Workbook()
        wb.remove(wb.active)

    if NEW_SHEET_NAME in wb.sheetnames:
        del wb[NEW_SHEET_NAME]
    ws = wb.create_sheet(NEW_SHEET_NAME)

    headers = ["university", "department", "name", "result_type", "value", "source_url"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, name="Arial")

    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial")

    widths = [28, 18, 24, 16, 55, 55]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    wb.save(results_path)
    print(f"\n[saved] '{NEW_SHEET_NAME}' tab written to: {os.path.abspath(results_path)}")


def resolve_input_path(filename: str) -> str | None:
    cwd_path = os.path.join(os.getcwd(), filename)
    if os.path.exists(cwd_path):
        return cwd_path

    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_dir_path = os.path.join(script_dir, filename)
    if os.path.exists(script_dir_path):
        print(f"[paths] '{filename}' not found in cwd, using script folder: {script_dir_path}")
        return script_dir_path

    return None


def main():
    print(f"[cwd] working directory: {os.getcwd()}")
    print(f"[cwd] script location: {os.path.dirname(os.path.abspath(__file__))}")

    input_path = resolve_input_path(INPUT_XLSX)
    if input_path is None:
        print(f"Couldn't find {INPUT_XLSX!r} in cwd or script folder. Set INPUT_XLSX to full path.")
        return

    people = load_people(input_path)
    print(f"\nLoaded {len(people)} people\n")

    rows = []
    for i, record in enumerate(people, 1):
        print(f"\n{'='*60}")
        print(f"Person {i}/{len(people)}: {record.get('name')}")
        print(f"{'='*60}")
        row = process_person(record)
        if row:
            rows.append(row)

    write_new_tab(rows)
    downloaded = sum(1 for r in rows if r["result_type"] == "cv_download")
    websites = sum(1 for r in rows if r["result_type"] == "website")
    manuals = sum(1 for r in rows if r["result_type"] == "manual_search")
    print(f"\n{'='*60}")
    print(f"DONE ✅")
    print(f"{'='*60}")
    print(f"CVs downloaded: {downloaded}")
    print(f"Websites recorded: {websites}")
    print(f"Manual search needed: {manuals}")
    print(f"CV folder: {CV_FOLDER_ABS}")
    print(f"Results file: {RESULTS_XLSX} tab '{NEW_SHEET_NAME}'")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[!!!] CRASHED -- full error below:")
        traceback.print_exc()
        sys.exit(1)