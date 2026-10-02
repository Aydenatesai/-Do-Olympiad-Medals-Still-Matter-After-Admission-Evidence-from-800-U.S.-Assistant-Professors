"""
Faculty CV / Website Finder (verified) -- profile-URL-first, Tavily for dept sites
====================================================================================

SEARCH PRIORITY:
  1. Extract personal website link từ profile page → trả về ngay
  2. Tìm CV trên profile page → tải về
  3. Tìm CV trên derived/Tavily pages → tải về
  4. Fallback: trả về profile URL

NEW (lab/group detection): trước khi coi một "personal website" candidate
là trang cá nhân của riêng người đó, script giờ kiểm tra xem đó có phải là
LAB / RESEARCH GROUP site không (nhiều site kiểu này có domain như
"smithlab.stanford.edu" hoặc "sites.google.com/view/xyz-group", và liệt kê
NHIỀU thành viên chứ không chỉ 1 người). Nếu đúng là lab site:
    lab homepage → tìm link "Team"/"People"/"Members" → trên trang đó, tìm
    đúng entry của người có chức danh "Assistant Professor" (hoặc
    Professor/PI/Principal Investigator/Group Leader) VÀ tên khớp với
    người đang tìm (không chỉ khớp tên suông, vì trang Team thường có cả
    postdoc/grad student trùng tên hoặc tên gần giống) → nếu entry đó có
    link riêng (sub-page) thì đi vào đó tìm CV, nếu không thì tìm CV ngay
    trên trang Team đó → cuối cùng nếu vẫn không thấy, thử luôn trang chủ
    lab như trước đây.

NEW (multi-batch): thay vì chỉ chạy 1 file input, script bây giờ chạy qua
danh sách BATCHES -- mỗi batch tương ứng 1 khoa (Biochem / Physics / Math),
với input xlsx, folder CV và tên sheet riêng. Tất cả các sheet được ghi
vào CÙNG MỘT results workbook (RESULTS_XLSX), mỗi khoa 1 tab.

SETUP:
  pip install openpyxl requests beautifulsoup4 pypdf
  Set TAVILY_API_KEY below (get one at https://tavily.com).

USAGE:
  1. Sửa danh sách BATCHES bên dưới: mỗi phần tử là 1 khoa với
     input_xlsx / new_sheet_name / cv_folder riêng.
  2. Set RESULTS_XLSX (file kết quả dùng chung cho cả 3 khoa).
  3. python find_cv_or_website.py
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

# Shared results workbook -- every batch below writes its own tab into this
# same file, so run all 3 departments in one process and get 1 xlsx with
# 3 sheets + 3 separate CV folders on disk.
RESULTS_XLSX = "cv_or_website_results.xlsx"

# One entry per department/batch. Add / remove / edit freely.
BATCHES = [

    {
        "input_xlsx": "Biochem.xlsx",
        "new_sheet_name": "CV or Website - Biochem",
        "cv_folder": "Biochem",
    }]

# --- Tavily -----------------------------------------------------------------
TAVILY_API_KEY = "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"
TAVILY_API_URL = "https://api.tavily.com/search"
TAVILY_MAX_RESULTS = 5

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# These two are set per-batch at the start of each run in main() / process_batch().
# Keeping them as module-level globals (instead of threading a "folder" argument
# through every function) keeps the rest of the pipeline code identical to the
# single-batch version -- process_person() etc. just read the current value.
CV_FOLDER = BATCHES[0]["cv_folder"]
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
    DEPT_PATH_INDICATORS = re.compile(
        r"/(departments?|schools?|colleges?|academics?)/|"
        r"/(computer[\s_-]?science|engineering|mathematics|physics|"
        r"chemistry|biology|biochemistry)(\.html|\.htm|/?)$|"
        r"(department|school)[\s_-]?of[\s_-]",
        re.IGNORECASE,
    )

    if DEPT_PATH_INDICATORS.search("/" + path):
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

        if not FACULTY_LINK_TEXT_RE.search(text):
            continue

        full_url = urljoin(homepage_url, href)

        if urlparse(full_url).netloc != homepage_domain:
            continue

        if PDF_URL_RE.search(full_url) or is_gdrive_link(full_url):
            continue

        url_path = urlparse(full_url).path.lower().rstrip("/")
        score = 0

        if url_path.endswith(("/people", "/faculty", "/directory", "/professors")):
            score = 10
        elif re.search(r"/(people|faculty|directory)/?$", url_path):
            score = 8
        elif re.search(r"/(people|faculty|directory)", url_path):
            score = 5
        else:
            score = 2

        if BAD_URL_PATTERNS.search(url_path):
            score -= 15

        if len(text) < 15:
            score += 2

        scored_candidates.append((score, full_url))

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
        return ranked[0][1]

    return None


def navigate_to_individual_page(homepage_url: str, name: str,
                                  university: str, department: str) -> str | None:
    """Navigate from department homepage to an individual faculty page."""

    print(f"    [navigate] looking for faculty listing link on {homepage_url}")
    faculty_listing = find_faculty_listing_on_homepage(homepage_url)

    if not faculty_listing:
        print(f"    [navigate] no faculty listing found on homepage, using Tavily")
        faculty_listing = tavily_find_faculty_listing_page(university, department)

    if not faculty_listing:
        print(f"    [navigate] couldn't find faculty listing for {department} @ {university}")
        return None

    print(f"    [navigate] crawling faculty listing: {faculty_listing}")
    link_map, dept_pattern = build_faculty_link_map(faculty_listing)
    candidates = find_person_url_candidates_in_link_map(link_map, name)

    if not candidates:
        print(f"    [navigate] no candidates matching name '{name}' in listing")
        return None

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
    """Extract personal website links from profile page."""
    soup = fetch_page(profile_url)
    if soup is None:
        return []

    profile_domain = urlparse(profile_url).netloc.lower()

    last_name_lower = ""
    if full_name:
        parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
        if parts:
            last_name_lower = re.sub(r"[^\w-]", "", parts[-1]).lower()

    WEBSITE_KEYWORDS = ["website", "homepage", "home page", "personal site",
                        "personal website", "personal page", "webpage", "web page"]

    PERSONAL_DOMAIN_PATTERNS = [
        "github.io", "netlify.app", "vercel.app", "gitlab.io",
        "wordpress.com", "sites.google.com", "wixsite.com",
    ]

    scored_candidates = []  # (score, url)

    for a in soup.find_all("a", href=True):
        href = (a["href"] or "").strip()
        text = (a.get_text(strip=True) or "").lower().strip()

        is_website_link = False
        for kw in WEBSITE_KEYWORDS:
            if text == kw or re.match(rf"^{re.escape(kw)}\b", text) or re.search(rf"\b{re.escape(kw)}$", text):
                is_website_link = True
                break

        if not is_website_link:
            continue

        if not href or href.startswith("#"):
            continue

        if any(sig in href.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
            continue

        if PDF_URL_RE.search(href) or is_gdrive_link(href):
            continue

        full_url = urljoin(profile_url, href)
        candidate_domain = urlparse(full_url).netloc.lower()
        clean_domain = candidate_domain.replace("www.", "").split(".")[0]

        score = 0

        if last_name_lower and len(last_name_lower) >= 3 and last_name_lower in clean_domain:
            score = 25
        elif any(p in candidate_domain for p in PERSONAL_DOMAIN_PATTERNS):
            score = 20
        elif candidate_domain != profile_domain:
            score = 10
        else:
            score = 2

        path = urlparse(full_url).path.strip("/")
        if len(path) < 3:
            score += 3

        BAD_PATH_PATTERNS = re.compile(
            r"/(alumni|professors|faculty|graduate|undergraduate|news|"
            r"about|resources|committee|admissions|apply|jobs|hiring)/",
            re.IGNORECASE,
        )
        if BAD_PATH_PATTERNS.search(full_url):
            score -= 15

        scored_candidates.append((score, full_url))

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
# NEW: Lab / research-group site detection + team-page navigation
# ---------------------------------------------------------------------------
#
# Nhiều "personal website" thực chất là trang của LAB / RESEARCH GROUP
# (vd: "smithlab.stanford.edu", "sites.google.com/view/xyz-group"),
# liệt kê NHIỀU người (PI + postdoc + grad student), không chỉ 1 mình
# giáo sư đang tìm. Nếu áp dụng try_site_for_cv() thẳng lên trang chủ lab,
# CV của PI có thể nằm ở 1 sub-page riêng (vd "/people/jane-smith") mà
# trang chủ không link PDF trực tiếp tới, nên sẽ bị miss.
#
# Luồng mới: lab homepage → tìm link "Team"/"People"/"Members" → trên
# trang đó, xác nhận đúng entry của PI bằng cách khớp CẢ tên (name) LẪN
# chức danh giáo sư (title) trong cùng 1 khối text nhỏ (không phải khớp
# tên/chức danh rời rạc trên cả trang, để tránh nhầm với 1 thành viên
# khác trùng tên hoặc trùng chữ "Professor" ở đâu đó không liên quan)
# → nếu entry đó có sub-page riêng thì crawl CV ở đó, nếu không thì tìm
# CV ngay trên trang Team → fallback cuối: thử luôn trang chủ lab.

# Token-separated match: "xyz-lab.edu", "xyz.group.edu", "sites.google.com/view/abc-lab"
LAB_GROUP_URL_SIGNAL_RE = re.compile(
    r"(?:^|[./_-])(lab|labs|group|grp)(?:[./_-]|$)",
    re.IGNORECASE,
)

# Domain-first-label SUFFIX match with NO separator required: catches the
# very common "smithlab.stanford.edu" / "chenlab.rice.edu" naming pattern,
# where "lab" is glued directly onto the PI's surname. Anchored to the END
# of the domain's first label (not a mid-word substring search), so this
# does NOT false-positive on unrelated words that merely CONTAIN "lab" in
# the middle (e.g. "collaborative", "elaborate").
_LAB_GROUP_SUFFIX_RE = re.compile(r"(labs?|group|grp)$", re.IGNORECASE)


def _domain_first_label(url: str) -> str:
    domain = urlparse(url).netloc.lower()
    if domain.startswith("www."):
        domain = domain[4:]
    return domain.split(".")[0] if domain else ""

LAB_GROUP_TITLE_SIGNAL_RE = re.compile(
    r"\b(lab|laboratory|research\s+group)\b",
    re.IGNORECASE,
)

TEAM_LINK_TEXT_RE = re.compile(
    r"^\s*(team|people|members|group\s+members|our\s+team|our\s+group|"
    r"lab\s+members|meet\s+the\s+team)\s*$",
    re.IGNORECASE,
)

# Broad professor/PI title signal used to confirm a team-page ENTRY really
# is the lab's PI (as opposed to a postdoc/grad student entry).
PI_TITLE_RE = re.compile(
    r"\b(assistant\s+professor|associate\s+professor|professor|"
    r"principal\s+investigator|\bpi\b|group\s+leader|lab\s+director|"
    r"faculty\s+director)\b",
    re.IGNORECASE,
)

# This project specifically tracks ASSISTANT professors, so this narrower
# signal is checked first / preferred when a team page has more than one
# plausible PI-titled entry (rare, but e.g. a lab co-run by two faculty).
ASSISTANT_PROFESSOR_RE = re.compile(r"\bassistant\s+professor\b", re.IGNORECASE)

# Reused from check_website()'s "looks like a listing" heuristic --
# defined here too so this section works if copy-pasted standalone.
NAME_LIKE_LINK_RE = re.compile(r"^[A-Z][a-zA-Z.'-]+(?:\s+[A-Z][a-zA-Z.'-]+){1,2}$")


def is_lab_or_group_site(url: str, soup: BeautifulSoup | None = None) -> bool:
    """Detect if a URL is a LAB / RESEARCH GROUP site rather than an
    individual's own single-author personal page. Checked BEFORE treating
    a 'personal website' candidate as directly being the professor's own
    bio page, since a lab site's CV link for the PI is often on a
    separate Team/People sub-page instead.

    Signals checked, in order (first match wins):
      1. URL has 'lab'/'labs'/'group'/'grp' as a token separated by
         punctuation anywhere in it (e.g. "xyz-group.mit.edu",
         "sites.google.com/view/abc-lab").
      2. The domain's FIRST LABEL ends with 'lab'/'labs'/'group'/'grp'
         even with NO separator -- catches the very common
         "smithlab.stanford.edu" / "chenlab.rice.edu" naming pattern.
         Anchored to the end of the label (not a mid-word substring
         search), so this does NOT false-positive on words that merely
         CONTAIN "lab" in the middle (e.g. "collaborative.edu").
      3. Page <title>/<h1>/<h2> mentions 'Lab' / 'Laboratory' /
         'Research Group'.
      4. Page lists 3+ name-shaped links (same heuristic used elsewhere
         to detect "this is a listing of several people, not one person's
         own page").
    """
    if LAB_GROUP_URL_SIGNAL_RE.search(url):
        return True

    first_label = _domain_first_label(url)
    if first_label and _LAB_GROUP_SUFFIX_RE.search(first_label):
        return True

    if soup is None:
        return False

    title_text = soup.title.get_text(strip=True) if soup.title else ""
    heading_text = " ".join(h.get_text(strip=True) for h in soup.find_all(["h1", "h2"])[:3])
    combined = f"{title_text} {heading_text}"

    if LAB_GROUP_TITLE_SIGNAL_RE.search(combined):
        return True

    name_like_links = sum(
        1 for a in soup.find_all("a", href=True)
        if NAME_LIKE_LINK_RE.match(a.get_text(strip=True))
    )
    if name_like_links >= 3:
        return True

    return False


def find_team_page_on_lab_site(lab_url: str) -> str | None:
    """Find the lab's 'Team' / 'People' / 'Members' page. Mirrors
    find_faculty_listing_on_homepage() but for lab sites, whose nav
    links are usually labeled 'Team' / 'People' / 'Members' / 'Our Group'
    rather than 'Faculty' / 'Directory'."""
    soup = fetch_page(lab_url)
    if soup is None:
        return None

    lab_domain = urlparse(lab_url).netloc
    scored_candidates = []  # (score, url)

    for a in soup.find_all("a", href=True):
        text = (a.get_text(strip=True) or "").lower()
        href = (a["href"] or "").strip()
        if not href or href.startswith("#"):
            continue
        if not TEAM_LINK_TEXT_RE.match(text):
            continue

        full_url = urljoin(lab_url, href)
        if urlparse(full_url).netloc != lab_domain:
            continue
        if PDF_URL_RE.search(full_url) or is_gdrive_link(full_url):
            continue

        url_path = urlparse(full_url).path.lower().rstrip("/")
        score = 10 if url_path.endswith(("/team", "/people", "/members", "/group")) else 5
        scored_candidates.append((score, full_url))

    scored_candidates.sort(key=lambda x: -x[0])
    if scored_candidates:
        print(f"    [lab] found team/people page: {scored_candidates[0][1]}")
        return scored_candidates[0][1]

    print(f"    [lab] no dedicated team/people page found on {lab_url}")
    return None


def find_pi_entry_on_team_page(team_url: str, full_name: str) -> tuple[str | None, str | None]:
    """On a lab's Team/People page, find the entry for THIS SPECIFIC
    person (the assistant professor / PI being researched). Lab team
    pages usually also list postdocs and grad students, so matching on
    NAME ALONE risks landing on the wrong member if names are similar or
    a postdoc happens to share a last name -- matching on an 'Assistant
    Professor' / 'PI' TITLE appearing in the SAME small block of text as
    the name is a much stronger confirmation than matching either signal
    against the whole page independently.

    Returns (individual_subpage_url_or_None, reason):
      - (url, "own sub-page linked from team entry") if the matched
        entry links out to the PI's own page.
      - (None, "confirmed inline on team page") if the PI's bio is
        directly on the team page with no separate link -- caller should
        then search the team page ITSELF for a CV.
      - (None, None) if no entry could be confirmed at all.
    """
    soup = fetch_page(team_url)
    if soup is None:
        return None, None

    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return None, None
    last_name = re.sub(r"[^\w-]", "", parts[-1]).lower()
    first_name = re.sub(r"[^\w-]", "", parts[0]).lower()
    if not last_name:
        return None, None

    team_domain = urlparse(team_url).netloc

    # Scan reasonably-sized "member card" containers (div/li/section/
    # article/p) rather than the whole page text blob, so the title
    # match has to be LOCAL to this person's own entry -- not just
    # "the word Professor appears somewhere on the page".
    candidate_containers = soup.find_all(["div", "li", "section", "article", "p"])
    best_container = None
    best_container_text = ""

    for el in candidate_containers:
        text = el.get_text(" ", strip=True)
        if not text or len(text) > 600:  # skip giant page-wrapper containers
            continue
        text_lower = text.lower()
        if last_name not in text_lower:
            continue
        if first_name and len(first_name) >= 3 and first_name not in text_lower:
            continue
        if not PI_TITLE_RE.search(text):
            continue

        best_container = el
        best_container_text = text
        # An "assistant professor" match is as strong a confirmation as
        # this heuristic gets -- stop here rather than keep scanning for
        # a hypothetically even-better container.
        if ASSISTANT_PROFESSOR_RE.search(text):
            break

    if best_container is None:
        print(f"    [lab] no PI-titled entry matching '{full_name}' found on {team_url}")
        return None, None

    print(f"    [lab] confirmed PI entry on team page: ...{best_container_text[:120]}...")

    GENERIC_LINK_TEXT = {"email", "e-mail", "twitter", "x", "linkedin",
                          "google scholar", "scholar", "orcid", "github"}

    for a in best_container.find_all("a", href=True):
        href = (a["href"] or "").strip()
        if not href or href.startswith("#"):
            continue
        full_url = urljoin(team_url, href)
        if urlparse(full_url).netloc != team_domain:
            continue
        if PDF_URL_RE.search(full_url) or is_gdrive_link(full_url):
            continue
        if any(sig in full_url.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
            continue
        link_text = (a.get_text(strip=True) or "").lower()
        if link_text in GENERIC_LINK_TEXT:
            continue
        print(f"    [lab] PI has own sub-page: {full_url}")
        return full_url, "own sub-page linked from team entry"

    return None, "confirmed inline on team page"


def try_lab_site_for_cv(lab_url: str, full_name: str) -> tuple[bytes | None, str | None]:
    """Full lab-site flow: confirm it's a lab/group site, find the
    Team/People page, confirm the PI's own entry there (name + title
    match), then look for a CV either on the PI's own sub-page (if
    linked from their entry) or on the team page itself. Falls back to
    trying the lab homepage directly if there's no team page or no
    confirmed PI entry, since some smaller lab sites just put everything
    on the one page.

    Returns (None, None) immediately (no work done) if the URL doesn't
    look like a lab/group site at all -- callers should fall through to
    the normal try_site_for_cv() path for genuine single-person sites.
    """
    soup = fetch_page(lab_url)
    if not is_lab_or_group_site(lab_url, soup):
        return None, None

    print(f"    [lab] {lab_url} looks like a lab/research-group site")
    team_url = find_team_page_on_lab_site(lab_url)

    search_url = lab_url
    if team_url:
        pi_subpage, reason = find_pi_entry_on_team_page(team_url, full_name)
        if pi_subpage:
            search_url = pi_subpage
        elif reason == "confirmed inline on team page":
            search_url = team_url
        # else: PI not confirmed on the team page -- fall through and
        # just try the lab homepage itself below, still a reasonable bet.

    content, source = try_site_for_cv(search_url, full_name)
    if content:
        return content, source

    if search_url != lab_url:
        print(f"    [lab] no CV on {search_url}, trying lab homepage too")
        content, source = try_site_for_cv(lab_url, full_name)
        if content:
            return content, source

    return None, None


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
    if not url or not (url.startswith("http://") or url.startswith("https://")):
        return False

    if "/goto?url=" in url:
        return False

    if any(sig in url.lower() for sig in NON_WEBSITE_DOMAIN_SIGNALS):
        return False

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
    """Search Tavily for this person's individual DEPARTMENT PROFILE page with STRICT filters."""
    site_filter = ""
    if profile_url:
        try:
            domain = urlparse(profile_url).netloc
            if domain:
                root = get_root_domain(domain)
                site_filter = f" site:{root}"
        except Exception:
            pass

    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return []
    last_name = re.sub(r"[^\w-]", "", parts[-1])
    first_name = re.sub(r"[^\w-]", "", parts[0])
    last_name_lower = last_name.lower()
    first_name_lower = first_name.lower()

    query = f"{full_name} {department}{site_filter}".strip()
    results = tavily_raw_search(query)

    PROFILE_TAG_RE = re.compile(
        r"/(bio|profile|faculty|faculty-directory|people|user|homes|members|"
        r"person|display|directory)/([^/?#]+)|"
        r"/~([a-z0-9_-]+)",
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

        tag_match = PROFILE_TAG_RE.search(urlparse(url).path)
        if not tag_match:
            continue
        segment = (tag_match.group(2) or tag_match.group(3) or "").strip()
        if not segment or segment.lower() in ["index", "home", "list"]:
            continue

        title_has_last = last_name_lower and last_name_lower in title
        title_has_first = first_name_lower and first_name_lower in title

        if not title_has_last:
            continue
        if len(first_name_lower) >= 3 and not title_has_first:
            continue

        filtered_urls.append(url)

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
    """Crawls faculty listing page and returns link map + department pattern."""
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
    """Extract root domain from a subdomain."""
    parts = domain.lower().split(".")
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return domain.lower()


def matches_department_pattern(url: str, dept_pattern: dict) -> bool:
    """Check if URL matches the shared department faculty page pattern."""
    if not dept_pattern:
        return False

    parsed = urlparse(url)
    url_domain = parsed.netloc
    pattern_domain = dept_pattern["domain"]

    if get_root_domain(url_domain) != get_root_domain(pattern_domain):
        return False

    url_path = parsed.path.lower()
    if NON_PROFILE_URL_PATTERNS.search(url_path):
        return False

    stripped = url_path.rstrip("/")
    LISTING_PATH_ENDINGS = ["/faculty", "/people", "/directory", "/staff",
                            "/members", "/researchers", "/professors"]
    if any(stripped.endswith(ending) for ending in LISTING_PATH_ENDINGS):
        return False

    if url_domain == pattern_domain:
        url_dirname = _dirname(url)
        if url_dirname == dept_pattern["dirname"]:
            return True
        if re.search(r"/~[a-z0-9_-]+", url_path, re.IGNORECASE):
            return True
        return False
    else:
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
    """Matches a person's name against department's link map."""
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

        if last_name not in text_clean:
            continue

        url_lower = url.lower()
        url_has_name = last_name in url_lower or (first_name and first_name in url_lower)

        score = 0

        if first_name and first_name in text_clean:
            score = 4
        elif first_name and re.search(rf"\b{re.escape(first_name[0])}\.?\s+{re.escape(last_name)}\b", text_clean):
            score = 3
        else:
            score = 1

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


NON_PROFILE_URL_PATTERNS = re.compile(
    r"/(news|story|stories|blog|press|announcements?|articles?|"
    r"post|posts|events?|awards?|prize|honor|welcome|grantee|"
    r"faculty-qa|q-and-a|interview|spotlight|profile-story)/",
    re.IGNORECASE,
)

NEWS_TITLE_SIGNALS = re.compile(
    r"\b(new faculty|welcomes?|joins?|announces?|q\s*&?a|"
    r"interview|spotlight|congratulations|awarded|"
    r"receives?\s+award|news:|press release)\b",
    re.IGNORECASE,
)


def check_website(url: str, full_name: str) -> tuple[bool, str]:
    """Fact-check that a candidate URL is THIS person's faculty profile page."""

    if NON_PROFILE_URL_PATTERNS.search(url):
        return False, f"URL pattern suggests news/story/blog page, not a profile"

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

    if NEWS_TITLE_SIGNALS.search(title_text) or NEWS_TITLE_SIGNALS.search(heading_text):
        return False, (f"title/heading has news-article signals "
                       f"(title: '{title_text[:80]}'), not a profile page")

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
    if source_url and PAPER_URL_PATTERNS.search(source_url):
        return False, f"URL pattern suggests research paper, not CV: {source_url}"

    if source_url:
        url_filename = urlparse(source_url).path.split("/")[-1].lower()
        if re.search(r"(?:^|[a-z_-])(19|20)\d{2}\.pdf$", url_filename):
            return False, f"URL filename has year pattern (likely paper: {url_filename})"

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "\n".join(p.extract_text() or "" for p in reader.pages[:3])
    except Exception as e:
        return False, f"couldn't parse PDF: {e}"

    if not text.strip():
        return False, "no extractable text (likely a scanned image)"

    text_lower = text.lower()

    last_name = re.sub(r"[^\w-]", "", full_name.strip().split()[-1].lower())
    if last_name and last_name not in text_lower:
        return False, f"person's name ('{last_name}') not found in document"

    url_lower = source_url.lower() if source_url else ""
    URL_CV_LABEL_RE = re.compile(
        r"(?:[/_.\-]|(?<=[a-z]))(cv|resume|vitae|curriculum[_-]?vitae)(?=[/_.\-\d]|$)",
        re.IGNORECASE,
    )
    url_labeled_as_cv = bool(URL_CV_LABEL_RE.search(url_lower))

    keyword_hits = sum(1 for kw in CV_SECTION_KEYWORDS if kw in text_lower)

    paper_signals = len(PAPER_STRUCTURE_RE.findall(text_lower))

    first_500 = text.strip()[:500].lower()
    starts_with_abstract = bool(PAPER_SIGNAL_RE.match(first_500))

    if url_labeled_as_cv:
        if starts_with_abstract:
            return False, "URL says CV but document starts with 'Abstract' (likely a paper)"
        return True, f"URL labeled as CV ({keyword_hits} CV keywords, {paper_signals} paper signals)"

    if starts_with_abstract:
        return False, "document starts with 'Abstract' — likely a paper"

    timetable_hits = len(TIMETABLE_SIGNAL_RE.findall(text))
    if timetable_hits >= 3 and keyword_hits < 4:
        return False, "document looks like a timetable/syllabus, not a CV"

    if paper_signals >= 5 and paper_signals >= keyword_hits * 2:
        return False, f"too many paper signals ({paper_signals}) vs CV keywords ({keyword_hits})"

    if keyword_hits < 4:
        return False, (f"only {keyword_hits} CV keyword(s) found — "
                       f"real CVs have multiple sections (need ≥4)")

    return True, f"ok ({keyword_hits} CV keywords, {paper_signals} paper signals)"


def guess_cv_subpages(url: str) -> list[str]:
    """Guess common CV subpaths for a website."""
    CV_SUBPATHS = ["cv", "resume", "vitae", "curriculum-vitae", "curriculum_vitae",
                   "about/cv", "about/resume", "bio/cv"]

    base = url.rstrip("/")

    return [f"{base}/{sub}" for sub in CV_SUBPATHS]


def try_site_for_cv(url: str, full_name: str, depth: int = 0) -> tuple[bytes | None, str | None]:
    direct_candidates, subpages = find_pdf_or_drive_candidates(url)

    for candidate_url in direct_candidates:
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
        for subpage_url in subpages[:2]:
            content, source = try_site_for_cv(subpage_url, full_name, depth=1)
            if content:
                return content, source

        already_tried = set(subpages[:2])
        for guessed_url in guess_cv_subpages(url):
            if guessed_url in already_tried:
                continue

            print(f"    [guess] trying {guessed_url}")
            content, source = try_site_for_cv(guessed_url, full_name, depth=1)
            if content:
                return content, source

    return None, None


# ---------------------------------------------------------------------------
# Per-person pipeline (NEW SEARCH PRIORITY)
# ---------------------------------------------------------------------------

def tavily_search_personal_website(full_name: str, university: str, department: str) -> str | None:
    """Fallback search when department profile page can't be fetched."""
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return None
    last_name = re.sub(r"[^\w-]", "", parts[-1])
    last_name_lower = last_name.lower()

    query = f"{full_name} {university} {department}".strip()
    results = tavily_raw_search(query)

    print(f"    [tavily-pw] '{query}' -> checking for personal website or fallback profile")

    PROFILE_TAG_RE = re.compile(
        r"/(bio|profile|faculty|faculty-directory|people|user|homes|members|"
        r"person|display|directory)/([^/?#]+)|"
        r"/~([a-z0-9_-]+)",
        re.IGNORECASE,
    )

    all_urls = []
    scored_candidates = []  # (score, url, reason)

    for r in results:
        url = (r.get("url") or "").strip()
        title = (r.get("title") or "").lower()

        if not _tavily_url_usable(url):
            continue
        all_urls.append(url)

        if last_name_lower not in title:
            continue

        domain = urlparse(url).netloc.lower()
        path = urlparse(url).path
        clean_domain = domain.replace("www.", "").split(".")[0]

        score = 0
        reason = ""

        if last_name_lower and len(last_name_lower) >= 3 and last_name_lower in clean_domain:
            score = 30
            reason = f"domain contains last name '{last_name_lower}'"

        elif "github.io" in domain:
            score = 25
            reason = "github.io site"

        elif "sites.google.com" in domain and "/site/" in path:
            score = 25
            reason = "sites.google.com site"

        elif PROFILE_TAG_RE.search(path):
            tag_match = PROFILE_TAG_RE.search(path)
            segment = (tag_match.group(2) or tag_match.group(3) or "").strip()
            if segment and segment.lower() not in ["index", "home", "list", "s"]:
                score = 10
                reason = f"profile tag with segment '{segment}'"

        if score > 0:
            print(f"      ✓ [{reason}] {url}")
            scored_candidates.append((score, url, reason))
        else:
            print(f"      ✗ {url}")

    if scored_candidates:
        scored_candidates.sort(key=lambda x: -x[0])
        best_score, best_url, best_reason = scored_candidates[0]
        print(f"    [tavily-pw] Best candidate [score={best_score}]: {best_url}")
        return best_url

    if all_urls:
        print(f"    [tavily-pw] no personal website found among {len(all_urls)} result(s)")
    else:
        print(f"    [tavily-pw] no results")
    return None


# ---------------------------------------------------------------------------
# Per-person pipeline (ONE GENERAL TAVILY SEARCH + RANKED URL FLOW)
# ---------------------------------------------------------------------------

def tavily_general_search_and_rank(
    full_name: str,
    university: str,
    department: str,
    profile_url: str = "",
) -> list[tuple[int, str, str]]:
    """STEP 1 + STEP 2: one general Tavily search, then rank ALL usable URLs.

    Scores:
      30 = personal website (github.io, name.com, sites.google.com/site/)
      10 = department/faculty profile (/faculty/, /people/, /bio/, /~name)
      0  = junk / irrelevant URLs (skipped)

    Returns:
        [(score, url, kind), ...] sorted by score DESC.
    """
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return []

    first_name = re.sub(r"[^\w-]", "", parts[0]).lower()
    last_name = re.sub(r"[^\w-]", "", parts[-1]).lower()

    # Build site filter from profile_url if available (helps narrow to university)
    site_filter = ""
    if profile_url and profile_url.startswith("http"):
        try:
            domain = urlparse(profile_url).netloc
            if domain:
                root = get_root_domain(domain)
                site_filter = f" site:{root}"
        except Exception:
            pass

    # STEP 1: EXACTLY ONE Tavily search
    query = f"{full_name} {university} {department}".strip()
    results = tavily_raw_search(query)

    print(f"    [tavily] GENERAL search: {query!r}")
    print(f"    [tavily] returned {len(results)} result(s)")

    # Department URL signals — includes /content/ and /profiles/
    DEPT_PATH_RE = re.compile(
        r"/(?:faculty|faculty-directory|people|bio|profile|profiles|"
        r"members|directory|staff|researchers?|content|homes|user|"
        r"person|display)/[^/?#]+"
        r"|/~[a-z0-9_-]+(?:/|$)",
        re.IGNORECASE,
    )

    # JUNK: NEWS/BLOG/EVENTS — NOT include /research/ or /projects/ 
    # (faculty pages often have those in path, e.g. cs.yale.edu/research/faculty/xyz)
    JUNK_PATH_RE = re.compile(
        r"/(?:news|story|stories|blog|press|announcements?|articles?|"
        r"post|posts|events?|awards?|prize|honor|welcome|grantee|"
        r"faculty-qa|q-and-a|interview|spotlight|profile-story)(?:/|$)",
        re.IGNORECASE,
    )

    # Paper URLs (arxiv, ecva, ...)
    PAPER_JUNK_RE = re.compile(
        r"/(?:papers?|proceedings?|preprints?|publications?|abstract|pdf/\d{4})/",
        re.IGNORECASE,
    )

    JUNK_DOMAIN_SIGNALS = [
        "linkedin.com", "researchgate.net", "scholar.google",
        "facebook.com", "twitter.com", "x.com", "instagram.com",
        "wikipedia.org", "ratemyprofessors.com", "alphaxiv",
    ]

    scored_candidates: list[tuple[int, str, str]] = []
    seen: set[str] = set()

    for r in results:
        url = (r.get("url") or "").strip()
        title = (r.get("title") or "").strip().lower()

        if not _tavily_url_usable(url):
            continue

        normalized = url.lower().rstrip("/")
        if normalized in seen:
            continue
        seen.add(normalized)

        parsed = urlparse(url)
        domain = parsed.netloc.lower().replace("www.", "")
        path = parsed.path.lower()

        score = 0
        kind = "other"
        reason = ""

        # ----- JUNK check first -----
        if any(sig in domain for sig in JUNK_DOMAIN_SIGNALS):
            print(f"      [score=0] {url}  <- junk domain")
            continue
        if JUNK_PATH_RE.search(path) or PAPER_JUNK_RE.search(path):
            print(f"      [score=0] {url}  <- junk/paper path")
            continue

        # ----- Score 30: PERSONAL -----
        domain_first_part = domain.split(".")[0]
        name_in_domain = (
            len(last_name) >= 3 and last_name in domain_first_part
        )
        firstlast = f"{first_name}{last_name}"
        firstlast_in_domain = (
            len(firstlast) >= 5 and firstlast in domain_first_part
        )
        is_github = "github.io" in domain
        is_google_site = "sites.google.com" in domain and "/site/" in path

        if name_in_domain or firstlast_in_domain or is_github or is_google_site:
            score = 30
            kind = "personal"
            if name_in_domain:
                reason = f"domain has last name '{last_name}'"
            elif firstlast_in_domain:
                reason = f"domain has full name '{firstlast}'"
            elif is_github:
                reason = "github.io site"
            else:
                reason = "sites.google.com/site"

        # ----- Score 10: DEPARTMENT -----
        elif DEPT_PATH_RE.search(path):
            # For /~username, prefer URLs that contain the person's name
            tilde_match = re.search(r"/~([a-z0-9_-]+)", path, re.IGNORECASE)
            if tilde_match:
                username = tilde_match.group(1).lower()
                if (
                    last_name in username
                    or first_name in username
                    or (first_name and f"{first_name[0]}{last_name}" in username)
                ):
                    score = 10
                    kind = "department"
                    reason = f"dept ~{username}"
                else:
                    print(f"      [score=0] {url}  <- unrelated ~username '{username}'")
                    continue
            else:
                # Also check title contains last name (avoid wrong person same last name)
                if last_name and last_name not in title:
                    print(f"      [score=0] {url}  <- title lacks last name")
                    continue
                score = 10
                kind = "department"
                reason = "dept/faculty profile"

        else:
            print(f"      [score=0] {url}  <- no personal/dept signal")
            continue

        print(f"      [score={score}] {url}  <- {reason}")
        scored_candidates.append((score, url, kind))

    scored_candidates.sort(key=lambda item: item[0], reverse=True)

    print("    [tavily] RANKED candidates:")
    if scored_candidates:
        for score, url, kind in scored_candidates:
            print(f"      [{score:>2}] [{kind:<10}] {url}")
    else:
        print("      (none)")

    return scored_candidates


def _save_cv_result(content, source, name, university, department):
    """Save a verified CV and return standard result row."""
    os.makedirs(CV_FOLDER_ABS, exist_ok=True)
    dest = os.path.join(CV_FOLDER_ABS, f"{safe_filename(name)}.pdf")
    with open(dest, "wb") as f:
        f.write(content)
    print(f"  ✅ FOUND CV: {dest}")
    return {"university": university, "department": department, "name": name,
            "result_type": "cv_download", "value": dest, "source_url": source}


def _website_result(url, name, university, department):
    """Return standard website result row."""
    return {"university": university, "department": department, "name": name,
            "result_type": "website", "value": url, "source_url": url}


def process_person(record: dict) -> dict | None:
    """New single-search ranked pipeline.

    FLOW:
      STEP 1: 1 Tavily search
      STEP 2: rank URLs (30 personal / 10 department / 0 junk)
      STEP 3: try candidates DESC by score
        - 30 personal → check if it's a LAB/GROUP site first (team page ->
          confirm PI entry by name+title -> CV) → else try CV directly →
          return CV or PW
        - 10 dept → try CV → extract PW → (same lab-aware check) → try CV →
          return, else next
        - exhausted → best URL or DW fallback
    """
    name = (record.get("name") or "").strip()
    university = (record.get("university") or "").strip()
    department = (record.get("department") or "").strip()
    profile_url = (record.get("profile_url") or "").strip()

    if not name:
        return None

    # Detect invalid profile_url (non-URL text like "Yes (personal site listed)")
    # Empty/NA/N/A → OK to proceed without site: filter
    # Text that's NOT a URL and NOT empty → flag as manual_search
    if profile_url and profile_url.lower() not in EMPTY_VALUES:
        if not profile_url.startswith(("http://", "https://")):
            print(f"  -> profile_url is text (not URL): {profile_url!r} — flagging manual_search")
            return {"university": university, "department": department, "name": name,
                    "result_type": "manual_search",
                    "value": f"INVALID URL in input: {profile_url}",
                    "source_url": ""}
    
    # Normalize: if profile_url is NA/empty, treat as no profile URL
    # (Tavily search will still run, just without site: filter)
    if profile_url.lower() in EMPTY_VALUES:
        profile_url = ""
        print(f"  -> no profile URL in input, will search without site: filter")

    print(f"\n  ========== SEARCH PRIORITY ==========")
    print("  1️⃣  Tavily GENERAL search — exactly once")
    print("  2️⃣  Rank ALL URLs: 30 Personal / 10 Department / 0 Junk")
    print("  3️⃣  Try candidates from highest to lowest score")
    print("      (personal/website candidates are checked for lab/group")
    print("       sites first -- if so, navigate Team/People -> PI -> CV)")
    print("  =====================================\n")

    # STEP 1 + 2
    candidates = tavily_general_search_and_rank(name, university, department, profile_url)

    if not candidates:
        print("  ⚠️  No usable candidates from Tavily")
        if not is_empty(profile_url):
            print(f"  ✅ DW FALLBACK: {profile_url}")
            return _website_result(profile_url, name, university, department)
        print("  ⚠️  No URL available → flagging manual_search")
        return {"university": university, "department": department, "name": name,
                "result_type": "manual_search",
                "value": "MANUAL SEARCH NEEDED (no profile URL, Tavily returned no candidates)",
                "source_url": ""}

    best_url = candidates[0][1]
    visited_urls: set[str] = set()

    # STEP 3: try candidates
    for score, candidate_url, kind in candidates:
        normalized = candidate_url.lower().rstrip("/")
        if normalized in visited_urls:
            print(f"  [skip] already tried: {candidate_url}")
            continue
        visited_urls.add(normalized)

        print(f"\n  [candidate] score={score}, kind={kind}")
        print(f"  [candidate] {candidate_url}")

        # ===== SCORE 30 — PERSONAL =====
        if score == 30 and kind == "personal":
            print("  [30] Personal-scored site → checking if it's actually a lab/group site")
            content, source = try_lab_site_for_cv(candidate_url, name)
            if content:
                return _save_cv_result(content, source, name, university, department)

            print("  [30] Not a lab site (or lab flow found nothing) → trying CV directly")
            content, source = try_site_for_cv(candidate_url, name)
            if content:
                return _save_cv_result(content, source, name, university, department)
            # MAX score exhausted → return PW
            print("  [30] No CV → returning personal website")
            return _website_result(candidate_url, name, university, department)

        # ===== SCORE 10 — DEPARTMENT =====
        if score == 10 and kind == "department":
            print("  [10] Department profile → trying CV")
            content, source = try_site_for_cv(candidate_url, name)
            if content:
                return _save_cv_result(content, source, name, university, department)

            print("  [10] No CV on dept → extracting Website link")
            website_links = find_website_links_on_profile(candidate_url, name)

            if website_links:
                for website_url in website_links:
                    pw_normalized = website_url.lower().rstrip("/")
                    if pw_normalized in visited_urls:
                        print(f"  [10] Skip already-tried website: {website_url}")
                        continue
                    visited_urls.add(pw_normalized)

                    print(f"  [10] Checking if linked website is a lab/group site: {website_url}")
                    content, source = try_lab_site_for_cv(website_url, name)
                    if content:
                        return _save_cv_result(content, source, name, university, department)

                    print(f"  [10] Trying CV on personal website: {website_url}")
                    content, source = try_site_for_cv(website_url, name)
                    if content:
                        return _save_cv_result(content, source, name, university, department)

                    # Found website link but no CV → return PW
                    print("  [10] No CV on PW → returning PW")
                    return _website_result(website_url, name, university, department)

            # No Website link → continue to next candidate
            print("  [10] No Website link → trying next candidate")
            continue

    # ALL EXHAUSTED
    print(f"\n  ⚠️  All candidates exhausted → best URL: {best_url}")
    return _website_result(best_url, name, university, department)


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


def write_new_tab(rows: list[dict], sheet_name: str) -> None:
    """Writes `rows` into `sheet_name` tab of the shared RESULTS_XLSX workbook.
    Re-opens/saves the workbook each call so multiple batches in the same run
    all land in the same file, each in its own tab."""
    results_path = resolve_input_path(RESULTS_XLSX)
    if results_path is not None:
        wb = load_workbook(results_path)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        results_path = os.path.join(script_dir, RESULTS_XLSX)
        from openpyxl import Workbook
        wb = Workbook()
        wb.remove(wb.active)

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

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
    print(f"\n[saved] '{sheet_name}' tab written to: {os.path.abspath(results_path)}")


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


def process_batch(batch: dict) -> dict:
    """Runs the full pipeline for ONE department batch: loads its input xlsx,
    processes every person, writes results into its own sheet, downloads CVs
    into its own folder. Returns a small summary dict for the final report."""
    global CV_FOLDER, CV_FOLDER_ABS

    input_xlsx = batch["input_xlsx"]
    new_sheet_name = batch["new_sheet_name"]
    cv_folder = batch["cv_folder"]

    # Point the module-level CV folder globals at THIS batch's folder so that
    # process_person()/try_site_for_cv() (which read CV_FOLDER_ABS) save CVs
    # into the right place for this department.
    CV_FOLDER = cv_folder
    CV_FOLDER_ABS = os.path.join(_SCRIPT_DIR, CV_FOLDER)

    print(f"\n\n{'#'*70}")
    print(f"#  BATCH: {new_sheet_name}  (input: {input_xlsx}, CV folder: {cv_folder})")
    print(f"{'#'*70}")

    input_path = resolve_input_path(input_xlsx)
    if input_path is None:
        print(f"Couldn't find {input_xlsx!r} in cwd or script folder -- skipping this batch.")
        return {"sheet": new_sheet_name, "downloaded": 0, "websites": 0, "manuals": 0,
                "skipped": True}

    people = load_people(input_path)
    print(f"\nLoaded {len(people)} people for {new_sheet_name}\n")

    rows = []
    for i, record in enumerate(people, 1):
        print(f"\n{'='*60}")
        print(f"[{new_sheet_name}] Person {i}/{len(people)}: {record.get('name')}")
        print(f"{'='*60}")
        row = process_person(record)
        if row:
            rows.append(row)

    write_new_tab(rows, new_sheet_name)

    downloaded = sum(1 for r in rows if r["result_type"] == "cv_download")
    websites = sum(1 for r in rows if r["result_type"] == "website")
    manuals = sum(1 for r in rows if r["result_type"] == "manual_search")

    print(f"\n{'-'*60}")
    print(f"BATCH DONE: {new_sheet_name}")
    print(f"  CVs downloaded: {downloaded}")
    print(f"  Websites recorded: {websites}")
    print(f"  Manual search needed: {manuals}")
    print(f"  CV folder: {CV_FOLDER_ABS}")
    print(f"{'-'*60}")

    return {"sheet": new_sheet_name, "downloaded": downloaded, "websites": websites,
            "manuals": manuals, "skipped": False}


def main():
    print(f"[cwd] working directory: {os.getcwd()}")
    print(f"[cwd] script location: {os.path.dirname(os.path.abspath(__file__))}")
    print(f"[batches] running {len(BATCHES)} batch(es): "
          f"{', '.join(b['new_sheet_name'] for b in BATCHES)}")

    summaries = [process_batch(batch) for batch in BATCHES]

    print(f"\n\n{'='*70}")
    print(f"ALL BATCHES DONE ✅")
    print(f"{'='*70}")
    total_downloaded = total_websites = total_manuals = 0
    for s in summaries:
        if s["skipped"]:
            print(f"  {s['sheet']}: SKIPPED (input file not found)")
            continue
        print(f"  {s['sheet']}: {s['downloaded']} CVs, {s['websites']} websites, "
              f"{s['manuals']} manual")
        total_downloaded += s["downloaded"]
        total_websites += s["websites"]
        total_manuals += s["manuals"]
    print(f"{'-'*70}")
    print(f"TOTAL: {total_downloaded} CVs downloaded, {total_websites} websites, "
          f"{total_manuals} manual search")
    print(f"Results file: {RESULTS_XLSX} (one tab per batch)")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[!!!] CRASHED -- full error below:")
        traceback.print_exc()
        sys.exit(1)