"""
Fetch OpenAlex impact scores (h-index) for assistant professors listed in
several xlsx files, and write the results into a new "research info" sheet
inside each of those same files.

Requirements:
    pip install openpyxl requests

Usage:
    Put this script in the same folder as your 5 xlsx files and run:
        python fetch_impact_scores_fixed.py

Notes:
- Set EMAIL below to your own email address. OpenAlex gives faster, more
  reliable service to requests that include a contact email ("polite pool").
- Author name matching is done by searching OpenAlex's /authors endpoint
  and picking the best match based on similarity between the candidate's
  last known institution and the "University" column in your sheet.
- "Impact score" = h-index from OpenAlex's summary_stats.
"""

import os
import time
import sys
from difflib import SequenceMatcher

import openpyxl
import requests

# ---------------------------------------------------------------------------
# CONFIG — edit these to match your setup
# ---------------------------------------------------------------------------

EMAIL = "hoangndsy7@gmail.com"

# IMPORTANT: Get your free API key from https://openalex.org/settings/api
# (Required as of Feb 2026 to avoid rate limits)
# Leave blank "" to try without a key (limited to ~100 requests)
API_KEY = "vQ3TtNToOv4hIRpkWsGTOd"  # <-- PASTE YOUR KEY HERE, e.g. API_KEY = "abc123xyz"

FOLDER = os.path.dirname(os.path.abspath(__file__))

FILENAMES = [
    "Math.xlsx",

]

FILES = [os.path.join(FOLDER, name) for name in FILENAMES]

# Column names expected in the SOURCE sheet (first sheet of each file)
COL_NAME = "Name"
COL_UNIVERSITY = "University"
COL_DEPARTMENT = "Department"
COL_TITLE = "Title"

OUTPUT_SHEET_NAME = "research info"

# "h_index", "i10_index", or "cited_by_count"
IMPACT_METRIC = "h_index"

# Minimum institution-name similarity (0-1) to accept a match
MIN_MATCH_SCORE = 0.4

REQUEST_DELAY_SECONDS = 0.5  # increased to be extra polite

BASE_URL = "https://api.openalex.org/authors"

# Retry settings
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds

# ---------------------------------------------------------------------------


def log(msg):
    """Print with timestamp for debugging."""
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def test_openalex_connection():
    """Test if we can reach OpenAlex API."""
    log("Testing OpenAlex connection...")
    try:
        headers = {
            "User-Agent": f"researcher-dataset-script (mailto:{EMAIL})",
        }
        params = {
            "search": "Albert Einstein",
            "per_page": 1,
            "mailto": EMAIL,
        }
        if API_KEY:
            params["api_key"] = API_KEY
            
        resp = requests.get(BASE_URL, params=params, headers=headers, timeout=15)
        log(f"  Status Code: {resp.status_code}")
        
        if resp.status_code != 200:
            log(f"  ERROR: Got status code {resp.status_code}")
            log(f"  Response: {resp.text[:200]}")
            return False
        
        data = resp.json()
        if "results" in data:
            log(f"  ✓ Connection OK. Found {len(data['results'])} results for 'Albert Einstein'")
            
            # Check for rate limit info in headers
            if 'X-RateLimit-Remaining' in resp.headers:
                remaining = resp.headers.get('X-RateLimit-Remaining', 'unknown')
                log(f"  ✓ API Key active. Remaining credits this hour: {remaining}")
            
            return True
        else:
            log(f"  ERROR: Unexpected response format")
            log(f"  Response keys: {list(data.keys())}")
            return False
            
    except requests.exceptions.Timeout:
        log("  ERROR: Request timeout. OpenAlex might be slow.")
        return False
    except requests.exceptions.ConnectionError as e:
        log(f"  ERROR: Cannot connect to OpenAlex: {e}")
        return False
    except Exception as e:
        log(f"  ERROR: {type(e).__name__}: {e}")
        return False


def similar(a: str, b: str) -> float:
    """Calculate string similarity ratio."""
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def find_author(name: str, university: str, attempt=1):
    """
    Search OpenAlex for an author and return (best_author, match_score, debug_note).
    Includes retry logic for transient failures.
    """
    params = {
        "search": name,
        "per_page": 25,
        "mailto": EMAIL,
    }
    if API_KEY:
        params["api_key"] = API_KEY
    
    headers = {
        "User-Agent": f"researcher-dataset-script (mailto:{EMAIL})",
    }
    
    try:
        resp = requests.get(BASE_URL, params=params, headers=headers, timeout=20)
        resp.raise_for_status()
        
    except requests.exceptions.Timeout:
        if attempt < MAX_RETRIES:
            log(f"    Timeout on attempt {attempt}, retrying...")
            time.sleep(RETRY_DELAY)
            return find_author(name, university, attempt + 1)
        return None, 0.0, f"TIMEOUT after {MAX_RETRIES} attempts"
        
    except requests.exceptions.ConnectionError as e:
        if attempt < MAX_RETRIES:
            log(f"    Connection error on attempt {attempt}, retrying...")
            time.sleep(RETRY_DELAY)
            return find_author(name, university, attempt + 1)
        return None, 0.0, f"CONNECTION ERROR: {str(e)[:50]}"
        
    except requests.exceptions.HTTPError as e:
        return None, 0.0, f"HTTP {resp.status_code}: {resp.text[:100]}"
        
    except Exception as e:
        return None, 0.0, f"ERROR: {type(e).__name__}: {str(e)[:50]}"

    try:
        results = resp.json().get("results", [])
    except ValueError as e:
        return None, 0.0, f"JSON PARSE ERROR: {str(e)[:50]}"

    if not results:
        return None, 0.0, "no OpenAlex results for this name"

    best_author = None
    best_score = 0.0
    
    for author in results:
        # Try to get institution name from various fields
        inst_name = ""
        
        # Check last_known_institutions (newer format, list)
        if author.get("last_known_institutions"):
            insts = author.get("last_known_institutions", [])
            if insts and isinstance(insts, list) and len(insts) > 0:
                inst_name = insts[0].get("display_name", "") or ""
        
        # Fall back to last_known_institution (older format, dict)
        if not inst_name and author.get("last_known_institution"):
            inst = author.get("last_known_institution", {})
            inst_name = inst.get("display_name", "") or ""

        # Calculate match score
        score = similar(inst_name, university)
        
        if score > best_score:
            best_score = score
            best_author = author

    if best_author is None:
        return None, 0.0, "no author with institution data found"

    return best_author, best_score, ""


def get_impact_score(author: dict):
    """Extract the impact metric from author data."""
    if not author:
        return "NA"
    
    # For cited_by_count, get it directly from author
    if IMPACT_METRIC == "cited_by_count":
        score = author.get("cited_by_count")
        return score if score is not None else "NA"
    
    # For h_index and i10_index, get from summary_stats
    stats = author.get("summary_stats") or {}
    score = stats.get(IMPACT_METRIC)
    return score if score is not None else "NA"


def process_file(filepath: str):
    """Process a single Excel file."""
    log(f"\nProcessing {os.path.basename(filepath)}...")
    
    try:
        wb = openpyxl.load_workbook(filepath)
    except FileNotFoundError:
        log(f"  [SKIPPED] File not found: {filepath}")
        return False
    except Exception as e:
        log(f"  [SKIPPED] Error loading workbook: {e}")
        return False

    src_sheet = wb.worksheets[0]

    # Get headers
    headers = [cell.value for cell in src_sheet[1]]
    
    # Verify required columns exist
    try:
        col_idx = {
            COL_NAME: headers.index(COL_NAME),
            COL_UNIVERSITY: headers.index(COL_UNIVERSITY),
            COL_DEPARTMENT: headers.index(COL_DEPARTMENT),
            COL_TITLE: headers.index(COL_TITLE),
        }
    except ValueError as exc:
        log(f"  [SKIPPED] Missing required column: {exc}")
        log(f"  Available columns: {headers}")
        return False

    rows = list(src_sheet.iter_rows(min_row=2, values_only=True))

    # Remove old research info sheet if it exists
    if OUTPUT_SHEET_NAME in wb.sheetnames:
        del wb[OUTPUT_SHEET_NAME]
    
    out_sheet = wb.create_sheet(OUTPUT_SHEET_NAME)
    out_sheet.append(
        ["Department", "Name", "Title", "School", "H-Index", "Match Confidence", "Note"]
    )

    total = len(rows)
    success_count = 0
    not_found_count = 0
    
    for i, row in enumerate(rows, start=1):
        if not row or all(v is None for v in row):
            continue
            
        name = row[col_idx[COL_NAME]]
        university = row[col_idx[COL_UNIVERSITY]]
        department = row[col_idx[COL_DEPARTMENT]]
        title = row[col_idx[COL_TITLE]]

        if not name:
            continue

        log(f"  ({i}/{total}) {name} @ {university or '(no university)'}")
        
        author, match_score, note = find_author(
            str(name).strip(), 
            str(university).strip() if university else ""
        )

        if author:
            score = get_impact_score(author)
            confidence = "low" if match_score < MIN_MATCH_SCORE else "ok"
            success_count += 1
        else:
            score = "NOT FOUND"
            confidence = ""
            not_found_count += 1

        out_sheet.append([department, name, title, university, score, confidence, note])
        time.sleep(REQUEST_DELAY_SECONDS)

    # Save the workbook
    try:
        wb.save(filepath)
        log(f"  ✓ Saved to {os.path.basename(filepath)}")
        log(f"    Success: {success_count}, Not found: {not_found_count}")
        return True
    except Exception as e:
        log(f"  [ERROR] Could not save file: {e}")
        return False


def main():
    log("=" * 60)
    log("OpenAlex H-Index Fetcher")
    log("=" * 60)
    
    # Warn if no API key
    if not API_KEY:
        log("\n⚠️  WARNING: No API key found!")
        log("   You'll be limited to ~100 searches before hitting rate limits.")
        log("   Get your FREE API key: https://openalex.org/settings/api")
        log("   Then add it to the script: API_KEY = 'your-key-here'")
        log("")
    else:
        log(f"\n✓ API Key detected (first 8 chars: {API_KEY[:8]}...)")
        log("  You have $1 free daily credit (~1,000 searches)")
        log("")
    
    # Test connection first
    if not test_openalex_connection():
        log("\n[CRITICAL] Cannot connect to OpenAlex. Aborting.")
        log("Possible fixes:")
        log("  1. Check your internet connection")
        log("  2. Check if OpenAlex is down (https://openalex.org)")
        log("  3. Try again in a few minutes")
        sys.exit(1)
    
    log(f"\nProcessing {len(FILES)} files...")
    
    successful_files = 0
    for filepath in FILES:
        if process_file(filepath):
            successful_files += 1
    
    log("\n" + "=" * 60)
    log(f"Completed: {successful_files}/{len(FILES)} files processed")
    log("=" * 60)


if __name__ == "__main__":
    main()