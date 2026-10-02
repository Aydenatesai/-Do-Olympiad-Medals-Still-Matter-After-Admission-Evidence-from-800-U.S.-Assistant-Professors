"""
AI CV Auditor -- Math + Computer Science (folder-scanning version)
======================================================================

Scans the actual CV folders on disk (Math_CV, Computer Science) for every
.pdf file that's really there, tries to match each one back to a row in
your results workbook for context (university/department/source_url), and
asks Claude to judge whether it's genuinely that person's own CV.

Why scan the folder instead of trusting the spreadsheet:
the spreadsheet's "value" column has been wrong before (scrambled order),
so trusting it as the list of "what exists" can miss files that are really
there, or try to open files that no longer exist. Scanning the folder is
the ground truth; the spreadsheet is only used afterward, to add context.

For each PDF found it reports one of:
  CORRECT        -- Claude confirms this is that person's own CV
  WRONG          -- Claude says it's the wrong document, with a short reason
  UNREADABLE     -- no extractable text (e.g. scanned image), can't judge
  NOT_IN_SHEET   -- the PDF exists on disk but no matching row was found in
                    either results tab (still AI-checked, just missing
                    university/department context)

RATE LIMITS: the Anthropic API enforces per-minute request/token limits
based on your usage tier. This script retries automatically with
exponential backoff on 429 (rate limit) errors instead of crashing, and
saves its progress after every single file -- so if you do hit a limit,
close the terminal, wait, and rerun, and it picks up from the audit tab
without losing anything already checked.

SETUP:
  pip install openpyxl pypdf anthropic

USAGE:
  1. Set ANTHROPIC_API_KEY below.
  2. Check FOLDERS_TO_CHECK and SHEETS_TO_CHECK match your actual names.
  3. python audit_cvs_ai.py
  4. Open RESULTS_XLSX -> "AI CV Audit" tab.
"""

from __future__ import annotations

import os
import re
import time
import unicodedata

from openpyxl import load_workbook
from openpyxl.styles import Font
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

ANTHROPIC_API_KEY = "PUT_YOUR_ANTHROPIC_API_KEY_HERE"
AI_MODEL = "claude-haiku-4-5-20251001"  # fast + cheap, plenty for this task

# --- AI backend ---------------------------------------------------------
# "ollama"     -- FREE, runs entirely on your own machine via Ollama.
#                 Install: https://ollama.com , then run once:
#                     ollama pull llama3.2
#                 No API key, no cost, no rate-limit paywall. Slightly less
#                 sharp than Claude on nuanced calls, but fine for a
#                 straightforward "is this really their CV" yes/no check.
# "anthropic"  -- uses the paid Claude API (ANTHROPIC_API_KEY above).
AI_BACKEND = "ollama"
OLLAMA_MODEL = "llama3.2"
OLLAMA_ENDPOINT = "http://localhost:11434/api/generate"

RESULTS_XLSX = "cv_or_website_results.xlsx"
SHEETS_TO_CHECK = ["CV or Website - Physics"]
FOLDERS_TO_CHECK = ["Physics","Math"]
AUDIT_SHEET_NAME = "AI CV Audit-Physics"

COL_UNIVERSITY, COL_DEPARTMENT, COL_NAME, COL_RESULT_TYPE, COL_VALUE, COL_SOURCE = range(1, 7)

MAX_CHARS_TO_SEND = 6000
DELAY_BETWEEN_CALLS_SECONDS = 1.0   # be gentle on rate limits
MAX_RETRIES = 5
INITIAL_BACKOFF_SECONDS = 5


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_filename(name: str) -> str:
    normalized = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    cleaned = re.sub(r"[^\w\s-]", "", normalized).strip().replace(" ", "_")
    return cleaned or "unknown"


def guess_name_from_filename(stem: str) -> str:
    # strip a trailing _2 / _3 collision suffix, then turn underscores into spaces
    stem = re.sub(r"_(\d+)$", "", stem)
    return stem.replace("_", " ").strip()


def build_sheet_index(wb) -> dict[str, dict]:
    """Maps safe_filename(name) -> {university, department, source_url, sheet}
    for every cv_download row across all sheets, so a file on disk can be
    matched back to its context."""
    index = {}
    for sheet_name in SHEETS_TO_CHECK:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in range(2, ws.max_row + 1):
            result_type = ws.cell(row=row, column=COL_RESULT_TYPE).value
            name = (ws.cell(row=row, column=COL_NAME).value or "").strip()
            if result_type != "cv_download" or not name:
                continue
            key = safe_filename(name)
            index[key] = {
                "university": ws.cell(row=row, column=COL_UNIVERSITY).value or "",
                "department": ws.cell(row=row, column=COL_DEPARTMENT).value or "",
                "source_url": ws.cell(row=row, column=COL_SOURCE).value or "",
                "sheet": sheet_name,
                "sheet_name_field": name,
            }
    return index


def extract_pdf_text(path: str) -> str | None:
    try:
        with open(path, "rb") as f:
            reader = PdfReader(f)
            text = "\n".join((p.extract_text() or "") for p in reader.pages[:3])
        return text
    except Exception as e:
        print(f"    [!] couldn't read {path}: {e}")
        return None


def build_prompt(text: str, full_name: str) -> str:
    return f"""Here is text extracted from the first few pages of a PDF document.

The document is supposed to be the CV (curriculum vitae / resume) of a person named "{full_name}".

Document text:
---
{text[:MAX_CHARS_TO_SEND]}
---

Is this document genuinely a CV/resume/curriculum vitae belonging to "{full_name}" specifically?

Answer in exactly this format, two lines:
VERDICT: YES or NO
REASON: a short (under 15 words) reason -- if NO, say what it actually looks like (e.g. "research paper by someone else", "class timetable", "CV of a different person named X")"""


def parse_verdict_reply(reply: str) -> tuple[str, str]:
    verdict_line = next((l for l in reply.splitlines() if l.upper().startswith("VERDICT")), "")
    reason_line = next((l for l in reply.splitlines() if l.upper().startswith("REASON")), "")
    reason = reason_line.split(":", 1)[-1].strip() if reason_line else reply[:100]
    if "YES" in verdict_line.upper():
        return "CORRECT", reason or "confirmed"
    return "WRONG", reason or "AI judged this is not the person's own CV"


def ask_ollama_is_this_their_cv(text: str, full_name: str) -> tuple[str, str]:
    """Free, local alternative to the Claude API call -- requires Ollama
    running on your machine (https://ollama.com, then `ollama pull llama3.2`).
    No retries needed here since there's no external rate limit -- the only
    failure mode is Ollama not being installed/running."""
    import requests as _requests

    prompt = build_prompt(text, full_name)
    try:
        resp = _requests.post(
            OLLAMA_ENDPOINT,
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=60,
        )
        resp.raise_for_status()
        reply = resp.json().get("response", "").strip()
        if not reply:
            return "WRONG", "Ollama returned an empty response"
        return parse_verdict_reply(reply)
    except _requests.exceptions.ConnectionError:
        return "WRONG", "couldn't reach Ollama -- is it running? (ollama serve / open the Ollama app)"
    except Exception as e:
        return "WRONG", f"Ollama call failed: {e}"


def ask_ai_is_this_their_cv(client, text: str, full_name: str) -> tuple[str, str]:
    """Returns (verdict, reason). Dispatches to whichever backend is
    configured. Retries with backoff on rate-limit errors instead of
    crashing the whole run (only relevant for the paid Anthropic backend)."""
    if AI_BACKEND == "ollama":
        return ask_ollama_is_this_their_cv(text, full_name)

    prompt = build_prompt(text, full_name)

    backoff = INITIAL_BACKOFF_SECONDS
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.messages.create(
                model=AI_MODEL,
                max_tokens=60,
                messages=[{"role": "user", "content": prompt}],
            )
            reply = response.content[0].text.strip()
            return parse_verdict_reply(reply)

        except Exception as e:
            status = getattr(e, "status_code", None)
            is_rate_limit = status == 429 or "rate_limit" in str(e).lower() or "429" in str(e)
            if is_rate_limit and attempt < MAX_RETRIES:
                print(f"    [!] rate limited, waiting {backoff}s (attempt {attempt}/{MAX_RETRIES})...")
                time.sleep(backoff)
                backoff *= 2
                continue
            return "WRONG", f"AI call failed: {e}"

    return "WRONG", "AI call failed after retries"


def save_audit(wb, audit_rows: list[list]) -> None:
    if AUDIT_SHEET_NAME in wb.sheetnames:
        del wb[AUDIT_SHEET_NAME]
    audit_ws = wb.create_sheet(AUDIT_SHEET_NAME)

    headers = ["sheet", "university", "department", "name", "status", "reason", "local_path", "source_url"]
    audit_ws.append(headers)
    for cell in audit_ws[1]:
        cell.font = Font(bold=True, name="Arial")

    status_order = {"WRONG": 0, "NOT_IN_SHEET": 1, "UNREADABLE": 2, "CORRECT": 3}
    rows_sorted = sorted(audit_rows, key=lambda r: status_order.get(r[4], 9))

    for r in rows_sorted:
        audit_ws.append(r)
    for row in audit_ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial")

    widths = [22, 26, 18, 24, 14, 45, 40, 45]
    for i, w in enumerate(widths, start=1):
        audit_ws.column_dimensions[audit_ws.cell(row=1, column=i).column_letter].width = w

    wb.save(RESULTS_XLSX)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    if not os.path.exists(RESULTS_XLSX):
        print(f"Couldn't find {RESULTS_XLSX} -- put it next to this script.")
        return

    client = None
    if AI_BACKEND == "anthropic":
        try:
            import anthropic
        except ImportError:
            print("Missing dependency: pip install anthropic")
            return
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    elif AI_BACKEND == "ollama":
        print(f"Using local Ollama model '{OLLAMA_MODEL}' at {OLLAMA_ENDPOINT}")
        print("(make sure the Ollama app/service is running, and you've run: ollama pull " + OLLAMA_MODEL + ")")
    else:
        print(f"Unknown AI_BACKEND: {AI_BACKEND!r} -- use 'ollama' or 'anthropic'")
        return

    wb = load_workbook(RESULTS_XLSX)
    sheet_index = build_sheet_index(wb)

    audit_rows = []
    counts = {"CORRECT": 0, "WRONG": 0, "UNREADABLE": 0, "NOT_IN_SHEET": 0}

    for folder in FOLDERS_TO_CHECK:
        if not os.path.isdir(folder):
            print(f"Folder '{folder}' not found, skipping.")
            continue
        print(f"\n=== Scanning folder: {folder} ===")

        pdf_files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        print(f"Found {len(pdf_files)} PDF(s)")

        for filename in pdf_files:
            local_path = os.path.join(folder, filename)
            stem = os.path.splitext(filename)[0]
            match = sheet_index.get(stem)

            if match:
                display_name = match["sheet_name_field"]
                university, department, source_url, sheet_label = (
                    match["university"], match["department"], match["source_url"], match["sheet"]
                )
                not_in_sheet = False
            else:
                display_name = guess_name_from_filename(stem)
                university = department = source_url = ""
                sheet_label = folder
                not_in_sheet = True

            print(f"Checking {display_name} ({local_path})")

            text = extract_pdf_text(local_path)
            if not text or not text.strip():
                print("    -> unreadable (no extractable text)")
                status = "UNREADABLE"
                reason = "No extractable text (likely scanned image)"
            else:
                status, reason = ask_ai_is_this_their_cv(client, text, display_name)
                if status == "CORRECT" and not_in_sheet:
                    status = "NOT_IN_SHEET"
                    reason = f"file exists but no matching sheet row found ({reason})"
                print(f"    -> {status}: {reason}")

            counts[status] += 1
            audit_rows.append([sheet_label, university, department, display_name, status, reason, local_path, source_url])

            # save after every file so a crash / rate-limit stall never loses progress
            save_audit(wb, audit_rows)
            time.sleep(DELAY_BETWEEN_CALLS_SECONDS)

    print(f"\n=== Done ===")
    print(f"Correct: {counts['CORRECT']}  Wrong: {counts['WRONG']}  "
    f"Unreadable: {counts['UNREADABLE']}  Not in sheet: {counts['NOT_IN_SHEET']}")
    print(f"Full results written to the '{AUDIT_SHEET_NAME}' tab in {RESULTS_XLSX}.")


if __name__ == "__main__":
    main()