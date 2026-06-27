#!/usr/bin/env python3
"""
Congruence Analysis Script
Drop PDFs into the 'documents/' folder, then run this script.
Results are saved to 'results/' as CSV and Excel files.
"""

import os
import re
import sys
import csv
from datetime import datetime
from pathlib import Path

# ── dependency check ──────────────────────────────────────────────────────────
missing = []
try:
    import pdfplumber
except ImportError:
    missing.append("pdfplumber")
try:
    import pandas as pd
except ImportError:
    missing.append("pandas")
try:
    import openpyxl
except ImportError:
    missing.append("openpyxl")

if missing:
    print(f"\n[ERROR] Missing packages: {', '.join(missing)}")
    print("Run setup.sh first (or: pip3 install " + " ".join(missing) + ")\n")
    sys.exit(1)

# ── paths ─────────────────────────────────────────────────────────────────────
BASE = Path(__file__).parent
DOCS_DIR = BASE / "documents"
RESULTS_DIR = BASE / "results"
KEYWORDS_FILE = BASE / "keywords.txt"

DOCS_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(exist_ok=True)


def load_keywords() -> list[str]:
    """Load keywords from keywords.txt, ignoring comment lines."""
    if not KEYWORDS_FILE.exists():
        return []
    with open(KEYWORDS_FILE, encoding="utf-8") as f:
        return [
            line.strip()
            for line in f
            if line.strip() and not line.startswith("#") and not line.startswith("**")
        ]


def prompt_extra_keywords(base: list[str]) -> list[str]:
    """Let the user add extra keywords for this run."""
    print("\n── Keyword Setup ──────────────────────────────────────")
    print(f"Default keywords loaded ({len(base)}):")
    for kw in base:
        print(f"  • {kw}")
    print()
    raw = input("Add extra keywords for this run (comma-separated, or press Enter to skip): ").strip()
    extra = [k.strip() for k in raw.split(",") if k.strip()] if raw else []
    all_kw = base + extra
    # deduplicate, preserve order, case-insensitive
    seen = set()
    result = []
    for kw in all_kw:
        key = kw.lower()
        if key not in seen:
            seen.add(key)
            result.append(kw)
    print(f"\nSearching for {len(result)} keyword(s): {', '.join(result)}")
    print("───────────────────────────────────────────────────────\n")
    return result


def extract_pages(pdf_path: Path) -> list[tuple[int, str]]:
    """Return list of (page_number, text) for each page in the PDF."""
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            pages.append((i, text))
    return pages


def get_context(text: str, match_start: int, match_end: int, window: int = 200) -> str:
    """Return the sentence or ~window chars surrounding a keyword match."""
    # try to grab the full sentence
    chunk_start = max(0, match_start - window)
    chunk_end = min(len(text), match_end + window)
    snippet = text[chunk_start:chunk_end]

    # find sentence boundaries within the snippet
    keyword_offset = match_start - chunk_start
    sentence_start = max(snippet.rfind(". ", 0, keyword_offset), snippet.rfind("\n", 0, keyword_offset))
    sentence_end = snippet.find(". ", keyword_offset)

    if sentence_start == -1:
        sentence_start = 0
    else:
        sentence_start += 2  # skip the ". "

    if sentence_end == -1:
        sentence_end = len(snippet)
    else:
        sentence_end += 1  # include the period

    context = snippet[sentence_start:sentence_end].strip()
    # collapse internal whitespace/newlines
    context = re.sub(r"\s+", " ", context)
    return context


def search_document(pdf_path: Path, keywords: list[str]) -> list[dict]:
    """Search a PDF for all keywords and return a list of match records."""
    print(f"  Analyzing: {pdf_path.name}")
    rows = []
    try:
        pages = extract_pages(pdf_path)
    except Exception as e:
        print(f"  [WARN] Could not read {pdf_path.name}: {e}")
        return rows

    for keyword in keywords:
        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
        for page_num, text in pages:
            for match in pattern.finditer(text):
                context = get_context(text, match.start(), match.end())
                rows.append({
                    "Document": pdf_path.name,
                    "Keyword": keyword,
                    "Page": page_num,
                    "Matched Text": match.group(),
                    "Context": context,
                })
    return rows


def save_results(all_rows: list[dict], keywords: list[str]) -> Path:
    """Save results to Excel (with a summary sheet) and return the file path."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = RESULTS_DIR / f"analysis_{timestamp}.xlsx"

    df = pd.DataFrame(all_rows)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # ── Matches sheet ──────────────────────────────────────────────────
        if df.empty:
            df_out = pd.DataFrame(columns=["Document", "Keyword", "Page", "Matched Text", "Context"])
        else:
            df_out = df[["Document", "Keyword", "Page", "Matched Text", "Context"]]
        df_out.to_excel(writer, sheet_name="Matches", index=False)

        # ── Summary sheet ──────────────────────────────────────────────────
        if not df.empty:
            summary = (
                df.groupby(["Document", "Keyword"])
                .size()
                .reset_index(name="Count")
                .sort_values(["Document", "Count"], ascending=[True, False])
            )
        else:
            summary = pd.DataFrame(columns=["Document", "Keyword", "Count"])
        summary.to_excel(writer, sheet_name="Summary", index=False)

        # ── auto-fit column widths ─────────────────────────────────────────
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)

    return out_path


def main():
    print("\n══════════════════════════════════════════════════════")
    print("          Congruence Analysis — Thesis Tool")
    print("══════════════════════════════════════════════════════")

    # find PDFs
    pdfs = sorted(DOCS_DIR.glob("*.pdf"))
    if not pdfs:
        print(f"\n[INFO] No PDFs found in '{DOCS_DIR}'.")
        print("Drop your PDF files into the 'documents/' folder and re-run.\n")
        sys.exit(0)

    print(f"\nFound {len(pdfs)} document(s):")
    for p in pdfs:
        print(f"  • {p.name}")

    # keywords
    base_keywords = load_keywords()
    if not base_keywords:
        print("\n[WARN] keywords.txt is empty. You'll need to enter keywords manually.")
    keywords = prompt_extra_keywords(base_keywords)

    if not keywords:
        print("[ERROR] No keywords to search for. Add some to keywords.txt and re-run.\n")
        sys.exit(1)

    # analyze
    all_rows = []
    for pdf in pdfs:
        rows = search_document(pdf, keywords)
        all_rows.extend(rows)
        total = len(rows)
        print(f"    → {total} match(es) found")

    # save
    print()
    out_path = save_results(all_rows, keywords)
    print(f"✓ Results saved to: {out_path}")
    print(f"  Total matches: {len(all_rows)}")
    print()

    # open in Finder
    os.system(f'open "{RESULTS_DIR}"')


if __name__ == "__main__":
    main()
