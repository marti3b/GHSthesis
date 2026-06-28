# GHS Thesis — Congruence Analysis

Keyword analysis tool for identifying **Normative Preferences (NP)** and **Strategic Autonomy (SA)** terms in Global Health Security documents.

## Structure

```
documents/    ← drop PDFs here
results/      ← Excel output appears here
keywords.txt  ← keyword list (NP and SA terms)
analyze.py    ← main analysis script
setup.sh      ← one-time setup (installs dependencies)
```

## Setup (run once)

```bash
bash setup.sh
```

Or manually:

```bash
pip3 install pdfplumber pandas openpyxl
```

## Usage

1. Drop PDF documents into `documents/`
2. Edit `keywords.txt` to update NP or SA terms
3. Run:

```bash
python3 analyze.py
```

The script will show the loaded keywords, let you add extras for that run, then save an Excel file to `results/` with two sheets:

- **Matches** — every hit with document name, page number, and surrounding sentence
- **Summary** — keyword count per document

## Keywords

Keywords are grouped into two categories in `keywords.txt`:

- **NP terms** — normative preference language (solidarity, equity, burden sharing, etc.)
- **SA terms** — strategic autonomy language (resilience, supply chain, geopolitical, etc.)

Section headers (`**NP terms:**`, `**SA terms:**`) and lines starting with `#` are ignored by the script.
