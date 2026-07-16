"""
pdf_to_excel.py
----------------
Converts a two-column balance sheet PDF (ATTIVITA' on the left,
PASSIVITA' on the right) into a structured Excel file.

STATUS: DRAFT — built from spec only, no sample PDF was available to
validate coordinates against. Every value marked "TUNE" below is a guess
and will very likely need adjustment once run against your real file.
Run with --debug first (see bottom of file) and inspect the printed
word coordinates before trusting the output.

Output columns: PAGE, SIDE, CODE, SUB_CODE, DESCRIPTION, AMOUNT

Usage:
    python pdf_to_excel.py input.pdf output.xlsx
    python pdf_to_excel.py input.pdf output.xlsx --debug   # dumps word coords for page 2
"""

import sys
import re
import pdfplumber
import pandas as pd

# ----------------------------------------------------------------------
# TUNABLE PARAMETERS — adjust these once you can see the real PDF layout
# ----------------------------------------------------------------------

# Row-clustering tolerance: words within this many points of vertical
# ("top") distance are considered part of the same text row.
ROW_TOLERANCE = 3  # TUNE: increase if rows are splitting; decrease if merging

# First page containing data (per spec: data starts on page 2, 0-indexed -> 1)
FIRST_DATA_PAGE_INDEX = 1  # page 2 in 1-indexed terms

# For test runs: limit how many data pages to process (None = process all)
# e.g. MAX_PAGES = 3 processes pages 2, 3, 4 only.
MAX_PAGES = 3  # TUNE: set to None once you're ready to run the full file

# Regex patterns for row classification
# TUNE: confirm these against real CODE values in the PDF.
# Hierarchy is 3 header levels + 1 detail level, e.g.:
#   AA        <- level 0 (top-level "parent category") -- EXCLUDED from output
#   AA01      <- level 1 header
#   AA0101    <- level 2 header
#   AA0101 000001 <- detail row (level-2 code + 6-digit sub-code)
PARENT_CODE_RE = re.compile(r"^[A-Z]{2}$")         # e.g. AA, BB -- top-level, excluded
LEVEL1_CODE_RE = re.compile(r"^[A-Z]{2}\d{2}$")    # e.g. AA01
LEVEL2_CODE_RE = re.compile(r"^[A-Z]{2}\d{4}$")    # e.g. AA0101 (header OR detail's CODE)
SUB_CODE_RE = re.compile(r"^\d{6}$")               # e.g. 000005

# Totals / subtotal rows to ignore
TOTALS_RE = re.compile(r"^(TOTALE|TOTALI|TOTAL|SUBTOTAL|GRAND\s+TOTAL)\b", re.IGNORECASE)
SEPARATOR_RE = re.compile(r"^[_=\-]+$")  # lines of underscores, equals, or dashes

# Italian-format amount, e.g. 26.447,61 or 2.160.000,00 or (1.234,56)
AMOUNT_RE = re.compile(r"^\(?-?\d{1,3}(\.\d{3})*,\d{2}\)?$")


def parse_italian_amount(token: str) -> float:
    """Convert an Italian-formatted number string to a float.
    Handles thousands separator '.' and decimal separator ','.
    Parenthesized values are treated as negative."""
    negative = token.startswith("(") and token.endswith(")")
    cleaned = token.strip("()")
    cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        value = float(cleaned)
    except ValueError:
        return 0.0
    return -value if negative else value


def cluster_rows(words, tolerance=ROW_TOLERANCE):
    """Group words into visual rows based on their 'top' coordinate."""
    if not words:
        return []
    words_sorted = sorted(words, key=lambda w: w["top"])
    rows = []
    current_row = [words_sorted[0]]
    current_top = words_sorted[0]["top"]

    for w in words_sorted[1:]:
        if abs(w["top"] - current_top) <= tolerance:
            current_row.append(w)
        else:
            rows.append(current_row)
            current_row = [w]
            current_top = w["top"]
    rows.append(current_row)

    # sort words left-to-right within each row
    for row in rows:
        row.sort(key=lambda w: w["x0"])
    return rows


def classify_and_parse_row(row_words):
    """
    Given a list of word dicts (already sorted left-to-right) for one row,
    return a dict with CODE/SUB_CODE/DESCRIPTION/AMOUNT, or None if the
    row should be skipped (totals/subtotals/blank/unrecognized).
    """
    tokens = [w["text"] for w in row_words]
    if not tokens:
        return None

    first_token = tokens[0]

    # Skip totals / separator rows (underscores, dashes, equals signs)
    full_line = " ".join(tokens)
    if TOTALS_RE.match(first_token) or SEPARATOR_RE.match(first_token):
        return None
    if any(SEPARATOR_RE.match(t) for t in tokens):
        return None

    # Skip a bare amount with no code/description (summary line fragment)
    if len(tokens) == 1 and AMOUNT_RE.match(tokens[0]):
        return None

    # Identify amount: last token matching AMOUNT_RE, if any
    amount = 0.0
    desc_tokens = tokens[:]
    if AMOUNT_RE.match(tokens[-1]):
        amount = parse_italian_amount(tokens[-1])
        desc_tokens = tokens[:-1]

    if not desc_tokens:
        return None

    code = None
    sub_code = ""
    remaining = desc_tokens[:]
    first = remaining[0]

    # Top-level "parent category" row (e.g. AA, BB) — excluded from output entirely
    if PARENT_CODE_RE.match(first):
        return None

    # Level-2 code (e.g. AA0101): could be a header OR a detail row.
    # It's a detail row if the next token is a 6-digit SUB_CODE.
    if LEVEL2_CODE_RE.match(first):
        code = first
        remaining = remaining[1:]
        if remaining and SUB_CODE_RE.match(remaining[0]):
            sub_code = remaining[0]
            remaining = remaining[1:]
        description = " ".join(remaining)
        return {
            "CODE": code,
            "SUB_CODE": sub_code,
            "DESCRIPTION": description,
            "AMOUNT": amount,
        }

    # Level-1 header row (e.g. AA01), no SUB_CODE
    if LEVEL1_CODE_RE.match(first):
        code = first
        remaining = remaining[1:]
        description = " ".join(remaining)
        return {
            "CODE": code,
            "SUB_CODE": "",
            "DESCRIPTION": description,
            "AMOUNT": amount,
        }

    # No recognizable code — likely a wrapped description continuation.
    # Signal this to the caller so it can be appended to the previous row.
    return {"CONTINUATION": full_line}


def extract_side_rows(words, page_num, side_label):
    """Cluster words into rows and classify each one for a single side
    (left or right) of a page."""
    results = []
    rows = cluster_rows(words)

    for row_words in rows:
        parsed = classify_and_parse_row(row_words)
        if parsed is None:
            continue
        if "CONTINUATION" in parsed:
            # Append to previous row's description, if one exists
            if results:
                results[-1]["DESCRIPTION"] = (
                    results[-1]["DESCRIPTION"] + " " + parsed["CONTINUATION"]
                ).strip()
            continue
        parsed["PAGE"] = page_num
        parsed["SIDE"] = side_label
        results.append(parsed)

    return results


def process_pdf(input_path, debug=False):
    all_rows = []
    pages_processed = 0

    with pdfplumber.open(input_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            if page_index < FIRST_DATA_PAGE_INDEX:
                continue

            if MAX_PAGES is not None and pages_processed >= MAX_PAGES:
                if debug:
                    print(f"\n[MAX_PAGES={MAX_PAGES} reached, stopping]")
                break

            page_num = page_index + 1
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            mid_x = page.width / 2  # TUNE: replace with real gap-detection if this misaligns

            left_words = [w for w in words if w["x0"] < mid_x]
            right_words = [w for w in words if w["x0"] >= mid_x]

            if debug:
                print(f"\n{'=' * 70}")
                print(f"PAGE {page_num}  |  page.width={page.width:.1f}  mid_x={mid_x:.1f}")
                print(f"  total words={len(words)}  left={len(left_words)}  right={len(right_words)}")
                print(f"\n  --- raw words (first 30) ---")
                for w in words[:30]:
                    side_tag = "L" if w["x0"] < mid_x else "R"
                    print(f"  [{side_tag}] x0={w['x0']:.1f} top={w['top']:.1f} text={w['text']!r}")

            left_rows = extract_side_rows(left_words, page_num, "ATTIVITA'")
            right_rows = extract_side_rows(right_words, page_num, "PASSIVITA'")

            if debug:
                print(f"\n  --- classified rows: ATTIVITA' (left), {len(left_rows)} rows ---")
                for r in left_rows:
                    print(f"    CODE={r['CODE']!r:10} SUB_CODE={r['SUB_CODE']!r:8} "
                          f"AMOUNT={r['AMOUNT']:>15} DESC={r['DESCRIPTION']!r}")
                print(f"\n  --- classified rows: PASSIVITA' (right), {len(right_rows)} rows ---")
                for r in right_rows:
                    print(f"    CODE={r['CODE']!r:10} SUB_CODE={r['SUB_CODE']!r:8} "
                          f"AMOUNT={r['AMOUNT']:>15} DESC={r['DESCRIPTION']!r}")

            all_rows.extend(left_rows)
            all_rows.extend(right_rows)
            pages_processed += 1

    return all_rows


def write_excel(rows, output_path):
    df = pd.DataFrame(rows, columns=["PAGE", "SIDE", "CODE", "SUB_CODE", "DESCRIPTION", "AMOUNT"])
    df.to_excel(output_path, index=False, sheet_name="Data")
    print(f"Wrote {len(df)} rows to {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pdf_to_excel.py input.pdf output.xlsx [--debug]")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_xlsx = sys.argv[2]
    debug_mode = "--debug" in sys.argv

    rows = process_pdf(input_pdf, debug=debug_mode)
    write_excel(rows, output_xlsx)
