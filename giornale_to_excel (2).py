"""
PDF -> Excel parser for the Italian financial ledger layout:
DATA / N.RIF. / N.Prot. / DESCRIZIONE / CONTO / DARE / AVERE

Built per the functional spec in `Prompts`, later revised to be more
transaction-aware and tolerant of misaligned rows:
  - pdfplumber only (extract_text() to locate the header, extract_words()
    for word coordinates)
  - DARE/AVERE and CONTO are still resolved by x-position, since they sit
    at the row's fixed right edge and rarely drift
  - DATA, N.RIF., N.Prot., and DESCRIZIONE are resolved by token CONTENT
    and SEQUENCE rather than x-position, since misaligned rows break a
    pure coordinate-bin approach for these columns:
      * DATA = first token anywhere in the remaining zone matching
        dd/mm/yyyy
      * N.RIF. = the numeric token immediately after DATA
      * N.Prot. = the numeric token immediately after N.RIF.
      * DESCRIZIONE = everything else, in original left-to-right order
  - amounts use a comma as a DECIMAL POINT (e.g. "61,00" -> 61.00), never
    split into two fields
  - DARE vs AVERE is decided purely by which column-x-range the amount
    word falls into
  - CONTO = everything (numeric identifiers + description words) sitting
    in the CONTO column's x-range
  - N.Prot. carries forward the last non-empty value ONLY on rows that
    otherwise contain real transaction content; fully blank lines never
    inherit it and never update it. DATA, N.RIF., DESCRIZIONE, CONTO are
    left blank when missing (no carry-forward for those)
  - runs across the full document by default (see `process_all_pages` flag
    on `convert()` if you ever need to restrict it to a subset for testing)

NOTE ON THE SPEC: the "Amount Rule" section says to "replace the comma
with 0", but its own worked example converts "61,00" -> "61.00" (comma
replaced with a decimal point). This implementation follows the worked
example. Flag this if that's not what you intended.
"""

import re
import sys
import traceback
from pathlib import Path

import pdfplumber
import pandas as pd


# -------------------- CONFIG --------------------

EXPECTED_HEADER_NORM = "DATANRIFNPROTDESCRIZIONECONTODAREAVERE"

# Internal column keys, in left-to-right order as they appear in the header
COLUMN_LABELS = ["DATA", "NRIF", "NPROT", "DESCRIZIONE", "CONTO", "DARE", "AVERE"]

# Output Excel columns, in required order
HEADERS = ["DATA", "N.RIF.", "N.Prot.", "DESCRIZIONE", "CONTO", "DARE", "AVERE"]

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
# N.RIF. and N.Prot. are both purely numeric tokens. They are told apart not
# by shape but by SEQUENCE: within a transaction, whichever numeric token
# comes immediately after DATA is N.RIF., and whichever numeric token comes
# immediately after THAT is N.Prot. (see parse_transaction_line, STEP 3).
NRIF_RE = re.compile(r"^\d+$")
NPROT_RE = re.compile(r"^\d+$")

Y_TOLERANCE = 3    # px tolerance for grouping words into the same visual line


# -------------------- HELPERS --------------------

def normalize_header(text):
    return re.sub(r"[^A-Z0-9]", "", text.upper())


def should_skip_row(line_text):
    n = normalize_header(line_text)
    return n.startswith("RIPORTI") or n.startswith("TOTALIPROGRESSIVI")


def is_separator_row(line_text):
    """True for divider/rule lines made up of nothing but dash-like
    characters (e.g. '---------------------------'), which some PDFs use
    as visual separators between sections and which should never become a
    data row.
    """
    compact = re.sub(r"\s+", "", line_text)
    return bool(compact) and bool(re.fullmatch(r"[-_=]+", compact))


def italian_amount_to_float(raw):
    """Comma is a decimal point here, not a field separator.
    '61,00' -> 61.00 ; '1.234,56' -> 1234.56
    """
    if not raw:
        return None
    val = raw.strip().replace(".", "").replace(",", ".")
    try:
        return float(val)
    except ValueError:
        return None


# -------------------- LINE GROUPING --------------------

def group_words_into_lines(words, y_tolerance=Y_TOLERANCE):
    """Cluster pdfplumber words into visual lines using their 'top' coordinate."""
    lines = []
    buffer = []
    ref_top = None

    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if ref_top is None or abs(w["top"] - ref_top) <= y_tolerance:
            buffer.append(w)
            ref_top = w["top"] if ref_top is None else ref_top
        else:
            lines.append(sorted(buffer, key=lambda w: w["x0"]))
            buffer = [w]
            ref_top = w["top"]

    if buffer:
        lines.append(sorted(buffer, key=lambda w: w["x0"]))

    return lines


# -------------------- HEADER / COLUMN DETECTION --------------------

def detect_header_columns(line_words):
    """
    Try to match this line's words against the expected header, consuming
    characters label-by-label. Handles labels split across multiple words
    by the PDF extractor (e.g. 'D A R E' as four separate word objects).
    Returns {label: x0_start} on a full match, else None.
    """
    col_starts = {}
    label_idx = 0
    remaining = COLUMN_LABELS[0]
    current_start_x = None

    for w in line_words:
        norm = normalize_header(w["text"])
        if not norm:
            continue
        if current_start_x is None:
            current_start_x = w["x0"]

        while norm:
            if label_idx >= len(COLUMN_LABELS):
                return None  # extra text beyond the 7 expected labels

            take = min(len(norm), len(remaining))
            chunk, norm = norm[:take], norm[take:]

            if not remaining.startswith(chunk):
                return None  # doesn't match the expected header at all

            remaining = remaining[len(chunk):]

            if remaining == "":
                col_starts[COLUMN_LABELS[label_idx]] = current_start_x
                label_idx += 1
                current_start_x = None
                if label_idx < len(COLUMN_LABELS):
                    remaining = COLUMN_LABELS[label_idx]

    return col_starts if label_idx == len(COLUMN_LABELS) else None


def build_column_bins(col_starts):
    """Return (labels, boundaries) sorted left-to-right for word binning.
    `boundaries` are the midpoints between each pair of adjacent column
    starts — a word is assigned to a column based on which side of these
    midpoints it falls on, rather than the raw column start. This is far
    more forgiving of real-world text not being perfectly left-aligned to
    the header label above it (e.g. a DESCRIZIONE word sitting slightly
    left of DESCRIZIONE's own header start would otherwise get misassigned
    to N.Prot. under a naive "nearest preceding start" rule).
    """
    ordered = sorted(col_starts.items(), key=lambda kv: kv[1])
    labels = [label for label, _ in ordered]
    starts = [x for _, x in ordered]
    boundaries = [(starts[i] + starts[i + 1]) / 2 for i in range(len(starts) - 1)]
    return labels, boundaries


def assign_column(x0, labels, boundaries):
    idx = 0
    for i, boundary in enumerate(boundaries):
        if x0 >= boundary:
            idx = i + 1
        else:
            break
    return labels[idx]


# -------------------- TRANSACTION ROW PARSER (RIGHT TO LEFT) --------------------

def parse_transaction_line(line_words, labels, boundaries, last_nprot):
    """
    Transaction-aware parser. Only DARE/AVERE and CONTO are still resolved
    by x-position (they sit at the fixed right edge of the row and rarely
    drift). Everything to their left -- DATA, N.RIF., N.Prot., DESCRIZIONE --
    is resolved by searching token CONTENT and SEQUENCE instead of raw
    x-coordinates, so slightly misaligned rows still parse correctly:

        1. AMOUNT (DARE or AVERE, whichever column the last word sits in)
        2. CONTO       - keep popping while still inside the CONTO column
        3. Search the remaining tokens (left-to-right) for the first one
           matching the DATA pattern (dd/mm/yyyy). If found:
             - the numeric token immediately after it is N.RIF.
             - the numeric token immediately after N.RIF. is N.Prot.
           If no DATA token exists, N.RIF./N.Prot. are left blank rather
           than guessed, since there is no longer a reliable position to
           anchor on.
        4. Every remaining token (including any date-like or numeric-looking
           token that wasn't consumed above) falls into DESCRIZIONE.
    """
    remaining = sorted(line_words, key=lambda w: w["x0"])

    def col_at(x0):
        return assign_column(x0, labels, boundaries)

    row = {
        "DATA": "",
        "N.RIF.": "",
        "N.Prot.": "",
        "DESCRIZIONE": "",
        "CONTO": "",
        "DARE": "",
        "AVERE": "",
    }

    # STEP 1: AMOUNT at the extreme right (position-based -- reliable, fixed
    # right edge of the row)
    if remaining:
        last = remaining[-1]
        col = col_at(last["x0"])
        if col in ("DARE", "AVERE"):
            amount = italian_amount_to_float(last["text"])
            if amount is not None:
                row[col] = round(amount, 2)
                remaining.pop()

    # STEP 2: CONTO block (position-based) - keep collecting while still in
    # the CONTO column. NOTE: the amount was already popped in STEP 1, so if
    # a word's x0 still falls in the DARE/AVERE zone at this point, it's
    # virtually certain to be CONTO text overflowing rightward, not a second
    # amount -- so DARE/AVERE also counts as "still CONTO" here.
    conto_words = []
    while remaining and col_at(remaining[-1]["x0"]) in ("CONTO", "DARE", "AVERE"):
        conto_words.insert(0, remaining.pop()["text"])
    row["CONTO"] = " ".join(conto_words)

    # From here on, `remaining` holds only the DATA/N.RIF./N.Prot./DESCRIZIONE
    # zone. Work with it as an ordered, left-to-right token list rather than
    # by x-position bins.
    tokens = [w["text"].strip() for w in remaining]

    # STEP 3: find DATA by content, anywhere in the remaining tokens -- not
    # by column position, so a date-shaped word sitting in DESCRIZIONE isn't
    # mistaken for the transaction's DATA (and vice versa).
    data_idx = None
    for i, t in enumerate(tokens):
        if DATE_RE.match(t):
            data_idx = i
            break

    nrif_idx = None
    nprot_idx = None

    if data_idx is not None:
        row["DATA"] = tokens[data_idx]

        # N.RIF. = the numeric token immediately after DATA
        if data_idx + 1 < len(tokens) and NRIF_RE.match(tokens[data_idx + 1]):
            nrif_idx = data_idx + 1
            row["N.RIF."] = tokens[nrif_idx]

            # N.Prot. = the numeric token immediately after N.RIF.
            if nrif_idx + 1 < len(tokens) and NPROT_RE.match(tokens[nrif_idx + 1]):
                nprot_idx = nrif_idx + 1
                row["N.Prot."] = tokens[nprot_idx]

    # STEP 4: DESCRIZIONE = every token not consumed by DATA/N.RIF./N.Prot.
    # above, in their original left-to-right order. This deliberately
    # includes date-like or numeric-looking tokens that didn't satisfy the
    # DATA -> N.RIF. -> N.Prot. sequence -- those are genuine description
    # content (e.g. "100567 Ts.bsst.n. 6724/8-2025 yep"), not misclassified
    # metadata.
    consumed = {i for i in (data_idx, nrif_idx, nprot_idx) if i is not None}
    desc_tokens = [t for i, t in enumerate(tokens) if i not in consumed]
    row["DESCRIZIONE"] = " ".join(desc_tokens).strip()

    # N.Prot. carry-forward: only apply it to rows that are genuinely part
    # of a transaction (i.e. some other field was actually populated).
    # Completely empty lines never inherit a value, and never advance
    # last_nprot either.
    has_content = any(
        row[c] != "" for c in ("DATA", "N.RIF.", "DESCRIZIONE", "CONTO", "DARE", "AVERE")
    )
    if row["N.Prot."] == "":
        if has_content:
            row["N.Prot."] = last_nprot
    else:
        last_nprot = row["N.Prot."]

    return row, last_nprot


# -------------------- PAGE PROCESSING --------------------

def process_page(page, last_nprot, labels, boundaries):
    """
    Parse one page. `labels`/`boundaries` carry the column layout forward from
    a previous page (None on the very first page, until the header is found).
    Returns (records, last_nprot, labels, boundaries).
    """
    words = page.extract_words()
    lines = group_words_into_lines(words)

    records = []
    inside_transactions = labels is not None

    for line_words in lines:
        if not line_words:
            continue

        if not inside_transactions:
            col_starts = detect_header_columns(line_words)
            if col_starts:
                labels, boundaries = build_column_bins(col_starts)
                inside_transactions = True
            continue

        # Header repeating on this page (e.g. per-page banner) -> skip it
        if detect_header_columns(line_words):
            continue

        line_text = " ".join(w["text"] for w in line_words)
        if should_skip_row(line_text) or is_separator_row(line_text):
            continue

        row, last_nprot = parse_transaction_line(line_words, labels, boundaries, last_nprot)

        if any(row[col] != "" for col in HEADERS):
            # DARE/AVERE default to 0 rather than blank when a real
            # transaction row simply has no value in that column.
            row["DARE"] = row["DARE"] if row["DARE"] != "" else 0
            row["AVERE"] = row["AVERE"] if row["AVERE"] != "" else 0
            records.append(row)

    return records, last_nprot, labels, boundaries


# -------------------- MAIN DRIVER --------------------

def convert(pdf_path, output_path, process_all_pages=True):
    """
    process_all_pages=True processes every page in the PDF (the normal,
    production behavior). Set to False if you ever need to restrict a run
    to just the first page, e.g. for quick testing against a new report
    layout before trusting it on the full document.
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF")

    records = []
    last_nprot = ""
    labels, boundaries = None, None

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages if process_all_pages else pdf.pages[:1]
        print(f"\nProcessing {len(pages)} page(s) of {len(pdf.pages)} total")

        for page_num, page in enumerate(pages, start=1):
            try:
                page_records, last_nprot, labels, boundaries = process_page(
                    page, last_nprot, labels, boundaries
                )
                records.extend(page_records)
            except Exception:
                print(f"\nPAGE PROCESSING ERROR (page {page_num})")
                traceback.print_exc()

    if labels is None:
        print("\nWARNING: transaction header was never found — 0 rows extracted.")

    df = pd.DataFrame(records, columns=HEADERS)
    df.to_excel(output_path, index=False)

    print(f"\nTotal rows extracted: {len(df)}")
    print(f"Excel saved to: {output_path}")

    return df


# -------------------- ENTRY POINT --------------------

if __name__ == "__main__":
    try:
        if len(sys.argv) < 2:
            print("Usage: python giornale_to_excel.py input.pdf [output.xlsx]")
            sys.exit(1)

        pdf = sys.argv[1]
        output = sys.argv[2] if len(sys.argv) > 2 else str(Path(pdf).with_suffix(".xlsx"))

        convert(pdf, output, process_all_pages=True)

    except Exception:
        print("\nPROGRAM FAILED")
        traceback.print_exc()
