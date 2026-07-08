#!/usr/bin/env python3
"""
test_extraction_approaches.py
==============================
Diagnostic tool: runs every viable pdfplumber extraction approach against the first
N pages of your ledger PDF and prints/saves the RAW output of each side by side, so we
can compare and pick the right one for your actual file BEFORE building the final
parser. Nothing here is opinionated about column mapping yet — this is purely
"what does each extraction method see".

100% local / offline (pdfplumber only). Safe for confidential client data.

Approaches tested per page:
  1. extract_tables()  — vertical='lines',  horizontal='lines'   (real ruled borders)
  2. extract_tables()  — vertical='text',   horizontal='text'    (whitespace/text-gap based)
  3. extract_tables()  — vertical='lines',  horizontal='text'    (mixed)
  4. extract_text()                          (plain reading-order text, no layout)
  5. extract_text(layout=True)               (text with preserved column spacing/whitespace)
  6. extract_words() + geometric clustering  (raw word positions grouped into rows/cols)

Usage
-----
    python test_extraction_approaches.py input.pdf [--pages 3] [--save]

    --pages N   how many pages to test (default 3)
    --save      also write full output to test_extraction_output.txt (console only otherwise
                shows a truncated preview per approach)
"""

import sys
import argparse
from pathlib import Path

import pdfplumber


def divider(title):
    line = "=" * 90
    return f"\n{line}\n{title}\n{line}"


def preview(text, max_lines=15):
    """Return only the first max_lines lines of a block of text, for console preview."""
    lines = text.splitlines()
    if len(lines) <= max_lines:
        return text
    return "\n".join(lines[:max_lines]) + f"\n... [{len(lines) - max_lines} more lines truncated — see full output if --save used]"


def format_table(table):
    """Pretty-print a pdfplumber table (list of list of cells) as a simple grid."""
    if not table:
        return "(no table detected)"
    out_lines = []
    for row_idx, row in enumerate(table):
        cells = [("" if c is None else str(c).strip()) for c in row]
        out_lines.append(f"  row {row_idx:>3}: {cells}")
    return "\n".join(out_lines)


def run_extract_tables(page, settings, label):
    try:
        tables = page.extract_tables(settings)
    except Exception as e:
        return f"[{label}] ERROR: {e}"

    if not tables:
        return f"[{label}] No tables detected."

    out = [f"[{label}] {len(tables)} table(s) detected."]
    for t_idx, table in enumerate(tables):
        out.append(f"\n--- Table {t_idx} ({len(table)} rows) ---")
        out.append(format_table(table))
    return "\n".join(out)


def run_extract_text_plain(page):
    try:
        text = page.extract_text() or ""
    except Exception as e:
        return f"ERROR: {e}"
    if not text.strip():
        return "(no text extracted)"
    return text


def run_extract_text_layout(page):
    try:
        text = page.extract_text(layout=True) or ""
    except Exception as e:
        return f"ERROR: {e}"
    if not text.strip():
        return "(no text extracted)"
    return text


def run_word_clustering(page):
    try:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    except Exception as e:
        return f"ERROR: {e}"
    if not words:
        return "(no words extracted)"

    lines = {}
    for w in words:
        key = round(w["top"] / 3) * 3
        lines.setdefault(key, []).append(w)

    out = []
    for key in sorted(lines.keys()):
        line_words = sorted(lines[key], key=lambda w: w["x0"])
        cells = []
        current_cell = [line_words[0]]
        for prev, curr in zip(line_words, line_words[1:]):
            gap = curr["x0"] - prev["x1"]
            if gap > 8:
                cells.append(" ".join(w["text"] for w in current_cell))
                current_cell = [curr]
            else:
                current_cell.append(curr)
        cells.append(" ".join(w["text"] for w in current_cell))
        out.append(f"  {cells}")

    return "\n".join(out) if out else "(no rows formed)"


def summarize_page(page, page_num):
    """Return (report_text, stats_dict) for one page across all approaches."""
    sections = []
    stats = {}

    # 1-3: table extraction variants
    table_settings = [
        ("lines/lines", {"vertical_strategy": "lines", "horizontal_strategy": "lines"}),
        ("text/text", {"vertical_strategy": "text", "horizontal_strategy": "text"}),
        ("lines/text", {"vertical_strategy": "lines", "horizontal_strategy": "text"}),
    ]
    for label, settings in table_settings:
        result = run_extract_tables(page, settings, label)
        sections.append(divider(f"PAGE {page_num} — extract_tables ({label})"))
        sections.append(result)
        try:
            n_tables = len(page.extract_tables(settings) or [])
            n_rows = sum(len(t) for t in (page.extract_tables(settings) or []))
        except Exception:
            n_tables, n_rows = 0, 0
        stats[f"tables_{label}"] = (n_tables, n_rows)

    # 4: plain text
    plain = run_extract_text_plain(page)
    sections.append(divider(f"PAGE {page_num} — extract_text() plain"))
    sections.append(preview(plain))
    stats["plain_text_lines"] = len(plain.splitlines())

    # 5: layout text
    layout = run_extract_text_layout(page)
    sections.append(divider(f"PAGE {page_num} — extract_text(layout=True)"))
    sections.append(preview(layout))
    stats["layout_text_lines"] = len(layout.splitlines())

    # 6: word clustering
    words_out = run_word_clustering(page)
    sections.append(divider(f"PAGE {page_num} — extract_words() + geometric clustering"))
    sections.append(preview(words_out))
    stats["word_cluster_lines"] = len(words_out.splitlines())

    return "\n".join(sections), stats


def recommend(all_stats):
    """Very simple heuristic recommendation based on which approach produced the most
    plausible row counts consistently across pages."""
    lines = [divider("RECOMMENDATION")]

    table_labels = ["lines/lines", "text/text", "lines/text"]
    totals = {label: 0 for label in table_labels}
    tables_found_any_page = {label: False for label in table_labels}

    for stats in all_stats:
        for label in table_labels:
            n_tables, n_rows = stats.get(f"tables_{label}", (0, 0))
            totals[label] += n_rows
            if n_tables > 0:
                tables_found_any_page[label] = True

    best_label = max(totals, key=lambda l: totals[l])
    best_rows = totals[best_label]

    if best_rows > 0:
        lines.append(
            f"Best table-based approach: extract_tables ({best_label}) — "
            f"{best_rows} total row(s) detected across tested pages."
        )
        lines.append("-> Your PDF appears to have a genuinely structured/detectable table. "
                      "We should build the final parser on extract_tables(), using this setting "
                      "as the primary strategy.")
    else:
        lines.append("No extract_tables() setting detected any table structure on the tested pages.")
        avg_plain = sum(s["plain_text_lines"] for s in all_stats) / max(len(all_stats), 1)
        avg_layout = sum(s["layout_text_lines"] for s in all_stats) / max(len(all_stats), 1)
        lines.append(f"extract_text() plain avg lines/page: {avg_plain:.1f}")
        lines.append(f"extract_text(layout=True) avg lines/page: {avg_layout:.1f}")
        lines.append("-> Table detection failed; fall back to word-position clustering "
                      "(extract_words()) or extract_text(layout=True) parsing with regex, "
                      "since layout=True preserves column spacing as whitespace.")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Test all pdfplumber extraction approaches on a PDF.")
    parser.add_argument("input_pdf", help="Path to the PDF file to test")
    parser.add_argument("--pages", type=int, default=3, help="Number of pages to test (default: 3)")
    parser.add_argument("--save", action="store_true", help="Save full output to test_extraction_output.txt")
    args = parser.parse_args()

    input_path = Path(args.input_pdf)
    if not input_path.exists():
        print(f"ERROR: file not found: {input_path}")
        sys.exit(1)

    full_report = []
    all_stats = []

    with pdfplumber.open(str(input_path)) as pdf:
        total_pages = len(pdf.pages)
        pages_to_test = min(args.pages, total_pages)
        print(f"Opened '{input_path}' — {total_pages} page(s) total. Testing first {pages_to_test} page(s).\n")

        for i, page in enumerate(pdf.pages[:pages_to_test], start=1):
            print(f"Testing page {i}/{pages_to_test}...")
            report, stats = summarize_page(page, i)
            full_report.append(report)
            all_stats.append(stats)

    recommendation = recommend(all_stats)
    print(recommendation)

    if args.save:
        out_path = Path("test_extraction_output.txt")
        out_path.write_text("\n".join(full_report) + "\n" + recommendation, encoding="utf-8")
        print(f"\nFull detailed output written to: {out_path.resolve()}")
    else:
        print(divider("PREVIEW OF DETAILED OUTPUT (per-page, truncated)"))
        print("\n".join(full_report))
        print("\n(Run with --save to write the FULL untruncated output to a file for close inspection.)")


if __name__ == "__main__":
    main()
