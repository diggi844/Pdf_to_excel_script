"""
Simplified PDF -> Excel parser (skeleton)

This version:
- Uses ONLY pdfplumber.extract_text()
- Detects transaction header from raw text
- Skips RIPORTI and TOTALI PROGRESSIVI rows
- Provides hooks for right-to-left parsing
"""

import re
import sys
from pathlib import Path

import pdfplumber
import pandas as pd


EXPECTED_HEADER = "DATANRIFNPROTDESCRIZIONECONTODAREAVERE"

HEADERS = [
    "DATA",
    "N.RIF.",
    "N.Prot.",
    "DESCRIZIONE",
    "CONTO",
    "DARE",
    "AVERE",
]


def normalize(text: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", text.upper())


def is_transaction_header(line: str) -> bool:
    return normalize(line) == EXPECTED_HEADER


def should_skip_row(line: str) -> bool:
    n = normalize(line)
    return (
        n.startswith("RIPORTI")
        or n.startswith("TOTALIPROGRESSIVI")
    )


def parse_dare_avere(token: str):
    """Comma is treated as separator between DARE and AVERE."""
    if "," not in token:
        return 0.0, 0.0

    left, right = token.split(",", 1)

    dare = float(left) if left else 0.0
    avere = float(right) if right else 0.0

    return dare, avere


DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")

def parse_row(line, last_nprot):

    row = {
        "DATA": "",
        "N.RIF.": "",
        "N.Prot.": last_nprot,
        "DESCRIZIONE": "",
        "CONTO": "",
        "DARE": "",
        "AVERE": "",
    }

    tokens = line.split()

    # ------------------------
    # STEP 1 : DARE / AVERE
    # ------------------------

    if tokens and "," in tokens[-1]:

        amount = tokens.pop()

        left, right = amount.split(",", 1)

        row["DARE"] = left
        row["AVERE"] = right

    # ------------------------
    # STEP 2 : DATA
    # ------------------------

    for i, t in enumerate(tokens):

        if DATE_RE.fullmatch(t):

            row["DATA"] = t

            tokens.pop(i)

            break

    # ------------------------
    # STEP 3 : N.RIF.
    # ------------------------

    if tokens and tokens[0].isalpha():

        row["N.RIF."] = tokens.pop(0)

    # ------------------------
    # STEP 4 : N.PROT.
    # ------------------------

    if tokens and tokens[0].isdigit():

        row["N.Prot."] = tokens.pop(0)

        last_nprot = row["N.Prot."]

    else:

        row["N.Prot."] = last_nprot

    # ------------------------
    # STEP 5 : TEMP
    # ------------------------

    row["DESCRIZIONE"] = " ".join(tokens)

    return row, last_nprot


def convert(pdf_path, output_path):

    records = []
    last_nprot = ""
    with pdfplumber.open(pdf_path) as pdf:

        # ---------- ONLY FIRST PAGE ----------
        page = pdf.pages[0]

        text = page.extract_text()

        if not text:
            print("No text found on first page.")
            return

        inside_transactions = False

        for line in text.splitlines():

            print(line)   # Debug

            if not inside_transactions:

                if is_transaction_header(line):

                    inside_transactions = True

                    print("\nTransaction Header Found\n")

                continue

            if should_skip_row(line):

                print(f"Skipping : {line}")

                continue

            # Write raw transaction row
            record, last_nprot = parse_row(line, last_nprot)

            records.append(record)

    df = pd.DataFrame(records, columns=HEADERS)

    print(df)

    df.to_excel(output_path, index=False)

    print(f"\nSaved : {output_path}")


if __name__ == "__main__":

    if len(sys.argv) < 2:
        print("Usage: python pdf_to_excel.py input.pdf [output.xlsx]")
        sys.exit(1)

    pdf = sys.argv[1]

    if len(sys.argv) > 2:
        output = sys.argv[2]
    else:
        output = str(Path(pdf).with_suffix(".xlsx"))

    convert(pdf, output)
