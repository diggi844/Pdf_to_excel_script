import fitz
import pandas as pd
import re
from pathlib import Path
import traceback


# -------------------- REGEX PATTERNS --------------------

HEADER_PATTERN = re.compile(
    r"""
    ^\s*
    (?P<nprogr>\d{6,8})\s+
    (?P<dtcpu>\d{6})\s+
    (?P<ndoc>\d{6,12})\s+
    (?P<dtreg>\d{6})\s+
    (?P<dtdoc>\d{6})\s+
    (?P<td>[A-Z0-9]{1,3})
    (?:\s+(?P<desc>.+))?
    """,
    re.VERBOSE
)

ACCOUNT_PATTERN = re.compile(
    r"""
    (?P<rdo>\d{2,6})\s+
    (?P<c>[A-Z])\s+
    (?P<account>[A-Z0-9]{5,12})
    """,
    re.VERBOSE
)

AMOUNT_PATTERN = re.compile(
    r"""
    -?\d{1,3}(?:\.\d{3})*,\d{2}-?
    """,
    re.VERBOSE
)

LOCATION_PATTERN = re.compile(r"^\d{5}\s+[A-ZÀ-ÖØ-Þ].+")
ADDRESS_PATTERN = re.compile(
    r"^(VIA|VIALE|PIAZZA|CORSO|STRADA|VICOLO|LARGO|PIAZZALE)\b",
    re.IGNORECASE
)


# -------------------- HELPERS --------------------

def normalize_line(line):
    return re.sub(r"\s+", " ", line.replace("\xa0", " ")).strip()


def italian_to_float(val):

    if not val or pd.isna(val):
        return 0.0

    val = str(val).strip()

    negative = val.endswith("-")

    val = val.replace(".", "").replace(",", ".").replace("-", "")

    try:
        num = float(val)
        return -num if negative else num
    except:
        return 0.0


# -------------------- FILE HANDLING --------------------

def open_pdf(pdf_path):

    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF")

    return fitz.open(pdf_path), pdf_path.with_suffix(".xlsx")


# -------------------- PAGE EXTRACTION --------------------

def extract_lines_from_page(doc, page_num):

    page = doc.load_page(page_num)

    text = page.get_text("text")

    lines = text.split("\n")

    return [normalize_line(l) for l in lines if l.strip()]


# -------------------- HEADER PARSER --------------------

def parse_header(line):

    match = HEADER_PATTERN.match(line)

    if not match:
        return None

    return {
        "N_progr": match.group("nprogr"),
        "DtCPU": match.group("dtcpu"),
        "N_doc": match.group("ndoc"),
        "DtReg": match.group("dtreg"),
        "DtDoc": match.group("dtdoc"),
        "TD": match.group("td"),
        "Testo_testata_doc": match.group("desc").strip()
    }


# -------------------- ACCOUNT PARSER --------------------

def parse_account_line(line):

    match = ACCOUNT_PATTERN.search(line)

    if not match:
        return None

    rdo = match.group("rdo")
    c = match.group("c")

    tokens = line.split()

    account = None

    # SAP account numbers are typically 10 digits
    accounts = [t for t in tokens if re.fullmatch(r"\d{8,12}", t)]

    if accounts:
        account = accounts[-1]
    else:
        account = match.group("account")

    amounts = AMOUNT_PATTERN.findall(line)

    imp_de = None
    dare = None
    avere = None
    div = None

    tokens = line.split()

    for t in tokens:
        if len(t) == 3 and t.isalpha():
            div = t
            break

    if len(amounts) == 3:
        imp_de, dare, avere = amounts

    elif len(amounts) == 2:
        dare, avere = amounts

    elif len(amounts) == 1:
        dare = amounts[0]

    account_desc = re.sub(r"\s+", " ", line[:match.start()]).strip()

    return {
        "Definizione_conto": account_desc,
        "RDO": rdo,
        "C": c,
        "N_conto": account,
        "imp_in_DE": imp_de,
        "DIV": div,
        "Imp_in_dare_DI": dare,
        "Imp_in_avere_DI": avere
    }


# -------------------- LINE PROCESSOR --------------------

def process_lines(lines, entry_data=None, entry_description=None):

    rows = []
    i = 0
    total_lines = len(lines)

    while i < total_lines:

        line = lines[i]

        try:

            # ---- SKIP PAGE / HEADER TEXT ----

            if ("Giornale bollato" in line
                or "Ledger" in line
                or "N.progr." in line
                or "Definizione conto" in line
                or "Tot. pagine" in line
                or "Accum." in line
                or line.startswith("Riporto")):

                i += 1
                continue

            if line.startswith(">"):
                i += 1
                continue

            if LOCATION_PATTERN.match(line) or ADDRESS_PATTERN.match(line):
                i += 1
                continue


            # ---- HEADER DETECTION ----

            header = parse_header(line)

            # DEBUG header detection
            if "10934" in line or "12122" in line:
                print("HEADER TEST LINE:", repr(line))
                print("HEADER MATCH:", header)

            if header:
                entry_data = header
                entry_description = header.get("Testo_testata_doc")
                i += 1
                continue


            # ---- MULTIPLE ACCOUNTS IN SAME LINE ----

            account_matches = ACCOUNT_PATTERN.findall(line)

            if len(account_matches) > 1:

                parts = re.split(
                    r"(?=\d{3,6}\s+[A-Z]\s+[A-Z0-9]{5,})",
                    line
                )

                for part in parts:

                    account = parse_account_line(part)

                    if account:
                        rows.append({
                            **(entry_data if entry_data else {
                                "N_progr": None,
                                "DtCPU":   None,
                                "N_doc":   None,
                                "DtReg":   None,
                                "DtDoc":   None,
                                "TD":      None,
                                "Testo_testata_doc": f"[CONTINUATION FROM PREVIOUS PDF]"
                            }),
                            **account
                        })
                i += 1
                continue


            # ---- SINGLE / MULTILINE ACCOUNT ----
            if not ACCOUNT_PATTERN.search(line):
                i += 1
                continue

            combined = line
            account = parse_account_line(combined)

            j = i

            while not account and j + 1 < total_lines:

                next_line = lines[j + 1]

                # STOP if new header encountered
                if parse_header(next_line):
                    break

                if ("Tot. pagine" in next_line
                    or "Accum." in next_line
                    or next_line.startswith("Riporto")
                    or "Ledger" in next_line):
                    break

                combined += " " + next_line
                account = parse_account_line(combined)

                j += 1

                if j - i > 20:
                    break


            if account:
                rows.append({
                    **(entry_data if entry_data else {
                        "N_progr": None,
                        "DtCPU":   None,
                        "N_doc":   None,
                        "DtReg":   None,
                        "DtDoc":   None,
                        "TD":      None,
                        "Testo_testata_doc": f"[CONTINUATION FROM PREVIOUS PDF]"
                    }),
                    **account
                })

                i = j

        except Exception:

            print("\nLINE PARSING ERROR")
            print(f"Line index: {i}")
            print(f"Line content: {line}")
            traceback.print_exc()

        i += 1

    return rows, entry_data, entry_description

# -------------------- SAVE EXCEL --------------------

def save_excel(rows, output_excel):

    df = pd.DataFrame(rows)

    column_order = [
        "N_progr",
        "DtCPU",
        "N_doc",
        "DtReg",
        "DtDoc",
        "TD",
        "Testo_testata_doc",
        "Definizione_conto",
        "RDO",
        "C",
        "N_conto",
        "imp_in_DE",
        "DIV",
        "Imp_in_dare_DI",
        "Imp_in_avere_DI"
    ]

    for col in column_order:
        if col not in df.columns:
            df[col] = None

    df = df[column_order]

    # ---------------- N_PROGR DEBUG ----------------

    df["N_progr"] = pd.to_numeric(df["N_progr"], errors="coerce")

    prev = df["N_progr"].shift(1)

    gap = df["N_progr"] - prev

    df["N_progr_missing"] = ""

    mask = gap > 1

    df.loc[mask, "N_progr_missing"] = (
        prev[mask].astype(int).astype(str)
        + "→"
        + df.loc[mask, "N_progr"].astype(int).astype(str)
    )

    # -----------------------------------------------

    # df.to_excel(output_excel, index=False)
    df.to_excel(output_excel, index=False, engine="xlsxwriter")

    print(f"\nTotal rows extracted: {len(df)}")
    print(f"Excel saved to: {output_excel}")

    return df

# -------------------- LEDGER AUDIT --------------------
def audit_ledger(df, pdf_number):
    
    continuation_rows = df[df["N_progr"].isna()]
    if not continuation_rows.empty:
        cont_dare  = continuation_rows["Imp_in_dare_DI"].apply(italian_to_float).sum()
        cont_avere = continuation_rows["Imp_in_avere_DI"].apply(italian_to_float).sum()
        print(f"\nContinuation rows (no header): {len(continuation_rows)}")
        print(f"  Dare  : {cont_dare:,.2f}")
        print(f"  Avere : {cont_avere:,.2f}")

    df_numbered = df[df["N_progr"].notna()].copy()
    df_numbered["N_progr"] = pd.to_numeric(df_numbered["N_progr"], errors="coerce")

    print("\n---------------- LEDGER AUDIT ----------------")

    if df.empty:
        print("No rows extracted.")
        return

    df["N_progr"] = pd.to_numeric(df["N_progr"], errors="coerce")

    first_nprog = int(df["N_progr"].min())
    last_nprog = int(df["N_progr"].max())

    print(f"\nFirst N_progr : {first_nprog}")
    print(f"Last  N_progr : {last_nprog}")

    expected = set(range(first_nprog, last_nprog + 1))
    extracted = set(df["N_progr"].dropna().astype(int))

    missing = sorted(expected - extracted)

    print(f"\nExpected entries : {len(expected)}")
    print(f"Extracted entries: {len(extracted)}")
    print(f"Missing entries  : {len(missing)}")

    if missing:
        print("\nFirst 20 missing N_progr:")
        print(missing[:20], len(missing))

    dare = df["Imp_in_dare_DI"].apply(italian_to_float).sum()
    avere = df["Imp_in_avere_DI"].apply(italian_to_float).sum()
    
    # Per-entry balance check
    df["dare_f"] = df["Imp_in_dare_DI"].apply(italian_to_float)
    df["avere_f"] = df["Imp_in_avere_DI"].apply(italian_to_float)

    entry_totals = df.groupby("N_progr")[["dare_f", "avere_f"]].sum()
    entry_totals["diff"] = entry_totals["dare_f"] - entry_totals["avere_f"]

    unbalanced = entry_totals[entry_totals["diff"].abs() > 0.01]
    print(f"\nUnbalanced entries: {len(unbalanced)}")
    print(unbalanced.head(20))

    print("\n------------- TOTALS -------------")
    print(f"Total Dare  : {dare:,.2f}")
    print(f"Total Avere : {avere:,.2f}")
    print(f"Difference  : {(dare - avere):,.2f}")
    print("----------------------------------\n")


# -------------------- MAIN DRIVER --------------------

def ledger_pdf_to_excel(pdf_path, pdf_number):

    try:

        doc, output_excel = open_pdf(pdf_path)

        start_page = 1 if pdf_number == 1 else 0
        max_pages = len(doc)

        print(f"\nProcessing pages {start_page + 1} to {max_pages}")

        # ------------------------------------------------
        # COLLECT ALL LINES FROM THE PDF FIRST
        # ------------------------------------------------

        all_lines = []

        for page_num in range(start_page, max_pages):

            try:

                lines = extract_lines_from_page(doc, page_num)

                all_lines.extend(lines)

            except Exception:

                print("\nPAGE READ ERROR")
                print(f"Page number: {page_num + 1}")
                traceback.print_exc()

        print(f"\nTotal lines collected: {len(all_lines)}")

        # ------------------------------------------------
        # PROCESS ALL LINES AS ONE CONTINUOUS STREAM
        # ------------------------------------------------

        entry_data = None
        entry_description = None

        all_rows, entry_data, entry_description = process_lines(
            all_lines,
            entry_data,
            entry_description
        )

        # ------------------------------------------------

        df = save_excel(all_rows, output_excel)

        audit_ledger(df,pdf_number)

        return df

    except Exception:

        print("\nPROGRAM FAILED")
        traceback.print_exc()

# -------------------- ENTRY POINT --------------------

if __name__ == "__main__":

    try:

        pdf_path = input("\nEnter PDF file path: ").strip().strip('"')

        pdf_number = int(input("Enter PDF number (1,2,3... if split PDFs): ").strip())

        ledger_pdf_to_excel(pdf_path, pdf_number)

    except Exception:

        print("\nPROGRAM FAILED")
        traceback.print_exc()