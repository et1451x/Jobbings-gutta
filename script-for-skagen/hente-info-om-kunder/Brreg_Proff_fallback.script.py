import argparse
import os
import sys
from urllib.parse import quote_plus

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

from openpyxl import Workbook, load_workbook
from brreg_proff_core import DEFAULT_TIMEOUT, fetch_primary_contact, safe


def set_hyperlink(cell, text, url):
    cell.value = text
    cell.hyperlink = url
    cell.style = "Hyperlink"


def open_output_file(path):
    if sys.platform.startswith("win") and hasattr(os, "startfile"):
        os.startfile(path)


def main():
    parser = argparse.ArgumentParser(
        description="Henter kontaktperson fra Brreg, med fallback til styreleder fra Proff."
    )
    parser.add_argument("--input", required=True, help="Input Excel-fil")
    parser.add_argument("--output", required=True, help="Output Excel-fil")
    parser.add_argument("--limit", type=int, default=None, help="Prosesser kun de første N radene")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Timeout i sekunder")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Fant ikke input-filen: {args.input}", file=sys.stderr)
        sys.exit(1)

    wb = load_workbook(args.input)
    ws = wb.active

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if args.limit is not None:
        rows = rows[:args.limit]

    iterator = tqdm(rows, desc="Prosesserer") if tqdm is not None else rows

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Resultat"

    out_ws.append([
        "Selskap",
        "Organisasjonsnummer",
        "Kontaktperson navn",
        "Rolle",
        "Telefon",
        "Adresse",
        "Postnr",
        "Poststed",
        "Fylke",
        "Regnskapsfører",
        "Sum Kasse/Bank/Post",
        "Sum investeringer",
        "Proff",
        "1881",
        "LinkedIn",
    ])
    out_ws.freeze_panes = "B2"

    total = len(rows)

    for i, row in enumerate(iterator, start=1):
        selskap = safe(row[0] if len(row) > 0 else "")
        orgnr = safe(row[1] if len(row) > 1 else "")

        navn, rolle, telefon, adresse, postnr, poststed, fylke, regnskapsforer = "", "", "", "", "", "", "", ""
        sum_kasse_bank, sum_investeringer = None, None

        if orgnr:
            navn, rolle, telefon, adresse, postnr, poststed, fylke, regnskapsforer, sum_kasse_bank, sum_investeringer = fetch_primary_contact(orgnr, args.timeout)

        out_ws.append([selskap, orgnr, navn, rolle, telefon, adresse, postnr, poststed, fylke, regnskapsforer, sum_kasse_bank, sum_investeringer, "Proff", "1881", "LinkedIn"])
        r = out_ws.max_row

        # Format financial columns as accounting
        ACCT_FMT = '#,##0 kr;-#,##0 kr'
        for col in (11, 12):
            cell = out_ws.cell(r, col)
            if cell.value is not None:
                cell.number_format = ACCT_FMT

        if orgnr:
            set_hyperlink(
                out_ws.cell(r, 13),
                "Proff",
                f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}",
            )

        if navn:
            q = quote_plus(navn)
            set_hyperlink(out_ws.cell(r, 14), "1881", f"https://www.1881.no/?query={q}")
            set_hyperlink(
                out_ws.cell(r, 15),
                "LinkedIn",
                f"https://www.linkedin.com/search/results/all/?keywords={q}",
            )

        if tqdm is None:
            print(f"{i}/{total} ferdig")

    # Enable auto-filter on header row
    out_ws.auto_filter.ref = out_ws.dimensions

    # Auto-fit column widths
    for col in out_ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(val))
        out_ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

    out_wb.save(args.output)
    print("Ferdig:", args.output)
    open_output_file(args.output)


if __name__ == "__main__":
    main()
