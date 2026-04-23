import argparse
import gzip
import io
import json
import os
import re
import sys
import time
from urllib.parse import quote_plus

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None

try:
    import requests
except ImportError:
    requests = None
    import urllib.request

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DEFAULT_TIMEOUT = 20
TOTALBESTAND_FILE = "roller_totalbestand.json.gz"
TOTALBESTAND_URL = "https://data.brreg.no/enhetsregisteret/api/roller/totalbestand"


def http_get_json(url, timeout=DEFAULT_TIMEOUT):
    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    if requests is not None:
        response = requests.get(url, timeout=timeout, headers=headers)
        response.raise_for_status()
        return response.json()
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def normalize_name(name):
    """Normaliser et navn for sammenligning."""
    return " ".join(name.upper().split())


def build_full_name(person):
    """Bygg fullt navn fra Brreg person-objekt."""
    navn = (person or {}).get("navn", {})
    parts = [
        (navn.get("fornavn") or "").strip(),
        (navn.get("mellomnavn") or "").strip(),
        (navn.get("etternavn") or "").strip(),
    ]
    return " ".join(p for p in parts if p)


def name_parts_match(search_name, candidate_name):
    """Sjekk om navnene er identiske (samme deler, uavhengig av rekkefølge)."""
    search_parts = set(search_name.upper().split())
    candidate_parts = set(candidate_name.upper().split())
    return search_parts and search_parts == candidate_parts


def safe(value):
    return "" if value is None else str(value).strip()


def normalize_space(text):
    return re.sub(r"\s+", " ", text).strip()


def http_get_text(url, timeout=DEFAULT_TIMEOUT):
    headers = {"User-Agent": "Mozilla/5.0"}
    if requests is not None:
        response = requests.get(url, timeout=timeout, headers=headers)
        response.raise_for_status()
        return response.text
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# Kommunenummer -> Fylke mapping
# ---------------------------------------------------------------------------

_KOMMUNENR_PREFIX_TO_FYLKE = {
    "03": "Oslo",
    "11": "Rogaland",
    "15": "Møre og Romsdal",
    "18": "Nordland",
    "31": "Østfold",
    "32": "Akershus",
    "33": "Buskerud",
    "34": "Innlandet",
    "39": "Vestfold",
    "40": "Telemark",
    "42": "Agder",
    "46": "Vestland",
    "50": "Trøndelag",
    "55": "Troms",
    "56": "Finnmark",
}


def _kommunenr_to_fylke(kommunenr):
    if not kommunenr or len(kommunenr) < 2:
        return ""
    return _KOMMUNENR_PREFIX_TO_FYLKE.get(kommunenr[:2], "")


# ---------------------------------------------------------------------------
# Hent selskapsdetaljer fra Brreg API + Proff (som i Brreg_Proff_fallback)
# ---------------------------------------------------------------------------

def _extract_address_from_brreg(data):
    """Hent adresse, postnr, poststed, fylke fra Brreg enhet-data."""
    adr = data.get("forretningsadresse") or data.get("postadresse") or {}
    adresse_lines = adr.get("adresse", [])
    adresse = ", ".join(a for a in adresse_lines if a) if adresse_lines else ""
    postnr = safe(adr.get("postnummer"))
    poststed = safe(adr.get("poststed"))
    kommunenr = safe(adr.get("kommunenummer"))
    fylke = _kommunenr_to_fylke(kommunenr)
    return adresse, postnr, poststed, fylke


def _extract_brreg_kontaktperson(roller_data):
    """Hent kontaktperson (KONT/DAGL/STYR) og regnskapsfører fra Brreg roller-data."""
    candidates = []
    regnskapsforer = ""

    for gruppe in roller_data.get("rollegrupper", []) or []:
        gruppe_kode = safe((gruppe.get("type") or {}).get("kode"))

        for rolle in gruppe.get("roller", []) or []:
            if rolle.get("fratraadt") is True or rolle.get("avregistrert") is True:
                continue

            # Regnskapsfører
            if gruppe_kode == "REGN" and not regnskapsforer:
                enhet = rolle.get("enhet")
                if enhet:
                    navn_list = enhet.get("navn", [])
                    if navn_list:
                        regnskapsforer = " ".join(navn_list).strip()
                person = rolle.get("person")
                if person and not regnskapsforer:
                    n = build_full_name(person)
                    if n:
                        regnskapsforer = n

            person = rolle.get("person")
            if not person:
                continue
            navn = build_full_name(person)
            if not navn:
                continue

            rolle_kode = safe((rolle.get("type") or {}).get("kode"))
            if rolle_kode == "KONT":
                candidates.append(("KONT", navn))
            elif rolle_kode == "DAGL":
                candidates.append(("DAGL", navn))
            elif rolle_kode == "LEDE" and gruppe_kode == "STYRE":
                candidates.append(("STYR", navn))

    # Prioritet: KONT > DAGL > STYR
    kontaktperson, rolle_kode = "", ""
    for wanted in ["KONT", "DAGL", "STYR"]:
        for rc, n in candidates:
            if rc == wanted:
                kontaktperson, rolle_kode = n, rc
                break
        if kontaktperson:
            break

    return kontaktperson, rolle_kode, regnskapsforer


def _extract_phone_from_html(html):
    text = normalize_space(re.sub(r"<[^>]+>", " ", html))
    match = re.search(r"Telefon\s*([\d\s]{8,15})", text)
    if match:
        return normalize_space(match.group(1))
    return ""


def _fetch_proff_data(orgnr, timeout, browser_page=None):
    """Hent telefon, KBPS og SIV fra Proff via Playwright (søk -> profil -> regnskap)."""
    telefon, kbps, siv = "", None, None
    if browser_page is None:
        return telefon, kbps, siv
    try:
        # Steg 1: Søk etter selskapet for å finne profil-URL
        search_url = f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}"
        browser_page.goto(search_url, wait_until="networkidle", timeout=timeout * 1000)
        html = browser_page.content()
        m = re.search(r'href="(/selskap/[^"]+)"', html)
        if not m:
            return telefon, kbps, siv
        profile_path = m.group(1)

        # Steg 2: Gå til profil-siden for å hente telefon
        profile_url = "https://www.proff.no" + profile_path
        browser_page.goto(profile_url, wait_until="networkidle", timeout=timeout * 1000)
        html = browser_page.content()
        telefon = _extract_phone_from_html(html)

        # Steg 3: Gå til regnskap-siden for KBPS og SIV
        regnskap_url = "https://www.proff.no" + profile_path.replace("/selskap/", "/regnskap/")
        browser_page.goto(regnskap_url, wait_until="networkidle", timeout=timeout * 1000)
        html = browser_page.content()

        for m in re.finditer(r'"code"\s*:\s*"(KBPS|SIV)"\s*,\s*"amount"\s*:\s*"([^"]*?)"', html):
            code, amount = m.group(1), m.group(2)
            try:
                val = int(amount) * 1000
            except (ValueError, TypeError):
                val = None
            if code == "KBPS" and kbps is None:
                kbps = val
            elif code == "SIV" and siv is None:
                siv = val
            if kbps is not None and siv is not None:
                break
    except Exception:
        pass
    return telefon, kbps, siv


def _fetch_brreg_regnskap(orgnr, timeout):
    """Hent regnskapsdata fra Brreg regnskapsregisteret."""
    try:
        url = f"https://data.brreg.no/regnskapsregisteret/regnskap/{orgnr}"
        data = http_get_json(url, timeout=timeout)
        if not data or not isinstance(data, list) or len(data) == 0:
            return None, None
        siste = data[0]
        resultat = siste.get("resultatregnskapResultat", {})
        driftsinntekter = (resultat.get("driftsresultat") or {}).get("driftsinntekter", {}).get("sumDriftsinntekter")
        aarsresultat = resultat.get("aarsresultat")
        return driftsinntekter, aarsresultat
    except Exception:
        return None, None


def fetch_all_details(orgnr, timeout=DEFAULT_TIMEOUT, browser_page=None):
    """Hent all info om et selskap: navn, adresse, fylke, kontaktperson, telefon, regnskapsfører, regnskap."""
    details = {
        "navn": "", "adresse": "", "postnr": "", "poststed": "", "fylke": "",
        "kontaktperson": "", "kontaktperson_rolle": "", "regnskapsforer": "",
        "telefon": "", "driftsinntekter": None, "aarsresultat": None,
        "sum_kasse_bank": None, "sum_investeringer": None,
    }

    # 1. Hent enhet-data (navn, adresse, fylke)
    try:
        url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}"
        enhet_data = http_get_json(url, timeout=timeout)
        details["navn"] = enhet_data.get("navn", "")
        adresse, postnr, poststed, fylke = _extract_address_from_brreg(enhet_data)
        details["adresse"] = adresse
        details["postnr"] = postnr
        details["poststed"] = poststed
        details["fylke"] = fylke
    except Exception:
        pass

    # 2. Hent roller (kontaktperson, regnskapsfører)
    try:
        url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}/roller"
        roller_data = http_get_json(url, timeout=timeout)
        kp, kp_rolle, regnf = _extract_brreg_kontaktperson(roller_data)
        details["kontaktperson"] = kp
        details["kontaktperson_rolle"] = kp_rolle
        details["regnskapsforer"] = regnf
    except Exception:
        pass

    # 3. Hent telefon, KBPS og SIV fra Proff (via Playwright)
    telefon, kbps, siv = _fetch_proff_data(orgnr, timeout, browser_page)
    details["telefon"] = telefon
    details["sum_kasse_bank"] = kbps
    details["sum_investeringer"] = siv

    # 4. Hent regnskap fra Brreg regnskapsregisteret
    details["driftsinntekter"], details["aarsresultat"] = _fetch_brreg_regnskap(orgnr, timeout)

    return details


def enrich_results_with_details(results, person_names):
    """Berik resultater med full info fra Brreg API og Proff."""
    unique_orgnr = set()
    for name in person_names:
        for comp in results.get(name, []):
            if comp.get("orgnr"):
                unique_orgnr.add(comp["orgnr"])

    if not unique_orgnr:
        return

    print(f"Henter detaljer for {len(unique_orgnr)} selskap(er) fra Brreg/Proff...")

    # Start Playwright-nettleser for Proff-scraping
    pw_context = None
    browser = None
    browser_page = None
    if sync_playwright is not None:
        try:
            pw_context = sync_playwright().__enter__()
            browser = pw_context.chromium.launch(headless=True)
            browser_page = browser.new_page()
            print("  (Playwright-nettleser startet for Proff-henting)")
        except Exception as e:
            print(f"  Advarsel: Kunne ikke starte Playwright ({e}), hopper over Proff-regnskap")
    else:
        print("  Advarsel: Playwright ikke installert, hopper over Proff-regnskap")

    details_cache = {}
    for i, orgnr in enumerate(unique_orgnr, 1):
        print(f"\r  {i}/{len(unique_orgnr)}: {orgnr}", end="", flush=True)
        details_cache[orgnr] = fetch_all_details(orgnr, browser_page=browser_page)
        time.sleep(0.3)
    print()

    # Lukk Playwright
    if browser is not None:
        try:
            browser.close()
        except Exception:
            pass
    if pw_context is not None:
        try:
            pw_context.__exit__(None, None, None)
        except Exception:
            pass

    for name in person_names:
        for comp in results.get(name, []):
            orgnr = comp.get("orgnr", "")
            det = details_cache.get(orgnr, {})
            if det.get("navn"):
                comp["selskap"] = det["navn"]
            comp["adresse"] = det.get("adresse", "")
            comp["postnr"] = det.get("postnr", "")
            comp["poststed"] = det.get("poststed", "")
            comp["fylke"] = det.get("fylke", "")
            comp["kontaktperson"] = det.get("kontaktperson", "")
            comp["kontaktperson_rolle"] = det.get("kontaktperson_rolle", "")
            comp["regnskapsforer"] = det.get("regnskapsforer", "")
            comp["telefon"] = det.get("telefon", "")
            comp["driftsinntekter"] = det.get("driftsinntekter")
            comp["aarsresultat"] = det.get("aarsresultat")
            comp["sum_kasse_bank"] = det.get("sum_kasse_bank")
            comp["sum_investeringer"] = det.get("sum_investeringer")


# ---------------------------------------------------------------------------
# Last ned og søk i Brreg roller totalbestand
# ---------------------------------------------------------------------------

def download_totalbestand(filepath, timeout=300):
    """Last ned roller totalbestand fra Brreg (ca 132 MB)."""
    print(f"Laster ned roller-totalbestand fra Brreg ({TOTALBESTAND_URL})...")
    print("  (dette kan ta et par minutter, ca 132 MB)")

    if requests is not None:
        response = requests.get(
            TOTALBESTAND_URL,
            headers={"User-Agent": "Mozilla/5.0"},
            stream=True,
            timeout=timeout,
        )
        response.raise_for_status()
        total = int(response.headers.get("content-length", 0))

        with open(filepath, "wb") as f:
            downloaded = 0
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                f.write(chunk)
                downloaded += len(chunk)
                if total > 0:
                    pct = downloaded * 100 // total
                    print(f"\r  Lastet ned: {downloaded // (1024*1024)} MB / {total // (1024*1024)} MB ({pct}%)", end="", flush=True)
                else:
                    print(f"\r  Lastet ned: {downloaded // (1024*1024)} MB", end="", flush=True)
        print()
    else:
        req = urllib.request.Request(TOTALBESTAND_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            with open(filepath, "wb") as f:
                while True:
                    chunk = resp.read(1024 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)

    print(f"  Lagret: {filepath} ({os.path.getsize(filepath) // (1024*1024)} MB)")


def search_totalbestand(filepath, person_names):
    """Søk gjennom roller-totalbestand for å finne selskaper knyttet til personene.

    Returnerer en dict: {person_name: [liste av treff]}
    """
    # Lag normaliserte søkenavn -> original navn mapping
    search_map = {}
    for name in person_names:
        norm = normalize_name(name)
        search_map[norm] = name

    # Lag et sett med alle unike etternavn for rask filtrering
    last_names = set()
    for name in person_names:
        parts = name.upper().split()
        if parts:
            last_names.add(parts[-1])

    results = {name: [] for name in person_names}
    total_enheter = 0

    print(f"Søker gjennom roller-totalbestand for {len(person_names)} person(er)...")

    with gzip.open(filepath, "rt", encoding="utf-8") as f:
        # Filen er en JSON-array med enheter. Vi parser den som en strøm.
        # Prøv først å laste hele filen (raskere for denne størrelsen)
        print("  Leser og parser JSON (dette tar litt tid)...")
        data = json.load(f)

    print(f"  Totalt {len(data)} enheter i totalbestand")

    iterator = tqdm(data, desc="  Søker") if tqdm is not None else data

    for enhet in iterator:
        total_enheter += 1
        orgnr = str(enhet.get("organisasjonsnummer", ""))
        enhet_navn = enhet.get("navn", "")

        for gruppe in enhet.get("rollegrupper", []) or []:
            gruppe_type = (gruppe.get("type") or {}).get("beskrivelse", "")

            for rolle in gruppe.get("roller", []) or []:
                if rolle.get("fratraadt") or rolle.get("avregistrert"):
                    continue

                person = rolle.get("person")
                if not person:
                    continue

                full_name = build_full_name(person)
                if not full_name:
                    continue

                # Rask filtrering: sjekk om etternavnet matcher
                name_upper = full_name.upper()
                if not any(ln in name_upper for ln in last_names):
                    continue

                # Detaljert sjekk mot hvert søkenavn
                for search_norm, original_name in search_map.items():
                    if name_parts_match(original_name, full_name):
                        rolle_type = (rolle.get("type") or {}).get("beskrivelse", "")
                        results[original_name].append({
                            "orgnr": orgnr,
                            "selskap": enhet_navn,
                            "rolle": rolle_type,
                            "gruppe": gruppe_type,
                            "person_i_brreg": full_name,
                            "kilde": "Brreg",
                        })

    print(f"  Ferdig! Søkte gjennom {total_enheter} enheter")
    return results


# ---------------------------------------------------------------------------
# Hovedprogram
# ---------------------------------------------------------------------------

def set_hyperlink(cell, text, url):
    cell.value = text
    cell.hyperlink = url
    cell.style = "Hyperlink"


def main():
    parser = argparse.ArgumentParser(
        description="Sjekk om personer i en XLSX-fil er registrert som bedriftseiere i Brønnøysundregistrene."
    )
    parser.add_argument(
        "--input", default="personer.xlsx",
        help="Input Excel-fil med personnavn (standard: personer.xlsx)",
    )
    parser.add_argument(
        "--output", default="bedriftseiere_resultat.xlsx",
        help="Output Excel-fil (standard: bedriftseiere_resultat.xlsx)",
    )
    parser.add_argument(
        "--cache", default=TOTALBESTAND_FILE,
        help="Cache-fil for roller-totalbestand (standard: roller_totalbestand.json.gz)",
    )
    parser.add_argument(
        "--force-download", action="store_true",
        help="Last ned totalbestand på nytt selv om cache finnes",
    )
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Fant ikke input-filen: {args.input}", file=sys.stderr)
        sys.exit(1)

    # Les personnavn fra input-fil
    wb = load_workbook(args.input)
    ws = wb.active
    person_names = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0] if row and len(row) > 0 and row[0] else "").strip()
        if name:
            person_names.append(name)

    print(f"Fant {len(person_names)} person(er) i {args.input}")
    for n in person_names:
        print(f"  - {n}")

    # Last ned totalbestand hvis nødvendig (auto-oppdater etter 30 dager)
    if args.force_download or not os.path.exists(args.cache):
        download_totalbestand(args.cache)
    else:
        age_days = (time.time() - os.path.getmtime(args.cache)) / 86400
        if age_days > 30:
            print(f"Cache er {age_days:.0f} dager gammel (>30 dager), laster ned på nytt...")
            download_totalbestand(args.cache)
        else:
            print(f"Bruker cached totalbestand: {args.cache} ({age_days:.0f} dager gammel)")

    # Søk gjennom totalbestand
    results = search_totalbestand(args.cache, person_names)

    # Berik med adresse og fylke fra Brreg API
    enrich_results_with_details(results, person_names)

    # Opprett output Excel
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Bedriftseiere"

    # Header-stil
    hf = Font(bold=True, color="FFFFFF", size=11)
    hfill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    ha = Alignment(horizontal="center", vertical="center")
    tb = Border(bottom=Side(style="thin", color="DDDDDD"))

    headers = [
        "Person",
        "Selskap",
        "Org.nr",
        "Rolle",
        "Rollegruppe",
        "Navn i Brreg",
        "Kontaktperson",
        "Telefon",
        "Adresse",
        "Postnr",
        "Poststed",
        "Fylke",
        "Regnskapsfører",
        "Driftsinntekter",
        "Årsresultat",
        "Sum Kasse/Bank/Post",
        "Sum investeringer",
        "Proff",
        "Brreg",
        "1881",
        "LinkedIn",
    ]
    for col, h in enumerate(headers, 1):
        c = out_ws.cell(row=1, column=col, value=h)
        c.font = hf
        c.fill = hfill
        c.alignment = ha

    out_ws.freeze_panes = "B2"

    row_num = 2
    total_companies = 0
    persons_with_companies = 0

    for name in person_names:
        companies = results.get(name, [])
        total_companies += len(companies)
        if companies:
            persons_with_companies += 1

        if companies:
            for comp in companies:
                out_ws.cell(row=row_num, column=1, value=name).border = tb
                out_ws.cell(row=row_num, column=2, value=comp["selskap"]).border = tb
                out_ws.cell(row=row_num, column=3, value=comp["orgnr"]).border = tb
                out_ws.cell(row=row_num, column=4, value=comp["rolle"]).border = tb
                out_ws.cell(row=row_num, column=5, value=comp["gruppe"]).border = tb
                out_ws.cell(row=row_num, column=6, value=comp["person_i_brreg"]).border = tb
                out_ws.cell(row=row_num, column=7, value=comp.get("kontaktperson", "")).border = tb
                out_ws.cell(row=row_num, column=8, value=comp.get("telefon", "")).border = tb
                out_ws.cell(row=row_num, column=9, value=comp.get("adresse", "")).border = tb
                out_ws.cell(row=row_num, column=10, value=comp.get("postnr", "")).border = tb
                out_ws.cell(row=row_num, column=11, value=comp.get("poststed", "")).border = tb
                out_ws.cell(row=row_num, column=12, value=comp.get("fylke", "")).border = tb
                out_ws.cell(row=row_num, column=13, value=comp.get("regnskapsforer", "")).border = tb

                # Regnskap med formatering
                ACCT_FMT = '#,##0 kr;-#,##0 kr'
                driftsinnt = comp.get("driftsinntekter")
                aarsres = comp.get("aarsresultat")
                c14 = out_ws.cell(row=row_num, column=14, value=driftsinnt)
                c14.border = tb
                if driftsinnt is not None:
                    c14.number_format = ACCT_FMT
                c15 = out_ws.cell(row=row_num, column=15, value=aarsres)
                c15.border = tb
                if aarsres is not None:
                    c15.number_format = ACCT_FMT

                kbps = comp.get("sum_kasse_bank")
                siv = comp.get("sum_investeringer")
                c16 = out_ws.cell(row=row_num, column=16, value=kbps)
                c16.border = tb
                if kbps is not None:
                    c16.number_format = ACCT_FMT
                c17 = out_ws.cell(row=row_num, column=17, value=siv)
                c17.border = tb
                if siv is not None:
                    c17.number_format = ACCT_FMT

                orgnr = comp["orgnr"]
                if orgnr:
                    set_hyperlink(
                        out_ws.cell(row=row_num, column=18),
                        "Proff",
                        f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}",
                    )
                    out_ws.cell(row=row_num, column=18).border = tb

                    set_hyperlink(
                        out_ws.cell(row=row_num, column=19),
                        "Brreg",
                        f"https://virksomhet.brreg.no/nb/oppslag/enheter/{orgnr}",
                    )
                    out_ws.cell(row=row_num, column=19).border = tb

                kontaktperson = comp.get("kontaktperson", "")
                if kontaktperson:
                    q = quote_plus(kontaktperson)
                    set_hyperlink(out_ws.cell(row=row_num, column=20), "1881", f"https://www.1881.no/?query={q}")
                    out_ws.cell(row=row_num, column=20).border = tb
                    set_hyperlink(out_ws.cell(row=row_num, column=21), "LinkedIn", f"https://www.linkedin.com/search/results/all/?keywords={q}")
                    out_ws.cell(row=row_num, column=21).border = tb

                row_num += 1
        else:
            out_ws.cell(row=row_num, column=1, value=name).border = tb
            c = out_ws.cell(row=row_num, column=2, value="Ingen treff")
            c.border = tb
            c.font = Font(italic=True, color="999999")
            for col in range(3, len(headers) + 1):
                out_ws.cell(row=row_num, column=col).border = tb

            # Legg til Proff rollesøk-lenke for manuell sjekk
            set_hyperlink(
                out_ws.cell(row=row_num, column=18),
                "Søk Proff",
                f"https://www.proff.no/rolles%C3%B8k?q={quote_plus(name)}",
            )
            out_ws.cell(row=row_num, column=18).border = tb
            row_num += 1

    # Auto-filter
    if row_num > 2:
        out_ws.auto_filter.ref = f"A1:U{row_num - 1}"

    # Kolonnebredder
    col_widths = {
        "A": 25, "B": 35, "C": 14, "D": 22, "E": 18, "F": 25, "G": 22,
        "H": 14, "I": 35, "J": 8, "K": 14, "L": 16, "M": 22,
        "N": 20, "O": 20, "P": 22, "Q": 20, "R": 10, "S": 10, "T": 10, "U": 12,
    }
    for letter, width in col_widths.items():
        out_ws.column_dimensions[letter].width = width

    # Oppsummerings-ark
    ws2 = out_wb.create_sheet("Oppsummering")
    ws2.cell(row=1, column=1, value="Søk etter bedriftseiere").font = Font(bold=True, size=14)
    ws2.cell(row=2, column=1, value=f"Input: {args.input}")
    ws2.cell(row=3, column=1, value="Kilde: Brønnøysundregistrenes roller-totalbestand")
    ws2.cell(row=5, column=1, value="Statistikk").font = Font(bold=True, size=12)

    stats = [
        ("Antall personer søkt:", len(person_names)),
        ("Personer med selskap:", persons_with_companies),
        ("Personer uten treff:", len(person_names) - persons_with_companies),
        ("Totalt antall roller funnet:", total_companies),
    ]
    for i, (lbl, val) in enumerate(stats):
        ws2.cell(row=6 + i, column=1, value=lbl)
        ws2.cell(row=6 + i, column=2, value=val).font = Font(bold=True)
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 10

    # Detaljer per person
    ws2.cell(row=11, column=1, value="Detaljer per person").font = Font(bold=True, size=12)
    for i, name in enumerate(person_names):
        count = len(results.get(name, []))
        ws2.cell(row=12 + i, column=1, value=name)
        ws2.cell(row=12 + i, column=2, value=f"{count} rolle(r)").font = Font(bold=True)

    out_wb.save(args.output)
    print(f"\nFerdig! Lagret: {args.output}")
    print(f"  Personer søkt:    {len(person_names)}")
    print(f"  Med selskap:      {persons_with_companies}")
    print(f"  Uten treff:       {len(person_names) - persons_with_companies}")
    print(f"  Roller funnet:    {total_companies}")

    try:
        os.startfile(args.output)
    except Exception:
        pass


if __name__ == "__main__":
    main()
