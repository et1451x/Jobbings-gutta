import re
from urllib.parse import quote_plus

try:
    import requests
except ImportError:
    requests = None
    import json
    import urllib.request


DEFAULT_TIMEOUT = 20

_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

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


def http_get_text(url, timeout=DEFAULT_TIMEOUT):
    if requests is not None:
        response = requests.get(url, timeout=timeout, headers=_HEADERS)
        response.raise_for_status()
        return response.text

    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def http_get_json(url, timeout=DEFAULT_TIMEOUT):
    headers = {**_HEADERS, "Accept": "application/json"}
    if requests is not None:
        response = requests.get(url, timeout=timeout, headers=headers)
        response.raise_for_status()
        return response.json()

    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def safe(value):
    return "" if value is None else str(value).strip()


def normalize_space(text):
    return re.sub(r"\s+", " ", text).strip()


def build_full_name(person):
    navn = (person or {}).get("navn", {})
    parts = [
        safe(navn.get("fornavn")),
        safe(navn.get("mellomnavn")),
        safe(navn.get("etternavn")),
    ]
    return " ".join([part for part in parts if part]).strip()


def extract_brreg_candidates(data):
    candidates = []

    for gruppe in data.get("rollegrupper", []) or []:
        gruppe_kode = safe((gruppe.get("type") or {}).get("kode"))

        for rolle in gruppe.get("roller", []) or []:
            if rolle.get("fratraadt") is True or rolle.get("avregistrert") is True:
                continue

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

    return candidates


def extract_regnskapsforer(data):
    for gruppe in data.get("rollegrupper", []) or []:
        if safe((gruppe.get("type") or {}).get("kode")) != "REGN":
            continue
        for rolle in gruppe.get("roller", []) or []:
            if rolle.get("fratraadt") is True or rolle.get("avregistrert") is True:
                continue
            enhet = rolle.get("enhet")
            if enhet:
                navn_list = enhet.get("navn", [])
                if navn_list:
                    return " ".join(navn_list).strip()
            person = rolle.get("person")
            if person:
                navn = build_full_name(person)
                if navn:
                    return navn
    return ""


def pick_primary_contact(candidates):
    for wanted in ["KONT", "DAGL", "STYR"]:
        for role_code, name in candidates:
            if role_code == wanted:
                return name, role_code
    return "", ""


def _extract_phone_from_html(html):
    text = normalize_space(re.sub(r"<[^>]+>", " ", html))
    match = re.search(r"Telefon\s*([\d\s]{8,15})", text)
    if match:
        return normalize_space(match.group(1))
    return ""


def _kommunenr_to_fylke(kommunenr):
    if not kommunenr or len(kommunenr) < 2:
        return ""
    return _KOMMUNENR_PREFIX_TO_FYLKE.get(kommunenr[:2], "")


def _extract_address_from_brreg(orgnr, timeout):
    url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}"
    data = http_get_json(url, timeout=timeout)
    adr = data.get("forretningsadresse") or data.get("postadresse") or {}
    adresse_lines = adr.get("adresse", [])
    adresse = ", ".join(line for line in adresse_lines if line) if adresse_lines else ""
    postnr = safe(adr.get("postnummer"))
    poststed = safe(adr.get("poststed"))
    kommunenr = safe(adr.get("kommunenummer"))
    fylke = _kommunenr_to_fylke(kommunenr)
    return adresse, postnr, poststed, fylke


def fetch_company_name(orgnr, timeout):
    try:
        url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}"
        data = http_get_json(url, timeout=timeout)
        return safe(data.get("navn"))
    except Exception:
        return ""


def fetch_proff_phone(orgnr, timeout):
    search_url = f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}"
    html = http_get_text(search_url, timeout=timeout)

    phone = _extract_phone_from_html(html)
    if phone:
        return phone

    profile_match = re.search(r'href="(/selskap/[^"]+)"', html)
    if profile_match:
        profile_url = "https://www.proff.no" + profile_match.group(1)
        try:
            profile_html = http_get_text(profile_url, timeout=timeout)
            phone = _extract_phone_from_html(profile_html)
            if phone:
                return phone
        except Exception:
            pass

    return ""


def fetch_from_brreg(orgnr, timeout):
    url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}/roller"
    data = http_get_json(url, timeout=timeout)
    candidates = extract_brreg_candidates(data)
    regnskapsforer = extract_regnskapsforer(data)
    return pick_primary_contact(candidates), regnskapsforer


def _extract_styreleder_from_html(html):
    link_match = re.search(
        r"styrets\s+leder\s*(?:<[^>]*>\s*)*<a[^>]*>\s*([^<]+?)\s*</a>",
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if link_match:
        name = normalize_space(link_match.group(1))
        if name and len(name) > 2:
            return name

    text = normalize_space(re.sub(r"<[^>]+>", " ", html))
    patterns = [
        r"Styrets leder\s+([A-ZÆØÅ][A-Za-zÆØÅæøåÉéÜüÖöÄä.\-\' ]+?)\s+(?:Adresse|\(f[\s\d])",
        r"Ledelse.administrasjon\s+Styrets leder\s+([A-ZÆØÅ][A-Za-zÆØÅæøåÉéÜüÖöÄä.\-\' ]+?)\s+\(",
        r"Styrets leder\s+([A-ZÆØÅ][A-Za-zÆØÅæøåÉéÜüÖöÄä.\-\' ]+?)\s+Kilde:\s*Brønnøysundregistrene",
        r"Styreleder\s+([A-ZÆØÅ][A-Za-zÆØÅæøåÉéÜüÖöÄä.\-\' ]+?)\s+Adresse",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            name = normalize_space(match.group(1))
            if name:
                return name

    return ""


def fetch_styreleder_from_proff(orgnr, timeout):
    search_url = f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}"
    html = http_get_text(search_url, timeout=timeout)

    name = _extract_styreleder_from_html(html)
    if name:
        return name, "STYR", html

    profile_match = re.search(r'href="(/selskap/[^"]+)"', html)
    if profile_match:
        profile_url = "https://www.proff.no" + profile_match.group(1)
        try:
            profile_html = http_get_text(profile_url, timeout=timeout)
            name = _extract_styreleder_from_html(profile_html)
            if name:
                return name, "STYR", profile_html
        except Exception:
            pass

    return "", "", html


def _find_proff_profile_path(orgnr, timeout):
    search_url = f"https://www.proff.no/bransjes%C3%B8k?q={orgnr}"
    html = http_get_text(search_url, timeout=timeout)
    match = re.search(r'href="(/selskap/[^"]+)"', html)
    return match.group(1) if match else ""


def fetch_proff_regnskap(orgnr, timeout):
    try:
        profile_path = _find_proff_profile_path(orgnr, timeout)
        if not profile_path:
            return None, None

        regnskap_url = "https://www.proff.no" + profile_path.replace("/selskap/", "/regnskap/")
        html = http_get_text(regnskap_url, timeout=timeout)

        kbps, siv = None, None
        for match in re.finditer(r'"code"\s*:\s*"(KBPS|SIV)"\s*,\s*"amount"\s*:\s*"([^"]*?)"', html):
            code, amount = match.group(1), match.group(2)
            try:
                value = int(amount)
            except (TypeError, ValueError):
                value = None
            if code == "KBPS" and kbps is None:
                kbps = value
            elif code == "SIV" and siv is None:
                siv = value
            if kbps is not None and siv is not None:
                break
        return kbps, siv
    except Exception:
        return None, None


def fetch_brreg_regnskap(orgnr, timeout):
    try:
        url = f"https://data.brreg.no/regnskapsregisteret/regnskap/{orgnr}"
        data = http_get_json(url, timeout=timeout)
        if not data or not isinstance(data, list):
            return None, None

        siste = data[0] if data else {}
        resultat = (siste or {}).get("resultatregnskapResultat") or {}
        driftsinntekter = (
            (resultat.get("driftsresultat") or {})
            .get("driftsinntekter", {})
            .get("sumDriftsinntekter")
        )
        aarsresultat = resultat.get("aarsresultat")
        return driftsinntekter, aarsresultat
    except Exception:
        return None, None


def fetch_primary_contact(orgnr, timeout):
    name, role, regnskapsforer = "", "", ""
    proff_html = None
    try:
        (name, role), regnskapsforer = fetch_from_brreg(orgnr, timeout)
    except Exception:
        pass

    if not name:
        try:
            name, role, proff_html = fetch_styreleder_from_proff(orgnr, timeout)
        except Exception:
            pass

    telefon, adresse, postnr, poststed, fylke = "", "", "", "", ""
    try:
        if proff_html:
            telefon = _extract_phone_from_html(proff_html)
        if not telefon:
            telefon = fetch_proff_phone(orgnr, timeout)
    except Exception:
        pass

    try:
        adresse, postnr, poststed, fylke = _extract_address_from_brreg(orgnr, timeout)
    except Exception:
        pass

    sum_kasse_bank, sum_investeringer = fetch_proff_regnskap(orgnr, timeout)

    return (
        name,
        role,
        telefon,
        adresse,
        postnr,
        poststed,
        fylke,
        regnskapsforer,
        sum_kasse_bank,
        sum_investeringer,
    )


def build_company_payload(orgnr, timeout=DEFAULT_TIMEOUT):
    normalized_orgnr = safe(orgnr)
    if not normalized_orgnr:
        raise ValueError("orgnr is required")

    selskapsnavn = fetch_company_name(normalized_orgnr, timeout)

    (
        navn,
        rolle,
        telefon,
        adresse,
        postnr,
        poststed,
        fylke,
        regnskapsforer,
        sum_kasse_bank,
        sum_investeringer,
    ) = fetch_primary_contact(normalized_orgnr, timeout)
    driftsinntekter, aarsresultat = fetch_brreg_regnskap(normalized_orgnr, timeout)

    person_query = quote_plus(navn) if navn else ""

    return {
        "selskapsnavn": selskapsnavn,
        "organisasjonsnummer": normalized_orgnr,
        "kontaktperson_navn": navn,
        "rolle": rolle,
        "telefon": telefon,
        "adresse": adresse,
        "postnr": postnr,
        "poststed": poststed,
        "fylke": fylke,
        "regnskapsforer": regnskapsforer,
        "driftsinntekter": driftsinntekter,
        "aarsresultat": aarsresultat,
        "sum_kasse_bank_post": sum_kasse_bank,
        "sum_investeringer": sum_investeringer,
        "links": {
            "proff": f"https://www.proff.no/bransjes%C3%B8k?q={normalized_orgnr}",
            "brreg": f"https://virksomhet.brreg.no/nb/oppslag/enheter/{normalized_orgnr}",
            "1881": f"https://www.1881.no/?query={person_query}" if person_query else "",
            "linkedin": f"https://www.linkedin.com/search/results/all/?keywords={person_query}" if person_query else "",
        },
    }