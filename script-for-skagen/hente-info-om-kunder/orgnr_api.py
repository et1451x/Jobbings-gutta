import argparse
import json
import re
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from brreg_proff_core import DEFAULT_TIMEOUT, build_company_payload


def parse_orgnrs_from_query(query):
    values = query.get("orgnr") or []
    candidates = []
    for raw in values:
        candidates.extend(re.split(r"[\s,;]+", raw.strip()))

    seen = set()
    orgnrs = []
    for value in candidates:
        orgnr = value.strip()
        if not orgnr or orgnr in seen:
            continue
        seen.add(orgnr)
        orgnrs.append(orgnr)
    return orgnrs


def build_payloads(orgnrs, timeout):
    payloads = []
    errors = []
    for orgnr in orgnrs:
        try:
            payloads.append(build_company_payload(orgnr, timeout=timeout))
        except Exception as exc:
            errors.append({"organisasjonsnummer": orgnr, "error": str(exc)})
    return payloads, errors


def format_value(value):
    if value is None or value == "":
        return "-"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, int):
        return f"{value:,}".replace(",", " ")
    return str(value)


STOREBRAND_LOGO_URL = "https://www.storebrand.no/privat/_/image/ac18e786-7254-42f5-9e43-bbb49c12ba18:2ab54e30cc57119ad83c5167947122cf03074110/block-1200-630/Storebrand%20logo%20beige_700x700.png"
SKAGEN_LOGO_URL = "https://www.skagenfondene.no/globalassets/skagen-funds/logos/skagen-logos/skagen_logo_web_no-logo-color.svg"


def build_index_html():
    example_orgnr = escape("988742163,979484534")
    return f"""<!DOCTYPE html>
<html lang=\"no\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Org.nr-oppslag</title>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
  <link href=\"https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Manrope:wght@400;500;700;800&display=swap\" rel=\"stylesheet\">
  <style>
    :root {{
      color-scheme: light;
      --bg: #edf4f2;
      --panel: #ffffff;
      --ink: #12322f;
      --muted: #4b6763;
      --accent: #007b5f;
      --accent-strong: #00644d;
      --accent-soft: #dff1ea;
      --border: #c5ddd5;
      --gold: #b78a58;
      --surface-deep: #0f3f36;
      --surface-deep-2: #0c322b;
    }}
    body {{
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at 0% 0%, rgba(0, 123, 95, 0.22), transparent 34%),
        radial-gradient(circle at 100% 100%, rgba(183, 138, 88, 0.16), transparent 36%),
        linear-gradient(145deg, #e4f0eb 0%, var(--bg) 55%, #f5faf8 100%);
      color: var(--ink);
      display: grid;
      place-items: center;
      padding: 24px;
      font-family: 'Manrope', sans-serif;
    }}
    main {{
      width: min(1040px, 100%);
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 26px;
      box-shadow: 0 24px 64px rgba(11, 49, 43, 0.14);
      padding: 30px;
      position: relative;
      overflow: hidden;
      animation: rise-in 460ms ease-out both;
    }}
    main::before {{
      content: "";
      position: absolute;
      inset: 0;
      pointer-events: none;
      background-image: linear-gradient(120deg, rgba(255, 255, 255, 0) 30%, rgba(0, 123, 95, 0.05));
    }}
    .brandbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 20px;
      padding-bottom: 16px;
      border-bottom: 1px solid var(--border);
      position: relative;
      z-index: 1;
    }}
    .brandstack {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }}
    .logo-chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      border: 1px solid var(--border);
      background: #fff;
      border-radius: 999px;
      padding: 7px 11px;
      box-shadow: 0 5px 18px rgba(0, 0, 0, 0.06);
      font-size: 0.85rem;
      color: var(--muted);
      animation: fade-up 420ms ease both;
    }}
    .logo-chip img {{
      display: block;
      max-height: 22px;
      width: auto;
    }}
    .kicker {{
      font-size: 0.76rem;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--accent);
      font-weight: 800;
    }}
    .hero {{
      display: grid;
      grid-template-columns: 1.05fr 0.95fr;
      gap: 20px;
      align-items: stretch;
      position: relative;
      z-index: 1;
    }}
    .hero-copy h1 {{
      margin: 0 0 10px;
      font-family: 'Fraunces', serif;
      font-size: clamp(2rem, 4.5vw, 3.3rem);
      line-height: 1.02;
    }}
    .hero-copy p {{
      margin: 0;
      color: var(--muted);
      line-height: 1.6;
      max-width: 64ch;
      font-size: 1rem;
    }}
    .hero-motif {{
      border: 1px solid var(--border);
      background:
        linear-gradient(135deg, rgba(183, 138, 88, 0.22), rgba(0, 123, 95, 0.12)),
        repeating-linear-gradient(45deg, rgba(18, 50, 47, 0.06), rgba(18, 50, 47, 0.06) 12px, rgba(255, 255, 255, 0.14) 12px, rgba(255, 255, 255, 0.14) 24px);
      border-radius: 18px;
      padding: 18px;
      display: grid;
      align-content: end;
      min-height: 180px;
    }}
    .hero-motif strong {{
      font-family: 'Fraunces', serif;
      font-size: 1.2rem;
      letter-spacing: 0.02em;
    }}
    .hero-motif span {{
      margin-top: 4px;
      color: #274845;
      font-size: 0.95rem;
    }}
    form {{ display: grid; gap: 12px; margin: 24px 0 16px; position: relative; z-index: 1; }}
    .form-actions {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
    textarea, button {{
      font: inherit;
      border-radius: 14px;
      border: 1px solid var(--border);
      padding: 13px 14px;
    }}
    textarea {{
      resize: vertical;
      min-height: 120px;
      background: white;
      transition: border-color 180ms ease, box-shadow 180ms ease;
    }}
    textarea:focus {{
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 4px rgba(0, 123, 95, 0.12);
    }}
    button {{
      width: fit-content;
      background: var(--accent);
      color: white;
      border-color: var(--accent);
      cursor: pointer;
      font-weight: 800;
      transition: transform 160ms ease, background-color 160ms ease;
    }}
    button:hover {{
      background: var(--accent-strong);
      transform: translateY(-1px);
    }}
    pre {{
      margin: 0;
      padding: 18px;
      border-radius: 16px;
      border: 1px solid #0f4f42;
      background: linear-gradient(145deg, var(--surface-deep), var(--surface-deep-2));
      color: #e9fff7;
      overflow: auto;
      min-height: 220px;
      white-space: pre-wrap;
      word-break: break-word;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.08);
    }}
    .quick-links {{ display: flex; gap: 8px; flex-wrap: wrap; }}
    .quick-links a {{
      color: var(--accent);
      text-decoration: none;
      border: 1px solid var(--border);
      border-radius: 999px;
      padding: 7px 11px;
      background: var(--accent-soft);
      font-weight: 700;
      font-size: 0.86rem;
    }}
    .quick-links a:hover {{ border-color: var(--accent); }}
    @keyframes rise-in {{
      from {{ opacity: 0; transform: translateY(12px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    @keyframes fade-up {{
      from {{ opacity: 0; transform: translateY(8px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    @media (max-width: 860px) {{
      .hero {{ grid-template-columns: 1fr; }}
      .hero-motif {{ min-height: 130px; }}
    }}
    @media (max-width: 720px) {{
      body {{ padding: 14px; }}
      main {{ padding: 18px; border-radius: 18px; }}
      .brandbar {{ align-items: flex-start; flex-direction: column; }}
      .logo-chip img {{ max-height: 18px; }}
      .form-actions {{ align-items: stretch; }}
    }}
  </style>
</head>
<body>
  <main>
    <section class=\"brandbar\">
      <div>
        <div class=\"kicker\">Storebrand x SKAGEN uttrykk</div>
      </div>
      <div class=\"brandstack\">
        <span class=\"logo-chip\"><img src=\"{STOREBRAND_LOGO_URL}\" alt=\"Storebrand\">Storebrand</span>
        <span class=\"logo-chip\"><img src=\"{SKAGEN_LOGO_URL}\" alt=\"SKAGEN Fondene\">SKAGEN</span>
      </div>
    </section>
    <section class="hero">
      <div class="hero-copy">
        <h1>Oppslag på organisasjonsnummer</h1>
        <p>Et raskt, visuelt og ryddig arbeidsrom for oppslag av selskapsdata. Lim inn ett eller flere org.nr. separert med komma, mellomrom eller linjeskift.</p>
      </div>
      <aside class="hero-motif">
        <strong>Nordic quality</strong>
        <span>Dataflyt inspirert av Storebrand og SKAGEN sitt uttrykk.</span>
      </aside>
    </section>
    <form id=\"lookup-form\">
      <textarea id=\"orgnr\" name=\"orgnr\" required>{example_orgnr}</textarea>
      <div class=\"form-actions\">
        <button type=\"submit\">Hent data</button>
        <div class=\"quick-links\">
          <a href=\"/health\" target=\"_blank\" rel=\"noreferrer\">Health</a>
          <a href=\"/lookup?orgnr=988742163\" target=\"_blank\" rel=\"noreferrer\">JSON eksempel</a>
          <a href=\"/lookup?orgnr=979484534,988742163\" target=\"_blank\" rel=\"noreferrer\">JSON flere org.nr</a>
          <a href=\"/lookup-html?orgnr=979484534,988742163\" target=\"_blank\" rel=\"noreferrer\">HTML flere org.nr</a>
        </div>
      </div>
    </form>
    <pre id=\"result\">Trykk på Hent data for å teste API-et.</pre>
  </main>
  <script>
    const form = document.getElementById('lookup-form');
    const input = document.getElementById('orgnr');
    const result = document.getElementById('result');

    async function loadLookup(orgnrInput) {{
      result.textContent = 'Henter...';
      try {{
        const response = await fetch(`/lookup?orgnr=${{encodeURIComponent(orgnrInput)}}`);
        const data = await response.json();
        result.textContent = JSON.stringify(data, null, 2);
      }} catch (error) {{
        result.textContent = `Feil: ${{error.message}}`;
      }}
    }}

    form.addEventListener('submit', (event) => {{
      event.preventDefault();
      const orgnrInput = input.value.trim();
      if (!orgnrInput) {{
        result.textContent = 'Du må skrive inn minst ett org.nr.';
        return;
      }}
      loadLookup(orgnrInput);
    }});
  </script>
</body>
</html>
"""


def build_result_html(payload):
    selskapsnavn = escape(payload.get("selskapsnavn") or "Ukjent selskap")
    orgnr = escape(payload.get("organisasjonsnummer") or "")
    links = payload.get("links") or {}

    rows = [
        ("Selskap", payload.get("selskapsnavn")),
        ("Organisasjonsnummer", payload.get("organisasjonsnummer")),
        ("Kontaktperson", payload.get("kontaktperson_navn")),
        ("Rolle", payload.get("rolle")),
        ("Telefon", payload.get("telefon")),
        ("Adresse", payload.get("adresse")),
        ("Postnr", payload.get("postnr")),
        ("Poststed", payload.get("poststed")),
        ("Fylke", payload.get("fylke")),
        ("Regnskapsfører", payload.get("regnskapsforer")),
        ("Driftsinntekter", payload.get("driftsinntekter")),
        ("Årsresultat", payload.get("aarsresultat")),
        ("Sum Kasse/Bank/Post", payload.get("sum_kasse_bank_post")),
        ("Sum investeringer", payload.get("sum_investeringer")),
    ]

    rows_html = "\n".join(
        f"<tr><th>{escape(label)}</th><td>{escape(format_value(value))}</td></tr>" for label, value in rows
    )

    link_items = []
    for label, key in [("Proff", "proff"), ("Brreg", "brreg"), ("1881", "1881"), ("LinkedIn", "linkedin")]:
        url = links.get(key) or ""
        if url:
            link_items.append(f"<a href=\"{escape(url)}\" target=\"_blank\" rel=\"noreferrer\">{label}</a>")
    links_html = "\n".join(link_items) if link_items else "<span>Ingen lenker tilgjengelig.</span>"

    return f"""<!DOCTYPE html>
<html lang=\"no\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Resultat {orgnr}</title>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
  <link href=\"https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Manrope:wght@400;500;700;800&display=swap\" rel=\"stylesheet\">
  <style>
    :root {{
      --bg: #edf4f2;
      --panel: #ffffff;
      --ink: #12322f;
      --muted: #4b6763;
      --accent: #007b5f;
      --accent-strong: #00644d;
      --border: #c5ddd5;
      --accent-soft: #dff1ea;
      --gold: #b78a58;
    }}
    body {{
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at 0% 0%, rgba(0, 123, 95, 0.22), transparent 30%),
        radial-gradient(circle at 100% 100%, rgba(183, 138, 88, 0.14), transparent 34%),
        var(--bg);
      color: var(--ink);
      padding: 24px;
      display: grid;
      place-items: center;
      font-family: 'Manrope', sans-serif;
    }}
    main {{
      width: min(960px, 100%);
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 24px;
      box-shadow: 0 24px 64px rgba(11, 49, 43, 0.14);
      padding: 24px;
      animation: rise-in 420ms ease-out both;
      position: relative;
      overflow: hidden;
    }}
    .brandbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 16px;
      padding-bottom: 14px;
      border-bottom: 1px solid var(--border);
    }}
    .brandstack {{ display: flex; gap: 10px; flex-wrap: wrap; }}
    .logo-chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      border: 1px solid var(--border);
      background: #fff;
      border-radius: 999px;
      padding: 7px 11px;
      box-shadow: 0 5px 18px rgba(0, 0, 0, 0.06);
      font-size: 0.85rem;
      color: var(--muted);
    }}
    .logo-chip img {{ max-height: 22px; width: auto; }}
    .headline {{
      display: grid;
      gap: 12px;
      grid-template-columns: 1fr auto;
      align-items: end;
      margin-bottom: 6px;
    }}
    h1 {{ margin: 0; font-family: 'Fraunces', serif; font-size: clamp(1.95rem, 4vw, 2.85rem); line-height: 1.03; }}
    .orgnr-tag {{
      border: 1px solid var(--border);
      background: var(--accent-soft);
      border-radius: 999px;
      padding: 7px 12px;
      font-size: 0.9rem;
      font-weight: 700;
      color: var(--accent-strong);
      white-space: nowrap;
    }}
    .meta {{ margin: 0; color: var(--muted); }}
    .insight {{
      margin-top: 16px;
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 12px;
      background: linear-gradient(130deg, #f9fcfb, #eff7f3);
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }}
    .insight div {{
      border: 1px dashed #bdd6ce;
      border-radius: 10px;
      padding: 10px;
      background: #fff;
    }}
    .insight strong {{
      display: block;
      font-size: 0.79rem;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--muted);
      margin-bottom: 4px;
    }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 18px; }}
    th, td {{ border-bottom: 1px solid var(--border); text-align: left; padding: 11px 8px; vertical-align: top; }}
    th {{ width: 32%; color: var(--muted); font-weight: 700; }}
    .links {{ display: flex; gap: 12px; flex-wrap: wrap; margin-top: 18px; }}
    .links a {{
      text-decoration: none;
      background: var(--accent);
      color: #fff;
      border-radius: 12px;
      padding: 9px 13px;
      font-size: 0.95rem;
      font-weight: 700;
      transition: transform 160ms ease, background-color 160ms ease;
    }}
    .links a:hover {{
      background: var(--accent-strong);
      transform: translateY(-1px);
    }}
    @keyframes rise-in {{
      from {{ opacity: 0; transform: translateY(10px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    @media (max-width: 720px) {{
      body {{ padding: 14px; }}
      main {{ padding: 18px; border-radius: 18px; }}
      .brandbar {{ align-items: flex-start; flex-direction: column; }}
      .logo-chip img {{ max-height: 18px; }}
      .headline {{ grid-template-columns: 1fr; align-items: start; }}
      .insight {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <main>
    <section class=\"brandbar\">
      <div></div>
      <div class=\"brandstack\">
        <span class=\"logo-chip\"><img src=\"{STOREBRAND_LOGO_URL}\" alt=\"Storebrand\">Storebrand</span>
        <span class=\"logo-chip\"><img src=\"{SKAGEN_LOGO_URL}\" alt=\"SKAGEN Fondene\">SKAGEN</span>
      </div>
    </section>
    <section class="headline">
      <h1>{selskapsnavn}</h1>
      <div class="orgnr-tag">Org.nr {orgnr}</div>
    </section>
    <p class="meta">Detaljert selskapsprofil hentet fra Brreg og Proff-kilder.</p>
    <section class="insight">
      <div><strong>Kontaktperson</strong>{escape(format_value(payload.get('kontaktperson_navn')))}</div>
      <div><strong>Regnskapsfører</strong>{escape(format_value(payload.get('regnskapsforer')))}</div>
    </section>
    <table>
      <tbody>
        {rows_html}
      </tbody>
    </table>
    <div class=\"links\">{links_html}</div>
  </main>
</body>
</html>
"""


def build_batch_result_html(payloads, errors):
    rows = []
    for payload in payloads:
        links = payload.get("links") or {}
        link_items = []
        for label, key in [("Proff", "proff"), ("Brreg", "brreg"), ("1881", "1881"), ("LinkedIn", "linkedin")]:
            url = links.get(key) or ""
            if url:
                link_items.append(f"<a href=\"{escape(url)}\" target=\"_blank\" rel=\"noreferrer\">{label}</a>")

        rows.append(
            "<tr>"
            f"<td>{escape(format_value(payload.get('organisasjonsnummer')))}</td>"
            f"<td>{escape(format_value(payload.get('selskapsnavn')))}</td>"
            f"<td>{escape(format_value(payload.get('kontaktperson_navn')))}</td>"
          f"<td>{escape(format_value(payload.get('rolle')))}</td>"
          f"<td>{escape(format_value(payload.get('telefon')))}</td>"
          f"<td>{escape(format_value(payload.get('adresse')))}</td>"
          f"<td>{escape(format_value(payload.get('postnr')))}</td>"
          f"<td>{escape(format_value(payload.get('poststed')))}</td>"
          f"<td>{escape(format_value(payload.get('fylke')))}</td>"
          f"<td>{escape(format_value(payload.get('regnskapsforer')))}</td>"
            f"<td>{escape(format_value(payload.get('driftsinntekter')))}</td>"
            f"<td>{escape(format_value(payload.get('aarsresultat')))}</td>"
          f"<td>{escape(format_value(payload.get('sum_kasse_bank_post')))}</td>"
          f"<td>{escape(format_value(payload.get('sum_investeringer')))}</td>"
            f"<td class=\"links\">{' '.join(link_items)}</td>"
            "</tr>"
        )

    error_html = ""
    if errors:
        items = "".join(
            f"<li>{escape(e.get('organisasjonsnummer', ''))}: {escape(e.get('error', ''))}</li>" for e in errors
        )
        error_html = f"<section><h2>Feil</h2><ul>{items}</ul></section>"

    table_rows = "\n".join(rows) if rows else "<tr><td colspan=\"15\">Ingen resultater.</td></tr>"

    return f"""<!DOCTYPE html>
<html lang=\"no\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Batch-resultat</title>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
  <link href=\"https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Manrope:wght@400;500;700;800&display=swap\" rel=\"stylesheet\">
  <style>
    :root {{
      --bg: #edf4f2;
      --panel: #ffffff;
      --ink: #12322f;
      --muted: #4b6763;
      --accent: #007b5f;
      --accent-strong: #00644d;
      --border: #c5ddd5;
      --stripe: #f4faf7;
    }}
    body {{
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at 0% 0%, rgba(0, 123, 95, 0.22), transparent 30%),
        radial-gradient(circle at 100% 100%, rgba(183, 138, 88, 0.14), transparent 34%),
        var(--bg);
      color: var(--ink);
      padding: 24px;
      font-family: 'Manrope', sans-serif;
    }}
    main {{
      max-width: 1240px;
      margin: 0 auto;
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 24px;
      box-shadow: 0 24px 64px rgba(11, 49, 43, 0.14);
      padding: 24px;
      animation: rise-in 420ms ease-out both;
    }}
    .brandbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 16px;
      padding-bottom: 14px;
      border-bottom: 1px solid var(--border);
    }}
    .brandstack {{ display: flex; gap: 10px; flex-wrap: wrap; }}
    .logo-chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      border: 1px solid var(--border);
      background: #fff;
      border-radius: 999px;
      padding: 7px 11px;
      box-shadow: 0 5px 18px rgba(0, 0, 0, 0.06);
      font-size: 0.85rem;
      color: var(--muted);
    }}
    .logo-chip img {{ max-height: 22px; width: auto; }}
    h1 {{ margin-top: 0; margin-bottom: 8px; font-family: 'Fraunces', serif; font-size: clamp(1.95rem, 4vw, 2.9rem); }}
    .summary {{ margin: 0 0 14px; color: var(--muted); }}
    .table-wrap {{ border: 1px solid var(--border); border-radius: 14px; overflow: auto; }}
    table {{ width: 100%; border-collapse: collapse; min-width: 1680px; }}
    th, td {{ border-bottom: 1px solid var(--border); text-align: left; padding: 11px 9px; vertical-align: top; }}
    th {{ color: var(--muted); font-weight: 700; background: #eff7f3; position: sticky; top: 0; z-index: 1; }}
    tbody tr:nth-child(even) {{ background: var(--stripe); }}
    .links a {{
      text-decoration: none;
      background: var(--accent);
      color: #fff;
      border-radius: 10px;
      padding: 7px 11px;
      margin-right: 6px;
      display: inline-block;
      margin-bottom: 4px;
      font-size: 0.9rem;
      font-weight: 700;
      transition: transform 160ms ease, background-color 160ms ease;
    }}
    .links a:hover {{
      background: var(--accent-strong);
      transform: translateY(-1px);
    }}
    section h2 {{ font-family: 'Fraunces', serif; margin-bottom: 6px; }}
    @keyframes rise-in {{
      from {{ opacity: 0; transform: translateY(10px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    @media (max-width: 720px) {{
      body {{ padding: 14px; }}
      main {{ padding: 18px; border-radius: 18px; }}
      .brandbar {{ align-items: flex-start; flex-direction: column; }}
      .logo-chip img {{ max-height: 18px; }}
    }}
  </style>
</head>
<body>
  <main>
    <section class=\"brandbar\">
      <div></div>
      <div class=\"brandstack\">
        <span class=\"logo-chip\"><img src=\"{STOREBRAND_LOGO_URL}\" alt=\"Storebrand\">Storebrand</span>
        <span class=\"logo-chip\"><img src=\"{SKAGEN_LOGO_URL}\" alt=\"SKAGEN Fondene\">SKAGEN</span>
      </div>
    </section>
    <h1>Batch-oppslag</h1>
    <p class="summary">{len(payloads)} resultat(er) fra kombinert datainnhenting.</p>
    <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Org.nr</th>
          <th>Selskap</th>
          <th>Kontaktperson</th>
          <th>Rolle</th>
          <th>Telefon</th>
          <th>Adresse</th>
          <th>Postnr</th>
          <th>Poststed</th>
          <th>Fylke</th>
          <th>Regnskapsfører</th>
          <th>Driftsinntekter</th>
          <th>Årsresultat</th>
          <th>Sum Kasse/Bank/Post</th>
          <th>Sum investeringer</th>
          <th>Lenker</th>
        </tr>
      </thead>
      <tbody>
        {table_rows}
      </tbody>
    </table>
    </div>
    {error_html}
  </main>
</body>
</html>
"""


class OrgnrRequestHandler(BaseHTTPRequestHandler):
    timeout = DEFAULT_TIMEOUT

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            return self._send_html(200, build_index_html())

        if parsed.path == "/health":
            return self._send_json(200, {"status": "ok"})

        query = parse_qs(parsed.query)
        orgnrs = parse_orgnrs_from_query(query)

        if parsed.path == "/lookup-html":
            if not orgnrs:
                return self._send_json(400, {"error": "Missing required query parameter: orgnr"})

            payloads, errors = build_payloads(orgnrs, timeout=self.timeout)
            if len(orgnrs) == 1 and payloads:
                return self._send_html(200, build_result_html(payloads[0]))
            return self._send_html(200, build_batch_result_html(payloads, errors))

        if parsed.path == "/lookup":
            if not orgnrs:
                return self._send_json(400, {"error": "Missing required query parameter: orgnr"})

            payloads, errors = build_payloads(orgnrs, timeout=self.timeout)
            if len(orgnrs) == 1 and payloads:
                return self._send_json(200, payloads[0])

            return self._send_json(
                200,
                {
                    "count": len(payloads),
                    "results": payloads,
                    "errors": errors,
                },
            )

        return self._send_json(404, {"error": "Not found"})

    def log_message(self, format, *args):
        return

    def _send_json(self, status_code, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, status_code, html):
        body = html.encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    parser = argparse.ArgumentParser(description="Start et API for oppslag på organisasjonsnummer.")
    parser.add_argument("--host", default="0.0.0.0", help="Host å lytte på")
    parser.add_argument("--port", type=int, default=8000, help="Port å lytte på")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Timeout i sekunder per forespørsel")
    args = parser.parse_args()

    handler = type(
        "ConfiguredOrgnrRequestHandler",
        (OrgnrRequestHandler,),
        {"timeout": args.timeout},
    )
    server = ThreadingHTTPServer((args.host, args.port), handler)
    print(f"API lytter på http://{args.host}:{args.port}")
    print("Bruk /lookup?orgnr=123456789")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
