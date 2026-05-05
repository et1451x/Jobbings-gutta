# Hent info om kunder

Script som tar en kundeliste med selskapsnavn og organisasjonsnummer fra en Excel-fil, og beriker den med kontaktperson, telefon, adresse, regnskapsfører og regnskapsdata. Resultatet er en ferdig Excel-fil med klikkbare lenker.

Det finnes også et lite HTTP-API i samme mappe hvis du kun vil slå opp ett organisasjonsnummer om gangen.

## Hvordan det fungerer

### 1. Les kundeliste fra input-fil
Scriptet leser første ark i input-filen. Rad 1 er overskrifter (hoppes over). Kolonne A = selskapsnavn, kolonne B = organisasjonsnummer.

### 2. Hent kontaktperson fra Brreg
For hvert organisasjonsnummer hentes roller fra Brønnøysundregistrene. Scriptet prioriterer:
1. **Kontaktperson (KONT)** — hvis registrert
2. **Daglig leder (DAGL)** — fallback
3. **Styreleder (STYR)** — siste fallback

### 3. Fallback til Proff.no
Hvis Brreg ikke gir noen kontaktperson, søkes det på Proff.no etter styreleder via HTML-scraping.

### 4. Berik med tilleggsdata

| Kilde | Data |
|---|---|
| **Brreg Enhetsregisteret** | Adresse, postnr, poststed, fylke |
| **Brreg Roller-API** | Kontaktperson, rolle, regnskapsfører |
| **Proff.no** (HTTP) | Telefonnummer, Sum Kasse/Bank/Post (KBPS), Sum investeringer (SIV) |

### 5. Skriv resultat til Excel
Output-filen inneholder ett ark med autofilter, formaterte beløp og klikkbare lenker til Proff, 1881 og LinkedIn.

## Installasjon

```bash
pip install openpyxl requests
```

> **Anbefalt:** Installer også `tqdm` for å se en fremdriftslinje:
> ```bash
> pip install tqdm
> ```
> Uten `tqdm` vises fremdrift som `1/125`, `2/125` osv.

## Bruk

### Start API for org.nr-oppslag

```bash
python orgnr_api.py --port 8000
```

Eksempel på kall:

```bash
curl "http://localhost:8000/lookup?orgnr=988742163"
```

Tilgjengelige endepunkt:

| Endepunkt | Beskrivelse |
|---|---|
| `GET /lookup?orgnr=...` | Returnerer JSON med selskapsnavn, kontaktperson, telefon, adresse, regnskapsfører og regnskap (`driftsinntekter`, `aarsresultat`, `KBPS`, `SIV`). Ved flere org.nr returneres `count`, `results` og `errors`. |
| `GET /lookup-html?orgnr=...` | Returnerer en HTML-resultatside med klikkbare lenker til Proff, Brreg, 1881 og LinkedIn. Ved flere org.nr vises en batch-tabell. |
| `GET /health` | Enkel helsesjekk |

Du kan sende flere org.nr i samme query med komma, mellomrom eller linjeskift, for eksempel:

```bash
curl "http://localhost:8000/lookup?orgnr=979484534,988742163"
```

### Forbered input-fil
Lag en Excel-fil med selskapsnavn i kolonne A og organisasjonsnummer i kolonne B. Rad 1 er overskrifter:

| Selskap | Organisasjonsnummer |
|---|---|
| BENDIKS AS | 988742163 |
| PROCEDO NOR AS | 989396749 |

### Kjør scriptet

```bash
python Brreg_Proff_fallback.script.py --input kundeliste.xlsx --output resultat.xlsx
```

### Argumenter

| Argument | Påkrevd | Beskrivelse |
|---|---|---|
| `--input` | Ja | Input Excel-fil med kundeliste |
| `--output` | Ja | Filnavn for resultatet |
| `--limit` | Nei | Prosesser kun de første N radene (nyttig for testing) |
| `--timeout` | Nei | Timeout i sekunder per forespørsel (standard: 20) |

### Eksempel med limit

```bash
python Brreg_Proff_fallback.script.py --input ebbekunder.xlsx --output resultat.xlsx --limit 5
```

## Output-kolonner

| # | Kolonne | Beskrivelse |
|---|---|---|
| 1 | Selskap | Selskapsnavn fra input-filen |
| 2 | Organisasjonsnummer | Org.nr fra input-filen |
| 3 | Kontaktperson navn | Daglig leder, kontaktperson eller styreleder |
| 4 | Rolle | KONT, DAGL eller STYR |
| 5 | Telefon | Telefonnummer fra Proff.no |
| 6 | Adresse | Forretningsadresse fra Brreg |
| 7 | Postnr | Postnummer |
| 8 | Poststed | Poststed |
| 9 | Fylke | Fylke (utledet fra kommunenummer) |
| 10 | Regnskapsfører | Regnskapsførerselskapet fra Brreg |
| 11 | Sum Kasse/Bank/Post | Fra Proff.no (KBPS) |
| 12 | Sum investeringer | Fra Proff.no (SIV) |
| 13 | Proff | Lenke til Proff.no |
| 14 | 1881 | Lenke til 1881-søk på kontaktperson |
| 15 | LinkedIn | Lenke til LinkedIn-søk på kontaktperson |

## Begrensninger

- Proff.no bruker JavaScript-rendering. HTTP-scraping fungerer for telefon og regnskap, men kan gi tomme resultater for noen selskaper.
- KBPS/SIV-verdier fra Proff brukes direkte uten ekstra multiplikasjon.
- Noen selskaper mangler regnskap på Proff (f.eks. nyregistrerte eller ENK).
