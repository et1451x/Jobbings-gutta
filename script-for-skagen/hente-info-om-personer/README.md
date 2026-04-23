# Finn bedriftseiere

Script som tar en liste med personnavn fra en Excel-fil og slår opp om de er registrert med roller (styreleder, daglig leder, osv.) i norske selskaper via Brønnøysundregistrene. Resultatet er en rik Excel-fil med selskapsinformasjon, regnskapsdata og lenker.

## Hvordan det fungerer

### 1. Les personnavn fra input-fil
Scriptet leser kolonnen **Navn** fra `personer.xlsx` (eller en annen fil angitt med `--input`).

### 2. Last ned roller-totalbestand fra Brreg
Ved første kjøring lastes hele roller-totalbestanden ned fra Brønnøysundregistrene (~132 MB gzippet). Filen caches lokalt som `roller_totalbestand.json.gz` og gjenbrukes ved senere kjøringer. Hvis cachen er eldre enn 30 dager, lastes den ned på nytt automatisk.

### 3. Søk gjennom totalbestanden
Alle ~1.1 millioner enheter i totalbestanden gjennomgås. For hver person sjekkes det om noen har en aktiv rolle (ikke fratrådt/avregistrert) der navnedelene matcher eksakt (case-insensitivt, rekkefølge-uavhengig). Etternavn brukes som rask forhåndsfiltrering for ytelse.

### 4. Berik med detaljert informasjon
For hvert unike organisasjonsnummer som ble funnet, hentes ytterligere data:

| Kilde | Data |
|---|---|
| **Brreg Enhetsregisteret** | Selskapsnavn, adresse, postnr, poststed, fylke |
| **Brreg Roller-API** | Kontaktperson (daglig leder/styreleder), regnskapsfører |
| **Brreg Regnskapsregisteret** | Driftsinntekter, årsresultat |
| **Proff.no** (via HTTP) | Telefonnummer, Sum Kasse/Bank/Post (KBPS), Sum investeringer (SIV) |

Proff.no hentes via HTTP med full User-Agent-header (søk → profilside → regnskapsside).

### 5. Skriv resultat til Excel
Output-filen inneholder to ark:
- **Bedriftseiere** — Én rad per person/rolle-kombinasjon med alle kolonner, autofilter, formaterte beløp og klikkbare lenker til Proff, Brreg, 1881 og LinkedIn.
- **Oppsummering** — Statistikk og antall roller per person.

## Installasjon

```bash
pip install openpyxl requests tqdm
```

> **Anbefalt:** `tqdm` gir en fremdriftslinje under søk i totalbestanden.
> Uten `tqdm` fungerer scriptet helt fint, men du ser ikke hvor langt søket har kommet.

## Bruk

### Forbered input-fil
Lag en Excel-fil (`personer.xlsx`) med én kolonne kalt **Navn** og ett personnavn per rad:

| Navn |
|---|
| Ola Nordmann |
| Kari Hansen |

### Kjør scriptet

```bash
python finn_bedriftseiere.py
```

### Valgfrie argumenter

| Argument | Standard | Beskrivelse |
|---|---|---|
| `--input` | `personer.xlsx` | Input Excel-fil med personnavn |
| `--output` | `bedriftseiere_resultat.xlsx` | Filnavn for resultatet |
| `--cache` | `roller_totalbestand.json.gz` | Cache-fil for totalbestanden |
| `--force-download` | — | Tving ny nedlasting av totalbestand |

### Eksempel

```bash
python finn_bedriftseiere.py --input mine_navn.xlsx --output resultat.xlsx
```

## Output-kolonner

| # | Kolonne | Beskrivelse |
|---|---|---|
| 1 | Person | Søkenavnet fra input-filen |
| 2 | Selskap | Selskapsnavn fra Brreg |
| 3 | Org.nr | Organisasjonsnummer |
| 4 | Rolle | Type rolle (f.eks. Styreleder, Daglig leder) |
| 5 | Rollegruppe | Rollegruppetype (f.eks. Styre, Ledelse) |
| 6 | Navn i Brreg | Personens fulle navn slik det står i Brreg |
| 7 | Kontaktperson | Daglig leder eller styreleder i selskapet |
| 8 | Telefon | Telefonnummer fra Proff.no |
| 9 | Adresse | Forretningsadresse |
| 10 | Postnr | Postnummer |
| 11 | Poststed | Poststed |
| 12 | Fylke | Fylke (utledet fra kommunenummer) |
| 13 | Regnskapsfører | Regnskapsførerselskapet |
| 14 | Driftsinntekter | Siste år, fra Brreg regnskapsregisteret |
| 15 | Årsresultat | Siste år, fra Brreg regnskapsregisteret |
| 16 | Sum Kasse/Bank/Post | Siste år, fra Proff.no (KBPS × 1000) |
| 17 | Sum investeringer | Siste år, fra Proff.no (SIV × 1000) |
| 18 | Proff | Lenke til Proff.no |
| 19 | Brreg | Lenke til Brreg virksomhetsoppslag |
| 20 | 1881 | Lenke til 1881-søk på kontaktperson |
| 21 | LinkedIn | Lenke til LinkedIn-søk på kontaktperson |

## Navnematching

Navnene matches med **eksakt sett-likhet** — alle navnedeler må være identiske, uavhengig av rekkefølge. For eksempel vil "Lars Holm" matche "Holm Lars" i Brreg, men **ikke** "Lars Erik Holm" (forskjellig antall deler).

## Begrensninger

- Kun personer med aktive roller i Enhetsregisteret gir treff. Historiske roller (fratrådt/avregistrert) filtreres bort.
- KBPS/SIV-verdier fra Proff er oppgitt i hele tusen og multipliseres med 1000 i output.
- Noen mindre selskaper mangler regnskap på Proff og/eller i Brreg regnskapsregisteret.
- Proff-scraping gjøres via HTTP (~3 forespørsler per selskap: søk, profil, regnskap).
