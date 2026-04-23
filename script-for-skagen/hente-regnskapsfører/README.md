# Hent kundeliste for regnskapsfører

Script som tar et organisasjonsnummer for et regnskapsførerselskap og henter ut alle selskaper det er registrert som regnskapsfører for i Brønnøysundregistrene. Resultatet er en Excel-fil med kundeliste og oppsummering.

## Hvordan det fungerer

### 1. Slå opp regnskapsførerselskapet
Scriptet henter firmanavnet fra Brreg Enhetsregisteret basert på oppgitt organisasjonsnummer.

### 2. Hent alle kunder via roller-API
Scriptet paginerer gjennom Brønnøysundregistrenes roller-API (`/roller/enheter/{orgnr}/juridiskeroller`) og henter alle enheter der regnskapsførerselskapet har en rolle. Sidene hentes automatisk til alle er lastet.

### 3. Klassifiser status
Hver kunde klassifiseres som:
- **Aktiv** — har minst én aktiv rolle
- **Fratrådt** — alle roller er fratrådt
- **Avregistrert** — alle roller er avregistrert

### 4. Skriv resultat til Excel
Output-filen inneholder to ark:
- **Kunder** — Nummerert liste med selskap, org.nr, rolle og status. Sortert med aktive først. Autofilter og fargekoding (grønn = aktiv, rød = fratrådt/avregistrert).
- **Oppsummering** — Firmanavn, dato, og statistikk (totalt, aktive, fratrådte, avregistrerte).

## Installasjon

```bash
pip install openpyxl requests
```

## Bruk

```bash
python Regnskapsfører_alle_kunder_xlsx.py --input <orgnr> --output <filnavn.xlsx>
```

`--input` er organisasjonsnummeret til regnskapsførerselskapet (ikke en fil).

### Argumenter

| Argument | Påkrevd | Beskrivelse |
|---|---|---|
| `--input` | Ja | Organisasjonsnummer for regnskapsførerselskapet |
| `--output` | Nei | Filnavn for resultatet (standard: `kunder_<orgnr>.xlsx`) |

### Eksempel

```bash
python Regnskapsfører_alle_kunder_xlsx.py --input 950836792 --output kundeliste.xlsx
```

## Output-kolonner (Kunder-arket)

| # | Kolonne | Beskrivelse |
|---|---|---|
| 1 | # | Radnummer |
| 2 | Selskap | Selskapsnavn |
| 3 | Org.nr | Organisasjonsnummer |
| 4 | Rolle | Rolletyper (f.eks. Regnskapsfører) |
| 5 | Status | Aktiv, Fratrådt eller Avregistrert |

## Begrensninger

- Scriptet henter kun roller registrert i Brønnøysundregistrene. Uformelle kundeforhold vises ikke.
- Filen åpnes automatisk etter generering (Windows).