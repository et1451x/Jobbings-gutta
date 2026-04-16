## Litt mer knask for Skagen Gutta

Last ned fra mappen `script-for-skagen/hente-regnskapsfører/Regnskapsfører_alle_kunder_xlsx.py`

Hvor enn du lagrer filen, kjør:

```
python Regnskapsfører_alle_kunder_xlsx.py --input <orgnr> --output <filnavn.xlsx>
```

### Eksempel

Du laster ned `Regnskapsfører_alle_kunder_xlsx.py` til en mappe, f.eks:  
`c:\scripts\Regnskapsfører_alle_kunder_xlsx.py`

Kjør:

```
python c:\scripts\Regnskapsfører_alle_kunder_xlsx.py --input <orgnr> --output <filnavn.xlsx>
```

`orgnr` = orgnr på det selskap som er regnskapsfører.

Eksempel:

```
python c:\scripts\Regnskapsfører_alle_kunder_xlsx.py --input 950836792 --output c:\kunder\test1.xlsx
```

Den viser så `test1.xlsx` som lister alle selskapene som org.nr `950836792` er regnskapsfører for.

> **NB:** 2 fliker i xlsx — **Kunder** og **Oppsummering**


