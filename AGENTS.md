# AGENTS.md

## Syfte

Detta repo ska stodja tva arbetssatt samtidigt:

- Manniskor ska kunna kora appen via GUI.
- Agenter och script ska kunna kora samma arbetsfloden via CLI.

Nar du bygger nya funktioner, utga fran att GUI och CLI ska dela samma motor.

For en praktisk korguide for runtime-agenter, se `AGENT_SIMULATION.md`.

## Huvudregel

Ny affarslogik ska inte byggas direkt in i `tkinter`-knappar, dialoger eller `messagebox`-floden.

Bygg i stallet i denna ordning:

1. Ren hjalpfunktion eller workflow-funktion
2. CLI-adapter som anropar samma workflow
3. GUI-adapter som anropar samma workflow

GUI ska vara presentationslager. CLI ska vara automationslager. Resultatet ska komma fran samma underliggande logik.

## Onskad struktur for nya funktioner

Nar en ny funktion laggs till, anvand helst detta monster:

1. Las in filer eller parametrar
2. Normalisera data i separata helpers
3. Kor en delad workflow-funktion
4. Returnera resultat, statistik, varningar och rapportdata
5. Lat GUI visa resultatet och CLI skriva resultatet

Undvik att lagga domanlogik i:

- `messagebox`
- clipboard-funktioner
- `open_*_in_excel()`-metoder
- `StringVar`-beroenden
- knapphandlers som blandar inlasning, berakning och export i samma block

## CLI-regler

Om en funktion kan koras utan manuell GUI-interaktion ska den normalt fa en CLI-vag.

CLI-kommandon bor:

- vara icke-interaktiva
- ta filer och parametrar via flaggor
- kunna skriva resultat till fil
- kunna skriva en kort JSON-sammanfattning med `--json`
- ge tydliga exit-koder vid fel
- undvika popup-fonster, clipboard och Excel-oppning

Bra standardflaggor nar de passar:

- `--json`
- `--report-out`
- `--output`
- `--result-out`
- `--details-out`

## GUI-regler

GUI far garna vara bekvamt for anvandaren, men ska helst bara:

- samla in filval och val
- anropa delad workflow
- visa varningar och status
- oppna exporterade filer

Om GUI och CLI riskerar att bete sig olika ar det ett tecken pa att logiken sitter for langt ut i UI-lagret.

## Testregler

Nar du bygger eller andrar funktioner, forsok fa med testning pa tva nivaer:

1. Servicetester for ren logik och normalisering
2. CLI end-to-end-tester for hela arbetsfloden

Minimikrav for en ny CLI-bar funktion:

- minst ett test som kor kommandot via CLI
- kontroll av viktig outputfil eller JSON-svar
- minst ett test for central gren eller kantfall om logiken ar mer an trivial

## Fler tester som ar hogt varderade har

Foljdande tester ar naturliga nasta steg:

- `hib-koppling` som komplett CLI end-to-end-test
- `allocate` med near-miss-traff
- `allocate` med refill-output
- `allocate` med pallet-space-output
- `overview-check` nar inga avvikelser hittas
- `overview-check` med flera HIB-rader och dubbletter
- `dispatch-check` nar allt matchar
- `dispatch-check` med flera olika feltyper i samma korning
- `vecka27-check` med tomt resultat
- `vecka27-check` med flera avvikande orderrader
- `eftersok` nar vissa valfria WMS-filer saknas
- `eftersok` nar artikel eller inkop inte ger traff
- tester for kolumnalias och fuzzy kolumnmatchning
- tester for fil med saknade obligatoriska kolumner
- tester for felaktiga datum- eller nummerformat
- tester for rapportskrivning till `csv`, `txt` och `xlsx` dar det ar relevant

## Floden som bor prioriteras for framtida CLI

Nar nya eller befintliga floden byggs ut, prioritera att gora dessa CLI-vanliga om de inte redan ar det:

- prognos/autoplock-rapport
- sales eller plocklogg-baserade analyser
- observationskontroller
- andra rapportfloden som idag bara gar via GUI-export

## Nar CLI inte ar rimligt

Om en funktion verkligen ar rent visuell eller starkt beroende av live-GUI-tillstand kan den fa vara GUI-only.

Men da ska det vara ett medvetet val, och koden eller dokumentationen bor kort saga varfor ingen CLI-vag finns.
