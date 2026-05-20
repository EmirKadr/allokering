# AGENT_SIMULATION.md

## Syfte

Denna guide ar till for agenter som ska anvanda `allokering` utan GUI.

Manniskor kan fortsatt anvanda GUI:t, men en agent ska i forsta hand:

- kora arbetsfloden via CLI
- lasa och skriva filer
- tolka JSON-sammanfattningar
- simulera anvandarens arbetssteg utan att oppna fonster

## Grundregel

Anvand inte GUI om samma sak kan goras via CLI.

For en agent ar CLI forstahandsvalet eftersom det ar:

- deterministiskt
- scriptbart
- loggbart
- batchbart
- enklare att testa

## Startpunkt

Visa tillgangliga CLI-kommandon:

```powershell
python allokering12.1.py --help
```

Om ett arbetsflode redan finns som subkommando ska agenten anvanda det direkt.

## Nuvarande CLI-kommandon

- `allocate`
- `ordersaldo`
- `lyx`
- `pafyllnadsprio`
- `hib-koppling`
- `overview-check`
- `dispatch-check`
- `vecka27-check`
- `eftersok`
- `prognos-report`
- `observations-update`
- `observations-sync`
- `split-values`
- `update-check`

## Rekommenderat arbetssatt for agenten

Nar agenten ska simulera ett anvandarbeteende bor den jobba i denna ordning:

1. Identifiera vilket arbetsflode anvandaren forsoker kora
2. Matcha det mot ett befintligt CLI-kommando
3. Kor kommandot med explicita in- och ut-filer
4. Begar `--json` om maskinlasbar summering behovs
5. Validera att utdatafiler skapats och att summeringen ar rimlig
6. Rapportera resultat, avvikelser och eventuella blockerare

## Standard for CLI-korning

Nar det ar mojligt bor agenten foredra:

- indata via filflaggor
- utdata via `--output`, `--report-out`, `--result-out` eller liknande
- `--json` for kort maskinlasbar summering

Exempel:

```powershell
python allokering12.1.py allocate `
  --orders .\orders.csv `
  --buffer .\buffer.csv `
  --result-out .\out\allocated.csv `
  --near-miss-out .\out\near_miss.csv `
  --json
```

## Hur agenten ska tanka om GUI-handlingar

Oversatt GUI-intentioner till CLI sa har:

- "Tryck pa kor allokering" -> `allocate`
- "Oppna resultat i Excel" -> skriv rapport till fil och las filen
- "Kopiera lista" -> skriv text- eller csv-fil
- "Kor kontroll" -> anvand motsvarande `*-check`-kommando
- "Kor eftersok" -> `eftersok` med WMS-filer som argument
- "Skapa prognosrapport" -> `prognos-report`
- "Uppdatera observations och artikel_max" -> `observations-update`
- "Synca observations med GitHub eller lokal kallfil" -> `observations-sync`
- "Dela inklistrad lista i kolumner" -> `split-values`
- "Kontrollera uppdatering" -> `update-check`

Agenten ska alltsa simulera resultatet av anvandarens handling, inte sjalva knapptrycket.

## Scenario-simulering

Om anvandaren vill simulera flera steg bor agenten dela upp det i separata CLI-korningar.

Exempel pa scenario:

1. Kor `ordersaldo`
2. Kor `pafyllnadsprio`
3. Kor `allocate`
4. Las rapporterna
5. Sammanfatta utfall och avvikelser

Det ar battre an att forsoka styra GUI:t som en manniska.

## Nar agenten upptacker att CLI saknas

Om ett onskat arbetsflode inte finns som CLI ska agenten:

1. Leta efter en delad workflow-funktion eller ren berakningsfunktion
2. Undvika att bygga ny logik direkt i GUI
3. Lagga till eller foresla ett nytt CLI-subkommando ovanpa samma motor
4. Lagga till tester for det nya kommandot
5. Uppdatera dokumentationen

## Krav pa nya CLI-floden

Nar agenten bygger ut CLI:t ska den forsoka se till att varje nytt kommando:

- fungerar utan interaktiv input
- accepterar tydliga filflaggor
- kan skriva resultat till fil
- kan skriva JSON-sammanfattning
- returnerar tydligt fel om obligatorisk indata saknas

## Output-kontrakt

Agenten bor behandla utdata sa har:

- `.txt` for korta listor eller rapporttext
- `.csv` for tabellresultat
- `.xlsx` nar flera blad eller Excel-format passar battre
- JSON endast som sammanfattning eller maskinlasbar respons, inte som enda rapportformat om tabellfiler behovs

## Nar GUI fortfarande ar tillatet

GUI ska fortfarande ses som den primara upplevelsen for manniskor.

Agenten ska inte forsoka ta bort GUI-stod. I stallet ska den hjalpa till att halla samma logik tillganglig pa bada sidor:

- GUI for manniskor
- CLI for agenter

## Viktig byggregel

Om agenten skapar eller andrar funktioner i repo:t ska den ocksa lasa:

- `AGENTS.md` for byggregler
- `TESTING.md` for test- och CLI-exempel
- `CLAUDE.md` for overgripande projektkontext

## Malbild

Malet ar att en agent ska kunna simulera sa mycket som mojligt av en verklig anvandare genom att:

- mata in samma filer som anvandaren skulle ladda upp
- kora samma arbetsfloden
- generera samma rapporttyper
- fatta beslut pa samma resultat

utan att vara beroende av GUI-automation.

I nuvarande lage betyder det att alla meningsfulla arbetsfloden i appen har en CLI-vag eller en filbaserad motsvarighet. Rent visuella GUI-detaljer som filvaljardialoger, drag-and-drop och popup-layout ersatts av explicita CLI-argument och utdatafiler.
