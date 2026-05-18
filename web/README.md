# Allokering - Webb/API

Ett modernt, API-styrt gränssnitt för allokeringsflödet. Samma motor som
`allokering12.1.py` (GUI + CLI) - det nya skiktet är ett rent
presentationslager ovanpå ett HTTP-API som också kan deployas som webbapp.

## Arkitektur

```
web/
  backend/
    engine.py    - laddar motorn från allokering12.1.py (delad logik)
    flows.py     - ett flöden-handtag per CLI-kommando + registret
    api.py       - FastAPI: /api/flows, /api/flow/{id}, /api/detect,
                   /api/open-excel, /api/download
    desktop.py   - pywebview-launcher (kör API + visar React-appen i ett fönster)
  frontend/      - React + Vite (sidebar, global drag&drop, filrader, resultattabs,
                   dark/light-tema, popups)
```

Frontenden byggs dynamiskt från `/api/flows` - lägg till ett nytt flöde i
`flows.py` så dyker det upp i menyn automatiskt.

CLI:t (`python allokering12.1.py allocate ...`) och tkinter-GUI:t är orörda -
webbskiktet är additivt och ligger helt i `web/`.

## Floden

CLI-kommandona finns kvar i API/motorn. I webben visas arbetsflödena som används
manuellt:

- **Allokering:** allocate
- **Order & saldo:** ordersaldo, lyx, pafyllnadsprio
- **Kontroller:** hib-koppling, overview-check, dispatch-check, vecka27-check
- **Sökning & prognos:** eftersök, prognos-report
- **Data & verktyg:** split-values, update-check

`observations-update` och `observations-sync` visas inte som egna menyval i webben.
Observations/artikel_max uppdateras automatiskt när en buffertfil laddas in,
samma grundbeteende som i tkinter-appen.

## Kom igång

Engångsinstallation:

```powershell
pip install -r web/requirements.txt
npm install --prefix web/frontend
npm run build --prefix web/frontend
```

Starta desktopappen (pywebview-fönster):

```powershell
python web/backend/desktop.py
```

Eller använd `web/start_web.bat` som bygger frontenden och startar appen.

## Utvecklingsläge (hot reload)

Tva terminaler:

```powershell
# 1. API
python -m uvicorn api:app --reload --port 8765 --app-dir web/backend

# 2. Frontend (Vite proxar /api -> 8765)
npm run dev --prefix web/frontend
```

Öppna http://localhost:5173.

## Funktioner som behalls

- En central Datauppladdning för alla filer, inklusive WMS-filer för Eftersök
- Drag & drop i hela aktiva vyn med automatisk filtypsigenkanning
- Automatisk observations-/artikel_max-uppdatering när buffertpallar laddas in
- Filval per rad i Datauppladdning när en fil behöver väljas manuellt
- Temp-fil-export: "Öppna i Excel" öppnar resultatet lokalt
- "Ladda ner CSV" för webbappsläge
- Popups för fel, okänd filtyp och hjälp

## Status

Hela huvudflödet är exponerat. Drag&drop sorterar filer i hela aktiva vyn via
samma filtypsdetektering som tkinter-GUI:t. Resultat öppnas i Excel lokalt eller
laddas ner som CSV. Tema-toggle för ljust/mörkt läge.
