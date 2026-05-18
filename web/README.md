# Allokering - Webb/API-demo

Ett modernt, API-styrt granssnitt for allokeringsflodet. Samma motor som
`allokering12.1.py` (GUI + CLI) - det nya skiktet ar ett rent
presentationslager ovanpa ett HTTP-API som ocksa kan deployas som webbapp.

## Arkitektur

```
web/
  backend/
    engine.py    - laddar motorn fran allokering12.1.py (delad logik)
    flows.py     - ett floden-handtag per CLI-kommando + registret
    api.py       - FastAPI: /api/flows, /api/flow/{id}, /api/detect,
                   /api/open-excel, /api/download
    desktop.py   - pywebview-launcher (kor API + visar React-appen i ett fonster)
  frontend/      - React + Vite (sidebar, drag&drop, filslots, resultattabs,
                   dark/light-tema, popups)
```

Frontenden byggs dynamiskt fran `/api/flows` - lagg till ett nytt flode i
`flows.py` sa dyker det upp i menyn automatiskt.

CLI:t (`python allokering12.1.py allocate ...`) och tkinter-GUI:t ar oroda -
demon ar additiv och ligger helt i `web/`.

## Floden

Alla 14 CLI-kommandon finns som floden i menyn:

- **Allokering:** allocate
- **Order & saldo:** ordersaldo, lyx, pafyllnadsprio
- **Kontroller:** hib-koppling, overview-check, dispatch-check, vecka27-check
- **Sokning & prognos:** eftersok, prognos-report
- **Data & verktyg:** observations-update, observations-sync, split-values, update-check

## Kom igang

Engangsinstallation:

```powershell
pip install -r web/requirements.txt
npm install --prefix web/frontend
npm run build --prefix web/frontend
```

Starta desktop-demon (pywebview-fonster):

```powershell
python web/backend/desktop.py
```

Eller anvand `web/start_web.bat` som bygger frontenden och startar appen.

## Utvecklingslage (hot reload)

Tva terminaler:

```powershell
# 1. API
python -m uvicorn api:app --reload --port 8765 --app-dir web/backend

# 2. Frontend (Vite proxar /api -> 8765)
npm run dev --prefix web/frontend
```

Oppna http://localhost:5173.

## Funktioner som behalls

- Drag & drop med automatisk filtypsigenkanning (samma logik som GUI:t)
- Filval per ruta (klick eller drop)
- Temp-fil-export: "Oppna i Excel" oppnar resultatet lokalt
- "Ladda ner CSV" for webbappslage
- Popups for fel, okand filtyp och hjalp

## Status

Hela appen ar exponerad - samtliga 14 floden. Drag&drop sorterar filer till
ratt ruta via samma filtypsdetektering som tkinter-GUI:t. Resultat oppnas i
Excel lokalt eller laddas ner som CSV. Tema-toggle for ljust/morkt lage.

`observations-update` och `observations-sync` skriver till temporara filer -
repo-datan ror demon aldrig.
