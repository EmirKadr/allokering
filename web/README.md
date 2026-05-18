# Allokering - Webb/API-demo

Ett modernt, API-styrt granssnitt for allokeringsflodet. Samma motor som
`allokering12.1.py` (GUI + CLI) - det nya skiktet ar ett rent
presentationslager ovanpa ett HTTP-API som ocksa kan deployas som webbapp.

## Arkitektur

```
web/
  backend/
    engine.py    - laddar motorn fran allokering12.1.py (delad logik)
    api.py       - FastAPI: /api/allocate, /api/detect, /api/open-excel, /api/download
    desktop.py   - pywebview-launcher (kor API + visar React-appen i ett fonster)
  frontend/      - React + Vite (drag&drop, filslots, resultattabs, popups)
```

CLI:t (`python allokering12.1.py allocate ...`) och tkinter-GUI:t ar oroda -
demon ar additiv och ligger helt i `web/`.

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

Demo v1 tacker `allocate`-flodet (resultat, near-miss, refill HP/AutoStore,
pallplatser). Ovriga CLI-kommandon kan kopplas in som fler endpoints enligt
samma monster.
