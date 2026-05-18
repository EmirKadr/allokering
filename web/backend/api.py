"""FastAPI-lager for allokerings-demon.

Exponerar hela appen: ett generiskt floden-API dar varje CLI-kommando i
allokering12.1.py har en motsvarande endpoint. Frontenden (React) ar ett
rent presentationslager - samma kontrakt fungerar i pywebview-fonstret
lokalt och som webbapp senare.
"""
from __future__ import annotations

import math
import tempfile
import traceback
import uuid
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from starlette.datastructures import UploadFile as StarletteUploadFile

import engine
import flows

app = FastAPI(title="Allokering API", version=engine.APP_VERSION)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Resultat fran en korning halls i minnet sa Excel/CSV-export kan ateranvanda dem.
# SESSIONS[session_id] = {"tables": {key: DataFrame}, "labels": {key: label}}
SESSIONS: dict[str, dict] = {}


# --- Hjalpfunktioner ---------------------------------------------------------

def _cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        return str(int(value)) if value.is_integer() else f"{value:g}"
    if isinstance(value, pd.Timestamp):
        return "" if pd.isna(value) else value.isoformat(sep=" ")
    text = str(value)
    return "" if text.lower() in ("nan", "nat", "none") else text


def _df_to_table(df: Optional[pd.DataFrame], preview_limit: int = 1000) -> dict:
    if not isinstance(df, pd.DataFrame) or df.empty:
        cols = [str(c) for c in df.columns] if isinstance(df, pd.DataFrame) else []
        return {"columns": cols, "rows": [], "row_count": 0, "truncated": False}
    columns = [str(c) for c in df.columns]
    preview = df.head(preview_limit)
    rows = [[_cell(v) for v in rec] for rec in preview.itertuples(index=False, name=None)]
    return {
        "columns": columns,
        "rows": rows,
        "row_count": int(len(df)),
        "truncated": len(df) > preview_limit,
    }


async def _save_upload(upload: UploadFile) -> Path:
    suffix = Path(upload.filename or "").suffix or ".csv"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(await upload.read())
    tmp.close()
    return Path(tmp.name)


class OpenExcelRequest(BaseModel):
    session_id: str
    key: str


# --- Endpoints ---------------------------------------------------------------

@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "version": engine.APP_VERSION, "title": engine.APP_TITLE}


@app.get("/api/flows")
def list_flows() -> dict:
    """Floden-registret - frontenden bygger UI:t dynamiskt fran detta."""
    return {"flows": flows.public_registry()}


@app.post("/api/detect")
async def detect(file: UploadFile = File(...)) -> dict:
    """Identifiera filtyp (samma logik som GUI:ts drag&drop)."""
    path = await _save_upload(file)
    try:
        file_type = engine.detect_file_type(str(path))
    except Exception:
        file_type = None
    finally:
        path.unlink(missing_ok=True)
    return {"file_type": file_type}


@app.post("/api/flow/{flow_id}")
async def run_flow(flow_id: str, request: Request) -> dict:
    """Kor ett flode. Filer och textfalt skickas som multipart/form-data."""
    flow = flows.FLOW_BY_ID.get(flow_id)
    if flow is None:
        raise HTTPException(status_code=404, detail=f"Okant flode: {flow_id}")

    form = await request.form()
    files: dict[str, Path] = {}
    params: dict[str, str] = {}
    temp_paths: list[Path] = []
    try:
        for key, value in form.multi_items():
            if isinstance(value, StarletteUploadFile):
                if value.filename:
                    path = await _save_upload(value)
                    files[key] = path
                    temp_paths.append(path)
            elif isinstance(value, str) and value.strip() != "":
                params[key] = value

        result = flow["handler"](files, params)
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(
            status_code=400,
            detail={"message": str(exc), "trace": traceback.format_exc()},
        )
    finally:
        for path in temp_paths:
            path.unlink(missing_ok=True)

    tables = result.get("tables", [])
    session_id = uuid.uuid4().hex
    SESSIONS[session_id] = {
        "tables": {key: df for key, _label, df in tables},
        "labels": {key: label for key, label, _df in tables},
    }

    return {
        "flow_id": flow_id,
        "session_id": session_id,
        "summary": result.get("summary", {}),
        "tables": [
            {"key": key, "label": label, "table": _df_to_table(df)}
            for key, label, df in tables
        ],
        "text": result.get("text"),
        "log": result.get("log", []),
    }


@app.post("/api/open-excel")
def open_excel(req: OpenExcelRequest) -> dict:
    """Skriv ett resultat till temporar fil och oppna det i OS:et (desktop-lage)."""
    session = SESSIONS.get(req.session_id)
    if session is None or req.key not in session["tables"]:
        raise HTTPException(status_code=404, detail="Resultatet hittades inte (kor floden igen).")
    label = session["labels"].get(req.key, req.key)
    path = engine.open_df_in_excel(session["tables"][req.key], label=label)
    return {"opened": True, "path": path}


@app.get("/api/download/{session_id}/{key}")
def download(session_id: str, key: str):
    """Ladda ner ett resultat som CSV (webbappslage)."""
    session = SESSIONS.get(session_id)
    if session is None or key not in session["tables"]:
        raise HTTPException(status_code=404, detail="Resultatet hittades inte.")
    label = session["labels"].get(key, key)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    session["tables"][key].to_csv(tmp.name, index=False, encoding="utf-8-sig")
    tmp.close()
    return FileResponse(tmp.name, filename=f"{label}.csv", media_type="text/csv")


# --- Statiska filer (byggd React-frontend) -----------------------------------
# Maste mountas SIST sa att /api/*-routerna far foretrade.
_DIST = Path(__file__).resolve().parents[1] / "frontend" / "dist"
if _DIST.exists():
    app.mount("/", StaticFiles(directory=str(_DIST), html=True), name="frontend")
