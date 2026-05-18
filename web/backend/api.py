"""FastAPI-lager for allokerings-demon.

Exponerar exakt samma motor som GUI och CLI som ett HTTP-API. Frontenden
(React) ar ett rent presentationslager ovanpa detta - samma kontrakt
fungerar bade i pywebview-fonstret lokalt och som webbapp senare.
"""
from __future__ import annotations

import math
import tempfile
import traceback
import uuid
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import engine

app = FastAPI(title="Allokering API", version=engine.APP_VERSION)

# Tillater Vite-dev-servern (npm run dev) att prata med API:t under utveckling.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Resultat fran en korning halls i minnet sa "Oppna i Excel" kan ateranvanda dem.
SESSIONS: dict[str, dict[str, pd.DataFrame]] = {}

# Vilka result-nycklar som far oppnas i Excel och deras filnamnsetikett.
EXCEL_KEYS = {
    "result": "allokerat_resultat",
    "near_miss": "near_miss",
    "refill_hp": "refill_huvudplock",
    "refill_autostore": "refill_autostore",
    "pallet_spaces": "pallplatser",
}

NEAR_MISS_COLUMNS = [
    "Artikel", "OrderID", "OrderRad", "PallID", "Kallplats", "Mottagen",
    "Behov_vid_tillfallet", "Pall_kvantitet", "Skillnad",
    "Procentuell skillnad (%)", "Anledning", "Galler (INSTEAD R/A)",
]


# --- Hjalpfunktioner ---------------------------------------------------------

def _cell(value: object) -> str:
    """Gor ett DataFrame-varde JSON-sakert for tabellvisning."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    if isinstance(value, pd.Timestamp):
        return "" if pd.isna(value) else value.isoformat(sep=" ")
    text = str(value)
    return "" if text.lower() in ("nan", "nat", "none") else text


def _df_to_table(df: Optional[pd.DataFrame], preview_limit: int = 1000) -> dict:
    """Konvertera ett DataFrame till {columns, rows, row_count, truncated}."""
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
    """Spara en uppladdad fil till en temporar fil med bevarat filtillagg."""
    suffix = Path(upload.filename or "").suffix or ".csv"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(await upload.read())
    tmp.close()
    return Path(tmp.name)


# --- API-modeller ------------------------------------------------------------

class OpenExcelRequest(BaseModel):
    session_id: str
    key: str


# --- Endpoints ---------------------------------------------------------------

@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "version": engine.APP_VERSION, "title": engine.APP_TITLE}


@app.post("/api/detect")
async def detect(file: UploadFile = File(...)) -> dict:
    """Identifiera vilken filtyp en uppladdad fil ar (samma logik som drag&drop i GUI)."""
    path = await _save_upload(file)
    try:
        file_type = engine.detect_file_type(str(path))
    except Exception:
        file_type = None
    finally:
        path.unlink(missing_ok=True)
    # Mappa motorns typer till allocate-flodets slots.
    slot_map = {"orders": "orders", "buffer": "buffer", "automation": "saldo", "item": "items"}
    return {"file_type": file_type, "slot": slot_map.get(file_type or "")}


@app.post("/api/allocate")
async def run_allocate(
    orders: UploadFile = File(...),
    buffer: UploadFile = File(...),
    saldo: Optional[UploadFile] = File(None),
    items: Optional[UploadFile] = File(None),
    not_putaway: Optional[UploadFile] = File(None),
) -> dict:
    """Kor allokeringsflodet - samma motor som CLI-kommandot ``allocate``."""
    temp_paths: list[Path] = []
    log_lines: list[str] = []

    def _log(msg: str) -> None:
        log_lines.append(str(msg))

    try:
        orders_path = await _save_upload(orders)
        buffer_path = await _save_upload(buffer)
        temp_paths += [orders_path, buffer_path]

        orders_raw = engine.read_table(str(orders_path))
        buffer_raw = engine.read_table(str(buffer_path))

        saldo_norm = None
        if saldo is not None:
            p = await _save_upload(saldo)
            temp_paths.append(p)
            saldo_norm = engine.normalize_saldo(engine.read_table(str(p)))

        item_norm = None
        if items is not None:
            p = await _save_upload(items)
            temp_paths.append(p)
            item_norm = engine.normalize_items(engine.read_table(str(p)))

        not_putaway_norm = None
        if not_putaway is not None:
            p = await _save_upload(not_putaway)
            temp_paths.append(p)
            not_putaway_norm = engine.normalize_not_putaway(engine.read_table(str(p)))

        _log(f"Laser in filer: {len(temp_paths)} fil(er).")
        result_df, near_miss_df = engine.allocate(orders_raw, buffer_raw, log=_log)
        result_df = engine.reclassify_skrymmande(result_df, saldo_norm)
        result_df = engine.merge_item_flags(result_df, item_norm)

        if near_miss_df.empty and len(near_miss_df.columns) == 0:
            near_miss_df = pd.DataFrame(columns=NEAR_MISS_COLUMNS)

        refill_hp_df, refill_autostore_df = engine.calculate_refill(
            result_df, buffer_raw, saldo_df=saldo_norm, not_putaway_df=not_putaway_norm,
        )
        pallet_spaces_df = engine.compute_pallet_spaces(result_df)
        _log("Allokering klar.")
    except Exception as exc:  # noqa: BLE001 - vi vill exponera felet for UI:t
        for p in temp_paths:
            p.unlink(missing_ok=True)
        raise HTTPException(
            status_code=400,
            detail={"message": str(exc), "trace": traceback.format_exc()},
        )
    finally:
        for p in temp_paths:
            p.unlink(missing_ok=True)

    session_id = uuid.uuid4().hex
    SESSIONS[session_id] = {
        "result": result_df,
        "near_miss": near_miss_df,
        "refill_hp": refill_hp_df,
        "refill_autostore": refill_autostore_df,
        "pallet_spaces": pallet_spaces_df,
    }

    return {
        "session_id": session_id,
        "summary": {
            "result_rows": int(len(result_df)),
            "near_miss_rows": int(len(near_miss_df)),
            "refill_hp_rows": int(len(refill_hp_df)),
            "refill_autostore_rows": int(len(refill_autostore_df)),
            "pallet_space_rows": int(len(pallet_spaces_df)),
        },
        "tables": {
            "result": _df_to_table(result_df),
            "near_miss": _df_to_table(near_miss_df),
            "refill_hp": _df_to_table(refill_hp_df),
            "refill_autostore": _df_to_table(refill_autostore_df),
            "pallet_spaces": _df_to_table(pallet_spaces_df),
        },
        "log": log_lines,
    }


@app.post("/api/open-excel")
def open_excel(req: OpenExcelRequest) -> dict:
    """Skriv ett resultat till en temporar Excel/CSV och oppna det i OS:et.

    Fungerar lokalt (pywebview/desktop). Som ren webbapp senare byts detta
    mot en nedladdnings-endpoint - se /api/download.
    """
    session = SESSIONS.get(req.session_id)
    if session is None:
        raise HTTPException(status_code=404, detail="Sessionen hittades inte (kor allokeringen igen).")
    if req.key not in EXCEL_KEYS:
        raise HTTPException(status_code=400, detail=f"Okand resultatnyckel: {req.key}")
    df = session[req.key]
    path = engine.open_df_in_excel(df, label=EXCEL_KEYS[req.key])
    return {"opened": True, "path": path}


@app.get("/api/download/{session_id}/{key}")
def download(session_id: str, key: str):
    """Ladda ner ett resultat som CSV (webbappslage, ingen lokal Excel)."""
    session = SESSIONS.get(session_id)
    if session is None or key not in EXCEL_KEYS:
        raise HTTPException(status_code=404, detail="Resultatet hittades inte.")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    session[key].to_csv(tmp.name, index=False, encoding="utf-8-sig")
    tmp.close()
    return FileResponse(tmp.name, filename=f"{EXCEL_KEYS[key]}.csv", media_type="text/csv")


# --- Statiska filer (byggd React-frontend) -----------------------------------
# Maste mountas SIST sa att /api/*-routerna far foretrade.
_DIST = Path(__file__).resolve().parents[1] / "frontend" / "dist"
if _DIST.exists():
    app.mount("/", StaticFiles(directory=str(_DIST), html=True), name="frontend")
