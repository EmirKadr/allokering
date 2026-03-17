#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py ? FastAPI-webbapp f?r Allokering.
Serverar frontend statiskt och exponerar REST-API + SSE.
"""

from __future__ import annotations

import asyncio
import csv
from dataclasses import dataclass, field
import mimetypes
import os
import random
import re
import shutil
import subprocess
import tempfile
import threading
import time
import unicodedata
import uuid
from urllib.parse import unquote, urlparse
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests as req
from fastapi import BackgroundTasks, FastAPI, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import sys
sys.path.insert(0, str(Path(__file__).parent))

from session_store import SessionData, create_session, delete_session, get_session
try:
    from classifier_v2 import router as classifier_v2_router
    _classifier_v2_import_error: Optional[Exception] = None
except Exception as e:
    classifier_v2_router = None
    _classifier_v2_import_error = e
from logic import (
    _clean_columns,
    _reclassify_skrymmande,
    allocate,
    apply_value_filters,
    calculate_refill,
    compute_hib_koppling,
    compute_missed_departures,
    compute_pallet_spaces,
    find_col,
    normalize_items,
    normalize_saldo,
    read_csv_auto,
    refresh_ordersaldo,
    save_df_to_excel,
    scan_filter_values,
    ORDER_SCHEMA,
)

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="Allokering WebApp")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# v2 classifier router (kept isolated from legacy v1 endpoints)
if classifier_v2_router is not None:
    app.include_router(classifier_v2_router)
else:
    print(f"[WARN] classifier_v2 disabled: {_classifier_v2_import_error}")

ASK_CSV_AGENT_URL = os.environ.get("ASK_CSV_AGENT_URL", "http://127.0.0.1:8010").rstrip("/")
ASK_CSV_DEFAULT_URL = os.environ.get(
    "ASK_CSV_DEFAULT_URL",
    "https://noeffectui-frey.nowastelogistics.com/desktop",
)
ASK_CSV_LOGIN_WAIT_SECONDS = int(os.environ.get("ASK_CSV_LOGIN_WAIT_SECONDS", "3600"))
ASK_FETCHABLE_FILE_KEYS = {
    "orders",
    "buffer",
    "automation",
    "item",
    "overview",
    "dispatch",
    "wms_receive",
    "wms_booking",
    "wms_trans",
    "wms_pick",
    "wms_correct",
}
CLASSIFIER_APP_DIR = os.environ.get("CLASSIFIER_APP_DIR", r"C:\artikelplacering\Artikelplacering").strip()
CLASSIFIER_SCRIPT = os.environ.get("CLASSIFIER_SCRIPT", "classifier.py").strip()
_classifier_lock = threading.Lock()
_classifier_proc: Optional[subprocess.Popen] = None
CLASSIFIER_DEFAULT_IMAGE_DIR = os.environ.get(
    "CLASSIFIER_DEFAULT_IMAGE_DIR",
    str(Path(CLASSIFIER_APP_DIR) / "bilder"),
).strip()
CLASSIFIER_DEFAULT_OUTPUT_DIR = os.environ.get(
    "CLASSIFIER_DEFAULT_OUTPUT_DIR",
    CLASSIFIER_APP_DIR,
).strip()
CLASSIFIER_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff"}
CLASSIFIER_DATA_FILE_KEYS = {"item", "item_alias", "item_attribute", "main_category"}
CLASSIFIER_WEB_DATA_DIR = Path(
    os.environ.get(
        "CLASSIFIER_WEB_DATA_DIR",
        str(Path(tempfile.gettempdir()) / "allok_classifier_data"),
    )
)


@dataclass
class WebClassifierSession:
    session_id: str
    test_name: str
    image_dir: str
    output_dir: str
    categories: List[str]
    images: List[str]
    image_source: str = "file"  # "file" | "url"
    image_rows: List[Dict[str, str]] = field(default_factory=list)  # for url mode
    counts: Dict[str, int] = field(default_factory=dict)
    index: int = 0
    skipped: int = 0
    finished: bool = False
    created_at: float = field(default_factory=time.time)


_web_classifier_lock = threading.Lock()
_web_classifier_sessions: Dict[str, WebClassifierSession] = {}
_classifier_data_lock = threading.Lock()
_classifier_data_files: Dict[str, str] = {}

# ---------------------------------------------------------------------------
# Pydantic modeller
# ---------------------------------------------------------------------------

class FilterBody(BaseModel):
    bolag: List[str] = []
    ordertyp: List[str] = []


class EftersokBody(BaseModel):
    purchase: str = ""
    article: str = ""


class ChunkedExcelBody(BaseModel):
    values: List[str] = []
    chunk_size: int = 2000


class AskCsvInitBody(BaseModel):
    url: str = ASK_CSV_DEFAULT_URL
    login_wait: int = ASK_CSV_LOGIN_WAIT_SECONDS
    headless: bool = False
    slow_mo: int = 80
    goto_timeout: int = 60


class AskCsvFetchBody(BaseModel):
    view_name: str = "Order\u00f6versikt"
    target_file_key: str = "overview"
    open_via: str = "shortcut"
    open_text: str = "Visa"
    export_text: str = "Exportera till CSV"
    grid_wait: int = 30
    download_timeout: int = 120
    output_name: str = ""


class WebClassifierStartBody(BaseModel):
    test_name: str = ""
    categories: List[str] = []
    image_dir: str = ""
    output_dir: str = ""
    shuffle: bool = False


class WebClassifierClassifyBody(BaseModel):
    category: str = ""


def _parse_content_disposition_filename(header_value: str) -> str:
    """Extract filename from Content-Disposition header."""
    if not header_value:
        return ""
    m = re.search(r"filename\*=UTF-8''([^;]+)", header_value, flags=re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    m = re.search(r'filename="([^"]+)"', header_value, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"filename=([^;]+)", header_value, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip().strip('"')
    return ""


def _safe_csv_name(name: str, fallback: str = "export.csv") -> str:
    safe = re.sub(r'[^\w.\-]+', "_", (name or "").strip())
    if not safe:
        safe = fallback
    if not safe.lower().endswith(".csv"):
        safe += ".csv"
    return safe


SLOT_FILENAME_HINTS: Dict[str, List[str]] = {
    "orders": ["customer_order_details", "order_details_all", "detalj"],
    "overview": ["order_overview", "overview"],
    "buffer": ["buffertpallet", "buffert", "buffer"],
    "dispatch": ["dispatch"],
    "wms_receive": ["receive_log", "receive"],
    "wms_booking": ["booking_putaway", "booking", "putaway"],
    "wms_trans": ["trans_log", "translog"],
    "wms_pick": ["pick_log_full", "pick_log", "plocklogg"],
    "wms_correct": ["correct_log", "correct", "saldojustering"],
}


def _csv_filename_matches_slot(filename: str, file_key: str) -> bool:
    hints = SLOT_FILENAME_HINTS.get(file_key, [])
    if not hints:
        return True
    lower = (filename or "").lower()
    if not lower:
        return True
    return any(h in lower for h in hints)


def _repair_mojibake_text(value: Any) -> str:
    text = str(value or "")
    if any(ch in text for ch in ("Ã", "Â", "â")):
        try:
            repaired = text.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
            if repaired:
                text = repaired
        except Exception:
            pass
    return text


def _norm_col_key(value: Any) -> str:
    text = _repair_mojibake_text(value).lower().strip()
    text = text.replace("?", "a")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text)


def _find_col_by_keywords(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    if df is None or not hasattr(df, "columns") or len(df.columns) == 0:
        return None
    normalized_cols = {col: _norm_col_key(col) for col in df.columns}
    normalized_keywords: List[str] = []
    for k in keywords:
        nk = _norm_col_key(k)
        if nk:
            normalized_keywords.append(nk)

    for kw in normalized_keywords:
        for col, nk in normalized_cols.items():
            if nk == kw:
                return col
    for kw in normalized_keywords:
        for col, nk in normalized_cols.items():
            if kw in nk or nk in kw:
                return col
    return None


def _classifier_script_path() -> Path:
    base = Path(CLASSIFIER_APP_DIR)
    return (base / CLASSIFIER_SCRIPT).resolve()


def _ensure_classifier_data_dir() -> Path:
    CLASSIFIER_WEB_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return CLASSIFIER_WEB_DATA_DIR


def _classifier_data_files_payload() -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    with _classifier_data_lock:
        for key in sorted(CLASSIFIER_DATA_FILE_KEYS):
            path = _classifier_data_files.get(key)
            exists = bool(path and os.path.exists(path))
            out[key] = {
                "exists": exists,
                "filename": os.path.basename(path) if exists and path else "",
                "path": path if exists and path else "",
            }
    return out


def _read_csv_loose(path: str) -> pd.DataFrame:
    # More tolerant read for uploaded data files.
    for enc in ("utf-8-sig", "latin1"):
        try:
            return pd.read_csv(path, dtype=str, sep=None, engine="python", encoding=enc)
        except Exception:
            continue
    # Last fallback
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig")


def _pick_col_by_names(df: pd.DataFrame, names: List[str]) -> Optional[str]:
    nmap: Dict[str, str] = {}
    for c in df.columns:
        nk = _norm_col_key(c)
        if nk and nk not in nmap:
            nmap[nk] = c
    nkeys = [_norm_col_key(n) for n in names if _norm_col_key(n)]
    for nk in nkeys:
        if nk in nmap:
            return nmap[nk]
    for nk in nkeys:
        for ck, c in nmap.items():
            if nk in ck or ck in nk:
                return c
    return None


def _extract_rows_from_item_attribute(path: str) -> List[Dict[str, str]]:
    try:
        df = _read_csv_loose(path)
    except Exception:
        return []
    if df is None or df.empty:
        return []

    art_col = _pick_col_by_names(df, ["Artikel", "Artikelnummer", "Artikelnr", "article"])
    name_col = _pick_col_by_names(df, ["Namn", "Name", "Attribut", "Attribute"])
    val_col = _pick_col_by_names(df, ["Värde", "Varde", "Value", "Val"])
    bolag_col = _pick_col_by_names(df, ["Bolag", "Company"])
    if not art_col or not val_col:
        return []

    rows: List[Dict[str, str]] = []
    seen = set()
    for _, r in df.iterrows():
        article = str(r.get(art_col, "") or "").strip()
        value = str(r.get(val_col, "") or "").strip()
        name = str(r.get(name_col, "") or "").strip().lower() if name_col else ""
        bolag = str(r.get(bolag_col, "") or "").strip() if bolag_col else ""
        if not article or not value.lower().startswith("http"):
            continue
        if name_col and name not in {"img", "image", "bild", "url"}:
            continue
        key = (article, bolag, value)
        if key in seen:
            continue
        seen.add(key)
        rows.append(
            {
                "article_number": article,
                "url": value,
                "bolag": bolag,
            }
        )
    return rows


def _download_url_bytes(url: str, timeout_sec: int = 40) -> tuple[bytes, str]:
    resp = req.get(url, timeout=timeout_sec)
    if not resp.ok or not resp.content:
        raise RuntimeError(f"Kunde inte ladda bild-URL ({resp.status_code}): {url}")
    content_type = (resp.headers.get("content-type") or "").split(";")[0].strip()
    return resp.content, (content_type or "application/octet-stream")


def _filename_from_url_row(row: Dict[str, str], index: int, content_type: str) -> str:
    article = _safe_folder_fragment(str(row.get("article_number", "") or "").strip())
    if not article:
        article = f"image_{index + 1}"
    url_path = urlparse(str(row.get("url", "") or "")).path
    ext = Path(url_path).suffix.lower()
    if ext not in CLASSIFIER_IMAGE_EXTENSIONS:
        guessed = mimetypes.guess_extension(content_type or "")
        ext = (guessed or ".jpg").lower()
    if not ext.startswith("."):
        ext = f".{ext}"
    return f"{article}{ext}"


def _classifier_status_payload() -> Dict[str, Any]:
    script = _classifier_script_path()
    available = script.exists()

    running = False
    pid: Optional[int] = None
    global _classifier_proc
    with _classifier_lock:
        if _classifier_proc is not None:
            if _classifier_proc.poll() is None:
                running = True
                pid = int(_classifier_proc.pid) if _classifier_proc.pid else None
            else:
                _classifier_proc = None

    if not available:
        message = "Hittar inte classifier.py i CLASSIFIER_APP_DIR."
    elif running:
        message = "Classifier är igång."
    else:
        message = "Classifier är stoppad."

    return {
        "ok": True,
        "available": available,
        "running": running,
        "pid": pid,
        "script_path": str(script),
        "message": message,
    }


def _safe_folder_fragment(name: str) -> str:
    text = str(name or "").strip()
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or "unnamed"


def _unique_target_path(dst_dir: Path, filename: str) -> Path:
    target = dst_dir / filename
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    for i in range(2, 10000):
        candidate = dst_dir / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
    return dst_dir / f"{stem}_{int(time.time())}{suffix}"


def _build_web_classifier_state(session: WebClassifierSession) -> Dict[str, Any]:
    total = len(session.images)
    done = bool(session.finished or session.index >= total)
    current_path = session.images[session.index] if not done else ""
    current_name = Path(current_path).name if current_path else ""
    current_article = ""
    current_source_url = ""
    if session.image_source == "url" and not done and session.index < len(session.image_rows):
        row = session.image_rows[session.index]
        current_article = str(row.get("article_number", "") or "").strip()
        current_source_url = str(row.get("url", "") or "").strip()
        if current_article:
            current_name = current_article

    return {
        "ok": True,
        "session_id": session.session_id,
        "test_name": session.test_name,
        "image_dir": session.image_dir,
        "output_dir": session.output_dir,
        "image_source": session.image_source,
        "categories": session.categories,
        "counts": session.counts,
        "skipped": session.skipped,
        "index": session.index,
        "total": total,
        "done": done,
        "current_filename": current_name,
        "current_article": current_article,
        "current_source_url": current_source_url,
        "image_url": (
            f"/api/classifier/web/{session.session_id}/image?ts={int(time.time() * 1000)}"
            if not done
            else ""
        ),
    }


def _find_web_classifier_session(session_id: str) -> Optional[WebClassifierSession]:
    with _web_classifier_lock:
        return _web_classifier_sessions.get(session_id)


# ---------------------------------------------------------------------------
# Session endpoints
# ---------------------------------------------------------------------------

@app.post("/api/session")
def api_create_session():
    s = create_session()
    return {"session_id": s.session_id}


@app.delete("/api/session/{sid}")
def api_delete_session(sid: str):
    delete_session(sid)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Upload endpoints
# ---------------------------------------------------------------------------

@app.post("/api/upload/{sid}")
async def api_upload(sid: str, file_key: str = Form(...), file: UploadFile = Form(...)):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    safe_name = re.sub(r"[^\w.\-]", "_", file.filename or "file")
    dest = os.path.join(session.temp_dir, f"{file_key}_{safe_name}")
    content = await file.read()
    with open(dest, "wb") as f:
        f.write(content)
    session.files[file_key] = dest
    return {"ok": True, "file_key": file_key, "filename": safe_name}


@app.delete("/api/upload/{sid}/{file_key}")
def api_remove_file(sid: str, file_key: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    path = session.files.pop(file_key, None)
    if path and os.path.exists(path):
        os.remove(path)
    return {"ok": True}


@app.get("/api/upload/{sid}")
def api_list_files(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    result: Dict[str, Optional[str]] = {}
    for k, v in list(session.files.items()):
        if v and os.path.exists(v):
            result[k] = os.path.basename(v)
        else:
            # Avoid stale "loaded" state in frontend when temp files were removed.
            session.files[k] = None
            result[k] = None
    return {"files": result}


# ---------------------------------------------------------------------------
# Filter endpoints
# ---------------------------------------------------------------------------

@app.get("/api/filters/{sid}")
def api_get_filter_options(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    combined: Dict[str, List[str]] = {}
    for path in session.files.values():
        if not path or not os.path.exists(path):
            continue
        try:
            df = read_csv_auto(path)
            vals = scan_filter_values(df)
            for fk, fvals in vals.items():
                existing = set(combined.get(fk, []))
                existing.update(fvals)
                combined[fk] = sorted(existing)
        except Exception as e:
            print(f"WARN filter-scan misslyckades för '{path}': {e}")
    return combined


@app.post("/api/filters/{sid}")
def api_set_filters(sid: str, body: FilterBody):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    session.active_filters = {"bolag": body.bolag, "ordertyp": body.ordertyp}
    return {"ok": True}


# ---------------------------------------------------------------------------
# ASK CSV bridge (external listener -> session files)
# ---------------------------------------------------------------------------

@app.get("/api/ask-csv/health")
def api_ask_csv_health():
    try:
        resp = req.get(f"{ASK_CSV_AGENT_URL}/health", timeout=10)
        if not resp.ok:
            raise HTTPException(status_code=502, detail=f"CSV-agent svarade {resp.status_code}")
        return resp.json()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Kunde inte n? CSV-agenten: {e}") from e


@app.post("/api/ask-csv/init")
def api_ask_csv_init(body: AskCsvInitBody):
    effective_payload = body.dict()
    effective_payload["url"] = (effective_payload.get("url") or ASK_CSV_DEFAULT_URL).strip() or ASK_CSV_DEFAULT_URL
    effective_payload["login_wait"] = ASK_CSV_LOGIN_WAIT_SECONDS
    timeout_sec = max(20, int(effective_payload["login_wait"]) + 30)
    try:
        resp = req.post(
            f"{ASK_CSV_AGENT_URL}/init-login",
            json=effective_payload,
            timeout=timeout_sec,
        )
        if not resp.ok:
            detail = resp.text[:500]
            raise HTTPException(status_code=502, detail=f"CSV-agent init misslyckades: {detail}")
        return resp.json()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Kunde inte initiera CSV-agent: {e}") from e


@app.post("/api/ask-csv/fetch/{sid}")
def api_ask_csv_fetch(sid: str, body: AskCsvFetchBody):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")

    if body.target_file_key not in ASK_FETCHABLE_FILE_KEYS:
        raise HTTPException(status_code=400, detail=f"Ogiltig target_file_key: {body.target_file_key}")

    payload = {
        "view_name": body.view_name,
        "open_via": "shortcut",
        "open_text": body.open_text,
        "export_text": body.export_text,
        "grid_wait": body.grid_wait,
        "download_timeout": body.download_timeout,
        "output_name": body.output_name,
    }

    timeout_sec = max(30, body.download_timeout + 60)
    try:
        resp = req.post(
            f"{ASK_CSV_AGENT_URL}/fetch-csv",
            json=payload,
            timeout=timeout_sec,
        )
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Kunde inte n? CSV-agenten: {e}") from e

    if not resp.ok:
        detail = resp.text[:1000]
        raise HTTPException(status_code=502, detail=f"CSV-agent fetch misslyckades: {detail}")

    if not resp.content:
        raise HTTPException(status_code=502, detail="CSV-agent returnerade tom fil")

    header_name = _parse_content_disposition_filename(resp.headers.get("content-disposition", ""))
    if header_name and not _csv_filename_matches_slot(header_name, body.target_file_key):
        raise HTTPException(
            status_code=502,
            detail=(
                "CSV-agent returnerade fel vy-fil för slot "
                f"'{body.target_file_key}': '{header_name}'."
            ),
        )
    safe_name = _safe_csv_name(body.output_name or header_name or f"{body.target_file_key}.csv")
    dest = os.path.join(session.temp_dir, f"{body.target_file_key}_{safe_name}")

    with open(dest, "wb") as f:
        f.write(resp.content)

    session.files[body.target_file_key] = dest
    return {
        "ok": True,
        "file_key": body.target_file_key,
        "filename": os.path.basename(dest),
        "bytes": len(resp.content),
    }


# ---------------------------------------------------------------------------
# Classifier bridge (external desktop app)
# ---------------------------------------------------------------------------

@app.get("/api/classifier/status")
def api_classifier_status():
    return _classifier_status_payload()


@app.post("/api/classifier/start")
def api_classifier_start():
    script = _classifier_script_path()
    if not script.exists():
        raise HTTPException(status_code=404, detail=f"Classifier-script saknas: {script}")

    global _classifier_proc
    with _classifier_lock:
        if _classifier_proc is not None and _classifier_proc.poll() is None:
            return {
                "ok": True,
                "already_running": True,
                "pid": int(_classifier_proc.pid) if _classifier_proc.pid else None,
                "message": "Classifier kör redan.",
            }

        try:
            creationflags = 0
            if os.name == "nt":
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

            _classifier_proc = subprocess.Popen(
                [sys.executable, str(script)],
                cwd=str(script.parent),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Kunde inte starta classifier: {e}") from e

        return {
            "ok": True,
            "started": True,
            "pid": int(_classifier_proc.pid) if _classifier_proc.pid else None,
            "message": f"Classifier startad (PID {_classifier_proc.pid}).",
        }


@app.post("/api/classifier/stop")
def api_classifier_stop():
    global _classifier_proc
    with _classifier_lock:
        if _classifier_proc is None or _classifier_proc.poll() is not None:
            _classifier_proc = None
            return {"ok": True, "stopped": False, "message": "Classifier var redan stoppad."}

        pid = int(_classifier_proc.pid) if _classifier_proc.pid else None
        try:
            _classifier_proc.terminate()
            _classifier_proc.wait(timeout=8)
        except Exception:
            try:
                _classifier_proc.kill()
            except Exception:
                pass
        finally:
            _classifier_proc = None

        return {"ok": True, "stopped": True, "pid": pid, "message": f"Classifier stoppad (PID {pid})."}


# ---------------------------------------------------------------------------
# Classifier web flow
# ---------------------------------------------------------------------------

@app.get("/api/classifier/web/config")
def api_classifier_web_config():
    return {
        "ok": True,
        "default_image_dir": CLASSIFIER_DEFAULT_IMAGE_DIR,
        "default_output_dir": CLASSIFIER_DEFAULT_OUTPUT_DIR,
        "supported_extensions": sorted(CLASSIFIER_IMAGE_EXTENSIONS),
        "data_files": _classifier_data_files_payload(),
    }


@app.get("/api/classifier/web/data-files")
def api_classifier_web_data_files():
    payload = _classifier_data_files_payload()
    item_attribute = payload.get("item_attribute", {})
    item_rows = 0
    if item_attribute.get("exists") and item_attribute.get("path"):
        item_rows = len(_extract_rows_from_item_attribute(str(item_attribute["path"])))
    return {
        "ok": True,
        "files": payload,
        "item_attribute_rows": item_rows,
    }


@app.post("/api/classifier/web/data-files/upload")
async def api_classifier_web_data_upload(file_key: str = Form(...), file: UploadFile = Form(...)):
    if file_key not in CLASSIFIER_DATA_FILE_KEYS:
        raise HTTPException(status_code=400, detail=f"Ogiltig file_key: {file_key}")
    base = _ensure_classifier_data_dir()
    safe_name = re.sub(r"[^\w.\-]", "_", file.filename or f"{file_key}.csv")
    dest = base / f"{file_key}_{safe_name}"
    content = await file.read()
    with open(dest, "wb") as fh:
        fh.write(content)
    with _classifier_data_lock:
        old = _classifier_data_files.get(file_key)
        _classifier_data_files[file_key] = str(dest)
    if old and old != str(dest):
        try:
            if os.path.exists(old):
                os.remove(old)
        except Exception:
            pass
    payload = _classifier_data_files_payload()
    item_rows = 0
    if file_key == "item_attribute":
        item_rows = len(_extract_rows_from_item_attribute(str(dest)))
    return {
        "ok": True,
        "file_key": file_key,
        "filename": os.path.basename(dest),
        "bytes": len(content),
        "item_attribute_rows": item_rows,
        "files": payload,
    }


@app.delete("/api/classifier/web/data-files/{file_key}")
def api_classifier_web_data_delete(file_key: str):
    if file_key not in CLASSIFIER_DATA_FILE_KEYS:
        raise HTTPException(status_code=400, detail=f"Ogiltig file_key: {file_key}")
    old = None
    with _classifier_data_lock:
        old = _classifier_data_files.pop(file_key, None)
    if old and os.path.exists(old):
        try:
            os.remove(old)
        except Exception:
            pass
    return {"ok": True, "file_key": file_key, "files": _classifier_data_files_payload()}


@app.post("/api/classifier/web/start")
def api_classifier_web_start(body: WebClassifierStartBody):
    test_name = str(body.test_name or "").strip()
    if not test_name:
        raise HTTPException(status_code=400, detail="Testnamn krävs.")

    categories: List[str] = []
    seen = set()
    for raw in body.categories or []:
        cat = str(raw or "").strip()
        if not cat:
            continue
        key = cat.lower()
        if key in seen:
            continue
        seen.add(key)
        categories.append(cat)
    if not categories:
        raise HTTPException(status_code=400, detail="Minst en kategori krävs.")

    output_dir = Path((body.output_dir or CLASSIFIER_DEFAULT_OUTPUT_DIR).strip()).expanduser()
    if not output_dir.exists() or not output_dir.is_dir():
        raise HTTPException(status_code=400, detail=f"Outputmapp saknas: {output_dir}")

    image_dir = Path((body.image_dir or CLASSIFIER_DEFAULT_IMAGE_DIR).strip()).expanduser()
    images: List[str] = []
    if image_dir.exists() and image_dir.is_dir():
        images = [
            str(p.resolve())
            for p in sorted(image_dir.rglob("*"))
            if p.is_file() and p.suffix.lower() in CLASSIFIER_IMAGE_EXTENSIONS
        ]

    image_rows: List[Dict[str, str]] = []
    item_attr_path = ""
    with _classifier_data_lock:
        item_attr_path = _classifier_data_files.get("item_attribute", "") or ""
    if item_attr_path and os.path.exists(item_attr_path):
        image_rows = _extract_rows_from_item_attribute(item_attr_path)

    image_source = "file"
    if images:
        if body.shuffle:
            random.shuffle(images)
    elif image_rows:
        image_source = "url"
        images = [str(r.get("url", "") or "").strip() for r in image_rows if str(r.get("url", "") or "").strip()]
        if body.shuffle:
            combo = list(zip(images, image_rows))
            random.shuffle(combo)
            images = [c[0] for c in combo]
            image_rows = [c[1] for c in combo]
    else:
        if image_dir.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Inga bilder hittades i {image_dir} och ingen item_attribute med IMG-URL är uppladdad.",
            )
        raise HTTPException(
            status_code=400,
            detail=f"Bildmapp saknas: {image_dir}. Ladda upp datafiler eller ange korrekt bildmapp.",
        )

    cls_id = uuid.uuid4().hex[:12]
    session = WebClassifierSession(
        session_id=cls_id,
        test_name=test_name,
        image_dir=str(image_dir.resolve()) if image_dir.exists() else str(image_dir),
        output_dir=str(output_dir.resolve()),
        categories=categories,
        images=images,
        image_source=image_source,
        image_rows=image_rows,
        counts={c: 0 for c in categories},
    )
    with _web_classifier_lock:
        _web_classifier_sessions[cls_id] = session

    state = _build_web_classifier_state(session)
    if image_source == "url":
        state["message"] = f"Klassificering startad med {len(images)} bilder från item_attribute."
    else:
        state["message"] = f"Klassificering startad med {len(images)} bilder."
    return state


@app.get("/api/classifier/web/{cls_id}")
def api_classifier_web_state(cls_id: str):
    session = _find_web_classifier_session(cls_id)
    if not session:
        raise HTTPException(status_code=404, detail="Klassificeringssession saknas.")
    return _build_web_classifier_state(session)


@app.get("/api/classifier/web/{cls_id}/image")
def api_classifier_web_image(cls_id: str):
    session = _find_web_classifier_session(cls_id)
    if not session:
        raise HTTPException(status_code=404, detail="Klassificeringssession saknas.")
    if session.finished or session.index >= len(session.images):
        raise HTTPException(status_code=404, detail="Ingen aktiv bild kvar.")

    if session.image_source == "url":
        if session.index >= len(session.image_rows):
            raise HTTPException(status_code=404, detail="Ingen aktiv bildrad kvar.")
        row = session.image_rows[session.index]
        url = str(row.get("url", "") or "").strip()
        if not url:
            raise HTTPException(status_code=404, detail="Bild-URL saknas i aktuell rad.")
        try:
            content, content_type = _download_url_bytes(url, timeout_sec=40)
        except Exception as e:
            raise HTTPException(status_code=502, detail=f"Kunde inte hämta bild från URL: {e}") from e
        return Response(content=content, media_type=content_type)

    path = Path(session.images[session.index])
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"Bilden saknas: {path.name}")
    media_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
    return FileResponse(str(path), media_type=media_type, filename=path.name)


@app.post("/api/classifier/web/{cls_id}/classify")
def api_classifier_web_classify(cls_id: str, body: WebClassifierClassifyBody):
    category = str(body.category or "").strip()
    if not category:
        raise HTTPException(status_code=400, detail="Kategori krävs.")

    with _web_classifier_lock:
        session = _web_classifier_sessions.get(cls_id)
        if not session:
            raise HTTPException(status_code=404, detail="Klassificeringssession saknas.")
        if session.finished or session.index >= len(session.images):
            raise HTTPException(status_code=400, detail="Sessionen är redan klar.")
        if category not in session.categories:
            raise HTTPException(status_code=400, detail=f"Okänd kategori: {category}")

        dst_dir = Path(session.output_dir) / f"{_safe_folder_fragment(session.test_name)}.{_safe_folder_fragment(category)}"
        dst_dir.mkdir(parents=True, exist_ok=True)

        saved_path = ""
        if session.image_source == "url":
            if session.index >= len(session.image_rows):
                raise HTTPException(status_code=400, detail="Ingen aktiv URL-rad kvar.")
            row = session.image_rows[session.index]
            url = str(row.get("url", "") or "").strip()
            if not url:
                raise HTTPException(status_code=400, detail="Aktuell rad saknar URL.")
            try:
                content, content_type = _download_url_bytes(url, timeout_sec=40)
            except Exception as e:
                raise HTTPException(status_code=502, detail=f"Kunde inte ladda bild för klassificering: {e}") from e
            fname = _filename_from_url_row(row, session.index, content_type)
            dst = _unique_target_path(dst_dir, fname)
            with open(dst, "wb") as fh:
                fh.write(content)
            saved_path = str(dst)
        else:
            src = Path(session.images[session.index])
            if not src.exists():
                session.skipped += 1
                session.index += 1
                if session.index >= len(session.images):
                    session.finished = True
                state = _build_web_classifier_state(session)
                state["message"] = f"Bilden saknades och hoppades över: {src.name}"
                return state
            dst = _unique_target_path(dst_dir, src.name)
            shutil.copy2(src, dst)
            saved_path = str(dst)

        session.counts[category] = int(session.counts.get(category, 0)) + 1
        session.index += 1
        if session.index >= len(session.images):
            session.finished = True
        state = _build_web_classifier_state(session)
        state["saved_to"] = saved_path
        state["message"] = f"Sparad till {dst_dir.name}"
        return state


@app.post("/api/classifier/web/{cls_id}/skip")
def api_classifier_web_skip(cls_id: str):
    with _web_classifier_lock:
        session = _web_classifier_sessions.get(cls_id)
        if not session:
            raise HTTPException(status_code=404, detail="Klassificeringssession saknas.")
        if session.finished or session.index >= len(session.images):
            raise HTTPException(status_code=400, detail="Sessionen är redan klar.")
        session.skipped += 1
        session.index += 1
        if session.index >= len(session.images):
            session.finished = True
        state = _build_web_classifier_state(session)
        state["message"] = "Bild hoppades över."
        return state


@app.post("/api/classifier/web/{cls_id}/finish")
def api_classifier_web_finish(cls_id: str):
    with _web_classifier_lock:
        session = _web_classifier_sessions.get(cls_id)
        if not session:
            raise HTTPException(status_code=404, detail="Klassificeringssession saknas.")
        session.finished = True
        state = _build_web_classifier_state(session)
        state["message"] = "Session avslutad."
        return state


# ---------------------------------------------------------------------------
# Run status
# ---------------------------------------------------------------------------

@app.get("/api/run/status/{sid}")
def api_run_status(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    return {
        "running": session.running,
        "results": list(session.results.keys())
    }


# ---------------------------------------------------------------------------
# Job-k?rningar
# ---------------------------------------------------------------------------

def _start_job(session: SessionData, coro):
    """Starta ett bakgrundsjobb."""
    if session.running:
        return False
    session.running = True
    asyncio.ensure_future(coro)
    return True


async def _job_allokering(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        orders_path = session.files.get("orders")
        buffer_path = session.files.get("buffer")
        automation_path = session.files.get("automation")
        item_path = session.files.get("item")

        if not orders_path or not buffer_path:
            log("FEL: orders och buffer m?ste vara uppladdade.")
            log("__ERROR__")
            return

        log("L?ser in filer...")
        orders_raw = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
        buffer_raw = await loop.run_in_executor(None, lambda: read_csv_auto(buffer_path))

        saldo_norm = None
        saldo_raw = None
        if automation_path and os.path.exists(automation_path):
            auto_raw = await loop.run_in_executor(None, lambda: read_csv_auto(automation_path))
            auto_raw = _clean_columns(auto_raw.copy())
            saldo_raw = auto_raw.copy()
            saldo_norm = await loop.run_in_executor(None, lambda: normalize_saldo(auto_raw))

        item_norm = None
        if item_path and os.path.exists(item_path):
            try:
                item_raw = await loop.run_in_executor(None, lambda: read_csv_auto(item_path))
                item_norm = await loop.run_in_executor(None, lambda: normalize_items(item_raw))
            except Exception as ie:
                log(f"Varning: Kunde inte l?sa item-fil: {ie}")

        orders_raw = _clean_columns(orders_raw)
        buffer_raw = _clean_columns(buffer_raw)

        # Applicera filter
        orders_raw = apply_value_filters(orders_raw, session.active_filters)
        buffer_raw = apply_value_filters(buffer_raw, session.active_filters)
        if saldo_raw is not None:
            saldo_raw = apply_value_filters(saldo_raw, session.active_filters)
            saldo_norm = await loop.run_in_executor(None, lambda: normalize_saldo(saldo_raw))

        log("K?r allokering (Helpall ? AutoStore ? Huvudplock, FIFO)...")
        result, near = await loop.run_in_executor(None, lambda: allocate(orders_raw, buffer_raw, log=log))

        result = await loop.run_in_executor(None, lambda: _reclassify_skrymmande(result, saldo_norm))

        # Sl? ihop item-fil
        if item_norm is not None and not item_norm.empty and not result.empty:
            try:
                art_col_res = find_col(result, ORDER_SCHEMA["artikel"], required=True)
                temp_merge = result.merge(item_norm, how="left", left_on=art_col_res, right_on="Artikel", suffixes=("", "_item"))
                for drop_col in ["Artikel_item", "Artikel_y", "Ej Staplingsbar_y"]:
                    if drop_col in temp_merge.columns:
                        temp_merge.drop(columns=[drop_col], inplace=True, errors="ignore")
                if "Ej Staplingsbar_x" in temp_merge.columns:
                    temp_merge["Ej Staplingsbar"] = temp_merge["Ej Staplingsbar_x"].fillna("")
                    temp_merge.drop(columns=["Ej Staplingsbar_x"], inplace=True, errors="ignore")
                if "Ej Staplingsbar" not in temp_merge.columns:
                    temp_merge["Ej Staplingsbar"] = ""
                result = temp_merge
            except Exception as e:
                log(f"Kunde inte sl? ihop item-fil: {e}")

        if "Ej Staplingsbar" not in result.columns:
            result["Ej Staplingsbar"] = ""

        log("Skapar Excel-filer...")

        # Spara allokerade
        allok_path = os.path.join(session.temp_dir, "allokerade.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(result, "allokerade", allok_path))
        session.results["allokerade"] = allok_path
        log("__RESULT:allokerade__")

        # Spara near-miss
        if not near.empty:
            nearmiss_path = os.path.join(session.temp_dir, "nearmiss.xlsx")
            await loop.run_in_executor(None, lambda: save_df_to_excel(near, "nearmiss", nearmiss_path))
            session.results["nearmiss"] = nearmiss_path
            log("__RESULT:nearmiss__")

        # Pallplatser
        try:
            pallet_spaces = await loop.run_in_executor(None, lambda: compute_pallet_spaces(result))
            if pallet_spaces is not None and not pallet_spaces.empty:
                ps_path = os.path.join(session.temp_dir, "pallplatser.xlsx")
                await loop.run_in_executor(None, lambda: save_df_to_excel(pallet_spaces, "pallplatser", ps_path))
                session.results["pallplatser"] = ps_path
                log("__RESULT:pallplatser__")
        except Exception as e:
            log(f"Pallplatser kunde inte ber?knas: {e}")

        # Refill
        try:
            hp_df, as_df = await loop.run_in_executor(None, lambda: calculate_refill(
                result, buffer_raw,
                saldo_df=saldo_norm,
                not_putaway_df=None
            ))
            has_refill = (hp_df is not None and not hp_df.empty) or (as_df is not None and not as_df.empty)
            if has_refill:
                refill_sheets = {}
                if hp_df is not None and not hp_df.empty:
                    refill_sheets["P?fyllning HP"] = hp_df
                if as_df is not None and not as_df.empty:
                    refill_sheets["P?fyllning AutoStore"] = as_df
                refill_path = os.path.join(session.temp_dir, "refill.xlsx")
                await loop.run_in_executor(None, lambda: save_df_to_excel(refill_sheets, "refill", refill_path))
                session.results["refill"] = refill_path
                log("__RESULT:refill__")
                log(f"Auto-refill klar: HP {len(hp_df)} rader, AUTOSTORE {len(as_df)} rader.")
        except Exception as e:
            log(f"Refill misslyckades: {e}")

        # Summering per zon
        try:
            zon_col = "Zon (ber?knad)"
            qty_col = find_col(result, ORDER_SCHEMA["qty"], required=True)
            summary = result.groupby(zon_col)[qty_col].apply(
                lambda s: pd.to_numeric(s, errors="coerce").sum()).reset_index(name="Totalt antal")
            log("\nSummering per zon:")
            for _, r in summary.iterrows():
                log(f"  Zon {r[zon_col]}: {r['Totalt antal']:.0f}")
        except Exception:
            pass

        # Summering per K?lltyp (samma som desktop-appen)
        try:
            import json as _json
            qty_col_kt = find_col(result, ORDER_SCHEMA["qty"], required=False, default=None)
            ktyp_series = result.get("K\u00e4lltyp", pd.Series([], dtype=object)).astype(str)
            unique_types = [k for k in sorted(set(ktyp_series.dropna())) if k]
            ordered_types = []
            for prv in ("HELPALL", "AUTOSTORE"):
                if prv in unique_types:
                    ordered_types.append(prv)
                    unique_types.remove(prv)
            ordered_types.extend(unique_types)
            kt_rows = []
            log("\nSummering per K\u00e4lltyp:")
            for ktyp in ordered_types:
                sub = result[ktyp_series == ktyp]
                row_count = int(len(sub))
                kolli = 0.0
                if qty_col_kt and not sub.empty:
                    kolli = float(pd.to_numeric(sub[qty_col_kt], errors="coerce").sum())
                if ktyp == "HELPALL":
                    row_text = f"{row_count} pallar"
                else:
                    row_text = f"{row_count} rader"
                kolli_int = int(round(kolli))
                log(f"  {ktyp}: {row_text}, {kolli_int} kolli")
                kt_rows.append({"kalltyp": ktyp, "antal_text": row_text, "kolli": kolli_int})
            log(f"__KALLTYP_SUMMARY:{_json.dumps(kt_rows)}__")
        except Exception:
            pass

        # Ber?kna ordersaldo-listor (kompletta ordrar / p?fyllningsbehov)
        try:
            if orders_path:
                list1, list2 = await loop.run_in_executor(
                    None,
                    lambda: refresh_ordersaldo(orders_path, session.active_filters),
                )
                session.ordersaldo_list1 = list1
                session.ordersaldo_list2 = list2
                if list1 or list2:
                    log(f"Ordersaldo: {len(list1)} kompletta ordrar, {len(list2)} artiklar med p?fyllningsbehov.")
        except Exception as e:
            log(f"Varning: Ordersaldo-ber?kning misslyckades: {e}")

        log("Allokeringen ?r klar.")
        log("__DONE__")
    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
    finally:
        session.running = False


async def _job_hib_koppling(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        orders_path = session.files.get("orders")
        overview_path = session.files.get("overview")

        if not orders_path or not overview_path:
            log("FEL: V?lj b?de best?llningslinjer och order?versikt.")
            log("__ERROR__")
            return

        log("L?ser in filer f?r HIB-koppling...")
        details_df = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
        overview_df = await loop.run_in_executor(None, lambda: read_csv_auto(overview_path))

        details_df = apply_value_filters(details_df, session.active_filters)
        overview_df = apply_value_filters(overview_df, session.active_filters)

        log("Ber?knar HIB-koppling...")
        changes_df = await loop.run_in_executor(None, lambda: compute_hib_koppling(details_df, overview_df))
        missed_df = await loop.run_in_executor(None, lambda: compute_missed_departures(details_df, overview_df))

        has_changes = isinstance(changes_df, pd.DataFrame) and not changes_df.empty
        has_missed = isinstance(missed_df, pd.DataFrame) and not missed_df.empty

        if not has_changes and not has_missed:
            log("Inga HIB-ordrar beh?ver ?ndras eller har missat sin avg?ng.")
            log("__DONE__")
            return

        instr_lines = [
            "?ndras i f?ljande ordning",
            "1. Ordernummer",
            "2. S?ndningsnummer",
            "3. Zon F p? orderlinjerna",
            "4. Samma multi p? alla Hibar till samma butik",
            "5. Generera",
            "6. Frisl?pp",
        ]
        instructions_df = pd.DataFrame({"Instruktioner": instr_lines})

        sheets: Dict[str, pd.DataFrame] = {}
        if has_changes:
            sheets["?ndringar"] = changes_df
            log(f"HIB-koppling: {len(changes_df)} ordrar att ?ndra.")
            for _, r in changes_df.iterrows():
                ordnr = str(r.get("ordernummer", "")).strip()
                kundnamn = str(r.get("kundnamn", "")).strip()
                fields = []
                if str(r.get("s?ndningsnummer", "")).strip():
                    fields.append(f"S?ndningsnr ? {str(r['s?ndningsnummer']).strip()}")
                if str(r.get("Orderdatum", "")).strip():
                    fields.append(f"Orderdatum ? {str(r['Orderdatum']).strip()}")
                if str(r.get("Zon", "")).strip():
                    fields.append(f"Zon ? {str(r['Zon']).strip()}")
                if str(r.get("Multi", "")).strip():
                    fields.append(f"Multi ? {str(r['Multi']).strip()}")
                if fields:
                    name_part = f" ({kundnamn})" if kundnamn else ""
                    log(f"Order {ordnr}{name_part}: {', '.join(fields)}")

        if has_missed:
            sheets["Missade avg?ngar"] = missed_df
            log(f"Missade avg?ngar: {len(missed_df)} st.")
            for _, r in missed_df.iterrows():
                ordnr = str(r.get("ordernummer", "")).strip()
                kundnamn = str(r.get("kundnamn", "")).strip()
                name_part = f" ({kundnamn})" if kundnamn else ""
                log(f"Order {ordnr}{name_part}: MISSAT SIN AVG?NG")

        sheets["Instruktion"] = instructions_df

        hib_path = os.path.join(session.temp_dir, "hib_koppling.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(sheets, "hib_koppling", hib_path))
        session.results["hib-koppling"] = hib_path
        log("__RESULT:hib-koppling__")
        log("HIB-kopplingen ?r ber?knad.")
        log("__DONE__")
    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
    finally:
        session.running = False


async def _job_orderkontroll(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        overview_path = session.files.get("overview")
        if not overview_path:
            log("FEL: V?lj order?versikten f?rst.")
            log("__ERROR__")
            return

        log("L?ser in order?versikt...")
        df = await loop.run_in_executor(None, lambda: read_csv_auto(overview_path))
        df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
        df = apply_value_filters(df, session.active_filters)

        if df.empty:
            log("Inga rader kvar efter filter i order?versikten.")
            log("__DONE__")
            return

        # Hitta kolumner robust (klarar teckenkodningsavvikelser).
        ship_col = _find_col_by_keywords(
            df,
            [
                "Sändningsnr",
                "Sändnings nr",
                "Sändningsnummer",
                "sandningsnr",
                "sandningsnummer",
                "shipment",
            ],
        )

        if not ship_col:
            log("FEL: Kunde inte identifiera s?ndningsnummer-kolumnen.")
            log("__ERROR__")
            return

        cust_col = _find_col_by_keywords(
            df,
            ["kundnr", "kundnummer", "kund nr", "kund", "customer", "customer number"],
        )

        if not cust_col:
            log("FEL: Kunde inte identifiera kund-kolumnen.")
            log("__ERROR__")
            return

        trans_col = _find_col_by_keywords(df, ["transportör", "transportor", "carrier", "transport"])
        if not trans_col:
            trans_col = "__transport_dummy__"
            df[trans_col] = ""

        order_col = _find_col_by_keywords(
            df,
            ["ordernr", "order nr", "ordernummer", "order number", "orderid", "order"],
        )

        # Bygg kundnamns-mapping
        order_to_customer: Dict[str, str] = {}
        orders_path = session.files.get("orders")
        if orders_path and os.path.exists(orders_path):
            try:
                ddf = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
                ddf.columns = [str(c).replace("\ufeff", "").strip() for c in ddf.columns]
                ddf = apply_value_filters(ddf, session.active_filters)
                det_order_col = _find_col_by_keywords(
                    ddf,
                    ["order nr", "ordernr", "ordernummer", "order number"],
                )
                det_customer_col = _find_col_by_keywords(
                    ddf,
                    ["kund.1", "kund1", "kund nr", "kund", "customer", "customer name"],
                )
                if det_order_col and det_customer_col:
                    order_to_customer = (
                        ddf.groupby(det_order_col)[det_customer_col]
                        .first()
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .to_dict()
                    )
            except Exception:
                pass

        df[ship_col] = df[ship_col].astype(str).str.strip()
        df[cust_col] = df[cust_col].astype(str).str.strip()
        df[trans_col] = df[trans_col].astype(str).str.strip()
        df = df[df[ship_col].astype(str).str.len() > 0].copy()

        if df.empty:
            log("Order?versikten inneh?ller inga s?ndningsnummer.")
            log("__DONE__")
            return

        # Avvikelsekontroll
        shipment_diff_rows: List[Dict[str, Any]] = []
        for ship, group in df.groupby(ship_col):
            try:
                customers = sorted(set(group[cust_col].dropna().astype(str).str.strip()))
                carriers = sorted(set(group[trans_col].dropna().astype(str).str.strip()))
                customers = [c for c in customers if c]
                carriers = [t for t in carriers if t]
                orders_list: List[str] = []
                if order_col:
                    order_vals = sorted(set(group[order_col].dropna().astype(str).str.strip()))
                    for o in order_vals:
                        nm = order_to_customer.get(o, "")
                        orders_list.append(f"{o} ({nm})" if nm else o)
                orders_str = ", ".join(orders_list)
                if len(customers) > 1 or len(carriers) > 1:
                    row: Dict[str, Any] = {
                        "Avvikelsetyp": "S?ndningsnr med flera kunder/transport?rer",
                        "S?ndningsnr": ship,
                        "Unika kunder": len(customers),
                        "Kunder": ", ".join(customers),
                        "Unika transport?rer": len(carriers),
                        "Transport?rer": ", ".join(carriers),
                        "Antal orderrader": int(len(group)),
                    }
                    if orders_str:
                        row["Ordernr (kundnamn)"] = orders_str
                    shipment_diff_rows.append(row)
            except Exception:
                continue

        result_df = pd.DataFrame(shipment_diff_rows) if shipment_diff_rows else pd.DataFrame()

        # HIB-kontroll
        ordertype_col = _find_col_by_keywords(df, ["ordertyp", "order type", "ordertype"])
        status_col = _find_col_by_keywords(df, ["status", "orderstatus", "radstatus", "state"])

        def _to_status_num(value) -> Optional[int]:
            try:
                raw = str(value).strip().replace(",", ".")
                if not raw:
                    return None
                return int(float(raw))
            except Exception:
                return None

        hib_rows: List[Dict[str, Any]] = []
        missing_hib_cols: List[str] = []
        if not order_col:
            missing_hib_cols.append("ordernummer")
        if not ordertype_col:
            missing_hib_cols.append("ordertyp")
        if not status_col:
            missing_hib_cols.append("status")

        if not missing_hib_cols:
            try:
                hib_df = df[[order_col, ship_col, cust_col, ordertype_col, status_col]].copy()
                hib_df[order_col] = hib_df[order_col].astype(str).str.strip()
                hib_df[ship_col] = hib_df[ship_col].astype(str).str.strip()
                hib_df[cust_col] = hib_df[cust_col].astype(str).str.strip()
                hib_df["_ordertype_norm"] = hib_df[ordertype_col].astype(str).str.strip().str.upper()
                hib_df["_status_num"] = hib_df[status_col].apply(_to_status_num)

                store_mask = hib_df["_ordertype_norm"].eq("N") | hib_df["_ordertype_norm"].str.contains("BUTIK", na=False)
                store_ships = set(hib_df.loc[store_mask, ship_col].dropna().astype(str).str.strip().tolist())
                store_ships.discard("")

                hib_only_df = hib_df[hib_df["_ordertype_norm"].str.contains("HIB", na=False)].copy()
                for ordnr, group in hib_only_df.groupby(order_col):
                    ordnr_str = str(ordnr).strip()
                    if not ordnr_str:
                        continue
                    status_values = [s for s in group["_status_num"].tolist() if s is not None]
                    if not status_values:
                        continue
                    max_status = max(status_values)
                    if max_status <= 31:
                        continue
                    hib_ships = sorted(set(group[ship_col].dropna().astype(str).str.strip()))
                    hib_ships = [s for s in hib_ships if s]
                    if not hib_ships:
                        continue
                    if any(ship_val in store_ships for ship_val in hib_ships):
                        continue
                    kundnamn = order_to_customer.get(ordnr_str, "")
                    if not kundnamn:
                        kunder = [k for k in group[cust_col].dropna().astype(str).str.strip().tolist() if k]
                        if kunder:
                            kundnamn = kunder[0]
                    row2: Dict[str, Any] = {
                        "Ordernr": ordnr_str,
                        "S?ndningsnr": ", ".join(hib_ships),
                        "Ordertyp": "HIB",
                        "Status": int(max_status),
                        "Anm?rkning": "HIB-order med status > 31 saknar matchande butikss?ndning",
                    }
                    if kundnamn:
                        row2["Kundnamn"] = kundnamn
                    hib_rows.append(row2)
            except Exception as e:
                log(f"HIB-kontrollen misslyckades delvis: {e}")

        hib_check_df = pd.DataFrame(hib_rows) if hib_rows else pd.DataFrame()

        has_any = not result_df.empty or not hib_check_df.empty
        if not has_any:
            msg = "Inga avvikelser hittades i order?versikten."
            if missing_hib_cols:
                msg += " HIB-kontrollen kunde inte k?ras fullt ut (saknar: " + ", ".join(missing_hib_cols) + ")."
            log(msg)
            log("__DONE__")
            return

        # Logga resultat
        if not result_df.empty:
            log(f"Order?versikt: {len(result_df)} s?ndningsnummer med flera kunder/transport?rer.")
        if not hib_check_df.empty:
            log(f"HIB-ordrar med status > 31 utan matchande butikss?ndning: {len(hib_check_df)} st.")
        if missing_hib_cols:
            log("HIB-kontrollen kunde inte k?ras fullt ut (saknar: " + ", ".join(missing_hib_cols) + ").")

        # Bygg Excel
        sheets: Dict[str, pd.DataFrame] = {}
        combined_parts = []
        if not result_df.empty:
            s_df = result_df.copy()
            if "Avvikelsetyp" not in s_df.columns:
                s_df.insert(0, "Avvikelsetyp", "S?ndningsnr med flera kunder/transport?rer")
            sheets["S?ndningskontroll"] = s_df
            combined_parts.append(s_df)
        if not hib_check_df.empty:
            h_df = hib_check_df.copy()
            if "Avvikelsetyp" not in h_df.columns:
                h_df.insert(0, "Avvikelsetyp", "HIB ?ver status 31 utan butikss?ndning")
            sheets["HIB utan butikss?ndning"] = h_df
            combined_parts.append(h_df)
        if combined_parts:
            combined = pd.concat(combined_parts, ignore_index=True, sort=False)
            sheets = {"Orderkontroll": combined, **sheets}

        ok_path = os.path.join(session.temp_dir, "orderkontroll.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(sheets, "orderkontroll", ok_path))
        session.results["orderkontroll"] = ok_path
        log("__RESULT:orderkontroll__")
        log("Orderkontrollen ?r klar.")
        log("__DONE__")
    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
    finally:
        session.running = False


async def _job_dispatchkontroll(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        overview_path = session.files.get("overview")
        dispatch_path = session.files.get("dispatch")

        if not overview_path or not dispatch_path:
            log("FEL: V?lj b?de order?versikt och dispatchpallar.")
            log("__ERROR__")
            return

        log("L?ser in filer f?r dispatchkontroll...")
        ov_df = await loop.run_in_executor(None, lambda: read_csv_auto(overview_path))
        dp_df = await loop.run_in_executor(None, lambda: read_csv_auto(dispatch_path))

        ov_df.columns = [str(c).replace("\ufeff", "").strip() for c in ov_df.columns]
        dp_df.columns = [str(c).replace("\ufeff", "").strip() for c in dp_df.columns]

        ov_df = apply_value_filters(ov_df, session.active_filters)
        dp_df = apply_value_filters(dp_df, session.active_filters)

        if ov_df.empty or dp_df.empty:
            log("Inga rader kvar efter filter.")
            log("__DONE__")
            return

        order_kws = ["ordernr", "order nr", "ordernummer", "order number", "orderid"]
        ship_kws = ["Sändningsnr", "Sändnings nr", "Sändningsnummer", "sandningsnr", "sandningsnummer", "shipment"]
        plock_kws = ["plockpallsnr", "plockpallsnr.", "plockpall", "plockpallnr", "plockpallsnummer", "plockpall nr"]

        ov_order_col = _find_col_by_keywords(ov_df, order_kws)
        ov_ship_col = _find_col_by_keywords(ov_df, ship_kws)
        if not ov_order_col or not ov_ship_col:
            log("FEL: Kunde inte identifiera order- eller s?ndningskolumnen i order?versikten.")
            log("__ERROR__")
            return

        dp_order_col = _find_col_by_keywords(dp_df, order_kws)
        dp_ship_col = _find_col_by_keywords(dp_df, ship_kws)
        plock_col = _find_col_by_keywords(dp_df, plock_kws)
        if not dp_order_col or not dp_ship_col or not plock_col:
            log("FEL: Kunde inte identifiera order-, s?ndnings- eller plockpallskolumnen i dispatchfilen.")
            log("__ERROR__")
            return

        ov_df[ov_order_col] = ov_df[ov_order_col].astype(str).str.strip()
        ov_df[ov_ship_col] = ov_df[ov_ship_col].astype(str).str.strip()
        dp_df[dp_order_col] = dp_df[dp_order_col].astype(str).str.strip()
        dp_df[dp_ship_col] = dp_df[dp_ship_col].astype(str).str.strip()
        dp_df[plock_col] = dp_df[plock_col].astype(str).str.strip()

        # Bygg kundnamns-mapping
        order_to_customer: Dict[str, str] = {}
        orders_path = session.files.get("orders")
        if orders_path and os.path.exists(orders_path):
            try:
                det_df = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
                det_df.columns = [str(c).replace("\ufeff", "").strip() for c in det_df.columns]
                det_df = apply_value_filters(det_df, session.active_filters)
                det_order_col = _find_col_by_keywords(
                    det_df,
                    ["order nr", "ordernr", "ordernummer", "order number"],
                )
                det_customer_col = _find_col_by_keywords(
                    det_df,
                    ["kund.1", "kund1", "kund nr", "kund", "customer", "customer name"],
                )
                if det_order_col and det_customer_col:
                    order_to_customer = (
                        det_df.groupby(det_order_col)[det_customer_col]
                        .first()
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .to_dict()
                    )
            except Exception:
                pass

        # Bygg order ? s?ndningsnr mapping fr?n order?versikten
        order_to_ship: Dict[str, str] = {}
        try:
            for ordnum, sub in ov_df.groupby(ov_order_col):
                ships = [s for s in sub[ov_ship_col] if isinstance(s, str) and s.strip()]
                if ships:
                    order_to_ship[str(ordnum)] = ships[0].strip()
        except Exception:
            pass

        diff_rows: List[Dict[str, Any]] = []
        for _, row in dp_df.iterrows():
            try:
                ordnr = str(row[dp_order_col]).strip()
                dp_ship = str(row[dp_ship_col]).strip()
                expected = order_to_ship.get(ordnr)
                if expected and expected != dp_ship:
                    diff_row: Dict[str, Any] = {
                        "Ordernr": ordnr,
                        "?versikt s?ndningsnr": expected,
                        "Dispatch s?ndningsnr": dp_ship,
                        "Plockpallsnr": str(row[plock_col]).strip(),
                        "kundnamn": order_to_customer.get(ordnr, ""),
                    }
                    diff_rows.append(diff_row)
            except Exception:
                continue

        if not diff_rows:
            log("Alla s?ndningsnummer st?mmer ?verens.")
            log("__DONE__")
            return

        diff_df = pd.DataFrame(diff_rows)
        log(f"Dispatchkontrollen hittade {len(diff_df)} avvikelser.")
        for _, row in diff_df.iterrows():
            name_part = f" ({row['kundnamn']})" if str(row.get("kundnamn", "")).strip() else ""
            log(f"Order {row['Ordernr']}{name_part}: s?ndningsnr {row['?versikt s?ndningsnr']} i ?versikten men {row['Dispatch s?ndningsnr']} i dispatch (plockpall {row['Plockpallsnr']})")

        dk_path = os.path.join(session.temp_dir, "dispatchkontroll.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel({"Dispatchkontroll": diff_df}, "dispatchkontroll", dk_path))
        session.results["dispatchkontroll"] = dk_path
        log("__RESULT:dispatchkontroll__")
        log("Dispatchkontrollen ?r klar.")
        log("__DONE__")
    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
    finally:
        session.running = False


async def _job_eftersok(session: SessionData, purchase: str, article: str):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        purchase = purchase.strip()
        article = article.strip()

        if not purchase or not article:
            log("FEL: Ange b?de ink?psnummer och artikelnummer.")
            log("__ERROR__")
            return

        # Samla alla tillg?ngliga WMS-filer
        wms_files: Dict[str, str] = {}
        for key in ["wms_receive", "wms_booking", "wms_buffert", "wms_trans", "wms_pick", "wms_correct"]:
            path = session.files.get(key)
            if path and os.path.exists(path):
                wms_files[key] = path

        log(f"Efters?k: ink?psnummer={purchase}, artikelnummer={article}")
        log(f"Tillg?ngliga WMS-filer: {list(wms_files.keys())}")

        results_sheets: Dict[str, pd.DataFrame] = {}

        # S?k igenom varje WMS-fil
        for key, path in wms_files.items():
            try:
                df = await loop.run_in_executor(None, lambda p=path: read_csv_auto(p))
                # S?k efter rader som matchar purchase eller article
                mask_purchase = pd.Series(False, index=df.index)
                mask_article = pd.Series(False, index=df.index)
                for col in df.columns:
                    col_str = df[col].astype(str)
                    if purchase:
                        mask_purchase = mask_purchase | col_str.str.contains(str(purchase), case=False, na=False)
                    if article:
                        mask_article = mask_article | col_str.str.contains(str(article), case=False, na=False)
                mask = mask_purchase | mask_article
                hits = df[mask].copy()
                if not hits.empty:
                    results_sheets[key] = hits
                    log(f"  {key}: {len(hits)} tr?ff(ar)")
                else:
                    log(f"  {key}: inga tr?ffar")
            except Exception as e:
                log(f"  {key}: fel vid l?sning ? {e}")

        if not results_sheets:
            log("Inga tr?ffar hittades i tillg?ngliga WMS-filer.")
            log("__DONE__")
            return

        eftersok_path = os.path.join(session.temp_dir, "eftersok.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(results_sheets, "eftersok", eftersok_path))
        session.results["eftersok"] = eftersok_path
        log("__RESULT:eftersok__")
        log("Efters?ket ?r klart.")
        log("__DONE__")
    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
    finally:
        session.running = False


# ---------------------------------------------------------------------------
# Run-endpoints
# ---------------------------------------------------------------------------

@app.post("/api/run/allokering/{sid}")
async def api_run_allokering(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En k?rning p?g?r redan")
    session.running = True
    background_tasks.add_task(_job_allokering, session)
    return {"job_id": "allokering", "status": "started"}


@app.post("/api/run/hib-koppling/{sid}")
async def api_run_hib(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En k?rning p?g?r redan")
    session.running = True
    background_tasks.add_task(_job_hib_koppling, session)
    return {"job_id": "hib-koppling", "status": "started"}


@app.post("/api/run/orderkontroll/{sid}")
async def api_run_orderkontroll(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En k?rning p?g?r redan")
    session.running = True
    background_tasks.add_task(_job_orderkontroll, session)
    return {"job_id": "orderkontroll", "status": "started"}


@app.post("/api/run/dispatchkontroll/{sid}")
async def api_run_dispatchkontroll(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En k?rning p?g?r redan")
    session.running = True
    background_tasks.add_task(_job_dispatchkontroll, session)
    return {"job_id": "dispatchkontroll", "status": "started"}


@app.post("/api/run/eftersok/{sid}")
async def api_run_eftersok(sid: str, body: EftersokBody, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En k?rning p?g?r redan")
    session.running = True
    background_tasks.add_task(_job_eftersok, session, body.purchase, body.article)
    return {"job_id": "eftersok", "status": "started"}


# ---------------------------------------------------------------------------
# Ordersaldo endpoints
# ---------------------------------------------------------------------------

@app.post("/api/ordersaldo/refresh/{sid}")
async def api_ordersaldo_refresh(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    orders_path = session.files.get("orders")
    if not orders_path or not os.path.exists(orders_path):
        raise HTTPException(status_code=400, detail="Best?llningslinjer-filen saknas")
    loop = asyncio.get_event_loop()
    list1, list2 = await loop.run_in_executor(
        None,
        lambda: refresh_ordersaldo(orders_path, session.active_filters),
    )
    session.ordersaldo_list1 = list1
    session.ordersaldo_list2 = list2
    return {
        "list1_count": len(list1),
        "list2_count": len(list2),
    }


@app.get("/api/ordersaldo/list1/{sid}")
def api_ordersaldo_list1(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    return {"values": session.ordersaldo_list1}


@app.get("/api/ordersaldo/list2/{sid}")
def api_ordersaldo_list2(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    return {"values": session.ordersaldo_list2}


# ---------------------------------------------------------------------------
# Debug
# ---------------------------------------------------------------------------

@app.get("/api/debug/columns/{sid}/{file_key}")
def api_debug_columns(sid: str, file_key: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    path = session.files.get(file_key)
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Fil saknas")
    try:
        df = read_csv_auto(path)
        return {"columns": list(df.columns), "rows": len(df)}
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# Rensa cache
# ---------------------------------------------------------------------------

@app.post("/api/session/reset-results/{sid}")
async def api_reset_results(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    # Rensa resultatfiler fr?n disk
    for path in session.results.values():
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    session.results.clear()
    session.ordersaldo_list1 = []
    session.ordersaldo_list2 = []
    # Rensa logg-k?n
    while not session.log_queue.empty():
        try:
            session.log_queue.get_nowait()
        except Exception:
            break
    return {"ok": True}


@app.post("/api/session/reset-all/{sid}")
async def api_reset_all(sid: str):
    """Rensa alla uppladdade filer OCH alla resultat för sessionen."""
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    # Rensa uppladdade filer
    for path in list(session.files.values()):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    session.files.clear()
    # Rensa resultatfiler
    for path in list(session.results.values()):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    session.results.clear()
    session.ordersaldo_list1 = []
    session.ordersaldo_list2 = []
    while not session.log_queue.empty():
        try:
            session.log_queue.get_nowait()
        except Exception:
            break
    return {"ok": True}


# ---------------------------------------------------------------------------
# SSE-logg
# ---------------------------------------------------------------------------

@app.get("/api/log/stream/{sid}")
async def api_log_stream(sid: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")

    async def event_generator():
        while True:
            try:
                msg = await asyncio.wait_for(session.log_queue.get(), timeout=30.0)
                safe = msg.replace("\n", "\\n").replace("\r", "")
                yield f"data: {safe}\n\n"
            except asyncio.TimeoutError:
                yield ": keepalive\n\n"
            except asyncio.CancelledError:
                break
            except Exception:
                break

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


# ---------------------------------------------------------------------------
# Nedladdning
# ---------------------------------------------------------------------------

@app.get("/api/result/preview/{sid}/{result_key}")
def api_result_preview(sid: str, result_key: str, limit: int = 120, sheet: str = ""):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")

    path = session.results.get(result_key)
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail=f"Resultat '{result_key}' ej tillgangligt")

    safe_limit = max(1, min(int(limit or 120), 500))
    ext = Path(path).suffix.lower()

    try:
        sheet_names: List[str] = []
        sheet_name = ""

        if ext in {".xlsx", ".xlsm", ".xls"}:
            xl = pd.ExcelFile(path)
            sheet_names = [str(s) for s in (xl.sheet_names or [])]
            if not sheet_names:
                return {
                    "result_key": result_key,
                    "sheet_name": "",
                    "sheet_names": [],
                    "row_count": 0,
                    "total_rows": 0,
                    "columns": [],
                    "rows": [],
                }
            sheet_name = sheet if sheet in sheet_names else sheet_names[0]
            preview_df = pd.read_excel(path, sheet_name=sheet_name, dtype=str, nrows=safe_limit)
        elif ext == ".csv":
            sheet_name = "csv"
            preview_df = read_csv_auto(path).head(safe_limit)
        else:
            raise HTTPException(status_code=400, detail=f"Format stods inte for preview: {ext}")

        if preview_df is None:
            preview_df = pd.DataFrame()

        if not preview_df.empty:
            preview_df = preview_df.fillna("")
            preview_df.columns = [_repair_mojibake_text(c).strip() for c in preview_df.columns]
            try:
                preview_df = preview_df.map(_repair_mojibake_text)
            except Exception:
                for col in preview_df.columns:
                    preview_df[col] = preview_df[col].map(_repair_mojibake_text)

        columns = [str(c) for c in preview_df.columns]
        rows = preview_df.to_dict(orient="records") if not preview_df.empty else []
        row_count = len(rows)
        return {
            "result_key": result_key,
            "sheet_name": sheet_name,
            "sheet_names": sheet_names,
            "row_count": row_count,
            "total_rows": row_count,
            "columns": columns,
            "rows": rows,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Kunde inte lasa resultat-preview: {e}") from e

@app.get("/api/download/{sid}/{result_key}")
def api_download(sid: str, result_key: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    path = session.results.get(result_key)
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail=f"Resultat '{result_key}' ej tillg?ngligt")
    filename = os.path.basename(path)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


# ---------------------------------------------------------------------------
# Chunked Excel
# ---------------------------------------------------------------------------

@app.post("/api/chunked-excel")
async def api_chunked_excel(body: ChunkedExcelBody):
    loop = asyncio.get_event_loop()
    values = [v for v in body.values if str(v).strip()]
    chunk_size = max(1, body.chunk_size)

    if not values:
        raise HTTPException(status_code=400, detail="Inga v?rden")

    def build_excel() -> str:
        chunks = [values[i:i + chunk_size] for i in range(0, len(values), chunk_size)]
        data: Dict[str, list] = {}
        for idx, chunk in enumerate(chunks):
            col_name = f"Kolumn {idx + 1}"
            data[col_name] = chunk
        max_len = max(len(c) for c in data.values())
        for k in data:
            while len(data[k]) < max_len:
                data[k].append("")
        df = pd.DataFrame(data)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix="_chunked.xlsx")
        tmp_path = tmp.name
        tmp.close()
        save_df_to_excel(df, "chunked", tmp_path)
        return tmp_path

    path = await loop.run_in_executor(None, build_excel)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="chunked_values.xlsx",
    )


# ---------------------------------------------------------------------------
# Statiska filer ? montera SIST
# ---------------------------------------------------------------------------

_frontend_dir = Path(__file__).parent.parent / "frontend"
if _frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(_frontend_dir), html=True), name="frontend")
