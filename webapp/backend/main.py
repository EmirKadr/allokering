#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py - FastAPI-webbapp för Allokering.
Serverar frontend statiskt och exponerar REST-API + SSE.
"""

from __future__ import annotations

import asyncio
import csv
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
import uuid
from urllib.parse import unquote
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests as req
from fastapi import FastAPI, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import sys
sys.path.insert(0, str(Path(__file__).parent))

from db import init_db
from session_store import SessionData, create_session, delete_session, get_session
from job_queue import (
    clear_session_logs,
    enqueue_job,
    fetch_session_logs_since,
    get_session_status,
)
from worker_runtime import process_claimed_job
from allokering_result_cache import (
    ALLOKERING_RESULT_KEYS,
    get_prepared_results,
    invalidate_allokering_cache,
    load_dataframe as load_allokering_dataframe,
    load_manifest as load_allokering_manifest,
    save_manifest as save_allokering_manifest,
    store_dataframe as store_allokering_dataframe,
)
from auth import (
    VALID_LISTS,
    create_auth_session,
    delete_auth_session,
    extract_token,
    find_user,
    get_auth_session,
    load_users,
    require_admin,
    require_auth,
    save_users,
    ADMIN_CODE,
    ADMIN_EMAIL,
    ADMIN_USERNAME,
)
try:
    from gg_kontroll import router as gg_kontroll_router
    _gg_kontroll_import_error: Optional[Exception] = None
except Exception as e:
    gg_kontroll_router = None
    _gg_kontroll_import_error = e
from logic import (
    _clean_columns,
    _reclassify_skrymmande,
    allocate,
    apply_value_filters,
    build_prognos_vs_autoplock_report,
    calculate_refill,
    compute_hib_koppling,
    compute_missed_departures,
    compute_pallet_spaces,
    find_col,
    normalize_items,
    normalize_not_putaway,
    normalize_saldo,
    read_campaign_xlsx,
    read_csv_auto,
    read_prognos_xlsx,
    refresh_ordersaldo,
    save_df_to_excel,
    scan_filter_values,
    ORDER_SCHEMA,
)

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="Stigamo WebApp")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup_event():
    init_db()
    # Start embedded worker so jobs run even without a separate worker process
    asyncio.create_task(_embedded_worker_loop())


async def _embedded_worker_loop():
    """Lightweight in-process worker that claims and runs queued jobs."""
    from job_queue import claim_next_job, requeue_stale_jobs
    import socket, uuid as _uuid
    worker_id = f"embedded-{socket.gethostname()}-{_uuid.uuid4().hex[:6]}"
    while True:
        try:
            await asyncio.to_thread(requeue_stale_jobs)
            job = await asyncio.to_thread(claim_next_job, worker_id)
            if not job:
                await asyncio.sleep(1.0)
                continue
            await process_claimed_job(job, worker_id)
        except Exception:
            await asyncio.sleep(5.0)

if gg_kontroll_router is not None:
    app.include_router(gg_kontroll_router)
else:
    print(f"[WARN] gg_kontroll disabled: {_gg_kontroll_import_error}")

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
    "wms_buffer",
    "wms_booking",
    "wms_trans",
    "wms_pick",
    "wms_correct",
}
FILTER_SCAN_TEXT_EXTENSIONS = {".csv", ".txt", ".tsv"}
ORDERSALDO_INVALIDATION_KEYS = {"orders"}
ALLOKERING_INVALIDATION_KEYS = {"orders", "buffer", "automation", "item", "wms_booking"}


def _file_cache_signature(path: str) -> tuple[int, int]:
    stat_result = os.stat(path)
    mtime_ns = getattr(stat_result, "st_mtime_ns", int(stat_result.st_mtime * 1_000_000_000))
    return int(stat_result.st_size), int(mtime_ns)


def _clear_filter_cache_entry(session: SessionData, path: Optional[str]) -> None:
    if path:
        session.filter_scan_cache.pop(path, None)


def _invalidate_ordersaldo_cache(session: SessionData) -> None:
    session.ordersaldo_list1 = []
    session.ordersaldo_list2 = []


def _invalidate_result_cache(session: SessionData, result_key: str) -> None:
    path = session.results.pop(result_key, None)
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass


def _handle_session_file_change(session: SessionData, file_key: str, old_path: Optional[str] = None) -> None:
    _clear_filter_cache_entry(session, old_path)
    _clear_filter_cache_entry(session, session.files.get(file_key))
    if file_key in ORDERSALDO_INVALIDATION_KEYS:
        _invalidate_ordersaldo_cache(session)
    if file_key in ALLOKERING_INVALIDATION_KEYS:
        invalidate_allokering_cache(session)
    if file_key in {"prognos", "campaign", "automation", "buffer"}:
        _invalidate_result_cache(session, "prognos")


def _scan_filter_values_for_path(session: SessionData, path: str) -> Dict[str, List[str]]:
    cached = session.filter_scan_cache.get(path)
    signature = _file_cache_signature(path)
    if cached and cached[0] == signature:
        return cached[1]
    df = read_csv_auto(path)
    values = scan_filter_values(df)
    session.filter_scan_cache[path] = (signature, values)
    return values


def _collect_filter_options(session: SessionData) -> Dict[str, List[str]]:
    combined: Dict[str, List[str]] = {}
    active_paths: set[str] = set()
    for path in session.files.values():
        if not path or not os.path.exists(path):
            continue
        active_paths.add(path)
        if Path(path).suffix.lower() not in FILTER_SCAN_TEXT_EXTENSIONS:
            continue
        try:
            vals = _scan_filter_values_for_path(session, path)
            for fk, fvals in vals.items():
                existing = set(combined.get(fk, []))
                existing.update(fvals)
                combined[fk] = sorted(existing)
        except Exception as e:
            print(f"WARN filter-scan misslyckades för '{path}': {e}")

    for cached_path in list(session.filter_scan_cache.keys()):
        if cached_path not in active_paths or not os.path.exists(cached_path):
            session.filter_scan_cache.pop(cached_path, None)

    return combined


def _build_prognos_excel_sheets(report_df: pd.DataFrame, meta: Optional[dict] = None) -> Dict[str, pd.DataFrame]:
    sheets: Dict[str, pd.DataFrame] = {}
    if isinstance(meta, dict) and (meta.get("partial") == "yes" or meta.get("note")):
        lines: List[str] = []
        if meta.get("partial") == "yes":
            missing = meta.get("missing", "")
            lines.append("PARTIELL RAPPORT - mer data krävs för fullständig bild.")
            if missing:
                lines.append(f"Saknar underlag: {missing}.")
        if meta.get("note"):
            lines.append(str(meta["note"]))
        if lines:
            sheets["Info"] = pd.DataFrame({"Info": [" ".join(lines)]})

    if not isinstance(report_df, pd.DataFrame):
        output_df = pd.DataFrame()
    else:
        output_df = report_df.copy()
        col_name = "FIFO-baserad beräkning (antal pall)"
        if col_name in output_df.columns:
            try:
                output_df = output_df.sort_values(by=col_name, ascending=False).reset_index(drop=True)
            except Exception:
                pass

    sheets["Prognos vs Autoplock"] = output_df
    return sheets


async def _generate_prognos_result(session: SessionData) -> Dict[str, Any]:
    loop = asyncio.get_event_loop()
    prognos_path = session.files.get("prognos")
    campaign_path = session.files.get("campaign")
    automation_path = session.files.get("automation")
    buffer_path = session.files.get("buffer")

    if not prognos_path and not campaign_path:
        raise RuntimeError("Ladda upp Prognos eller Kampanjvolymer först.")

    prognos_df: Optional[pd.DataFrame] = None
    campaign_df: Optional[pd.DataFrame] = None
    saldo_raw: Optional[pd.DataFrame] = None
    buffer_raw: Optional[pd.DataFrame] = None

    if prognos_path and os.path.exists(prognos_path):
        prognos_df = await loop.run_in_executor(None, lambda p=prognos_path: read_prognos_xlsx(p))
    if campaign_path and os.path.exists(campaign_path):
        campaign_df = await loop.run_in_executor(None, lambda p=campaign_path: read_campaign_xlsx(p))
    if automation_path and os.path.exists(automation_path):
        saldo_raw = await loop.run_in_executor(None, lambda p=automation_path: read_csv_auto(p))
    if buffer_path and os.path.exists(buffer_path):
        buffer_raw = await loop.run_in_executor(None, lambda p=buffer_path: read_csv_auto(p))

    has_prognos = isinstance(prognos_df, pd.DataFrame) and not prognos_df.empty
    has_campaign = isinstance(campaign_df, pd.DataFrame) and not campaign_df.empty
    if not has_prognos and not has_campaign:
        raise RuntimeError("Kunde inte läsa Prognos eller Kampanjvolymer.")

    if has_prognos:
        combined_df = prognos_df.copy()
    else:
        combined_df = pd.DataFrame(
            columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"]
        )

    if has_campaign:
        camp_df = campaign_df.copy()
        if isinstance(saldo_raw, pd.DataFrame) and not saldo_raw.empty:
            saldo_df = saldo_raw.copy()
            art_col_sal = None
            robot_col_sal = None
            for c in saldo_df.columns:
                lc = str(c).strip().lower()
                if not art_col_sal and lc in ("artikel", "artikelnummer", "artnr", "art.nr", "sku", "article"):
                    art_col_sal = c
                if not robot_col_sal and lc == "robot":
                    robot_col_sal = c
            if art_col_sal and robot_col_sal:
                saldo_df = saldo_df[[art_col_sal, robot_col_sal]].copy()
                saldo_df.columns = ["Artikelnummer", "Robot"]
                saldo_df["Artikelnummer"] = saldo_df["Artikelnummer"].astype(str).str.strip()
                saldo_df["Robot"] = saldo_df["Robot"].astype(str).str.upper().str.strip()
                saldo_df = saldo_df.loc[saldo_df["Robot"] == "Y"]
                if not saldo_df.empty:
                    camp_df = camp_df.merge(saldo_df[["Artikelnummer"]], on="Artikelnummer", how="inner")
                else:
                    camp_df = camp_df.iloc[0:0]
            else:
                camp_df = camp_df.iloc[0:0]
        else:
            camp_df = camp_df.iloc[0:0]

        if not camp_df.empty:
            vol_by_art = camp_df.groupby("Artikelnummer")["Antal styck"].sum().to_dict()
            combined_df["Artikelnummer"] = combined_df["Artikelnummer"].astype(str).str.strip()
            combined_df["Antal styck"] = (
                pd.to_numeric(combined_df.get("Antal styck", 0), errors="coerce").fillna(0).astype(int)
            )
            existing_arts = set(combined_df["Artikelnummer"].astype(str))
            for art, vol in vol_by_art.items():
                if art in existing_arts:
                    mask = combined_df["Artikelnummer"] == art
                    combined_df.loc[mask, "Antal styck"] = (
                        combined_df.loc[mask, "Antal styck"].astype(int) + int(vol)
                    ).astype(int)
                else:
                    combined_df = pd.concat(
                        [
                            combined_df,
                            pd.DataFrame(
                                {
                                    "Artikelnummer": [art],
                                    "Beskrivning": [None],
                                    "Antal styck": [int(vol)],
                                    "Antal rader": [0],
                                    "Antal butiker": [0],
                                }
                            ),
                        ],
                        ignore_index=True,
                    )

    report_df, meta = await loop.run_in_executor(
        None,
        lambda: build_prognos_vs_autoplock_report(
            prognos_df=combined_df,
            saldo_norm_df=(saldo_raw if isinstance(saldo_raw, pd.DataFrame) else None),
            buffer_df=(buffer_raw if isinstance(buffer_raw, pd.DataFrame) else None),
            exclude_source_ids=None,
            allocated_df=None,
        ),
    )

    result_path = os.path.join(session.results_dir, "prognos.xlsx")
    sheets = _build_prognos_excel_sheets(report_df, meta)
    await loop.run_in_executor(None, lambda: save_df_to_excel(sheets, "prognos", result_path))
    session.results["prognos"] = result_path

    return {
        "ok": True,
        "result_key": "prognos",
        "rows": int(len(report_df)) if isinstance(report_df, pd.DataFrame) else 0,
        "partial": bool(isinstance(meta, dict) and meta.get("partial") == "yes"),
        "missing": str(meta.get("missing", "")) if isinstance(meta, dict) else "",
        "note": str(meta.get("note", "")) if isinstance(meta, dict) else "",
    }

# ---------------------------------------------------------------------------
# Pydantic modeller
# ---------------------------------------------------------------------------

class FilterBody(BaseModel):
    bolag: Optional[List[str]] = None
    ordertyp: Optional[List[str]] = None


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
    "wms_buffer": ["buffertpallet", "buffert", "buffer"],
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


def _detect_uploaded_file_type(path: str) -> Optional[str]:
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    base_name = os.path.basename(path).lower()
    wms_name_hints = {
        "v_ask_receive_log": "wms_receive",
        "v_ask_booking_putaway": "wms_booking",
        "v_ask_article_buffertpallet": "buffer",
        "v_ask_trans_log": "wms_trans",
        "v_ask_pick_log_full": "wms_pick",
        "v_ask_correct_log": "wms_correct",
    }
    for hint, file_type in wms_name_hints.items():
        if hint in base_name:
            return file_type

    generic_buffer_name_hints = (
        "buffertpall",
        "buffertpallet",
        "buffert_pall",
        "bufferpall",
        "bufferpallet",
        "buffer_pallet",
    )
    if any(hint in base_name for hint in generic_buffer_name_hints):
        return "buffer"

    if ext in ("xlsx", "xlsm", "xls"):
        try:
            df_campaign = read_campaign_xlsx(path)
            if isinstance(df_campaign, pd.DataFrame) and not df_campaign.empty and list(df_campaign.columns) == ["Artikelnummer", "Antal styck"]:
                return "campaign"
        except Exception:
            pass
        try:
            df_prognos = read_prognos_xlsx(path)
            if (
                isinstance(df_prognos, pd.DataFrame)
                and not df_prognos.empty
                and len(df_prognos.columns) >= 3
                and any(str(c).strip().lower() in ("antal styck", "quantity", "qty") for c in df_prognos.columns)
            ):
                return "prognos"
        except Exception:
            pass
        return None

    try:
        df = pd.read_csv(path, dtype=str, nrows=50, sep=None, engine="python", encoding="utf-8-sig")
        if df.shape[1] == 1:
            df = pd.read_csv(path, dtype=str, nrows=50, sep="\t", engine="python", encoding="utf-8-sig")
    except Exception:
        try:
            df = pd.read_csv(path, dtype=str, nrows=50, sep="\t", engine="python", encoding="utf-8-sig")
        except Exception:
            return None

    df = _clean_columns(df)
    cols = [_norm_col_key(c) for c in df.columns]

    def has_exact(*candidates: str) -> bool:
        normalized = {_norm_col_key(candidate) for candidate in candidates}
        return any(col in normalized for col in cols)

    def has_fragment(*candidates: str) -> bool:
        normalized = [_norm_col_key(candidate) for candidate in candidates]
        return any(candidate and candidate in col for col in cols for candidate in normalized)

    has_art = has_exact("artikel", "artikelnummer", "artnr", "art.nr", "sku", "article")
    has_qty = has_exact("beställt", "antal", "qty", "quantity", "bestalld", "order qty", "antal styck")
    has_ord = has_exact("ordernr", "order nr", "order number", "kund", "kundnr", "order id")
    has_rad = has_exact("radnr", "rad nr", "line id", "rad", "struktur", "radsnr")
    if has_art and has_qty and (has_ord or has_rad):
        return "orders"

    has_lagerplats = has_fragment("lagerplats") or has_exact("plats", "location", "bin")
    has_pallid = has_exact("pallid", "pall id", "id", "sscc", "etikett", "batch")
    has_status = has_exact("status")
    has_inkop = has_fragment("inköpsnr", "inkopsnr")
    has_mottaget = has_fragment("mottaget")
    has_pallnr = has_fragment("pall nr", "pallnr")
    has_till = has_exact("till") or has_fragment(" till", "till ")
    has_fran = has_fragment("från", "fran")
    has_plockat = has_fragment("plockat")
    has_anledning = has_fragment("anledning")

    if has_inkop and has_art and has_pallid and has_mottaget:
        return "wms_receive"
    if has_inkop and (has_pallnr or has_pallid) and not has_mottaget and not has_plockat:
        return "wms_booking"
    if has_lagerplats and has_pallid and has_inkop:
        return "buffer"
    if has_pallid and has_till and has_fran:
        return "wms_trans"
    if has_pallid and has_plockat and has_ord:
        return "wms_pick"
    if has_anledning and has_qty:
        return "wms_correct"

    if has_art and has_qty and has_lagerplats:
        return "buffer"

    buffer_marker_count = sum(1 for flag in (has_lagerplats, has_pallid, has_status, has_inkop, has_mottaget, has_pallnr) if flag)
    if has_art and (has_qty or has_pallid) and buffer_marker_count >= 2:
        return "buffer"

    has_pack = has_fragment("pack klass", "staplingsbar")
    if has_pack:
        has_plockpall = has_fragment("plockpall")
        has_dispatch_order = has_exact("ordernr", "order nr", "order number", "ordernummer")
        has_dispatch_ship = has_fragment("sändnings", "sandnings", "sändningsnr", "sandningsnr", "sändningsnr.", "sandningsnr.")
        if has_plockpall and has_dispatch_order and has_dispatch_ship:
            return "dispatch"
        return "item"

    has_ordernr = has_exact("ordernr", "order nr", "order number")
    has_orderdatum = has_fragment("orderdatum")
    has_sandning = has_fragment("sändningsnr", "sandningsnr", "sändningsnr.", "sandnr")
    has_ordertyp = has_fragment("ordertyp")
    if has_ordernr and has_orderdatum and has_sandning and has_ordertyp:
        return "overview"

    has_plockpall = has_fragment("plockpall")
    has_dispatch_order = has_exact("ordernr", "order nr", "order number", "ordernummer")
    has_dispatch_ship = has_fragment("sändnings", "sandnings", "sändningsnr", "sandningsnr", "sändningsnr.", "sandningsnr.")
    if has_plockpall and has_dispatch_order and has_dispatch_ship:
        return "dispatch"

    return None


def _resolve_upload_slot(requested_key: str, detected_type: Optional[str], page_id: str) -> str:
    if not detected_type:
        return requested_key
    if detected_type == "buffer":
        if requested_key in {"wms_buffer", "wms_buffert"} or page_id == "eftersok-page":
            return "wms_buffer"
        return "buffer"
    return detected_type


# ---------------------------------------------------------------------------
# Auth middleware
# ---------------------------------------------------------------------------

_AUTH_EXEMPT = {"/api/auth/login"}


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path
    # Tillåt inloggning och icke-API-anrop (statiska filer, login.html etc.)
    if path in _AUTH_EXEMPT or not path.startswith("/api/"):
        return await call_next(request)
    # Extrahera token
    auth_header = request.headers.get("authorization")
    token_param = request.query_params.get("token")
    token = extract_token(auth_header, token_param)
    if not token or not get_auth_session(token):
        return JSONResponse(status_code=401, content={"detail": "Ej inloggad"})
    request.state.auth_user = get_auth_session(token)
    return await call_next(request)


def _current_user(request: Request) -> dict:
    user = getattr(request.state, "auth_user", None)
    if not user:
        raise HTTPException(status_code=401, detail="Ej inloggad")
    return user


def _get_owned_session(request: Request, sid: str) -> SessionData:
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    user = _current_user(request)
    if user.get("role") != "admin" and session.owner_username != user.get("username"):
        raise HTTPException(status_code=403, detail="Du har inte åtkomst till denna session")
    return session


# ---------------------------------------------------------------------------
# Auth endpoints
# ---------------------------------------------------------------------------


class LoginRequest(BaseModel):
    identifier: str
    code: str = ""


@app.post("/api/auth/login")
def api_auth_login(body: LoginRequest):
    ident = body.identifier.strip()
    if not ident:
        raise HTTPException(400, "Ange e-post eller användarnamn")

    # Kolla om admin
    ident_lower = ident.lower()
    if ident_lower in (ADMIN_EMAIL.lower(), ADMIN_USERNAME.lower()):
        if not body.code:
            return {"needs_code": True}
        if body.code != ADMIN_CODE:
            raise HTTPException(401, "Fel kod")
        user_info = find_user(ident)
        token = create_auth_session(user_info)
        return {"token": token, "role": "admin", "lists": list(VALID_LISTS), "username": ADMIN_USERNAME}

    user_info = find_user(ident)
    if not user_info:
        return {"not_authorized": True}

    token = create_auth_session(user_info)
    return {
        "token": token,
        "role": user_info["role"],
        "lists": user_info["lists"],
        "username": user_info["username"],
    }


@app.get("/api/auth/me")
def api_auth_me(authorization: str = Header(None), token: str = Query(None)):
    t = extract_token(authorization, token)
    user = require_auth(t)
    return {"username": user["username"], "email": user.get("email", ""), "role": user["role"], "lists": user["lists"]}


@app.post("/api/auth/logout")
def api_auth_logout(authorization: str = Header(None)):
    t = extract_token(authorization)
    if t:
        delete_auth_session(t)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Admin endpoints – användarlista
# ---------------------------------------------------------------------------


@app.get("/api/admin/users")
def api_admin_get_users(authorization: str = Header(None)):
    require_admin(extract_token(authorization))
    return load_users()["lists"]


@app.post("/api/admin/users/{list_key}")
def api_admin_add_user(list_key: str, authorization: str = Header(None), username: str = Form(...), email: str = Form("")):
    require_admin(extract_token(authorization))
    if list_key not in VALID_LISTS:
        raise HTTPException(400, f"Ogiltig lista: {list_key}")
    data = load_users()
    users = data["lists"][list_key]["users"]
    # Kolla dubbletter
    uname_lower = username.strip().lower()
    for u in users:
        if u["username"].lower() == uname_lower:
            raise HTTPException(409, f"Användaren '{username}' finns redan i listan")
    users.append({"username": username.strip(), "email": email.strip()})
    save_users(data)
    return {"ok": True}


@app.delete("/api/admin/users/{list_key}/{username}")
def api_admin_delete_user(list_key: str, username: str, authorization: str = Header(None)):
    require_admin(extract_token(authorization))
    if list_key not in VALID_LISTS:
        raise HTTPException(400, f"Ogiltig lista: {list_key}")
    data = load_users()
    users = data["lists"][list_key]["users"]
    uname_lower = username.strip().lower()
    data["lists"][list_key]["users"] = [u for u in users if u["username"].lower() != uname_lower]
    save_users(data)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Session endpoints
# ---------------------------------------------------------------------------

@app.post("/api/session")
def api_create_session(request: Request):
    user = _current_user(request)
    s = create_session(user["username"])
    return {"session_id": s.session_id}


@app.delete("/api/session/{sid}")
def api_delete_session(request: Request, sid: str):
    _get_owned_session(request, sid)
    delete_session(sid)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Upload endpoints
# ---------------------------------------------------------------------------

@app.post("/api/upload/{sid}")
async def api_upload(request: Request, sid: str, file_key: str = Form(...), page_id: str = Form(""), file: UploadFile = Form(...)):
    session = _get_owned_session(request, sid)
    safe_name = re.sub(r"[^\w.\-]", "_", file.filename or "file")
    upload_tmp = os.path.join(session.uploads_dir, f"upload_{uuid.uuid4().hex}_{safe_name}")
    content = await file.read()
    with open(upload_tmp, "wb") as f:
        f.write(content)
    detected_type = _detect_uploaded_file_type(upload_tmp)
    actual_file_key = _resolve_upload_slot(file_key, detected_type, page_id)
    dest = os.path.join(session.uploads_dir, f"{actual_file_key}_{safe_name}")
    old_path = session.files.get(actual_file_key)
    if old_path and old_path != dest and os.path.exists(old_path):
        try:
            os.remove(old_path)
        except Exception:
            pass
    if os.path.exists(dest) and dest != upload_tmp:
        try:
            os.remove(dest)
        except Exception:
            pass
    shutil.move(upload_tmp, dest)
    session.files[actual_file_key] = dest
    _handle_session_file_change(session, actual_file_key, old_path=old_path)
    return {
        "ok": True,
        "file_key": actual_file_key,
        "requested_file_key": file_key,
        "filename": safe_name,
        "detected_type": detected_type,
    }


@app.delete("/api/upload/{sid}/{file_key}")
def api_remove_file(request: Request, sid: str, file_key: str):
    session = _get_owned_session(request, sid)
    path = session.files.pop(file_key, None)
    _handle_session_file_change(session, file_key, old_path=path)
    if path and os.path.exists(path):
        os.remove(path)
    return {"ok": True}


@app.get("/api/upload/{sid}")
def api_list_files(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    result: Dict[str, Optional[str]] = {}
    for k, v in list(session.files.items()):
        if v and os.path.exists(v):
            result[k] = os.path.basename(v)
        else:
            # Avoid stale "loaded" state in frontend when temp files were removed.
            _clear_filter_cache_entry(session, v)
            session.files.pop(k, None)
            result[k] = None
    return {"files": result}


# ---------------------------------------------------------------------------
# Filter endpoints
# ---------------------------------------------------------------------------

@app.get("/api/filters/{sid}")
def api_get_filter_options(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    return {
        "options": _collect_filter_options(session),
        "selected": session.active_filters,
    }


@app.post("/api/filters/{sid}")
def api_set_filters(request: Request, sid: str, body: FilterBody):
    session = _get_owned_session(request, sid)
    session.active_filters = {"bolag": body.bolag, "ordertyp": body.ordertyp}
    _invalidate_ordersaldo_cache(session)
    invalidate_allokering_cache(session)
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
        raise HTTPException(status_code=502, detail=f"Kunde inte nå CSV-agenten: {e}") from e


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
def api_ask_csv_fetch(request: Request, sid: str, body: AskCsvFetchBody):
    session = _get_owned_session(request, sid)

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
        raise HTTPException(status_code=502, detail=f"Kunde inte nå CSV-agenten: {e}") from e

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
    dest = os.path.join(session.uploads_dir, f"{body.target_file_key}_{safe_name}")
    old_path = session.files.get(body.target_file_key)

    with open(dest, "wb") as f:
        f.write(resp.content)

    if old_path and old_path != dest and os.path.exists(old_path):
        try:
            os.remove(old_path)
        except Exception:
            pass

    session.files[body.target_file_key] = dest
    _handle_session_file_change(session, body.target_file_key, old_path=old_path)
    return {
        "ok": True,
        "file_key": body.target_file_key,
        "filename": os.path.basename(dest),
        "bytes": len(resp.content),
    }


# ---------------------------------------------------------------------------
# Run status
# ---------------------------------------------------------------------------

@app.get("/api/run/status/{sid}")
def api_run_status(request: Request, sid: str):
    _get_owned_session(request, sid)
    try:
        return get_session_status(sid)
    except RuntimeError:
        raise HTTPException(status_code=404, detail="Session saknas")


@app.post("/api/result/prognos/{sid}")
async def api_build_prognos_result(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    try:
        return await _generate_prognos_result(session)
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/result/build/{sid}/{result_key}")
async def api_build_allokering_result(request: Request, sid: str, result_key: str):
    session = _get_owned_session(request, sid)
    try:
        return await _build_allokering_result(session, result_key)
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))


# ---------------------------------------------------------------------------
# Jobbkörningar
# ---------------------------------------------------------------------------

def _start_job(session: SessionData, coro):
    """Starta ett bakgrundsjobb."""
    if session.running:
        return False
    session.running = True
    asyncio.ensure_future(coro)
    return True


def _build_prepared_allokering_results(result: pd.DataFrame, near: pd.DataFrame) -> list[str]:
    prepared: list[str] = []
    if isinstance(result, pd.DataFrame) and not result.empty:
        prepared.extend(["allokerade", "pallplatser", "refill"])
    if isinstance(near, pd.DataFrame) and not near.empty:
        prepared.append("nearmiss")
    return prepared


async def _store_allokering_artifacts(
    session: SessionData,
    result: pd.DataFrame,
    near: pd.DataFrame,
    buffer_raw: pd.DataFrame,
    saldo_norm: Optional[pd.DataFrame],
    not_putaway_norm: Optional[pd.DataFrame],
) -> dict:
    loop = asyncio.get_event_loop()
    prepared_results = _build_prepared_allokering_results(result, near)
    meta = {
        "allokerade_rows": int(len(result)) if isinstance(result, pd.DataFrame) else 0,
        "nearmiss_rows": int(len(near)) if isinstance(near, pd.DataFrame) else 0,
    }
    await loop.run_in_executor(None, lambda: store_allokering_dataframe(session, "result", result))
    await loop.run_in_executor(None, lambda: store_allokering_dataframe(session, "near", near))
    await loop.run_in_executor(None, lambda: store_allokering_dataframe(session, "buffer_raw", buffer_raw))
    await loop.run_in_executor(None, lambda: store_allokering_dataframe(session, "saldo_norm", saldo_norm))
    await loop.run_in_executor(None, lambda: store_allokering_dataframe(session, "not_putaway_norm", not_putaway_norm))
    await loop.run_in_executor(None, lambda: save_allokering_manifest(session, prepared_results, meta))
    return {"prepared_results": prepared_results, "meta": meta}


async def _refresh_ordersaldo_after_allokering(session: SessionData, orders_path: Optional[str]) -> None:
    if not orders_path:
        return
    loop = asyncio.get_event_loop()
    list1, list2 = await loop.run_in_executor(
        None, lambda: refresh_ordersaldo(orders_path, session.active_filters),
    )
    session.ordersaldo_list1 = list1
    session.ordersaldo_list2 = list2


async def _build_allokering_result(session: SessionData, result_key: str) -> dict:
    if result_key not in ALLOKERING_RESULT_KEYS:
        raise RuntimeError(f"Ogiltig resultatnyckel: {result_key}")

    existing_path = session.results.get(result_key)
    if existing_path and os.path.exists(existing_path):
        return {"ok": True, "result_key": result_key, "available": True, "cached": True}

    manifest = load_allokering_manifest(session)
    prepared_results = list(manifest.get("prepared_results") or [])
    if result_key not in prepared_results:
        raise RuntimeError("Resultatet är inte förberett ännu. Kör allokering först.")

    loop = asyncio.get_event_loop()
    result_df = await loop.run_in_executor(None, lambda: load_allokering_dataframe(session, "result"))
    if result_df is None or result_df.empty:
        raise RuntimeError("Allokeringsresultatet saknas. Kör allokering igen.")

    rows = 0
    available = False

    if result_key == "allokerade":
        rows = int(len(result_df))
        out_path = os.path.join(session.results_dir, "allokerade.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(result_df, "allokerade", out_path))
        session.results["allokerade"] = out_path
        available = True
    elif result_key == "nearmiss":
        near_df = await loop.run_in_executor(None, lambda: load_allokering_dataframe(session, "near"))
        if isinstance(near_df, pd.DataFrame) and not near_df.empty:
            rows = int(len(near_df))
            out_path = os.path.join(session.results_dir, "nearmiss.xlsx")
            await loop.run_in_executor(None, lambda: save_df_to_excel(near_df, "nearmiss", out_path))
            session.results["nearmiss"] = out_path
            available = True
    elif result_key == "pallplatser":
        pallet_spaces_df = await loop.run_in_executor(None, lambda: compute_pallet_spaces(result_df.copy()))
        if isinstance(pallet_spaces_df, pd.DataFrame) and not pallet_spaces_df.empty:
            rows = int(len(pallet_spaces_df))
            out_path = os.path.join(session.results_dir, "pallplatser.xlsx")
            await loop.run_in_executor(None, lambda: save_df_to_excel(pallet_spaces_df, "pallplatser", out_path))
            session.results["pallplatser"] = out_path
            available = True
    elif result_key == "refill":
        buffer_raw = await loop.run_in_executor(None, lambda: load_allokering_dataframe(session, "buffer_raw"))
        saldo_norm = await loop.run_in_executor(None, lambda: load_allokering_dataframe(session, "saldo_norm"))
        not_putaway_norm = await loop.run_in_executor(None, lambda: load_allokering_dataframe(session, "not_putaway_norm"))
        if buffer_raw is None or buffer_raw.empty:
            raise RuntimeError("Buffertunderlaget saknas. Kör allokering igen.")
        hp_df, as_df = await loop.run_in_executor(
            None,
            lambda: calculate_refill(
                result_df.copy(),
                buffer_raw.copy(),
                saldo_df=(saldo_norm.copy() if isinstance(saldo_norm, pd.DataFrame) else None),
                not_putaway_df=(not_putaway_norm.copy() if isinstance(not_putaway_norm, pd.DataFrame) else None),
            ),
        )
        has_refill = (isinstance(hp_df, pd.DataFrame) and not hp_df.empty) or (
            isinstance(as_df, pd.DataFrame) and not as_df.empty
        )
        if has_refill:
            rows = int(len(hp_df)) + int(len(as_df))
            refill_sheets = {}
            if isinstance(hp_df, pd.DataFrame) and not hp_df.empty:
                refill_sheets["Påfyllning HP"] = hp_df
            if isinstance(as_df, pd.DataFrame) and not as_df.empty:
                refill_sheets["Påfyllning AutoStore"] = as_df
            out_path = os.path.join(session.results_dir, "refill.xlsx")
            await loop.run_in_executor(None, lambda: save_df_to_excel(refill_sheets, "refill", out_path))
            session.results["refill"] = out_path
            available = True

    if not available:
        prepared_results = [key for key in prepared_results if key != result_key]
        await loop.run_in_executor(
            None,
            lambda: save_allokering_manifest(session, prepared_results, manifest.get("meta") or {}),
        )
        return {"ok": True, "result_key": result_key, "available": False, "rows": 0}

    session.log_queue.put_nowait(f"__RESULT:{result_key}__")
    return {"ok": True, "result_key": result_key, "available": True, "rows": rows}


async def _job_allokering(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        orders_path = session.files.get("orders")
        buffer_path = session.files.get("buffer")
        automation_path = session.files.get("automation")
        item_path = session.files.get("item")
        booking_path = session.files.get("wms_booking")

        if not orders_path or not buffer_path:
            log("FEL: orders och buffer m\u00e5ste vara uppladdade.")
            log("__ERROR__")
            return

        # --- Fas 1: Läs alla filer parallellt ---
        log("L\u00e4ser in filer...")
        read_tasks: Dict[str, Any] = {}
        read_tasks["orders"] = loop.run_in_executor(None, lambda p=orders_path: read_csv_auto(p))
        read_tasks["buffer"] = loop.run_in_executor(None, lambda p=buffer_path: read_csv_auto(p))
        if automation_path and os.path.exists(automation_path):
            read_tasks["automation"] = loop.run_in_executor(None, lambda p=automation_path: read_csv_auto(p))
        if item_path and os.path.exists(item_path):
            read_tasks["item"] = loop.run_in_executor(None, lambda p=item_path: read_csv_auto(p))
        if booking_path and os.path.exists(booking_path):
            read_tasks["wms_booking"] = loop.run_in_executor(None, lambda p=booking_path: read_csv_auto(p))

        keys = list(read_tasks.keys())
        read_results = await asyncio.gather(*read_tasks.values())
        raw_data = dict(zip(keys, read_results))

        orders_raw = _clean_columns(raw_data["orders"])
        buffer_raw = _clean_columns(raw_data["buffer"])

        saldo_norm = None
        saldo_raw = None
        if "automation" in raw_data:
            auto_raw = _clean_columns(raw_data["automation"].copy())
            saldo_raw = auto_raw.copy()
            saldo_norm = await loop.run_in_executor(None, lambda: normalize_saldo(auto_raw))

        item_norm = None
        if "item" in raw_data:
            try:
                item_norm = await loop.run_in_executor(None, lambda: normalize_items(raw_data["item"]))
            except Exception as ie:
                log(f"Varning: Kunde inte l\u00e4sa item-fil: {ie}")

        not_putaway_norm = None
        if "wms_booking" in raw_data:
            try:
                booking_raw = apply_value_filters(_clean_columns(raw_data["wms_booking"].copy()), session.active_filters)
                not_putaway_norm = await loop.run_in_executor(None, lambda: normalize_not_putaway(booking_raw))
            except Exception as npe:
                log(f"Varning: Kunde inte l\u00e4sa ej inlagrade artiklar: {npe}")

        # Applicera filter
        orders_raw = apply_value_filters(orders_raw, session.active_filters)
        buffer_raw = apply_value_filters(buffer_raw, session.active_filters)
        if saldo_raw is not None:
            saldo_raw = apply_value_filters(saldo_raw, session.active_filters)
            saldo_norm = await loop.run_in_executor(None, lambda: normalize_saldo(saldo_raw))

        # --- Fas 2: Kör allokering ---
        log("K\u00f6r allokering (Helpall \u2192 AutoStore \u2192 Huvudplock, FIFO)...")
        result, near = await loop.run_in_executor(None, lambda: allocate(orders_raw, buffer_raw, log=log))

        result = await loop.run_in_executor(None, lambda: _reclassify_skrymmande(result, saldo_norm))

        # Slå ihop item-fil
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
                log(f"Kunde inte sl\u00e5 ihop item-fil: {e}")

        if "Ej Staplingsbar" not in result.columns:
            result["Ej Staplingsbar"] = ""

        # --- Fas 3: Summeringar + Klar ---
        # Summering per zon
        try:
            zon_col = "Zon (ber\u00e4knad)"
            qty_col = find_col(result, ORDER_SCHEMA["qty"], required=True)
            summary = result.groupby(zon_col)[qty_col].apply(
                lambda s: pd.to_numeric(s, errors="coerce").sum()).reset_index(name="Totalt antal")
            log("\nSummering per zon:")
            for _, r in summary.iterrows():
                log(f"  Zon {r[zon_col]}: {r['Totalt antal']:.0f}")
        except Exception:
            pass

        # Summering per Källtyp
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
        except Exception as _e_kt:
            log(f"Summering per K\u00e4lltyp kunde inte ber\u00e4knas: {_e_kt}")

        artifact_info = await _store_allokering_artifacts(
            session, result, near, buffer_raw, saldo_norm, not_putaway_norm
        )
        for prepared_key in artifact_info.get("prepared_results") or []:
            log(f"__PREPARED:{prepared_key}__")
        log("Allokeringen \u00e4r klar.")
        try:
            await _refresh_ordersaldo_after_allokering(session, orders_path)
            list1 = session.ordersaldo_list1
            list2 = session.ordersaldo_list2
            if list1 or list2:
                log(f"Ordersaldo: {len(list1)} kompletta ordrar, {len(list2)} artiklar med p\u00e5fyllningsbehov.")
        except Exception as e:
            log(f"Varning: Ordersaldo-ber\u00e4kning misslyckades: {e}")
        log("__DONE__")
        session.running = False

    except Exception as e:
        import traceback
        log(f"FEL: {e}")
        log(traceback.format_exc())
        log("__ERROR__")
        session.running = False


async def _job_hib_koppling(session: SessionData):
    loop = asyncio.get_event_loop()

    def log(msg: str):
        session.log_queue.put_nowait(msg)

    try:
        orders_path = session.files.get("orders")
        overview_path = session.files.get("overview")

        if not orders_path or not overview_path:
            log("FEL: Välj både beställningslinjer och orderöversikt.")
            log("__ERROR__")
            return

        log("Läser in filer för HIB-koppling...")
        details_df = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
        overview_df = await loop.run_in_executor(None, lambda: read_csv_auto(overview_path))

        details_df = apply_value_filters(details_df, session.active_filters)
        overview_df = apply_value_filters(overview_df, session.active_filters)

        log("Beräknar HIB-koppling...")
        changes_df = await loop.run_in_executor(None, lambda: compute_hib_koppling(details_df, overview_df))
        missed_df = await loop.run_in_executor(None, lambda: compute_missed_departures(details_df, overview_df))

        has_changes = isinstance(changes_df, pd.DataFrame) and not changes_df.empty
        has_missed = isinstance(missed_df, pd.DataFrame) and not missed_df.empty

        if not has_changes and not has_missed:
            log("Inga HIB-ordrar behöver ändras eller har missat sin avgång.")
            log("__DONE__")
            return

        instr_lines = [
            "Ändras i följande ordning",
            "1. Ordernummer",
            "2. Sändningsnummer",
            "3. Zon F på orderlinjerna",
            "4. Samma multi på alla HIB:ar till samma butik",
            "5. Generera",
            "6. Frisläpp",
        ]
        instructions_df = pd.DataFrame({"Instruktioner": instr_lines})

        sheets: Dict[str, pd.DataFrame] = {}
        if has_changes:
            sheets["Ändringar"] = changes_df
            log(f"HIB-koppling: {len(changes_df)} ordrar att ändra.")
            for _, r in changes_df.iterrows():
                ordnr = str(r.get("ordernummer", "")).strip()
                kundnamn = str(r.get("kundnamn", "")).strip()
                fields = []
                if str(r.get("sändningsnummer", "")).strip():
                    fields.append(f"Sändningsnr -> {str(r['sändningsnummer']).strip()}")
                if str(r.get("Orderdatum", "")).strip():
                    fields.append(f"Orderdatum -> {str(r['Orderdatum']).strip()}")
                if str(r.get("Zon", "")).strip():
                    fields.append(f"Zon -> {str(r['Zon']).strip()}")
                if str(r.get("Multi", "")).strip():
                    fields.append(f"Multi -> {str(r['Multi']).strip()}")
                if fields:
                    name_part = f" ({kundnamn})" if kundnamn else ""
                    log(f"Order {ordnr}{name_part}: {', '.join(fields)}")

        if has_missed:
            sheets["Missade avgångar"] = missed_df
            log(f"Missade avgångar: {len(missed_df)} st.")
            for _, r in missed_df.iterrows():
                ordnr = str(r.get("ordernummer", "")).strip()
                kundnamn = str(r.get("kundnamn", "")).strip()
                name_part = f" ({kundnamn})" if kundnamn else ""
                log(f"Order {ordnr}{name_part}: MISSAT SIN AVGÅNG")

        sheets["Instruktion"] = instructions_df

        hib_path = os.path.join(session.results_dir, "hib_koppling.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(sheets, "hib_koppling", hib_path))
        session.results["hib-koppling"] = hib_path
        log("__RESULT:hib-koppling__")
        log("HIB-kopplingen är beräknad.")
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
            log("FEL: Välj orderöversikten först.")
            log("__ERROR__")
            return

        log("Läser in orderöversikt...")
        df = await loop.run_in_executor(None, lambda: read_csv_auto(overview_path))
        df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
        df = apply_value_filters(df, session.active_filters)

        if df.empty:
            log("Inga rader kvar efter filter i orderöversikten.")
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
            log("FEL: Kunde inte identifiera sändningsnummer-kolumnen.")
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

        date_col = _find_col_by_keywords(
            df,
            ["orderdatum", "order datum", "order date", "orderdate"],
        )

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
            log("Orderöversikten innehåller inga sändningsnummer.")
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
                        "Avvikelsetyp": "Sändningsnr med flera kunder/transportörer",
                        "Sändningsnr": ship,
                        "Unika kunder": len(customers),
                        "Kunder": ", ".join(customers),
                        "Unika transportörer": len(carriers),
                        "Transportörer": ", ".join(carriers),
                        "Antal orderrader": int(len(group)),
                    }
                    if orders_str:
                        row["Ordernr (kundnamn)"] = orders_str
                    shipment_diff_rows.append(row)
            except Exception:
                continue

        # Kontroll: sändningsnr med olika orderdatum
        date_diff_rows: List[Dict[str, Any]] = []
        if date_col:
            for ship, group in df.groupby(ship_col):
                try:
                    dates = sorted(set(group[date_col].dropna().astype(str).str.strip()))
                    dates = [d for d in dates if d]
                    if len(dates) > 1:
                        order_vals: List[str] = []
                        if order_col:
                            for _, r in group.drop_duplicates(subset=[order_col]).iterrows():
                                o = str(r[order_col]).strip()
                                d = str(r[date_col]).strip()
                                nm = order_to_customer.get(o, "")
                                order_vals.append(f"{o} ({nm}, {d})" if nm else f"{o} ({d})")
                        date_diff_rows.append({
                            "Avvikelsetyp": "Sändningsnr med olika orderdatum",
                            "Sändningsnr": ship,
                            "Antal unika datum": len(dates),
                            "Orderdatum": ", ".join(dates),
                            "Antal orderrader": int(len(group)),
                            "Ordernr (kundnamn, datum)": ", ".join(order_vals) if order_vals else "",
                        })
                except Exception:
                    continue

        result_df = pd.DataFrame(shipment_diff_rows) if shipment_diff_rows else pd.DataFrame()
        date_diff_df = pd.DataFrame(date_diff_rows) if date_diff_rows else pd.DataFrame()

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
                        "Sändningsnr": ", ".join(hib_ships),
                        "Ordertyp": "HIB",
                        "Status": int(max_status),
                        "Anmärkning": "HIB-order med status > 31 saknar matchande butikssändning",
                    }
                    if kundnamn:
                        row2["Kundnamn"] = kundnamn
                    hib_rows.append(row2)
            except Exception as e:
                log(f"HIB-kontrollen misslyckades delvis: {e}")

        hib_check_df = pd.DataFrame(hib_rows) if hib_rows else pd.DataFrame()

        has_any = not result_df.empty or not date_diff_df.empty or not hib_check_df.empty
        if not has_any:
            msg = "Inga avvikelser hittades i orderöversikten."
            if missing_hib_cols:
                msg += " HIB-kontrollen kunde inte köras fullt ut (saknar: " + ", ".join(missing_hib_cols) + ")."
            log(msg)
            log("__DONE__")
            return

        # Logga resultat
        if not result_df.empty:
            log(f"Orderöversikt: {len(result_df)} sändningsnummer med flera kunder/transportörer.")
        if not date_diff_df.empty:
            log(f"Orderöversikt: {len(date_diff_df)} sändningsnummer med olika orderdatum.")
        if not hib_check_df.empty:
            log(f"HIB-ordrar med status > 31 utan matchande butikssändning: {len(hib_check_df)} st.")
        if missing_hib_cols:
            log("HIB-kontrollen kunde inte köras fullt ut (saknar: " + ", ".join(missing_hib_cols) + ").")

        # Bygg Excel
        sheets: Dict[str, pd.DataFrame] = {}
        combined_parts = []
        if not result_df.empty:
            s_df = result_df.copy()
            if "Avvikelsetyp" not in s_df.columns:
                s_df.insert(0, "Avvikelsetyp", "Sändningsnr med flera kunder/transportörer")
            sheets["Sändningskontroll"] = s_df
            combined_parts.append(s_df)
        if not date_diff_df.empty:
            dd_df = date_diff_df.copy()
            if "Avvikelsetyp" not in dd_df.columns:
                dd_df.insert(0, "Avvikelsetyp", "Sändningsnr med olika orderdatum")
            sheets["Datumkontroll"] = dd_df
            combined_parts.append(dd_df)
        if not hib_check_df.empty:
            h_df = hib_check_df.copy()
            if "Avvikelsetyp" not in h_df.columns:
                h_df.insert(0, "Avvikelsetyp", "HIB över status 31 utan butikssändning")
            sheets["HIB utan butikssändning"] = h_df
            combined_parts.append(h_df)
        if combined_parts:
            combined = pd.concat(combined_parts, ignore_index=True, sort=False)
            sheets = {"Orderkontroll": combined, **sheets}

        ok_path = os.path.join(session.results_dir, "orderkontroll.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(sheets, "orderkontroll", ok_path))
        session.results["orderkontroll"] = ok_path
        log("__RESULT:orderkontroll__")
        log("Orderkontrollen är klar.")
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
            log("FEL: Välj både orderöversikt och dispatchpallar.")
            log("__ERROR__")
            return

        log("Läser in filer för dispatchkontroll...")
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
            log("FEL: Kunde inte identifiera order- eller sändningskolumnen i orderöversikten.")
            log("__ERROR__")
            return

        dp_order_col = _find_col_by_keywords(dp_df, order_kws)
        dp_ship_col = _find_col_by_keywords(dp_df, ship_kws)
        plock_col = _find_col_by_keywords(dp_df, plock_kws)
        if not dp_order_col or not dp_ship_col or not plock_col:
            log("FEL: Kunde inte identifiera order-, sändnings- eller plockpallskolumnen i dispatchfilen.")
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

        # Bygg order -> sändningsnr-mapping från orderöversikten
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
                        "Översikt sändningsnr": expected,
                        "Dispatch sändningsnr": dp_ship,
                        "Plockpallsnr": str(row[plock_col]).strip(),
                        "kundnamn": order_to_customer.get(ordnr, ""),
                    }
                    diff_rows.append(diff_row)
            except Exception:
                continue

        if not diff_rows:
            log("Alla sändningsnummer stämmer överens.")
            log("__DONE__")
            return

        diff_df = pd.DataFrame(diff_rows)
        log(f"Dispatchkontrollen hittade {len(diff_df)} avvikelser.")
        for _, row in diff_df.iterrows():
            name_part = f" ({row['kundnamn']})" if str(row.get("kundnamn", "")).strip() else ""
            log(f"Order {row['Ordernr']}{name_part}: sändningsnr {row['Översikt sändningsnr']} i översikten men {row['Dispatch sändningsnr']} i dispatch (plockpall {row['Plockpallsnr']})")

        dk_path = os.path.join(session.results_dir, "dispatchkontroll.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel({"Dispatchkontroll": diff_df}, "dispatchkontroll", dk_path))
        session.results["dispatchkontroll"] = dk_path
        log("__RESULT:dispatchkontroll__")
        log("Dispatchkontrollen är klar.")
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
            log("FEL: Ange både inköpsnummer och artikelnummer.")
            log("__ERROR__")
            return

        wms_slot_labels = {
            "wms_receive": "Varumottagningslogg",
            "wms_buffer": "Buffertpall",
            "wms_booking": "Ej inlagrade artiklar",
            "wms_trans": "Translogg",
            "wms_pick": "Plocklogg full",
            "wms_correct": "Saldojusteringar & inventeringsavvikelser",
        }

        # Samla alla tillgängliga WMS-filer
        wms_files: Dict[str, str] = {}
        for key in ["wms_receive", "wms_buffer", "wms_booking", "wms_trans", "wms_pick", "wms_correct"]:
            path = session.files.get(key)
            if not path and key == "wms_buffer":
                path = session.files.get("wms_buffert")
            if path and os.path.exists(path):
                wms_files[key] = path

        log(f"Eftersök: inköpsnummer={purchase}, artikelnummer={article}")
        if wms_files:
            readable_files = [wms_slot_labels.get(key, key) for key in wms_files.keys()]
            log(f"Tillgängliga WMS-filer: {', '.join(readable_files)}")
        else:
            log("Inga WMS-filer uppladdade för Eftersök.")

        results_sheets: Dict[str, pd.DataFrame] = {}

        # Sök igenom varje WMS-fil
        for key, path in wms_files.items():
            try:
                df = await loop.run_in_executor(None, lambda p=path: read_csv_auto(p))
                # Sök efter rader som matchar purchase eller article
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
                sheet_name = wms_slot_labels.get(key, key)
                if not hits.empty:
                    results_sheets[sheet_name] = hits
                    log(f"  {sheet_name}: {len(hits)} träff(ar)")
                else:
                    log(f"  {sheet_name}: inga träffar")
            except Exception as e:
                log(f"  {wms_slot_labels.get(key, key)}: fel vid läsning - {e}")

        if not results_sheets:
            log("Inga träffar hittades i tillgängliga WMS-filer.")
            log("__DONE__")
            return

        eftersok_path = os.path.join(session.results_dir, "eftersok.xlsx")
        await loop.run_in_executor(None, lambda: save_df_to_excel(results_sheets, "eftersok", eftersok_path))
        session.results["eftersok"] = eftersok_path
        log("__RESULT:eftersok__")
        log("Eftersöket är klart.")
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
async def api_run_allokering(request: Request, sid: str):
    _get_owned_session(request, sid)
    try:
        return JSONResponse(status_code=202, content=enqueue_job(sid, "allokering"))
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


@app.post("/api/run/hib-koppling/{sid}")
async def api_run_hib(request: Request, sid: str):
    _get_owned_session(request, sid)
    try:
        return JSONResponse(status_code=202, content=enqueue_job(sid, "hib-koppling"))
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


@app.post("/api/run/orderkontroll/{sid}")
async def api_run_orderkontroll(request: Request, sid: str):
    _get_owned_session(request, sid)
    try:
        return JSONResponse(status_code=202, content=enqueue_job(sid, "orderkontroll"))
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


@app.post("/api/run/dispatchkontroll/{sid}")
async def api_run_dispatchkontroll(request: Request, sid: str):
    _get_owned_session(request, sid)
    try:
        return JSONResponse(status_code=202, content=enqueue_job(sid, "dispatchkontroll"))
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


@app.post("/api/run/eftersok/{sid}")
async def api_run_eftersok(request: Request, sid: str, body: EftersokBody):
    _get_owned_session(request, sid)
    try:
        return JSONResponse(
            status_code=202,
            content=enqueue_job(
                sid,
                "eftersok",
                {"purchase": body.purchase, "article": body.article},
            ),
        )
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


# ---------------------------------------------------------------------------
# Ordersaldo endpoints
# ---------------------------------------------------------------------------

@app.post("/api/ordersaldo/refresh/{sid}")
async def api_ordersaldo_refresh(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    orders_path = session.files.get("orders")
    if not orders_path or not os.path.exists(orders_path):
        raise HTTPException(status_code=400, detail="Beställningslinjer-filen saknas")
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
def api_ordersaldo_list1(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    return {"values": session.ordersaldo_list1}


@app.get("/api/ordersaldo/list2/{sid}")
def api_ordersaldo_list2(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    return {"values": session.ordersaldo_list2}


# ---------------------------------------------------------------------------
# Debug
# ---------------------------------------------------------------------------

@app.get("/api/debug/columns/{sid}/{file_key}")
def api_debug_columns(request: Request, sid: str, file_key: str):
    session = _get_owned_session(request, sid)
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
async def api_reset_results(request: Request, sid: str):
    session = _get_owned_session(request, sid)
    invalidate_allokering_cache(session)
    for result_key in list(session.results.keys()):
        _invalidate_result_cache(session, result_key)
    session.ordersaldo_list1 = []
    session.ordersaldo_list2 = []
    session.filter_scan_cache.clear()
    clear_session_logs(sid)
    return {"ok": True}


@app.post("/api/session/reset-all/{sid}")
async def api_reset_all(request: Request, sid: str):
    """Rensa alla uppladdade filer OCH alla resultat för sessionen."""
    session = _get_owned_session(request, sid)
    # Rensa uppladdade filer
    for path in list(session.files.values()):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    session.files.clear()
    invalidate_allokering_cache(session)
    for result_key in list(session.results.keys()):
        _invalidate_result_cache(session, result_key)
    session.ordersaldo_list1 = []
    session.ordersaldo_list2 = []
    session.filter_scan_cache.clear()
    clear_session_logs(sid)
    return {"ok": True}


# ---------------------------------------------------------------------------
# SSE-logg
# ---------------------------------------------------------------------------

@app.get("/api/log/stream/{sid}")
async def api_log_stream(request: Request, sid: str):
    _get_owned_session(request, sid)
    last_event_id = 0
    try:
        header_value = request.headers.get("last-event-id", "") or request.headers.get("Last-Event-ID", "")
        if header_value:
            last_event_id = int(header_value)
        elif request.query_params.get("last_event_id"):
            last_event_id = int(request.query_params.get("last_event_id", "0"))
    except Exception:
        last_event_id = 0

    async def event_generator():
        nonlocal last_event_id
        while True:
            try:
                rows = await asyncio.to_thread(fetch_session_logs_since, sid, last_event_id)
                if rows:
                    for row in rows:
                        last_event_id = int(row["id"])
                        safe = str(row["message"]).replace("\n", "\\n").replace("\r", "")
                        yield f"id: {last_event_id}\ndata: {safe}\n\n"
                else:
                    await asyncio.sleep(1.0)
                    yield ": keepalive\n\n"
            except asyncio.CancelledError:
                break
            except ValueError:
                yield ": keepalive\n\n"
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

@app.get("/api/download/{sid}/{result_key}")
def api_download(request: Request, sid: str, result_key: str):
    session = _get_owned_session(request, sid)
    path = session.results.get(result_key)
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail=f"Resultat '{result_key}' ej tillgängligt")
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
        raise HTTPException(status_code=400, detail="Inga värden")

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
# Statiska filer - montera SIST
# ---------------------------------------------------------------------------

_frontend_dir = Path(__file__).parent.parent / "frontend"
if _frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(_frontend_dir), html=True), name="frontend")
