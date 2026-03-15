#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py – FastAPI-webbapp för Allokering.
Serverar frontend statiskt och exponerar REST-API + SSE.
"""

from __future__ import annotations

import asyncio
import os
import re
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import BackgroundTasks, FastAPI, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import sys
sys.path.insert(0, str(Path(__file__).parent))

from session_store import SessionData, create_session, delete_session, get_session
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
    for k, v in session.files.items():
        result[k] = os.path.basename(v) if v else None
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
    for key in ("orders", "automation", "overview"):
        path = session.files.get(key)
        if not path or not os.path.exists(path):
            continue
        try:
            df = read_csv_auto(path)
            vals = scan_filter_values(df)
            for fk, fvals in vals.items():
                existing = set(combined.get(fk, []))
                existing.update(fvals)
                combined[fk] = sorted(existing)
        except Exception:
            pass
    return combined


@app.post("/api/filters/{sid}")
def api_set_filters(sid: str, body: FilterBody):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    session.active_filters = {"bolag": body.bolag, "ordertyp": body.ordertyp}
    return {"ok": True}


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
# Job-körningar
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
            log("FEL: orders och buffer måste vara uppladdade.")
            log("__ERROR__")
            return

        log("Läser in filer...")
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
                log(f"Varning: Kunde inte läsa item-fil: {ie}")

        orders_raw = _clean_columns(orders_raw)
        buffer_raw = _clean_columns(buffer_raw)

        # Applicera filter
        orders_raw = apply_value_filters(orders_raw, session.active_filters)
        buffer_raw = apply_value_filters(buffer_raw, session.active_filters)
        if saldo_raw is not None:
            saldo_raw = apply_value_filters(saldo_raw, session.active_filters)
            saldo_norm = await loop.run_in_executor(None, lambda: normalize_saldo(saldo_raw))

        log("Kör allokering (Helpall → AutoStore → Huvudplock, FIFO)...")
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
                log(f"Kunde inte slå ihop item-fil: {e}")

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
            log(f"Pallplatser kunde inte beräknas: {e}")

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
                    refill_sheets["Påfyllning HP"] = hp_df
                if as_df is not None and not as_df.empty:
                    refill_sheets["Påfyllning AutoStore"] = as_df
                refill_path = os.path.join(session.temp_dir, "refill.xlsx")
                await loop.run_in_executor(None, lambda: save_df_to_excel(refill_sheets, "refill", refill_path))
                session.results["refill"] = refill_path
                log("__RESULT:refill__")
                log(f"Auto-refill klar: HP {len(hp_df)} rader, AUTOSTORE {len(as_df)} rader.")
        except Exception as e:
            log(f"Refill misslyckades: {e}")

        # Summering per zon
        try:
            zon_col = "Zon (beräknad)"
            qty_col = find_col(result, ORDER_SCHEMA["qty"], required=True)
            summary = result.groupby(zon_col)[qty_col].apply(
                lambda s: pd.to_numeric(s, errors="coerce").sum()).reset_index(name="Totalt antal")
            log("\nSummering per zon:")
            for _, r in summary.iterrows():
                log(f"  Zon {r[zon_col]}: {r['Totalt antal']:.0f}")
        except Exception:
            pass

        # Beräkna ordersaldo-listor (kompletta ordrar / påfyllningsbehov)
        try:
            if orders_path:
                list1, list2 = await loop.run_in_executor(
                    None,
                    lambda: refresh_ordersaldo(orders_path, session.active_filters),
                )
                session.ordersaldo_list1 = list1
                session.ordersaldo_list2 = list2
                if list1 or list2:
                    log(f"Ordersaldo: {len(list1)} kompletta ordrar, {len(list2)} artiklar med påfyllningsbehov.")
        except Exception as e:
            log(f"Varning: Ordersaldo-beräkning misslyckades: {e}")

        log("Allokeringen är klar.")
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
            "4. Samma multi på alla Hibar till samma butik",
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
                    fields.append(f"Sändningsnr → {str(r['sändningsnummer']).strip()}")
                if str(r.get("Orderdatum", "")).strip():
                    fields.append(f"Orderdatum → {str(r['Orderdatum']).strip()}")
                if str(r.get("Zon", "")).strip():
                    fields.append(f"Zon → {str(r['Zon']).strip()}")
                if str(r.get("Multi", "")).strip():
                    fields.append(f"Multi → {str(r['Multi']).strip()}")
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

        hib_path = os.path.join(session.temp_dir, "hib_koppling.xlsx")
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

        # Hitta kolumner
        ship_col = None
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if "sändning" in cl or "sandning" in cl or "sändningsnummer" in cl:
                ship_col = c
                break

        if not ship_col:
            log("FEL: Kunde inte identifiera sändningsnummer-kolumnen.")
            log("__ERROR__")
            return

        cust_col = None
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if "kundnr" in cl or "kundnummer" in cl:
                cust_col = c
                break
        if not cust_col:
            for c in df.columns:
                if "kund" in c.lower():
                    cust_col = c
                    break

        if not cust_col:
            log("FEL: Kunde inte identifiera kund-kolumnen.")
            log("__ERROR__")
            return

        trans_col = None
        for c in df.columns:
            cl = c.lower()
            if "transportör" in cl or "transportor" in cl:
                trans_col = c
                break
        if not trans_col:
            trans_col = "__transport_dummy__"
            df[trans_col] = ""

        order_col = None
        order_keywords = ["ordernr", "order nr", "ordernummer", "order number"]
        for c in df.columns:
            for kw in order_keywords:
                if kw.replace(" ", "") == c.lower().replace(" ", ""):
                    order_col = c
                    break
            if order_col:
                break
        if not order_col:
            for c in df.columns:
                if "order" in c.lower():
                    order_col = c
                    break

        # Bygg kundnamns-mapping
        order_to_customer: Dict[str, str] = {}
        orders_path = session.files.get("orders")
        if orders_path and os.path.exists(orders_path):
            try:
                ddf = await loop.run_in_executor(None, lambda: read_csv_auto(orders_path))
                ddf.columns = [str(c).replace("\ufeff", "").strip() for c in ddf.columns]
                ddf = apply_value_filters(ddf, session.active_filters)
                if "Order nr" in ddf.columns and "Kund.1" in ddf.columns:
                    order_to_customer = (ddf.groupby("Order nr")["Kund.1"].first()
                                         .fillna("").astype(str).str.strip().to_dict())
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

        result_df = pd.DataFrame(shipment_diff_rows) if shipment_diff_rows else pd.DataFrame()

        # HIB-kontroll
        ordertype_col = None
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if cl in {"ordertyp", "ordertype"} or ("order" in cl and "typ" in cl):
                ordertype_col = c
                break
        status_col = None
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if cl in {"status", "orderstatus", "radstatus", "state"}:
                status_col = c
                break
        if not status_col:
            for c in df.columns:
                if "status" in c.lower():
                    status_col = c
                    break

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

        has_any = not result_df.empty or not hib_check_df.empty
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
        if not hib_check_df.empty:
            h_df = hib_check_df.copy()
            if "Avvikelsetyp" not in h_df.columns:
                h_df.insert(0, "Avvikelsetyp", "HIB över status 31 utan butikssändning")
            sheets["HIB utan butikssändning"] = h_df
            combined_parts.append(h_df)
        if combined_parts:
            combined = pd.concat(combined_parts, ignore_index=True, sort=False)
            sheets = {"Orderkontroll": combined, **sheets}

        ok_path = os.path.join(session.temp_dir, "orderkontroll.xlsx")
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

        def _find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
            for kw in keywords:
                kw_norm = kw.lower().replace(" ", "")
                for col in df.columns:
                    if col.lower().replace(" ", "") == kw_norm:
                        return col
            for kw in keywords:
                kw_lower = kw.lower()
                for col in df.columns:
                    if kw_lower in col.lower():
                        return col
            return None

        order_kws = ["ordernr", "order nr", "ordernummer", "order number", "orderid"]
        ship_kws = ["sändningsnr", "sändnings nr", "sändningsnummer", "sandningsnr", "sandningsnummer", "shipment"]
        plock_kws = ["plockpallsnr", "plockpallsnr.", "plockpall", "plockpallnr", "plockpallsnummer", "plockpall nr"]

        ov_order_col = _find_col(ov_df, order_kws)
        ov_ship_col = _find_col(ov_df, ship_kws)
        if not ov_order_col or not ov_ship_col:
            log("FEL: Kunde inte identifiera order- eller sändningskolumnen i orderöversikten.")
            log("__ERROR__")
            return

        dp_order_col = _find_col(dp_df, order_kws)
        dp_ship_col = _find_col(dp_df, ship_kws)
        plock_col = _find_col(dp_df, plock_kws)
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
                if "Order nr" in det_df.columns and "Kund.1" in det_df.columns:
                    order_to_customer = (det_df.groupby("Order nr")["Kund.1"].first()
                                         .fillna("").astype(str).str.strip().to_dict())
            except Exception:
                pass

        # Bygg order → sändningsnr mapping från orderöversikten
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

        dk_path = os.path.join(session.temp_dir, "dispatchkontroll.xlsx")
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

        # Samla alla tillgängliga WMS-filer
        wms_files: Dict[str, str] = {}
        for key in ["wms_receive", "wms_booking", "wms_buffert", "wms_trans", "wms_pick", "wms_correct"]:
            path = session.files.get(key)
            if path and os.path.exists(path):
                wms_files[key] = path

        log(f"Eftersök: inköpsnummer={purchase}, artikelnummer={article}")
        log(f"Tillgängliga WMS-filer: {list(wms_files.keys())}")

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
                if not hits.empty:
                    results_sheets[key] = hits
                    log(f"  {key}: {len(hits)} träff(ar)")
                else:
                    log(f"  {key}: inga träffar")
            except Exception as e:
                log(f"  {key}: fel vid läsning – {e}")

        if not results_sheets:
            log("Inga träffar hittades i tillgängliga WMS-filer.")
            log("__DONE__")
            return

        eftersok_path = os.path.join(session.temp_dir, "eftersok.xlsx")
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
async def api_run_allokering(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En körning pågår redan")
    session.running = True
    background_tasks.add_task(_job_allokering, session)
    return {"job_id": "allokering", "status": "started"}


@app.post("/api/run/hib-koppling/{sid}")
async def api_run_hib(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En körning pågår redan")
    session.running = True
    background_tasks.add_task(_job_hib_koppling, session)
    return {"job_id": "hib-koppling", "status": "started"}


@app.post("/api/run/orderkontroll/{sid}")
async def api_run_orderkontroll(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En körning pågår redan")
    session.running = True
    background_tasks.add_task(_job_orderkontroll, session)
    return {"job_id": "orderkontroll", "status": "started"}


@app.post("/api/run/dispatchkontroll/{sid}")
async def api_run_dispatchkontroll(sid: str, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En körning pågår redan")
    session.running = True
    background_tasks.add_task(_job_dispatchkontroll, session)
    return {"job_id": "dispatchkontroll", "status": "started"}


@app.post("/api/run/eftersok/{sid}")
async def api_run_eftersok(sid: str, body: EftersokBody, background_tasks: BackgroundTasks):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
    if session.running:
        raise HTTPException(status_code=409, detail="En körning pågår redan")
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

@app.get("/api/download/{sid}/{result_key}")
def api_download(sid: str, result_key: str):
    session = get_session(sid)
    if not session:
        raise HTTPException(status_code=404, detail="Session saknas")
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
# Statiska filer – montera SIST
# ---------------------------------------------------------------------------

_frontend_dir = Path(__file__).parent.parent / "frontend"
if _frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(_frontend_dir), html=True), name="frontend")
