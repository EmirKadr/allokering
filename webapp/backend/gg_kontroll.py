"""
gg_kontroll.py – GG Kontrollpanel: filuppladdning, bearbetning och resultat-API.

Delade GG-data (en LKON-användare laddar upp, alla ser resultaten).
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import APIRouter, Header, HTTPException, Query, UploadFile
from fastapi.responses import JSONResponse
from psycopg.types.json import Jsonb

try:
    from .db import connect, get_state_dir, utcnow
    from .auth import extract_token, require_auth
    from .logic import read_csv_auto
except ImportError:
    from db import connect, get_state_dir, utcnow
    from auth import extract_token, require_auth
    from logic import read_csv_auto

router = APIRouter(prefix="/api/gg", tags=["gg"])

GG_FILE_SLOTS = [
    {"key": "gg_orders", "label": "Data Child (Detalj Kundorder)"},
    {"key": "gg_overview", "label": "Data OÖ (Orderöversikt)"},
    {"key": "gg_plocklogg", "label": "Data Plocklogg"},
    {"key": "gg_palluppdrag", "label": "Inmatning Palluppdrag"},
    {"key": "gg_dispatch", "label": "Dispatchpallar"},
]

VALID_FILE_KEYS = {s["key"] for s in GG_FILE_SLOTS}

# Kolumner som behövs per fil (case-insensitive matching)
GG_NEEDED_COLS: Dict[str, List[str]] = {
    "gg_plocklogg":   ["status", "zon", "användare", "datum", "timestamp", "ändrad", "andrad", "bolag"],
    "gg_orders":      ["zon", "bolag", "status", "kund"],
    "gg_overview":    ["avgångstid", "zon", "rader", "transportör", "status", "bolag"],
    "gg_palluppdrag": ["zon", "lagerplats", "krangång", "bolag"],
    "gg_dispatch":    ["pallplacering", "kund", "palltyp", "pallbeskrivning", "bolag"],
}


def _gg_dir() -> Path:
    d = get_state_dir() / "gg_kontroll"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _require_gg_or_admin(token: Optional[str]) -> dict:
    user = require_auth(token)
    if user.get("role") == "admin":
        return user
    if "gg" in (user.get("lists") or []):
        return user
    raise HTTPException(403, "GG-behörighet krävs")


# ---------------------------------------------------------------------------
# Upload / list / delete
# ---------------------------------------------------------------------------


@router.post("/upload")
async def gg_upload(
    file: UploadFile,
    key: str = Query(...),
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    user = _require_gg_or_admin(tok)

    if key not in VALID_FILE_KEYS:
        raise HTTPException(400, f"Ogiltig filnyckel: {key}")

    dest_dir = _gg_dir()
    ext = Path(file.filename or "file").suffix or ".csv"
    dest_path = dest_dir / f"{key}{ext}"

    content = await file.read()
    dest_path.write_bytes(content)

    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO gg_files (file_key, path, original_filename, uploaded_by, uploaded_at)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT (file_key)
                DO UPDATE SET
                    path = EXCLUDED.path,
                    original_filename = EXCLUDED.original_filename,
                    uploaded_by = EXCLUDED.uploaded_by,
                    uploaded_at = EXCLUDED.uploaded_at
                """,
                (key, str(dest_path), file.filename or "", user.get("username", ""), now),
            )
        conn.commit()

    return {"ok": True, "key": key, "filename": file.filename}


@router.get("/files")
def gg_files(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    require_auth(tok)

    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT file_key, path, original_filename, uploaded_by, uploaded_at
                FROM gg_files
                ORDER BY file_key
                """
            )
            rows = cur.fetchall()

    files = {}
    stale_keys = []
    for r in rows:
        # Only report file as uploaded if it actually exists on disk
        if r["path"] and Path(r["path"]).exists():
            files[r["file_key"]] = {
                "filename": r["original_filename"],
                "uploaded_by": r["uploaded_by"],
                "uploaded_at": r["uploaded_at"].isoformat() if r["uploaded_at"] else None,
            }
        else:
            stale_keys.append(r["file_key"])

    # Clean up stale DB entries where file no longer exists
    if stale_keys:
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "DELETE FROM gg_files WHERE file_key = ANY(%s)", (stale_keys,)
                )
            conn.commit()

    return {"files": files}


@router.delete("/upload/{key}")
def gg_delete_file(
    key: str,
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    _require_gg_or_admin(tok)

    if key not in VALID_FILE_KEYS:
        raise HTTPException(400, f"Ogiltig filnyckel: {key}")

    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT path FROM gg_files WHERE file_key = %s", (key,))
            row = cur.fetchone()
            if row:
                try:
                    Path(row["path"]).unlink(missing_ok=True)
                except Exception:
                    pass
                cur.execute("DELETE FROM gg_files WHERE file_key = %s", (key,))
        conn.commit()

    return {"ok": True, "key": key}


@router.post("/reset")
def gg_reset(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    """Rensa alla GG-filer och resultat."""
    tok = extract_token(authorization, token)
    _require_gg_or_admin(tok)

    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT path FROM gg_files")
            rows = cur.fetchall()
            for row in rows:
                try:
                    Path(row["path"]).unlink(missing_ok=True)
                except Exception:
                    pass
            cur.execute("DELETE FROM gg_files")
            cur.execute("DELETE FROM gg_results")
        conn.commit()

    return {"ok": True}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_gg_file_path(key: str) -> Optional[str]:
    """Return the file path for a GG file key, or None."""
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT path FROM gg_files WHERE file_key = %s", (key,))
            row = cur.fetchone()
    if not row:
        return None
    path = row["path"]
    if not Path(path).exists():
        return None
    return path


def _load_gg_file(key: str) -> Optional[pd.DataFrame]:
    """Ladda GG-fil (redan strippade vid upload, så den är liten)."""
    path = _load_gg_file_path(key)
    if not path:
        return None
    if path.endswith(".xlsx"):
        df = pd.read_excel(path, dtype=str)
        needed = GG_NEEDED_COLS.get(key)
        return _keep_cols(df, needed) if needed else df
    return read_csv_auto(path)


def _keep_cols(df: pd.DataFrame, needed: List[str]) -> pd.DataFrame:
    """Keep only columns whose name (lowered) matches the needed list."""
    needed_set = {n.lower() for n in needed}
    keep = [c for c in df.columns if c.strip().lower() in needed_set]
    return df[keep] if keep else df


def _find_col(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    """Find first matching column name (case-insensitive, stripped)."""
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    return None


def _save_result(key: str, data: dict, username: str) -> None:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO gg_results (result_key, data_json, processed_by, processed_at)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (result_key)
                DO UPDATE SET
                    data_json = EXCLUDED.data_json,
                    processed_by = EXCLUDED.processed_by,
                    processed_at = EXCLUDED.processed_at
                """,
                (key, Jsonb(data), username, now),
            )
        conn.commit()


def _get_result(key: str) -> Optional[dict]:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT data_json, processed_by, processed_at FROM gg_results WHERE result_key = %s",
                (key,),
            )
            row = cur.fetchone()
    if not row:
        return None
    data = row["data_json"] or {}
    data["_processed_by"] = row.get("processed_by", "")
    data["_processed_at"] = row["processed_at"].isoformat() if row.get("processed_at") else None
    return data


# ---------------------------------------------------------------------------
# Processing logic
# ---------------------------------------------------------------------------

HOURS = [f"{h:02d}:00-{h+1:02d}:00" if h < 23 else "23:00-00:00" for h in range(6, 24)]
ZONES_PRODUKTION = ["A", "S", "F", "E"]


def _process_produktion(
    plocklogg: pd.DataFrame,
    bolag_filter: Optional[List[str]] = None,
    orders_df: Optional[pd.DataFrame] = None,
    overview_df: Optional[pd.DataFrame] = None,
) -> dict:
    """Beräkna plockproduktion per zon och timme från plockloggen."""
    df = plocklogg.copy()

    # Find relevant columns
    status_col = _find_col(df, "Status")
    zon_col = _find_col(df, "Zon")
    user_col = _find_col(df, "Användare")
    datum_col = _find_col(df, "Datum")
    timestamp_col = _find_col(df, "Timestamp")
    bolag_col = _find_col(df, "Bolag")
    beställt_col = _find_col(df, "Beställt")

    if not zon_col or not user_col:
        return {"error": "Saknar kolumner Zon/Användare i plockloggen"}

    # Filter by bolag
    df = _filter_bolag(df, bolag_filter)

    # Filter completed picks (status 35 = slutförd)
    if status_col:
        df[status_col] = pd.to_numeric(df[status_col], errors="coerce")
        df = df[df[status_col] == 35]

    # Parse hour – prefer Ändrad (has full timestamp), fall back to Datum
    andrad_col = _find_col(df, "Ändrad", "Andrad")
    time_col = andrad_col or timestamp_col or datum_col
    if not time_col:
        return {"error": "Saknar Datum/Ändrad/Timestamp-kolumn i plockloggen"}

    df["_ts"] = pd.to_datetime(df[time_col], errors="coerce", dayfirst=False)
    df = df.dropna(subset=["_ts"])

    # Use the most common date as "today"
    df["_date"] = df["_ts"].dt.date
    if df.empty:
        today = None
    else:
        today = df["_date"].mode().iloc[0]
        df = df[df["_date"] == today]

    df["_hour"] = df["_ts"].dt.hour

    zones = {}
    for zone in ZONES_PRODUKTION:
        zdf = df[df[zon_col].str.strip().str.upper() == zone]
        pickers = []
        avg = []
        rows_list = []
        for h in range(6, 24):
            hdf = zdf[zdf["_hour"] == h]
            unique_users = hdf[user_col].nunique()
            row_count = len(hdf)
            pickers.append(unique_users)
            rows_list.append(row_count)
            avg.append(round(row_count / unique_users, 1) if unique_users > 0 else 0)
        zones[zone] = {
            "label": f"Zon {zone}",
            "pickers": pickers,
            "rows": rows_list,
            "avg": avg,
        }

    # --- Statistik dagen ---
    stats_dagen: List[Dict[str, Any]] = []
    # Count total ordered rows per zone (from orders_df)
    orders_per_zone: Dict[str, int] = {}
    if orders_df is not None and not orders_df.empty:
        odf = orders_df.copy()
        odf = _filter_bolag(odf, bolag_filter)
        zon_col_o = _find_col(odf, "Zon")
        if zon_col_o:
            zon_counts = odf[zon_col_o].str.strip().str.upper().value_counts()
            for z in ZONES_PRODUKTION:
                orders_per_zone[z] = int(zon_counts.get(z, 0))

    for zone in ZONES_PRODUKTION:
        z_data = zones[zone]
        total_picked = sum(z_data["rows"])
        active_hours = sum(1 for r in z_data["rows"] if r > 0)
        avg_per_hour = round(total_picked / active_hours, 1) if active_hours > 0 else 0
        total_orders = orders_per_zone.get(zone, 0)
        rader_kvar = max(0, total_orders - total_picked)
        tid_kvar = round(rader_kvar / avg_per_hour, 1) if avg_per_hour > 0 else 0
        stats_dagen.append({
            "zone": zone,
            "label": f"Zon {zone}",
            "avg_per_hour": avg_per_hour,
            "rader_kvar": rader_kvar,
            "tid_kvar": tid_kvar,
        })

    # --- Tid kvar till avgång ---
    tid_kvar_avgang: List[Dict[str, Any]] = []
    if overview_df is not None and not overview_df.empty:
        ovdf = overview_df.copy()
        ovdf = _filter_bolag(ovdf, bolag_filter)
        avgtid_col = _find_col(ovdf, "Avgångstid", "Avg\u00e5ngstid")
        zon_col_ov = _find_col(ovdf, "Zon")
        rader_col_ov = _find_col(ovdf, "Rader")
        status_col_ov = _find_col(ovdf, "Status")

        if avgtid_col and zon_col_ov and rader_col_ov:
            ovdf["_avgtid"] = pd.to_datetime(ovdf[avgtid_col], errors="coerce", dayfirst=True)
            ovdf = ovdf.dropna(subset=["_avgtid"])
            ovdf["_avgtid_str"] = ovdf["_avgtid"].dt.strftime("%H:%M")
            ovdf["_zon_upper"] = ovdf[zon_col_ov].str.strip().str.upper()
            ovdf["_rader"] = pd.to_numeric(ovdf[rader_col_ov], errors="coerce").fillna(0).astype(int)

            # Check status to determine if departure is done (all orders picked)
            if status_col_ov:
                ovdf["_status_num"] = pd.to_numeric(ovdf[status_col_ov], errors="coerce").fillna(0)

            for tid, group in sorted(ovdf.groupby("_avgtid_str"), key=lambda x: x[0]):
                row_data: Dict[str, Any] = {"time": tid}
                for z in ZONES_PRODUKTION:
                    zg = group[group["_zon_upper"] == z]
                    total_rader = int(zg["_rader"].sum())
                    # If status >= 35 for all orders in this zone/departure → done
                    if status_col_ov and len(zg) > 0:
                        all_done = bool((zg["_status_num"] >= 35).all())
                    else:
                        all_done = total_rader == 0
                    row_data[z] = "Avgång klar" if all_done else str(total_rader)
                tid_kvar_avgang.append(row_data)

    return {
        "date": str(today) if today else None,
        "hours": HOURS,
        "zones": zones,
        "stats_dagen": stats_dagen,
        "tid_kvar_avgang": tid_kvar_avgang,
    }


def _process_dagsoversikt(
    orders_df: Optional[pd.DataFrame],
    overview_df: Optional[pd.DataFrame],
    palluppdrag_df: Optional[pd.DataFrame],
    bolag_filter: Optional[List[str]] = None,
) -> dict:
    """Beräkna dagsöversikt: totala rader, rader per avgång, transportörer."""

    totals: List[Dict[str, Any]] = []
    departures: List[Dict[str, Any]] = []
    transporters: List[Dict[str, Any]] = []
    ehandel: List[Dict[str, Any]] = []
    avgang_klar: List[Dict[str, Any]] = []

    # --- Totala rader (from orders / customer_order_details_all) ---
    if orders_df is not None and not orders_df.empty:
        df = orders_df.copy()
        zon_col = _find_col(df, "Zon")
        bolag_col = _find_col(df, "Bolag")
        status_col = _find_col(df, "Status")
        # Find the "Kund" column that has customer names (not numbers)
        kund_col = None
        kund_candidates = [c for c in df.columns if c.strip().lower().startswith("kund")]
        for c in kund_candidates:
            if df[c].astype(str).str.contains("Ehandel|ehandel", case=False, na=False).any():
                kund_col = c
                break
        if not kund_col and kund_candidates:
            kund_col = kund_candidates[-1]  # last Kund column is usually the name

        df = _filter_bolag(df, bolag_filter)

        if zon_col:
            zone_counts = df[zon_col].str.strip().str.upper().value_counts()
            zone_labels = {"A": "Huvudplock", "S": "Skrymmande", "F": "HIB", "E": "E-handel", "R": "Autostore"}
            for z in ["A", "S", "F", "E", "R"]:
                totals.append({
                    "zone": z,
                    "label": zone_labels.get(z, z),
                    "count": int(zone_counts.get(z, 0)),
                })

        # Ogenererat AS: status < 30 and zone R
        if status_col and zon_col:
            df_num = df.copy()
            df_num["_status_num"] = pd.to_numeric(df_num[status_col], errors="coerce")
            df_num["_zon_upper"] = df_num[zon_col].str.strip().str.upper()
            ogenererat = len(df_num[(df_num["_status_num"] < 30) & (df_num["_zon_upper"] == "R")])
            totals.append({"zone": "", "label": "Ogenererat AS", "count": ogenererat})

        # --- Ehandel ---
        if zon_col and status_col and kund_col:
            df_eh = df.copy()
            df_eh["_status_num"] = pd.to_numeric(df_eh[status_col], errors="coerce")
            df_eh["_zon_upper"] = df_eh[zon_col].str.strip().str.upper()
            df_eh["_is_ehandel"] = df_eh[kund_col].astype(str).str.contains("Ehandel", case=False, na=False)
            # EH Q (Manuellt): zone E, status 30, kund = Ehandel
            eh_q = len(df_eh[
                (df_eh["_status_num"] == 30) &
                (df_eh["_zon_upper"] == "E") &
                (df_eh["_is_ehandel"])
            ])
            # EH AS (Autostore): zone R, status in (31,32,33), kund = Ehandel
            eh_as = len(df_eh[
                (df_eh["_status_num"].isin([31, 32, 33])) &
                (df_eh["_zon_upper"] == "R") &
                (df_eh["_is_ehandel"])
            ])
            ehandel = [
                {"key": "EH Q", "label": "Manuellt", "count": eh_q},
                {"key": "EH AS", "label": "Autostore", "count": eh_as},
            ]

    # Helpall/Påfyll from palluppdrag
    if palluppdrag_df is not None and not palluppdrag_df.empty:
        pdf = palluppdrag_df.copy()
        pdf = _filter_bolag(pdf, bolag_filter)

        zon_col_p = _find_col(pdf, "Zon")
        lagerplats_col = _find_col(pdf, "Lagerplats")
        status_col_p = _find_col(pdf, "Status")
        krangång_col = _find_col(pdf, "Krangång", "Krang\u00e5ng")

        helpall_kran = 0
        helpall_skjutare = 0
        pafyll_kran = 0
        pafyll_skjutare = 0
        helpall_eltak = 0

        if lagerplats_col:
            for _, row in pdf.iterrows():
                lp = str(row.get(lagerplats_col, "") or "").strip().upper()
                zon = str(row.get(zon_col_p, "") or "").strip().upper() if zon_col_p else ""
                kg = str(row.get(krangång_col, "") or "").strip() if krangång_col else ""

                if lp.startswith("UT") or lp.startswith("HELPALL"):
                    if "ELTAK" in lp:
                        helpall_eltak += 1
                    elif kg and kg != "0":
                        helpall_kran += 1
                    else:
                        helpall_skjutare += 1
                elif lp.startswith("PÅFYLL") or lp.startswith("PAFYLL"):
                    if kg and kg != "0":
                        pafyll_kran += 1
                    else:
                        pafyll_skjutare += 1

        totals.extend([
            {"zone": "", "label": "Helpall Kran", "count": helpall_kran},
            {"zone": "", "label": "Helpall Skjutare", "count": helpall_skjutare},
            {"zone": "", "label": "Påfyll Kran", "count": pafyll_kran},
            {"zone": "", "label": "Påfyll Skjutare", "count": pafyll_skjutare},
            {"zone": "", "label": "Helpall ELTAK", "count": helpall_eltak},
        ])

    # --- Rader per avgång (from overview) ---
    if overview_df is not None and not overview_df.empty:
        odf = overview_df.copy()
        odf = _filter_bolag(odf, bolag_filter)

        avgangstid_col = _find_col(odf, "Avgångstid", "Avg\u00e5ngstid")
        zon_col_o = _find_col(odf, "Zon")
        rader_col = _find_col(odf, "Rader")

        if avgangstid_col and zon_col_o and rader_col:
            odf["_avgtid"] = pd.to_datetime(odf[avgangstid_col], errors="coerce", dayfirst=True)
            odf = odf.dropna(subset=["_avgtid"])
            odf["_avgtid_str"] = odf["_avgtid"].dt.strftime("%H:%M")
            odf["_zon_upper"] = odf[zon_col_o].str.strip().str.upper()
            odf["_rader"] = pd.to_numeric(odf[rader_col], errors="coerce").fillna(0).astype(int)

            status_col_o = _find_col(odf, "Status")
            if status_col_o:
                odf["_status_num"] = pd.to_numeric(odf[status_col_o], errors="coerce").fillna(0)

            dep_groups = odf.groupby("_avgtid_str")
            for tid, group in sorted(dep_groups, key=lambda x: x[0]):
                row_data = {"time": tid}
                klar_data: Dict[str, Any] = {"time": tid}
                for z in ["A", "S", "F", "E"]:
                    zg = group[group["_zon_upper"] == z]
                    row_data[z] = int(zg["_rader"].sum())
                    # Avgång klar: all orders status >= 35 → done
                    if status_col_o and len(zg) > 0:
                        all_done = bool((zg["_status_num"] >= 35).all())
                        klar_data[z] = "Klar" if all_done else int(zg["_rader"].sum())
                    else:
                        klar_data[z] = "" if len(zg) == 0 else int(zg["_rader"].sum())
                departures.append(row_data)
                avgang_klar.append(klar_data)

        # --- Transportörer (from overview) ---
        transportor_col = _find_col(odf, "Transportör", "Transport\u00f6r")
        if transportor_col and zon_col_o and rader_col:
            odf["_zon_cat"] = odf["_zon_upper"].map(
                lambda z: "A" if z == "A" else ("O" if z in ("S",) else "F" if z in ("F", "E", "R") else z)
            )
            trans_groups = odf.groupby(transportor_col)
            for name, group in sorted(trans_groups, key=lambda x: x[0]):
                if not name or str(name).strip() == "":
                    continue
                a_sum = int(group[group["_zon_cat"] == "A"]["_rader"].sum())
                o_sum = int(group[group["_zon_cat"] == "O"]["_rader"].sum())
                f_sum = int(group[group["_zon_cat"] == "F"]["_rader"].sum())
                transporters.append({
                    "name": str(name).strip(),
                    "A": a_sum,
                    "O": o_sum,
                    "F": f_sum,
                    "total": a_sum + o_sum + f_sum,
                })

    return {
        "totals": totals,
        "departures": departures,
        "transporters": transporters,
        "ehandel": ehandel,
        "avgang_klar": avgang_klar,
    }


def _process_lastlista(dispatch_df: pd.DataFrame, bolag_filter: Optional[List[str]] = None) -> dict:
    """Beräkna lastlista från dispatch-data."""
    df = dispatch_df.copy()

    df = _filter_bolag(df, bolag_filter)

    pallplacering_col = _find_col(df, "Pallplacering")
    kund_col = _find_col(df, "Kund")
    palltyp_col = _find_col(df, "Palltyp")
    pallbeskrivning_col = _find_col(df, "Pallbeskrivning")

    if not pallplacering_col or not kund_col:
        return {"rows": [], "error": "Saknar Pallplacering/Kund i dispatch-filen"}

    result_rows: Dict[str, Dict[str, Any]] = {}

    for _, row in df.iterrows():
        ytor = str(row.get(pallplacering_col, "") or "").strip()
        kund = str(row.get(kund_col, "") or "").strip()
        if not ytor:
            continue

        group_key = f"{ytor}|{kund}"
        if group_key not in result_rows:
            result_rows[group_key] = {"ytor": ytor, "kund": kund, "pallar": 0, "halvpallar": 0, "sma_paket": 0}

        beskr = str(row.get(pallbeskrivning_col, "") or "").strip().upper() if pallbeskrivning_col else ""
        typ_val = str(row.get(palltyp_col, "") or "").strip() if palltyp_col else ""

        if "EUROPALL" in beskr or "EUROPALL" in typ_val.upper():
            result_rows[group_key]["pallar"] += 1
        elif "HALVPALL" in beskr or "HALVPALL" in typ_val.upper():
            result_rows[group_key]["halvpallar"] += 1
        elif "LITET" in beskr or "SMÅ" in beskr or "SMA" in beskr or "PAKET" in beskr:
            result_rows[group_key]["sma_paket"] += 1
        else:
            # Default: count as pall
            result_rows[group_key]["pallar"] += 1

    rows_list = sorted(result_rows.values(), key=lambda r: r["ytor"])
    return {"rows": rows_list}


# ---------------------------------------------------------------------------
# Process endpoint
# ---------------------------------------------------------------------------


def _read_bolag_values(path: str) -> set:
    """Read only the Bolag column from a CSV file (memory-efficient)."""
    try:
        import csv as csv_mod
        sep = _detect_separator_for_bolag(path)
        with open(path, encoding="utf-8-sig") as f:
            reader = csv_mod.reader(f, delimiter=sep)
            header = next(reader, None)
            if not header:
                return set()
            bolag_idx = None
            for i, col in enumerate(header):
                if col.strip().lower() == "bolag":
                    bolag_idx = i
                    break
            if bolag_idx is None:
                return set()
            vals = set()
            for row in reader:
                if bolag_idx < len(row):
                    v = row[bolag_idx].strip().upper()
                    if v:
                        vals.add(v)
            return vals
    except Exception:
        return set()


def _detect_separator_for_bolag(path: str) -> str:
    """Quick separator detection for Bolag extraction."""
    try:
        with open(path, encoding="utf-8-sig") as f:
            first_line = f.readline()
        if "\t" in first_line:
            return "\t"
        if ";" in first_line:
            return ";"
    except Exception:
        pass
    return ","


@router.get("/filter-options")
def gg_filter_options(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    """Returnera unika Bolag-värden från uppladdade filer (minneseffektivt)."""
    tok = extract_token(authorization, token)
    require_auth(tok)
    bolag_set: set = set()
    for key in ["gg_orders", "gg_overview", "gg_plocklogg", "gg_palluppdrag", "gg_dispatch"]:
        path = _load_gg_file_path(key)
        if path:
            bolag_set.update(_read_bolag_values(path))
    return {"bolag": sorted(bolag_set)}


def _filter_bolag(df: pd.DataFrame, bolag_filter: Optional[List[str]]) -> pd.DataFrame:
    """Filtrera DataFrame på Bolag om filter är satt."""
    if not bolag_filter:
        return df
    col = _find_col(df, "Bolag")
    if not col:
        return df
    return df[df[col].str.strip().str.upper().isin(bolag_filter)]


@router.post("/process")
def gg_process(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
    bolag: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    user = _require_gg_or_admin(tok)
    username = user.get("username", "")

    # Parse bolag filter
    bolag_filter = [b.strip().upper() for b in bolag.split(",") if b.strip()] if bolag else None

    errors = []

    # Load shared files (used by both Produktion and Dagsöversikt)
    orders = None
    overview = None
    try:
        orders = _load_gg_file("gg_orders")
        overview = _load_gg_file("gg_overview")
    except Exception as e:
        errors.append(f"Laddningsfel orders/overview: {e}")

    # Produktion
    try:
        plocklogg = _load_gg_file("gg_plocklogg")
        if plocklogg is not None:
            result = _process_produktion(plocklogg, bolag_filter, orders_df=orders, overview_df=overview)
            _save_result("produktion", result, username)
            del plocklogg  # free memory
        else:
            errors.append("Plocklogg saknas — Produktion ej beräknad")
    except Exception as e:
        errors.append(f"Produktion-fel: {e}")

    # Dagsöversikt
    try:
        palluppdrag = _load_gg_file("gg_palluppdrag")

        if orders is not None or overview is not None:
            result = _process_dagsoversikt(orders, overview, palluppdrag, bolag_filter)
            _save_result("dagsoversikt", result, username)
        else:
            errors.append("Orders/Overview saknas — Dagsöversikt ej beräknad")
        del orders, overview, palluppdrag
    except Exception as e:
        errors.append(f"Dagsöversikt-fel: {e}")

    # Lastlista
    try:
        dispatch = _load_gg_file("gg_dispatch")
        if dispatch is not None:
            result = _process_lastlista(dispatch, bolag_filter)
            _save_result("lastlista", result, username)
            del dispatch
        else:
            errors.append("Dispatch saknas — Lastlista ej beräknad")
    except Exception as e:
        errors.append(f"Lastlista-fel: {e}")

    return {"ok": True, "errors": errors}


# ---------------------------------------------------------------------------
# Read endpoints
# ---------------------------------------------------------------------------


@router.get("/produktion")
def gg_produktion(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    require_auth(tok)
    data = _get_result("produktion")
    if not data:
        return {"data": None}
    return {"data": data}


@router.get("/dagsoversikt")
def gg_dagsoversikt(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    require_auth(tok)
    data = _get_result("dagsoversikt")
    if not data:
        return {"data": None}
    return {"data": data}


@router.get("/lastlista")
def gg_lastlista(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None),
):
    tok = extract_token(authorization, token)
    require_auth(tok)
    data = _get_result("lastlista")
    if not data:
        return {"data": None}
    return {"data": data}
