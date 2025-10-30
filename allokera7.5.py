#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
allokera5.0.py
--------------
- Nytt: Auto-refill
  - Refill (HP = HUVUDPLOCK + SKRYMMANDE samt AUTOSTORE) beräknas automatiskt direkt efter allokering.
  - "Öppna påfyllningspallar" visar den senast auto-beräknade rapporten (ingen extra beräkning krävs).
- Nytt: Sales (frivilligt)
  - Filväljare för plocklogg (CSV). När filen laddas beräknas försäljningsinsikter direkt:
    Total_7/30/90, ADV_30/90, SenastPlockad, DagarSedanSenast, UnikaPlockdagar_90,
    ABC_klass (Pareto 80/15/5), NollraderPerPlockdag_90 (medel antal rader med Plockat=0 per aktiv plockdag).
  - Ny knapp "Öppna försäljningsinsikter" öppnar en Excel med:
    "Top sellers (90d)", "Slow movers (≥90d/0)" och "Sammanställning".
  - Valfri annotering: Om sales finns läggs kolumner (ADV_90, UnikaPlockdagar_90, ABC_klass,
    DagarSedanSenast, NollraderPerPlockdag_90) till vid visning av refill-bladen. Påverkar inte logiken.
- Förbättrat: Zonbaserad omklassificering
  - Efter att HELPALL och AUTOSTORE har allokerats används orderfilens "Zon"‑kolumn
    för att klassificera huvudplock. Om zonkoden är "S", "E", "A", "Q", "O" eller "F"
    sätts Källtyp och "Zon (beräknad)" enligt mappningen: S→SKRYMMANDE, E→EHANDEL,
    A→HUVUDPLOCK, Q→EHANDEL, O→SKRYMMANDE, F→BRAND. Ingen plockplats/saldofil krävs.
- Refill-logik:
  - HP-bladet inkluderar zon A + S; plocksaldo dras EN gång per artikel och fördelas proportionellt.
  - 0-rader filtreras bort (HP + AUTOSTORE). HELPALL som använts exkluderas alltid från refill.
- Rensning/underhåll:
  - Tidigare checkbox för “exkludera HELPALL” och all relaterad kod borttagen (beteende alltid aktivt).
  - Ingen "Öppna ej inlagrade"-knapp längre (filen kan fortfarande användas för kolumnen i refill).
"""

from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Deque, Dict, List, Tuple, Optional

# Drag & drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

from collections import defaultdict, deque
import pandas as pd
import tempfile
import os
import sys
import subprocess
import numpy as np

# --- Prognos-inläsning (frivillig) ---
# Följande funktion normaliserar en prognosfil (XLSX) genom att kasta bort
# onödiga rader/kolumner och standardisera rubriker. Den returnerar kolumnerna:
#  Artikelnummer[str], Beskrivning[str], Antal styck[int], Antal rader[int], Antal butiker[int].
def read_prognos_xlsx(path: str) -> pd.DataFrame:
    """
    Läser en prognos (XLSX) och returnerar ett normaliserat DataFrame.
    Steg:
      1) Ta bort de tre första raderna (index 0,1,3) om de finns.
      2) Ta bort kolumn A (första kolumnen).
      3) Använd första kvarvarande rad som rubriker och plocka ut relevanta kolumner.

    Returnerar DataFrame med kolumner:
      - Artikelnummer (str)
      - Beskrivning (str)
      - Antal styck (int)
      - Antal rader (int)
      - Antal butiker (int)
    """
    df = pd.read_excel(path, header=None, dtype=object)
    # Om tom fil
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
    # Rader att droppa: index 0,1,3
    drop_idx = [i for i in (0, 1, 3) if i < len(df.index)]
    df = df.drop(index=drop_idx, errors="ignore").reset_index(drop=True)
    # Droppa första kolumnen (kolumn A)
    if df.shape[1] > 0:
        df = df.drop(columns=[df.columns[0]]).reset_index(drop=True)
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
    # Första rad som rubriker
    header = df.iloc[0].astype(str).str.strip().tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = header
    # Matcha kolumnnamn (case-insensitivt, alfanumeriskt)
    def _ci_match(name: str) -> str:
        return "".join(c.lower() for c in str(name).strip() if c.isalnum())
    def _pick_col(cols: List[str], candidates: List[str]) -> str | None:
        s_cols = { _ci_match(c): c for c in cols }
        for cand in candidates:
            key = _ci_match(cand)
            if key in s_cols:
                return s_cols[key]
        return None
    need_map: Dict[str, List[str]] = {
        "Artikelnummer": ["Product code", "SKU", "Artikelnr", "Artikelnummer"],
        "Beskrivning":   ["Product name", "Name", "Benämning", "Beskrivning"],
        "Antal styck":   ["Antal styck", "Antal stycken", "Qty", "Quantity"],
        "Antal rader":   ["Antal rader", "Rows", "Number of rows"],
        "Antal butiker": ["Antal butiker", "Stores", "Butiker", "Number of stores"],
    }
    picked: Dict[str, str] = {}
    for out_name, candidates in need_map.items():
        col = _pick_col(list(df.columns), candidates)
        if col:
            picked[out_name] = col
    out = pd.DataFrame()
    for out_name in ["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"]:
        if out_name in picked:
            out[out_name] = df[picked[out_name]]
        else:
            out[out_name] = pd.Series([None] * len(df), dtype=object)
    out["Artikelnummer"] = out["Artikelnummer"].astype(str).str.strip()
    out["Beskrivning"]   = out["Beskrivning"].astype(str).str.strip()
    for num_col in ["Antal styck", "Antal rader", "Antal butiker"]:
        out[num_col] = pd.to_numeric(out[num_col], errors="coerce").fillna(0).astype(int)
    mask_keep = out["Artikelnummer"].str.len().gt(0) | out["Beskrivning"].str.len().gt(0)
    out = out.loc[mask_keep].reset_index(drop=True)
    return out

# --- Kampanjvolymer (frivillig) ---
# Denna funktion normaliserar en kampanjvolymfil (XLSX) enligt en specifik sekvens av rader och kolumner som ska tas bort.
# Den förväntar sig ett Excel‑ark med samma struktur som "Granngården prognos kampanjvolymer per dag". Stegen är:
#   1. Ta bort rad 5 (1‑baserat index) — motsvarar index 4.
#   2. Ta bort raderna 1–3 (index 0–2).
#   3. Ta bort alla kolumner från och med kolumn G *efter* första behållna kolumn (det vill säga, behåll upp till kolumn index 6). Detta
#      innebär att vi sparar de sju första kolumnerna (index 0–6) och kastar resten.
#   4. Ta bort kolumn F, E, D, B och A i just den ordningen. Efter dessa steg återstår kolumn C (Produktkod) och kolumn G
#      (Projicerat antal) från originalfilen. Dessa döps till 'Artikelnummer' respektive 'Antal styck'.
# Resultatet är en DataFrame med två kolumner: Artikelnummer (str) och Antal styck (int).

def read_campaign_xlsx(path: str) -> pd.DataFrame:
    """
    Läs och normalisera en kampanjvolymfil (XLSX) enligt en fördefinierad sekvens av borttagningar av rader och kolumner.
    Returnerar ett DataFrame med kolumnerna:
      - Artikelnummer (str)
      - Antal styck (int)
    """
    # Läs hela arket utan rubriker
    df = pd.read_excel(path, header=None, dtype=object)
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Antal styck"])
    # 1) Ta bort rad 5 (index 4)
    if len(df.index) > 4:
        df = df.drop(index=[4])
    # 2) Ta bort rader 1–3 (index 0–2)
    drop_idx = [i for i in (0, 1, 2) if i < len(df.index)]
    df = df.drop(index=drop_idx)
    # Återställ index efter rad‑borttagning
    df = df.reset_index(drop=True)
    # 3) Ta bort alla kolumner från G och höger. G är kolumn 7 (A=1) → index 6 (0‑baserat).
    # Vi vill behålla kolumnerna 0–6 (inklusive) och kasta övriga.
    keep_cols = [c for c in df.columns if c <= 6]
    df = df.loc[:, keep_cols]
    # 4) Ta bort kolumn F (index 5) om den finns
    if 5 in df.columns:
        df = df.drop(columns=[5])
    # 5) Ta bort kolumn E (index 4) om den finns
    if 4 in df.columns:
        df = df.drop(columns=[4])
    # 6) Ta bort kolumn D (index 3) om den finns
    if 3 in df.columns:
        df = df.drop(columns=[3])
    # 7) Ta bort kolumn B (index 1) om den finns
    if 1 in df.columns:
        df = df.drop(columns=[1])
    # 8) Ta bort kolumn A (index 0) om den finns
    if 0 in df.columns:
        df = df.drop(columns=[0])
    # Efter dessa steg bör endast två kolumner återstå, med index som ursprungligen var 2 (Produktkod) och 6 (Projicerat antal)
    # Säkerställ att vi bara har två kolumner; annars returnera tom df
    if df.shape[1] != 2:
        return pd.DataFrame(columns=["Artikelnummer", "Antal styck"])
    # Ge dem rubriker
    df = df.reset_index(drop=True)
    df.columns = ["Artikelnummer", "Antal styck"]
    # Rensa och konvertera
    df["Artikelnummer"] = df["Artikelnummer"].astype(str).str.strip()
    df["Antal styck"] = pd.to_numeric(df["Antal styck"], errors="coerce").fillna(0).astype(int)
    # Filtrera bort rader där artikelnummer saknas
    df = df.loc[df["Artikelnummer"].astype(str).str.len().gt(0)].reset_index(drop=True)
    # Droppa header-rad om Artikelnummer är exakt "Produktkod" (kan förekomma efter radborttagning)
    if not df.empty and str(df.loc[0, "Artikelnummer"]).lower() in ("produktkod", "#"):
        df = df.drop(index=[0]).reset_index(drop=True)
    return df

# --- Konfig/konstanter ----------------------------------------------------

APP_TITLE = "Buffertpallar → Order-allokering (GUI) — 5.0"
DEFAULT_OUTPUT = "allocated_orders.csv"

INVALID_LOC_PREFIXES: Tuple[str, ...] = ("AA",)
INVALID_LOC_EXACT: set[str] = {"TRANSIT", "TRANSIT_ERROR", "MISSING", "UT2"}

# Allokering använder 29/30/32
ALLOC_BUFFER_STATUSES: set[int] = {29, 30, 32}
# Refill använder 29/30
REFILL_BUFFER_STATUSES: set[int] = {29, 30}

#
# Near‐miss threshold expressed as a proportion of the remaining need.  A pallet will be
# considered a “near miss” if it exceeds the outstanding demand by no more than this
# fraction.  The default used to be 50 % (0.50) but has been lowered to 30 % (0.30)
# based on user feedback.
NEAR_MISS_PCT: float = 0.30  # 30 % över behov

ORDER_SCHEMA: Dict[str, List[str]] = {
    "artikel": ["artikel", "artikelnummer", "sku", "article", "artnr", "art.nr"],
    "qty":     ["beställt", "antal", "qty", "quantity", "bestalld", "order qty"],
    "status":  ["status", "radstatus", "orderstatus", "state"],
    "ordid":   ["ordernr", "order nr", "order number", "kund", "kundnr"],
    "radid":   ["radnr", "rad nr", "line id", "rad", "struktur", "radsnr"],
}
BUFFER_SCHEMA: Dict[str, List[str]] = {
    "artikel": ["artikel", "article", "artnr", "art.nr", "artikelnummer"],
    "qty":     ["antal", "qty", "quantity", "pallantal", "colli", "units"],
    "loc":     ["lagerplats", "plats", "location", "bin", "hyllplats"],
    "dt":      ["datum/tid", "datum", "mottagen", "received", "inleverans", "inleveransdatum", "timestamp", "arrival"],
    "id":      ["pallid", "pall id", "id", "sscc", "etikett", "batch", "lpn"],
    "status":  ["status", "pallstatus", "state"],
}

# Frivillig fil: "Ej inlagrade artiklar" – mappning
NOT_PUTAWAY_SCHEMA: Dict[str, List[str]] = {
    "artikel":  ["artikel", "artnr", "art.nr", "artikelnummer"],
    "namn":     ["artikelnamn", "artikelbenämning", "benämning", "produktnamn", "namn", "artikel.1"],
    "antal":    ["antal", "qty", "quantity", "kolli"],
    "status":   ["status"],
    "pallnr":   ["pall nr", "pallid", "pall id", "pall"],
    "sscc":     ["sscc"],
    "andrad":   ["ändrad", "senast ändrad", "timestamp"],
    "utgang":   ["utgång", "bäst före", "utgångsdatum", "utgangsdatum", "best före"],
}

# Nytt schema: "Saldo inkl. automation" – mappning
SALDO_SCHEMA: Dict[str, List[str]] = {
    "artikel":    ["artikel", "artnr", "art.nr", "artikelnummer", "sku", "article"],
    "plocksaldo": ["plocksaldo", "plock saldo", "plock-saldo", "saldo", "pick saldo", "pick qty",
                   "tillgängligt plock", "tillgangligt plock", "available pick", "plock"],
    "plockplats": ["plockplats", "huvudplock", "mainpick", "hyllplats", "bin", "location", "lagerplats"],
}

# Schema för item option-fil: artikel och staplingsbar-flagga. Kolumnnamnet i filen kan vara
# "Staplingsbar" eller "Ej staplingsbar" beroende på filversion. Inkludera vanliga varianter.
ITEM_SCHEMA: Dict[str, List[str]] = {
    "artikel": ORDER_SCHEMA["artikel"],  # återanvänd artikel-kandidater från beställningar
    # Kandidater för staplingsbar/ej staplingsbar
    "staplingsbar": [
        "staplingsbar", "staplings bar", "staplbar", "stackable",
        "ej staplingsbar", "ejstaplingsbar", "ej_staplingsbar", "non stackable"
    ]
}

# --- Hjälpare -------------------------------------------------------------

def _open_df_in_excel(df, label: str = "data") -> str:
    """Skriv DF (eller {blad: DF}) till temporär fil och öppna i OS:et."""
    import importlib
    if isinstance(df, dict):
        engine = None
        if importlib.util.find_spec("openpyxl"):
            engine = "openpyxl"
        elif importlib.util.find_spec("xlsxwriter"):
            engine = "xlsxwriter"
        else:
            raise RuntimeError("Saknar Excel-skrivare (installera 'openpyxl' eller 'xlsxwriter').")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        path = tmp.name; tmp.close()
        with pd.ExcelWriter(path, engine=engine) as writer:
            for sheet, d in df.items():
                dd = d if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
                dd.to_excel(writer, sheet_name=str(sheet)[:31] or "Sheet1", index=False)
    else:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.csv")
        path = tmp.name; tmp.close()
        (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)).to_csv(path, index=False, encoding="utf-8-sig")
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass
    return path

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ta bort BOM/whitespace i kolumnnamn för robustare kolumnmatchning."""
    try:
        df.rename(columns=lambda c: str(c).replace("\ufeff", "").strip(), inplace=True)
    except Exception:
        pass
    return df

def smart_to_datetime(s) -> pd.Series:
    """Robust datumtolkning (ISO→dayfirst=False, annars True; fallback tvärtom)."""
    try:
        ser = pd.Series(s) if not isinstance(s, pd.Series) else s
        vals = ser.dropna().astype(str).str.strip()
        sample = vals.head(50)
        # Försök YYYYMMDD först
        numeric_like = (sample.str.match(r"^\d{8}$").sum() >= max(1, int(len(sample) * 0.6)))
        if numeric_like:
            dt = pd.to_datetime(ser, format="%Y%m%d", errors="coerce")
            if not dt.isna().all():
                return dt
        iso_like = (sample.str.match(r"^\d{4}-\d{2}-\d{2}").sum() >= max(1, int(len(sample) * 0.6)))
        primary_dayfirst = False if iso_like else True
        dt = pd.to_datetime(ser, errors="coerce", dayfirst=primary_dayfirst)
        if hasattr(dt, "isna") and getattr(dt, "isna")().all():
            dt = pd.to_datetime(ser, errors="coerce", dayfirst=not primary_dayfirst)
        return dt
    except Exception:
        try: return pd.to_datetime(s, errors="coerce", dayfirst=True)
        except Exception: return pd.to_datetime(s, errors="coerce", dayfirst=False)

def to_num(x) -> float:
    if pd.isna(x): return 0.0
    s = str(x).replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else 0.0

def find_col(df: pd.DataFrame, candidates: List[str], required: bool = True, default=None) -> str:
    """Hitta en kolumn genom exakt eller substring-match mot kandidatnamn (case-insensitive)."""
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols: return cols[cand.lower()]
    for key, orig in cols.items():
        for cand in candidates:
            if cand.lower() in key: return orig
    if required and default is None:
        raise KeyError(f"Hittar inte kolumnerna {candidates} i {list(df.columns)}")
    return default

def logprintln(txt_widget: tk.Text, msg: str) -> None:
    txt_widget.configure(state="normal")
    txt_widget.insert("end", msg + "\n")
    txt_widget.see("end")
    txt_widget.configure(state="disabled")
    txt_widget.update()

def _first_path_from_dnd(event_data: str) -> str:
    raw = str(event_data).strip()
    if raw.startswith("{") and raw.endswith("}"): raw = raw[1:-1]
    if raw.startswith('"') and raw.endswith('"'): raw = raw[1:-1]
    return raw

# -------- "Ej inlagrade" (mappning) --------------------------------------

def _read_not_putaway_csv(path: str) -> pd.DataFrame:
    """Läs CSV för 'Ej inlagrade'. Försök auto-separator, fallback TAB."""
    try:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        if df.shape[1] == 1 and len(df):
            first = str(df.iloc[0, 0])
            if "\t" in first:
                df = pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
        return _clean_columns(df)
    except Exception:
        return _clean_columns(pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig"))

def normalize_not_putaway(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Mappa 'Ej inlagrade' till enkel struktur. Ingen påverkan på allokering/refill."""
    df = df_raw.copy()
    def col(key: str, required: bool, default=None) -> str:
        return find_col(df, NOT_PUTAWAY_SCHEMA[key], required=required, default=default)
    art_col  = col("artikel", True)
    name_col = col("namn", False, default=None)
    qty_col  = col("antal", True)
    st_col   = col("status", False, default=None)
    pall_col = col("pallnr", False, default=None)
    sscc_col = col("sscc", False, default=None)
    chg_col  = col("andrad", False, default=None)
    exp_col  = col("utgang", False, default=None)
    out = pd.DataFrame({
        "Artikel": df[art_col].astype(str).str.strip(),
        "Namn":    df[name_col].astype(str).str.strip() if name_col else "",
        "Antal":   df[qty_col].map(to_num).astype(float),
        "Status":  pd.to_numeric(df[st_col], errors="coerce") if st_col else pd.Series([np.nan]*len(df)),
        "Pall nr": df[pall_col].astype(str) if pall_col else "",
        "SSCC":    df[sscc_col].astype(str) if sscc_col else "",
        "Ändrad":  smart_to_datetime(df[chg_col]) if chg_col else pd.NaT,
        "Utgång":  smart_to_datetime(df[exp_col]) if exp_col else pd.NaT,
    })
    for c in ["Namn","Pall nr","SSCC"]:
        if c in out.columns: out[c] = out[c].fillna("").astype(str).str.strip()
    return out

# -------- "Saldo inkl. automation" (mappning) -----------------------------

def normalize_saldo(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Mappa saldofil till struktur per artikel: Plocksaldo (sum) + Plockplats (första icke-tom)."""
    df = _clean_columns(df_raw.copy())
    def col(key: str, required: bool, default=None) -> str:
        return find_col(df, SALDO_SCHEMA[key], required=required, default=default)
    art_col   = col("artikel", True)
    saldo_col = col("plocksaldo", False, default=None)
    plats_col = col("plockplats", False, default=None)

    if saldo_col is None:
        return pd.DataFrame(columns=["Artikel", "Plocksaldo", "Plockplats"])

    out = pd.DataFrame({
        "Artikel": df[art_col].astype(str).str.strip(),
        "Plocksaldo": pd.to_numeric(df[saldo_col].map(to_num), errors="coerce").fillna(0.0),
        "Plockplats": (df[plats_col].astype(str).str.strip() if plats_col else pd.Series([""]*len(df))),
    })
    agg = (out.groupby("Artikel", as_index=False)
              .agg({"Plocksaldo":"sum","Plockplats":lambda s: next((x for x in s if isinstance(x,str) and x.strip()), "")}))
    return agg

# ====================== SALES (inbakat) ======================

# Robust kolumnmatchning för plocklogg
PICK_LOG_SCHEMA: dict[str, list[str]] = {
    "artikel": ["artikel", "artikelnr", "artnr", "art.nr", "artikelnummer", "sku", "article"],
    "antal":   ["plockat", "antal", "quantity", "qty", "picked", "units"],
    "datum":   ["datum", "datumtid", "timestamp", "date", "tid", "time"]
}

def normalize_pick_log(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisera plocklogg.
    Ut: Artikelnummer[str], Artikel[str] (namn eller =Artikelnummer om saknas),
        Plockat[float≥0], Datum[datetime].
    """
    df = _clean_columns(df_raw.copy())

    art_col = find_col(df, PICK_LOG_SCHEMA["artikel"], required=True)
    qty_col = find_col(df, PICK_LOG_SCHEMA["antal"], required=True)
    dt_col  = find_col(df, PICK_LOG_SCHEMA["datum"], required=True)

    # Försök hitta ett namn/benämning
    name_col = None
    for cand in ["artikelnamn","namn","benämning","artikelbenämning","produktnamn"]:
        try:
            nc = find_col(df, [cand], required=False, default=None)
            if nc:
                name_col = nc
                break
        except KeyError:
            pass

    out = pd.DataFrame({
        "Artikelnummer": df[art_col].astype(str).str.strip(),
        "Plockat": pd.to_numeric(df[qty_col].map(to_num), errors="coerce").fillna(0.0).astype(float),
        "Datum": smart_to_datetime(df[dt_col])
    })

    if name_col:
        out["Artikel"] = df[name_col].astype(str).str.strip()
    else:
        out["Artikel"] = out["Artikelnummer"]

    return out

def compute_sales_metrics(df_norm: pd.DataFrame, today=None) -> pd.DataFrame:
    """
    Beräkna sales-mått per Artikelnummer.
    Kolumner:
      - Artikelnummer, Artikel
      - Total_7, Total_30, Total_90
      - ADV_30 (=Total_30/30), ADV_90 (=Total_90/90)
      - SenastPlockad, DagarSedanSenast
      - UnikaPlockdagar_90 (unika datum med Plockat>0 sista 90)
      - NollraderPerPlockdag_90 (medel antal rader med Plockat=0 per aktiv plockdag sista 90)
      - ABC_klass (Pareto på Total_90; 80/15/5 → A/B/C)
    """
    if df_norm is None or df_norm.empty:
        cols = [
            "Artikelnummer","Artikel","Total_7","Total_30","Total_90","ADV_30","ADV_90",
            "SenastPlockad","DagarSedanSenast","UnikaPlockdagar_90","NollraderPerPlockdag_90","ABC_klass"
        ]
        return pd.DataFrame(columns=cols)

    if today is None:
        today = pd.Timestamp.now().normalize()
    else:
        today = pd.to_datetime(today).normalize()

    df = df_norm.copy()
    df["DatumNorm"] = pd.to_datetime(df["Datum"]).dt.normalize()
    df["Plockat"] = pd.to_numeric(df["Plockat"], errors="coerce").fillna(0.0)

    # Fönsterfilter
    mask7  = df["DatumNorm"] >= (today - pd.Timedelta(days=7))
    mask30 = df["DatumNorm"] >= (today - pd.Timedelta(days=30))
    mask90 = df["DatumNorm"] >= (today - pd.Timedelta(days=90))

    # Summeringar per Artikelnummer
    total7  = df.loc[mask7].groupby("Artikelnummer")["Plockat"].sum()
    total30 = df.loc[mask30].groupby("Artikelnummer")["Plockat"].sum()
    total90 = df.loc[mask90].groupby("Artikelnummer")["Plockat"].sum()

    # Senaste plock (>0)
    positive = df[df["Plockat"] > 0]
    last_pick = positive.groupby("Artikelnummer")["DatumNorm"].max() if not positive.empty else pd.Series(dtype="datetime64[ns]")
    last_pick = last_pick.reindex(df["Artikelnummer"].unique())

    # Dagar sedan senaste
    days_since = (today - last_pick).dt.days
    days_since = days_since.where(~days_since.isna(), other=pd.NA)

    # Unika plockdagar under 90 (Plockat>0)
    sub90_pos = df.loc[mask90 & (df["Plockat"] > 0)]
    unique_days_90 = sub90_pos.groupby("Artikelnummer")["DatumNorm"].nunique()

    # Nollrader per plockdag sista 90
    sub90 = df.loc[mask90].copy()
    zero_rows = (sub90.assign(IsZero=(sub90["Plockat"]==0))
                        .groupby(["Artikelnummer","DatumNorm"])["IsZero"].sum()
                        .rename("ZeroRows"))
    zero_avg = zero_rows.reset_index().groupby("Artikelnummer")["ZeroRows"].mean()
    zero_avg = zero_avg.reindex(df["Artikelnummer"].unique()).fillna(0.0)

    # Bygg metrics
    idx = pd.Index(sorted(df["Artikelnummer"].astype(str).unique()), name="Artikelnummer")
    out = pd.DataFrame(index=idx)
    out["Total_7"]  = total7.reindex(idx).fillna(0).round().astype(int)
    out["Total_30"] = total30.reindex(idx).fillna(0).round().astype(int)
    out["Total_90"] = total90.reindex(idx).fillna(0).round().astype(int)
    out["ADV_30"] = (out["Total_30"] / 30.0).astype(float)
    out["ADV_90"] = (out["Total_90"] / 90.0).astype(float)
    out["SenastPlockad"] = last_pick.reindex(idx)
    out["DagarSedanSenast"] = days_since.reindex(idx)
    out["UnikaPlockdagar_90"] = unique_days_90.reindex(idx).fillna(0).astype(int)
    out["NollraderPerPlockdag_90"] = zero_avg.reindex(idx).fillna(0.0).astype(float)

    # ABC-klass baserat på Total_90 (Pareto 80/15/5)
    tmp = out["Total_90"].astype(float).sort_values(ascending=False)
    total_sum = float(tmp.sum())
    if total_sum <= 0:
        out["ABC_klass"] = "C"
    else:
        cum = tmp.cumsum() / total_sum
        cls = pd.Series(index=tmp.index, dtype=object)
        cls[cum <= 0.80] = "A"
        cls[(cum > 0.80) & (cum <= 0.95)] = "B"
        cls[cum > 0.95] = "C"
        out["ABC_klass"] = cls.reindex(out.index).fillna("C")

    out = out.reset_index()

    # Lägg på Artikelnamn om det finns i df_norm
    if "Artikel" in df_norm.columns:
        out = out.merge(df_norm[["Artikelnummer","Artikel"]].drop_duplicates(),
                        on="Artikelnummer", how="left")
    else:
        out["Artikel"] = out["Artikelnummer"]

    # Lägg om kolumnordningen
    cols = ["Artikelnummer","Artikel"] + [c for c in out.columns if c not in ["Artikelnummer","Artikel"]]
    out = out[cols]

    return out


def _open_sales_excel(df_or_dict, label: str = "sales") -> str:
    """Skriv DF eller {blad: DF} till temporär Excel/CSV och öppna (med säkra bladnamn)."""
    import importlib

    def _sanitize_sheet_name(name: str) -> str:
        s = str(name)
        for ch in ['\\', '/', '?', '*', ':', '[', ']']:
            s = s.replace(ch, '-')
        s = s.strip("'")  # ledande/avslutande apostrof ställer också till det
        if not s:
            s = "Sheet"
        # Max 31 tecken i Excel
        return s[:31]

    def _dedupe(name: str, used: set[str]) -> str:
        base = name
        n = 2
        out = name
        while out in used:
            suffix = f" ({n})"
            out = (base[:31 - len(suffix)] + suffix)
            n += 1
        used.add(out)
        return out

    if isinstance(df_or_dict, dict):
        engine = None
        if importlib.util.find_spec("openpyxl"):
            engine = "openpyxl"
        elif importlib.util.find_spec("xlsxwriter"):
            engine = "xlsxwriter"
        else:
            raise RuntimeError("Saknar Excel-skrivare (installera 'openpyxl' eller 'xlsxwriter').")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        path = tmp.name; tmp.close()
        used_names: set[str] = set()
        with pd.ExcelWriter(path, engine=engine) as writer:
            for sheet, d in df_or_dict.items():
                safe = _sanitize_sheet_name(sheet)
                safe = _dedupe(safe, used_names)
                dd = d if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
                dd.to_excel(writer, sheet_name=safe, index=False)
    else:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.csv")
        path = tmp.name; tmp.close()
        (df_or_dict if isinstance(df_or_dict, pd.DataFrame) else pd.DataFrame(df_or_dict)).to_csv(path, index=False, encoding="utf-8-sig")

    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass
    return path

def open_sales_insights(df_metrics: pd.DataFrame) -> str:
    """
    Skapar Excel med:
      - Top sellers (90d)
      - Slow movers (≥90d eller 0)
      - Sammanställning
    Inkluderar alltid kolumnen Artikel (artikelnummer).
    """
    if df_metrics is None or df_metrics.empty:
        raise RuntimeError("Inga försäljningsinsikter att visa (tom metrics).")

    # Se till att Artikel finns först
    cols = ["Artikel"] + [c for c in df_metrics.columns if c != "Artikel"]
    df = df_metrics[cols].copy()

    top = df.sort_values(["Total_90","ADV_90"], ascending=[False, False]).reset_index(drop=True)
    slow = df[(df["DagarSedanSenast"].fillna(10**9) >= 90) | (df["Total_90"] == 0)] \
              .sort_values(["DagarSedanSenast","Total_90"], ascending=[False, True]) \
              .reset_index(drop=True)

    sheets = {
        "Top sellers (90d)": top,
        "Slow movers (≥90d eller 0)": slow,
        "Sammanställning": df
    }
    return _open_sales_excel(sheets, label="sales_insights")

def annotate_refill(refill_df: pd.DataFrame, df_metrics: pd.DataFrame) -> pd.DataFrame:
    """
    Lägg på sales-kolumner i refill-blad (påverkar inte logiken). Returnerar nytt DF.
    Adderar: ADV_90, ABC_klass, DagarSedanSenast, UnikaPlockdagar_90, NollraderPerPlockdag_90
    """
    if refill_df is None or refill_df.empty or df_metrics is None or df_metrics.empty:
        return refill_df
    cols = ["Artikel", "ADV_90", "ABC_klass", "DagarSedanSenast", "UnikaPlockdagar_90", "NollraderPerPlockdag_90"]
    cols = [c for c in cols if c in df_metrics.columns or c == "Artikel"]
    out = refill_df.merge(df_metrics[cols], on="Artikel", how="left")
    return out

# --- Item normalisering -------------------------------------------------

def normalize_items(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisera item-fil för att extrahera artikelnummer och staplingsbar-flagga.
    Returnerar DataFrame med kolumner ["Artikel", "Staplingsbar"].

    Parametrar:
        df_raw: O-normaliserad DataFrame från item-CSV.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["Artikel", "Staplingsbar"])
    df = df_raw.copy()
    # Städa kolumner från BOM och whitespace
    df = _clean_columns(df)
    try:
        art_col = find_col(df, ITEM_SCHEMA["artikel"], required=True)
    except Exception:
        art_col = None
    try:
        stap_col = find_col(df, ITEM_SCHEMA["staplingsbar"], required=False, default=None)
    except Exception:
        stap_col = None
    if not art_col:
        return pd.DataFrame(columns=["Artikel", "Staplingsbar"])
    if not stap_col or stap_col not in df.columns:
        # Om Staplingsbar inte finns, returnera bara artikel-kolumn med tom flagga
        tmp = df[[art_col]].copy()
        tmp.columns = ["Artikel"]
        # Skapa en kolumn för Ej Staplingsbar med tom sträng
        tmp["Ej Staplingsbar"] = ""
        # Ta bort dubbletter, behåll första
        return tmp.drop_duplicates(subset=["Artikel"]).reset_index(drop=True)
    # Extrahera och normalisera
    tmp = df[[art_col, stap_col]].copy()
    # Döp om kolumnerna till standardnamn
    tmp.columns = ["Artikel", "Ej Staplingsbar"]
    # Städa upp: trimma strängar och ersätt NaN med tomt
    tmp["Artikel"] = tmp["Artikel"].astype(str).str.strip()
    tmp["Ej Staplingsbar"] = tmp["Ej Staplingsbar"].fillna("").astype(str).str.strip()
    # Ta bort dubbletter, behåll första förekomsten
    return tmp.drop_duplicates(subset=["Artikel"]).reset_index(drop=True)

# --- Pallplatser (kundsammanställning) ---------------------------------

def compute_pallet_spaces(result_df: pd.DataFrame) -> pd.DataFrame:
    """
    Beräkna pallplatsbehov per kund baserat på allokeringsresultatet.

    Parametrar:
        result_df: DataFrame med allokerade orderrader efter saldofil-omklassificering och item/ej staplingsbar-sammanfogning.

    Returnerar:
        Ett DataFrame med kolumnerna ["Kund", "Kund1", "Botten Pallar", "Topp Pallar", "Totalt Pallar", "Pallplatser"].
        Om nödvändiga kolumner saknas returneras ett tomt DataFrame.
    """
    if result_df is None or result_df.empty:
        return pd.DataFrame(columns=["Kund", "Kund1", "Botten Pallar", "Topp Pallar", "Totalt Pallar", "Pallplatser"])
    df = result_df.copy()
    # Identifiera kolumner
    try:
        kund_col = find_col(df, ["kund", "customer"], required=True)
    except Exception:
        return pd.DataFrame(columns=["Kund", "Kund1", "Botten Pallar", "Topp Pallar", "Totalt Pallar", "Pallplatser"])
    try:
        kund1_col = find_col(df, ["kund1", "kund 1", "customer1", "kund.1"], required=False, default=None)
    except Exception:
        kund1_col = None
    zone_col = "Zon (beräknad)" if "Zon (beräknad)" in df.columns else None
    stack_col = None
    try:
        stack_col = find_col(df, ["ej staplingsbar", "ejstaplingsbar", "staplingsbar", "staplings bar"], required=False, default=None)
    except Exception:
        stack_col = None
    palltyp_col = "Palltyp (matchad)" if "Palltyp (matchad)" in df.columns else None
    if zone_col is None or palltyp_col is None:
        return pd.DataFrame(columns=["Kund", "Kund1", "Botten Pallar", "Topp Pallar", "Totalt Pallar", "Pallplatser"])

    # Förbereda kolumner
    df[zone_col] = df[zone_col].fillna("").astype(str).str.strip().str.upper()
    if stack_col:
        df[stack_col] = df[stack_col].fillna("").astype(str).str.strip().str.upper()
    else:
        df["_stack_tmp"] = ""
        stack_col = "_stack_tmp"
    df[palltyp_col] = df[palltyp_col].fillna("").astype(str).str.strip().str.upper()

    # Gruppberäkning
    groups = df.groupby([kund_col] if kund1_col is None else [kund_col, kund1_col])
    records: list[dict] = []
    import math
    for keys, sub in groups:
        # unpack keys
        if kund1_col is None:
            kund_val = keys
            kund1_val = ""
        else:
            kund_val, kund1_val = keys
        # Bottenpallar: zon H och Ej Staplingsbar = N eller tomt
        # Om kolumnen är tom (""), räknas den som staplingsbar=N.
        mask_bottom = (sub[zone_col] == "H") & ((sub[stack_col] == "N") | (sub[stack_col] == ""))
        B = int(mask_bottom.sum())
        # Toppallar A
        rows_A = int((sub[zone_col] == "A").sum())
        # En pall rymmer 20 rader. Om det finns några rader ska minst 1 pall avsättas, så runda uppåt.
        if rows_A > 0:
            top_A = math.ceil(rows_A / 20.0)
        else:
            top_A = 0
        # Toppallar H: ej staplingsbar = Y och palltyp != SJÖ
        mask_topH = (sub[zone_col] == "H") & (sub[stack_col] == "Y") & (sub[palltyp_col] != "SJÖ")
        top_H = int(mask_topH.sum())
        # Toppallar R: baserat på antal rader
        rows_R = int((sub[zone_col] == "R").sum())
        if rows_R == 0:
            top_R = 0
        elif rows_R <= 96:
            top_R = 1
        elif rows_R <= 163:
            top_R = 2
        elif rows_R <= 204:
            top_R = 3
        else:
            top_R = 4
        # Toppallar S: baserat på antal rader
        rows_S = int((sub[zone_col] == "S").sum())
        if rows_S == 0:
            top_S = 0
        elif rows_S <= 10:
            top_S = 1
        elif rows_S <= 15:
            top_S = 2
        elif rows_S <= 20:
            top_S = 3
        elif rows_S <= 26:
            top_S = 4
        else:
            top_S = 5
        # Sjörader (H-zon med palltyp SJÖ)
        mask_sjo = (sub[zone_col] == "H") & (sub[palltyp_col] == "SJÖ")
        S_rows = int(mask_sjo.sum())
        # Totala topplar (kan vara float)
        T = top_A + top_H + top_R + top_S
        # Beräkna pallplatser
        half_sum = (B + T) / 2.0
        P_component = math.ceil(half_sum)
        # Välj max mellan T och avrundade halvsumman
        max_val = T if T > P_component else P_component
        P = max_val + 2 * S_rows
        # Totalt pallar (sjöpallar räknas som 1 pall per rad)
        total_pallar = B + T + S_rows
        # Bygg resultatrad med extra rubriker för helpall/hellpall, sjöpall, skrymme, plockpall och autostore
        # helpall stapelbar = antal rader i H-zon med Ej Staplingsbar = N eller tomt (B)
        helpall_stapelbar = B
        # helpall ej stapelbar = antal rader i H-zon med Ej Staplingsbar = Y och palltyp != SJÖ
        helpall_ej_stapelbar = top_H
        # Sjö pall = antal rader i H-zon där palltyp = SJÖ
        sjo_pall = S_rows
        # Skrymme = antal pallar från S-zon (skrymmande) enligt toppal-logik
        skrymme_pallar = top_S
        # Plockpall = antal pallar från A-zon (huvudplock), en pall per 20 rader (avrundat uppåt)
        plockpall = top_A
        # autostore = antal pallar från R-zon (AUTOSTORE) enligt toppal-logik
        autostore_pallar = top_R
        # Totalt pallar har samma beräkning som tidigare (B + T + S_rows)
        record = {
            "Kund": kund_val,
            "Kund1": kund1_val,
            "hellpall stapelbar": helpall_stapelbar,
            "hellpall ej stapelbar": helpall_ej_stapelbar,
            "Sjö pall": sjo_pall,
            "Skrymme": skrymme_pallar,
            "Plockpall": plockpall,
            "autostore": autostore_pallar,
            "Botten Pallar": B,
            "Topp Pallar": T,
            "Totalt Pallar": total_pallar,
            "Pallplatser": P
        }
        records.append(record)
    return pd.DataFrame(records)

# == Prognosrapport (prognos vs autoplock) ==

def _safe_str_series(s: pd.Series) -> pd.Series:
    """
    Returnera en strängserie där varje värde är trimmat och NaN ersätts med tom sträng.
    """
    return s.astype(str).fillna("").str.strip()


def _str_to_num(x) -> float:
    """
    Extrahera första numeriska värdet ur ett godtyckligt objekt/sträng och returnera som float.
    Saknas numeriskt värde → 0.0.
    """
    import re
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    s = str(x).replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else 0.0


def _num_series(s: pd.Series) -> pd.Series:
    """
    Konvertera en Serie till numeriska värden med hjälp av _str_to_num. NaN ersätts med 0.
    """
    return pd.to_numeric(s.map(_str_to_num), errors="coerce").fillna(0)


def _sum_not_putaway(not_putaway_df: Optional[pd.DataFrame]) -> pd.Series:
    """
    Summera kolumnen 'Antal' per artikel i en normaliserad ej-inlagrade-DataFrame.
    Returnerar en Series med artikelnummer som index och summa antal som värde.
    Om underlaget saknas eller fel format returneras en tom Series.
    """
    if not isinstance(not_putaway_df, pd.DataFrame) or not len(not_putaway_df):
        return pd.Series(dtype=float)
    df = not_putaway_df.copy()
    if "Artikel" not in df.columns or "Antal" not in df.columns:
        return pd.Series(dtype=float)
    df["Artikel"] = _safe_str_series(df["Artikel"])
    df["Antal"] = _num_series(df["Antal"])
    return df.groupby("Artikel")["Antal"].sum()


def _collect_exclude_source_ids(allocated_df: Optional[pd.DataFrame]) -> set[str]:
    """
    Samla ihop de käll-ID:n från en allokerad DataFrame som motsvarar HELPALL-rader.
    Dessa ID används för att exkludera källor i refill/FIFO-beräkningen.
    """
    exclude: set[str] = set()
    if isinstance(allocated_df, pd.DataFrame) and not allocated_df.empty:
        if "Källtyp" in allocated_df.columns and "Källa" in allocated_df.columns:
            mask = _safe_str_series(allocated_df["Källtyp"]) == "HELPALL"
            vals = _safe_str_series(allocated_df.loc[mask, "Källa"]).replace("", pd.NA).dropna().unique().tolist()
            exclude = set(vals)
    return exclude


def _fifo_pallar_for_article(buffer_df: Optional[pd.DataFrame], article: str, needed_units: float, exclude_source_ids: Optional[set[str]] = None) -> float:
    """
    FIFO-baserad beräkning för hur många pallar som behövs för att täcka 'needed_units' av en given artikel.
    Filtrerar bufferten enligt REFILL_BUFFER_STATUSES och exkluderar angivna käll-ID.
    Returnerar ett flyttal med antalet pallar (heltal). Om inget behövs → 0. Om underlag saknas → NaN.
    """
    # Om inget behov finns
    if needed_units <= 0:
        return 0.0
    if not isinstance(buffer_df, pd.DataFrame) or buffer_df.empty:
        return np.nan
    df = buffer_df.copy()
    # Rensa BOM och trimma kolumnnamn
    try:
        df.rename(columns=lambda c: str(c).replace("\ufeff", "").strip(), inplace=True)
    except Exception:
        pass
    # Identifiera kolumner
    try:
        art_col = find_col(df, BUFFER_SCHEMA["artikel"], required=True)
        qty_col = find_col(df, BUFFER_SCHEMA["qty"], required=True)
        dt_col = find_col(df, BUFFER_SCHEMA["dt"], required=False, default=None)
        status_col = find_col(df, BUFFER_SCHEMA["status"], required=False, default=None)
        id_col = find_col(df, BUFFER_SCHEMA["id"], required=False, default=None)
    except Exception:
        return np.nan
    # Filtrera på artikel
    sub = df.loc[_safe_str_series(df[art_col]) == str(article)].copy()
    if sub.empty:
        return 0.0
    # Filtrera status enligt REFILL_BUFFER_STATUSES (kan vara sträng eller numerisk)
    if status_col and status_col in sub.columns:
        s = _safe_str_series(sub[status_col])
        s_num = pd.to_numeric(s.str.extract(r"(-?\d+)")[0], errors="coerce")
        allowed_str = {str(x) for x in REFILL_BUFFER_STATUSES}
        sub = sub[s.isin(allowed_str) | s_num.isin(REFILL_BUFFER_STATUSES)].copy()
        if sub.empty:
            return 0.0
    # Exkludera käll-ID:n om angivet
    if exclude_source_ids:
        if id_col and id_col in sub.columns:
            sub["_source_id"] = _safe_str_series(sub[id_col])
        else:
            sub["_source_id"] = "SRC-" + sub.index.astype(str)
        sub = sub[~sub["_source_id"].isin(exclude_source_ids)].copy()
        if sub.empty:
            return 0.0
    # Konvertera kvantitet till numeriskt
    sub["__qty__"] = _num_series(sub[qty_col])
    # FIFO-sortering efter datum om det finns
    if dt_col and dt_col in sub.columns:
        sub = sub.sort_values(dt_col, kind="mergesort", na_position="last")
    # Summera pallar tills behovet täcks
    acc = 0.0
    pall_count = 0
    for q in sub["__qty__"]:
        if q <= 0:
            continue
        acc += float(q)
        pall_count += 1
        if acc >= float(needed_units):
            break
    if pall_count == 0:
        return 0.0
    return float(pall_count)


def build_prognos_vs_autoplock_report(
    prognos_df: pd.DataFrame,
    saldo_norm_df: Optional[pd.DataFrame] = None,
    buffer_df: Optional[pd.DataFrame] = None,
    *,
    exclude_source_ids: Optional[set[str]] = None,
    allocated_df: Optional[pd.DataFrame] = None,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Bygg en rapport som jämför prognosens behov med saldo i autoplock och buffertpallar (FIFO‑baserad
    pallberäkning). Kolumnen för ej inlagrade artiklar (E) har tagits bort.
    Returnerar ett DataFrame med kolumnerna A–D samt F och en meta‑dikt som anger om rapporten är
    partiell och eventuella notes om vad som saknas.
    """
    meta: Dict[str, str] = {"partial": "no", "missing": "", "note": ""}
    missing: List[str] = []
    # Om prognos saknas → tomt resultat
    if not isinstance(prognos_df, pd.DataFrame) or prognos_df.empty:
        empty = pd.DataFrame(columns=[
            "Artikelnummer",
            "Behov i prognosen (antal styck)",
            "Saldo i autoplock",
            "Behov efter saldo",
            "Summa antal i ej inlagrade artiklar",
            "FIFO-baserad beräkning (antal pall)",
        ])
        meta.update({"partial": "yes", "missing": "prognos", "note": "Ingen prognos inläst."})
        return empty, meta
    pr = prognos_df.copy()
    # Säkerställ kolumnnamn Artikelnummer och Antal styck
    if "Artikelnummer" not in pr.columns or "Antal styck" not in pr.columns:
        rename_map: Dict[str, str] = {}
        for col in pr.columns:
            lc = str(col).strip().lower()
            if lc in ("product code", "artikelnummer", "artnr", "sku", "article"):
                rename_map[col] = "Artikelnummer"
            elif lc in ("antal styck", "antal", "qty", "quantity"):
                rename_map[col] = "Antal styck"
        if rename_map:
            pr = pr.rename(columns=rename_map)
    pr["Artikelnummer"] = _safe_str_series(pr.get("Artikelnummer", ""))
    pr["Antal styck"] = _num_series(pr.get("Antal styck", 0))
    # --- Saldomappning ---
    if isinstance(saldo_norm_df, pd.DataFrame) and not saldo_norm_df.empty:
        # Kontrollera om saldodatat verkligen innehåller autoplock-fält (Robot + Saldo autoplock).
        # Om inte, behandla saldo som saknat underlag (så att vi inte filtrerar bort allt).
        orig_cols = [str(c).strip().lower() for c in saldo_norm_df.columns]
        has_robot_col = any("robot" == c for c in orig_cols)
        has_auto_col = any("saldo autoplock" in c for c in orig_cols)
        # Om Robot‑kolumnen saknas betraktas saldo som helt saknat underlag. I så fall filtrerar vi inte på robot.
        if not has_robot_col:
            missing.append("saldo")
            pr["Robot"] = "N"
            pr["Saldo i autoplock"] = 0.0
        else:
            # Saldodatat innehåller åtminstone Robot‑kolumnen. Vi kan filtrera prognosen på Robot=Y.
            s = saldo_norm_df.copy()
            # Mappa artikelkolumnen om nödvändigt
            if "Artikel" not in s.columns:
                for c in s.columns:
                    lc = str(c).strip().lower()
                    if lc in ("artikel", "artikelnummer", "sku", "artnr", "art.nr", "article"):
                        s = s.rename(columns={c: "Artikel"})
                        break
            # Se till att Robot-kolumnen finns
            if "Robot" not in s.columns:
                s["Robot"] = "N"
            # Om saldokolumn saknas sätter vi 0. Detta påverkar bara kolumnen "Saldo i autoplock" (C)
            if "Saldo autoplock" not in s.columns:
                s["Saldo autoplock"] = 0.0
            s["Artikel"] = _safe_str_series(s["Artikel"])
            s["Robot"] = _safe_str_series(s["Robot"]).str.upper().map(lambda x: "Y" if x == "Y" else "N")
            s["Saldo autoplock"] = _num_series(s["Saldo autoplock"])
            pr = pr.merge(s[["Artikel", "Robot", "Saldo autoplock"]], left_on="Artikelnummer", right_on="Artikel", how="left")
            pr = pr.drop(columns=["Artikel"], errors="ignore")
            pr["Robot"].fillna("N", inplace=True)
            pr["Saldo i autoplock"] = pr["Saldo autoplock"].fillna(0.0)
    else:
        missing.append("saldo")
        pr["Robot"] = "N"
        pr["Saldo i autoplock"] = 0.0
    # Kolumner B, C, D
    pr["Behov i prognosen (antal styck)"] = _num_series(pr["Antal styck"])
    pr["Saldo i autoplock"] = _num_series(pr["Saldo i autoplock"])
    pr["Behov efter saldo"] = (pr["Behov i prognosen (antal styck)"] - pr["Saldo i autoplock"]).clip(lower=0)
    # E: ej inlagrade har tagits bort. Sätt kolumnen till 0 (bevaras enbart för bakåtkompatibilitet)
    pr["Summa antal i ej inlagrade artiklar"] = 0.0
    # F: FIFO-baserad pallberäkning. Kvarvarande behov = behov efter saldo (ingen avdrag för ej inlagrade)
    shortage = pr["Behov efter saldo"].copy()
    # Hämta exclude_source_ids från allocated_df om ej angivet
    if exclude_source_ids is None and isinstance(allocated_df, pd.DataFrame):
        exclude_source_ids = _collect_exclude_source_ids(allocated_df)
    if not exclude_source_ids:
        exclude_source_ids = None
    if isinstance(buffer_df, pd.DataFrame) and not buffer_df.empty:
        pr["FIFO-baserad beräkning (antal pall)"] = [
            _fifo_pallar_for_article(buffer_df, art, need, exclude_source_ids=exclude_source_ids)
            for art, need in zip(pr["Artikelnummer"], shortage)
        ]
    else:
        missing.append("buffert")
        pr["FIFO-baserad beräkning (antal pall)"] = np.nan
    # Om saldo finns filtrera Robot=Y och behov efter saldo > 0
    if "saldo" not in missing:
        pr = pr.loc[(pr["Robot"] == "Y") & (pr["Behov efter saldo"] > 0)].copy()
    # Slutlig kolumnordning
    # Slutlig kolumnordning: utan 'Summa antal i ej inlagrade artiklar'
    out_cols = [
        "Artikelnummer",
        "Behov i prognosen (antal styck)",
        "Saldo i autoplock",
        "Behov efter saldo",
        "FIFO-baserad beräkning (antal pall)",
    ]
    for c in out_cols:
        if c not in pr.columns:
            pr[c] = np.nan if c.startswith("FIFO") else 0.0
    report = pr[out_cols].reset_index(drop=True)
    # Meta
    if missing:
        notes: List[str] = []
        if "saldo" in missing:
            notes.append("Saldo saknas → visar ej Robot-filter eller Saldo i autoplock (C=0, D=B).")
        if "buffert" in missing:
            notes.append("Buffert saknas → F kan inte beräknas.")
        meta = {
            "partial": "yes",
            "missing": ",".join(sorted(set(missing))),
            "note": " ".join(notes),
        }
    else:
        meta = {"partial": "no", "missing": "", "note": ""}
    return report, meta


def open_prognos_vs_autoplock_excel(report_df: pd.DataFrame, meta: Optional[dict] = None) -> str:
    """
    Skriv en prognosrapport (A–F) till en temporär Excel-fil och öppna den. Om meta anger att
    rapporten är partiell eller innehåller anteckningar skapas även ett Info-blad.
    Returnerar sökvägen till den skapade filen.
    """
    sheets: dict[str, pd.DataFrame] = {}
    # Info-blad vid partiell rapport
    if isinstance(meta, dict) and (meta.get("partial") == "yes" or meta.get("note")):
        lines: list[str] = []
        if meta.get("partial") == "yes":
            missing = meta.get("missing", "")
            lines.append("PARTIELL RAPPORT – mer data krävs för fullständig bild.")
            if missing:
                lines.append(f"Saknar underlag: {missing}.")
        if meta.get("note"):
            lines.append(str(meta["note"]))
        if lines:
            sheets["Info"] = pd.DataFrame({"Info": [" ".join(lines)]})
    # Huvudbladet
    if not isinstance(report_df, pd.DataFrame):
        report_df = pd.DataFrame()
    sheets["Prognos vs Autoplock"] = report_df
    return _open_df_in_excel(sheets, label="prognos_vs_autoplock")

# =================== SLUT SALES (inbakat) ===================


# --- Kärnlogik: allokering -----------------------------------------------

def allocate(orders_raw: pd.DataFrame, buffer_raw: pd.DataFrame, log=None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Allokera beställningsrader mot buffert enligt HELPALL→AUTOSTORE→HUVUDPLOCK.
    - Buffert filter: status {29,30,32} + platsfilter (ej AA*, TRANSIT, TRANSIT_ERROR, MISSING, UT2).
    - Ignorera orderrader med Status=35.
    Returnerar (allocated_df, near_miss_df).
    """
    def _log(msg: str):
        if log:
            log(msg)

    # Kolumnupptäckt via schema
    order_article_col = find_col(orders_raw, ORDER_SCHEMA["artikel"])
    order_qty_col     = find_col(orders_raw, ORDER_SCHEMA["qty"])
    order_id_col      = find_col(orders_raw, ORDER_SCHEMA["ordid"], required=False, default=None)
    order_line_col    = find_col(orders_raw, ORDER_SCHEMA["radid"], required=False, default=None)
    order_status_col  = find_col(orders_raw, ORDER_SCHEMA["status"], required=False, default=None)

    buff_article_col  = find_col(buffer_raw, BUFFER_SCHEMA["artikel"])
    buff_qty_col      = find_col(buffer_raw, BUFFER_SCHEMA["qty"])
    buff_loc_col      = find_col(buffer_raw, BUFFER_SCHEMA["loc"])
    buff_dt_col       = find_col(buffer_raw, BUFFER_SCHEMA["dt"], required=False, default=None)
    buff_id_col       = find_col(buffer_raw, BUFFER_SCHEMA["id"], required=False, default=None)
    buff_status_col   = find_col(buffer_raw, BUFFER_SCHEMA["status"], required=False, default=None)
    # Försök hitta en kolumn för palltyp i bufferten (t.ex. "Palltyp").
    # Innehåller typ av pall (t.ex. EUR, EIG, SJÖ eller BIN-x) som vi vill föra vidare till resultatet.
    try:
        # Hitta kolumn för palltyp. Vi matchar på "palltyp" och vanliga variationer.
        buff_type_col = find_col(buffer_raw, [
            "palltyp", "pall typ", "pall type"
        ], required=False, default=None)
    except Exception:
        buff_type_col = None

    _log(f"Order-kolumner: Artikel='{order_article_col}', Antal='{order_qty_col}', OrderId='{order_id_col}', Rad='{order_line_col}', Status='{order_status_col}'")
    _log(f"Buffert-kolumner: Artikel='{buff_article_col}', Antal='{buff_qty_col}', Lagerplats='{buff_loc_col}', Tid='{buff_dt_col}', ID='{buff_id_col}', Status='{buff_status_col}'")

    # Normalisera orders
    orders = orders_raw.copy()
    orders["_artikel"] = orders[order_article_col].astype(str).str.strip()
    orders["_qty"] = orders[order_qty_col].map(to_num).astype(float)
    orders["_order_id"] = orders[order_id_col].astype(str) if order_id_col and order_id_col in orders.columns else ""
    orders["_order_line"] = orders[order_line_col].astype(str) if order_line_col and order_line_col in orders.columns else orders.index.astype(str)

    # Ignorera Status=35
    if order_status_col and order_status_col in orders.columns:
        _status_str = orders[order_status_col].astype(str).str.strip()
        _status_num = pd.to_numeric(_status_str.str.extract(r"(-?\d+)")[0], errors="coerce")
        _before = len(orders)
        orders = orders[~(_status_num == 35)].copy()
        _removed = _before - len(orders)
        if _removed:
            _log(f"Ignorerar {_removed} orderrad(er) pga Status = 35.")
    else:
        _log("OBS: Ingen order-statuskolumn hittad; kan inte filtrera Status = 35.")

    # Normalisera buffert
    buffer_df = buffer_raw.copy()
    buffer_df["_artikel"] = buffer_df[buff_article_col].astype(str).str.strip()
    buffer_df["_qty"] = buffer_df[buff_qty_col].map(to_num).astype(float)
    buffer_df["_loc"] = buffer_df[buff_loc_col].astype(str).str.strip()
    buffer_df["_received"] = smart_to_datetime(buffer_df[buff_dt_col]) if buff_dt_col and buff_dt_col in buffer_df.columns else pd.NaT
    buffer_df["_source_id"] = buffer_df[buff_id_col].astype(str) if buff_id_col and buff_id_col in buffer_df.columns else "SRC-" + buffer_df.index.astype(str)
    # Kopiera palltyp (om kolumn hittades) till intern kolumn.
    # Fyll ut NaN och eventuella tomma/vita strängar med "" så att "nan" inte sprids till resultatet.
    if buff_type_col and buff_type_col in buffer_df.columns:
        # str() på NaN ger "nan", så ersätt detta i efterhand. fillna hanterar även None.
        tmp_palltyp = buffer_df[buff_type_col].fillna("").astype(str).str.strip()
        # Ersätt "nan" (som str(NaN) ger) med tom sträng
        buffer_df["_palltyp"] = tmp_palltyp.replace({"nan": "", "": ""})
    else:
        buffer_df["_palltyp"] = ""

    # Statusfilter (29/30/32)
    if buff_status_col and buff_status_col in buffer_df.columns:
        status_series = buffer_df[buff_status_col].astype(str).str.strip()
        status_num = pd.to_numeric(status_series.str.extract(r"(-?\d+)")[0], errors="coerce")
        allowed_str = {str(x) for x in ALLOC_BUFFER_STATUSES}
        mask_allowed = status_series.isin(allowed_str) | status_num.isin(ALLOC_BUFFER_STATUSES)
        removed = int((~mask_allowed).sum())
        if removed:
            _log(f"Filtrerar bort {removed} buffertpall(ar) pga Status ej i {sorted(ALLOC_BUFFER_STATUSES)}.")
        buffer_df = buffer_df[mask_allowed].copy()
    else:
        _log("OBS: Hittade ingen statuskolumn; ingen statusfiltrering tillämpas.")

    # Platsfilter
    loc_upper = buffer_df["_loc"].str.upper()
    mask_exclude = loc_upper.str.startswith(INVALID_LOC_PREFIXES, na=False) | loc_upper.isin(INVALID_LOC_EXACT)
    excluded_count = int(mask_exclude.sum())
    if excluded_count:
        _log(f"Filtrerar bort {excluded_count} rad(er) från bufferten pga lagerplats-regler ({INVALID_LOC_PREFIXES}*, {', '.join(sorted(INVALID_LOC_EXACT))}).")
    buffer_df = buffer_df[~mask_exclude].copy()

    # Säker, lätt minnesoptimering
    try:
        buffer_df["_artikel"] = buffer_df["_artikel"].astype("category")
    except Exception:
        pass

    buffer_df["_is_autostore"] = buffer_df["_loc"].str.contains("AUTOSTORE", case=False, na=False)
    buffer_df = buffer_df[buffer_df["_qty"] > 0].copy()

    far_future = pd.Timestamp("2262-04-11")
    buffer_df["_received_ord"] = buffer_df["_received"].fillna(far_future)

    pallets = buffer_df[~buffer_df["_is_autostore"]].copy().sort_values(by=["_artikel", "_received_ord", "_source_id"])
    bins = buffer_df[buffer_df["_is_autostore"]].copy().sort_values(by=["_artikel", "_received_ord", "_source_id"])

    pallet_queues: Dict[str, Deque[dict]] = defaultdict(deque)
    for _, r in pallets.iterrows():
        pallet_queues[str(r["_artikel"]).strip()].append({
            "source_id": r["_source_id"],
            "qty": float(r["_qty"]),
            "loc": r["_loc"],
            "received": r["_received"],
            # För intern pallsort: använd alltid sträng och ersätt NaN med tom sträng
            "palltyp": (r.get("_palltyp", "") if pd.notna(r.get("_palltyp", "")) else "")
        })

    bin_queues: Dict[str, Deque[dict]] = defaultdict(deque)
    for _, r in bins.iterrows():
        bin_queues[str(r["_artikel"]).strip()].append({
            "source_id": r["_source_id"],
            "qty": float(r["_qty"]),
            "loc": r["_loc"],
            "received": r["_received"],
            # För intern pallsort: använd alltid sträng och ersätt NaN med tom sträng
            "palltyp": (r.get("_palltyp", "") if pd.notna(r.get("_palltyp", "")) else "")
        })

    allocated_rows: List[dict] = []
    near_miss_rows: List[dict] = []
    # Track which articles have already been recorded for near-miss to avoid logging the same article many times
    near_miss_article_set: set[str] = set()

    def clone_row(orow: pd.Series) -> dict:
        return orow.to_dict()

    def record_near_miss(orow: pd.Series, pal: dict, need: float) -> None:
        """
        Record a near-miss event when a pallet is up to the configured NEAR_MISS_PCT larger than the
        remaining need for an article. To prevent excessive logging when the same article triggers
        multiple near-miss events across many order lines, this function will only record the first
        near-miss for each unique article. Additional near misses for the same article are ignored.
        """
        if need <= 0:
            return
        diff = pal["qty"] - need
        if diff <= 0:
            return
        pct = diff / need
        if pct <= NEAR_MISS_PCT:
            art_id = str(orow["_artikel"]).strip()
            # Only record the first near-miss event per article
            if art_id in near_miss_article_set:
                return
            near_miss_article_set.add(art_id)
            near_miss_rows.append({
                "Artikel": art_id,
                "OrderID": str(orow["_order_id"]),
                "OrderRad": str(orow["_order_line"]),
                "PallID": str(pal["source_id"]),
                "Källplats": str(pal["loc"]),
                "Mottagen": pal["received"],
                "Behov_vid_tillfället": need,
                "Pall_kvantitet": pal["qty"],
                "Skillnad": diff,
                "Procentuell skillnad (%)": pct * 100.0,
                # Beskrivning av anledningen använder den aktuella near‑miss‑gränsen i procent.
                "Anledning": f"Pallen var ≤{int(NEAR_MISS_PCT * 100)}% större än återstående behov (kan ej brytas)",
                # Placeholder for classification; will be updated later in run_allocation
                "Gäller (INSTEAD R/A)": None
            })

    for _, orow in orders.iterrows():
        art = str(orow["_artikel"]).strip()
        need = float(orow["_qty"])
        if need <= 0:
            continue

        # 1) HELPALL
        pq = pallet_queues.get(art, deque())
        new_pq = deque()
        tmp = deque(pq)
        any_helpall = False
        while tmp and need > 0:
            pal = tmp.popleft()
            pal_qty = pal["qty"]
            if pal_qty <= need:
                sub = clone_row(orow)
                sub[order_qty_col] = pal_qty
                sub["Zon (beräknad)"] = "H"
                sub["Källtyp"] = "HELPALL"
                sub["Källa"] = pal["source_id"]
                sub["Källplats"] = pal["loc"]
                # Lägg till matchad palltyp. Ersätt NaN/"nan"/None med tom sträng.
                paltyp_val = pal.get("palltyp", "")
                if not paltyp_val or str(paltyp_val).lower() == "nan":
                    paltyp_val = ""
                sub["Palltyp (matchad)"] = paltyp_val
                allocated_rows.append(sub)
                need -= pal_qty
                any_helpall = True
            else:
                record_near_miss(orow, pal, need)
                new_pq.append(pal)
        while tmp:
            new_pq.append(tmp.popleft())
        pallet_queues[art] = new_pq

        # 2) AUTOSTORE
        any_autostore = False
        bq = bin_queues.get(art, deque())
        new_bq = deque()
        while bq and need > 0:
            binr = bq.popleft()
            take = min(binr["qty"], need)
            if take > 0:
                sub = clone_row(orow)
                sub[order_qty_col] = take
                sub["Zon (beräknad)"] = "R"
                sub["Källtyp"] = "AUTOSTORE"
                sub["Källa"] = binr["source_id"]
                sub["Källplats"] = binr["loc"]
                # Lägg till matchad palltyp. Ersätt NaN/"nan"/None med tom sträng.
                bin_palltyp_val = binr.get("palltyp", "")
                if not bin_palltyp_val or str(bin_palltyp_val).lower() == "nan":
                    bin_palltyp_val = ""
                sub["Palltyp (matchad)"] = bin_palltyp_val
                allocated_rows.append(sub)
                binr["qty"] -= take
                need -= take
                any_autostore = True
            if binr["qty"] > 0:
                new_bq.append(binr)
        while bq:
            new_bq.append(bq.popleft())
        bin_queues[art] = new_bq

        # 3) HUVUDPLOCK
        any_mainpick = False
        if need > 0:
            sub = clone_row(orow)
            sub[order_qty_col] = need
            sub["Zon (beräknad)"] = "A"
            sub["Källtyp"] = "HUVUDPLOCK"
            sub["Källa"] = ""
            sub["Källplats"] = ""
            # Ingen palltyp vid huvudplock
            sub["Palltyp (matchad)"] = ""
            allocated_rows.append(sub)
            any_mainpick = True
            need = 0.0

        # Near-miss markering
        if not any_helpall and (any_autostore or any_mainpick):
            for r in near_miss_rows:
                if r["OrderID"] == str(orow["_order_id"]) and r["OrderRad"] == str(orow["_order_line"]):
                    r["Gäller (INSTEAD R/A)"] = True
        else:
            for r in near_miss_rows:
                if r["OrderID"] == str(orow["_order_id"]) and r["OrderRad"] == str(orow["_order_line"]):
                    r["Gäller (INSTEAD R/A)"] = False

    allocated_df = pd.DataFrame(allocated_rows)

    # Om en artikel har AUTOSTORE-rad → gör alla dess icke-HELPALL till AUTOSTORE
    try:
        if not allocated_df.empty and ("Källtyp" in allocated_df.columns):
            if "Zon (beräknad)" not in allocated_df.columns:
                allocated_df["Zon (beräknad)"] = ""
            low = {c.lower(): c for c in allocated_df.columns}
            art_col_res = None
            for n in ["artikel", "article", "artnr", "art.nr", "artikelnummer", "_artikel"]:
                if n.lower() in low:
                    art_col_res = low[n.lower()]
                    break
            if art_col_res:
                auto_arts = set(allocated_df.loc[allocated_df["Källtyp"].astype(str) == "AUTOSTORE", art_col_res].astype(str).str.strip())
                if auto_arts:
                    mask_same = allocated_df[art_col_res].astype(str).str.strip().isin(auto_arts)
                    mask_change = mask_same & (allocated_df["Källtyp"].astype(str) != "HELPALL")
                    allocated_df.loc[mask_change, "Källtyp"] = "AUTOSTORE"
                    allocated_df.loc[mask_change, "Zon (beräknad)"] = "R"
    except Exception:
        pass

    # Utökade kolumner som läggs till i resultatet
    added_cols = ["Zon (beräknad)", "Källtyp", "Källa", "Källplats", "Palltyp (matchad)"]
    ordered_cols = [c for c in orders_raw.columns] + [c for c in added_cols if c not in orders_raw.columns]
    if not allocated_df.empty:
        allocated_df = allocated_df[ordered_cols]
    else:
        allocated_df = pd.DataFrame(columns=ordered_cols)

    near_miss_df = pd.DataFrame(near_miss_rows)
    return allocated_df, near_miss_df

# --- Refill (HP: A+S, med saldoreduktion proportionerligt; 0-filter) ------

def calculate_refill(allocated_df: pd.DataFrame,
                     buffer_raw: pd.DataFrame,
                     saldo_df: pd.DataFrame | None = None,
                     not_putaway_df: pd.DataFrame | None = None
                     ) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Beräkna påfyllningspallar.
    - HP-blad inkluderar både HUVUDPLOCK (A) och SKRYMMANDE (S).
    - Plocksaldo dras en gång per artikel och fördelas proportionerligt mellan A och S.
    - 0-rader tas bort.
    - AUTOSTORE-blad (R) oförändrat, men 0-rader tas också bort.
    - Buffert filtreras till status {29,30}. HELPALL-pallar som redan används exkluderas alltid.
    """

    result = allocated_df.copy()
    buff = buffer_raw.copy()

    # Identifiera kolumner
    art_col_res = find_col(result, ORDER_SCHEMA["artikel"])
    qty_col_res = find_col(result, ORDER_SCHEMA["qty"])

    art_col_buf = find_col(buff, BUFFER_SCHEMA["artikel"])
    qty_col_buf = find_col(buff, BUFFER_SCHEMA["qty"])
    dt_col_buf  = find_col(buff, BUFFER_SCHEMA["dt"], required=False, default=None)
    id_col_buf  = find_col(buff, BUFFER_SCHEMA["id"], required=False, default=None)
    status_col_buf = find_col(buff, BUFFER_SCHEMA["status"], required=False, default=None)

    # Normalisera buffert
    b = buff.copy()
    b["_artikel"] = b[art_col_buf].astype(str).str.strip()
    b["_qty"] = b[qty_col_buf].map(to_num).astype(float)
    b["_received"] = smart_to_datetime(b[dt_col_buf]) if dt_col_buf and dt_col_buf in b.columns else pd.NaT
    b["_source_id"] = b[id_col_buf].astype(str) if id_col_buf and id_col_buf in b.columns else "SRC-" + b.index.astype(str)

    # Statusfilter refill
    if status_col_buf and status_col_buf in b.columns:
        _s = b[status_col_buf].astype(str).str.strip()
        _snum = pd.to_numeric(_s.str.extract(r"(-?\d+)")[0], errors="coerce")
        allowed_str = {str(x) for x in REFILL_BUFFER_STATUSES}
        b = b[_s.isin(allowed_str) | _snum.isin(REFILL_BUFFER_STATUSES)].copy()

    # HELPALL-pallar som redan använts (exkludera alltid)
    used_help_ids: set[str] = set()
    if "Källtyp" in result.columns and "Källa" in result.columns:
        used_help_ids = set(result[result["Källtyp"].astype(str) == "HELPALL"]["Källa"].dropna().astype(str).tolist())

    # Bygg plocksaldo/plockplats
    saldo_sum: Dict[str, float] = {}
    plockplats_by_art: Dict[str, str] = {}
    if isinstance(saldo_df, pd.DataFrame) and not saldo_df.empty:
        try:
            s_norm = normalize_saldo(saldo_df)
            for _, r in s_norm.iterrows():
                art = str(r["Artikel"]).strip()
                saldo_sum[art] = float(saldo_sum.get(art, 0.0) + float(r.get("Plocksaldo", 0.0)))
                pp = str(r.get("Plockplats", "") or "").strip()
                if pp and art not in plockplats_by_art:
                    plockplats_by_art[art] = pp
        except Exception:
            saldo_sum = {}
            plockplats_by_art = {}

    # Bygg "ej inlagrade" summor
    npu_sum: Dict[str, float] = {}
    if isinstance(not_putaway_df, pd.DataFrame) and not not_putaway_df.empty:
        try:
            npu = not_putaway_df.copy()
            npu_art_col = find_col(npu, NOT_PUTAWAY_SCHEMA["artikel"])
            npu_qty_col = find_col(npu, NOT_PUTAWAY_SCHEMA["antal"])
            grp = npu.groupby(npu[npu_art_col].astype(str).str.strip())[npu_qty_col].apply(lambda s: float(pd.to_numeric(s, errors="coerce").fillna(0).sum()))
            npu_sum = {str(k): float(v) for k, v in grp.to_dict().items()}
        except Exception:
            npu_sum = {}

    # FIFO helper per artikel
    def fifo_for_art(art_key: str) -> pd.DataFrame:
        d = b[b["_artikel"] == art_key].copy()
        if not d.empty and used_help_ids:
            d = d[~d["_source_id"].astype(str).isin(used_help_ids)].copy()
        return d.sort_values("_received")

    # --- HP (A + S) ---
    hp_like = result[result.get("Källtyp", "").isin(["HUVUDPLOCK", "SKRYMMANDE"])].copy()
    rows_hp: List[dict] = []
    if not hp_like.empty:
        # lägg zon-kod
        hp_like["_zon"] = np.where(hp_like["Källtyp"].astype(str) == "SKRYMMANDE", "S", "A")
        # behov per artikel och zon
        needs = (hp_like
                 .assign(_art=hp_like[art_col_res].astype(str).str.strip(),
                         _qty=pd.to_numeric(hp_like[qty_col_res], errors="coerce").fillna(0.0))
                 .groupby(["_art", "_zon"], as_index=False)["_qty"].sum())

        for art_key, grp_art in needs.groupby("_art"):
            total_need = float(grp_art["_qty"].sum())
            if total_need <= 0:
                continue
            # plocksaldo dras en gång
            adjusted_total = max(0.0, round(total_need) - float(saldo_sum.get(art_key, 0.0)))

            # proportionell fördelning mellan zoner
            if adjusted_total <= 0:
                continue  # 0-rad; hoppa över helt

            parts = []
            allocated_sum = 0
            for _, r in grp_art.iterrows():
                zone = str(r["_zon"])
                part = (float(r["_qty"]) / total_need) * adjusted_total if total_need > 0 else 0.0
                val = int(round(part))
                parts.append([zone, val])
                allocated_sum += val
            # justera avrundningsdiff på första posten
            diff = int(adjusted_total) - int(allocated_sum)
            if parts:
                parts[0][1] += diff

            # FIFO info för artikeln
            fifo_df = fifo_for_art(art_key)
            tillgangligt = float(pd.to_numeric(fifo_df["_qty"], errors="coerce").sum()) if not fifo_df.empty else 0.0

            for zone, behov_int in parts:
                behov_int = int(max(0, behov_int))
                if behov_int <= 0:
                    continue  # 0-rad → bort
                # simulera pallar
                behov_kvar = float(behov_int)
                pall_count = 0
                for q in (fifo_df["_qty"].astype(float) if not fifo_df.empty else []):
                    if behov_kvar <= 0:
                        break
                    pall_count += 1
                    behov_kvar -= float(q)

                rows_hp.append({
                    "Artikel": art_key,
                    "Zon": zone,  # A eller S
                    "Behov (kolli)": behov_int,
                    "FIFO-baserad beräkning": int(pall_count),
                    "Tillräckligt tillgängligt saldo i buffert": "Ja" if tillgangligt >= behov_int else "Nej",
                    "Plockplats": plockplats_by_art.get(art_key, ""),
                    "Ej inlagrade (antal)": int(round(npu_sum.get(art_key, 0.0)))
                })

    refill_hp_df = pd.DataFrame(rows_hp)
    if not refill_hp_df.empty:
        refill_hp_df = refill_hp_df.sort_values(["Zon", "FIFO-baserad beräkning"], ascending=[True, False])

    # --- AUTOSTORE (R) ---
    refill_autostore_df = pd.DataFrame()
    try:
        as_df = result.copy()
        if not as_df.empty:
            mask_autostore = as_df["Källtyp"].astype(str) == "AUTOSTORE" if "Källtyp" in as_df.columns else pd.Series(False, index=as_df.index)
            k_blank = as_df["Källa"].isna() | (as_df["Källa"].astype(str).str.strip() == "") if "Källa" in as_df.columns else pd.Series(True, index=as_df.index)
            as_df = as_df[mask_autostore & k_blank].copy()
        if not as_df.empty:
            art_col_res_as = find_col(as_df, ORDER_SCHEMA["artikel"])
            qty_col_res_as = find_col(as_df, ORDER_SCHEMA["qty"])
            behov_per_art_as = as_df.groupby(as_df[art_col_res_as].astype(str).str.strip())[qty_col_res_as] \
                                   .apply(lambda s: float(pd.to_numeric(s, errors="coerce").fillna(0).sum())) \
                                   .to_dict()

            rows_as: List[dict] = []
            for art, behov in behov_per_art_as.items():
                art_key = str(art).strip()
                fifo_df = fifo_for_art(art_key)
                tillgangligt = float(pd.to_numeric(fifo_df["_qty"], errors="coerce").sum()) if not fifo_df.empty else 0.0
                behov_int = int(max(0, round(behov) - float(saldo_sum.get(art_key, 0.0))))
                if behov_int <= 0:
                    continue  # 0-rad bort
                remaining = float(behov_int)
                pall_count = 0
                for q in (fifo_df["_qty"].astype(float) if not fifo_df.empty else []):
                    if remaining <= 0:
                        break
                    pall_count += 1
                    remaining -= float(q)

                rows_as.append({
                    "Artikel": art_key,
                    "Behov (kolli)": behov_int,
                    "FIFO-baserad beräkning": int(pall_count),
                    "Tillräckligt tillgängligt saldo i buffert": "Ja" if tillgangligt >= behov_int else "Nej",
                    "Plockplats": plockplats_by_art.get(art_key, ""),
                    "Ej inlagrade (antal)": int(round(npu_sum.get(art_key, 0.0)))
                })

            refill_autostore_df = pd.DataFrame(rows_as)
            if not refill_autostore_df.empty:
                refill_autostore_df = refill_autostore_df.sort_values("FIFO-baserad beräkning", ascending=False)
    except Exception:
        refill_autostore_df = pd.DataFrame()

    return refill_hp_df, refill_autostore_df

# --- GUI ------------------------------------------------------------------

class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.pack(fill="both", expand=True)
        self._create_widgets()
        # Interna variabler för kampanjvolymer
        self._campaign_norm: Optional[pd.DataFrame] = None
        self._campaign_raw: Optional[pd.DataFrame] = None

    def _log(self, msg: str, level: str = "info") -> None:
        logprintln(self.log, msg)

    def _create_widgets(self) -> None:
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="Beställningslinjer (CSV):").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.orders_var = tk.StringVar()
        self.orders_entry = ttk.Entry(self, textvariable=self.orders_var)
        self.orders_entry.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_orders).grid(row=0, column=2, padx=8)

        ttk.Label(self, text="Buffertpallar (CSV):").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.buffer_var = tk.StringVar()
        self.buffer_entry = ttk.Entry(self, textvariable=self.buffer_var)
        self.buffer_entry.grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_buffer).grid(row=1, column=2, padx=8)

        ttk.Label(self, text="Saldo inkl. automation (CSV):").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        self.automation_var = tk.StringVar()
        self.automation_entry = ttk.Entry(self, textvariable=self.automation_var)
        self.automation_entry.grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_automation).grid(row=2, column=2, padx=8)

        # Item option-fil för Ej Staplingsbar (frivillig)
        ttk.Label(self, text="Item option (CSV):").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.item_var = tk.StringVar()
        self.item_entry = ttk.Entry(self, textvariable=self.item_var)
        self.item_entry.grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_item).grid(row=3, column=2, padx=8)

        # Prognos (XLSX)
        ttk.Label(self, text="Prognos (XLSX):").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        self.prognos_var = tk.StringVar()
        self.prognos_entry = ttk.Entry(self, textvariable=self.prognos_var)
        self.prognos_entry.grid(row=4, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_prognos).grid(row=4, column=2, padx=8)

        # Kampanjvolymer (XLSX)
        ttk.Label(self, text="Kampanjvolymer (XLSX):").grid(row=5, column=0, sticky="w", padx=8, pady=6)
        self.campaign_var = tk.StringVar()
        self.campaign_entry = ttk.Entry(self, textvariable=self.campaign_var)
        self.campaign_entry.grid(row=5, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Bläddra...", command=self.pick_campaign).grid(row=5, column=2, padx=8)

        # Kör-knapp
        self.run_btn = ttk.Button(self, text="Kör allokering", command=self.run_allocation)
        self.run_btn.grid(row=6, column=0, columnspan=3, pady=10)

        # Öppna-knappar (för resultat, near-miss, pallplatsrapport och prognos) och rensa-cache.
        # Resultat
        self.open_result_btn = ttk.Button(self, text="Öppna allokerade pallar", command=self.open_result_in_excel, state="disabled")
        # Near-miss (INSTEAD R/A)
        self.open_nearmiss_btn = ttk.Button(self, text="Öppna near-miss", command=self.open_nearmiss_in_excel, state="disabled")
        # Pallplatser per kund
        self.open_palletspaces_btn = ttk.Button(self, text="Öppna pallplatser", command=self.open_pallet_spaces_in_excel, state="disabled")
        # Prognos
        self.open_prognos_btn = ttk.Button(self, text="Öppna prognos", command=self.open_prognos_in_excel, state="disabled")
        # Positionera knappar på samma rad. Lägg till extra kolumn för pallplatser och prognos.
        self.open_result_btn.grid(row=99, column=0, pady=10)
        self.open_nearmiss_btn.grid(row=99, column=1, pady=10)
        self.open_palletspaces_btn.grid(row=99, column=2, pady=10)
        self.open_prognos_btn.grid(row=99, column=3, pady=10)

        # Rensa cache-knapp: återställ alla internt beräknade data och töm loggen
        self.reset_cache_btn = ttk.Button(self, text="Rensa cache", command=self.reset_cache)
        # Placera den i nästa kolumn på samma rad som övriga öppna-knappar
        self.reset_cache_btn.grid(row=99, column=4, pady=10)

        # Logg
        ttk.Label(self, text="Logg / Summering:").grid(row=6, column=0, sticky="w", padx=8)
        self.log = tk.Text(self, height=14, width=110, state="disabled")
        self.log.grid(row=7, column=0, columnspan=4, sticky="nsew", padx=8, pady=8)
        self.rowconfigure(7, weight=1)

        # Summeringstabell (inkl. SKRYMMANDE)
        ttk.Label(self, text="Summering per Källtyp").grid(row=8, column=0, sticky="w", padx=8)
        self.summary_table = ttk.Treeview(self, columns=("ktyp", "antal_rader", "antal_kolli"), show="headings", height=5)
        self.summary_table.heading("ktyp", text="Källtyp")
        self.summary_table.heading("antal_rader", text="antal rader")
        self.summary_table.heading("antal_kolli", text="antal kolli")
        self.summary_table.column("ktyp", anchor="w", width=160)
        self.summary_table.column("antal_rader", anchor="e", width=140)
        self.summary_table.column("antal_kolli", anchor="e", width=140)
        self.summary_table.grid(row=9, column=0, columnspan=4, sticky="ew", padx=8, pady=(0,8))
        # Tabellrader sätts dynamiskt i update_summary_table() baserat på faktiska Källtyp-värden

        # Senaste resultat i minnet
        self.last_result_df: pd.DataFrame | None = None
        self.last_nearmiss_instead_df: pd.DataFrame | None = None
        self._orders_raw: pd.DataFrame | None = None
        self._buffer_raw: pd.DataFrame | None = None
        self._result_df: pd.DataFrame | None = None

        # Lagring: ej-inlagrade & saldo normaliserat (utan öppna-knapp)
        self._not_putaway_raw: pd.DataFrame | None = None
        self._not_putaway_norm: pd.DataFrame | None = None
        self._saldo_norm: pd.DataFrame | None = None

        # Rå saldodata (inkl. Robot-kolumn)
        self._saldo_raw: pd.DataFrame | None = None

        # Lagring: item-fil (staplingsbar)
        self._item_raw: pd.DataFrame | None = None
        self._item_norm: pd.DataFrame | None = None

        # Sales cache
        self._sales_metrics_df: pd.DataFrame | None = None

        # Auto-refill cache (beräknas i run_allocation)
        self._last_refill_hp_df: pd.DataFrame | None = None
        self._last_refill_autostore_df: pd.DataFrame | None = None

        # Pallplatser (kundsammanställning)
        self._pallet_spaces_df: pd.DataFrame | None = None

        # Prognos data (normaliserad)
        self._prognos_df: pd.DataFrame | None = None

        #
        # Inga cache-data att återställa här; rensningslogik implementeras i reset_cache()


        # DnD – bind en gång per entry
        if TkinterDnD and DND_FILES:
            def bind_drop(entry_widget: ttk.Entry, var: tk.StringVar) -> None:
                try:
                    entry_widget.drop_target_register(DND_FILES)
                    def _on_drop(event, _var=var):
                        path = _first_path_from_dnd(event.data)
                        if path:
                            _var.set(path)
                            # Om prognos, ladda prognosen direkt
                            if _var is self.prognos_var:
                                self._load_prognos(path)
                            # Om kampanjvolymer, ladda kampanjfilen direkt
                            elif _var is self.campaign_var:
                                self._load_campaign(path)
                    entry_widget.dnd_bind("<<Drop>>", _on_drop)
                except Exception:
                    pass
            bind_drop(self.orders_entry, self.orders_var)
            bind_drop(self.buffer_entry, self.buffer_var)
            bind_drop(self.automation_entry, self.automation_var)
            bind_drop(self.item_entry, self.item_var)
            # Prognos DnD
            bind_drop(self.prognos_entry, self.prognos_var)
            # Kampanjvolymer DnD
            if hasattr(self, "campaign_entry"):
                bind_drop(self.campaign_entry, self.campaign_var)
            # Fält för ej inlagrade artiklar och plocklogg har tagits bort.

    # --- File pickers ------------------------------------------------------

    def pick_orders(self) -> None:
        path = filedialog.askopenfilename(title="Välj beställningsrader (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path: self.orders_var.set(path)

    def pick_automation(self) -> None:
        path = filedialog.askopenfilename(title="Välj Saldo inkl. automation (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path: self.automation_var.set(path)

    def pick_buffer(self) -> None:
        path = filedialog.askopenfilename(title="Välj buffertpallar (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path: self.buffer_var.set(path)

    def pick_item(self) -> None:
        """
        Öppna dialog för att välja item-fil (CSV) med staplingsbar-uppgift.
        """
        path = filedialog.askopenfilename(title="Välj item-fil (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self.item_var.set(path)

    def pick_not_putaway(self) -> None:
        """
        Stub för filval av 'Ej inlagrade artiklar'. Denna funktion gör inget i denna version.
        """
        return

    def pick_sales(self) -> None:
        """
        Stub för filval av plocklogg. Denna funktion gör inget i denna version.
        """
        return

    # --- Öppna/Export -----------------------------------------------------

    def open_result_in_excel(self) -> None:
        if isinstance(self.last_result_df, pd.DataFrame) and not self.last_result_df.empty:
            try:
                # Spara resultatet i en temporär Excel-fil med eget bladnamn
                path = _open_df_in_excel({"Allokerade order": self.last_result_df.copy()}, label="allocated_orders")
                self._log(f"Öppnade resultat i Excel (temporär fil): {path}")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna resultat i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget resultat att öppna ännu. Kör allokeringen först.")

    def open_nearmiss_in_excel(self) -> None:
        if isinstance(self.last_nearmiss_instead_df, pd.DataFrame) and not self.last_nearmiss_instead_df.empty:
            try:
                # Spara near-miss i en temporär Excel-fil med eget bladnamn
                nm_df = self.last_nearmiss_instead_df.copy()
                # Visa endast den första near-miss för varje artikel
                if "Artikel" in nm_df.columns:
                    # Behåll första förekomsten per artikel
                    nm_df = nm_df.drop_duplicates(subset=["Artikel"], keep="first").reset_index(drop=True)
                # Använd ett bladnamn utan snedstreck ("/") eftersom Excel inte tillåter det
                # Bladnamnet och etiketten ska spegla den aktuella near‑miss‑gränsen
                pct_str = f"{int(NEAR_MISS_PCT * 100)}%"
                sheet_name = f"Near-miss {pct_str} (unika artiklar)"
                label = f"near_miss_{int(NEAR_MISS_PCT * 100)}pct"
                path = _open_df_in_excel({sheet_name: nm_df}, label=label)
                self._log(f"Öppnade near-miss (INSTEAD R or A) i Excel (temporär fil): {path}")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna near-miss i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns ingen near-miss INSTEAD R/A att öppna ännu.")

    def open_pallet_spaces_in_excel(self) -> None:
        """
        Öppna den beräknade pallplatsrapporten per kund i en temporär Excel-fil.
        Rapporten innehåller antal bottenpallar, toppallar, totalt pallar och pallplatser per kund.
        """
        if isinstance(self._pallet_spaces_df, pd.DataFrame) and not self._pallet_spaces_df.empty:
            try:
                # Spara pallplatsrapport i en temporär Excel-fil med eget bladnamn
                ps_df = self._pallet_spaces_df.copy()
                path = _open_df_in_excel({"Pallplatser": ps_df}, label="pallplatser")
                self._log(f"Öppnade pallplatser i Excel (temporär fil): {path}")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna pallplatser i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns ingen pallplatsrapport att öppna ännu. Kör allokeringen först.")

    # --- Prognos (inläsning och öppning) ---
    def pick_prognos(self) -> None:
        """Visa en filväljare för att välja en prognosfil (XLSX)."""
        path = filedialog.askopenfilename(title="Välj prognos (XLSX)", filetypes=[("Excel", "*.xlsx"), ("Alla filer","*.*")])
        if path:
            self.prognos_var.set(path)
            self._load_prognos(path)
        else:
            # Om inget valdes, nollställ prognosen och inaktivera knappen
            self._prognos_df = None
            self.open_prognos_btn.configure(state="disabled")

    def _load_prognos(self, path: str) -> None:
        """Läs in prognosfilen och aktivera knappen för öppning."""
        try:
            df = read_prognos_xlsx(path)
            self._prognos_df = df
            # Logga antal rader/artiklar
            try:
                n_art = int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else len(df)
                self._log(f"Prognos inläst: {len(df)} rader, {n_art} artiklar.")
            except Exception:
                self._log(f"Prognos inläst: {len(df)} rader.")
            # Aktivera öppna-knappen
            self.open_prognos_btn.configure(state="normal")
        except Exception as e:
            self._prognos_df = None
            self.open_prognos_btn.configure(state="disabled")
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa prognosfilen:\n{e}")

    def pick_campaign(self) -> None:
        """Visa en filväljare för att välja en kampanjvolymfil (XLSX)."""
        path = filedialog.askopenfilename(title="Välj kampanjvolymer (XLSX)", filetypes=[("Excel", "*.xlsx"), ("Alla filer", "*.*")])
        if path:
            self.campaign_var.set(path)
            self._load_campaign(path)
        else:
            # Om inget valdes, nollställ kampanjdata
            self._campaign_norm = None

    def _load_campaign(self, path: str) -> None:
        """Läs in kampanjvolymer och lagra den normaliserade datan."""
        try:
            df = read_campaign_xlsx(path)
            self._campaign_norm = df
            # Logga antal kampanjrader/artiklar
            try:
                n_art = int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else len(df)
                self._log(f"Kampanjvolymer inlästa: {len(df)} rader, {n_art} artiklar.")
            except Exception:
                self._log(f"Kampanjvolymer inlästa: {len(df)} rader.")
            # Aktivera prognos-knappen om vi har minst en av prognos eller kampanjvolymer
            try:
                if (self._prognos_df is not None and isinstance(self._prognos_df, pd.DataFrame) and not self._prognos_df.empty) or (isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty):
                    self.open_prognos_btn.configure(state="normal")
            except Exception:
                pass
        except Exception as e:
            self._campaign_norm = None
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa kampanjfilen:\n{e}")

    def open_prognos_in_excel(self) -> None:
        """
        Skapa och öppna en prognosrapport i en temporär Excel‑fil.

        Rapporten jämför prognosbehovet med saldo i autoplock, ej inlagrade artiklar samt buffertpallar
        (FIFO‑logik) och följer exakt samma uträkningar som i originalprojektet. Om prognosen inte
        har lästs in ännu visas ett meddelande istället.
        """
        # Kontrollera att det finns underlag: antingen prognos eller kampanjvolymer
        has_prognos = isinstance(self._prognos_df, pd.DataFrame) and not self._prognos_df.empty
        has_campaign = isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty
        if not has_prognos and not has_campaign:
            messagebox.showinfo(APP_TITLE, "Välj och läs in antingen prognosfilen eller kampanjvolymerna först.")
            return
        try:
            # Utgå från inläst prognos eller skapa ett tomt DataFrame
            if has_prognos:
                combined_df: pd.DataFrame = self._prognos_df.copy()
            else:
                # Om ingen prognos finns, utgå från kampanjvolymer och skapa motsvarande kolumner
                # Skapa en DataFrame med samma format som prognos_df (Artikelnummer + Beskrivning + Antal styck + Antal rader + Antal butiker)
                combined_df = pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
            # Om en normaliserad kampanjvolym finns och saldodata finns kan vi slå ihop kampanjvolymerna
            if isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty:
                # Filtrera kampanjvolymer på Robot=Y om saldo finns och har Robot-kolumn
                camp_df = self._campaign_norm.copy()
                if isinstance(self._saldo_raw, pd.DataFrame) and not self._saldo_raw.empty:
                    s = self._saldo_raw.copy()
                    # Identifiera artikel- och robotkolumn i saldofilen
                    art_col_sal = None
                    robot_col_sal = None
                    for c in s.columns:
                        lc = str(c).strip().lower()
                        if not art_col_sal and lc in ("artikel", "artikelnummer", "artnr", "art.nr", "sku", "article"):
                            art_col_sal = c
                        if not robot_col_sal and lc == "robot":
                            robot_col_sal = c
                    if art_col_sal and robot_col_sal:
                        s = s[[art_col_sal, robot_col_sal]].copy()
                        s.columns = ["Artikelnummer", "Robot"]
                        s["Artikelnummer"] = s["Artikelnummer"].astype(str).str.strip()
                        s["Robot"] = s["Robot"].astype(str).str.upper().str.strip()
                        # Endast Robot = Y
                        s = s.loc[s["Robot"] == "Y"]
                        if not s.empty:
                            camp_df = camp_df.merge(s[["Artikelnummer"]], on="Artikelnummer", how="inner")
                        else:
                            camp_df = camp_df.iloc[0:0]
                    else:
                        # Om robot eller artikelkolumn saknas, ta bort alla kampanjrader (ingen match)
                        camp_df = camp_df.iloc[0:0]
                # Om camp_df fortfarande har rader, summera volym per artikel
                if not camp_df.empty:
                    # Summera per artikel
                    vol_by_art = camp_df.groupby("Artikelnummer")["Antal styck"].sum().to_dict()
                    # Lägg samman med ordinarie prognos
                    # Konvertera Artikelnummer i prognosen
                    combined_df["Artikelnummer"] = combined_df["Artikelnummer"].astype(str).str.strip()
                    combined_df["Antal styck"] = pd.to_numeric(combined_df.get("Antal styck", 0), errors="coerce").fillna(0).astype(int)
                    # Håll koll på vilka artiklar som finns
                    existing_arts = set(combined_df["Artikelnummer"].astype(str))
                    # Addera kampanjvolymer till befintliga rader
                    for art, vol in vol_by_art.items():
                        if art in existing_arts:
                            mask = combined_df["Artikelnummer"] == art
                            combined_df.loc[mask, "Antal styck"] = (combined_df.loc[mask, "Antal styck"].astype(int) + int(vol)).astype(int)
                        else:
                            # Lägg till en ny rad för artiklar som inte finns i prognosen
                            combined_df = pd.concat([
                                combined_df,
                                pd.DataFrame({
                                    "Artikelnummer": [art],
                                    "Beskrivning": [None],
                                    "Antal styck": [int(vol)],
                                    "Antal rader": [0],
                                    "Antal butiker": [0],
                                })
                            ], ignore_index=True)
            # Beräkna prognosrapporten baserat på kombinerad prognos, saldo, buffert
            report_df, meta = build_prognos_vs_autoplock_report(
                prognos_df=combined_df,
                saldo_norm_df=(self._saldo_raw if isinstance(self._saldo_raw, pd.DataFrame) else None),
                buffer_df=(self._buffer_raw if isinstance(self._buffer_raw, pd.DataFrame) else None),
                exclude_source_ids=None,
                allocated_df=None,
            )
            # Skapa och öppna Excel-filen med rapporten och eventuellt ett Info-blad
            path = open_prognos_vs_autoplock_excel(report_df, meta)
            # Logga en kort rapport i loggrutan
            msg = f"Prognosrapport skapad ({len(report_df)} rader)."
            if isinstance(meta, dict) and meta.get("partial") == "yes":
                miss = meta.get("missing", "").replace(",", ", ")
                if miss:
                    msg += f" PARTIELL: saknar {miss}."
            self._log(msg)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte skapa/öppna prognosrapporten:\n{e}")

    # --- Cache-reset ------------------------------------------------------
    def reset_cache(self) -> None:
        """
        Rensa alla cacher och temporära variabler i applikationen. Detta nollställer
        internt lagrade DataFrames (resultat, near-miss, saldo, item, sales m.m.),
        tömmer loggrutan, återställer summeringstabellen till noll och inaktiverar
        öppna-knapparna. Pathvariabler för filval påverkas inte.
        """
        try:
            # Töm internt lagrade resultat
            self.last_result_df = None
            self.last_nearmiss_instead_df = None
            self._orders_raw = None
            self._buffer_raw = None
            self._result_df = None
            self._not_putaway_raw = None
            self._not_putaway_norm = None
            self._saldo_norm = None
            self._saldo_raw = None
            self._item_raw = None
            self._item_norm = None
            self._sales_metrics_df = None
            self._last_refill_hp_df = None
            self._last_refill_autostore_df = None
            self._pallet_spaces_df = None
            self._prognos_df = None
            self._campaign_raw = None
            self._campaign_norm = None

            # Töm loggen
            self.log.configure(state="normal")
            self.log.delete("1.0", tk.END)
            self.log.configure(state="disabled")

            # Återställ summeringstabellen: ta bort alla rader
            try:
                for child in self.summary_table.get_children(""):
                    self.summary_table.delete(child)
            except Exception:
                pass

            # Inaktivera öppna-knapparna
            for btn in (self.open_result_btn, self.open_nearmiss_btn, self.open_palletspaces_btn, self.open_prognos_btn):
                try:
                    btn.configure(state="disabled")
                except Exception:
                    pass

            # Töm även filvägarna i gränssnittet så att användaren ser att alla val är återställda
            try:
                self.orders_var.set("")
                self.buffer_var.set("")
                self.automation_var.set("")
                self.item_var.set("")
                self.prognos_var.set("")
                self.campaign_var.set("")
            except Exception:
                # Om någon av variablerna saknas av någon anledning, ignorera
                pass

            # Logga att cache rensades
            self._log("Cache och temporära data har rensats.")
        except Exception:
            # Om något går fel i reset-cache, visa ändå en loggrad
            try:
                self._log("Kunde inte genomföra fullständig cache-rensning (internt fel).")
            except Exception:
                pass

    def open_refill_in_excel(self) -> None:
        """Öppnar den senast auto-beräknade refill-rapporten; annoterar med sales vid öppning om tillgängligt."""
        if isinstance(self._last_refill_hp_df, pd.DataFrame) or isinstance(self._last_refill_autostore_df, pd.DataFrame):
            try:
                hp = self._last_refill_hp_df.copy() if isinstance(self._last_refill_hp_df, pd.DataFrame) else pd.DataFrame()
                asr = self._last_refill_autostore_df.copy() if isinstance(self._last_refill_autostore_df, pd.DataFrame) else pd.DataFrame()
                # Annotera med sales om finns
                if isinstance(self._sales_metrics_df, pd.DataFrame) and not self._sales_metrics_df.empty:
                    hp = annotate_refill(hp, self._sales_metrics_df)
                    asr = annotate_refill(asr, self._sales_metrics_df)
                path = _open_df_in_excel({"Refill HP": hp, "Refill AUTOSTORE": asr}, label="refill")
                self._log(f"Öppnade påfyllningspallar (cache) i Excel (temporär fil): {path}")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna påfyllningspallar i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns ingen påfyllningspallsrapport att öppna ännu. Kör allokeringen först.")

    def open_sales_in_excel(self) -> None:
        if isinstance(self._sales_metrics_df, pd.DataFrame) and not self._sales_metrics_df.empty:
            try:
                path = open_sales_insights(self._sales_metrics_df)
                self._log(f"Öppnade försäljningsinsikter i Excel (temporär fil): {path}")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna försäljningsinsikter:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inga försäljningsinsikter att öppna ännu. Läs in en plocklogg först.")

    # --- Sales inläsning ---------------------------------------------------

    def _on_sales_file_selected(self) -> None:
        """
        Stub för hantering av plocklogg. Funktionen för att läsa in plockloggar och beräkna försäljningsinsikter är borttagen i denna version.

        Denna metod finns kvar för kompatibilitet men gör inget längre.
        """
        return

    # --- Summering ---------------------------------------------------------

    def update_summary_table(self, result_df: pd.DataFrame) -> None:
        """
        Uppdatera sammanställningstabellen med alla förekommande Källtyp‑värden.

        HELPALL visas som antal pallar, AUTOSTORE som antal rader, och övriga typer som
        antal rader samt motsvarande pallantal (20 rader per pall).
        """
        # Rensa befintliga rader i sammanställningstabellen
        for child in self.summary_table.get_children(""):
            self.summary_table.delete(child)
        # Hämta kolumn för kvantitet (antal) om den finns
        try:
            qty_col = find_col(result_df, ORDER_SCHEMA["qty"], required=False, default=None)
        except Exception:
            qty_col = None
        # Samtliga unika Källtyp-värden (strängar), sortera med HELPALL och AUTOSTORE först
        ktyp_series = result_df.get("Källtyp", pd.Series([], dtype=object)).astype(str)
        unique_types = [k for k in sorted(set(ktyp_series.dropna())) if k]
        # Prioriterad ordning
        ordered = []
        for prv in ("HELPALL", "AUTOSTORE"):
            if prv in unique_types:
                ordered.append(prv)
                unique_types.remove(prv)
        ordered.extend(unique_types)
        # Uppdatera tabellen för varje Källtyp
        for ktyp in ordered:
            try:
                sub = result_df[ktyp_series == ktyp]
                row_count = int(len(sub))
                kolli = 0.0
                if qty_col and not sub.empty:
                    kolli = float(pd.to_numeric(sub[qty_col], errors="coerce").sum())
            except Exception:
                row_count, kolli = 0, 0.0
            # Bestäm text för radantal och pallantal
            if ktyp == "HELPALL":
                row_text = f"{row_count} pallar"
            elif ktyp == "AUTOSTORE":
                row_text = f"{row_count} rader"
            else:
                pallar = (row_count / 20.0) if row_count else 0.0
                pallar_str = f"{pallar:.2f}".replace(".", ",")
                row_text = f"{row_count} rader ({pallar_str} pallar)"
            kolli_text = f"{int(round(kolli))}"
            self.summary_table.insert("", "end", iid=ktyp, values=(ktyp, row_text, kolli_text))

# --- Zonbaserad omklassificering --------------------------------------

    @staticmethod
    def _reclassify_skrymmande(result_df: pd.DataFrame, saldo_norm: pd.DataFrame | None) -> pd.DataFrame:
        """
        Omklassificera rader utifrån orderfilens zonkod.

        Efter att HELPALL‑ och AUTOSTORE‑allokeringar är bestämda (dvs. Källtyp
        "HELPALL" respektive "AUTOSTORE"), sätts Källtyp och "Zon (beräknad)"
        för övriga rader baserat på den befintliga "Zon"‑kolumnen i
        beställningsfilen. Följande mappning används (zon → (källtyp, zon)):

          * "S" → ("SKRYMMANDE",   "S")
          * "E" → ("EHANDEL",      "E")
          * "A" → ("HUVUDPLOCK",   "A")
          * "Q" → ("EHANDEL",      "Q")
          * "O" → ("SKRYMMANDE",   "O")
          * "F" → ("BRAND",        "F")

        Rader vars Källtyp redan är "HELPALL" eller "AUTOSTORE" lämnas
        oförändrade. Om ingen "Zon"‑kolumn hittas returneras oförändrat DataFrame.
        Den medskickade saldofil används inte i denna metod.
        """
        # Inga omklassificeringar om result_df saknas eller är tomt
        if result_df is None or result_df.empty:
            return result_df
        res = result_df.copy()
        # Hitta zonkolumnen (case-insensitive match mot exakt "zon")
        zon_col = None
        for c in res.columns:
            if str(c).strip().lower() == "zon":
                zon_col = c
                break
        if not zon_col:
            return res
        # Se till att kolumnen "Zon (beräknad)" finns
        if "Zon (beräknad)" not in res.columns:
            res["Zon (beräknad)"] = ""
        # Identifiera vilka rader som kan omklassificeras (inte HELPALL/AUTOSTORE)
        ktyp_series = res.get("Källtyp", pd.Series("", index=res.index)).astype(str)
        mask_to_change = ~(ktyp_series.isin(["HELPALL", "AUTOSTORE"]))
        if not mask_to_change.any():
            return res
        # Skapa mappning från zonkod till (källtyp, zon)
        mapping: Dict[str, Tuple[str, str]] = {
            "S": ("SKRYMMANDE",   "S"),
            "E": ("EHANDEL",      "E"),
            "A": ("HUVUDPLOCK",   "A"),
            "Q": ("EHANDEL",      "Q"),
            "O": ("SKRYMMANDE",   "O"),
            "F": ("BRAND",        "F"),
        }
        # Standardisera zonkod till stor bokstav
        zones = res.loc[mask_to_change, zon_col].astype(str).str.strip().str.upper()
        for zone_code, (ktyp_val, zon_val) in mapping.items():
            idx = res.loc[mask_to_change].index[zones == zone_code]
            if len(idx) > 0:
                res.loc[idx, "Källtyp"] = ktyp_val
                res.loc[idx, "Zon (beräknad)"] = zon_val
        return res

    # --- Orkestrering ------------------------------------------------------

    def run_allocation(self) -> None:
        orders_path = self.orders_var.get().strip()
        buffer_path = self.buffer_var.get().strip()
        automation_path = self.automation_var.get().strip()
        item_path = self.item_var.get().strip()
        # Fältet för ej inlagrade artiklar är borttaget, så ingen separat fil hämtas
        not_putaway_path = ""

        if not orders_path or not buffer_path:
            messagebox.showerror(APP_TITLE, "Välj både beställningsfil och buffertfil.")
            return

        try:
            self._log("Läser in filer...")
            orders_raw = pd.read_csv(orders_path, dtype=str, sep=None, engine="python")
            buffer_raw = pd.read_csv(buffer_path, dtype=str, sep=None, engine="python")

            # Ej inlagrade är borttaget; sätt alltid till None
            self._not_putaway_raw = None
            self._not_putaway_norm = None

            # Saldofil (frivillig)
            if automation_path:
                auto_raw = pd.read_csv(automation_path, dtype=str, sep=None, engine="python")
                # Rensa eventuella BOM/whitespace i kolumnnamn för robustare matchning
                auto_raw_clean = _clean_columns(auto_raw.copy())
                # Spara både rå- och normaliserad saldodata. Den normaliserade används i allokering/refill,
                # medan den rensade råfilen (med Robot-kolumn) används för prognosrapporten.
                self._saldo_raw = auto_raw_clean.copy()
                self._saldo_norm = normalize_saldo(auto_raw_clean)
            else:
                self._saldo_norm = None
                self._saldo_raw = None

            # Item-fil (staplingsbar), frivillig
            self._item_raw = None
            self._item_norm = None
            if item_path:
                try:
                    # Försök läsa automatiskt med python-engine (sniffar separator). dtype=str bevarar strängar.
                    item_raw = pd.read_csv(item_path, dtype=str, sep=None, engine="python")
                except Exception:
                    try:
                        # Fallback: anta tab-separerad
                        item_raw = pd.read_csv(item_path, dtype=str, sep="\t", quoting=3, engine="python")
                    except Exception as ie:
                        raise RuntimeError(f"Kunde inte läsa item-fil: {ie}")
                self._item_raw = item_raw.copy()
                self._item_norm = normalize_items(item_raw)

            orders_raw = _clean_columns(orders_raw)
            buffer_raw = _clean_columns(buffer_raw)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa CSV-filerna:\n{e}")
            return

        # --- Allokering + omklassificering SKRYMMANDE
        try:
            self._log("\n--------------")
            # Log message reflecting the current near‑miss threshold (NEAR_MISS_PCT)
            self._log(f"Kör allokering (Helpall → AutoStore → Huvudplock, FIFO) + {int(NEAR_MISS_PCT * 100)}%-near-miss loggning + Status {sorted(ALLOC_BUFFER_STATUSES)}-filter...")
            result, near = allocate(orders_raw, buffer_raw, log=self._log)

            # Omklassificera huvudplock till SKRYMMANDE utifrån plockplats i saldofil
            result = self._reclassify_skrymmande(result, self._saldo_norm)

            # Lägg till staplingsbar från item-fil om tillgänglig
            try:
                if isinstance(self._item_norm, pd.DataFrame) and not self._item_norm.empty and isinstance(result, pd.DataFrame) and not result.empty:
                    # Hitta kolumn för artikel i resultatet
                    try:
                        art_col_res = find_col(result, ORDER_SCHEMA["artikel"], required=True)
                    except Exception:
                        art_col_res = None
                    if art_col_res:
                        # Slå ihop på artikelnumret
                        temp_merge = result.merge(self._item_norm, how="left", left_on=art_col_res, right_on="Artikel", suffixes=("", "_item"))
                        # Ta bort extra Artikel-kolumn från item_df
                        if "Artikel_item" in temp_merge.columns:
                            temp_merge.drop(columns=["Artikel_item"], inplace=True, errors=False)
                        # Om båda Artikel-kolumner kvar, försök att droppa den vi inte behöver
                        if "Artikel_y" in temp_merge.columns:
                            temp_merge.drop(columns=["Artikel_y"], inplace=True, errors=False)
                        # Efter merge: kombinera staplingsbar från item-fil med befintlig (om någon)
                        # Om både _x och _y finns, prioritera _y (från item), annars _x.
                        if "Ej Staplingsbar_y" in temp_merge.columns or "Ej Staplingsbar_x" in temp_merge.columns:
                            if "Ej Staplingsbar_y" in temp_merge.columns:
                                temp_merge["Ej Staplingsbar"] = temp_merge["Ej Staplingsbar_y"].fillna("")
                            elif "Ej Staplingsbar_x" in temp_merge.columns:
                                temp_merge["Ej Staplingsbar"] = temp_merge["Ej Staplingsbar_x"].fillna("")
                            # Ta bort de temporära kolumnerna
                            for _col in ["Ej Staplingsbar_x", "Ej Staplingsbar_y"]:
                                if _col in temp_merge.columns:
                                    temp_merge.drop(columns=[_col], inplace=True)
                        # Om det fortfarande inte finns kolumn, skapa en tom
                        if "Ej Staplingsbar" not in temp_merge.columns:
                            temp_merge["Ej Staplingsbar"] = ""
                        # Se till att kolumnen för Ej Staplingsbar ligger sist i kolumnordningen
                        cols = [c for c in temp_merge.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
                        temp_merge = temp_merge[cols]
                        result = temp_merge
                # Om kolumnen för Ej Staplingsbar inte har lagts till efter merges (ingen item-fil), skapa en tom kolumn
                if isinstance(result, pd.DataFrame) and ("Ej Staplingsbar" not in result.columns):
                    result["Ej Staplingsbar"] = ""
                    cols = [c for c in result.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
                    result = result[cols]
            except Exception as e:
                # Logga men fortsätt
                try:
                    self._log(f"Kunde inte slå ihop item-fil: {e}")
                except Exception:
                    pass
            self._log("Skapar resultat i minnet...")

            # Spara resultat och near-miss innan statistik bearbetas
            # near_instead_df will be overwritten in the near-miss statistics section but initialise with the raw near-miss data now
            self.last_result_df = result.copy()
            # Store the raw near-miss dataframe here; it will later be enriched with zone information and reassigned
            self.last_nearmiss_instead_df = near.copy()
            self._orders_raw = orders_raw.copy()
            self._buffer_raw = buffer_raw.copy()
            self._result_df = result.copy()

            # Beräkna pallplatsbehov per kund på allokeringsresultatet
            try:
                self._pallet_spaces_df = compute_pallet_spaces(self._result_df)
            except Exception:
                self._pallet_spaces_df = None

            # Uppdatera summering
            try:
                self.update_summary_table(result)
            except Exception as _e_upd:
                self._log(f"Summering per Källtyp kunde inte uppdateras: {_e_upd}")

            # Auto-refill: beräkna och cacha direkt efter allokering
            try:
                hp_df, as_df = calculate_refill(
                    result, buffer_raw,
                    saldo_df=self._saldo_norm,
                    not_putaway_df=self._not_putaway_norm
                )
                # 0-rader är redan bortfiltrerade i beräkningen
                self._last_refill_hp_df = hp_df.copy()
                self._last_refill_autostore_df = as_df.copy()
                self._log(f"Auto-refill klar: HP {len(hp_df)} rader, AUTOSTORE {len(as_df)} rader (cachad).")
            except Exception as e:
                self._last_refill_hp_df = None
                self._last_refill_autostore_df = None
                self._log(f"Auto-refill misslyckades: {e}")

            # Enable-knappar för resultat och near-miss
            self.open_result_btn.configure(state="normal" if not result.empty else "disabled")
            # Enable the near‑miss button only if a non‑empty near‑miss DataFrame was produced.
            # The variable `near` holds the raw near‑miss rows returned from allocate().
            try:
                self.open_nearmiss_btn.configure(state="normal" if isinstance(near, pd.DataFrame) and not near.empty else "disabled")
            except Exception:
                self.open_nearmiss_btn.configure(state="disabled")
            # Aktivera pallplats-knappen om rapporten finns
            try:
                has_pallet = isinstance(self._pallet_spaces_df, pd.DataFrame) and not self._pallet_spaces_df.empty
                self.open_palletspaces_btn.configure(state="normal" if has_pallet else "disabled")
            except Exception:
                # Om något går fel, inaktivera knappen
                self.open_palletspaces_btn.configure(state="disabled")
            # Knappar för refill och försäljningsinsikter finns inte längre
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Fel under allokering:\n{e}")
            return

        # Summering per zon
        try:
            zon_col = "Zon (beräknad)"
            qty_col = find_col(result, ORDER_SCHEMA["qty"], required=True)
            summary = result.groupby(zon_col)[qty_col].apply(lambda s: pd.to_numeric(s, errors="coerce").sum()).reset_index(name="Totalt antal")
            self._log("\nSummering per zon:")
            for _, r in summary.iterrows():
                self._log(f"  Zon {r[zon_col]}: {r['Totalt antal']:.0f}")
        except Exception:
            pass

        # Near-miss statistik
        try:
            # Header for the near‑miss statistics uses the current threshold in percent
            self._log(f"\n{int(NEAR_MISS_PCT * 100)}% near-miss statistik:")
            # Only compute statistics if we have a non-empty near-miss DataFrame
            if isinstance(near, pd.DataFrame) and not near.empty:
                # Identify the article column in the near-miss DataFrame
                near_art_col = None
                for c in ["Artikel", "artikel", "Artikelnummer", "artikelnummer", "_artikel"]:
                    if c in near.columns:
                        near_art_col = c
                        break
                # Identify the article column in the result DataFrame
                res_art_col = None
                try:
                    res_art_col = find_col(result, ORDER_SCHEMA["artikel"], required=False)
                except Exception:
                    # fallback: attempt to find a sensible article column manually
                    for c in ["Artikel", "artikel", "Artikelnummer", "artikelnummer", "_artikel"]:
                        if c in result.columns:
                            res_art_col = c
                            break
                # Zone column in the result DataFrame
                zone_col = "Zon (beräknad)"
                # Create a copy to avoid mutating original near
                near_with_zone = near.copy()
                if near_art_col and res_art_col and zone_col in result.columns:
                    # Build a mapping from article -> zone based on the majority zone in result
                    zone_map: Dict[str, str] = {}
                    # Pre-normalise article column in result for performance
                    res_art_series = result[res_art_col].astype(str).str.strip()
                    # For each unique article in near-miss list, determine which zone it ended up in
                    for art in near_with_zone[near_art_col].astype(str).str.strip().unique():
                        # mask rows for this article
                        mask = res_art_series == art
                        if not mask.any():
                            continue
                        # Determine the most frequent zone for this article
                        zones = result.loc[mask, zone_col].astype(str)
                        if not zones.empty:
                            zone_counts = zones.value_counts()
                            # idxmax returns the first occurrence of the maximum count
                            chosen_zone = zone_counts.idxmax()
                            zone_map[art] = chosen_zone
                    # Assign the chosen zone to each near-miss row
                    near_with_zone["Slutade som Zon"] = near_with_zone[near_art_col].astype(str).str.strip().map(lambda x: zone_map.get(x, ""))
                else:
                    # Fallback: if we cannot determine zones, create an empty column
                    near_with_zone["Slutade som Zon"] = ""
                # Zones to report statistics for
                zones_to_report = ["R", "A", "E", "S", "Q", "O", "F"]
                # Compute and log counts per zone
                for z in zones_to_report:
                    try:
                        cnt = 0
                        if near_art_col:
                            # Count unique articles whose near-miss ended as this zone
                            cnt = int(near_with_zone.loc[near_with_zone["Slutade som Zon"] == z, near_art_col].astype(str).str.strip().nunique())
                        self._log(f"  Near-miss som slutade som {z}: {cnt:,}")
                    except Exception:
                        # If any error occurs for this zone, log zero
                        self._log(f"  Near-miss som slutade som {z}: 0")
                # Log the list of articles with near-miss, one per row
                try:
                    if near_art_col:
                        arts = near_with_zone[near_art_col].astype(str).str.strip().unique().tolist()
                        arts_sorted = sorted(arts)
                        if arts_sorted:
                            self._log("  Artiklar med near-miss:")
                            for art in arts_sorted:
                                self._log(f"    {art}")
                        else:
                            self._log("  Inga near-miss artiklar hittades.")
                    else:
                        self._log("  Inga near-miss artiklar hittades.")
                except Exception:
                    self._log("  Inga near-miss artiklar hittades.")
                # Store the updated near-miss dataframe with zone information for later use (e.g., opening from the GUI)
                self.last_nearmiss_instead_df = near_with_zone.copy()
            else:
                # No near-miss rows present
                self._log("  Inga near-miss artiklar hittades.")
                self.last_nearmiss_instead_df = pd.DataFrame()
        except Exception:
            # Swallow any exception in near-miss stats to avoid breaking the UI
            try:
                self._log("  Inga near-miss artiklar hittades.")
            except Exception:
                pass

# --- main -----------------------------------------------------------------

def main() -> None:
    root_class = TkinterDnD.Tk if TkinterDnD else tk.Tk
    root = root_class()
    root.title(APP_TITLE)
    app = App(root)
    root.geometry("1160x780")
    root.mainloop()

if __name__ == "__main__":
    main()
