#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
allokering12.1.py
---------------
Denna version (12.1) bygger vidare på tidigare versioner och lägger till
ytterligare förbättrad HIB‑koppling och cache‑hantering.

**Nyheter i version 12.1**

* **Förbättrad matchning av HIB‑ordrar**: Programmet matchar nu varje
  HIB‑orders sändningsnummer **och kundnamn** mot butikens ordrar. Om en
  butiksorder har samma sändningsnummer **och tillhör samma butik** (identiskt
  kundnamn) väljs den som referens (den tidigaste om flera finns). Om ingen
  sådan order finns matchas endast på sändningsnummer. I sista hand används
  den äldsta giltiga butiksordern som fallback.

  Denna version korrigerar även ett problem där HIB‑ordrar felaktigt
  föreslogs kopplas om till en annan butik när butikens order med rätt
  sändningsnummer hade status ≥ 34 och därför inte räknades som giltig. Nu
  används **alla butiksordrar** (oavsett status) för att hitta matchning på
  sändningsnummer och kundnamn. Endast om ingen sådan matchning hittas
  används fallback‑butiken.

* **Statushantering för butiksorder**: Butiksorder som saknar status i
  orderdetaljerna behandlas som giltiga (status 0) istället för att uteslutas.

* **Rensning av dispatchpallar vid cache‑reset**: När man väljer “Rensa
  cache” i GUI:et rensas nu även den valda dispatchpallsfilen och alla
  temporära dispatchresultat. Detta förhindrar att en gammal dispatchfil
  ligger kvar i minnet efter att man bytt dataset.

* **Övriga förbättringar från version 10.7** behålls, såsom robust
  kolumnmatchning, förbättrat GUI för filuppladdning och mer intuitiv
  statusvisning. Multi‑reglerna (fel zon, saknad multi) gäller bara när det
  finns mer än en HIB‑order per kundnummer. Med endast en HIB‑order sätts
  inte multi. Instruktionerna för ändringsordning skrivs ut i loggen och
  exportfilen. Kolumnmatchningen i orderöversikten är robust mot olika
  ordning och namn. Indata‑filvalet använder tydliga statusrutor med text
  ("Uppladdad" med grön bakgrund respektive "Ej fil" med grå bakgrund) och en
  röd borttagningsknapp. Drag‑och‑släpp‑zonen kan även klickas för att välja
  flera filer samtidigt. Fixar från 10.6 för korrekt initiering av
  statusikoner gäller fortsatt.
"""

from __future__ import annotations

import argparse
import base64
import json
import math
import platform
import queue
import re
import shutil
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Deque, Dict, List, Tuple, Optional
import importlib.util
import webbrowser
from pathlib import Path

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

from app_info import (
    APP_NAME,
    APP_TITLE,
    APP_VERSION,
    ANALYTICS_ENABLED_DEFAULT,
    ANALYTICS_ENABLED_ENV,
    ANALYTICS_HOST_ENV,
    ANALYTICS_POSTHOG_HOST,
    ANALYTICS_POSTHOG_PROJECT_API_KEY,
    ANALYTICS_PROJECT_API_KEY_ENV,
    GITHUB_RELEASES_URL,
    UPDATE_DISABLED_ENV,
)
from update_service import UpdateInfo, check_for_update, download_update_installer


SILENT_UPDATE_ARGS = [
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART",
    "/CLOSEAPPLICATIONS",
    "/FORCECLOSEAPPLICATIONS",
]


ANALYTICS_CONFIG_KEY_ENABLED = "analytics_enabled"
ANALYTICS_CONFIG_KEY_DISTINCT_ID = "analytics_distinct_id"
ANALYTICS_CONFIG_KEY_HOST = "analytics_host"
ANALYTICS_CONFIG_KEY_PROJECT_API_KEY = "analytics_project_api_key"


def _parse_boolish(value: object) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "ja", "on"}:
        return True
    if text in {"0", "false", "no", "nej", "off"}:
        return False
    return None


def _app_config_path() -> Path:
    appdata = os.environ.get("APPDATA") or str(Path.home())
    return Path(appdata) / "allokering" / "config.json"


def _load_app_config() -> dict:
    path = _app_config_path()
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_app_config(config: dict) -> None:
    path = _app_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(config, indent=2, ensure_ascii=False, sort_keys=True)
    path.write_text(payload, encoding="utf-8")


def _get_or_create_analytics_distinct_id(config: Optional[dict] = None) -> str:
    cfg = config if isinstance(config, dict) else _load_app_config()
    existing = str(cfg.get(ANALYTICS_CONFIG_KEY_DISTINCT_ID, "")).strip()
    if existing:
        return existing
    distinct_id = uuid.uuid4().hex
    cfg[ANALYTICS_CONFIG_KEY_DISTINCT_ID] = distinct_id
    try:
        _save_app_config(cfg)
    except Exception:
        pass
    return distinct_id


def _resolve_analytics_settings(enable_analytics: bool = True) -> dict:
    cfg = _load_app_config()
    enabled_pref = _parse_boolish(os.environ.get(ANALYTICS_ENABLED_ENV))
    if enabled_pref is None:
        enabled_pref = _parse_boolish(cfg.get(ANALYTICS_CONFIG_KEY_ENABLED))
    if enabled_pref is None:
        enabled_pref = ANALYTICS_ENABLED_DEFAULT

    host = str(
        os.environ.get(ANALYTICS_HOST_ENV)
        or cfg.get(ANALYTICS_CONFIG_KEY_HOST)
        or ANALYTICS_POSTHOG_HOST
        or ""
    ).strip()
    api_key = str(
        os.environ.get(ANALYTICS_PROJECT_API_KEY_ENV)
        or cfg.get(ANALYTICS_CONFIG_KEY_PROJECT_API_KEY)
        or ANALYTICS_POSTHOG_PROJECT_API_KEY
        or ""
    ).strip()

    distinct_id = ""
    if enable_analytics or enabled_pref:
        distinct_id = _get_or_create_analytics_distinct_id(config=cfg)

    if not enable_analytics:
        reason = "Analytics är tillfälligt avstängt för detta körläge."
    elif not enabled_pref:
        reason = "Användaren har stängt av anonym användningsstatistik."
    elif not api_key:
        reason = "Ingen PostHog project API key är konfigurerad ännu."
    else:
        reason = ""

    active = bool(enable_analytics and enabled_pref and api_key and host)
    return {
        "active": active,
        "enabled_preference": bool(enabled_pref),
        "host": host,
        "api_key": api_key,
        "distinct_id": distinct_id,
        "reason": reason,
    }


class AnalyticsClient:
    def __init__(self, settings: dict):
        self.active = bool(settings.get("active"))
        self.host = str(settings.get("host", "")).rstrip("/")
        self.api_key = str(settings.get("api_key", "")).strip()
        self.distinct_id = str(settings.get("distinct_id", "")).strip()
        self.reason = str(settings.get("reason", "")).strip()
        self.session_id = uuid.uuid4().hex
        self._queue: "queue.Queue[object]" = queue.Queue()
        self._stop_token = object()
        self._thread: Optional[threading.Thread] = None
        if self.active:
            self._thread = threading.Thread(target=self._worker, name="analytics-worker", daemon=True)
            self._thread.start()

    def capture(self, event: str, properties: Optional[dict] = None) -> None:
        if not self.active:
            return
        props = dict(properties or {})
        props.setdefault("distinct_id", self.distinct_id)
        props.setdefault("session_id", self.session_id)
        props.setdefault("app_name", APP_NAME)
        props.setdefault("app_version", APP_VERSION)
        props.setdefault("platform", platform.system())
        props.setdefault("platform_release", platform.release())
        props.setdefault("python_version", platform.python_version())
        props.setdefault("frozen", bool(getattr(sys, "frozen", False)))
        payload = {
            "api_key": self.api_key,
            "event": event,
            "properties": props,
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        }
        try:
            self._queue.put_nowait(payload)
        except Exception:
            pass

    def shutdown(self, timeout: float = 1.5) -> None:
        if not self._thread or not self._thread.is_alive():
            return
        try:
            self._queue.put_nowait(self._stop_token)
        except Exception:
            return
        self._thread.join(timeout=max(0.1, timeout))

    def _worker(self) -> None:
        capture_url = f"{self.host}/capture/"
        while True:
            item = self._queue.get()
            if item is self._stop_token:
                break
            try:
                body = json.dumps(item).encode("utf-8")
                request = urllib.request.Request(
                    capture_url,
                    data=body,
                    method="POST",
                    headers={
                        "Content-Type": "application/json",
                        "User-Agent": "allokering-analytics",
                    },
                )
                with urllib.request.urlopen(request, timeout=5):
                    pass
            except Exception:
                continue


def _bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resource_path(*parts: str) -> Path:
    return _bundle_root().joinpath(*parts)


LOWFREQDATA_DIR = "lowfreqdata"
BUFFERPALL_DIR = "buffertpall"
ITEM_OPTION_DIR = "item-option"
BUFFERPALL_PATH_PARTS = (LOWFREQDATA_DIR, BUFFERPALL_DIR)


def _bufferpall_resource_path(*parts: str) -> Path:
    return _resource_path(*BUFFERPALL_PATH_PARTS, *parts)


def _bufferpall_runtime_dir() -> Path:
    return _runtime_root().joinpath(*BUFFERPALL_PATH_PARTS)


def _bufferpall_source_dir() -> Path:
    return Path(__file__).resolve().parent.joinpath(*BUFFERPALL_PATH_PARTS)


def _seed_bufferpall_runtime_file(filename: str) -> Path:
    runtime_path = _bufferpall_runtime_dir() / filename
    runtime_path.parent.mkdir(parents=True, exist_ok=True)

    resource_path = _bufferpall_resource_path(filename)
    if not runtime_path.exists() and resource_path.exists() and resource_path.resolve() != runtime_path.resolve():
        shutil.copy2(resource_path, runtime_path)
    return runtime_path

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
    df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
    drop_idx = [i for i in (0, 1, 3) if i < len(df.index)]
    df = df.drop(index=drop_idx, errors="ignore").reset_index(drop=True)
    if df.shape[1] > 0:
        df = df.drop(columns=[df.columns[0]]).reset_index(drop=True)
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
    header = df.iloc[0].astype(str).str.strip().tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = header
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


def read_campaign_xlsx(path: str) -> pd.DataFrame:
    """
    Läs och normalisera en kampanjvolymfil (XLSX) enligt en fördefinierad sekvens av borttagningar av rader och kolumner.
    Returnerar ett DataFrame med kolumnerna:
      - Artikelnummer (str)
      - Antal styck (int)
    """
    df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
    if df.empty:
        return pd.DataFrame(columns=["Artikelnummer", "Antal styck"])
    if len(df.index) > 4:
        df = df.drop(index=[4])
    drop_idx = [i for i in (0, 1, 2) if i < len(df.index)]
    df = df.drop(index=drop_idx)
    df = df.reset_index(drop=True)
    keep_cols = [c for c in df.columns if c <= 6]
    df = df.loc[:, keep_cols]
    if 5 in df.columns:
        df = df.drop(columns=[5])
    if 4 in df.columns:
        df = df.drop(columns=[4])
    if 3 in df.columns:
        df = df.drop(columns=[3])
    if 1 in df.columns:
        df = df.drop(columns=[1])
    if 0 in df.columns:
        df = df.drop(columns=[0])
    if df.shape[1] != 2:
        return pd.DataFrame(columns=["Artikelnummer", "Antal styck"])
    df = df.reset_index(drop=True)
    df.columns = ["Artikelnummer", "Antal styck"]
    df["Artikelnummer"] = df["Artikelnummer"].astype(str).str.strip()
    df["Antal styck"] = pd.to_numeric(df["Antal styck"], errors="coerce").fillna(0).astype(int)
    df = df.loc[df["Artikelnummer"].astype(str).str.len().gt(0)].reset_index(drop=True)
    if not df.empty and str(df.loc[0, "Artikelnummer"]).lower() in ("produktkod", "#"):
        df = df.drop(index=[0]).reset_index(drop=True)
    return df


# Uppdaterad programversion 12.1
DEFAULT_OUTPUT = "allocated_orders.csv"

INVALID_LOC_PREFIXES: Tuple[str, ...] = ("AA",)
INVALID_LOC_EXACT: set[str] = {"TRANSIT", "TRANSIT_ERROR", "MISSING", "UT2"}

ALLOC_BUFFER_STATUSES: set[int] = {29, 30, 32}
REFILL_BUFFER_STATUSES: set[int] = {29, 30}

# Vecka 27 - tak/hus -> tillåtna matchande gräsklippare (per order krävs minst lika många gräsklippare som tak)
VECKA27_ROOF_TO_MOWERS: dict[str, frozenset[str]] = {
    "2002039": frozenset({"2003511", "2003512", "2002034", "2002035", "2002036"}),
    "2001926": frozenset({"2003708", "2003709"}),
    "2005080": frozenset({"2003482", "2003483", "2003484", "2003485", "2003486"}),
    "2001928": frozenset({"2001921", "2001922", "2001923"}),
    "2003711": frozenset({"2003710"}),
}

NEAR_MISS_PCT: float = 0.30  # 30 % över behov

# Artiklar som undantas från R+F-räkningen i compute_pallet_spaces
RF_PALLPLATS_EXCLUDE_ARTICLES: set[str] = {
    "1075621","1154474","1265531","1265532","1265533","1265534","1265535","1265536","1265537","1265539",
    "1265541","1265542","1265543","1265545","1265547","1265548","1265549","1265550","1265551","1265552",
    "1265553","1265554","1265555","1265557","1265558","1265559","1265560","1265561","1265562","1265563",
    "1265564","1265565","1265566","1265567","1265568","1265569","1265570","1265571","1265572","1265573",
    "1265575","1265576","1265578","1265579","1265580","1265581","1265582","1265583","1265584","1265585",
    "1265586","1265588","1265589","1265590","1265591","1265592","1265593","1265594","1265595","1265596",
    "1265598","1265601","1265602","1265603","1265604","1265605","1265606","1265607","1265608","1265609",
    "1265610","1265612","1265613","1265614","1265615","1265617","1265618","1265619","1265620","1265621",
    "1265622","1265623","1265624","1265625","1265626","1265627","1265628","1265629","1265630","1265631",
    "1265632","1265633","1265634","1265635","1265636","1265637","1265638","1265639","1265640","1265641",
    "1265642","1265643","1265644","1265645","1265646","1265651","1265652","1265653","1265654","1265655",
    "1265656","1265657","1265658","1265659","1265660","1265661","1265662","1265663","1265664","1265665",
    "1265666","1265667","1265669","1265671","1265672","1265673","1265674","1265675","1265676","1265677",
    "1265678","1265679","1265680","1265681","1265682","1265683","1265684","1265685","1265687","1265689",
    "1265690","1265692","1265693","1265694","1265695","1265696","1265697","1265698","1265699","1265700",
    "1265701","1265702","1265703","1265704","1265705","1265706","1265707","1265708","1265709","1265710",
    "1265711","1265712","1265713","1265714","1265715","1265716","1265717","1265718","1265719","1265720",
    "1265721","1265722","1265723","1265724","1265725","1265727","1265728","1265729","1265730","1265731",
    "1265733","1265734","1265735","1265737","1265738","1265739","1265740","1265741","1265742","1265743",
    "1265744","1265745","1265746","1265747","1265748","1265749","1265750","1265751","1265754","1265755",
    "1265756","1265757","1265760","1265762","1265763","1265764","1265765","1265766","1265768","1265770",
    "1265771","1265772","1265773","1265774","1265775","1265778","1265779","1265780","1265781","1265782",
    "1265783","1265784","1265785","1265786","1265787","1265788","1265789","1265790","1265791","1265793",
    "1265794","1265795","1265797","1265798","1265799","1265800","1265801","1265802","1265803","1265804",
    "1265805","1265806","1265807","1265808","1265809","1265810","1265811","1265812","1265813","1265814",
    "1265815","1265816","1265817","1265818","1265821","1265822","1265823","1265826","1265827","1265828",
    "1265829","1265830","1265832","1265833","1265834","1265835","1265837","1265838","1265839","1265840",
    "1265841","1265842","1265843","1265844","1265846","1265847","1265848","1265849","1265850","1265851",
    "1265852","1265853","1265854","1265855","1265856","1265857","1265858","1265859","1265860","1265861",
    "1265862","1265863","1265864","1265865","1265866","1265867","1265868","1265869","1265870","1265871",
    "1265872","1265873","1265874","1265876","1265877","1265878","1265879","1265880","1265881","1265882",
    "1265883","1265884","1265885","1265886","1265887","1265888","1265889","1265890","1265891","1265892",
    "1265894","1265895","1265896","1265897","1265899","1265900","1265902","1265903","1265904","1265905",
    "1265906","1265907","1265908","1265909","1265910","1265911","1265912","1265913","1265915","1265916",
    "1265917","1265918","1265919","1265920","1265921","1265923","1265924","1265925","1265926","1265927",
    "1265928","1265929","1265930","1265931","1265932","1265933","1265934","1265935","1265936","1265937",
    "1265938","1265939","1265940","1265941","1265942","1265943","1265944","1265945","1265946","1265947",
    "1265948","1265951","1265952","1265953","1265954","1265955","1265956","1265957","1265958","1265959",
    "1265960","1265961","1265963","1265965","1265966","1265967","1265968","1265969","1265970","1265971",
    "1265972","1265973","1265974","1265975","1265976","1265977","1265978","1265979","1265980","1265981",
    "1265983","1265984","1265985","1265986","1265987","1265988","1265989","1265991","1265992","1265993",
    "1265994","1265995","1265996","1265997","1265998","1265999","1266000","1266001","1266002","1266003",
    "1266004","1266005","1266006","1266008","1266009","1266010","1266011","1266012","1266013","1266014",
    "1266015","1266017","1266018","1266019","1266020","1266021","1266022","1266023","1266024","1266025",
    "1266026","1266027","1266034","1266035","1266036","1266037","1266038","1266039","1266040","1266041",
    "1266042","1266043","1266044","1266045","1266046","1266047","1266048","1266049","1266050","1266051",
    "1266052","1266053","1266054","1266056","1266057","1266058","1266059","1266060","1266061","1266062",
    "1266063","1266065","1266066","1266067","1266068","1266069","1266070","1266072","1266073","1266074",
    "1266075","1266076","1266077","1266078","1266079","1266081","1266082","1266084","1266085","1266086",
    "1266087","1266088","1266089","1266091","1266093","1266094","1266095","1266096","1266097","1266099",
    "1266100","1266101","1266102","1266231","1266233","1266234","1266236","1266237","1266238","1266239",
    "1266240","1266241","1266242","1266244","1266245","1266246","1266247","1266248","1266249","1266251",
    "1266252","1266253","1266254","1266255","1266256","1266257","1266260","1266261","1266262","1266263",
    "1266264","1266265","1266266","1266268","1266270","1266271","1266272","1266273","1266274","1266275",
    "1266276","1266277","1266279","1266280","1266283","1266284","1266285","1266863","1266864","1266865",
    "1266866","1266868","1266872","1266873","1266874","1266875","1266876","1267022","1267023","1267024",
    "1267025","1267031","1267033","1267034","1267043","1267044","1267045","1267046","1267048","1267050",
    "1267054","1267055","1267059","1267064","1267067","1267086","1267090","1267093","1267104","1267116",
    "1267119","1267121","1267122","1267124","1267127","1268095","1268097","1268167","1268168","1268169",
    "1268170","1268171","1268172","1268173","1268174","1268175","1268176","1268177","1268178","1268179",
    "1268180","1268181","1268182","1268183","1268184","1269119","1269120","1269189","1269190","1269191",
    "1269192","1269193","1269194","1269195","1269196","1269197","1269198","1269199","1269200","1269201",
    "1269202","1269203","1269204","1269205","1269206","1269207","1269208","1269239","1269243","1269244",
    "1269245","1269246","1269247","1269250","1269251","1269252","1269253","1269254","1269255","1269256",
    "1269258","1269259","1269260","1269263","1269264","1269265","1269267","1269268","1269270","1269271",
    "1269272","1269273","1270087","1270088","1270089","1270090","1270091","1270092","1270093","1270094",
    "1270095","1270096","1270097","1270098","1270099","1270100","1270101","1270102","1270103","1270104",
    "1270105","1270106","1270107","1270108","1270109","1270110","1270111","1270112","1270113","1270114",
    "1270115","1270116","1270117","1270118","1270119","1270120","1270121","1270122","1270123","1270124",
    "1270125","1270126","1270127","1270128","1270129","1270130","1270131","1270132","1270133","1270134",
    "1270135","1270136","1270137","1270138","1270139","1270140","1270141","1270142","1270143","1270144",
    "1270145","1270146","1270147","1270148","1270149","1270150","1270151","1270152","1270153","1270154",
    "1270155","1270156","1270157","1270158","1270159","1270160","1270161","1270162","1270163","1270164",
    "1270165","1270166","1270167","1270168","1270169","1270170","1270171","1270172","1270173","1270174",
    "1270175","1270176","1270177","1270178","1270179","1270180","1270181","1270182","1270183","1270184",
    "1270185","1270186","1270187","1270188","1270189","1270190","1270191","1270192","1270193","1270194",
    "1270195","1270196","1270197","1270198","1270199","1270200","1270201","1270202","1270203","1270204",
    "1270205","1270206","1270207","1270208","1270209","1270210","1270211","1270212","1270213","1270214",
    "1270215","1270216","1270217","1270218","1270219","1270220","1270221","1270222","1270223","1270224",
    "1270225","1270226","1270227","1270228","1270229","1270230","1270231","1270232","1270233","1270234",
    "1270235","1270547","1270548","1270549","1270550","1270551","1270552","1270553","1270554","1270555",
    "1270556","1270557","1270558","1270559","1270560","1270561","1270634","2001334","2001335","2001336",
    "2001337","2001338","2001339","2001340","2001341","2001342","2001343","2001344","2001345","2001346",
    "2001347","2001348","2001349","2001350","2001351","2001352","2001353","2001354","2001355","2001356",
    "2001357","2001358","2001359","2001360","2001361","2001362","2001363","2001364","2001365","2001366",
    "2001367","2001368","2001369","2001370","2001371","2001372","2001373","2001374","2001375","2001376",
    "2001377","2001378","2001379","2001380","2001381","2001382","2001433","2001434","2001435","2001436",
    "2003381","2003382","2003383","2003384","2003385","2003386","2003387","2003388","2003389","2003390",
    "2003391","2003392","2003393","2003394","2003395","2003396","2003397","2003398","2003399","2003400",
    "2003401","2003402","2003403","2003404","2003405","2003406","2003407","2003408","2003409","2003410",
    "2003411","2003412","2003413","2003414","2003415","2003416","2003417","2003418","2003419","2003420",
    "2003423","2003424","2003425","2003426","2003427","2003428","2003429","2003430","2003431","2003432",
    "2003433","2003434","2003435","2003436","2003437","2003438","2003439","2003440","2003441","2003442",
    "2003443","2003444","2003445","2003448","2003449","2003450","2003451","2003452","2003453","2003454",
    "2003455","2003456","2003457","2003458","2003459","2003460","2003461","2003462","2003463","2003464",
    "2003465","2003466","2003467","2003514","2003515","1169745","1267354","1267355","1169747","1169746",
    "1267358","1267357","1267356",
}


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

SALDO_SCHEMA: Dict[str, List[str]] = {
    "artikel":    ["artikel", "artnr", "art.nr", "artikelnummer", "sku", "article"],
    "plocksaldo": ["plocksaldo", "plock saldo", "plock-saldo", "saldo", "pick saldo", "pick qty",
                   "tillgängligt plock", "tillgangligt plock", "available pick", "plock"],
    "plockplats": ["plockplats", "huvudplock", "mainpick", "hyllplats", "bin", "location", "lagerplats"],
}

ITEM_SCHEMA: Dict[str, List[str]] = {
    "artikel": ORDER_SCHEMA["artikel"],  # återanvänd artikel-kandidater från beställningar
    "staplingsbar": [
        "staplingsbar", "staplings bar", "staplbar", "stackable",
        "ej staplingsbar", "ejstaplingsbar", "ej_staplingsbar", "non stackable"
    ]
}


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


def _open_df_in_excel_with_bold_cells(
    df,
    *,
    sheet_name: str,
    label: str,
    bold_cells: Optional[set[tuple[int, int]]] = None,
    bold_sheet_name: Optional[str] = None,
) -> str:
    """Skriv en Excel-fil och fetstila utvalda dataceller i vald sheet."""
    import importlib

    bold_targets = {(int(r), int(c)) for r, c in (bold_cells or set()) if r >= 0 and c >= 0}
    if isinstance(df, dict):
        sheet_map = {
            (str(name)[:31] or "Sheet1"): (value if isinstance(value, pd.DataFrame) else pd.DataFrame(value))
            for name, value in df.items()
        }
        target_sheet = str(bold_sheet_name or sheet_name)[:31] or "Sheet1"
        if target_sheet not in sheet_map and sheet_map:
            target_sheet = next(iter(sheet_map))
    else:
        safe_sheet = str(sheet_name)[:31] or "Sheet1"
        sheet_map = {safe_sheet: (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df))}
        target_sheet = safe_sheet

    if importlib.util.find_spec("openpyxl"):
        from openpyxl.styles import Font

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        path = tmp.name
        tmp.close()
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, out_df in sheet_map.items():
                out_df.to_excel(writer, sheet_name=name, index=False)
            if target_sheet in writer.sheets:
                out_df = sheet_map[target_sheet]
                ws = writer.sheets[target_sheet]
                bold_font = Font(bold=True)
                for row_idx, col_idx in bold_targets:
                    if row_idx >= len(out_df.index) or col_idx >= len(out_df.columns):
                        continue
                    ws.cell(row=row_idx + 2, column=col_idx + 1).font = bold_font
    elif importlib.util.find_spec("xlsxwriter"):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        path = tmp.name
        tmp.close()
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            for name, out_df in sheet_map.items():
                out_df.to_excel(writer, sheet_name=name, index=False)
            if target_sheet in writer.sheets:
                out_df = sheet_map[target_sheet]
                ws = writer.sheets[target_sheet]
                bold_format = writer.book.add_format({"bold": True})
                for row_idx, col_idx in bold_targets:
                    if row_idx >= len(out_df.index) or col_idx >= len(out_df.columns):
                        continue
                    ws.write(row_idx + 1, col_idx, out_df.iat[row_idx, col_idx], bold_format)
    else:
        raise RuntimeError("Saknar Excel-skrivare (installera 'openpyxl' eller 'xlsxwriter').")

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


def _open_text_in_editor(text: str, label: str = "rapport") -> str:
    """Skriv text till temporär .txt-fil och öppna i system-editor. Returnerar sökvägen."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.txt")
    path = tmp.name
    tmp.close()
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
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

# -----------------------------------------------
# Ny funktion för HIB‑koppling
# Denna funktion tar beställningsrader och orderöversikt och räknar ut vilka HIB‑ordrar
# som behöver uppdateras. Resultatet returneras som ett DataFrame med kolumnerna
# "ordernummer", "Orderdatum", "sändningsnummer", "Zon" och "Multi". Endast
# ordrar med minst en ändring inkluderas i resultatet. Zonuppdateringar anges alltid
# som "F" om någon orderrad inte ligger i zon F/H/R. Multi sätts till "MULTI"
# om antingen flera olika multi‑nummer finns för kundens HIB‑ordrar i zon F, om
# något multi‑nummer saknas eller om någon HIB‑order behöver zonuppdatering.

def _hib_orders_with_today_origin(kund_df: pd.DataFrame) -> set[str]:
    """Returnera HIB-ordrar vars Ursprungsdatum är samma som dagens kördatum."""
    if (
        not isinstance(kund_df, pd.DataFrame)
        or kund_df.empty
        or "Ordernr" not in kund_df.columns
        or "Ordertyp" not in kund_df.columns
        or "Ursprungsdatum" not in kund_df.columns
    ):
        return set()

    try:
        parsed = smart_to_datetime(kund_df["Ursprungsdatum"])
        today = pd.Timestamp.now().date()
        mask = (
            kund_df["Ordertyp"].astype(str).str.strip().str.upper().eq("HIB")
            & parsed.notna()
            & parsed.dt.date.eq(today)
        )
    except Exception:
        today_str = pd.Timestamp.now().strftime("%Y-%m-%d")
        mask = (
            kund_df["Ordertyp"].astype(str).str.strip().str.upper().eq("HIB")
            & kund_df["Ursprungsdatum"].astype(str).str.strip().eq(today_str)
        )

    if not mask.any():
        return set()

    return set(kund_df.loc[mask, "Ordernr"].astype(str).str.strip())


def compute_hib_koppling(
    details_df: pd.DataFrame, overview_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Analysera orderdetaljer (beställningslinjer) och orderöversikt för att
    identifiera vilka HIB‑ordrar som behöver ändras.  Resultatet innehåller
    följande kolumner:

      - ordernummer: Ordernumret för HIB‑ordern.
      - Ursprungsdatum: Ursprungsdatum från orderöversikten om det finns.
      - Orderdatum: Nytt orderdatum om det skiljer sig från butikens orderdatum.
      - sändningsnummer: Nytt sändningsnummer om det skiljer sig från butikens order.
      - Zon: "F" om någon rad inte ligger i zon F/H/R och därför måste sättas till F.
      - Multi: "MULTI" om det behövs ett nytt multi‑nummer för kundens HIB‑ordrar.

    Endast ordrar där minst en kolumn behöver uppdateras inkluderas i resultatet.
    """
    # Kopiera och städa kolumnnamn (ta bort BOM och trimma blanksteg)
    details = details_df.copy()
    overview = overview_df.copy()
    details.columns = [str(c).replace("\ufeff", "").strip() for c in details.columns]
    overview.columns = [str(c).replace("\ufeff", "").strip() for c in overview.columns]
    # Map synonyms in overview columns to canonical names so that column order and variations do not matter.
    synonyms = {
        "Ordernr": ["Ordernr", "Order nr", "Order number", "Ordernummer"],
        "Ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
        "Kund nr": ["Kund nr", "Kundnr", "Kundnummer", "Customer number", "Kund NR"],
        "Bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
        "Orderdatum": ["Orderdatum", "Order datum", "Order date", "Orderdate"],
        "Sändningsnr": [
            "Sändningsnr",
            "Sändnings nr",
            "Sändningsnummer",
            "Sendingsnr",
            "Sändnings number",
        ],
        "Zon": ["Zon", "Zone"],
        "Multi": ["Multi", "Multi nr", "Multinr", "Multi number"],
        "Ursprungsdatum": ["Ursprungsdatum", "Ursprungs datum", "Original date", "Ursprungsdate"],
    }
    for canonical, syns in synonyms.items():
        if canonical in overview.columns:
            continue
        for candidate in syns:
            # search for a matching column, case-insensitive after stripping spaces
            for col in list(overview.columns):
                if col.strip().lower() == candidate.strip().lower():
                    overview.rename(columns={col: canonical}, inplace=True)
                    break
            if canonical in overview.columns:
                break

    # Säkerställ att nödvändiga kolumner finns, annars returnera tomt df
    required_overview_cols = {"Ordernr", "Ordertyp", "Kund nr", "Orderdatum", "Sändningsnr", "Zon", "Multi"}
    missing = [c for c in required_overview_cols if c not in overview.columns]
    if missing:
        return pd.DataFrame(columns=["ordernummer", "Ursprungsdatum", "Orderdatum", "sändningsnummer", "Zon", "Multi"])

    # Normalisera ordertyp men låt indata-filtret styra vilka rader som ingår.
    ov = overview.copy()
    ov["Ordertyp"] = ov["Ordertyp"].astype(str).str.strip().str.upper()
    if ov.empty:
        return pd.DataFrame(columns=["ordernummer", "Ursprungsdatum", "Orderdatum", "sändningsnummer", "Zon", "Multi"])

    # Samla status per order från beställningslinjerna
    details.columns = [c.replace("\ufeff", "").strip() for c in details.columns]
    # Säkerställ att vi har nödvändiga kolumner även där
    if "Order nr" not in details.columns or "Status" not in details.columns:
        return pd.DataFrame(columns=["ordernummer", "Ursprungsdatum", "Orderdatum", "sändningsnummer", "Zon", "Multi"])

    # Konvertera status till tal när möjligt (allt som inte går tolkas som stort tal för att markera ej OK)
    def to_status_numeric(x):
        try:
            return int(float(str(x).strip()))
        except Exception:
            return 9999

    details["_status_num"] = details["Status"].apply(to_status_numeric)

    # Map för order -> max status
    order_status_max = details.groupby("Order nr")["_status_num"].max().to_dict()

    # Map för order -> zoner i beställningslinjer
    order_zones = details.groupby("Order nr")["Zon"].apply(lambda x: list(x.dropna().astype(str))).to_dict()

    # Skapa mappning från ordernummer till kundnamn (butiksnamn) om möjligt
    order_to_kundnamn: dict[str, str] = {}
    if "Order nr" in details.columns and "Kund.1" in details.columns:
        try:
            order_to_kundnamn = (details.groupby("Order nr")["Kund.1"].first()
                                 .fillna("")
                                 .astype(str)
                                 .str.strip()
                                 .to_dict())
        except Exception:
            order_to_kundnamn = {}

    # Resultatlista
    rows: list[dict] = []

    # Gruppera orderöversikten efter kundnummer
    for kund_nr, kund_df in ov.groupby("Kund nr"):
        # Hämta butikens order (Ordertyp N) och hämta deras ordernummer
        # Hitta butiksordrar (Ordertyp N) och HIB‑ordrar, men deduplicera per ordernummer
        store_df = kund_df[kund_df["Ordertyp"] == "N"].copy()
        hib_df = kund_df[kund_df["Ordertyp"] == "HIB"].copy()
        ignored_hib_orders = _hib_orders_with_today_origin(kund_df)
        if ignored_hib_orders:
            hib_df = hib_df[~hib_df["Ordernr"].astype(str).str.strip().isin(ignored_hib_orders)].copy()
        # Deduplicera för att undvika att samma order behandlas flera gånger (en rad per zon i orderöversikten)
        if not store_df.empty:
            store_df = store_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
        if not hib_df.empty:
            hib_df = hib_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
        if store_df.empty or hib_df.empty:
            # inga HIB att koppla eller ingen butik => hoppa över
            continue
        # Filtrera butiksordrar där alla statusar är < 34
        # Detta bildar listan av "giltiga" butiksordrar som kan användas som referens.
        #
        # OBS! Vissa butiksorder kan sakna status i orderdetaljerna. Tidigare
        # användes ett defaultvärde på 9999 vilket uteslöt dessa order från matchning.
        # Det ledde till att en HIB-order med korrekt sändningsnummer och datum ändå
        # kopplades om till en annan butik. För att behandla sådana butiksorder som
        # giltiga sätts nu defaultstatus till 0 istället för 9999. Då inkluderas
        # butiksorder som saknar statusuppgift.
        valid_store_df = store_df[store_df["Ordernr"].map(lambda ordnum: order_status_max.get(ordnum, 0) < 34)].copy()
        if valid_store_df.empty:
            # ingen giltig butiksorder att koppla mot
            continue

        # Hjälpfunktion: välj den butiksorder som har äldst orderdatum i ett givet DataFrame
        def _choose_earliest(df: pd.DataFrame) -> pd.Series:
            # Börja med första rad som referens
            earliest_row = df.iloc[0]
            earliest_date = str(earliest_row["Orderdatum"]).strip()
            for _, r in df.iterrows():
                date_str = str(r["Orderdatum"]).strip()
                try:
                    d_new = pd.to_datetime(date_str, errors="coerce")
                    d_old = pd.to_datetime(earliest_date, errors="coerce")
                    if (pd.isna(d_old) and not pd.isna(d_new)) or (
                        not pd.isna(d_old) and not pd.isna(d_new) and d_new < d_old
                    ):
                        earliest_row = r
                        earliest_date = date_str
                    elif pd.isna(d_new) and pd.isna(d_old) and date_str < earliest_date:
                        earliest_row = r
                        earliest_date = date_str
                except Exception:
                    # Fallback: jämför strängar om datumkonvertering misslyckas
                    if date_str < earliest_date:
                        earliest_row = r
                        earliest_date = date_str
            return earliest_row

        # Fallback‑butiksorder: den med äldst orderdatum bland giltiga
        fallback_store_row = _choose_earliest(valid_store_df)
        # Undersök HIB‑ordrar som är tillåtna (alla status < 34)
        hib_orders: list[dict] = []
        for _, hib_row in hib_df.iterrows():
            h_ord = hib_row["Ordernr"]
            # Kontrollera status
            maxstatus = order_status_max.get(h_ord, 9999)
            if maxstatus >= 34:
                continue  # denna hib får inte ändras
            hib_orders.append({"row": hib_row, "ordernr": h_ord})
        if not hib_orders:
            continue
        # Bestäm zon‑flagga per hibordernummer
        zone_flag = False  # om någon rad ej är F/H/R => True
        hib_zone_updates = {}  # ordernummer -> zon_update ("F" eller "")
        for hib in hib_orders:
            h_ord = hib["ordernr"]
            zones = [z.strip().upper() for z in order_zones.get(h_ord, []) if str(z).strip()]
            # Om det finns minst en zon som inte är F, H eller R
            if any(z not in ("F", "H", "R") for z in zones):
                zone_flag = True
                hib_zone_updates[h_ord] = "F"
            else:
                hib_zone_updates[h_ord] = ""
        # Bestäm multi‑nummer per order i zon F
        # Samla multi‑nummer för varje HIB‑order i zon F (i orderöversikten)
        hib_f_multi: dict[str, list[str]] = {}
        missing_multi_per_order: dict[str, bool] = {}
        for hib in hib_orders:
            h_ord = hib["ordernr"]
            # Alla rader i kund_df för denna order där zon är F
            hib_zone_rows = kund_df[(kund_df["Ordernr"] == h_ord) & (kund_df["Zon"].astype(str).str.strip().str.upper() == "F")]
            mlist: list[str] = []
            if hib_zone_rows.empty:
                # ingen rad i zon F => saknar multi för denna order
                missing_multi_per_order[h_ord] = True
            else:
                missing_flag = True
                for _, zrow in hib_zone_rows.iterrows():
                    mval = str(zrow.get("Multi", "")).strip()
                    if mval:
                        mlist.append(mval)
                        missing_flag = False
                missing_multi_per_order[h_ord] = missing_flag
            hib_f_multi[h_ord] = mlist
        # Global unik mängd av alla multi-värden (icke-tomma) i zon F för denna kund
        multi_vals_global: set[str] = set()
        for mlist in hib_f_multi.values():
            for m in mlist:
                if m:
                    multi_vals_global.add(m)
        # Det finns en gemensam multi om mängden har exakt ett värde
        common_multi_exists = len(multi_vals_global) == 1
        # Om det finns en gemensam multi, extrahera den
        common_multi_value = next(iter(multi_vals_global)) if common_multi_exists else None
        # Generera rader
        for hib in hib_orders:
            h_row = hib["row"]
            h_ord = hib["ordernr"]
            # Beräkna uppdateringar
            ship_update = ""
            date_update = ""
            z_update = hib_zone_updates.get(h_ord, "")
            # Jämför sändningsnummer och orderdatum mot matchande butiksorder
            cur_ship = str(h_row["Sändningsnr"]).strip()
            cur_date = str(h_row["Orderdatum"]).strip()

            # Kundnamn för HIB‑ordern, används för att prioritera matchning mot samma butik
            hib_kundnamn = order_to_kundnamn.get(h_ord, "").strip().lower()

            # Försök hitta butiksorder som matchar både sändningsnummer och kundnamn
            def _store_kname(ordnr: str) -> str:
                return order_to_kundnamn.get(ordnr, "").strip().lower()

            # Kandidater med samma sändningsnummer och samma kundnamn
            # Använd alla butiksordrar (store_df) för att hitta matchning på sändningsnummer
            # oavsett status. Detta säkerställer att en HIB‑order som redan är kopplad till
            # en butik med ett avslutat orderstatus (>34) inte kopplas om till en annan butik
            # bara för att dess butik inte finns i valid_store_df.
            ship_kname_candidates = store_df[
                (store_df["Sändningsnr"].astype(str).str.strip() == cur_ship)
                & (store_df["Ordernr"].map(lambda x: _store_kname(x) == hib_kundnamn))
            ]
            if not ship_kname_candidates.empty:
                # Välj den tidigaste av de butiksorder som matchar både sändningsnummer och kundnamn
                candidate_row = _choose_earliest(ship_kname_candidates)
            else:
                # Annars matcha endast på sändningsnummer (oavsett kundnamn) i alla butiksordrar
                ship_candidates = store_df[store_df["Sändningsnr"].astype(str).str.strip() == cur_ship]
                if not ship_candidates.empty:
                    candidate_row = _choose_earliest(ship_candidates)
                else:
                    # Om ingen matchande sändningsnummer hittas används fallback‑butiken
                    candidate_row = fallback_store_row

            # Hämta referensdata från vald butiksorder
            ref_ship = str(candidate_row["Sändningsnr"]).strip()
            ref_date = str(candidate_row["Orderdatum"]).strip()

            # Om HIB‑orderns värde inte matchar referensen anges uppdatering
            if cur_ship != ref_ship:
                ship_update = ref_ship
            if cur_date != ref_date:
                date_update = ref_date
            # Bestäm multi‑uppdatering per order
            multi_update = ""
            if len(hib_orders) > 1:
                # saknar F‑zon eller multi för denna order
                if missing_multi_per_order.get(h_ord, False):
                    multi_update = "MULTI"
                else:
                    if common_multi_exists:
                        # det finns exakt en gemensam multi; kontrollera om denna order har samma värde
                        if set(hib_f_multi.get(h_ord, [])) != {common_multi_value}:
                            multi_update = "MULTI"
                    else:
                        # flera olika multi-värden existerar globalt; föreslå att enas på en multi
                        multi_update = "MULTI"
            ursprungsdatum = ""
            if "Ursprungsdatum" in ov.columns:
                udat_vals = kund_df.loc[kund_df["Ordernr"] == h_ord, "Ursprungsdatum"].dropna().astype(str).str.strip()
                if not udat_vals.empty:
                    ursprungsdatum = udat_vals.iloc[0]
            # Inkludera endast om någon kolumn behöver ändras
            if ship_update or date_update or z_update or multi_update:
                rows.append({
                    "ordernummer": h_ord,
                    "kundnamn": order_to_kundnamn.get(h_ord, ""),
                    "Ursprungsdatum": ursprungsdatum,
                    "Orderdatum": date_update,
                    "sändningsnummer": ship_update,
                    "Zon": z_update,
                    "Multi": multi_update
                })
    # Skapa DataFrame
    if not rows:
        return pd.DataFrame(columns=["ordernummer", "Ursprungsdatum", "Orderdatum", "sändningsnummer", "Zon", "Multi"])
    result_df = pd.DataFrame(rows)
    if result_df.empty:
        return result_df
    # Sortera efter kundnamn (A→Z) och sedan ordernummer för stabilitet
    # Detta gör att Excel-filen hamnar i alfabetisk ordning på kundnamn
    result_df = result_df.sort_values(by=["kundnamn", "ordernummer"]).reset_index(drop=True)
    # Placera kolumner i ordning: ordernr, kundnamn, ursprungsdatum, orderdatum, sändningsnummer, Zon, Multi
    cols = ["ordernummer", "kundnamn", "Ursprungsdatum", "Orderdatum", "sändningsnummer", "Zon", "Multi"]
    result_df = result_df[cols]
    return result_df


def compute_missed_departures(details_df: pd.DataFrame, overview_df: pd.DataFrame) -> pd.DataFrame:
    """
    Identifiera HIB‑ordrar som har orderrader med status > 34 och vars sändningsnummer inte matchar
    någon butiksorder för samma kund.  Returnerar ett DataFrame med kolumnerna:
      - ordernummer: HIB‑ordernummer.
      - kundnamn: Kundnamn om tillgängligt.
      - Missat: alltid "MISSAT SIN AVGÅNG" för dessa ordrar.
    """
    try:
        # Kopiera och städa kolumnnamn
        details = details_df.copy()
        overview = overview_df.copy()
        details.columns = [str(c).replace("\ufeff", "").strip() for c in details.columns]
        overview.columns = [str(c).replace("\ufeff", "").strip() for c in overview.columns]
        # Synonym‑mappning som i compute_hib_koppling
        synonyms = {
            "Ordernr": ["Ordernr", "Order nr", "Order number", "Ordernummer"],
            "Ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
            "Kund nr": ["Kund nr", "Kundnr", "Kundnummer", "Customer number", "Kund NR"],
            "Bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
            "Orderdatum": ["Orderdatum", "Order datum", "Order date", "Orderdate"],
            "Sändningsnr": [
                "Sändningsnr",
                "Sändnings nr",
                "Sändningsnummer",
                "Sendingsnr",
                "Sändnings number",
            ],
            "Zon": ["Zon", "Zone"],
            "Multi": ["Multi", "Multi nr", "Multinr", "Multi number"],
        }
        for canonical, syns in synonyms.items():
            if canonical not in overview.columns:
                for candidate in syns:
                    for col in list(overview.columns):
                        if col.strip().lower() == candidate.strip().lower():
                            overview.rename(columns={col: canonical}, inplace=True)
                            break
                    if canonical in overview.columns:
                        break
        # Kontrollera att nödvändiga kolumner finns
        required_overview_cols = {"Ordernr", "Ordertyp", "Kund nr", "Sändningsnr"}
        if any(c not in overview.columns for c in required_overview_cols):
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        # Normalisera ordertyp men låt indata-filtret styra vilka rader som ingår.
        ov = overview.copy()
        ov["Ordertyp"] = ov["Ordertyp"].astype(str).str.strip().str.upper()
        if ov.empty:
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        # Säkerställ att details har ordernr och status
        if "Order nr" not in details.columns or "Status" not in details.columns:
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        # Konvertera status till numeriskt
        def to_status_numeric(x):
            try:
                return int(float(str(x).strip()))
            except Exception:
                return 9999
        details["_status_num"] = details["Status"].apply(to_status_numeric)
        order_status_max = details.groupby("Order nr")["_status_num"].max().to_dict()
        # Mappning ordernummer -> kundnamn
        order_to_kundnamn: dict[str, str] = {}
        if "Order nr" in details.columns and "Kund.1" in details.columns:
            try:
                order_to_kundnamn = (
                    details.groupby("Order nr")["Kund.1"].first()
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .to_dict()
                )
            except Exception:
                order_to_kundnamn = {}
        rows: list[dict] = []
        # Gruppera efter kundnummer
        for kund_nr, kund_df in ov.groupby("Kund nr"):
            store_df = kund_df[kund_df["Ordertyp"] == "N"].copy()
            hib_df = kund_df[kund_df["Ordertyp"] == "HIB"].copy()
            ignored_hib_orders = _hib_orders_with_today_origin(kund_df)
            if ignored_hib_orders:
                hib_df = hib_df[~hib_df["Ordernr"].astype(str).str.strip().isin(ignored_hib_orders)].copy()
            if not store_df.empty:
                store_df = store_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
            if not hib_df.empty:
                hib_df = hib_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
            if store_df.empty or hib_df.empty:
                continue
            # Sändningsnummer för butikens ordrar
            store_ships: set[str] = set()
            for _, row in store_df.iterrows():
                ship = str(row.get("Sändningsnr", "")).strip()
                if ship:
                    store_ships.add(ship)
            for _, hib_row in hib_df.iterrows():
                h_ord = hib_row["Ordernr"]
                maxstatus = order_status_max.get(h_ord, 9999)
                # Intressanta HIB‑ordrar har status > 34
                if maxstatus <= 34:
                    continue
                cur_ship = str(hib_row.get("Sändningsnr", "")).strip()
                # Om sändningsnumret finns i butikernas sändningar är det inte en missad avgång
                if cur_ship and cur_ship in store_ships:
                    continue
                rows.append({
                    "ordernummer": h_ord,
                    "kundnamn": order_to_kundnamn.get(h_ord, ""),
                    "Missat": "MISSAT SIN AVGÅNG",
                })
        if not rows:
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        result = pd.DataFrame(rows)
        result = result.sort_values(by=["kundnamn", "ordernummer"]).reset_index(drop=True)
        return result
    except Exception:
        return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])

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

def _find_lyx_max_csv() -> Optional[Path]:
    candidates = [
        _seed_bufferpall_runtime_file("artikel_max.csv"),
        _bufferpall_resource_path("artikel_max.csv"),
        _bufferpall_source_dir() / "artikel_max.csv",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


ORDERSALDO_COLUMN_CANDIDATES: Dict[str, List[str]] = {
    "order": ["ordernr", "ordernummer", "order number", "order no", "orderid", "order"],
    "article": ["artikel", "artikelnr", "artikelnummer", "artnr", "sku", "item", "productcode"],
    "demand": ["beställt", "bestalld", "ordered", "orderqty", "qty", "quantity", "antal"],
    "pick": ["plock", "plocksaldo", "saldo", "available", "stock", "qtyavailable", "saldo autoplock"],
}

PAFYLLNADSPRIO_COLUMNS: List[str] = ["ALLA", "PRIO 1", "PRIO 2", "PRIO 3", "PRIO 4", "PRIO 5"]
LASTNINGSFONSTER_PRIO_COLUMNS: List[str] = ["PRIO", "LASTNINGSFÖNSTER"]
LASTNINGSFONSTER_UNKNOWN_SORT = pd.Timestamp("2262-04-11 23:47:00")
LASTNINGSFONSTER_UNKNOWN_LABEL = "Saknar lastningsfönster"


def _ordersaldo_norm(value: str) -> str:
    """Normalisera kolumnnamn för robust matchning."""
    txt = str(value).lower()
    txt = txt.replace("å", "a").replace("ä", "a").replace("ö", "o")
    txt = re.sub(r"[^a-z0-9]+", "", txt)
    return txt


def _ordersaldo_find_col(df: pd.DataFrame, candidates: List[str], used_cols: set[str]) -> Optional[str]:
    """Hitta kolumn via exakt/fuzzy match mot kandidater."""
    cols = [str(c) for c in df.columns]
    norm_cols = {col: _ordersaldo_norm(col) for col in cols}
    cand_norm = [_ordersaldo_norm(c) for c in candidates]
    for cand in cand_norm:
        for col, norm_col in norm_cols.items():
            if col in used_cols:
                continue
            if norm_col == cand:
                return col
    for cand in cand_norm:
        for col, norm_col in norm_cols.items():
            if col in used_cols:
                continue
            if cand and cand in norm_col:
                return col
    return None


def _find_ordersaldo_columns(
    df: pd.DataFrame,
    column_candidates: Optional[Dict[str, List[str]]] = None,
) -> Dict[str, Optional[str]]:
    candidates = column_candidates or ORDERSALDO_COLUMN_CANDIDATES
    used: set[str] = set()
    order_col = _ordersaldo_find_col(df, candidates["order"], used)
    if order_col:
        used.add(order_col)
    article_col = _ordersaldo_find_col(df, candidates["article"], used)
    if article_col:
        used.add(article_col)
    demand_col = _ordersaldo_find_col(df, candidates["demand"], used)
    if demand_col:
        used.add(demand_col)
    pick_col = _ordersaldo_find_col(df, candidates["pick"], used)
    return {
        "order": order_col,
        "article": article_col,
        "demand": demand_col,
        "pick": pick_col,
    }


def compute_ordersaldo_data(
    df: pd.DataFrame,
    utbest_map: Optional[Dict[str, float]] = None,
    column_names: Optional[Dict[str, Optional[str]]] = None,
) -> Tuple[list[str], pd.DataFrame]:
    """
    Beräkna ordersaldo-listor och underskottsdata per artikel från beställningslinjer.

    Returnerar (kompletta_ordrar, underskott_df) där underskott_df har index=artikelnummer
    och kolumnerna Total beställt, Tillgängligt saldo (Plock), Utbeställt, Underskott.
    """
    empty = pd.DataFrame(
        columns=["Total beställt", "Tillgängligt saldo (Plock)", "Utbeställt", "Underskott"]
    )
    if not isinstance(df, pd.DataFrame) or df.empty:
        return [], empty

    calc_df = _clean_columns(df.copy())
    cols = column_names or _find_ordersaldo_columns(calc_df)
    order_col = cols.get("order")
    article_col = cols.get("article")
    demand_col = cols.get("demand")
    pick_col = cols.get("pick")
    if not order_col or not article_col or not demand_col or not pick_col:
        raise KeyError("Hittar inte order-, artikel-, antal- eller plockkolumn i beställningsfilen.")

    calc_df[order_col] = calc_df[order_col].astype(str).str.strip()
    calc_df[article_col] = calc_df[article_col].astype(str).str.strip()
    calc_df[demand_col] = calc_df[demand_col].map(to_num).astype(float)
    calc_df[pick_col] = calc_df[pick_col].map(to_num).astype(float)

    calc_df["_enough_row"] = calc_df[pick_col] >= calc_df[demand_col]
    complete_mask = calc_df.groupby(order_col)["_enough_row"].all()
    complete_orders = sorted(complete_mask[complete_mask].index.astype(str).tolist())

    demand_by_art = calc_df.groupby(article_col)[demand_col].sum(min_count=1)
    stock_by_art = calc_df.groupby(article_col)[pick_col].max()
    holistic = pd.DataFrame({
        "Total beställt": demand_by_art,
        "Tillgängligt saldo (Plock)": stock_by_art,
    }).fillna(0)
    holistic.index = holistic.index.map(lambda value: str(value).strip())

    if utbest_map is None:
        utbest_map = {}
    holistic["Utbeställt"] = holistic.index.to_series().map(utbest_map).fillna(0.0)
    holistic["Underskott"] = (
        holistic["Total beställt"] + holistic["Utbeställt"] - holistic["Tillgängligt saldo (Plock)"]
    ).clip(lower=0)
    holistic_short = holistic[holistic["Underskott"] > 0].copy().sort_index()
    return complete_orders, holistic_short


def _build_article_max_map(max_df: pd.DataFrame) -> Dict[str, float]:
    max_df = _clean_columns(max_df.copy())
    max_art_col = find_col(max_df, ["artikelnummer", "artikel", "artnr", "art.nr", "sku"])
    max_val_col = find_col(max_df, ["max"])
    tmp = pd.DataFrame({
        "_art": _safe_str_series(max_df[max_art_col]),
        "_max": max_df[max_val_col].map(to_num),
    })
    tmp = tmp[tmp["_art"].ne("")].dropna(subset=["_max"])
    return tmp.drop_duplicates(subset="_art").set_index("_art")["_max"].to_dict()


def _classify_pafyllnadsprio(underskott: float, reference_value: float) -> Tuple[str, bool]:
    try:
        reference = float(reference_value)
    except Exception:
        reference = 0.0
    if pd.isna(reference) or reference <= 0:
        return "PRIO 5", True

    ratio = float(underskott) / reference
    if ratio <= 0.25:
        return "PRIO 1", False
    if ratio <= 0.40:
        return "PRIO 2", False
    if ratio <= 0.55:
        return "PRIO 3", False
    if ratio <= 0.70:
        return "PRIO 4", False
    return "PRIO 5", False


def _build_pafyllnadsprio_dataframe(groups: Dict[str, List[str]]) -> pd.DataFrame:
    max_len = max((len(values) for values in groups.values()), default=0)
    padded = {
        column: values + [""] * (max_len - len(values))
        for column, values in groups.items()
    }
    return pd.DataFrame(padded, columns=PAFYLLNADSPRIO_COLUMNS)


def build_pafyllnadsprio_report(shortage_df: pd.DataFrame, max_df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Bygg enkel fallback-rapport för Påfyllnadsprio utan lastningsfönster."""
    groups = {column: [] for column in PAFYLLNADSPRIO_COLUMNS}
    if not isinstance(shortage_df, pd.DataFrame) or shortage_df.empty:
        return pd.DataFrame(columns=PAFYLLNADSPRIO_COLUMNS), 0

    max_map = _build_article_max_map(max_df)
    missing_reference_count = 0
    work = shortage_df.copy()
    work.index = work.index.map(lambda value: str(value).strip())
    work = work[work.index != ""].sort_index()

    for article, row in work.iterrows():
        groups["ALLA"].append(article)
        reference_value = max_map.get(article, 0.0)
        prio, missing_reference = _classify_pafyllnadsprio(to_num(row.get("Underskott", 0.0)), reference_value)
        if missing_reference:
            missing_reference_count += 1
        groups[prio].append(article)

    return _build_pafyllnadsprio_dataframe(groups), missing_reference_count


def _combine_orderdatum_and_laststart(orderdatum_value: object, laststart_value: object) -> pd.Timestamp | pd.NaT:
    """Kombinera datum från Orderdatum med klockslag från Laststarttid."""
    order_ts = pd.to_datetime(orderdatum_value, errors="coerce", dayfirst=False)
    if pd.isna(order_ts):
        return pd.NaT
    match = re.search(r"(\d{1,2}):(\d{2})", str(laststart_value or "").strip())
    if not match:
        return pd.NaT
    hour = int(match.group(1))
    minute = int(match.group(2))
    try:
        return pd.Timestamp(
            year=int(order_ts.year),
            month=int(order_ts.month),
            day=int(order_ts.day),
            hour=hour,
            minute=minute,
        )
    except Exception:
        return pd.NaT


def _format_lastningsfonster_label(window_ts: pd.Timestamp) -> str:
    if pd.isna(window_ts):
        return LASTNINGSFONSTER_UNKNOWN_LABEL
    return pd.Timestamp(window_ts).strftime("%Y-%m-%d %H:%M")


def _prepare_lastningsfonster_overview(overview_df: pd.DataFrame) -> pd.DataFrame:
    """Normalisera orderöversikten till en order -> lastningsfönster-tabell."""
    overview = _clean_columns(overview_df.copy())
    order_col = find_col(overview, ["ordernr", "ordernummer", "order number", "order no", "orderid", "order"])
    orderdatum_col = find_col(overview, ["orderdatum", "order datum", "order date", "orderdate"])
    laststart_col = find_col(
        overview,
        ["laststarttid", "laststart tid", "laststart", "laststart time", "last start time"],
    )

    tmp = pd.DataFrame({
        "_order": overview[order_col].astype(str).str.strip(),
        "_orderdatum": overview[orderdatum_col],
        "_laststart": overview[laststart_col],
    })
    tmp = tmp[tmp["_order"] != ""].copy()
    if tmp.empty:
        return pd.DataFrame(columns=["_order", "_window_sort", "_window_label", "_prio", "_missing_window"])

    tmp["_window_sort"] = [
        _combine_orderdatum_and_laststart(orderdatum_value, laststart_value)
        for orderdatum_value, laststart_value in zip(tmp["_orderdatum"], tmp["_laststart"])
    ]

    rows: List[dict] = []
    for order, grp in tmp.groupby("_order", sort=True):
        valid = grp.dropna(subset=["_window_sort"]).sort_values("_window_sort")
        if not valid.empty:
            window_ts = pd.Timestamp(valid.iloc[0]["_window_sort"])
            rows.append({
                "_order": str(order).strip(),
                "_window_sort": window_ts,
                "_window_label": _format_lastningsfonster_label(window_ts),
                "_missing_window": False,
            })
        else:
            rows.append({
                "_order": str(order).strip(),
                "_window_sort": LASTNINGSFONSTER_UNKNOWN_SORT,
                "_window_label": LASTNINGSFONSTER_UNKNOWN_LABEL,
                "_missing_window": True,
            })

    out = pd.DataFrame(rows)
    valid_windows = sorted(
        {pd.Timestamp(value) for value in out.loc[~out["_missing_window"], "_window_sort"].tolist()}
    )
    prio_map: Dict[pd.Timestamp, str] = {}
    for idx, window_ts in enumerate(valid_windows):
        prio_map[window_ts] = f"PRIO {idx + 1}" if idx < 4 else "PRIO 5"
    out["_prio"] = out["_window_sort"].map(prio_map).fillna("PRIO 5")
    return out[["_order", "_window_sort", "_window_label", "_prio", "_missing_window"]]


def _build_lastningsfonster_prio_dataframe(overview_windows: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Bygg en enkel översikt över vilka lastningsfönster som motsvarar PRIO 1-4."""
    if not isinstance(overview_windows, pd.DataFrame) or overview_windows.empty:
        return pd.DataFrame(columns=LASTNINGSFONSTER_PRIO_COLUMNS), []

    prio_map = (
        overview_windows.loc[
            (~overview_windows["_missing_window"]) & (overview_windows["_prio"].isin(PAFYLLNADSPRIO_COLUMNS[1:5])),
            ["_window_sort", "_window_label", "_prio"],
        ]
        .drop_duplicates()
        .sort_values(["_window_sort", "_prio", "_window_label"])
    )
    if prio_map.empty:
        return pd.DataFrame(columns=LASTNINGSFONSTER_PRIO_COLUMNS), []

    rows: List[dict] = []
    log_lines: List[str] = []
    for _, row in prio_map.iterrows():
        prio = str(row["_prio"]).strip()
        label = str(row["_window_label"]).strip()
        rows.append({"PRIO": prio, "LASTNINGSFÖNSTER": label})
        log_lines.append(f"{prio} = {label}")

    return pd.DataFrame(rows, columns=LASTNINGSFONSTER_PRIO_COLUMNS), log_lines


def build_pafyllnadsprio_lastningsfonster_report(
    orders_df: pd.DataFrame,
    shortage_df: pd.DataFrame,
    overview_df: pd.DataFrame,
    max_df: pd.DataFrame,
    *,
    column_names: Optional[Dict[str, Optional[str]]] = None,
) -> Tuple[pd.DataFrame, set[tuple[int, int]], List[str], int, pd.DataFrame]:
    """Bygg Påfyllnadsprio med lastningsfönster där samma artikel kan ligga i flera PRIO-kolumner."""
    groups = {column: [] for column in PAFYLLNADSPRIO_COLUMNS}
    empty_window_map = pd.DataFrame(columns=LASTNINGSFONSTER_PRIO_COLUMNS)
    if not isinstance(shortage_df, pd.DataFrame) or shortage_df.empty:
        return pd.DataFrame(columns=PAFYLLNADSPRIO_COLUMNS), set(), [], 0, empty_window_map

    orders_work = _clean_columns(orders_df.copy())
    cols = column_names or _find_ordersaldo_columns(orders_work)
    order_col = cols.get("order")
    article_col = cols.get("article")
    demand_col = cols.get("demand")
    if not order_col or not article_col or not demand_col:
        raise KeyError("Hittar inte order-, artikel- eller antalskolumn i beställningsfilen.")

    overview_windows = _prepare_lastningsfonster_overview(overview_df)
    demand_rows = pd.DataFrame({
        "_order": orders_work[order_col].astype(str).str.strip(),
        "_article": orders_work[article_col].astype(str).str.strip(),
        "_qty": orders_work[demand_col].map(to_num).astype(float),
    })
    demand_rows = demand_rows[(demand_rows["_article"] != "") & (demand_rows["_qty"] > 0)].copy()
    if demand_rows.empty:
        return pd.DataFrame(columns=PAFYLLNADSPRIO_COLUMNS), set(), [], 0, empty_window_map

    demand_rows = demand_rows.merge(
        overview_windows,
        how="left",
        left_on="_order",
        right_on="_order",
    )
    demand_rows["_window_sort"] = pd.to_datetime(demand_rows["_window_sort"], errors="coerce")
    demand_rows["_window_sort"] = demand_rows["_window_sort"].fillna(LASTNINGSFONSTER_UNKNOWN_SORT)
    demand_rows["_window_label"] = demand_rows["_window_label"].fillna(LASTNINGSFONSTER_UNKNOWN_LABEL)
    demand_rows["_prio"] = demand_rows["_prio"].fillna("PRIO 5")
    demand_by_window = (
        demand_rows.groupby(["_article", "_window_sort", "_window_label", "_prio"], as_index=False)["_qty"]
        .sum()
        .sort_values(["_article", "_window_sort", "_window_label"])
    )
    window_map_df, window_map_log_lines = _build_lastningsfonster_prio_dataframe(overview_windows)

    max_map = _build_article_max_map(max_df)
    prio_counts: Dict[str, Dict[str, int]] = {prio: {} for prio in PAFYLLNADSPRIO_COLUMNS[1:]}
    log_lines: List[str] = []
    missing_reference_count = 0

    work = shortage_df.copy()
    work.index = work.index.map(lambda value: str(value).strip())
    work = work[work.index != ""].sort_index()

    for article, row in work.iterrows():
        groups["ALLA"].append(article)
        reference_value = max_map.get(article, 0.0)
        try:
            reference_float = float(reference_value)
        except Exception:
            reference_float = 0.0
        if pd.isna(reference_float) or reference_float <= 0:
            missing_reference_count += 1
            prio_counts["PRIO 5"][article] = max(prio_counts["PRIO 5"].get(article, 0), 1)
            log_lines.append(f"Artikel {article} saknar referensvärde och placerades i PRIO 5.")
            continue

        article_windows = demand_by_window[demand_by_window["_article"] == article].copy()
        if article_windows.empty:
            prio_counts["PRIO 5"][article] = max(prio_counts["PRIO 5"].get(article, 0), 1)
            log_lines.append(f"Artikel {article} saknar matchning mot lastningsfönster och placerades i PRIO 5.")
            continue

        available_start = float(to_num(row.get("Tillgängligt saldo (Plock)", 0.0))) - float(
            to_num(row.get("Utbeställt", 0.0))
        )
        cumulative_need = 0.0
        previous_total_pallets = 0
        event_found = False

        for _, window_row in article_windows.iterrows():
            qty = float(window_row["_qty"])
            if qty <= 0:
                continue
            cumulative_need += qty
            cumulative_shortage = max(0.0, cumulative_need - available_start)
            total_pallets = int(math.ceil(cumulative_shortage / reference_float)) if cumulative_shortage > 0 else 0
            new_pallets = max(0, total_pallets - previous_total_pallets)
            previous_total_pallets = total_pallets
            if new_pallets <= 0:
                continue

            event_found = True
            prio = str(window_row["_prio"])
            prio_counts[prio][article] = max(prio_counts[prio].get(article, 0), new_pallets)
            pall_text = "pall" if new_pallets == 1 else "pallar"
            log_lines.append(
                f"Artikel {article} behöver {new_pallets} {pall_text} i lastningsfönster "
                f"{window_row['_window_label']} ({prio})."
            )

        if not event_found:
            prio_counts["PRIO 5"][article] = max(prio_counts["PRIO 5"].get(article, 0), 1)
            log_lines.append(f"Artikel {article} saknar beräknat påfyllningsfönster och placerades i PRIO 5.")

    bold_cells: set[tuple[int, int]] = set()
    for col_idx, prio in enumerate(PAFYLLNADSPRIO_COLUMNS[1:], start=1):
        counts = prio_counts[prio]
        multi_articles = sorted(
            [article for article, pallet_count in counts.items() if pallet_count > 1],
            key=lambda article: (-counts[article], article),
        )
        single_articles = sorted(article for article, pallet_count in counts.items() if pallet_count <= 1)
        ordered_articles = multi_articles + single_articles
        groups[prio] = ordered_articles
        for row_idx, article in enumerate(ordered_articles):
            if counts.get(article, 0) > 1:
                bold_cells.add((row_idx, col_idx))

    if window_map_log_lines:
        log_lines.append("Lastningsfönster per prio:")
        log_lines.extend(window_map_log_lines)

    return _build_pafyllnadsprio_dataframe(groups), bold_cells, log_lines, missing_reference_count, window_map_df


# ----------------------------------------------------------------------
# Observations (crowdsourcad pallid-historik) + GitHub-sync
# ----------------------------------------------------------------------

OBSERVATIONS_FILENAME = "observations.csv.gz"
OBSERVATIONS_COLS = ["artikelnummer", "pallid", "antal"]
GITHUB_REPO = "EmirKadr/allokering"
GITHUB_OBS_BRANCH = "data/community-observations"
GITHUB_OBS_FILE = "lowfreqdata/buffertpall/observations.csv.gz"


def _observations_path() -> Path:
    """Returnera lokalt path for observations.csv.gz."""
    return _seed_bufferpall_runtime_file(OBSERVATIONS_FILENAME)


def _artikel_max_path() -> Path:
    return _seed_bufferpall_runtime_file("artikel_max.csv")


def _read_observations(path: Path) -> pd.DataFrame:
    if path.exists() and path.stat().st_size > 0:
        df = pd.read_csv(path, dtype=str)
        for col in OBSERVATIONS_COLS:
            if col not in df.columns:
                df[col] = ""
        return df[OBSERVATIONS_COLS]
    return pd.DataFrame(columns=OBSERVATIONS_COLS)


def _max_utan_outlier(grupp: pd.DataFrame) -> Tuple[float, str]:
    """Returnera (max, pallid) efter Tukey IQR-outlier-filter.

    grupp ska ha kolumner 'antal' och 'pallid' med unikt index.
    Filtret aktiveras bara nar gruppen har >2 pallar.
    """
    if len(grupp) > 2:
        q1, q3 = np.percentile(grupp["antal"], [25, 75])
        ovre = q3 + 1.5 * (q3 - q1)
        sub = grupp[grupp["antal"] <= ovre]
        if not sub.empty:
            row = sub.loc[sub["antal"].idxmax()]
            return float(row["antal"]), str(row["pallid"])
    row = grupp.loc[grupp["antal"].idxmax()]
    return float(row["antal"]), str(row["pallid"])


def _recompute_artikel_max(observations: pd.DataFrame, ut_path: Path) -> int:
    """Racka om artikel_max.csv fran observations. Returnerar antal artiklar."""
    if observations.empty:
        pd.DataFrame(columns=["artikelnummer", "max", "pallid"]).to_csv(
            ut_path, index=False, encoding="utf-8-sig"
        )
        return 0

    df = observations.copy()
    df["antal"] = pd.to_numeric(df["antal"], errors="coerce")
    df = df.dropna(subset=["artikelnummer", "antal", "pallid"])
    df["artikelnummer"] = df["artikelnummer"].astype(str).str.strip()
    df["pallid"] = df["pallid"].astype(str).str.strip()
    df = df.drop_duplicates(subset="pallid").reset_index(drop=True)

    rader = []
    for art, grupp in df.groupby("artikelnummer"):
        max_val, pall_id = _max_utan_outlier(grupp)
        rader.append({"artikelnummer": art, "max": max_val, "pallid": pall_id})

    pd.DataFrame(rader, columns=["artikelnummer", "max", "pallid"]).to_csv(
        ut_path, index=False, encoding="utf-8-sig"
    )
    return len(rader)


def update_observations_from_buffer(buffer_raw: pd.DataFrame) -> Tuple[int, pd.DataFrame]:
    """Lagg till nya status-30-pallid i observations.csv.gz och racka om artikel_max.csv.

    Returnerar (antal_nya, dataframe_med_endast_nya_rader).
    Endast pallar med Status == 30 sparas.
    """
    art_col = find_col(buffer_raw, BUFFER_SCHEMA["artikel"], required=False)
    qty_col = find_col(buffer_raw, BUFFER_SCHEMA["qty"], required=False)
    id_col = find_col(buffer_raw, BUFFER_SCHEMA["id"], required=False)
    status_col = find_col(buffer_raw, BUFFER_SCHEMA["status"], required=False)
    if not all([art_col, qty_col, id_col, status_col]):
        return 0, pd.DataFrame(columns=OBSERVATIONS_COLS)

    df = buffer_raw[[art_col, qty_col, id_col, status_col]].copy()
    df.columns = ["artikelnummer", "antal", "pallid", "status"]
    df["antal"] = pd.to_numeric(df["antal"], errors="coerce")
    df["status"] = pd.to_numeric(df["status"], errors="coerce")
    df = df.dropna(subset=["artikelnummer", "antal", "pallid", "status"])
    df = df[df["status"] == 30]
    df["artikelnummer"] = df["artikelnummer"].astype(str).str.strip()
    df["pallid"] = df["pallid"].astype(str).str.strip()
    df["antal"] = df["antal"].astype(int).astype(str)
    df = df[["artikelnummer", "pallid", "antal"]].drop_duplicates(subset="pallid")

    obs_path = _observations_path()
    befintliga = _read_observations(obs_path)
    befintliga_ids = set(befintliga["pallid"].astype(str))

    nya = df[~df["pallid"].isin(befintliga_ids)]
    if nya.empty:
        return 0, nya

    kombinerat = pd.concat([befintliga, nya], ignore_index=True)
    kombinerat.to_csv(obs_path, index=False, compression="gzip")
    _recompute_artikel_max(kombinerat, _artikel_max_path())
    return len(nya), nya


def _github_token_path() -> Path:
    return _app_config_path()


def _load_github_token() -> Optional[str]:
    p = _github_token_path()
    if not p.exists():
        return None
    try:
        # utf-8-sig hanterar UTF-8 med BOM (PowerShell Out-File -Encoding utf8 skriver BOM)
        cfg = json.loads(p.read_text(encoding="utf-8-sig"))
        token = cfg.get("github_token", "").strip()
        return token or None
    except Exception:
        return None


def _github_request(url: str, method: str = "GET", token: Optional[str] = None,
                    payload: Optional[dict] = None, timeout: int = 15) -> Tuple[int, dict]:
    headers = {"Accept": "application/vnd.github+json", "User-Agent": "allokering-app"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    if data:
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read()
            return resp.status, (json.loads(body) if body else {})
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read())
        except Exception:
            body = {}
        return e.code, body


def push_new_observations_to_github(nya: pd.DataFrame) -> bool:
    """Pusha nya observationer till GitHub som en sessions-CSV. Tyst no-op om token saknas."""
    if nya is None or nya.empty:
        return False
    token = _load_github_token()
    if not token:
        return False
    try:
        from io import BytesIO
        import gzip
        buf = BytesIO()
        with gzip.GzipFile(fileobj=buf, mode="wb") as gz:
            gz.write(nya.to_csv(index=False).encode("utf-8"))
        gz_bytes = buf.getvalue()
    except Exception:
        return False

    user = re.sub(r"[^A-Za-z0-9_-]", "_", os.environ.get("USERNAME") or "user")
    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    remote_name = f"lowfreqdata/buffertpall/observations_{user}_{ts}.csv.gz"
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{urllib.parse.quote(remote_name)}"
    payload = {
        "message": f"User observations {user} {ts}",
        "content": base64.b64encode(gz_bytes).decode("ascii"),
        "branch": GITHUB_OBS_BRANCH,
    }
    status, _ = _github_request(api_url, method="PUT", token=token, payload=payload)
    return 200 <= status < 300


def fetch_observations_from_github() -> Tuple[int, int]:
    """Tvavags-sync med GitHub master:
    1. Hamta nya rader fran master och merga in i lokal observations.csv.gz
    2. Hitta orphaned lokala pallid (sparade offline / push misslyckats) och push:a dem

    Returnerar (antal_hamtade, antal_pushade_orphaned).
    Tyst no-op pa natfel, JSON-fel eller saknade kolumner.
    """
    raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_OBS_BRANCH}/{GITHUB_OBS_FILE}"
    token = _load_github_token()
    headers = {"User-Agent": "allokering-app"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(raw_url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
    except Exception:
        return 0, 0

    from io import BytesIO
    try:
        remote = pd.read_csv(BytesIO(data), compression="gzip", dtype=str)
    except Exception:
        return 0, 0
    for col in OBSERVATIONS_COLS:
        if col not in remote.columns:
            return 0, 0
    remote = remote[OBSERVATIONS_COLS]

    obs_path = _observations_path()
    lokal = _read_observations(obs_path)
    remote_ids = set(remote["pallid"].astype(str))
    lokal_ids = set(lokal["pallid"].astype(str))

    # 1. Hamta nya rader fran master
    nya_fran_remote = remote[~remote["pallid"].astype(str).isin(lokal_ids)]
    n_hamtade = len(nya_fran_remote)
    if n_hamtade:
        kombinerat = pd.concat([lokal, nya_fran_remote], ignore_index=True)
        kombinerat.to_csv(obs_path, index=False, compression="gzip")
        _recompute_artikel_max(kombinerat, _artikel_max_path())
        lokal = kombinerat

    # 2. Hitta orphaned lokala (finns lokalt, inte pa master) och push:a dem
    orphaned = lokal[~lokal["pallid"].astype(str).isin(remote_ids)]
    n_pushade = 0
    if not orphaned.empty:
        try:
            if push_new_observations_to_github(orphaned):
                n_pushade = len(orphaned)
        except Exception:
            pass

    return n_hamtade, n_pushade


def utbest_per_article(saldo_df: pd.DataFrame) -> Dict[str, float]:
    """Summera 'utbeställt' per artikel från saldofilen.

    Returnerar tom dict om artikel- eller utbest-kolumn saknas.
    Inga bolagsfilter — alla rader summeras.
    """
    df = _clean_columns(saldo_df.copy())
    art_col = find_col(df, ["artikel", "artnr", "art.nr", "artikelnummer", "sku"], required=False)
    utbest_col = find_col(df, ["utbeställt", "utbestallt"], required=False)
    if not art_col or not utbest_col:
        return {}
    tmp = pd.DataFrame({
        "_art": _safe_str_series(df[art_col]),
        "_utbest": df[utbest_col].map(to_num).fillna(0.0),
    })
    return tmp.groupby("_art")["_utbest"].sum().to_dict()


def compute_lyx_articles(saldo_df: pd.DataFrame, max_df: pd.DataFrame) -> Tuple[list[str], int]:
    """
    Returnera artikelnummer där (plocksaldo + utbeställt) är högst 20 % av
    max buffertantalet. Returvärdet är (artikellista, antal filtrerade rader).
    """
    saldo_df = _clean_columns(saldo_df.copy())
    max_df = _clean_columns(max_df.copy())

    art_col = find_col(saldo_df, ["artikel", "artnr", "art.nr", "artikelnummer", "sku"])
    saldo_col = find_col(
        saldo_df,
        ["plocksaldo", "plock saldo", "plock-saldo", "tillgängligt plock", "tillgangligt plock", "plock"],
        required=False,
    )
    plats_col = find_col(
        saldo_df,
        ["plockplats", "huvudplock", "mainpick", "hyllplats", "bin", "location", "lagerplats"],
        required=False,
    )
    bolag_col = find_col(saldo_df, ["bolag", "company", "bolagskod"], required=False)
    utbest_col = find_col(saldo_df, ["utbeställt", "utbestallt"], required=False)

    if saldo_col is None:
        raise KeyError("Kunde inte hitta plocksaldo-kolumnen i saldofilen.")

    mask = pd.Series(True, index=saldo_df.index)
    if plats_col:
        mask &= _safe_str_series(saldo_df[plats_col]).ne("")
    if bolag_col:
        mask &= _safe_str_series(saldo_df[bolag_col]).str.upper() == "MG"

    saldo_filt = saldo_df[mask].copy()
    if saldo_filt.empty:
        return [], 0

    saldo_filt["_art"] = _safe_str_series(saldo_filt[art_col])
    saldo_filt["_saldo"] = saldo_filt[saldo_col].map(to_num).fillna(0)
    if utbest_col:
        saldo_filt["_utbest"] = saldo_filt[utbest_col].map(to_num).fillna(0)
    else:
        saldo_filt["_utbest"] = 0.0
    saldo_filt["_total"] = saldo_filt["_saldo"] + saldo_filt["_utbest"]

    max_map = _build_article_max_map(max_df)

    saldo_filt["_max"] = saldo_filt["_art"].map(max_map)
    lyx_mask = saldo_filt["_max"].notna() & (saldo_filt["_total"] <= saldo_filt["_max"] * 0.20)
    lyx_arts = sorted(saldo_filt.loc[lyx_mask, "_art"].unique().tolist())
    return lyx_arts, len(saldo_filt)

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

    mask7  = df["DatumNorm"] >= (today - pd.Timedelta(days=7))
    mask30 = df["DatumNorm"] >= (today - pd.Timedelta(days=30))
    mask90 = df["DatumNorm"] >= (today - pd.Timedelta(days=90))

    total7  = df.loc[mask7].groupby("Artikelnummer")["Plockat"].sum()
    total30 = df.loc[mask30].groupby("Artikelnummer")["Plockat"].sum()
    total90 = df.loc[mask90].groupby("Artikelnummer")["Plockat"].sum()

    positive = df[df["Plockat"] > 0]
    last_pick = positive.groupby("Artikelnummer")["DatumNorm"].max() if not positive.empty else pd.Series(dtype="datetime64[ns]")
    last_pick = last_pick.reindex(df["Artikelnummer"].unique())

    days_since = (today - last_pick).dt.days
    days_since = days_since.where(~days_since.isna(), other=pd.NA)

    sub90_pos = df.loc[mask90 & (df["Plockat"] > 0)]
    unique_days_90 = sub90_pos.groupby("Artikelnummer")["DatumNorm"].nunique()

    sub90 = df.loc[mask90].copy()
    zero_rows = (sub90.assign(IsZero=(sub90["Plockat"]==0))
                        .groupby(["Artikelnummer","DatumNorm"])["IsZero"].sum()
                        .rename("ZeroRows"))
    zero_avg = zero_rows.reset_index().groupby("Artikelnummer")["ZeroRows"].mean()
    zero_avg = zero_avg.reindex(df["Artikelnummer"].unique()).fillna(0.0)

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

    if "Artikel" in df_norm.columns:
        out = out.merge(df_norm[["Artikelnummer","Artikel"]].drop_duplicates(),
                        on="Artikelnummer", how="left")
    else:
        out["Artikel"] = out["Artikelnummer"]

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
        tmp = df[[art_col]].copy()
        tmp.columns = ["Artikel"]
        tmp["Ej Staplingsbar"] = ""
        return tmp.drop_duplicates(subset=["Artikel"]).reset_index(drop=True)
    tmp = df[[art_col, stap_col]].copy()
    tmp.columns = ["Artikel", "Ej Staplingsbar"]
    tmp["Artikel"] = tmp["Artikel"].astype(str).str.strip()
    tmp["Ej Staplingsbar"] = tmp["Ej Staplingsbar"].fillna("").astype(str).str.strip()
    return tmp.drop_duplicates(subset=["Artikel"]).reset_index(drop=True)


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

    df[zone_col] = df[zone_col].fillna("").astype(str).str.strip().str.upper()
    if stack_col:
        df[stack_col] = df[stack_col].fillna("").astype(str).str.strip().str.upper()
    else:
        df["_stack_tmp"] = ""
        stack_col = "_stack_tmp"
    df[palltyp_col] = df[palltyp_col].fillna("").astype(str).str.strip().str.upper()

    art_col_ps = None
    try:
        art_col_ps = find_col(df, ORDER_SCHEMA["artikel"], required=False, default=None)
    except Exception:
        art_col_ps = None
    groups = df.groupby([kund_col] if kund1_col is None else [kund_col, kund1_col])
    records: list[dict] = []
    import math
    for keys, sub in groups:
        if kund1_col is None:
            kund_val = keys
            kund1_val = ""
        else:
            kund_val, kund1_val = keys
        mask_bottom = (sub[zone_col] == "H") & ((sub[stack_col] == "N") | (sub[stack_col] == ""))
        B = int(mask_bottom.sum())
        rows_A = int((sub[zone_col] == "A").sum())
        if rows_A > 0:
            top_A = math.ceil(rows_A / 20.0)
        else:
            top_A = 0
        mask_topH = (sub[zone_col] == "H") & (sub[stack_col] == "Y") & (sub[palltyp_col] != "SJÖ")
        top_H = int(mask_topH.sum())
        mask_rf = sub[zone_col].isin({"R", "F"})
        if art_col_ps and art_col_ps in sub.columns:
            mask_rf = mask_rf & ~sub[art_col_ps].astype(str).str.strip().isin(RF_PALLPLATS_EXCLUDE_ARTICLES)
        rows_R = int(mask_rf.sum())
        if rows_R < 27:
            top_R = 0
        elif rows_R <= 96:
            top_R = 1
        elif rows_R <= 163:
            top_R = 2
        elif rows_R <= 204:
            top_R = 3
        else:
            top_R = 4
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
        mask_sjo = (sub[zone_col] == "H") & (sub[palltyp_col] == "SJÖ")
        S_rows = int(mask_sjo.sum())
        T = top_A + top_H + top_R + top_S
        half_sum = (B + T) / 2.0
        P_component = math.ceil(half_sum)
        max_val = T if T > P_component else P_component
        P = max_val + 2 * S_rows
        total_pallar = B + T + S_rows
        helpall_stapelbar = B
        helpall_ej_stapelbar = top_H
        sjo_pall = S_rows
        skrymme_pallar = top_S
        plockpall = top_A
        autostore_pallar = top_R
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


def _safe_str_series(s: pd.Series) -> pd.Series:
    """
    Returnera en strängserie där varje värde är trimmat och NaN ersätts med tom sträng.
    """
    return s.fillna("").astype(str).str.strip()


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
    if needed_units <= 0:
        return 0.0
    if not isinstance(buffer_df, pd.DataFrame) or buffer_df.empty:
        return np.nan
    df = buffer_df.copy()
    try:
        df.rename(columns=lambda c: str(c).replace("\ufeff", "").strip(), inplace=True)
    except Exception:
        pass
    try:
        art_col = find_col(df, BUFFER_SCHEMA["artikel"], required=True)
        qty_col = find_col(df, BUFFER_SCHEMA["qty"], required=True)
        dt_col = find_col(df, BUFFER_SCHEMA["dt"], required=False, default=None)
        status_col = find_col(df, BUFFER_SCHEMA["status"], required=False, default=None)
        id_col = find_col(df, BUFFER_SCHEMA["id"], required=False, default=None)
    except Exception:
        return np.nan
    sub = df.loc[_safe_str_series(df[art_col]) == str(article)].copy()
    if sub.empty:
        return 0.0
    if status_col and status_col in sub.columns:
        s = _safe_str_series(sub[status_col])
        s_num = pd.to_numeric(s.str.extract(r"(-?\d+)")[0], errors="coerce")
        allowed_str = {str(x) for x in REFILL_BUFFER_STATUSES}
        sub = sub[s.isin(allowed_str) | s_num.isin(REFILL_BUFFER_STATUSES)].copy()
        if sub.empty:
            return 0.0
    if exclude_source_ids:
        if id_col and id_col in sub.columns:
            sub["_source_id"] = _safe_str_series(sub[id_col])
        else:
            sub["_source_id"] = "SRC-" + sub.index.astype(str)
        sub = sub[~sub["_source_id"].isin(exclude_source_ids)].copy()
        if sub.empty:
            return 0.0
    sub["__qty__"] = _num_series(sub[qty_col])
    if dt_col and dt_col in sub.columns:
        sub = sub.sort_values(dt_col, kind="mergesort", na_position="last")
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
    if isinstance(saldo_norm_df, pd.DataFrame) and not saldo_norm_df.empty:
        orig_cols = [str(c).strip().lower() for c in saldo_norm_df.columns]
        has_robot_col = any("robot" == c for c in orig_cols)
        has_auto_col = any("saldo autoplock" in c for c in orig_cols)
        if not has_robot_col:
            missing.append("saldo")
            pr["Robot"] = "N"
            pr["Saldo i autoplock"] = 0.0
        else:
            s = saldo_norm_df.copy()
            if "Artikel" not in s.columns:
                for c in s.columns:
                    lc = str(c).strip().lower()
                    if lc in ("artikel", "artikelnummer", "sku", "artnr", "art.nr", "article"):
                        s = s.rename(columns={c: "Artikel"})
                        break
            if "Robot" not in s.columns:
                s["Robot"] = "N"
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
    pr["Behov i prognosen (antal styck)"] = _num_series(pr["Antal styck"])
    pr["Saldo i autoplock"] = _num_series(pr["Saldo i autoplock"])
    pr["Behov efter saldo"] = (pr["Behov i prognosen (antal styck)"] - pr["Saldo i autoplock"]).clip(lower=0)
    pr["Summa antal i ej inlagrade artiklar"] = 0.0
    shortage = pr["Behov efter saldo"].copy()
    if exclude_source_ids is None and isinstance(allocated_df, pd.DataFrame):
        exclude_source_ids = _collect_exclude_source_ids(allocated_df)
    if not exclude_source_ids:
        exclude_source_ids = None
    if isinstance(buffer_df, pd.DataFrame) and not buffer_df.empty:
        buf = buffer_df.copy()
        try:
            buf.rename(columns=lambda c: str(c).replace("\ufeff", "").strip(), inplace=True)
        except Exception:
            pass
        try:
            art_col = find_col(buf, BUFFER_SCHEMA["artikel"], required=True)
            qty_col = find_col(buf, BUFFER_SCHEMA["qty"], required=True)
            dt_col = find_col(buf, BUFFER_SCHEMA["dt"], required=False, default=None)
            status_col = find_col(buf, BUFFER_SCHEMA["status"], required=False, default=None)
            id_col = find_col(buf, BUFFER_SCHEMA["id"], required=False, default=None)
        except Exception:
            missing.append("buffert")
            pr["FIFO-baserad beräkning (antal pall)"] = np.nan
            pr["Buffertsaldo (status 29,30)"] = 0.0
        if status_col and status_col in buf.columns:
            s_str = _safe_str_series(buf[status_col])
            s_num = pd.to_numeric(s_str.str.extract(r"(-?\d+)")[0], errors="coerce")
            allowed_str = {str(x) for x in REFILL_BUFFER_STATUSES}
            mask_status = s_str.isin(allowed_str) | s_num.isin(REFILL_BUFFER_STATUSES)
            buf = buf.loc[mask_status].copy()
        if exclude_source_ids:
            if id_col and id_col in buf.columns:
                buf["_source_id"] = _safe_str_series(buf[id_col])
            else:
                buf["_source_id"] = "SRC-" + buf.index.astype(str)
            buf = buf[~buf["_source_id"].isin(exclude_source_ids)].copy()
        buf["__qty__"] = _num_series(buf[qty_col])
        prefix_dict: Dict[str, np.ndarray] = {}
        if dt_col and dt_col in buf.columns:
            buf = buf.sort_values([art_col, dt_col], kind="mergesort", na_position="last")
        for art, group in buf.groupby(buf[art_col]):
            qty_vals = group["__qty__"].to_numpy()
            if qty_vals.size == 0:
                continue
            prefix = np.cumsum(qty_vals)
            prefix_dict[str(art)] = prefix

        buffer_sum_series = buf.groupby(buf[art_col])["__qty__"].sum()
        buffer_sum_dict = {str(k): v for k, v in buffer_sum_series.items()}
        pr["Buffertsaldo (status 29,30)"] = pr["Artikelnummer"].map(lambda x: buffer_sum_dict.get(str(x), 0.0))
        def calc_pallar(art: Any, need: float) -> float:
            if need <= 0:
                return 0.0
            pref = prefix_dict.get(str(art))
            if pref is None:
                return 0.0
            idx = np.searchsorted(pref, float(need), side="left")
            if idx >= len(pref):
                return float(len(pref))
            else:
                return float(idx + 1)
        pr["FIFO-baserad beräkning (antal pall)"] = [calc_pallar(a, n) for a, n in zip(pr["Artikelnummer"], shortage)]
    else:
        missing.append("buffert")
        pr["FIFO-baserad beräkning (antal pall)"] = np.nan
        pr["Buffertsaldo (status 29,30)"] = 0.0
    pr = pr.loc[(pr["Robot"].astype(str).str.upper() == "Y") & (pr["Behov efter saldo"] > 0)].copy()
    out_cols = [
        "Artikelnummer",
        "Behov i prognosen (antal styck)",
        "Saldo i autoplock",
        "Behov efter saldo",
        "Buffertsaldo (status 29,30)",
        "FIFO-baserad beräkning (antal pall)",
    ]
    for c in out_cols:
        if c not in pr.columns:
            pr[c] = np.nan if c.startswith("FIFO") else 0.0
    report = pr[out_cols].reset_index(drop=True)
    if missing:
        notes: List[str] = []
        if "saldo" in missing:
            notes.append("Saldo saknas → Saldo i autoplock antas 0 (C=0, D=B).")
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
    if not isinstance(report_df, pd.DataFrame):
        report_df = pd.DataFrame()
    else:
        col_name = "FIFO-baserad beräkning (antal pall)"
        if col_name in report_df.columns:
            try:
                report_df = report_df.sort_values(by=col_name, ascending=False).reset_index(drop=True)
            except Exception:
                pass
    sheets["Prognos vs Autoplock"] = report_df
    return _open_df_in_excel(sheets, label="prognos_vs_autoplock")




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
    try:
        buff_type_col = find_col(buffer_raw, [
            "palltyp", "pall typ", "pall type"
        ], required=False, default=None)
    except Exception:
        buff_type_col = None

    _log(f"Order-kolumner: Artikel='{order_article_col}', Antal='{order_qty_col}', OrderId='{order_id_col}', Rad='{order_line_col}', Status='{order_status_col}'")
    _log(f"Buffert-kolumner: Artikel='{buff_article_col}', Antal='{buff_qty_col}', Lagerplats='{buff_loc_col}', Tid='{buff_dt_col}', ID='{buff_id_col}', Status='{buff_status_col}'")

    orders = orders_raw.copy()
    orders["_artikel"] = orders[order_article_col].astype(str).str.strip()
    orders["_qty"] = orders[order_qty_col].map(to_num).astype(float)
    orders["_order_id"] = orders[order_id_col].astype(str) if order_id_col and order_id_col in orders.columns else ""
    orders["_order_line"] = orders[order_line_col].astype(str) if order_line_col and order_line_col in orders.columns else orders.index.astype(str)

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

    buffer_df = buffer_raw.copy()
    buffer_df["_artikel"] = buffer_df[buff_article_col].astype(str).str.strip()
    buffer_df["_qty"] = buffer_df[buff_qty_col].map(to_num).astype(float)
    buffer_df["_loc"] = buffer_df[buff_loc_col].astype(str).str.strip()
    buffer_df["_received"] = smart_to_datetime(buffer_df[buff_dt_col]) if buff_dt_col and buff_dt_col in buffer_df.columns else pd.NaT
    buffer_df["_source_id"] = buffer_df[buff_id_col].astype(str) if buff_id_col and buff_id_col in buffer_df.columns else "SRC-" + buffer_df.index.astype(str)
    if buff_type_col and buff_type_col in buffer_df.columns:
        tmp_palltyp = buffer_df[buff_type_col].fillna("").astype(str).str.strip()
        buffer_df["_palltyp"] = tmp_palltyp.replace({"nan": "", "": ""})
    else:
        buffer_df["_palltyp"] = ""

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

    loc_upper = buffer_df["_loc"].str.upper()
    mask_exclude = loc_upper.str.startswith(INVALID_LOC_PREFIXES, na=False) | loc_upper.isin(INVALID_LOC_EXACT)
    excluded_count = int(mask_exclude.sum())
    if excluded_count:
        _log(f"Filtrerar bort {excluded_count} rad(er) från bufferten pga lagerplats-regler ({INVALID_LOC_PREFIXES}*, {', '.join(sorted(INVALID_LOC_EXACT))}).")
    buffer_df = buffer_df[~mask_exclude].copy()

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
            "palltyp": (r.get("_palltyp", "") if pd.notna(r.get("_palltyp", "")) else "")
        })

    bin_queues: Dict[str, Deque[dict]] = defaultdict(deque)
    for _, r in bins.iterrows():
        bin_queues[str(r["_artikel"]).strip()].append({
            "source_id": r["_source_id"],
            "qty": float(r["_qty"]),
            "loc": r["_loc"],
            "received": r["_received"],
            "palltyp": (r.get("_palltyp", "") if pd.notna(r.get("_palltyp", "")) else "")
        })

    allocated_rows: List[dict] = []
    near_miss_rows: List[dict] = []
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
                "Anledning": f"Pallen var ≤{int(NEAR_MISS_PCT * 100)}% större än återstående behov (kan ej brytas)",
                "Gäller (INSTEAD R/A)": None
            })

    for _, orow in orders.iterrows():
        art = str(orow["_artikel"]).strip()
        need = float(orow["_qty"])
        if need <= 0:
            continue

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

        any_mainpick = False
        if need > 0:
            sub = clone_row(orow)
            sub[order_qty_col] = need
            sub["Zon (beräknad)"] = "A"
            sub["Källtyp"] = "HUVUDPLOCK"
            sub["Källa"] = ""
            sub["Källplats"] = ""
            sub["Palltyp (matchad)"] = ""
            allocated_rows.append(sub)
            any_mainpick = True
            need = 0.0

        if not any_helpall and (any_autostore or any_mainpick):
            for r in near_miss_rows:
                if r["OrderID"] == str(orow["_order_id"]) and r["OrderRad"] == str(orow["_order_line"]):
                    r["Gäller (INSTEAD R/A)"] = True
        else:
            for r in near_miss_rows:
                if r["OrderID"] == str(orow["_order_id"]) and r["OrderRad"] == str(orow["_order_line"]):
                    r["Gäller (INSTEAD R/A)"] = False

    allocated_df = pd.DataFrame(allocated_rows)

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

    added_cols = ["Zon (beräknad)", "Källtyp", "Källa", "Källplats", "Palltyp (matchad)"]
    ordered_cols = [c for c in orders_raw.columns] + [c for c in added_cols if c not in orders_raw.columns]
    if not allocated_df.empty:
        allocated_df = allocated_df[ordered_cols]
    else:
        allocated_df = pd.DataFrame(columns=ordered_cols)

    near_miss_df = pd.DataFrame(near_miss_rows)
    return allocated_df, near_miss_df


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

    art_col_res = find_col(result, ORDER_SCHEMA["artikel"])
    qty_col_res = find_col(result, ORDER_SCHEMA["qty"])

    art_col_buf = find_col(buff, BUFFER_SCHEMA["artikel"])
    qty_col_buf = find_col(buff, BUFFER_SCHEMA["qty"])
    dt_col_buf  = find_col(buff, BUFFER_SCHEMA["dt"], required=False, default=None)
    id_col_buf  = find_col(buff, BUFFER_SCHEMA["id"], required=False, default=None)
    status_col_buf = find_col(buff, BUFFER_SCHEMA["status"], required=False, default=None)

    b = buff.copy()
    b["_artikel"] = b[art_col_buf].astype(str).str.strip()
    b["_qty"] = b[qty_col_buf].map(to_num).astype(float)
    b["_received"] = smart_to_datetime(b[dt_col_buf]) if dt_col_buf and dt_col_buf in b.columns else pd.NaT
    b["_source_id"] = b[id_col_buf].astype(str) if id_col_buf and id_col_buf in b.columns else "SRC-" + b.index.astype(str)

    if status_col_buf and status_col_buf in b.columns:
        _s = b[status_col_buf].astype(str).str.strip()
        _snum = pd.to_numeric(_s.str.extract(r"(-?\d+)")[0], errors="coerce")
        allowed_str = {str(x) for x in REFILL_BUFFER_STATUSES}
        b = b[_s.isin(allowed_str) | _snum.isin(REFILL_BUFFER_STATUSES)].copy()

    used_help_ids: set[str] = set()
    if "Källtyp" in result.columns and "Källa" in result.columns:
        used_help_ids = set(result[result["Källtyp"].astype(str) == "HELPALL"]["Källa"].dropna().astype(str).tolist())

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

    def fifo_for_art(art_key: str) -> pd.DataFrame:
        d = b[b["_artikel"] == art_key].copy()
        if not d.empty and used_help_ids:
            d = d[~d["_source_id"].astype(str).isin(used_help_ids)].copy()
        return d.sort_values("_received")

    hp_like = result[result.get("Källtyp", "").isin(["HUVUDPLOCK", "SKRYMMANDE", "HIB", "EHANDEL"])].copy()
    rows_hp: List[dict] = []
    if not hp_like.empty:
        hp_like["_zon"] = hp_like["Källtyp"].astype(str).map({"SKRYMMANDE": "S", "HIB": "F"}).fillna("A")
        needs = (hp_like
                 .assign(_art=hp_like[art_col_res].astype(str).str.strip(),
                         _qty=pd.to_numeric(hp_like[qty_col_res], errors="coerce").fillna(0.0))
                 .groupby(["_art", "_zon"], as_index=False)["_qty"].sum())

        for art_key, grp_art in needs.groupby("_art"):
            total_need = float(grp_art["_qty"].sum())
            if total_need <= 0:
                continue
            adjusted_total = max(0.0, round(total_need) - float(saldo_sum.get(art_key, 0.0)))

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
            diff = int(adjusted_total) - int(allocated_sum)
            if parts:
                parts[0][1] += diff

            fifo_df = fifo_for_art(art_key)
            tillgangligt = float(pd.to_numeric(fifo_df["_qty"], errors="coerce").sum()) if not fifo_df.empty else 0.0

            for zone, behov_int in parts:
                behov_int = int(max(0, behov_int))
                if behov_int <= 0:
                    continue  # 0-rad → bort
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


class SlideToggle(tk.Canvas):
    """Enkel iOS-liknande slide-toggle. Kör callback(bool) vid klick."""

    def __init__(self, parent, command=None, width=50, height=24, **kwargs):
        super().__init__(parent, width=width, height=height,
                         highlightthickness=0, **kwargs)
        self._on = False
        self._command = command
        self._w = width
        self._h = height
        self._draw()
        self.bind("<Button-1>", self._toggle)

    def _draw(self) -> None:
        self.delete("all")
        w, h = self._w, self._h
        r = h // 2
        bg = "#CC2222" if self._on else "#999999"
        # Bakgrundskapsel
        self.create_oval(0, 0, h, h, fill=bg, outline="")
        self.create_oval(w - h, 0, w, h, fill=bg, outline="")
        self.create_rectangle(r, 0, w - r, h, fill=bg, outline="")
        # Vit cirkel
        pad = 3
        cx = w - r if self._on else r
        self.create_oval(cx - r + pad, pad, cx + r - pad, h - pad, fill="white", outline="")

    def _toggle(self, _event=None) -> None:
        self._on = not self._on
        self._draw()
        if self._command:
            self._command(self._on)

    @property
    def value(self) -> bool:
        return self._on

    def set(self, state: bool) -> None:
        self._on = state
        self._draw()


class App(ttk.Frame):
    def __init__(self, master, enable_update_checks: bool = True, enable_analytics: bool = True):
        super().__init__(master)
        self.master = master
        self.enable_update_checks = enable_update_checks
        self.enable_analytics = enable_analytics
        self.pack(fill="both", expand=True)
        # Set up a default style for the application
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            # Om temat inte finns installerat, använd standardtemat
            pass
        style.configure("Accent.TButton", padding=10, foreground="white", background="#2D7FF9")
        style.map(
            "Accent.TButton",
            background=[
                ("disabled", "#6DAEFF"),
                ("pressed", "#1F6ED8"),
                ("active", "#3A8DFF"),
            ],
            foreground=[
                ("disabled", "#EAF3FF"),
                ("pressed", "white"),
                ("active", "white"),
            ],
        )
        style.configure("Green.TButton", padding=10, foreground="white", background="#28a745")
        style.configure("Warning.TButton", padding=10, foreground="#212529", background="#f0ad4e")
        style.map(
            "Warning.TButton",
            background=[
                ("disabled", "#f5d8a8"),
                ("pressed", "#d99834"),
                ("active", "#f4bc66"),
            ],
            foreground=[
                ("disabled", "#7a6441"),
                ("pressed", "#212529"),
                ("active", "#212529"),
            ],
        )

        # Dictionaries used to track status icons and associated StringVars for indatafiler.
        # Dessa måste initieras innan widgets skapas eftersom _create_widgets refererar
        # till dem när den sätter upp filvalsraderna.
        self.file_status_widgets: dict[str, tuple[tk.Label, tk.Button]] = {}
        self.file_vars: dict[str, tk.StringVar] = {}
        self.filter_options: dict[str, list[str]] = {"bolag": [], "ordertyp": []}
        self.filter_vars: dict[str, dict[str, tk.BooleanVar]] = {"bolag": {}, "ordertyp": {}}
        self.filter_column_candidates: dict[str, list[str]] = {
            "bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
            "ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
        }
        self.filter_titles: dict[str, str] = {"bolag": "Bolag", "ordertyp": "Ordertyp"}
        self.filter_null_token = "__NULL__"
        self.filter_null_label = "(Tomt/Null)"
        self._filter_group_frames: dict[str, ttk.LabelFrame] = {}
        self._filter_checkbuttons: dict[str, dict[str, ttk.Checkbutton]] = {"bolag": {}, "ordertyp": {}}
        self._filter_pairs: set[tuple[str, str]] = set()
        self.ordersaldo_list1_values: list[str] = []
        self.ordersaldo_list2_values: list[str] = []
        self.ordersaldo_column_candidates: dict[str, list[str]] = {
            key: values[:] for key, values in ORDERSALDO_COLUMN_CANDIDATES.items()
        }
        self.wms_expected_filenames: dict[str, str] = {
            "wms_receive": "v_ask_receive_log.csv",
            "wms_booking": "v_ask_booking_putaway.csv",
            "wms_buffert": "v_ask_article_buffertpallet.csv",
            "wms_trans": "v_ask_trans_log.csv",
            "wms_pick": "v_ask_pick_log_full.csv",
            "wms_correct": "v_ask_correct_log.csv",
        }
        self._wms_analyzer_cls = None
        self._action_requirements: dict[ttk.Button, list[tuple[str, str]]] = {}
        self._action_missing_files: dict[ttk.Button, list[str]] = {}
        self._open_button_hints: dict[ttk.Button, str] = {}
        self._hover_tooltip: Optional[tk.Toplevel] = None
        self._hover_tooltip_label: Optional[tk.Label] = None
        self._update_check_in_progress = False
        self._update_download_in_progress = False
        self._update_cancel_event: Optional[threading.Event] = None
        self._update_progress_window: Optional[tk.Toplevel] = None
        self._update_progress_value = tk.IntVar(value=0)
        self._update_status_var = tk.StringVar(value="")
        self._help_mode: str = ""  # "" | "enkel" | "avancerat"
        self._help_overlay: Optional[tk.Toplevel] = None
        self._help_popup: Optional[tk.Toplevel] = None
        self._help_bar: Optional[tk.Toplevel] = None
        self._help_toggle: Optional[SlideToggle] = None
        self._HELP_REGISTRY: dict[int, str] = {}
        self._help_widget_map: dict[int, tk.Widget] = {}
        self._analytics_settings: dict = {}
        self._analytics_client = AnalyticsClient({"active": False, "reason": "Analytics inte initierat."})
        self._session_started_at = time.monotonic()

        self._initialize_analytics()
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

        self._setup_menu()
        self._set_window_icon()

        # Build the GUI widgets
        self._create_widgets()
        # Initialize optional campaign DataFrames
        self._campaign_norm: Optional[pd.DataFrame] = None
        self._campaign_raw: Optional[pd.DataFrame] = None
        # Uppdatera statusikonerna initialt så att rätt symboler visas
        try:
            self.update_file_status_icons()
        except Exception:
            pass
        self._track_event(
            "app_started",
            automatic_update_checks=self._automatic_update_checks_enabled(),
        )
        self._schedule_update_check()
        threading.Thread(target=fetch_observations_from_github, daemon=True).start()

    def _initialize_analytics(self) -> None:
        self._rebuild_analytics_client()

    def _rebuild_analytics_client(self) -> None:
        try:
            self._analytics_client.shutdown(timeout=1.0)
        except Exception:
            pass
        self._analytics_settings = _resolve_analytics_settings(enable_analytics=self.enable_analytics)
        self._analytics_client = AnalyticsClient(self._analytics_settings)

    def _track_event(self, event: str, **properties) -> None:
        try:
            self._analytics_client.capture(event, properties)
        except Exception:
            pass

    def _track_feature(self, feature: str, action: str, **properties) -> None:
        payload = {"feature": feature, "action": action}
        payload.update(properties)
        self._track_event("feature_usage", **payload)

    def _on_close(self) -> None:
        try:
            session_seconds = int(max(0, round(time.monotonic() - self._session_started_at)))
            self._track_event("app_closed", session_seconds=session_seconds)
            self._analytics_client.shutdown(timeout=1.5)
        except Exception:
            pass
        self.master.destroy()

    def _log(self, msg: str, level: str = "info") -> None:
        logprintln(self.log, msg)

    def _setup_menu(self) -> None:
        menu = tk.Menu(self.master)
        help_menu = tk.Menu(menu, tearoff=0)
        help_menu.add_command(label="Hjälpläge  (klicka på valfri del av appen)",
                              command=self._toggle_help_mode)
        help_menu.add_separator()
        help_menu.add_command(
            label="Sök efter uppdateringar",
            command=lambda: self._check_for_updates(manual=True),
        )
        help_menu.add_command(
            label="Öppna releasesida",
            command=lambda: webbrowser.open(GITHUB_RELEASES_URL),
        )
        help_menu.add_separator()
        help_menu.add_command(label=f"Om {APP_NAME}", command=self._show_about_dialog)
        menu.add_cascade(label="Hjälp", menu=help_menu)
        self.master.configure(menu=menu)

    def _set_window_icon(self) -> None:
        candidates = [
            _resource_path("app.ico"),
            Path(__file__).resolve().parent / "packaging" / "windows" / "app.ico",
        ]
        for icon_path in candidates:
            if not icon_path.exists():
                continue
            try:
                self.master.iconbitmap(default=str(icon_path))
                return
            except Exception:
                continue

    def _show_about_dialog(self) -> None:
        messagebox.showinfo(APP_NAME, f"{APP_NAME}\nVersion {APP_VERSION}")

    def _automatic_update_checks_enabled(self) -> bool:
        return self.enable_update_checks and os.environ.get(UPDATE_DISABLED_ENV) != "1"

    def _schedule_update_check(self) -> None:
        if self._automatic_update_checks_enabled():
            self.after(2500, lambda: self._check_for_updates(manual=False))

    def _check_for_updates(self, manual: bool = False) -> None:
        if self._update_check_in_progress:
            if manual:
                messagebox.showinfo(APP_NAME, "Söker redan efter uppdateringar.")
            return

        self._update_check_in_progress = True
        threading.Thread(
            target=self._check_for_updates_worker,
            args=(manual,),
            daemon=True,
        ).start()

    def _check_for_updates_worker(self, manual: bool) -> None:
        info: Optional[UpdateInfo] = None
        error: Optional[str] = None
        try:
            info = check_for_update(current_version=APP_VERSION)
        except Exception as exc:
            error = str(exc)
        self.after(
            0,
            lambda: self._finish_update_check(manual=manual, info=info, error=error),
        )

    def _finish_update_check(
        self,
        *,
        manual: bool,
        info: Optional[UpdateInfo],
        error: Optional[str],
    ) -> None:
        self._update_check_in_progress = False
        if error:
            if manual:
                messagebox.showwarning(
                    APP_NAME,
                    f"Kunde inte söka efter uppdatering.\n\n{error}",
                )
            return
        if info is None:
            if manual:
                messagebox.showinfo(
                    APP_NAME,
                    f"Du kör senaste versionen av {APP_NAME}.",
                )
            return
        self._on_update_available(info, manual=manual)

    def _on_update_available(self, info: UpdateInfo, manual: bool) -> None:
        if not info.installer_url:
            open_release = messagebox.askyesno(
                "Uppdatering finns",
                (
                    f"Version {info.version} finns tillgänglig, men releasen saknar "
                    "en Setup.exe-fil.\n\nVill du öppna releasesidan?"
                ),
            )
            if open_release:
                webbrowser.open(info.release_url)
            return

        install_now = messagebox.askyesno(
            "Uppdatering finns",
            (
                f"Version {info.version} finns tillgänglig.\n\n"
                "Vill du ladda ner och installera uppdateringen nu? "
                "Appen stängs automatiskt medan uppdateringen installeras."
            ),
        )
        if install_now:
            self._download_update(info)
        elif manual:
            self._log("Uppdatering hittad men användaren avstod just nu.")

    def _download_update(self, info: UpdateInfo) -> None:
        if self._update_download_in_progress:
            messagebox.showinfo(APP_NAME, "Uppdateringen laddas redan ner.")
            return

        self._update_download_in_progress = True
        self._update_cancel_event = threading.Event()
        self._update_progress_value.set(0)
        self._update_status_var.set("Laddar ner uppdatering...")
        self._show_update_progress()

        target_dir = Path(tempfile.gettempdir()) / APP_NAME / "updates"
        threading.Thread(
            target=self._download_update_worker,
            args=(info, target_dir, self._update_cancel_event),
            daemon=True,
        ).start()

    def _show_update_progress(self) -> None:
        window = tk.Toplevel(self.master)
        window.title("Uppdatering")
        window.transient(self.master)
        window.resizable(False, False)
        window.grab_set()
        window.protocol("WM_DELETE_WINDOW", self._cancel_update_download)

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, textvariable=self._update_status_var).pack(anchor="w")
        ttk.Progressbar(
            frame,
            maximum=100,
            variable=self._update_progress_value,
            length=320,
            mode="determinate",
        ).pack(fill="x", pady=(12, 8))
        ttk.Button(frame, text="Avbryt", command=self._cancel_update_download).pack(
            anchor="e"
        )
        self._update_progress_window = window

    def _cancel_update_download(self) -> None:
        if self._update_cancel_event is not None:
            self._update_cancel_event.set()
        self._update_status_var.set("Avbryter nedladdning...")

    def _download_update_worker(
        self,
        info: UpdateInfo,
        target_dir: Path,
        cancel_event: threading.Event,
    ) -> None:
        try:
            installer_path = download_update_installer(
                info,
                target_dir=target_dir,
                progress_cb=lambda value: self.after(
                    0, lambda value=value: self._update_progress_value.set(value)
                ),
                stop_flag=cancel_event.is_set,
            )
        except Exception as exc:
            cancelled = cancel_event.is_set()
            error_message = str(exc)
            self.after(
                0,
                lambda error_message=error_message, cancelled=cancelled: self._finish_update_download(
                    installer_path=None,
                    error=error_message,
                    cancelled=cancelled,
                ),
            )
            return

        self.after(
            0,
            lambda: self._finish_update_download(
                installer_path=installer_path,
                error=None,
                cancelled=False,
            ),
        )

    def _finish_update_download(
        self,
        *,
        installer_path: Optional[Path],
        error: Optional[str],
        cancelled: bool,
    ) -> None:
        self._close_update_progress()
        self._update_download_in_progress = False
        self._update_cancel_event = None

        if cancelled:
            messagebox.showinfo(APP_NAME, "Nedladdningen avbröts.")
            return
        if error:
            messagebox.showwarning(
                APP_NAME,
                f"Kunde inte ladda ner uppdateringen.\n\n{error}",
            )
            return
        if installer_path is None:
            messagebox.showwarning(APP_NAME, "Installeraren kunde inte hittas.")
            return

        if self._start_silent_update(installer_path):
            self.master.after(150, self.master.destroy)

    def _close_update_progress(self) -> None:
        if self._update_progress_window is not None:
            try:
                self._update_progress_window.grab_release()
            except Exception:
                pass
            self._update_progress_window.destroy()
            self._update_progress_window = None

    def _start_silent_update(self, installer_path: Path) -> bool:
        try:
            subprocess.Popen([str(installer_path), *SILENT_UPDATE_ARGS])
        except Exception as exc:
            messagebox.showerror(
                APP_NAME,
                f"Kunde inte starta installeraren:\n{installer_path}\n\n{exc}",
            )
            return False
        return True

    def _create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.columnconfigure(2, weight=0)
        self.columnconfigure(3, weight=0)
        indata_frame = ttk.LabelFrame(self, text="Indatafiler")
        indata_frame.grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=8)
        # Keep label/status/remove tightly grouped on the left.
        indata_frame.columnconfigure(0, minsize=280)
        indata_frame.columnconfigure(1, minsize=120)
        # Row for Beställningslinjer (CSV)
        ttk.Label(indata_frame, text="Beställningslinjer (CSV):").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.orders_var = tk.StringVar()
        # Use tk.Label for status so we can control background and font for better visibility
        status_orders = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_orders.grid(row=0, column=1, sticky="w", padx=4)
        remove_orders = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("orders"),
                                   fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                   relief="raised", width=2)
        remove_orders.grid(row=0, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["orders"] = (status_orders, remove_orders)
        self.file_vars["orders"] = self.orders_var
        # Row for Buffertpallar (CSV)
        ttk.Label(indata_frame, text="Buffertpallar (CSV):").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.buffer_var = tk.StringVar()
        self._last_observations_path: Optional[str] = None
        self.buffer_var.trace_add("write", self._on_buffer_var_changed)
        status_buffer = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_buffer.grid(row=1, column=1, sticky="w", padx=4)
        remove_buffer = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("buffer"),
                                  fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                  relief="raised", width=2)
        remove_buffer.grid(row=1, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["buffer"] = (status_buffer, remove_buffer)
        self.file_vars["buffer"] = self.buffer_var
        # Row for Saldo inkl. automation (CSV)
        ttk.Label(indata_frame, text="Saldo inkl. automation (CSV):").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.automation_var = tk.StringVar()
        status_automation = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_automation.grid(row=2, column=1, sticky="w", padx=4)
        remove_automation = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("automation"),
                                      fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                      relief="raised", width=2)
        remove_automation.grid(row=2, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["automation"] = (status_automation, remove_automation)
        self.file_vars["automation"] = self.automation_var
        # Row for Item option (CSV)
        ttk.Label(indata_frame, text="Item option (CSV):").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.item_var = tk.StringVar()
        status_item = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_item.grid(row=3, column=1, sticky="w", padx=4)
        remove_item = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("item"),
                                fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                relief="raised", width=2)
        remove_item.grid(row=3, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["item"] = (status_item, remove_item)
        self.file_vars["item"] = self.item_var

        # Row for Orderöversikt (CSV)
        ttk.Label(indata_frame, text="Orderöversikt (CSV):").grid(row=4, column=0, sticky="w", padx=4, pady=4)
        self.overview_var = tk.StringVar()
        status_overview = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_overview.grid(row=4, column=1, sticky="w", padx=4)
        remove_overview = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("overview"),
                                    fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                    relief="raised", width=2)
        remove_overview.grid(row=4, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["overview"] = (status_overview, remove_overview)
        self.file_vars["overview"] = self.overview_var

        # Row for Dispatchpallar (CSV)
        ttk.Label(indata_frame, text="Dispatchpallar (CSV):").grid(row=5, column=0, sticky="w", padx=4, pady=4)
        self.dispatch_var = tk.StringVar()
        status_dispatch = tk.Label(
            indata_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_dispatch.grid(row=5, column=1, sticky="w", padx=4)
        remove_dispatch = tk.Button(indata_frame, text="✗", command=lambda: self.clear_file("dispatch"),
                                   fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                   relief="raised", width=2)
        remove_dispatch.grid(row=5, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["dispatch"] = (status_dispatch, remove_dispatch)
        self.file_vars["dispatch"] = self.dispatch_var

        prog_frame = ttk.LabelFrame(self, text="Prognos / Kampanj")
        prog_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=8)
        # Match horizontal coordinates with indata_frame.
        prog_frame.columnconfigure(0, minsize=280)
        prog_frame.columnconfigure(1, minsize=120)
        ttk.Label(prog_frame, text="Prognos (XLSX):").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.prognos_var = tk.StringVar()
        status_prognos = tk.Label(
            prog_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_prognos.grid(row=0, column=1, sticky="w", padx=4)
        remove_prognos = tk.Button(prog_frame, text="✗", command=lambda: self.clear_file("prognos"),
                                   fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                   relief="raised", width=2)
        remove_prognos.grid(row=0, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["prognos"] = (status_prognos, remove_prognos)
        self.file_vars["prognos"] = self.prognos_var
        ttk.Label(prog_frame, text="Kampanjvolymer (XLSX):").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.campaign_var = tk.StringVar()
        status_campaign = tk.Label(
            prog_frame,
            text="Ej fil",
            fg="white",
            bg="#6c757d",
            width=10,
            anchor="w",
            font=("Arial", 11, "bold"),
        )
        status_campaign.grid(row=1, column=1, sticky="w", padx=4)
        remove_campaign = tk.Button(prog_frame, text="✗", command=lambda: self.clear_file("campaign"),
                                    fg="white", bg="#dc3545", activebackground="#c82333", activeforeground="white",
                                    relief="raised", width=2)
        remove_campaign.grid(row=1, column=2, sticky="w", padx=(4, 0))
        self.file_status_widgets["campaign"] = (status_campaign, remove_campaign)
        self.file_vars["campaign"] = self.campaign_var

        # Klick på valfri statusruta öppnar gemensam filväljare (multi-select).
        for status_lbl, _ in self.file_status_widgets.values():
            status_lbl.bind("<Button-1>", self.open_files_dialog)

        # Dynamiska värdefilter (Bolag/Ordertyp), nära indatafilerna till vänster.
        self.value_filter_frame = ttk.LabelFrame(self, text="Filtrering")
        self.value_filter_frame.grid(row=2, column=0, columnspan=1, sticky="w", padx=8, pady=(0, 8))
        self.value_filter_frame.grid_remove()
        for idx, filter_key in enumerate(("bolag", "ordertyp")):
            group_frame = ttk.LabelFrame(self.value_filter_frame, text=self.filter_titles[filter_key])
            group_frame.grid(row=0, column=idx, sticky="nw", padx=(0, 12), pady=4)
            group_frame.grid_remove()
            self._filter_group_frames[filter_key] = group_frame

        # DEMO-yta: Eftersok (WMS) med inmatning och egna indatafiler.
        self.demo_frame = ttk.LabelFrame(self, text="")
        self.demo_frame.grid(row=0, column=2, rowspan=2, sticky="nw", padx=(0, 8), pady=8)
        self.demo_frame.columnconfigure(0, minsize=160)
        self.demo_frame.columnconfigure(1, minsize=120)
        tk.Label(
            self.demo_frame,
            text="DEMO",
            fg="#28a745",
            font=("Arial", 11, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=4, pady=(4, 2))

        ttk.Label(self.demo_frame, text="Inköpsnummer:").grid(row=1, column=0, sticky="w", padx=4, pady=(2, 2))
        self.eftersok_purchase_var = tk.StringVar()
        self.eftersok_purchase_entry = ttk.Entry(self.demo_frame, textvariable=self.eftersok_purchase_var, width=16)
        self.eftersok_purchase_entry.grid(row=1, column=1, sticky="w", padx=4, pady=(2, 2))

        ttk.Label(self.demo_frame, text="Artikelnummer:").grid(row=2, column=0, sticky="w", padx=4, pady=(2, 4))
        self.eftersok_article_var = tk.StringVar()
        self.eftersok_article_entry = ttk.Entry(self.demo_frame, textvariable=self.eftersok_article_var, width=16)
        self.eftersok_article_entry.grid(row=2, column=1, sticky="w", padx=4, pady=(2, 4))

        self.eftersok_purchase_var.trace_add("write", self._on_eftersok_input_changed)
        self.eftersok_article_var.trace_add("write", self._on_eftersok_input_changed)

        def _add_demo_file_row(row_idx: int, file_key: str, label_text: str, var_obj: tk.StringVar) -> None:
            ttk.Label(self.demo_frame, text=label_text).grid(row=row_idx, column=0, sticky="w", padx=4, pady=2)
            status_lbl = tk.Label(
                self.demo_frame,
                text="Ej fil",
                fg="white",
                bg="#6c757d",
                width=10,
                anchor="w",
                font=("Arial", 11, "bold"),
            )
            status_lbl.grid(row=row_idx, column=1, sticky="w", padx=4)
            status_lbl.bind("<Button-1>", self.open_files_dialog)
            remove_btn = tk.Button(
                self.demo_frame,
                text="✗",
                command=lambda fk=file_key: self.clear_file(fk),
                fg="white",
                bg="#dc3545",
                activebackground="#c82333",
                activeforeground="white",
                relief="raised",
                width=2,
            )
            remove_btn.grid(row=row_idx, column=2, sticky="w", padx=(4, 0))
            self.file_status_widgets[file_key] = (status_lbl, remove_btn)
            self.file_vars[file_key] = var_obj

        self.wms_receive_var = tk.StringVar()
        self.wms_booking_var = tk.StringVar()
        self.wms_trans_var = tk.StringVar()
        self.wms_pick_var = tk.StringVar()
        self.wms_correct_var = tk.StringVar()

        _add_demo_file_row(3, "wms_receive", "Mottagningslogg (CSV):", self.wms_receive_var)
        _add_demo_file_row(4, "wms_booking", "Ej inlagrade (CSV):", self.wms_booking_var)
        _add_demo_file_row(5, "wms_trans", "Translogg (CSV):", self.wms_trans_var)
        _add_demo_file_row(6, "wms_pick", "Plocklogg (CSV):", self.wms_pick_var)
        _add_demo_file_row(7, "wms_correct", "Saldojustering (CSV):", self.wms_correct_var)

        # 2000-tal-verktyg på ytan där filter tidigare låg.
        self.split2000_frame = ttk.LabelFrame(self, text="")
        self.split2000_frame.grid(row=0, column=3, rowspan=3, sticky="nw", padx=(8, 8), pady=8)
        ttk.Label(self.split2000_frame, text="Klistra in värden (en per rad):").grid(row=0, column=0, sticky="w", padx=6, pady=(6, 4))
        self.split_input_text = scrolledtext.ScrolledText(self.split2000_frame, width=24, height=12)
        self.split_input_text.grid(row=1, column=0, sticky="nw", padx=6, pady=(0, 6))
        split_bottom = ttk.Frame(self.split2000_frame)
        split_bottom.grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))
        ttk.Label(split_bottom, text="Antal per kolumn:").pack(side="left")
        self.split_chunk_var = tk.StringVar(value="2000")
        ttk.Entry(split_bottom, textvariable=self.split_chunk_var, width=8).pack(side="left", padx=(6, 10))
        ttk.Button(split_bottom, text="Öppna i Excel direkt", command=self.open_chunked_values_in_excel).pack(side="left")

        # Placera run-knappar i ett eget ram för att kunna ha flera knappar bredvid varandra
        run_frame = ttk.Frame(self)
        run_frame.grid(row=3, column=0, columnspan=4, sticky="w", pady=10)
        self.run_btn = ttk.Button(run_frame, text="Kör allokering", command=self.run_allocation, style="Accent.TButton")
        self.run_btn.pack(side="left", padx=4)
        # Knapp för HIB‑koppling
        self.koppla_btn = ttk.Button(run_frame, text="Kör HIB‑koppling", command=self.run_koppla, style="Accent.TButton")
        self.koppla_btn.pack(side="left", padx=4)
        # Knapp för kontroll av orderöversikt (sändningsnr vs kunder/transportörer)
        self.overview_check_btn = ttk.Button(run_frame, text="Kontrollera orderöversikt", command=self.run_overview_check, style="Accent.TButton")
        self.overview_check_btn.pack(side="left", padx=4)
        # Knapp för kontroll av dispatchpallar (ordernr och sändningsnr)
        self.dispatch_check_btn = ttk.Button(run_frame, text="Kontrollera dispatchpallar", command=self.run_dispatch_check, style="Accent.TButton")
        self.dispatch_check_btn.pack(side="left", padx=4)
        self.eftersok_btn = ttk.Button(
            run_frame,
            text="Eftersök",
            command=self.run_eftersok,
            style="Accent.TButton",
            state="disabled",
        )
        self.eftersok_btn.pack(side="left", padx=4)

        # Egna knappar för listor, på separat rad (likt Chromium-layouten).
        ordersaldo_frame = ttk.Frame(self)
        ordersaldo_frame.grid(row=4, column=0, columnspan=4, sticky="w", pady=(0, 10))
        self.ordersaldo_copy_list1_btn = ttk.Button(
            ordersaldo_frame,
            text="Kompletta ordrar",
            command=self.copy_ordersaldo_list1,
            style="Warning.TButton",
            state="disabled",
        )
        self.ordersaldo_copy_list1_btn.pack(side="left", padx=4)
        self.ordersaldo_copy_list2_btn = ttk.Button(
            ordersaldo_frame,
            text="Påfyllningsbehov",
            command=self.copy_ordersaldo_list2,
            style="Warning.TButton",
            state="disabled",
        )
        self.ordersaldo_copy_list2_btn.pack(side="left", padx=4)
        self.vecka27_btn = ttk.Button(
            ordersaldo_frame,
            text="Vecka 27",
            command=self.run_vecka27_check,
            style="Warning.TButton",
            state="disabled",
        )
        self.vecka27_btn.pack(side="left", padx=4)
        self.lyx_btn = ttk.Button(
            ordersaldo_frame,
            text="LYX",
            command=self.run_lyx,
            style="Warning.TButton",
            state="disabled",
        )
        self.lyx_btn.pack(side="left", padx=4)
        self.pafyllnadsprio_btn = ttk.Button(
            ordersaldo_frame,
            text="Påfyllnadsprio",
            command=self.run_pafyllnadsprio,
            style="Warning.TButton",
            state="disabled",
        )
        self.pafyllnadsprio_btn.pack(side="left", padx=4)
        self.reset_cache_btn = ttk.Button(ordersaldo_frame, text="Rensa cache", command=self.reset_cache, style="Green.TButton")
        self.reset_cache_btn.pack(side="left", padx=(16, 4))
        self.hjälp_btn = ttk.Button(ordersaldo_frame, text="? Hjälp", command=self._toggle_help_mode)
        self.hjälp_btn.pack(side="left", padx=(16, 4))
        self._action_requirements = {
            self.run_btn: [
                ("orders", "Bestallningslinjer (CSV)"),
                ("buffer", "Buffertpallar (CSV)"),
            ],
            self.koppla_btn: [
                ("orders", "Bestallningslinjer (CSV)"),
                ("overview", "Orderoversikt (CSV)"),
            ],
            self.overview_check_btn: [
                ("overview", "Orderoversikt (CSV)"),
            ],
            self.dispatch_check_btn: [
                ("overview", "Orderoversikt (CSV)"),
                ("dispatch", "Dispatchpallar (CSV)"),
            ],
            self.eftersok_btn: [
                ("wms_receive", "Mottagningslogg (CSV)"),
            ],
            self.lyx_btn: [
                ("automation", "Saldo inkl. automation (CSV)"),
            ],
            self.pafyllnadsprio_btn: [
                ("orders", "Bestallningslinjer (CSV)"),
            ],
        }
        for action_btn in self._action_requirements:
            action_btn.bind("<Enter>", self._on_action_button_hover, add="+")
            action_btn.bind("<Motion>", self._on_action_button_hover, add="+")
            action_btn.bind("<Leave>", self._hide_hover_tooltip, add="+")
            action_btn.bind("<ButtonPress-1>", self._hide_hover_tooltip, add="+")
        self._update_action_buttons_state()

        open_frame = ttk.Frame(self)
        open_frame.grid(row=5, column=0, columnspan=4, sticky="w", pady=10)
        self.open_result_btn = ttk.Button(open_frame, text="Öppna allokerade pallar", command=self.open_result_in_excel, state="disabled")
        self.open_result_btn.grid(row=0, column=0, padx=4)
        self.open_nearmiss_btn = ttk.Button(open_frame, text="Öppna near-miss", command=self.open_nearmiss_in_excel, state="disabled")
        self.open_nearmiss_btn.grid(row=0, column=1, padx=4)
        self.open_palletspaces_btn = ttk.Button(open_frame, text="Öppna pallplatser", command=self.open_pallet_spaces_in_excel, state="disabled")
        self.open_palletspaces_btn.grid(row=0, column=2, padx=4)
        self.open_prognos_btn = ttk.Button(open_frame, text="Öppna prognos", command=self.open_prognos_in_excel, state="disabled")
        self.open_prognos_btn.grid(row=0, column=3, padx=4)
        self.open_refill_btn = ttk.Button(open_frame, text="Öppna refill", command=self.open_refill_in_excel, state="disabled")
        self.open_refill_btn.grid(row=0, column=4, padx=4)
        # Flytta knappen för att öppna HIB‑kopplingen till vänster om Rensa cache
        self.open_koppla_btn = ttk.Button(open_frame, text="Öppna HIB‑koppling", command=self.open_koppla_in_excel, state="disabled")
        self.open_koppla_btn.grid(row=0, column=5, padx=4)
        # Nya knappar för att öppna resultatet av order- och dispatchkontroller
        self.open_overview_check_btn = ttk.Button(open_frame, text="Öppna orderkontroll", command=self.open_overview_check_in_excel, state="disabled")
        self.open_overview_check_btn.grid(row=0, column=6, padx=4)
        self.open_dispatch_check_btn = ttk.Button(open_frame, text="Öppna dispatchkontroll", command=self.open_dispatch_check_in_excel, state="disabled")
        self.open_dispatch_check_btn.grid(row=0, column=7, padx=4)
        self.open_eftersok_btn = ttk.Button(open_frame, text="Öppna Eftersök", command=self.open_eftersok_in_excel, state="disabled")
        self.open_eftersok_btn.grid(row=0, column=8, padx=4)
        self._open_button_hints = {
            self.open_result_btn: "Tryck forst: Kor allokering",
            self.open_nearmiss_btn: "Tryck forst: Kor allokering",
            self.open_palletspaces_btn: "Tryck forst: Kor allokering",
            self.open_refill_btn: "Tryck forst: Kor allokering",
            self.open_koppla_btn: "Tryck forst: Kor HIB-koppling",
            self.open_overview_check_btn: "Tryck forst: Kontrollera orderoversikt",
            self.open_dispatch_check_btn: "Tryck forst: Kontrollera dispatchpallar",
            self.open_eftersok_btn: "Tryck forst: Eftersok",
            self.open_prognos_btn: "Ladda upp Prognos eller Kampanjvolymer forst",
        }
        for open_btn in self._open_button_hints:
            open_btn.bind("<Enter>", self._on_open_button_hover, add="+")
            open_btn.bind("<Motion>", self._on_open_button_hover, add="+")
            open_btn.bind("<Leave>", self._hide_hover_tooltip, add="+")
            open_btn.bind("<ButtonPress-1>", self._hide_hover_tooltip, add="+")

        # Dragbar delare mellan logg och summering så man kan höja/sänka loggrutan.
        self.log_splitter = ttk.Panedwindow(self, orient="vertical")
        self.log_splitter.grid(row=6, column=0, columnspan=4, sticky="nsew", padx=8, pady=(0, 8))
        self.rowconfigure(6, weight=1)

        log_frame = ttk.Frame(self.log_splitter)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        ttk.Label(log_frame, text="Logg / Summering:").grid(row=0, column=0, sticky="w")
        self.log = tk.Text(log_frame, height=14, width=110, state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

        summary_frame = ttk.Frame(self.log_splitter)
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(1, weight=1)
        ttk.Label(summary_frame, text="Summering per Källtyp").grid(row=0, column=0, sticky="w")
        self.summary_table = ttk.Treeview(summary_frame, columns=("ktyp", "antal_rader", "antal_kolli"), show="headings", height=5)
        self.summary_table.heading("ktyp", text="Källtyp")
        self.summary_table.heading("antal_rader", text="antal rader")
        self.summary_table.heading("antal_kolli", text="antal kolli")
        self.summary_table.column("ktyp", anchor="w", width=160)
        self.summary_table.column("antal_rader", anchor="e", width=140)
        self.summary_table.column("antal_kolli", anchor="e", width=140)
        self.summary_table.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

        try:
            self.log_splitter.add(log_frame, weight=3)
            self.log_splitter.add(summary_frame, weight=1)
        except tk.TclError:
            self.log_splitter.add(log_frame)
            self.log_splitter.add(summary_frame)
        # Starta med en balanserad delning så handtaget går att dra både upp/ner direkt.
        self.after(120, self._set_log_splitter_default_position)

        # Hjälp-registry: registrerar widget → ämne för hjälptexten
        def _hreg(widget: tk.Widget, topic: str) -> None:
            self._HELP_REGISTRY[id(widget)] = topic
            self._help_widget_map[id(widget)] = widget

        _hreg(self.run_btn, "run_btn")
        _hreg(self.koppla_btn, "koppla_btn")
        _hreg(self.overview_check_btn, "overview_check_btn")
        _hreg(self.dispatch_check_btn, "dispatch_check_btn")
        _hreg(self.eftersok_btn, "eftersok_btn")
        _hreg(self.hjälp_btn, "hjalp_btn")
        _hreg(self.lyx_btn, "lyx_btn")
        _hreg(self.ordersaldo_copy_list1_btn, "ordersaldo_list1_btn")
        _hreg(self.ordersaldo_copy_list2_btn, "ordersaldo_list2_btn")
        _hreg(self.vecka27_btn, "vecka27_btn")
        _hreg(self.pafyllnadsprio_btn, "pafyllnadsprio_btn")
        _hreg(self.reset_cache_btn, "reset_cache_btn")
        _hreg(self.open_result_btn, "open_result_btn")
        _hreg(self.open_nearmiss_btn, "open_nearmiss_btn")
        _hreg(self.open_palletspaces_btn, "open_palletspaces_btn")
        _hreg(self.open_prognos_btn, "open_prognos_btn")
        _hreg(self.open_refill_btn, "open_refill_btn")
        _hreg(self.open_koppla_btn, "open_koppla_btn")
        _hreg(self.open_overview_check_btn, "open_overview_check_btn")
        _hreg(self.open_dispatch_check_btn, "open_dispatch_check_btn")
        _hreg(self.open_eftersok_btn, "open_eftersok_btn")
        _hreg(self.log, "log_widget")
        _hreg(self.summary_table, "summary_table")
        for ft, (lbl, _btn) in self.file_status_widgets.items():
            _hreg(lbl, f"file_{ft}")

        self.last_result_df: pd.DataFrame | None = None
        self.last_nearmiss_instead_df: pd.DataFrame | None = None
        self._orders_raw: pd.DataFrame | None = None
        self._buffer_raw: pd.DataFrame | None = None
        self._result_df: pd.DataFrame | None = None

        self._not_putaway_raw: pd.DataFrame | None = None
        self._not_putaway_norm: pd.DataFrame | None = None
        self._saldo_norm: pd.DataFrame | None = None

        self._saldo_raw: pd.DataFrame | None = None

        self._item_raw: pd.DataFrame | None = None
        self._item_norm: pd.DataFrame | None = None

        self._sales_metrics_df: pd.DataFrame | None = None

        self._last_refill_hp_df: pd.DataFrame | None = None
        self._last_refill_autostore_df: pd.DataFrame | None = None

        self._pallet_spaces_df: pd.DataFrame | None = None

        self._prognos_df: pd.DataFrame | None = None

        # För HIB‑kopplingens resultat
        self.last_koppla_df: pd.DataFrame | None = None
        # För missade avgångar i HIB‑kopplingen
        self.last_koppla_missed_df: pd.DataFrame | None = None
        self.last_koppla_path: str | None = None

        # För orderöversikt- och dispatchkontroll
        self.last_overview_check_df: pd.DataFrame | None = None
        self.last_hib_status_check_df: pd.DataFrame | None = None
        self.last_overview_check_path: str | None = None
        self.last_dispatch_check_df: pd.DataFrame | None = None
        self.last_dispatch_check_path: str | None = None
        self.last_eftersok_df: pd.DataFrame | None = None
        self.last_eftersok_path: str | None = None
        self.last_eftersok_report: str | None = None



        # Hela fönstret är droppbart när TkinterDnD finns tillgängligt.
        self._bind_global_drop_targets()


    def pick_orders(self) -> None:
        path = filedialog.askopenfilename(title="Välj beställningsrader (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("orders", path, source="picker")
            try:
                self.update_file_status_icons()
            except Exception:
                pass

    def pick_automation(self) -> None:
        path = filedialog.askopenfilename(title="Välj Saldo inkl. automation (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("automation", path, source="picker")
            try:
                self.update_file_status_icons()
            except Exception:
                pass

    def pick_buffer(self) -> None:
        path = filedialog.askopenfilename(title="Välj buffertpallar (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("buffer", path, source="picker")
            try:
                self.update_file_status_icons()
            except Exception:
                pass

    def _on_buffer_var_changed(self, *args) -> None:
        """Triggas nar buffer_var andras. Lasin filen i bakgrundstrad och uppdatera observations."""
        path = self.buffer_var.get().strip()
        if not path or path == self._last_observations_path:
            return
        if not Path(path).is_file():
            return
        self._last_observations_path = path
        threading.Thread(target=self._observations_worker, args=(path,), daemon=True).start()

    def _observations_worker(self, path: str) -> None:
        """Las buffer-filen och uppdatera observations.csv.gz + push till GitHub. Korbar fran bakgrundstrad."""
        try:
            buffer_raw = pd.read_csv(path, dtype=str, sep=None, engine="python")
            buffer_raw = _clean_columns(buffer_raw)
            n_nya, nya_rader = update_observations_from_buffer(buffer_raw)
        except Exception as e:
            self.master.after(0, lambda: self._log(f"[Observations] Misslyckades lasa buffertpall: {e}"))
            return
        if n_nya:
            self.master.after(0, lambda: self._log(
                f"[Observations] {n_nya} nya pallid sparade lokalt (artikel_max.csv uppdaterad)."
            ))
            try:
                pushed = push_new_observations_to_github(nya_rader)
                if pushed:
                    self.master.after(0, lambda: self._log(
                        f"[Observations] {n_nya} pallid pushade till GitHub."
                    ))
                else:
                    self.master.after(0, lambda: self._log(
                        "[Observations] GitHub-push hoppades over (ingen token eller API-fel)."
                    ))
            except Exception as e:
                err = str(e)
                self.master.after(0, lambda: self._log(f"[Observations] GitHub-push misslyckades: {err}"))

    # ------------------------------------------------------------------ #
    #  Hjälpläge                                                          #
    # ------------------------------------------------------------------ #

    def _toggle_help_mode(self) -> None:
        if self._help_mode:
            self._exit_help_mode()
        else:
            self._enter_help_mode("enkel")

    def _enter_help_mode(self, mode: str) -> None:
        self._help_mode = mode
        self.hjälp_btn.configure(text="✕ Stäng hjälp")
        self._show_help_overlay()
        self._build_help_bar()
        self.master.bind("<Escape>", lambda _e: self._exit_help_mode())

    def _exit_help_mode(self) -> None:
        self._help_mode = ""
        self.hjälp_btn.configure(text="? Hjälp")
        self._remove_help_overlay()
        self._close_help_popup()
        self._close_help_bar()
        try:
            self.master.unbind("<Escape>")
        except Exception:
            pass

    def _show_help_overlay(self) -> None:
        self._remove_help_overlay()
        # Toplevel-overlay med alpha ger riktig genomskinlighet (Canvas kan inte det)
        overlay = tk.Toplevel(self.master)
        overlay.wm_overrideredirect(True)
        overlay.wm_attributes("-topmost", True)
        overlay.wm_attributes("-alpha", 0.35)
        overlay.configure(cursor="question_arrow")
        self._help_overlay = overlay
        self._update_overlay_color()
        self._position_overlay()
        overlay.bind("<Button-1>", self._on_help_click)
        overlay.bind("<Escape>", lambda _e: self._exit_help_mode())
        self.master.bind("<Configure>", self._on_main_configure_help, add="+")

    def _update_overlay_color(self) -> None:
        if self._help_overlay is None:
            return
        color = "#CC2222" if self._help_mode == "avancerat" else "#808080"
        self._help_overlay.configure(bg=color)

    def _position_overlay(self) -> None:
        if self._help_overlay is None:
            return
        self.master.update_idletasks()
        x = self.master.winfo_rootx()
        y = self.master.winfo_rooty()
        w = self.master.winfo_width()
        h = self.master.winfo_height()
        self._help_overlay.geometry(f"{w}x{h}+{x}+{y}")

    def _on_main_configure_help(self, event) -> None:
        if event.widget is self.master:
            self._position_overlay()
            self._reposition_help_bar()

    def _remove_help_overlay(self) -> None:
        try:
            self.master.unbind("<Configure>")
        except Exception:
            pass
        if self._help_overlay is not None:
            try:
                self._help_overlay.destroy()
            except Exception:
                pass
            self._help_overlay = None

    def _on_help_click(self, event) -> None:
        topic = self._find_topic_at(event.x_root, event.y_root)
        self._show_help_popup(topic, event.x_root, event.y_root)

    def _find_topic_at(self, x_root: int, y_root: int) -> Optional[str]:
        """Hitta hjälpämne för widgeten under skärmkoordinat (x_root, y_root)."""
        for widget_id, topic in self._HELP_REGISTRY.items():
            widget = self._help_widget_map.get(widget_id)
            if widget is None:
                continue
            try:
                if not widget.winfo_exists():
                    continue
                wx = widget.winfo_rootx()
                wy = widget.winfo_rooty()
                ww = widget.winfo_width()
                wh = widget.winfo_height()
                if wx <= x_root < wx + ww and wy <= y_root < wy + wh:
                    return topic
            except Exception:
                pass
        return None

    def _get_help_text(self, topic: str) -> str:
        folder = "avancerat" if self._help_mode == "avancerat" else "enkel"
        for base in [_resource_path(f"hjalp/{folder}/{topic}.txt"),
                     _resource_path(f"hjalp/enkel/{topic}.txt")]:
            if base.exists():
                try:
                    return base.read_text(encoding="utf-8").strip()
                except Exception:
                    pass
        return "Ingen hjälptext hittades för det här elementet."

    def _show_help_popup(self, topic: Optional[str], x_root: int, y_root: int) -> None:
        self._close_help_popup()
        text = self._get_help_text(topic) if topic else "Klicka på en knapp eller ett fält för mer information."
        title = topic.replace("_", " ").title() if topic else "Hjälp"

        popup = tk.Toplevel(self.master)
        popup.wm_title(title)
        popup.wm_attributes("-topmost", True)
        popup.resizable(False, False)
        self._help_popup = popup

        # Positionera nära klicket men håll den inom skärmen
        pw, ph = 360, 200
        sx = self.master.winfo_screenwidth()
        sy = self.master.winfo_screenheight()
        px = min(x_root + 12, sx - pw - 8)
        py = min(y_root + 12, sy - ph - 8)
        popup.geometry(f"{pw}x{ph}+{px}+{py}")

        frm = tk.Frame(popup, bg="#1f2933", padx=10, pady=8)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text=title, bg="#1f2933", fg="white",
                 font=("Arial", 11, "bold"), anchor="w").pack(fill="x")
        tk.Frame(frm, bg="#444444", height=1).pack(fill="x", pady=(4, 6))

        txt = tk.Text(frm, wrap="word", bg="#1f2933", fg="#e0e0e0",
                      font=("Arial", 10), relief="flat", height=7,
                      state="normal", cursor="arrow")
        txt.insert("1.0", text)
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)

        close_btn = tk.Button(frm, text="✕ Stäng", command=self._close_help_popup,
                              bg="#444444", fg="white", relief="flat",
                              activebackground="#666666", activeforeground="white",
                              padx=8, pady=2)
        close_btn.pack(anchor="e", pady=(6, 0))

        popup.bind("<Escape>", lambda _e: self._exit_help_mode())
        popup.protocol("WM_DELETE_WINDOW", self._close_help_popup)

    def _close_help_popup(self) -> None:
        if self._help_popup is not None:
            try:
                self._help_popup.destroy()
            except Exception:
                pass
            self._help_popup = None

    def _build_help_bar(self) -> None:
        self._close_help_bar()
        bar = tk.Toplevel(self.master)
        bar.wm_overrideredirect(True)
        bar.wm_attributes("-topmost", True)
        bar.configure(bg="#CC2222")   # röd kant runt hela baren
        self._help_bar = bar

        frm = tk.Frame(bar, bg="#1a1a1a", padx=10, pady=8)
        frm.pack(fill="both", expand=True, padx=2, pady=2)

        tk.Label(frm, text="❓ Hjälpläge – klicka på valfri del av appen",
                 bg="#1a1a1a", fg="white",
                 font=("Arial", 10, "bold")).pack(side="left", padx=(0, 14))

        tk.Label(frm, text="Avancerat", bg="#1a1a1a", fg="#CCCCCC",
                 font=("Arial", 9)).pack(side="left")
        toggle = SlideToggle(frm, command=self._on_advanced_toggle, bg="#1a1a1a")
        toggle.pack(side="left", padx=(6, 16))
        self._help_toggle = toggle

        tk.Button(frm, text="✕  Stäng hjälp", command=self._exit_help_mode,
                  bg="#CC2222", fg="white", relief="flat",
                  activebackground="#AA0000", activeforeground="white",
                  font=("Arial", 10, "bold"), padx=8, pady=2).pack(side="left")

        bar.bind("<Escape>", lambda _e: self._exit_help_mode())
        # Ge tkinter tid att rendera baren innan positionering
        bar.update_idletasks()
        self._reposition_help_bar()

    def _reposition_help_bar(self) -> None:
        if self._help_bar is None:
            return
        self._help_bar.update_idletasks()
        bw = self._help_bar.winfo_reqwidth()
        bh = self._help_bar.winfo_reqheight()
        wx = self.master.winfo_rootx()
        wy = self.master.winfo_rooty()
        ww = self.master.winfo_width()
        x = wx + ww - bw - 8
        y = wy + 8
        self._help_bar.geometry(f"+{x}+{y}")

    def _close_help_bar(self) -> None:
        if self._help_bar is not None:
            try:
                self._help_bar.destroy()
            except Exception:
                pass
            self._help_bar = None
            self._help_toggle = None

    def _on_advanced_toggle(self, state: bool) -> None:
        self._help_mode = "avancerat" if state else "enkel"
        self._update_overlay_color()
        self._close_help_popup()

    def pick_item(self) -> None:
        """
        Öppna dialog för att välja item-fil (CSV) med staplingsbar-uppgift.
        """
        path = filedialog.askopenfilename(title="Välj item-fil (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("item", path, source="picker")
            try:
                self.update_file_status_icons()
            except Exception:
                pass

    def pick_overview(self) -> None:
        """
        Öppna dialog för att välja orderöversikt (CSV).  Denna fil innehåller
        övergripande information om ordrar inklusive ordertyp, kundnummer,
        orderdatum, sändningsnummer, zoner och multi.  Endast en fil behöver
        väljas och sparas i overview_var.
        """
        path = filedialog.askopenfilename(title="Välj orderöversikt (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("overview", path, source="picker")
            try:
                self.update_file_status_icons()
            except Exception:
                pass

    def pick_not_putaway(self) -> None:
        """
        Stub för filval av 'Ej inlagrade artiklar'. Denna funktion gör inget i denna version.
        """
        return

    def _on_global_drop(self, event) -> None:
        """Gemensam drop-handler för hela fönstret."""
        self._handle_drop_all(event)
        try:
            self.update_file_status_icons()
        except Exception:
            pass

    def _bind_global_drop_targets(self) -> None:
        """Gör hela appfönstret droppbart när TkinterDnD finns tillgängligt."""
        if not (TkinterDnD and DND_FILES):
            return
        seen: set[str] = set()

        def _register(widget) -> None:
            try:
                wid = str(widget)
                if wid in seen:
                    return
                seen.add(wid)
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_global_drop)
            except Exception:
                pass
            try:
                for child in widget.winfo_children():
                    _register(child)
            except Exception:
                pass

        _register(self.master)

    def _parse_dnd_paths(self, event_data: str) -> list[str]:
        """Tolka en DnD-sträng (kan innehålla en eller flera filvägar inom klamrar) till en lista med paths."""
        raw = str(event_data).strip()
        paths: list[str] = []
        i = 0
        while raw:
            raw = raw.strip()
            if not raw:
                break
            if raw.startswith("{"):
                end = raw.find("}")
                if end == -1:
                    break
                path = raw[1:end]
                paths.append(path)
                raw = raw[end+1:]
            else:
                if ' ' in raw:
                    part, raw = raw.split(' ', 1)
                else:
                    part, raw = raw, ''
                if part:
                    paths.append(part)
        return paths

    def _detect_file_type(self, path: str) -> str | None:
        """Försök avgöra vilken sorts fil det är (orders, buffer, automation, item, prognos, campaign).
        Returnerar en sträng med typen eller None om okänd.
        """
        import os
        import pandas as _pd
        ext = os.path.splitext(path)[1].lower().lstrip('.')
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
        # Filename hint fallback for buffer exports with varying schemas.
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
                df_c = read_campaign_xlsx(path)
                if isinstance(df_c, _pd.DataFrame) and not df_c.empty and list(df_c.columns) == ["Artikelnummer", "Antal styck"]:
                    return "campaign"
            except Exception:
                pass
            try:
                df_p = read_prognos_xlsx(path)
                if isinstance(df_p, _pd.DataFrame) and not df_p.empty and len(df_p.columns) >= 3 and any(str(c).strip().lower() in ("antal styck", "quantity", "qty") for c in df_p.columns):
                    return "prognos"
            except Exception:
                pass
            return None
        try:
            df = _pd.read_csv(path, dtype=str, nrows=50, sep=None, engine="python", encoding="utf-8-sig")
            if df.shape[1] == 1:
                df = _pd.read_csv(path, dtype=str, nrows=50, sep="\t", engine="python", encoding="utf-8-sig")
        except Exception:
            try:
                df = _pd.read_csv(path, dtype=str, nrows=50, sep="\t", engine="python", encoding="utf-8-sig")
            except Exception:
                return None
        cols = [str(c).strip().lower() for c in df.columns]
        has_art = any(c in ("artikel", "artikelnummer", "artnr", "art.nr", "sku", "article") for c in cols)
        has_qty = any(c in ("beställt", "antal", "qty", "quantity", "bestalld", "order qty", "antal styck") for c in cols)
        has_ord = any(c in ("ordernr", "order nr", "order number", "kund", "kundnr", "order id") for c in cols)
        has_rad = any(c in ("radnr", "rad nr", "line id", "rad", "struktur", "radsnr") for c in cols)
        if has_art and has_qty and (has_ord or has_rad):
            return "orders"
        has_lagerplats = any("lagerplats" in c or "plats" == c or "location" == c or "bin" == c for c in cols)
        has_pallid = any(c in ("pallid", "pall id", "id", "sscc", "etikett", "batch") for c in cols)
        has_status = any(c == "status" for c in cols)
        has_inkop = any("inköpsnr" in c or "inkopsnr" in c for c in cols)
        has_mottaget = any("mottaget" in c for c in cols)
        has_pallnr = any("pall nr" in c or "pallnr" in c for c in cols)
        has_till = any(c == "till" or c.endswith(" till") or c.startswith("till ") for c in cols)
        has_fran = any("från" in c or "fran" in c for c in cols)
        has_plockat = any("plockat" in c for c in cols)
        has_anledning = any("anledning" in c for c in cols)

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
        # Buffer exports can include saldo-like columns. Keep buffer precedence.
        buffer_marker_count = sum(
            1 for flag in (has_lagerplats, has_pallid, has_status, has_inkop, has_mottaget, has_pallnr) if flag
        )
        if has_art and (has_qty or has_pallid) and buffer_marker_count >= 2:
            return "buffer"
        has_pack = any("pack klass" in c or "staplingsbar" in c for c in cols)
        # Om filen innehåller pack‑relaterade kolumner ("pack klass" eller "staplingsbar"),
        # kontrollera först om den också motsvarar en dispatchfil. Dispatchfiler har
        # plockpallskolumn samt både ordernummer och sändningsinformation. Utan denna
        # kontroll klassificerades dispatchpallar felaktigt som item.
        if has_pack:
            # dispatch‑indikatorer
            has_plockpall = any("plockpall" in c for c in cols)
            has_dispatch_order = any(c in ("ordernr", "order nr", "order number", "ordernummer") for c in cols)
            has_dispatch_ship = any(
                ("sändnings" in c) or ("sandnings" in c) or ("sändningsnr" in c) or ("sandningsnr" in c) or ("sändningsnr." in c) or ("sandningsnr." in c)
                for c in cols
            )
            # om dispatchindikatorer hittas, återgå "dispatch" istället för "item"
            if has_plockpall and has_dispatch_order and has_dispatch_ship:
                return "dispatch"
            return "item"
        # Ny detektering för orderöversikt (overview)
        has_ordernr = any(c in ("ordernr", "order nr", "order number") for c in cols)
        has_orderdatum = any("orderdatum" in c for c in cols)
        has_sandning = any("sändningsnr" in c or "sandningsnr" in c or "sändningsnr." in c or "sandnr" in c for c in cols)
        has_ordertyp = any("ordertyp" in c for c in cols)
        has_multi = any("multi" == c for c in cols)
        # kräver flera av dessa kolumner för att identifiera en orderöversikt
        if has_ordernr and has_orderdatum and has_sandning and has_ordertyp:
            return "overview"
        # Ny detektering för dispatchpallar (dispatch)
        has_plockpall = any("plockpall" in c for c in cols)
        has_dispatch_order = any(c in ("ordernr", "order nr", "order number", "ordernummer") for c in cols)
        has_dispatch_ship = any(
            ("sändnings" in c) or ("sandnings" in c) or ("sändningsnr" in c) or ("sandningsnr" in c) or ("sändningsnr." in c) or ("sandningsnr." in c)
            for c in cols
        )
        if has_plockpall and has_dispatch_order and has_dispatch_ship:
            return "dispatch"
        has_robot = any(c == "robot" for c in cols)
        has_saldo = any(("saldo autoplock" in c) or (c == "plocksaldo") or (c == "plock saldo") for c in cols)
        if has_art and (has_robot or has_saldo):
            return "automation"
        return None

    def _handle_drop_all(self, event) -> None:
        """Hantera drop av en eller flera filer i den gemensamma drop-zonen."""
        paths = self._parse_dnd_paths(event.data)
        for p in paths:
            p = p.strip()
            if not p:
                continue
            file_type = self._detect_file_type(p)
            if file_type == "orders":
                self._set_file_path("orders", p, source="drag_drop")
            elif file_type == "buffer":
                self._set_file_path("buffer", p, source="drag_drop")
            elif file_type == "automation":
                self._set_file_path("automation", p, source="drag_drop")
            elif file_type == "item":
                self._set_file_path("item", p, source="drag_drop")
            elif file_type == "prognos":
                self._set_file_path("prognos", p, source="drag_drop")
            elif file_type == "campaign":
                self._set_file_path("campaign", p, source="drag_drop")
            elif file_type == "overview":
                self._set_file_path("overview", p, source="drag_drop")
            elif file_type == "dispatch":
                self._set_file_path("dispatch", p, source="drag_drop")
            elif file_type == "wms_receive":
                self._set_file_path("wms_receive", p, source="drag_drop")
            elif file_type == "wms_booking":
                self._set_file_path("wms_booking", p, source="drag_drop")
            elif file_type == "wms_trans":
                self._set_file_path("wms_trans", p, source="drag_drop")
            elif file_type == "wms_pick":
                self._set_file_path("wms_pick", p, source="drag_drop")
            elif file_type == "wms_correct":
                self._set_file_path("wms_correct", p, source="drag_drop")
            else:
                self._log(f"Okänd filtyp: {p}")

    def pick_sales(self) -> None:
        """
        Stub för filval av plocklogg. Denna funktion gör inget i denna version.
        """
        return

    def update_file_status_icons(self) -> None:
        """
        Uppdatera ikonerna för filinmatningsraderna. Grön bock för uppladdad fil,
        grått streck för ingen fil och inaktivera röd kryss vid tomt fält.
        """
        try:
            for ft, (lbl, btn) in self.file_status_widgets.items():
                var = self.file_vars.get(ft)
                path = var.get().strip() if var else ""
                if path:
                    # fil har valts: visa "Uppladdad" med grön bakgrund och vit text
                    lbl.config(text="Uppladdad", fg="white", bg="#28a745")
                    btn.config(state="normal")
                else:
                    # ingen fil: visa "Ej fil" med grå bakgrund
                    lbl.config(text="Ej fil", fg="white", bg="#6c757d")
                    btn.config(state="disabled")
        except Exception:
            pass
        try:
            self._update_action_buttons_state()
        except Exception:
            pass
        try:
            self._refresh_value_filter_options()
        except Exception:
            pass
        try:
            self._refresh_ordersaldo_from_orders()
        except Exception:
            pass

    def _normalize_filter_value(self, value: object) -> str:
        """Normalisera filtervärden till trimmat versalt textvärde eller NULL-token."""
        try:
            if pd.isna(value):
                return self.filter_null_token
        except Exception:
            pass
        raw = str(value).strip()
        if not raw:
            return self.filter_null_token
        if raw.lower() in {"null", "none", "nan", "nat", "na"}:
            return self.filter_null_token
        return raw.upper()

    def _display_filter_value(self, value: str) -> str:
        """Visa intern NULL-token som läsbar etikett."""
        return self.filter_null_label if value == self.filter_null_token else str(value)

    def _find_filter_column(self, df: pd.DataFrame, filter_key: str) -> Optional[str]:
        """Hitta kolumn för filtergruppen via robust kandidatmatchning."""
        try:
            return find_col(
                df,
                self.filter_column_candidates.get(filter_key, []),
                required=False,
                default=None,
            )
        except Exception:
            return None

    def _read_tabular_for_filter_scan(self, path: str) -> Optional[pd.DataFrame]:
        """Läs en tabulär fil för att extrahera unika filtervärden."""
        ext = os.path.splitext(str(path))[1].lower()
        try:
            if ext == ".csv":
                df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
                if df.shape[1] == 1:
                    try:
                        df = pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
                    except Exception:
                        pass
                return _clean_columns(df)
            if ext in {".xlsx", ".xlsm", ".xls"}:
                try:
                    df = pd.read_excel(path, dtype=str)
                except Exception:
                    return None
                return _clean_columns(df)
        except Exception:
            return None
        return None

    def _refresh_value_filter_options(self) -> None:
        """Bygg om Bolag/Ordertyp-filter utifrån uppladdade filer och behåll tidigare val."""
        discovered: dict[str, set[str]] = {"bolag": set(), "ordertyp": set()}
        discovered_pairs: set[tuple[str, str]] = set()
        for var in self.file_vars.values():
            path = var.get().strip() if var else ""
            if not path:
                continue
            df = self._read_tabular_for_filter_scan(path)
            if not isinstance(df, pd.DataFrame):
                continue
            for filter_key in ("bolag", "ordertyp"):
                col = self._find_filter_column(df, filter_key)
                if not col or col not in df.columns:
                    continue
                try:
                    values = df[col].tolist()
                except Exception:
                    values = []
                for val in values:
                    discovered[filter_key].add(self._normalize_filter_value(val))
            bolag_col = self._find_filter_column(df, "bolag")
            ordertyp_col = self._find_filter_column(df, "ordertyp")
            if bolag_col and ordertyp_col and bolag_col in df.columns and ordertyp_col in df.columns:
                try:
                    pair_df = df[[bolag_col, ordertyp_col]].copy()
                    for _, row in pair_df.iterrows():
                        discovered_pairs.add(
                            (
                                self._normalize_filter_value(row.get(bolag_col)),
                                self._normalize_filter_value(row.get(ordertyp_col)),
                            )
                        )
                except Exception:
                    pass

        for filter_key in ("bolag", "ordertyp"):
            previous_values = {
                value: bool(var.get())
                for value, var in self.filter_vars.get(filter_key, {}).items()
            }
            sorted_values = sorted(
                discovered[filter_key],
                key=lambda v: (v == self.filter_null_token, self._display_filter_value(v)),
            )
            new_var_map: dict[str, tk.BooleanVar] = {}
            old_var_map = self.filter_vars.get(filter_key, {})
            for value in sorted_values:
                is_checked = previous_values.get(value, True)
                existing_var = old_var_map.get(value)
                if existing_var is None:
                    existing_var = tk.BooleanVar(value=is_checked)
                else:
                    existing_var.set(is_checked)
                new_var_map[value] = existing_var
            self.filter_options[filter_key] = sorted_values
            self.filter_vars[filter_key] = new_var_map
        self._filter_pairs = discovered_pairs

        self._render_value_filter_options()
        self._enforce_filter_dependencies()

    def _render_value_filter_options(self) -> None:
        """Rendera checkboxes för alla upptäckta filtervärden."""
        any_group_visible = False
        for filter_key in ("bolag", "ordertyp"):
            frame = self._filter_group_frames.get(filter_key)
            if frame is None:
                continue
            self._filter_checkbuttons[filter_key] = {}
            for child in frame.winfo_children():
                child.destroy()
            values = self.filter_options.get(filter_key, [])
            if not values:
                frame.grid_remove()
                continue
            frame.grid()
            any_group_visible = True
            for idx, value in enumerate(values):
                cb = ttk.Checkbutton(
                    frame,
                    text=self._display_filter_value(value),
                    variable=self.filter_vars[filter_key][value],
                    command=self._on_filter_selection_changed,
                )
                cb.grid(row=idx, column=0, sticky="w", padx=4, pady=1)
                self._filter_checkbuttons[filter_key][value] = cb
        if any_group_visible:
            self.value_filter_frame.grid()
        else:
            self.value_filter_frame.grid_remove()

    def _on_filter_selection_changed(self) -> None:
        """Hantera klick på filter-checkboxar."""
        try:
            self._enforce_filter_dependencies()
        except Exception:
            pass
        try:
            self._refresh_ordersaldo_from_orders()
        except Exception:
            pass
        try:
            self._hide_hover_tooltip()
        except Exception:
            pass

    def _enforce_filter_dependencies(self) -> None:
        """
        Säkerställ kompatibla kombinationer mellan Bolag och Ordertyp.
        Ogiltiga val bockas ur och motsvarande checkboxar inaktiveras.
        """
        if not self._filter_pairs:
            for filter_key in ("bolag", "ordertyp"):
                for value, cb in self._filter_checkbuttons.get(filter_key, {}).items():
                    try:
                        cb.configure(state="normal")
                    except Exception:
                        pass
            return

        all_bolag = set(self.filter_options.get("bolag", []))
        all_ordertyp = set(self.filter_options.get("ordertyp", []))

        for _ in range(6):
            selected_bolag = self._get_selected_filter_values("bolag")
            selected_ordertyp = self._get_selected_filter_values("ordertyp")

            if selected_ordertyp:
                enabled_bolag = {b for (b, o) in self._filter_pairs if o in selected_ordertyp}
            else:
                enabled_bolag = set(all_bolag)

            if selected_bolag:
                enabled_ordertyp = {o for (b, o) in self._filter_pairs if b in selected_bolag}
            else:
                enabled_ordertyp = set(all_ordertyp)

            changed = False
            for value, var in self.filter_vars.get("bolag", {}).items():
                if value not in enabled_bolag and bool(var.get()):
                    var.set(False)
                    changed = True
            for value, var in self.filter_vars.get("ordertyp", {}).items():
                if value not in enabled_ordertyp and bool(var.get()):
                    var.set(False)
                    changed = True
            if not changed:
                break

        selected_bolag = self._get_selected_filter_values("bolag")
        selected_ordertyp = self._get_selected_filter_values("ordertyp")
        if selected_ordertyp:
            enabled_bolag = {b for (b, o) in self._filter_pairs if o in selected_ordertyp}
        else:
            enabled_bolag = set(all_bolag)
        if selected_bolag:
            enabled_ordertyp = {o for (b, o) in self._filter_pairs if b in selected_bolag}
        else:
            enabled_ordertyp = set(all_ordertyp)

        for value, cb in self._filter_checkbuttons.get("bolag", {}).items():
            try:
                cb.configure(state="normal" if value in enabled_bolag else "disabled")
            except Exception:
                pass
        for value, cb in self._filter_checkbuttons.get("ordertyp", {}).items():
            try:
                cb.configure(state="normal" if value in enabled_ordertyp else "disabled")
            except Exception:
                pass

    def _get_selected_filter_values(self, filter_key: str) -> set[str]:
        """Returnera markerade värden för en filtergrupp."""
        selected: set[str] = set()
        for value, var in self.filter_vars.get(filter_key, {}).items():
            try:
                if bool(var.get()):
                    selected.add(value)
            except Exception:
                continue
        return selected

    def _log_active_value_filters(self) -> None:
        """Logga vilka filtervärden som är aktiva för nästa körning."""
        parts: list[str] = []
        for filter_key in ("bolag", "ordertyp"):
            available = self.filter_options.get(filter_key, [])
            if not available:
                continue
            selected = self._get_selected_filter_values(filter_key)
            if selected:
                selected_labels = [self._display_filter_value(v) for v in available if v in selected]
                parts.append(f"{self.filter_titles[filter_key]}: {', '.join(selected_labels)}")
            else:
                parts.append(f"{self.filter_titles[filter_key]}: inga val")
        if parts:
            self._log("Aktiva filter: " + " | ".join(parts))
        else:
            self._log("Aktiva filter: inga filterkolumner hittades i uppladdade filer.")

    def _apply_value_filters(self, df: pd.DataFrame, source_name: str, log_result: bool = True) -> pd.DataFrame:
        """
        Filtrera DataFrame med valda Bolag/Ordertyp-värden.
        AND-logik används när båda kolumnerna finns i samma DataFrame.
        """
        if not isinstance(df, pd.DataFrame):
            return df
        out_df = df.copy()
        before = len(out_df)
        applied_groups: list[str] = []
        for filter_key in ("bolag", "ordertyp"):
            available = self.filter_options.get(filter_key, [])
            if not available:
                continue
            col = self._find_filter_column(out_df, filter_key)
            if not col or col not in out_df.columns:
                continue
            selected = self._get_selected_filter_values(filter_key)
            if not selected:
                out_df = out_df.iloc[0:0].copy()
                applied_groups.append(f"{self.filter_titles[filter_key]} ({col}): inga val")
                break
            normalized_series = out_df[col].map(self._normalize_filter_value)
            out_df = out_df[normalized_series.isin(selected)].copy()
            applied_groups.append(f"{self.filter_titles[filter_key]} ({col})")
        after = len(out_df)
        if log_result:
            if applied_groups:
                self._log(f"Filter {source_name}: {before} -> {after} rader ({', '.join(applied_groups)})")
            else:
                self._log(f"Filter {source_name}: {before} -> {after} rader (ingen matchande filterkolumn)")
        return out_df

    def _ordersaldo_norm(self, value: str) -> str:
        """Normalisera kolumnnamn för robust matchning."""
        return _ordersaldo_norm(value)

    def _ordersaldo_find_col(self, df: pd.DataFrame, candidates: list[str], used_cols: set[str]) -> Optional[str]:
        """Hitta kolumn via exakt/fuzzy match mot kandidater."""
        return _ordersaldo_find_col(df, candidates, used_cols)

    def _load_ordersaldo_source_df(self, source_name: str = "OrderSaldo5 (beställningslinjer)") -> Optional[pd.DataFrame]:
        """Läs beställningslinjer och applicera aktiva filter för ordersaldo-funktioner."""
        path = self.orders_var.get().strip() if hasattr(self, "orders_var") else ""
        if not path:
            return None
        df = self._read_tabular_for_filter_scan(path)
        if not isinstance(df, pd.DataFrame) or df.empty:
            return None
        try:
            df = self._apply_value_filters(df, source_name, log_result=False)
        except Exception:
            pass
        return df

    def _load_ordersaldo_utbest_map(self) -> Dict[str, float]:
        """Läs utbeställt per artikel från automation/saldofilen."""
        saldo_path = self.automation_var.get().strip() if hasattr(self, "automation_var") else ""
        if not saldo_path:
            return {}
        try:
            saldo_df_raw = _clean_columns(
                pd.read_csv(saldo_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
            )
        except Exception:
            return {}
        return utbest_per_article(saldo_df_raw)

    def _refresh_ordersaldo_from_orders(self) -> None:
        """Beräkna data för Kompletta ordrar/Påfyllningsbehov från beställningslinjer."""
        self.ordersaldo_list1_values = []
        self.ordersaldo_list2_values = []

        def _disable_all() -> None:
            self.ordersaldo_copy_list1_btn.configure(state="disabled")
            self.ordersaldo_copy_list2_btn.configure(state="disabled")
            try:
                self.vecka27_btn.configure(state="disabled")
            except Exception:
                pass

        df = self._load_ordersaldo_source_df()
        if not isinstance(df, pd.DataFrame) or df.empty:
            _disable_all()
            return

        column_names = _find_ordersaldo_columns(df, self.ordersaldo_column_candidates)
        order_col = column_names.get("order")
        article_col = column_names.get("article")
        demand_col = column_names.get("demand")
        pick_col = column_names.get("pick")

        # Vecka 27 behöver bara order + artikel + antal (inte pick).
        try:
            self.vecka27_btn.configure(
                state="normal" if (order_col and article_col and demand_col) else "disabled"
            )
        except Exception:
            pass

        if not order_col or not article_col or not demand_col or not pick_col:
            self.ordersaldo_copy_list1_btn.configure(state="disabled")
            self.ordersaldo_copy_list2_btn.configure(state="disabled")
            return

        try:
            complete_orders, holistic_short = compute_ordersaldo_data(
                df,
                utbest_map=self._load_ordersaldo_utbest_map(),
                column_names=column_names,
            )
        except Exception:
            self.ordersaldo_copy_list1_btn.configure(state="disabled")
            self.ordersaldo_copy_list2_btn.configure(state="disabled")
            return

        self.ordersaldo_list1_values = complete_orders
        self.ordersaldo_list2_values = sorted(holistic_short.index.astype(str).tolist())

        self.ordersaldo_copy_list1_btn.configure(state="normal" if self.ordersaldo_list1_values else "disabled")
        self.ordersaldo_copy_list2_btn.configure(state="normal" if self.ordersaldo_list2_values else "disabled")

    def copy_ordersaldo_list1(self) -> None:
        """Kopiera ordernummer från Lista1."""
        if not self.ordersaldo_list1_values:
            messagebox.showinfo(APP_TITLE, "Lista1 är tom.")
            return
        copied_count = len(self.ordersaldo_list1_values)
        try:
            self.master.clipboard_clear()
            self.master.clipboard_append("\n".join(self.ordersaldo_list1_values))
            self.master.update()
        except Exception:
            pass
        messagebox.showinfo(APP_TITLE, f"{copied_count} ordernummer kopierade.")

    def copy_ordersaldo_list2(self) -> None:
        """Kopiera artikelnummer från Lista2."""
        if not self.ordersaldo_list2_values:
            messagebox.showinfo(APP_TITLE, "Lista2 är tom.")
            return
        copied_count = len(self.ordersaldo_list2_values)
        try:
            self.master.clipboard_clear()
            self.master.clipboard_append("\n".join(self.ordersaldo_list2_values))
            self.master.update()
        except Exception:
            pass
        messagebox.showinfo(APP_TITLE, f"{copied_count} artikelnummer kopierade.")

    def run_lyx(self) -> None:
        """LYX: Kopiera artikelnummer där (plocksaldo + utbeställt) ≤ 20 % av max buffertantal."""
        saldo_path = self.automation_var.get().strip()
        self._track_feature("lyx", "run_started")
        if not saldo_path:
            messagebox.showinfo(APP_TITLE, "Ladda upp saldofil först.")
            return

        max_csv = _find_lyx_max_csv()
        if max_csv is None:
            messagebox.showerror(APP_TITLE, "Kunde inte hitta lowfreqdata/buffertpall/artikel_max.csv.")
            return

        try:
            saldo_df = _clean_columns(pd.read_csv(saldo_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig"))
            max_df = _clean_columns(pd.read_csv(str(max_csv), dtype=str, sep=None, engine="python", encoding="utf-8-sig"))
            lyx_arts, filtered_row_count = compute_lyx_articles(saldo_df, max_df)
            if filtered_row_count == 0:
                messagebox.showinfo(APP_TITLE, "Ingen data kvar efter filtrering (Bolag=MG & Plockplats ej tom).")
                return

            if not lyx_arts:
                messagebox.showinfo(APP_TITLE, "Inga artiklar med (plocksaldo + utbeställt) ≤ 20 % av max buffertantal.")
                return

            self.master.clipboard_clear()
            self.master.clipboard_append("\n".join(lyx_arts))
            self.master.update()
            self._track_feature("lyx", "run_completed", copied_articles=int(len(lyx_arts)))
            messagebox.showinfo(APP_TITLE, f"{len(lyx_arts)} artikelnummer kopierade.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"LYX-beräkning misslyckades:\n{e}")

    def run_pafyllnadsprio(self) -> None:
        """Beräkna Påfyllnadsprio och öppna resultatet i en temporär Excel-fil."""
        orders_path = self.orders_var.get().strip() if hasattr(self, "orders_var") else ""
        self._track_feature("pafyllnadsprio", "run_started")
        if not orders_path:
            messagebox.showinfo(APP_TITLE, "Ladda upp beställningslinjer först.")
            return

        max_csv = _find_lyx_max_csv()
        if max_csv is None:
            messagebox.showerror(APP_TITLE, "Kunde inte hitta lowfreqdata/buffertpall/artikel_max.csv.")
            return

        try:
            orders_df = self._load_ordersaldo_source_df("Påfyllnadsprio (beställningslinjer)")
            if not isinstance(orders_df, pd.DataFrame) or orders_df.empty:
                messagebox.showinfo(APP_TITLE, "Beställningslinjefilen är tom eller kunde inte läsas.")
                return

            column_names = _find_ordersaldo_columns(orders_df, self.ordersaldo_column_candidates)
            saldo_path = self.automation_var.get().strip() if hasattr(self, "automation_var") else ""
            if not saldo_path:
                self._log("Påfyllnadsprio: ingen saldofil uppladdad, utbeställt antas vara 0.")
            utbest_map = self._load_ordersaldo_utbest_map()
            _, shortage_df = compute_ordersaldo_data(
                orders_df,
                utbest_map=utbest_map,
                column_names=column_names,
            )
            if shortage_df.empty:
                messagebox.showinfo(APP_TITLE, "Inga artiklar med underskott.")
                return

            max_df = _clean_columns(pd.read_csv(str(max_csv), dtype=str, sep=None, engine="python", encoding="utf-8-sig"))
            overview_path = self.overview_var.get().strip() if hasattr(self, "overview_var") else ""

            if overview_path:
                try:
                    overview_df = self._read_tabular_for_filter_scan(overview_path)
                    if not isinstance(overview_df, pd.DataFrame) or overview_df.empty:
                        raise ValueError("Orderöversikten är tom eller kunde inte läsas.")
                    try:
                        overview_df = self._apply_value_filters(
                            overview_df,
                            "Påfyllnadsprio (orderöversikt)",
                            log_result=False,
                        )
                    except Exception:
                        pass
                    if overview_df.empty:
                        raise ValueError("Inga rader kvar efter aktiva filter i orderöversikten.")

                    report_df, bold_cells, log_lines, missing_reference_count, window_map_df = build_pafyllnadsprio_lastningsfonster_report(
                        orders_df,
                        shortage_df,
                        overview_df,
                        max_df,
                        column_names=column_names,
                    )
                    path = _open_df_in_excel_with_bold_cells(
                        {
                            "Påfyllnadsprio": report_df,
                            "Lastningsfönster": window_map_df,
                        },
                        sheet_name="Påfyllnadsprio",
                        bold_sheet_name="Påfyllnadsprio",
                        label="pafyllnadsprio",
                        bold_cells=bold_cells,
                    )
                    self._log("Påfyllnadsprio: använder lastningsfönster från orderöversikten.")
                    self._log(f"Öppnade Påfyllnadsprio i Excel (temporär fil): {path}")
                    for line in log_lines:
                        self._log(f"[Påfyllnadsprio] {line}")
                    if missing_reference_count:
                        messagebox.showinfo(
                            APP_TITLE,
                            (
                                "För bättre träff, ladda upp buffertpallar.csv. "
                                f"{missing_reference_count} artiklar saknade referensvärde och placerades i PRIO 5."
                            ),
                        )
                    return
                except Exception as overview_error:
                    self._log(
                        "Påfyllnadsprio: kunde inte använda orderöversikten för lastningsfönster "
                        f"({overview_error}). Använder fallback utan orderöversikt."
                    )
                    messagebox.showwarning(
                        APP_TITLE,
                        (
                            "Orderöversikten kunde inte användas för lastningsfönster.\n"
                            "Påfyllnadsprio körs i fallback-läge utan orderöversikt.\n\n"
                            f"{overview_error}"
                        ),
                    )

            report_df, missing_reference_count = build_pafyllnadsprio_report(shortage_df, max_df)
            path = _open_df_in_excel({"Påfyllnadsprio": report_df}, label="pafyllnadsprio")
            self._log("Påfyllnadsprio: använder fallback utan orderöversikt.")
            self._log(f"Öppnade Påfyllnadsprio i Excel (temporär fil): {path}")
            if missing_reference_count:
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        "För bättre träff, ladda upp buffertpallar.csv. "
                        f"{missing_reference_count} artiklar saknade referensvärde och placerades i PRIO 5."
                    ),
                )
        except KeyError as e:
            messagebox.showerror(APP_TITLE, str(e))
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Påfyllnadsprio misslyckades:\n{e}")

    @staticmethod
    def _vecka27_fmt_qty(q: float) -> str:
        """Formatera antal: heltal utan decimaler, annars float."""
        try:
            f = float(q)
        except Exception:
            return str(q)
        return str(int(f)) if f.is_integer() else str(f)

    def run_vecka27_check(self) -> None:
        """Kontrollera tak/hus vs gräsklippare per order. Vid avvikelse: öppna .txt-fil."""
        path = self.orders_var.get().strip() if hasattr(self, "orders_var") else ""
        self._track_feature("vecka27_check", "run_started")
        if not path:
            messagebox.showinfo(APP_TITLE, "Ladda upp beställningslinjefilen först.")
            return

        df = self._read_tabular_for_filter_scan(path)
        if not isinstance(df, pd.DataFrame) or df.empty:
            messagebox.showinfo(APP_TITLE, "Beställningslinjefilen är tom eller kunde inte läsas.")
            return

        try:
            df = self._apply_value_filters(df, "Vecka 27 (beställningslinjer)", log_result=False)
        except Exception:
            pass

        if df.empty:
            messagebox.showinfo(APP_TITLE, "Inga rader kvar efter aktiva filter.")
            return

        used: set[str] = set()
        order_col = self._ordersaldo_find_col(df, self.ordersaldo_column_candidates["order"], used)
        if order_col:
            used.add(order_col)
        article_col = self._ordersaldo_find_col(df, self.ordersaldo_column_candidates["article"], used)
        if article_col:
            used.add(article_col)
        demand_col = self._ordersaldo_find_col(df, self.ordersaldo_column_candidates["demand"], used)

        if not order_col or not article_col or not demand_col:
            messagebox.showerror(
                APP_TITLE,
                "Hittar inte order-, artikel- eller antalskolumn i beställningsfilen.",
            )
            return

        work = df[[order_col, article_col, demand_col]].copy()
        work[order_col] = work[order_col].astype(str).str.strip()
        work[article_col] = work[article_col].astype(str).str.strip()
        work[demand_col] = work[demand_col].map(to_num).astype(float)

        # Summera per (order, artikel) - en order kan ha samma artikel på flera rader.
        grp = work.groupby([order_col, article_col])[demand_col].sum(min_count=1)

        deviations: list[str] = []
        for order_id, sub in grp.groupby(level=0):
            if not str(order_id).upper().startswith("PR"):
                continue
            art_qty: dict[str, float] = {}
            for (_, art), qty in sub.items():
                if pd.notna(qty):
                    art_qty[str(art)] = float(qty)
            for roof, mowers in VECKA27_ROOF_TO_MOWERS.items():
                roof_qty = art_qty.get(roof, 0.0)
                if roof_qty <= 0:
                    continue  # Tak saknas helt - ingen kontroll, ingen varning.
                mower_qty = sum(art_qty.get(m, 0.0) for m in mowers)
                if mower_qty < roof_qty:
                    mower_list = "/".join(sorted(mowers))
                    deviations.append(
                        f"Order {order_id} har {self._vecka27_fmt_qty(roof_qty)} st av {roof} "
                        f"men endast {self._vecka27_fmt_qty(mower_qty)} st gräsklippare av {mower_list}."
                    )

        if not deviations:
            self._track_feature("vecka27_check", "run_completed", deviation_count=0)
            messagebox.showinfo(APP_TITLE, "Allt stämmer för Vecka 27.")
            try:
                self._log("Vecka 27: inga avvikelser.")
            except Exception:
                pass
            return

        body = "Hej Lina!\n" + "\n".join(deviations) + "\nHur gör vi med denna/dessa?\n"
        try:
            tmp_path = _open_text_in_editor(body, label="vecka27")
            self._log(f"Vecka 27: {len(deviations)} avvikelse(r) - öppnade {tmp_path}")
            self._track_feature("vecka27_check", "run_completed", deviation_count=int(len(deviations)))
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte öppna Vecka 27-rapport:\n{e}")

    def _on_eftersok_input_changed(self, *_args) -> None:
        """Uppdatera blå knappstatus när inköpsnummer/artikelnummer ändras."""
        try:
            self._update_action_buttons_state()
        except Exception:
            pass

    def _update_action_buttons_state(self) -> None:
        """Aktivera/inaktivera blå actions beroende på vilka indatafiler som finns."""
        missing_per_button: dict[ttk.Button, list[str]] = {}
        for btn, requirements in self._action_requirements.items():
            missing_labels: list[str] = []
            for file_key, file_label in requirements:
                var = self.file_vars.get(file_key)
                if not var or not var.get().strip():
                    missing_labels.append(file_label)
            if btn == getattr(self, "eftersok_btn", None):
                try:
                    if not str(self.eftersok_purchase_var.get()).strip():
                        missing_labels.append("Inköpsnummer")
                    if not str(self.eftersok_article_var.get()).strip():
                        missing_labels.append("Artikelnummer")
                except Exception:
                    missing_labels.extend(["Inköpsnummer", "Artikelnummer"])
            try:
                btn.configure(state="normal" if not missing_labels else "disabled")
            except Exception:
                pass
            missing_per_button[btn] = missing_labels
        self._action_missing_files = missing_per_button
        self._hide_hover_tooltip()

    def _on_action_button_hover(self, event) -> None:
        """Visa tooltip med saknade filer när en inaktiv blå knapp hovras."""
        btn = event.widget
        try:
            state = str(btn.cget("state")).strip().lower()
        except Exception:
            state = "normal"
        missing = self._action_missing_files.get(btn, [])
        if state != "disabled" or not missing:
            self._hide_hover_tooltip()
            return
        tip_text = "Saknar filer: " + ", ".join(missing)
        self._show_hover_tooltip(tip_text, event.x_root + 12, event.y_root + 10)

    def _on_open_button_hover(self, event) -> None:
        """Visa tooltip med vad som maste goras for att kunna oppna Excel."""
        btn = event.widget
        hint = self._open_button_hints.get(btn, "")
        if not hint:
            self._hide_hover_tooltip()
            return
        try:
            state = str(btn.cget("state")).strip().lower()
        except Exception:
            state = "normal"
        if state != "disabled":
            self._hide_hover_tooltip()
            return
        self._show_hover_tooltip(hint, event.x_root + 12, event.y_root + 10)

    def _show_hover_tooltip(self, text: str, x_root: int, y_root: int) -> None:
        """Rendera en enkel tooltip nära muspekaren."""
        if not text:
            self._hide_hover_tooltip()
            return
        tooltip = self._hover_tooltip
        if tooltip is None or not tooltip.winfo_exists():
            tooltip = tk.Toplevel(self)
            tooltip.wm_overrideredirect(True)
            try:
                tooltip.wm_attributes("-topmost", True)
            except Exception:
                pass
            label = tk.Label(
                tooltip,
                text=text,
                bg="#1f2933",
                fg="white",
                relief="solid",
                borderwidth=1,
                padx=6,
                pady=4,
                justify="left",
            )
            label.pack()
            self._hover_tooltip = tooltip
            self._hover_tooltip_label = label
        else:
            label = self._hover_tooltip_label
            if label is not None:
                label.config(text=text)
        if self._hover_tooltip is not None and self._hover_tooltip.winfo_exists():
            self._hover_tooltip.geometry(f"+{x_root}+{y_root}")

    def _hide_hover_tooltip(self, event=None) -> None:
        """Dolj och nollställ tooltip."""
        tooltip = self._hover_tooltip
        if tooltip is not None:
            try:
                tooltip.destroy()
            except Exception:
                pass
        self._hover_tooltip = None
        self._hover_tooltip_label = None

    def _set_log_splitter_default_position(self) -> None:
        """Placera logg-delaren mitt/balanserat vid start."""
        splitter = getattr(self, "log_splitter", None)
        if splitter is None:
            return
        try:
            total_h = int(splitter.winfo_height())
        except Exception:
            total_h = 0
        if total_h <= 120:
            try:
                self.after(120, self._set_log_splitter_default_position)
            except Exception:
                pass
            return
        # Ge loggen lite mer yta som standard men lämna god marginal åt summeringen.
        desired = int(total_h * 0.62)
        min_top = 90
        min_bottom = 110
        max_top = max(min_top, total_h - min_bottom)
        pos = max(min_top, min(desired, max_top))
        try:
            splitter.sashpos(0, pos)
        except Exception:
            pass

    def _set_file_path(self, file_type: str, path: str, source: str = "dialog") -> None:
        path = str(path or "").strip()
        if not path:
            return

        var = self.file_vars.get(file_type)
        if var is not None:
            var.set(path)
        elif file_type == "wms_receive":
            self.wms_receive_var.set(path)
        elif file_type == "wms_booking":
            self.wms_booking_var.set(path)
        elif file_type == "wms_trans":
            self.wms_trans_var.set(path)
        elif file_type == "wms_pick":
            self.wms_pick_var.set(path)
        elif file_type == "wms_correct":
            self.wms_correct_var.set(path)
        else:
            return

        if file_type == "prognos":
            try:
                self._load_prognos(path)
            except Exception:
                pass
        elif file_type == "campaign":
            try:
                self._load_campaign(path)
            except Exception:
                pass

        self._track_event(
            "input_selected",
            file_type=file_type,
            source=source,
            extension=Path(path).suffix.lower(),
        )

    def clear_file(self, file_type: str) -> None:
        """
        Töm filvalet för angiven filtyp och uppdatera ikonerna.
        """
        try:
            var = self.file_vars.get(file_type)
            if var:
                var.set("")
            # Rensa även eventuellt laddad prognos eller kampanjdata
            if file_type == "prognos":
                self._prognos_df = None
            if file_type == "campaign":
                self._campaign_raw = None
                self._campaign_norm = None
            if file_type in ("prognos", "campaign"):
                has_prognos = isinstance(self._prognos_df, pd.DataFrame) and not self._prognos_df.empty
                has_campaign = isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty
                self.open_prognos_btn.configure(state="normal" if (has_prognos or has_campaign) else "disabled")
        except Exception:
            pass
        self.update_file_status_icons()

    def open_files_dialog(self, event=None) -> None:
        """
        Öppna en fil-dialog för att välja en eller flera filer. Filtyperna
        identifieras automatiskt och tilldelas rätt fält.
        """
        paths = filedialog.askopenfilenames(title="Välj filer", filetypes=[
            ("CSV och Excel", "*.csv *.xlsx *.xlsm *.xls"),
            ("Alla filer", "*.*")
        ])
        if not paths:
            return
        for p in paths:
            p = str(p)
            file_type = self._detect_file_type(p)
            if file_type == "orders":
                self._set_file_path("orders", p, source="file_dialog")
            elif file_type == "buffer":
                self._set_file_path("buffer", p, source="file_dialog")
            elif file_type == "automation":
                self._set_file_path("automation", p, source="file_dialog")
            elif file_type == "item":
                self._set_file_path("item", p, source="file_dialog")
            elif file_type == "prognos":
                self._set_file_path("prognos", p, source="file_dialog")
            elif file_type == "campaign":
                self._set_file_path("campaign", p, source="file_dialog")
            elif file_type == "overview":
                self._set_file_path("overview", p, source="file_dialog")
            elif file_type == "dispatch":
                self._set_file_path("dispatch", p, source="file_dialog")
            elif file_type == "wms_receive":
                self._set_file_path("wms_receive", p, source="file_dialog")
            elif file_type == "wms_booking":
                self._set_file_path("wms_booking", p, source="file_dialog")
            elif file_type == "wms_trans":
                self._set_file_path("wms_trans", p, source="file_dialog")
            elif file_type == "wms_pick":
                self._set_file_path("wms_pick", p, source="file_dialog")
            elif file_type == "wms_correct":
                self._set_file_path("wms_correct", p, source="file_dialog")
            else:
                try:
                    self._log(f"Okänd filtyp: {p}")
                except Exception:
                    pass
        # Uppdatera ikoner efter alla filer har satts
        self.update_file_status_icons()

    def _load_wms_analyzer_class(self):
        """Ladda WMSAnalyzerUpdated från wms_sök79.py."""
        if self._wms_analyzer_cls is not None:
            return self._wms_analyzer_cls
        search_roots: list[Path] = []
        for root in (_runtime_root(), _bundle_root(), Path(__file__).resolve().parent):
            if root not in search_roots:
                search_roots.append(root)
        module_path: Optional[Path] = None
        for root in search_roots:
            for filename in ("wms_sok79.py", "wms_sök79.py"):
                candidate = root / filename
                if candidate.exists():
                    module_path = candidate
                    break
            if module_path is not None:
                break

        if module_path is None:
            raise FileNotFoundError(
                "Hittar inte wms_sök79.py i appmappen."
            )
        spec = importlib.util.spec_from_file_location("wms_sok79_module", module_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Kunde inte ladda filen: {module_path}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        analyzer_cls = getattr(module, "WMSAnalyzerUpdated", None)
        if analyzer_cls is None:
            raise AttributeError("WMSAnalyzerUpdated saknas i wms_sök79.py")
        self._wms_analyzer_cls = analyzer_cls
        return analyzer_cls

    def run_eftersok(self) -> None:
        """Kör DEMO Eftersök med logiken från wms_sök79.py."""
        purchase = str(self.eftersok_purchase_var.get()).strip()
        article = str(self.eftersok_article_var.get()).strip()
        if not purchase or not article:
            messagebox.showwarning(APP_TITLE, "Både inköpsnummer och artikelnummer måste fyllas i.")
            return

        receive_path = str(self.wms_receive_var.get()).strip()
        if not receive_path:
            messagebox.showwarning(APP_TITLE, "Ladda minst Mottagningslogg (CSV) för att köra Eftersök.")
            return

        self._track_feature("eftersok", "run_started")
        self.last_eftersok_df = None
        self.last_eftersok_report = None
        self.last_eftersok_path = None
        try:
            self.open_eftersok_btn.configure(state="disabled")
        except Exception:
            pass

        wms_paths = {
            "wms_receive": str(self.wms_receive_var.get()).strip(),
            "wms_booking": str(self.wms_booking_var.get()).strip(),
            "wms_buffert": str(self.buffer_var.get()).strip(),
            "wms_trans": str(self.wms_trans_var.get()).strip(),
            "wms_pick": str(self.wms_pick_var.get()).strip(),
            "wms_correct": str(self.wms_correct_var.get()).strip(),
        }

        try:
            analyzer_cls = self._load_wms_analyzer_class()
        except Exception as e:
            self._track_feature("eftersok", "run_failed", stage="load_analyzer")
            messagebox.showerror(APP_TITLE, f"Kunde inte ladda wms_sök79.py:\n{e}")
            return

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                for key, src in wms_paths.items():
                    if not src:
                        continue
                    dst_name = self.wms_expected_filenames.get(key)
                    if not dst_name:
                        continue
                    dst = os.path.join(tmpdir, dst_name)
                    shutil.copy(src, dst)
                analyzer = analyzer_cls(data_path=tmpdir)
                report_text = str(analyzer.analyze(purchase, article) or "").strip()
        except Exception as e:
            self._track_feature("eftersok", "run_failed", stage="analyze")
            messagebox.showerror(APP_TITLE, f"Ett fel uppstod under Eftersök:\n{e}")
            return

        if not report_text:
            report_text = "Ingen text returnerades från Eftersök."

        report_lines = report_text.splitlines()
        if not report_lines:
            report_lines = [report_text]

        self.last_eftersok_report = report_text
        self.last_eftersok_df = pd.DataFrame({"Rapport": report_lines})
        try:
            self.open_eftersok_btn.configure(state="normal")
        except Exception:
            pass
        self._log(f"Eftersök klart för inköpsnummer {purchase}, artikelnummer {article}.")
        self._log(f"Eftersök-rader: {len(report_lines)}")
        self._track_feature("eftersok", "run_completed", report_lines=int(len(report_lines)))

    def open_eftersok_in_excel(self) -> None:
        """Öppna senaste Eftersök-resultatet i Excel."""
        if isinstance(self.last_eftersok_df, pd.DataFrame) and not self.last_eftersok_df.empty:
            try:
                path = _open_df_in_excel({"Eftersök": self.last_eftersok_df.copy()}, label="eftersok")
                self.last_eftersok_path = path
                self._log(f"Öppnade Eftersök i Excel (temporär fil): {path}")
                self._track_feature("eftersok", "opened_result")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna Eftersök i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget Eftersök-resultat att öppna. Kör Eftersök först.")

    def open_chunked_values_in_excel(self) -> None:
        """
        Dela inklistrade värden i kolumner (default 2000 rader per kolumn)
        och öppna resultatet direkt i en temporär Excel-fil.
        """
        try:
            lines = [r.strip() for r in self.split_input_text.get("1.0", tk.END).splitlines() if r.strip()]
        except Exception:
            lines = []
        if not lines:
            messagebox.showwarning(APP_TITLE, "Klistra in värden först (en per rad).")
            return
        try:
            chunk_size = int(str(self.split_chunk_var.get()).strip())
            if chunk_size <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror(APP_TITLE, "Antal per kolumn måste vara ett heltal > 0.")
            return

        chunks = [lines[i:i + chunk_size] for i in range(0, len(lines), chunk_size)]
        out_cols: dict[str, pd.Series] = {}
        for idx, chunk in enumerate(chunks, start=1):
            out_cols[f"Kolumn {idx}"] = pd.Series([str(v) for v in chunk], dtype="string")
        out_df = pd.DataFrame(out_cols).fillna("")
        try:
            path = _open_df_in_excel({"Delade värden": out_df}, label="2000tal_split")
            self._log(f"2000-tal: öppnade temporär Excel-fil: {path}")
            self._track_feature("chunked_values", "opened_result", rows=int(len(out_df)))
            messagebox.showinfo(APP_TITLE, "Excel-filen öppnades direkt. Spara i Excel om du vill behålla den.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte skapa/öppna Excel-filen:\n{e}")


    def open_result_in_excel(self) -> None:
        if isinstance(self.last_result_df, pd.DataFrame) and not self.last_result_df.empty:
            try:
                path = _open_df_in_excel({"Allokerade order": self.last_result_df.copy()}, label="allocated_orders")
                self._log(f"Öppnade resultat i Excel (temporär fil): {path}")
                self._track_feature("allocation", "opened_result", result_type="allocated_orders")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna resultat i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget resultat att öppna ännu. Kör allokeringen först.")

    def open_nearmiss_in_excel(self) -> None:
        if isinstance(self.last_nearmiss_instead_df, pd.DataFrame) and not self.last_nearmiss_instead_df.empty:
            try:
                nm_df = self.last_nearmiss_instead_df.copy()
                if "Artikel" in nm_df.columns:
                    nm_df = nm_df.drop_duplicates(subset=["Artikel"], keep="first").reset_index(drop=True)
                pct_str = f"{int(NEAR_MISS_PCT * 100)}%"
                sheet_name = f"Near-miss {pct_str} (unika artiklar)"
                label = f"near_miss_{int(NEAR_MISS_PCT * 100)}pct"
                path = _open_df_in_excel({sheet_name: nm_df}, label=label)
                self._log(f"Öppnade near-miss (INSTEAD R or A) i Excel (temporär fil): {path}")
                self._track_feature("allocation", "opened_result", result_type="near_miss")
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
                ps_df = self._pallet_spaces_df.copy()
                path = _open_df_in_excel({"Pallplatser": ps_df}, label="pallplatser")
                self._log(f"Öppnade pallplatser i Excel (temporär fil): {path}")
                self._track_feature("allocation", "opened_result", result_type="pallet_spaces")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna pallplatser i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns ingen pallplatsrapport att öppna ännu. Kör allokeringen först.")

    def pick_prognos(self) -> None:
        """Visa en filväljare för att välja en prognosfil (XLSX)."""
        path = filedialog.askopenfilename(title="Välj prognos (XLSX)", filetypes=[("Excel", "*.xlsx"), ("Alla filer","*.*")])
        if path:
            self._set_file_path("prognos", path, source="picker")
            # Uppdatera statusikoner även om prognosfilen laddas in via egen knapp
            try:
                self.update_file_status_icons()
            except Exception:
                pass
        else:
            self._prognos_df = None
            self.open_prognos_btn.configure(state="disabled")

    def _load_prognos(self, path: str) -> None:
        """Läs in prognosfilen och aktivera knappen för öppning."""
        try:
            df = read_prognos_xlsx(path)
            self._prognos_df = df
            self._track_event(
                "input_loaded",
                file_type="prognos",
                rows=int(len(df)),
                unique_articles=int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else int(len(df)),
            )
            try:
                n_art = int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else len(df)
                self._log(f"Prognos inläst: {len(df)} rader, {n_art} artiklar.")
            except Exception:
                self._log(f"Prognos inläst: {len(df)} rader.")
            self.open_prognos_btn.configure(state="normal")
        except Exception as e:
            self._prognos_df = None
            self.open_prognos_btn.configure(state="disabled")
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa prognosfilen:\n{e}")

    def pick_campaign(self) -> None:
        """Visa en filväljare för att välja en kampanjvolymfil (XLSX)."""
        path = filedialog.askopenfilename(title="Välj kampanjvolymer (XLSX)", filetypes=[("Excel", "*.xlsx"), ("Alla filer", "*.*")])
        if path:
            self._set_file_path("campaign", path, source="picker")
            # Uppdatera statusikoner även när kampanjfilen laddas in via egen knapp
            try:
                self.update_file_status_icons()
            except Exception:
                pass
        else:
            self._campaign_norm = None

    def _load_campaign(self, path: str) -> None:
        """Läs in kampanjvolymer och lagra den normaliserade datan."""
        try:
            df = read_campaign_xlsx(path)
            self._campaign_norm = df
            self._track_event(
                "input_loaded",
                file_type="campaign",
                rows=int(len(df)),
                unique_articles=int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else int(len(df)),
            )
            try:
                n_art = int(df["Artikelnummer"].nunique()) if "Artikelnummer" in df.columns else len(df)
                self._log(f"Kampanjvolymer inlästa: {len(df)} rader, {n_art} artiklar.")
            except Exception:
                self._log(f"Kampanjvolymer inlästa: {len(df)} rader.")
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
        has_prognos = isinstance(self._prognos_df, pd.DataFrame) and not self._prognos_df.empty
        has_campaign = isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty
        if not has_prognos and not has_campaign:
            messagebox.showinfo(APP_TITLE, "Välj och läs in antingen prognosfilen eller kampanjvolymerna först.")
            return
        try:
            if has_prognos:
                combined_df: pd.DataFrame = self._prognos_df.copy()
            else:
                combined_df = pd.DataFrame(columns=["Artikelnummer", "Beskrivning", "Antal styck", "Antal rader", "Antal butiker"])
            if isinstance(self._campaign_norm, pd.DataFrame) and not self._campaign_norm.empty:
                camp_df = self._campaign_norm.copy()
                if isinstance(self._saldo_raw, pd.DataFrame) and not self._saldo_raw.empty:
                    s = self._saldo_raw.copy()
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
                        s = s.loc[s["Robot"] == "Y"]
                        if not s.empty:
                            camp_df = camp_df.merge(s[["Artikelnummer"]], on="Artikelnummer", how="inner")
                        else:
                            camp_df = camp_df.iloc[0:0]
                    else:
                        camp_df = camp_df.iloc[0:0]
                if not camp_df.empty:
                    vol_by_art = camp_df.groupby("Artikelnummer")["Antal styck"].sum().to_dict()
                    combined_df["Artikelnummer"] = combined_df["Artikelnummer"].astype(str).str.strip()
                    combined_df["Antal styck"] = pd.to_numeric(combined_df.get("Antal styck", 0), errors="coerce").fillna(0).astype(int)
                    existing_arts = set(combined_df["Artikelnummer"].astype(str))
                    for art, vol in vol_by_art.items():
                        if art in existing_arts:
                            mask = combined_df["Artikelnummer"] == art
                            combined_df.loc[mask, "Antal styck"] = (combined_df.loc[mask, "Antal styck"].astype(int) + int(vol)).astype(int)
                        else:
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
            report_df, meta = build_prognos_vs_autoplock_report(
                prognos_df=combined_df,
                saldo_norm_df=(self._saldo_raw if isinstance(self._saldo_raw, pd.DataFrame) else None),
                buffer_df=(self._buffer_raw if isinstance(self._buffer_raw, pd.DataFrame) else None),
                exclude_source_ids=None,
                allocated_df=None,
            )
            path = open_prognos_vs_autoplock_excel(report_df, meta)
            msg = f"Prognosrapport skapad ({len(report_df)} rader)."
            if isinstance(meta, dict) and meta.get("partial") == "yes":
                miss = meta.get("missing", "").replace(",", ", ")
                if miss:
                    msg += f" PARTIELL: saknar {miss}."
            self._log(msg)
            self._track_feature(
                "prognos_report",
                "opened_result",
                report_rows=int(len(report_df)),
                partial=bool(isinstance(meta, dict) and meta.get("partial") == "yes"),
            )
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte skapa/öppna prognosrapporten:\n{e}")

    def reset_cache(self) -> None:
        """
        Rensa alla cacher och temporära variabler i applikationen. Detta nollställer
        internt lagrade DataFrames (resultat, near-miss, saldo, item, sales m.m.),
        tömmer loggrutan, återställer summeringstabellen till noll och inaktiverar
        öppna-knapparna. Pathvariabler för filval påverkas inte.
        """
        try:
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
            self.last_eftersok_df = None
            self.last_eftersok_path = None
            self.last_eftersok_report = None

            # Rensa HIB‑kopplingsresultat
            self.last_koppla_df = None
            self.last_koppla_missed_df = None
            self.last_koppla_path = None

            self.log.configure(state="normal")
            self.log.delete("1.0", tk.END)
            self.log.configure(state="disabled")

            try:
                for child in self.summary_table.get_children(""):
                    self.summary_table.delete(child)
            except Exception:
                pass

            # Stäng av alla öppna-knappar inklusive HIB-kopplingsknappen
            for btn in (self.open_result_btn, self.open_nearmiss_btn, self.open_palletspaces_btn, self.open_prognos_btn, self.open_refill_btn, self.open_koppla_btn, self.open_eftersok_btn):
                try:
                    btn.configure(state="disabled")
                except Exception:
                    pass

            try:
                # Rensa alla filval (inklusive orderöversikt) så att texten försvinner från GUI
                self.orders_var.set("")
                self.buffer_var.set("")
                self.automation_var.set("")
                self.item_var.set("")
                self.prognos_var.set("")
                self.campaign_var.set("")
                # Orderöversikten (overview) ska också nollställas vid cache-rensning
                if hasattr(self, "overview_var"):
                    self.overview_var.set("")
                # Rensa även dispatchpallar vid cache-rensning
                if hasattr(self, "dispatch_var"):
                    self.dispatch_var.set("")
                if hasattr(self, "wms_receive_var"):
                    self.wms_receive_var.set("")
                if hasattr(self, "wms_booking_var"):
                    self.wms_booking_var.set("")
                if hasattr(self, "wms_trans_var"):
                    self.wms_trans_var.set("")
                if hasattr(self, "wms_pick_var"):
                    self.wms_pick_var.set("")
                if hasattr(self, "wms_correct_var"):
                    self.wms_correct_var.set("")
                if hasattr(self, "eftersok_purchase_var"):
                    self.eftersok_purchase_var.set("")
                if hasattr(self, "eftersok_article_var"):
                    self.eftersok_article_var.set("")
                # Nollställ eventuella resultat från dispatch-kontrollen
                self.last_dispatch_check_df = None
                self.last_dispatch_check_path = None
                # Inaktivera dispatchkontroll-knappen
                try:
                    self.open_dispatch_check_btn.configure(state="disabled")
                except Exception:
                    pass
            except Exception:
                pass

            # Uppdatera statusikoner efter att alla filval har nollställts
            try:
                self.update_file_status_icons()
            except Exception:
                pass

            self._log("Cache och temporära data har rensats.")
        except Exception:
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
                if isinstance(self._sales_metrics_df, pd.DataFrame) and not self._sales_metrics_df.empty:
                    hp = annotate_refill(hp, self._sales_metrics_df)
                    asr = annotate_refill(asr, self._sales_metrics_df)
                path = _open_df_in_excel({"Refill HP": hp, "Refill AUTOSTORE": asr}, label="refill")
                self._log(f"Öppnade påfyllningspallar (cache) i Excel (temporär fil): {path}")
                self._track_feature("allocation", "opened_result", result_type="refill")
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


    def run_koppla(self) -> None:
        """
        Utför HIB‑koppling.  Läs in beställningslinjer och orderöversikt,
        filtrera enligt de regler som angetts och bygg en resultatlista
        över HIB‑ordrar som behöver uppdateras.  Efter körning lagras
        resultatet i `self.last_koppla_df` och knappen för att öppna
        resultatet aktiveras om det finns något att visa.
        """
        details_path = self.orders_var.get().strip()
        overview_path = self.overview_var.get().strip()
        if not details_path or not overview_path:
            messagebox.showerror(APP_TITLE, "Välj både beställningslinjer och orderöversikt.")
            return
        self._track_feature("hib_koppling", "run_started")
        self._log_active_value_filters()
        try:
            # Läs in beställningslinjer
            details_df = pd.read_csv(details_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        except Exception as e:
            self._track_feature("hib_koppling", "run_failed", stage="read_details")
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa beställningslinjer:\n{e}")
            return
        try:
            # Läs in orderöversikt
            overview_df = pd.read_csv(overview_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        except Exception as e:
            self._track_feature("hib_koppling", "run_failed", stage="read_overview")
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa orderöversikten:\n{e}")
            return
        try:
            details_df = self._apply_value_filters(details_df, "Beställningslinjer (HIB-koppling)")
            overview_df = self._apply_value_filters(overview_df, "Orderöversikt (HIB-koppling)")
        except Exception:
            pass
        # Beräkna ändringar och missade avgångar
        try:
            changes_df = compute_hib_koppling(details_df, overview_df)
        except Exception as e:
            self._track_feature("hib_koppling", "run_failed", stage="compute_changes")
            messagebox.showerror(APP_TITLE, f"Fel vid beräkning av HIB‑kopplingen:\n{e}")
            return
        try:
            missed_df = compute_missed_departures(details_df, overview_df)
        except Exception as e:
            self._track_feature("hib_koppling", "run_failed", stage="compute_missed_departures")
            messagebox.showerror(APP_TITLE, f"Fel vid beräkning av missade avgångar:\n{e}")
            missed_df = pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        # Spara resultat
        self.last_koppla_df = changes_df.copy() if isinstance(changes_df, pd.DataFrame) else pd.DataFrame()
        self.last_koppla_missed_df = missed_df.copy() if isinstance(missed_df, pd.DataFrame) else pd.DataFrame()
        # Om varken ändringar eller missade avgångar finns, meddela användaren och stäng av öppna‑knappen
        if (changes_df is None or changes_df.empty) and (missed_df is None or missed_df.empty):
            self.open_koppla_btn.config(state="disabled")
            self._track_feature("hib_koppling", "run_completed", changes_rows=0, missed_rows=0)
            messagebox.showinfo(APP_TITLE, "Inga HIB‑ordrar behöver ändras eller har missat sin avgång.")
            return
        # Det finns något att visa – aktivera öppna‑knappen
        self.open_koppla_btn.config(state="normal")
        # Logga resultatet i loggfönstret
        try:
            if changes_df is not None and not changes_df.empty:
                self._log("HIB‑koppling ändringar:")
                for _, r in changes_df.iterrows():
                    try:
                        ordnr = str(r.get("ordernummer", "")).strip()
                        kundnamn = str(r.get("kundnamn", "")).strip()
                        fields: list[str] = []
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
                            self._log(f"Order {ordnr}{name_part}: {', '.join(fields)}")
                    except Exception:
                        pass
            if missed_df is not None and not missed_df.empty:
                self._log("Missade avgångar:")
                for _, r in missed_df.iterrows():
                    try:
                        ordnr = str(r.get("ordernummer", "")).strip()
                        kundnamn = str(r.get("kundnamn", "")).strip()
                        name_part = f" ({kundnamn})" if kundnamn else ""
                        self._log(f"Order {ordnr}{name_part}: MISSAT SIN AVGÅNG")
                    except Exception:
                        pass
            self._log("HIB‑kopplingen är beräknad och redo att öppnas i Excel.")
            instr_lines = [
                "\nInstruktion:",
                "Ändras i följande ordning",
                "1. Orderdatum",
                "2. Sändningsnummer",
                "3. Zon F på orderlinjerna",
                "4. Samma multi på alla Hibar till samma butik",
                "5. Generera",
                "6. Frisläpp",
            ]
            for line in instr_lines:
                try:
                    self._log(line)
                except Exception:
                    pass
        except Exception:
            # Om loggning misslyckas fortsätter vi utan att avbryta
            self._log("HIB‑kopplingen är beräknad och redo att öppnas i Excel.")


        self._track_feature(
            "hib_koppling",
            "run_completed",
            changes_rows=int(len(changes_df)) if isinstance(changes_df, pd.DataFrame) else 0,
            missed_rows=int(len(missed_df)) if isinstance(missed_df, pd.DataFrame) else 0,
        )

    def open_koppla_in_excel(self) -> None:
        """
        Öppna det senast beräknade HIB‑kopplingsresultatet i en temporär
        Excel‑fil tillsammans med instruktioner.  Om ingen körning har gjorts
        ännu eller om resultatet saknas visas ett informationsmeddelande.
        """
        # Endast öppna i Excel om det finns ändringar eller missade avgångar
        has_changes = isinstance(self.last_koppla_df, pd.DataFrame) and not self.last_koppla_df.empty
        has_missed = isinstance(getattr(self, "last_koppla_missed_df", None), pd.DataFrame) and not getattr(self, "last_koppla_missed_df").empty
        if has_changes or has_missed:
            try:
                instr_lines = [
                    "Ändras i följande ordning",
                    "1. Orderdatum",
                    "2. Sändningsnummer",
                    "3. Zon F på orderlinjerna",
                    "4. Samma multi på alla Hibar till samma butik",
                    "5. Generera",
                    "6. Frisläpp",
                ]
                instructions_df = pd.DataFrame({"Instruktioner": instr_lines})
                sheets: dict[str, pd.DataFrame] = {}
                if has_changes:
                    sheets["Ändringar"] = self.last_koppla_df.copy()
                if has_missed:
                    sheets["Missade avgångar"] = self.last_koppla_missed_df.copy()
                sheets["Instruktion"] = instructions_df
                path = _open_df_in_excel(sheets, label="hib_koppling")
                self.last_koppla_path = path
                self._log(f"Öppnade HIB‑koppling i Excel (temporär fil): {path}")
                self._track_feature("hib_koppling", "opened_result")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna HIB‑koppling i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget HIB‑kopplingsresultat att öppna. Kör HIB‑kopplingen först.")

    def run_overview_check(self) -> None:
        """
        Gå igenom orderöversikten och hitta sändningsnummer som förekommer hos flera kunder
        eller med olika transportörer. Resultatet loggas och kan öppnas i Excel.
        """
        path = self.overview_var.get().strip()
        self._track_feature("overview_check", "run_started")
        if not path:
            messagebox.showerror(APP_TITLE, "Välj orderöversikten först.")
            return
        self._log_active_value_filters()
        try:
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
            if df.shape[1] == 1:
                # Försök tab-separerad om endast en kolumn hittades
                df = pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa orderöversikten:\n{e}")
            return
        # Normalisera kolumnnamn
        df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
        try:
            df = self._apply_value_filters(df, "Orderöversikt (orderkontroll)")
        except Exception:
            pass
        if df.empty:
            self.open_overview_check_btn.config(state="disabled")
            self.last_overview_check_df = pd.DataFrame()
            self.last_hib_status_check_df = pd.DataFrame()
            messagebox.showinfo(APP_TITLE, "Inga rader kvar efter filter i orderöversikten.")
            return
        # Identifiera relevant kolumner
        ship_col = None
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if "sändning" in cl or "sandning" in cl or "sändningsnr" in cl or "sandningsnr" in cl or "sändningsnummer" in cl:
                ship_col = c
                break
        if not ship_col:
            messagebox.showerror(APP_TITLE, "Kunde inte identifiera sändningsnummer-kolumnen i orderöversikten.")
            return
        cust_col = None
        # Försök hitta kundnummerkolumn
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            if "kundnr" in cl or "kundnr." in cl or "kundnummer" in cl:
                cust_col = c
                break
        if not cust_col:
            # Om kundnummer saknas, använd kundnamn
            for c in df.columns:
                if "kund" in c.lower():
                    cust_col = c
                    break
        if not cust_col:
            messagebox.showerror(APP_TITLE, "Kunde inte identifiera kund-kolumnen i orderöversikten.")
            return
        trans_col = None
        for c in df.columns:
            cl = c.lower()
            if "transportör" in cl or "transportor" in cl:
                trans_col = c
                break
        if not trans_col:
            for c in df.columns:
                cl = c.lower().replace(" ", "")
                if "transportörsnr" in cl or "transportorsnr" in cl:
                    trans_col = c
                    break
        # Fyll i tom transportörskolumn om den saknas
        if not trans_col:
            trans_col = "__transport_dummy__"
            df[trans_col] = ""
        # Rensa strängar
        df[ship_col] = df[ship_col].astype(str).str.strip()
        df[cust_col] = df[cust_col].astype(str).str.strip()
        df[trans_col] = df[trans_col].astype(str).str.strip()

        # ----- Ny funktionalitet: hitta ordernummer och kundnamn -----
        # Vi vill kunna visa vilka ordernummer (och deras kundnamn) som ingår i varje sändning.
        # Försök hitta ordernummer-kolumnen i orderöversikten.
        order_col = None
        order_keywords = ["ordernr", "order nr", "ordernummer", "order number", "orderid", "order id"]
        for c in df.columns:
            for kw in order_keywords:
                if kw.replace(" ", "") == c.lower().replace(" ", ""):
                    order_col = c
                    break
            if order_col:
                break
        if not order_col:
            # Om vi inte hittade en exakt matchning, ta första kolumn som innehåller "order"
            for c in df.columns:
                if "order" in c.lower():
                    order_col = c
                    break
        # Bygg en mappning från ordernummer till kundnamn. Prioritera beställningsfilen om sådan finns.
        order_to_customer: Dict[str, str] = {}
        try:
            details_path = getattr(self, "orders_var", tk.StringVar()).get().strip()
            if details_path:
                try:
                    ddf = pd.read_csv(details_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
                    if ddf.shape[1] == 1:
                        ddf = pd.read_csv(details_path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
                    ddf.columns = [str(c).replace("\ufeff", "").strip() for c in ddf.columns]
                    try:
                        ddf = self._apply_value_filters(ddf, "Beställningslinjer (orderkontroll)")
                    except Exception:
                        pass
                    if "Order nr" in ddf.columns and "Kund.1" in ddf.columns:
                        try:
                            order_to_customer = (
                                ddf.groupby("Order nr")["Kund.1"].first()
                                .fillna("")
                                .astype(str)
                                .str.strip()
                                .to_dict()
                            )
                        except Exception:
                            order_to_customer = {}
                except Exception:
                    order_to_customer = {}
        except Exception:
            order_to_customer = {}
        # Om vi fortfarande saknar mappning, använd orderöversikten som fallback
        if not order_to_customer and order_col:
            try:
                order_to_customer = (
                    df.groupby(order_col)[cust_col]
                    .first()
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .to_dict()
                )
            except Exception:
                order_to_customer = {}

        # Identifiera ordertyp- och statuskolumn för extra HIB-kontroll.
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

        def _to_status_num(value: object) -> Optional[int]:
            try:
                raw = str(value).strip().replace(",", ".")
                if not raw:
                    return None
                return int(float(raw))
            except Exception:
                return None

        # Filtrera bort tomma sändningsnummer
        df = df[df[ship_col].astype(str).str.len() > 0].copy()
        if df.empty:
            self.open_overview_check_btn.config(state="disabled")
            self.last_overview_check_df = pd.DataFrame()
            messagebox.showinfo(APP_TITLE, "Orderöversikten innehåller inga sändningsnummer att analysera.")
            return

        # Del 1: befintlig avvikelsekontroll (flera kunder/transportörer per sändning).
        shipment_diff_rows: List[Dict[str, object]] = []
        try:
            grouped = df.groupby(ship_col)
        except Exception:
            messagebox.showerror(APP_TITLE, "Kunde inte gruppera orderöversikten på sändningsnummer.")
            return
        for ship, group in grouped:
            try:
                customers = sorted(set(group[cust_col].dropna().astype(str).str.strip()))
                carriers = sorted(set(group[trans_col].dropna().astype(str).str.strip()))
                # Ta bort tomma strängar
                customers = [c for c in customers if c]
                carriers = [t for t in carriers if t]
                # Sammanställ ordernummer (med kundnamn) för denna sändning
                orders_list: List[str] = []
                if order_col:
                    try:
                        order_vals = sorted(set(group[order_col].dropna().astype(str).str.strip()))
                    except Exception:
                        order_vals = []
                    for o in order_vals:
                        try:
                            nm = order_to_customer.get(o, "")
                        except Exception:
                            nm = ""
                        if nm:
                            orders_list.append(f"{o} ({nm})")
                        else:
                            orders_list.append(o)
                orders_str = ", ".join(orders_list)
                if len(customers) > 1 or len(carriers) > 1:
                    res_row = {
                        "Avvikelsetyp": "Sändningsnr med flera kunder/transportörer",
                        "Sändningsnr": ship,
                        "Unika kunder": len(customers),
                        "Kunder": ", ".join(customers),
                        "Unika transportörer": len(carriers),
                        "Transportörer": ", ".join(carriers),
                        "Antal orderrader": int(len(group)),
                    }
                    if orders_str:
                        res_row["Ordernr (kundnamn)"] = orders_str
                    shipment_diff_rows.append(res_row)
            except Exception:
                continue
        result_df = pd.DataFrame(shipment_diff_rows) if shipment_diff_rows else pd.DataFrame()
        self.last_overview_check_df = result_df.copy() if not result_df.empty else pd.DataFrame()

        # Del 2: HIB-order med status > 31 som saknar matchande butikssändning.
        hib_rows: List[Dict[str, object]] = []
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
                    # Avvikelse bara om ingen av HIB-orderns sändningar finns hos butik.
                    if any(ship_val in store_ships for ship_val in hib_ships):
                        continue

                    kundnamn = ""
                    try:
                        kundnamn = order_to_customer.get(ordnr_str, "")
                    except Exception:
                        kundnamn = ""
                    if not kundnamn:
                        try:
                            kunder = [k for k in group[cust_col].dropna().astype(str).str.strip().tolist() if k]
                            if kunder:
                                kundnamn = kunder[0]
                        except Exception:
                            kundnamn = ""

                    row: Dict[str, object] = {
                        "Ordernr": ordnr_str,
                        "Sändningsnr": ", ".join(hib_ships),
                        "Ordertyp": "HIB",
                        "Status": int(max_status),
                        "Anmärkning": "HIB-order med status > 31 saknar matchande butikssändning",
                    }
                    if kundnamn:
                        row["Kundnamn"] = kundnamn
                    hib_rows.append(row)
            except Exception:
                pass
        hib_check_df = pd.DataFrame(hib_rows) if hib_rows else pd.DataFrame()
        self.last_hib_status_check_df = hib_check_df.copy() if not hib_check_df.empty else pd.DataFrame()

        # Aktivera knappen om någon kontroll gav resultat
        has_any = not result_df.empty or not hib_check_df.empty
        self.open_overview_check_btn.config(state="normal" if has_any else "disabled")
        if not has_any:
            msg = "Inga avvikelser hittades i orderöversikten."
            if missing_hib_cols:
                msg += "\nHIB-kontrollen kunde inte köras fullt ut (saknar kolumner: " + ", ".join(missing_hib_cols) + ")."
            self._track_feature("overview_check", "run_completed", shipment_rows=0, hib_rows=0)
            messagebox.showinfo(APP_TITLE, msg)
            return

        # Logga
        try:
            if not result_df.empty:
                self._log("Orderöversikt: sändningsnummer med flera kunder eller transportörer:")
                for _, row in result_df.iterrows():
                    try:
                        if int(row.get("Unika kunder", 0)) > 1:
                            self._log(f"  Sändningsnr {row['Sändningsnr']} har flera kunder: {row['Kunder']}")
                        if int(row.get("Unika transportörer", 0)) > 1:
                            self._log(f"  Sändningsnr {row['Sändningsnr']} har flera transportörer: {row['Transportörer']}")
                    except Exception:
                        pass
            if not hib_check_df.empty:
                self._log(f"HIB-ordrar med status > 31 utan matchande butikssändning ({len(hib_check_df)} st):")
                for _, row in hib_check_df.iterrows():
                    try:
                        name_part = f" ({row['Kundnamn']})" if str(row.get("Kundnamn", "")).strip() else ""
                        self._log(f"  Order {row['Ordernr']}{name_part}: sändning {row['Sändningsnr']} (status {row['Status']})")
                    except Exception:
                        pass
            if missing_hib_cols:
                self._log("HIB-kontrollen kunde inte köras fullt ut (saknar kolumner: " + ", ".join(missing_hib_cols) + ").")
            self._log("Orderkontrollen är beräknad och redo att öppnas i Excel.")
        except Exception:
            pass
        self._track_feature(
            "overview_check",
            "run_completed",
            shipment_rows=int(len(result_df)) if isinstance(result_df, pd.DataFrame) else 0,
            hib_rows=int(len(hib_check_df)) if isinstance(hib_check_df, pd.DataFrame) else 0,
        )

    def open_overview_check_in_excel(self) -> None:
        """
        Öppna resultatet av den senaste orderöversiktkontrollen i Excel.
        Innehåller ett blad för sändningsnummer med flera kunder/transportörer
        och ett blad för HIB-ordrar med status > 31 utan matchande butikssändning.
        """
        has_sändning = isinstance(self.last_overview_check_df, pd.DataFrame) and not self.last_overview_check_df.empty
        has_hib = isinstance(getattr(self, "last_hib_status_check_df", None), pd.DataFrame) and not self.last_hib_status_check_df.empty
        if has_sändning or has_hib:
            try:
                sheets: dict[str, pd.DataFrame] = {}
                combined_parts: List[pd.DataFrame] = []
                if has_sändning:
                    s_df = self.last_overview_check_df.copy()
                    if "Avvikelsetyp" not in s_df.columns:
                        s_df.insert(0, "Avvikelsetyp", "Sändningsnr med flera kunder/transportörer")
                    sheets["Sändningskontroll"] = s_df.copy()
                    combined_parts.append(s_df)
                if has_hib:
                    h_df = self.last_hib_status_check_df.copy()
                    if "Avvikelsetyp" not in h_df.columns:
                        h_df.insert(0, "Avvikelsetyp", "HIB över status 31 utan butikssändning")
                    sheets["HIB utan butikssändning"] = h_df.copy()
                    combined_parts.append(h_df)
                if combined_parts:
                    sheets = {
                        "Orderkontroll": pd.concat(combined_parts, ignore_index=True, sort=False),
                        **sheets,
                    }
                path = _open_df_in_excel(sheets, label="orderkontroll")
                self.last_overview_check_path = path
                self._log(f"Öppnade orderkontroll i Excel (temporär fil): {path}")
                self._track_feature("overview_check", "opened_result")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna orderkontroll i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget orderkontroll-resultat att öppna. Kör kontrollen först.")

    def run_dispatch_check(self) -> None:
        """
        Kontrollera att ordernummer och sändningsnummer i dispatchpallarna stämmer
        överens med orderöversikten.  Identifierar och loggar avvikelser.
        """
        overview_path = self.overview_var.get().strip()
        dispatch_path = getattr(self, "dispatch_var", tk.StringVar()).get().strip()
        self._track_feature("dispatch_check", "run_started")
        if not overview_path or not dispatch_path:
            messagebox.showerror(APP_TITLE, "Välj både orderöversikt och dispatchpallar först.")
            return
        self._log_active_value_filters()
        # Läs in orderöversikt
        try:
            ov_df = pd.read_csv(overview_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
            if ov_df.shape[1] == 1:
                ov_df = pd.read_csv(overview_path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa orderöversikten:\n{e}")
            return
        # Läs in dispatch
        try:
            dp_df = pd.read_csv(dispatch_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
            if dp_df.shape[1] == 1:
                dp_df = pd.read_csv(dispatch_path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa dispatchpallarna:\n{e}")
            return
        # Normalisera kolumnnamn
        ov_df.columns = [str(c).replace("\ufeff", "").strip() for c in ov_df.columns]
        dp_df.columns = [str(c).replace("\ufeff", "").strip() for c in dp_df.columns]
        try:
            ov_df = self._apply_value_filters(ov_df, "Orderöversikt (dispatchkontroll)")
            dp_df = self._apply_value_filters(dp_df, "Dispatchpallar (dispatchkontroll)")
        except Exception:
            pass
        if ov_df.empty or dp_df.empty:
            self.open_dispatch_check_btn.config(state="disabled")
            self.last_dispatch_check_df = pd.DataFrame()
            messagebox.showinfo(APP_TITLE, "Inga rader kvar efter filter för dispatchkontrollen.")
            return
        # Hjälpfunktion för att hitta kolumn baserat på nyckelord
        def _find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
            # exakt match
            for kw in keywords:
                kw_norm = kw.lower().replace(" ", "")
                for col in df.columns:
                    if col.lower().replace(" ", "") == kw_norm:
                        return col
            # delmatch
            for kw in keywords:
                kw_lower = kw.lower()
                for col in df.columns:
                    if kw_lower in col.lower():
                        return col
            return None
        order_keywords = ["ordernr", "order nr", "ordernummer", "order number", "orderid", "order id"]
        ship_keywords = ["sändningsnr", "sändnings nr", "sändningsnummer", "sandningsnr", "sandnings nr", "sandningsnummer", "shipment"]
        plock_keywords = ["plockpallsnr", "plockpallsnr.", "plockpall", "plockpallnr", "plockpallsnummer", "plockpall nr"]
        ov_order_col = _find_col(ov_df, order_keywords)
        ov_ship_col = _find_col(ov_df, ship_keywords)
        if not ov_order_col or not ov_ship_col:
            messagebox.showerror(APP_TITLE, "Kunde inte identifiera order- eller sändningskolumnen i orderöversikten.")
            return
        dp_order_col = _find_col(dp_df, order_keywords)
        dp_ship_col = _find_col(dp_df, ship_keywords)
        plock_col = _find_col(dp_df, plock_keywords)
        if not dp_order_col or not dp_ship_col or not plock_col:
            messagebox.showerror(APP_TITLE, "Kunde inte identifiera order-, sändnings- eller plockpallskolumnen i dispatchfilen.")
            return
        # Rensa strängar
        ov_df[ov_order_col] = ov_df[ov_order_col].astype(str).str.strip()
        ov_df[ov_ship_col] = ov_df[ov_ship_col].astype(str).str.strip()
        dp_df[dp_order_col] = dp_df[dp_order_col].astype(str).str.strip()
        dp_df[dp_ship_col] = dp_df[dp_ship_col].astype(str).str.strip()
        dp_df[plock_col] = dp_df[plock_col].astype(str).str.strip()

        # ----- Ny funktionalitet: hämta kundnamn per order -----
        # Försök bygga en mappning från ordernummer till kundnamn. Detta gör att
        # dispatchkontrollen kan visa vilken butik varje order tillhör, på samma sätt
        # som HIB‑kopplingen visar kundnamn. Prioritera beställningsfilen om den är
        # inläst, annars använd orderöversikten som fallback.
        order_to_customer: Dict[str, str] = {}
        try:
            # Försök läsa beställningsrader om en sådan fil är angiven
            details_path = getattr(self, "orders_var", tk.StringVar()).get().strip()
            if details_path:
                try:
                    det_df = pd.read_csv(details_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
                    if det_df.shape[1] == 1:
                        det_df = pd.read_csv(details_path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
                    det_df.columns = [str(c).replace("\ufeff", "").strip() for c in det_df.columns]
                    try:
                        det_df = self._apply_value_filters(det_df, "Beställningslinjer (dispatchkontroll)")
                    except Exception:
                        pass
                    # Standardkolumn för kundnamn i beställningsrader är "Kund.1" enligt HIB‑logiken
                    if "Order nr" in det_df.columns and "Kund.1" in det_df.columns:
                        try:
                            order_to_customer = (
                                det_df.groupby("Order nr")["Kund.1"].first()
                                .fillna("")
                                .astype(str)
                                .str.strip()
                                .to_dict()
                            )
                        except Exception:
                            order_to_customer = {}
                except Exception:
                    # Ignorera fel från läsning av beställningsfilen
                    order_to_customer = {}
        except Exception:
            order_to_customer = {}
        # Om vi fortfarande saknar kundnamn, försök hämta det från orderöversikten
        if not order_to_customer:
            try:
                # Försök hitta en kolumn i ov_df som innehåller kundnamn
                cust_candidate = None
                for col in ov_df.columns:
                    # Uteslut order‑ och sändningskolumnerna
                    if col in (ov_order_col, ov_ship_col):
                        continue
                    cn = col.lower().replace(" ", "")
                    # Ignorera kolumner som bara innehåller nummer (ex kund nr), leta efter namn
                    if "kund" in cn and not cn.endswith("nr"):
                        cust_candidate = col
                        break
                if cust_candidate:
                    try:
                        order_to_customer = (
                            ov_df.groupby(ov_order_col)[cust_candidate]
                            .first()
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            .to_dict()
                        )
                    except Exception:
                        order_to_customer = {}
            except Exception:
                order_to_customer = {}
        # Skapa mapping order → sändningsnummer från orderöversikten
        order_to_ship: Dict[str, str] = {}
        try:
            for ordnum, sub in ov_df.groupby(ov_order_col):
                ships = [s for s in sub[ov_ship_col] if isinstance(s, str) and s.strip()]
                if ships:
                    order_to_ship[str(ordnum)] = ships[0].strip()
        except Exception:
            pass
        # Jämför dispatch mot orderöversikten
        diff_rows: List[Dict[str, object]] = []
        for _, row in dp_df.iterrows():
            try:
                ordnr = str(row[dp_order_col]).strip()
                dp_ship = str(row[dp_ship_col]).strip()
                expected = order_to_ship.get(ordnr)
                # Om det finns ett förväntat sändningsnummer och det skiljer sig från dispatchens
                if expected and expected != dp_ship:
                    diff_row: Dict[str, object] = {
                        "Ordernr": ordnr,
                        "Översikt sändningsnr": expected,
                        "Dispatch sändningsnr": dp_ship,
                        "Plockpallsnr": str(row[plock_col]).strip(),
                    }
                    # Lägg till kundnamn om tillgängligt
                    try:
                        kundnamn_val = order_to_customer.get(ordnr, "")
                    except Exception:
                        kundnamn_val = ""
                    diff_row["kundnamn"] = kundnamn_val
                    diff_rows.append(diff_row)
            except Exception:
                continue
        if not diff_rows:
            self.open_dispatch_check_btn.config(state="disabled")
            self.last_dispatch_check_df = pd.DataFrame()
            self._track_feature("dispatch_check", "run_completed", mismatch_rows=0)
            messagebox.showinfo(APP_TITLE, "Alla sändningsnummer stämmer överens mellan orderöversikten och dispatchpallar.")
            return
        diff_df = pd.DataFrame(diff_rows)
        self.last_dispatch_check_df = diff_df.copy()
        # Aktivera öppna-knappen
        self.open_dispatch_check_btn.config(state="normal")
        # Logga
        try:
            self._log("Dispatchkontrollen har hittat avvikelser mellan orderöversikten och dispatchpallar:")
            for _, row in diff_df.iterrows():
                try:
                    # Ta med kundnamn i loggen om det finns
                    name_part = ""
                    try:
                        nm = str(row.get("kundnamn", "")).strip()
                        if nm:
                            name_part = f" ({nm})"
                    except Exception:
                        name_part = ""
                    self._log(
                        f"Order {row['Ordernr']}{name_part} har sändningsnr {row['Översikt sändningsnr']} i översikten men {row['Dispatch sändningsnr']} i dispatch (plockpall {row['Plockpallsnr']})"
                    )
                except Exception:
                    pass
            self._log("Dispatchkontrollen är beräknad och redo att öppnas i Excel.")
        except Exception:
            pass
        self._track_feature(
            "dispatch_check",
            "run_completed",
            mismatch_rows=int(len(diff_df)) if isinstance(diff_df, pd.DataFrame) else 0,
        )

    def open_dispatch_check_in_excel(self) -> None:
        """
        Öppna resultatet av den senaste dispatchkontrollen i Excel.
        """
        if isinstance(self.last_dispatch_check_df, pd.DataFrame) and not self.last_dispatch_check_df.empty:
            try:
                path = _open_df_in_excel({"Dispatchkontroll": self.last_dispatch_check_df.copy()}, label="dispatchkontroll")
                self.last_dispatch_check_path = path
                self._log(f"Öppnade dispatchkontroll i Excel (temporär fil): {path}")
                self._track_feature("dispatch_check", "opened_result")
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Kunde inte öppna dispatchkontroll i Excel:\n{e}")
        else:
            messagebox.showinfo(APP_TITLE, "Det finns inget dispatchkontroll-resultat att öppna. Kör kontrollen först.")


    def _on_sales_file_selected(self) -> None:
        """
        Stub för hantering av plocklogg. Funktionen för att läsa in plockloggar och beräkna försäljningsinsikter är borttagen i denna version.

        Denna metod finns kvar för kompatibilitet men gör inget längre.
        """
        return


    def update_summary_table(self, result_df: pd.DataFrame) -> None:
        """
        Uppdatera sammanställningstabellen med alla förekommande Källtyp‑värden.

        HELPALL visas som antal pallar, AUTOSTORE som antal rader, och övriga typer som
        antal rader samt motsvarande pallantal (20 rader per pall).
        """
        for child in self.summary_table.get_children(""):
            self.summary_table.delete(child)
        try:
            qty_col = find_col(result_df, ORDER_SCHEMA["qty"], required=False, default=None)
        except Exception:
            qty_col = None
        ktyp_series = result_df.get("Källtyp", pd.Series([], dtype=object)).astype(str)
        unique_types = [k for k in sorted(set(ktyp_series.dropna())) if k]
        ordered = []
        for prv in ("HELPALL", "AUTOSTORE"):
            if prv in unique_types:
                ordered.append(prv)
                unique_types.remove(prv)
        ordered.extend(unique_types)
        for ktyp in ordered:
            try:
                sub = result_df[ktyp_series == ktyp]
                row_count = int(len(sub))
                kolli = 0.0
                if qty_col and not sub.empty:
                    kolli = float(pd.to_numeric(sub[qty_col], errors="coerce").sum())
            except Exception:
                row_count, kolli = 0, 0.0
            if ktyp == "HELPALL":
                row_text = f"{row_count} pallar"
            elif ktyp == "AUTOSTORE":
                row_text = f"{row_count} rader"
            else:
                # För övriga Källtyper visas endast antal rader (ingen pallar‑beräkning i parentes)
                row_text = f"{row_count} rader"
            kolli_text = f"{int(round(kolli))}"
            self.summary_table.insert("", "end", iid=ktyp, values=(ktyp, row_text, kolli_text))


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
          * "F" → ("HIB",          "F")

        Rader vars Källtyp redan är "HELPALL" eller "AUTOSTORE" lämnas
        oförändrade. Om ingen "Zon"‑kolumn hittas returneras oförändrat DataFrame.
        Den medskickade saldofil används inte i denna metod.
        """
        if result_df is None or result_df.empty:
            return result_df
        res = result_df.copy()
        zon_col = None
        for c in res.columns:
            if str(c).strip().lower() == "zon":
                zon_col = c
                break
        if not zon_col:
            return res
        if "Zon (beräknad)" not in res.columns:
            res["Zon (beräknad)"] = ""
        ktyp_series = res.get("Källtyp", pd.Series("", index=res.index)).astype(str)
        mask_to_change = ~(ktyp_series.isin(["HELPALL", "AUTOSTORE"]))
        if not mask_to_change.any():
            return res
        mapping: Dict[str, Tuple[str, str]] = {
            "S": ("SKRYMMANDE",   "S"),
            "E": ("EHANDEL",      "E"),
            "A": ("HUVUDPLOCK",   "A"),
            "Q": ("EHANDEL",      "Q"),
            "O": ("SKRYMMANDE",   "O"),
            "F": ("HIB",          "F"),
            "D": ("DISPLAY",      "D"),
        }
        zones = res.loc[mask_to_change, zon_col].astype(str).str.strip().str.upper()
        for zone_code, (ktyp_val, zon_val) in mapping.items():
            idx = res.loc[mask_to_change].index[zones == zone_code]
            if len(idx) > 0:
                res.loc[idx, "Källtyp"] = ktyp_val
                res.loc[idx, "Zon (beräknad)"] = zon_val
        return res


    def run_allocation(self) -> None:
        orders_path = self.orders_var.get().strip()
        buffer_path = self.buffer_var.get().strip()
        automation_path = self.automation_var.get().strip()
        item_path = self.item_var.get().strip()
        not_putaway_path = ""
        self._track_feature(
            "allocation",
            "run_started",
            has_automation=bool(automation_path),
            has_item=bool(item_path),
            has_prognos=bool(str(self.prognos_var.get()).strip()),
            has_campaign=bool(str(self.campaign_var.get()).strip()),
        )

        if not orders_path or not buffer_path:
            messagebox.showerror(APP_TITLE, "Välj både beställningsfil och buffertfil.")
            return

        self._log_active_value_filters()

        try:
            self._log("Läser in filer...")
            orders_raw = pd.read_csv(orders_path, dtype=str, sep=None, engine="python")
            buffer_raw = pd.read_csv(buffer_path, dtype=str, sep=None, engine="python")

            self._not_putaway_raw = None
            self._not_putaway_norm = None

            if automation_path:
                auto_raw = pd.read_csv(automation_path, dtype=str, sep=None, engine="python")
                auto_raw_clean = _clean_columns(auto_raw.copy())
                self._saldo_raw = auto_raw_clean.copy()
                self._saldo_norm = normalize_saldo(auto_raw_clean)
            else:
                self._saldo_norm = None
                self._saldo_raw = None

            self._item_raw = None
            self._item_norm = None
            if item_path:
                try:
                    item_raw = pd.read_csv(item_path, dtype=str, sep=None, engine="python")
                except Exception:
                    try:
                        item_raw = pd.read_csv(item_path, dtype=str, sep="\t", quoting=3, engine="python")
                    except Exception as ie:
                        raise RuntimeError(f"Kunde inte läsa item-fil: {ie}")
                self._item_raw = item_raw.copy()
                self._item_norm = normalize_items(item_raw)

            orders_raw = _clean_columns(orders_raw)
            buffer_raw = _clean_columns(buffer_raw)

            orders_raw = self._apply_value_filters(orders_raw, "Beställningslinjer")
            buffer_raw = self._apply_value_filters(buffer_raw, "Buffertpallar")
            if isinstance(self._saldo_raw, pd.DataFrame):
                self._saldo_raw = self._apply_value_filters(self._saldo_raw, "Saldo inkl. automation")
            if isinstance(self._item_raw, pd.DataFrame):
                self._item_raw = self._apply_value_filters(self._item_raw, "Item option")
            if isinstance(self._saldo_raw, pd.DataFrame):
                self._saldo_norm = normalize_saldo(self._saldo_raw)
            if isinstance(self._item_raw, pd.DataFrame):
                self._item_norm = normalize_items(self._item_raw)
        except Exception as e:
            self._track_feature("allocation", "run_failed", stage="load_inputs")
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa CSV-filerna:\n{e}")
            return

        try:
            self._log("\n--------------")
            self._log(f"Kör allokering (Helpall → AutoStore → Huvudplock, FIFO) + {int(NEAR_MISS_PCT * 100)}%-near-miss loggning + Status {sorted(ALLOC_BUFFER_STATUSES)}-filter...")
            result, near = allocate(orders_raw, buffer_raw, log=self._log)

            result = self._reclassify_skrymmande(result, self._saldo_norm)

            try:
                if isinstance(self._item_norm, pd.DataFrame) and not self._item_norm.empty and isinstance(result, pd.DataFrame) and not result.empty:
                    try:
                        art_col_res = find_col(result, ORDER_SCHEMA["artikel"], required=True)
                    except Exception:
                        art_col_res = None
                    if art_col_res:
                        temp_merge = result.merge(self._item_norm, how="left", left_on=art_col_res, right_on="Artikel", suffixes=("", "_item"))
                        if "Artikel_item" in temp_merge.columns:
                            temp_merge.drop(columns=["Artikel_item"], inplace=True, errors=False)
                        if "Artikel_y" in temp_merge.columns:
                            temp_merge.drop(columns=["Artikel_y"], inplace=True, errors=False)
                        if "Ej Staplingsbar_y" in temp_merge.columns or "Ej Staplingsbar_x" in temp_merge.columns:
                            if "Ej Staplingsbar_y" in temp_merge.columns:
                                temp_merge["Ej Staplingsbar"] = temp_merge["Ej Staplingsbar_y"].fillna("")
                            elif "Ej Staplingsbar_x" in temp_merge.columns:
                                temp_merge["Ej Staplingsbar"] = temp_merge["Ej Staplingsbar_x"].fillna("")
                            for _col in ["Ej Staplingsbar_x", "Ej Staplingsbar_y"]:
                                if _col in temp_merge.columns:
                                    temp_merge.drop(columns=[_col], inplace=True)
                        if "Ej Staplingsbar" not in temp_merge.columns:
                            temp_merge["Ej Staplingsbar"] = ""
                        cols = [c for c in temp_merge.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
                        temp_merge = temp_merge[cols]
                        result = temp_merge
                if isinstance(result, pd.DataFrame) and ("Ej Staplingsbar" not in result.columns):
                    result["Ej Staplingsbar"] = ""
                    cols = [c for c in result.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
                    result = result[cols]
            except Exception as e:
                try:
                    self._log(f"Kunde inte slå ihop item-fil: {e}")
                except Exception:
                    pass
            self._log("Skapar resultat i minnet...")

            self.last_result_df = result.copy()
            self.last_nearmiss_instead_df = near.copy()
            self._orders_raw = orders_raw.copy()
            self._buffer_raw = buffer_raw.copy()
            self._result_df = result.copy()

            try:
                self._pallet_spaces_df = compute_pallet_spaces(self._result_df)
            except Exception:
                self._pallet_spaces_df = None

            try:
                self.update_summary_table(result)
            except Exception as _e_upd:
                self._log(f"Summering per Källtyp kunde inte uppdateras: {_e_upd}")

            try:
                hp_df, as_df = calculate_refill(
                    result, buffer_raw,
                    saldo_df=self._saldo_norm,
                    not_putaway_df=self._not_putaway_norm
                )
                self._last_refill_hp_df = hp_df.copy()
                self._last_refill_autostore_df = as_df.copy()
                self._log(f"Auto-refill klar: HP {len(hp_df)} rader, AUTOSTORE {len(as_df)} rader (cachad).")
            except Exception as e:
                self._last_refill_hp_df = None
                self._last_refill_autostore_df = None
                self._log(f"Auto-refill misslyckades: {e}")

            self.open_result_btn.configure(state="normal" if not result.empty else "disabled")
            try:
                self.open_nearmiss_btn.configure(state="normal" if isinstance(near, pd.DataFrame) and not near.empty else "disabled")
            except Exception:
                self.open_nearmiss_btn.configure(state="disabled")
            try:
                has_pallet = isinstance(self._pallet_spaces_df, pd.DataFrame) and not self._pallet_spaces_df.empty
                self.open_palletspaces_btn.configure(state="normal" if has_pallet else "disabled")
            except Exception:
                self.open_palletspaces_btn.configure(state="disabled")
            try:
                has_refill = isinstance(self._last_refill_hp_df, pd.DataFrame) or isinstance(self._last_refill_autostore_df, pd.DataFrame)
                self.open_refill_btn.configure(state="normal" if has_refill else "disabled")
            except Exception:
                self.open_refill_btn.configure(state="disabled")
        except Exception as e:
            self._track_feature("allocation", "run_failed", stage="allocation")
            messagebox.showerror(APP_TITLE, f"Fel under allokering:\n{e}")
            return

        try:
            zon_col = "Zon (beräknad)"
            qty_col = find_col(result, ORDER_SCHEMA["qty"], required=True)
            summary = result.groupby(zon_col)[qty_col].apply(lambda s: pd.to_numeric(s, errors="coerce").sum()).reset_index(name="Totalt antal")
            self._log("\nSummering per zon:")
            for _, r in summary.iterrows():
                self._log(f"  Zon {r[zon_col]}: {r['Totalt antal']:.0f}")
        except Exception:
            pass

        try:
            self._log(f"\n{int(NEAR_MISS_PCT * 100)}% near-miss statistik:")
            if isinstance(near, pd.DataFrame) and not near.empty:
                near_art_col = None
                for c in ["Artikel", "artikel", "Artikelnummer", "artikelnummer", "_artikel"]:
                    if c in near.columns:
                        near_art_col = c
                        break
                res_art_col = None
                try:
                    res_art_col = find_col(result, ORDER_SCHEMA["artikel"], required=False)
                except Exception:
                    for c in ["Artikel", "artikel", "Artikelnummer", "artikelnummer", "_artikel"]:
                        if c in result.columns:
                            res_art_col = c
                            break
                zone_col = "Zon (beräknad)"
                near_with_zone = near.copy()
                if near_art_col and res_art_col and zone_col in result.columns:
                    zone_map: Dict[str, str] = {}
                    res_art_series = result[res_art_col].astype(str).str.strip()
                    for art in near_with_zone[near_art_col].astype(str).str.strip().unique():
                        mask = res_art_series == art
                        if not mask.any():
                            continue
                        zones = result.loc[mask, zone_col].astype(str)
                        if not zones.empty:
                            zone_counts = zones.value_counts()
                            chosen_zone = zone_counts.idxmax()
                            zone_map[art] = chosen_zone
                    near_with_zone["Slutade som Zon"] = near_with_zone[near_art_col].astype(str).str.strip().map(lambda x: zone_map.get(x, ""))
                else:
                    near_with_zone["Slutade som Zon"] = ""
                zones_to_report = ["R", "A", "E", "S", "Q", "O", "F", "D"]
                for z in zones_to_report:
                    try:
                        cnt = 0
                        if near_art_col:
                            cnt = int(near_with_zone.loc[near_with_zone["Slutade som Zon"] == z, near_art_col].astype(str).str.strip().nunique())
                        self._log(f"  Near-miss som slutade som {z}: {cnt:,}")
                    except Exception:
                        self._log(f"  Near-miss som slutade som {z}: 0")
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
                self.last_nearmiss_instead_df = near_with_zone.copy()
            else:
                self._log("  Inga near-miss artiklar hittades.")
                self.last_nearmiss_instead_df = pd.DataFrame()
        except Exception:
            try:
                self._log("  Inga near-miss artiklar hittades.")
            except Exception:
                pass

        self._track_feature(
            "allocation",
            "run_completed",
            result_rows=int(len(result)) if isinstance(result, pd.DataFrame) else 0,
            near_miss_rows=int(len(self.last_nearmiss_instead_df)) if isinstance(self.last_nearmiss_instead_df, pd.DataFrame) else 0,
            pallet_space_rows=int(len(self._pallet_spaces_df)) if isinstance(self._pallet_spaces_df, pd.DataFrame) else 0,
            refill_hp_rows=int(len(self._last_refill_hp_df)) if isinstance(self._last_refill_hp_df, pd.DataFrame) else 0,
            refill_autostore_rows=int(len(self._last_refill_autostore_df)) if isinstance(self._last_refill_autostore_df, pd.DataFrame) else 0,
        )


def _read_cli_table(path: str) -> pd.DataFrame:
    """Las en tabellfil for CLI-kommandon och normalisera kolumnnamn."""
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(f"Filen finns inte: {target}")

    suffix = target.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}:
        return _clean_columns(pd.read_excel(target, dtype=str))

    try:
        df = pd.read_csv(target, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        if df.shape[1] == 1 and len(df):
            first = str(df.iloc[0, 0])
            if "\t" in first:
                df = pd.read_csv(target, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(target, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
    return _clean_columns(df)


def _write_cli_dataframe(df: pd.DataFrame, path: str) -> str:
    """Skriv DataFrame till CSV/XLSX/JSON beroende pa filandelse."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    suffix = target.suffix.lower()

    if suffix == ".xlsx":
        df.to_excel(target, index=False, engine="openpyxl")
    elif suffix == ".json":
        target.write_text(df.to_json(orient="records", force_ascii=False, indent=2), encoding="utf-8")
    else:
        df.to_csv(target, index=False, encoding="utf-8-sig")
    return str(target.resolve())


def _write_cli_workbook(sheets: Dict[str, pd.DataFrame], path: str) -> str:
    """Skriv flera blad till en Excel-fil."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.suffix.lower() != ".xlsx":
        raise ValueError("Flerbladsutskrift kraver .xlsx som utfil.")
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=str(sheet_name)[:31] or "Sheet1", index=False)
    return str(target.resolve())


def _write_cli_list(values: list[str], path: str, column_name: str) -> str:
    """Skriv en lista till TXT, CSV, XLSX eller JSON."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    suffix = target.suffix.lower()

    if suffix in {"", ".txt"}:
        target.write_text("\n".join(str(value) for value in values), encoding="utf-8")
        return str(target.resolve())
    if suffix == ".json":
        target.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
        return str(target.resolve())
    return _write_cli_dataframe(pd.DataFrame({column_name: values}), str(target))


def _emit_cli_summary(summary: dict, as_json: bool) -> None:
    if as_json:
        print(json.dumps(summary, ensure_ascii=False))
        return
    for key, value in summary.items():
        print(f"{key}: {value}")


def _load_utbest_map_from_saldo_path(path: Optional[str]) -> Dict[str, float]:
    if not path:
        return {}
    saldo_df = _read_cli_table(path)
    return utbest_per_article(saldo_df)


def _df_with_named_index(df: pd.DataFrame, index_name: str) -> pd.DataFrame:
    out = df.copy()
    out.index = out.index.map(lambda value: str(value).strip())
    return out.reset_index().rename(columns={"index": index_name})


def _merge_item_flags(result_df: pd.DataFrame, item_norm: pd.DataFrame | None) -> pd.DataFrame:
    if not isinstance(result_df, pd.DataFrame) or result_df.empty:
        return result_df

    result = result_df.copy()
    if not isinstance(item_norm, pd.DataFrame) or item_norm.empty:
        if "Ej Staplingsbar" not in result.columns:
            result["Ej Staplingsbar"] = ""
        cols = [c for c in result.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
        return result[cols]

    art_col_res = find_col(result, ORDER_SCHEMA["artikel"], required=True)
    merged = result.merge(item_norm, how="left", left_on=art_col_res, right_on="Artikel", suffixes=("", "_item"))
    if "Artikel_item" in merged.columns:
        merged.drop(columns=["Artikel_item"], inplace=True, errors=False)
    if "Artikel_y" in merged.columns:
        merged.drop(columns=["Artikel_y"], inplace=True, errors=False)
    if "Ej Staplingsbar_y" in merged.columns:
        merged["Ej Staplingsbar"] = merged["Ej Staplingsbar_y"].fillna("")
    elif "Ej Staplingsbar_x" in merged.columns:
        merged["Ej Staplingsbar"] = merged["Ej Staplingsbar_x"].fillna("")
    elif "Ej Staplingsbar" not in merged.columns:
        merged["Ej Staplingsbar"] = ""
    for col in ["Ej Staplingsbar_x", "Ej Staplingsbar_y"]:
        if col in merged.columns:
            merged.drop(columns=[col], inplace=True)
    cols = [c for c in merged.columns if c != "Ej Staplingsbar"] + ["Ej Staplingsbar"]
    return merged[cols]


def _resolve_max_csv_path(explicit_path: Optional[str]) -> Path:
    if explicit_path:
        target = Path(explicit_path)
        if not target.exists():
            raise FileNotFoundError(f"Filen finns inte: {target}")
        return target

    found = _find_lyx_max_csv()
    if found is None:
        raise FileNotFoundError("Kunde inte hitta lowfreqdata/buffertpall/artikel_max.csv.")
    return found


def _cli_allocate(args: argparse.Namespace) -> int:
    orders_raw = _read_cli_table(args.orders)
    buffer_raw = _read_cli_table(args.buffer)

    saldo_norm = None
    if args.saldo:
        saldo_norm = normalize_saldo(_read_cli_table(args.saldo))

    item_norm = None
    if args.items:
        item_norm = normalize_items(_read_cli_table(args.items))

    not_putaway_norm = None
    if args.not_putaway:
        not_putaway_norm = normalize_not_putaway(_read_cli_table(args.not_putaway))

    result_df, near_miss_df = allocate(orders_raw, buffer_raw)
    result_df = App._reclassify_skrymmande(result_df, saldo_norm)
    result_df = _merge_item_flags(result_df, item_norm)
    if near_miss_df.empty and len(near_miss_df.columns) == 0:
        near_miss_df = pd.DataFrame(
            columns=[
                "Artikel",
                "OrderID",
                "OrderRad",
                "PallID",
                "Källplats",
                "Mottagen",
                "Behov_vid_tillfället",
                "Pall_kvantitet",
                "Skillnad",
                "Procentuell skillnad (%)",
                "Anledning",
                "Gäller (INSTEAD R/A)",
            ]
        )

    output_paths: dict[str, str] = {}
    if args.result_out:
        output_paths["result"] = _write_cli_dataframe(result_df, args.result_out)
    if args.near_miss_out:
        output_paths["near_miss"] = _write_cli_dataframe(near_miss_df, args.near_miss_out)

    refill_hp_df = None
    refill_autostore_df = None
    if args.refill_hp_out or args.refill_autostore_out:
        refill_hp_df, refill_autostore_df = calculate_refill(
            result_df,
            buffer_raw,
            saldo_df=saldo_norm,
            not_putaway_df=not_putaway_norm,
        )
        if args.refill_hp_out:
            output_paths["refill_hp"] = _write_cli_dataframe(refill_hp_df, args.refill_hp_out)
        if args.refill_autostore_out:
            output_paths["refill_autostore"] = _write_cli_dataframe(refill_autostore_df, args.refill_autostore_out)

    pallet_spaces_df = None
    if args.pallet_spaces_out:
        pallet_spaces_df = compute_pallet_spaces(result_df)
        output_paths["pallet_spaces"] = _write_cli_dataframe(pallet_spaces_df, args.pallet_spaces_out)

    summary = {
        "command": "allocate",
        "result_rows": int(len(result_df)),
        "near_miss_rows": int(len(near_miss_df)),
        "refill_hp_rows": int(len(refill_hp_df)) if isinstance(refill_hp_df, pd.DataFrame) else 0,
        "refill_autostore_rows": int(len(refill_autostore_df)) if isinstance(refill_autostore_df, pd.DataFrame) else 0,
        "pallet_space_rows": int(len(pallet_spaces_df)) if isinstance(pallet_spaces_df, pd.DataFrame) else 0,
        "outputs": output_paths,
    }
    _emit_cli_summary(summary, args.json)
    return 0


def _cli_ordersaldo(args: argparse.Namespace) -> int:
    orders_df = _read_cli_table(args.orders)
    column_names = _find_ordersaldo_columns(orders_df)
    complete_orders, shortage_df = compute_ordersaldo_data(
        orders_df,
        utbest_map=_load_utbest_map_from_saldo_path(args.saldo),
        column_names=column_names,
    )

    output_paths: dict[str, str] = {}
    if args.complete_orders_out:
        output_paths["complete_orders"] = _write_cli_list(complete_orders, args.complete_orders_out, "Ordernr")
    if args.shortage_out:
        output_paths["shortage"] = _write_cli_dataframe(_df_with_named_index(shortage_df, "Artikel"), args.shortage_out)

    summary = {
        "command": "ordersaldo",
        "complete_order_count": int(len(complete_orders)),
        "shortage_article_count": int(len(shortage_df)),
        "complete_orders": complete_orders if args.json else None,
        "shortage_articles": sorted(shortage_df.index.astype(str).tolist()) if args.json else None,
        "outputs": output_paths,
    }
    _emit_cli_summary(summary, args.json)
    return 0


def _cli_lyx(args: argparse.Namespace) -> int:
    saldo_df = _read_cli_table(args.saldo)
    max_df = _read_cli_table(str(_resolve_max_csv_path(args.max_csv)))
    articles, filtered_row_count = compute_lyx_articles(saldo_df, max_df)

    output_paths: dict[str, str] = {}
    if args.output:
        output_paths["articles"] = _write_cli_list(articles, args.output, "Artikel")

    summary = {
        "command": "lyx",
        "filtered_row_count": int(filtered_row_count),
        "article_count": int(len(articles)),
        "articles": articles if args.json else None,
        "outputs": output_paths,
    }
    _emit_cli_summary(summary, args.json)
    return 0


def _cli_pafyllnadsprio(args: argparse.Namespace) -> int:
    orders_df = _read_cli_table(args.orders)
    column_names = _find_ordersaldo_columns(orders_df)
    _, shortage_df = compute_ordersaldo_data(
        orders_df,
        utbest_map=_load_utbest_map_from_saldo_path(args.saldo),
        column_names=column_names,
    )
    max_df = _read_cli_table(str(_resolve_max_csv_path(args.max_csv)))

    mode = "fallback"
    overview_error = None
    log_lines: list[str] = []
    missing_reference_count = 0
    window_map_df = None

    if args.overview:
        try:
            overview_df = _read_cli_table(args.overview)
            report_df, _bold_cells, log_lines, missing_reference_count, window_map_df = (
                build_pafyllnadsprio_lastningsfonster_report(
                    orders_df,
                    shortage_df,
                    overview_df,
                    max_df,
                    column_names=column_names,
                )
            )
            mode = "lastningsfonster"
        except Exception as exc:
            overview_error = str(exc)
            report_df, missing_reference_count = build_pafyllnadsprio_report(shortage_df, max_df)
    else:
        report_df, missing_reference_count = build_pafyllnadsprio_report(shortage_df, max_df)

    output_paths: dict[str, str] = {}
    if args.report_out:
        report_path = Path(args.report_out)
        if isinstance(window_map_df, pd.DataFrame) and report_path.suffix.lower() == ".xlsx":
            output_paths["report"] = _write_cli_workbook(
                {
                    "Påfyllnadsprio": report_df,
                    "Lastningsfönster": window_map_df,
                },
                args.report_out,
            )
        else:
            output_paths["report"] = _write_cli_dataframe(report_df, args.report_out)
    if args.window_map_out and isinstance(window_map_df, pd.DataFrame):
        output_paths["window_map"] = _write_cli_dataframe(window_map_df, args.window_map_out)

    summary = {
        "command": "pafyllnadsprio",
        "mode": mode,
        "shortage_article_count": int(len(shortage_df)),
        "report_rows": int(len(report_df)),
        "missing_reference_count": int(missing_reference_count),
        "overview_error": overview_error,
        "log_lines": log_lines if args.json else None,
        "outputs": output_paths,
    }
    _emit_cli_summary(summary, args.json)
    return 0


def _cli_hib_koppling(args: argparse.Namespace) -> int:
    details_df = _read_cli_table(args.details)
    overview_df = _read_cli_table(args.overview)
    changes_df = compute_hib_koppling(details_df, overview_df)
    missed_df = compute_missed_departures(details_df, overview_df)

    output_paths: dict[str, str] = {}
    if args.changes_out:
        output_paths["changes"] = _write_cli_dataframe(changes_df, args.changes_out)
    if args.missed_out:
        output_paths["missed"] = _write_cli_dataframe(missed_df, args.missed_out)

    summary = {
        "command": "hib-koppling",
        "change_rows": int(len(changes_df)),
        "missed_rows": int(len(missed_df)),
        "outputs": output_paths,
    }
    _emit_cli_summary(summary, args.json)
    return 0


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--version", action="store_true")

    subparsers = parser.add_subparsers(dest="command")

    allocate_parser = subparsers.add_parser("allocate", help="Kor allokering utan GUI.")
    allocate_parser.add_argument("--orders", required=True, help="Bestallningslinjer CSV/XLSX.")
    allocate_parser.add_argument("--buffer", required=True, help="Buffertpallar CSV/XLSX.")
    allocate_parser.add_argument("--saldo", help="Saldo/automation CSV/XLSX.")
    allocate_parser.add_argument("--items", help="Item option CSV/XLSX.")
    allocate_parser.add_argument("--not-putaway", help="Ej inlagrade CSV/XLSX.")
    allocate_parser.add_argument("--result-out", help="Utfil for allokerat resultat.")
    allocate_parser.add_argument("--near-miss-out", help="Utfil for near-miss.")
    allocate_parser.add_argument("--refill-hp-out", help="Utfil for refill huvudplock.")
    allocate_parser.add_argument("--refill-autostore-out", help="Utfil for refill autostore.")
    allocate_parser.add_argument("--pallet-spaces-out", help="Utfil for pallplatser.")
    allocate_parser.add_argument("--json", action="store_true", help="Skriv sammanfattning som JSON.")
    allocate_parser.set_defaults(cli_handler=_cli_allocate)

    ordersaldo_parser = subparsers.add_parser("ordersaldo", help="Berakna kompletta ordrar och underskott.")
    ordersaldo_parser.add_argument("--orders", required=True, help="Bestallningslinjer CSV/XLSX.")
    ordersaldo_parser.add_argument("--saldo", help="Saldo/automation CSV/XLSX for Utbestallt.")
    ordersaldo_parser.add_argument("--complete-orders-out", help="Utfil for kompletta ordernummer.")
    ordersaldo_parser.add_argument("--shortage-out", help="Utfil for artiklar med underskott.")
    ordersaldo_parser.add_argument("--json", action="store_true", help="Skriv sammanfattning som JSON.")
    ordersaldo_parser.set_defaults(cli_handler=_cli_ordersaldo)

    lyx_parser = subparsers.add_parser("lyx", help="Berakna LYX-artiklar utan GUI.")
    lyx_parser.add_argument("--saldo", required=True, help="Saldofil CSV/XLSX.")
    lyx_parser.add_argument("--max-csv", help="artikel_max.csv. Default ar lowfreqdata/buffertpall/artikel_max.csv.")
    lyx_parser.add_argument("--output", help="Utfil for artikelnummer.")
    lyx_parser.add_argument("--json", action="store_true", help="Skriv sammanfattning som JSON.")
    lyx_parser.set_defaults(cli_handler=_cli_lyx)

    pafyllnadsprio_parser = subparsers.add_parser("pafyllnadsprio", help="Kor pafyllnadsprio utan GUI.")
    pafyllnadsprio_parser.add_argument("--orders", required=True, help="Bestallningslinjer CSV/XLSX.")
    pafyllnadsprio_parser.add_argument("--saldo", help="Saldo/automation CSV/XLSX for Utbestallt.")
    pafyllnadsprio_parser.add_argument("--overview", help="Orderoversikt CSV/XLSX for lastningsfonster.")
    pafyllnadsprio_parser.add_argument("--max-csv", help="artikel_max.csv. Default ar lowfreqdata/buffertpall/artikel_max.csv.")
    pafyllnadsprio_parser.add_argument("--report-out", help="Utfil for rapporten.")
    pafyllnadsprio_parser.add_argument("--window-map-out", help="Utfil for lastningsfonster-tabell.")
    pafyllnadsprio_parser.add_argument("--json", action="store_true", help="Skriv sammanfattning som JSON.")
    pafyllnadsprio_parser.set_defaults(cli_handler=_cli_pafyllnadsprio)

    hib_parser = subparsers.add_parser("hib-koppling", help="Kor HIB-koppling utan GUI.")
    hib_parser.add_argument("--details", required=True, help="Bestallningslinjer CSV/XLSX.")
    hib_parser.add_argument("--overview", required=True, help="Orderoversikt CSV/XLSX.")
    hib_parser.add_argument("--changes-out", help="Utfil for andringar.")
    hib_parser.add_argument("--missed-out", help="Utfil for missade avgangar.")
    hib_parser.add_argument("--json", action="store_true", help="Skriv sammanfattning som JSON.")
    hib_parser.set_defaults(cli_handler=_cli_hib_koppling)

    return parser


def _parse_args(argv: Optional[list[str]] = None):
    parser = _build_arg_parser()
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = _parse_args(argv)
    if args.version:
        print(APP_VERSION)
        return 0

    cli_handler = getattr(args, "cli_handler", None)
    if cli_handler:
        return int(cli_handler(args) or 0)

    root_class = TkinterDnD.Tk if TkinterDnD else tk.Tk
    root = root_class()
    root.title(APP_TITLE)
    root.geometry("1360x860")
    if args.smoke_test:
        root.withdraw()
    _app = App(root, enable_update_checks=not args.smoke_test)
    if args.smoke_test:
        root.update_idletasks()
        root.update()
        root.destroy()
        return 0
    root.mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
