#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
logic.py – Rena logikfunktioner porterade från allokera11.py.
Inga tkinter-beroenden. Loggning sker via log_callback-parameter.
"""

from __future__ import annotations

import re
import os
import sys
import math
import subprocess
import importlib
import importlib.util
import tempfile
import unicodedata
from collections import defaultdict, deque
from typing import Deque, Dict, List, Optional, Tuple, Any

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Konstanter
# ---------------------------------------------------------------------------

APP_TITLE = "Buffertpallar → Order-allokering (GUI) — 10.9"
DEFAULT_OUTPUT = "allocated_orders.csv"

INVALID_LOC_PREFIXES: Tuple[str, ...] = ("AA",)
INVALID_LOC_EXACT: set = {"TRANSIT", "TRANSIT_ERROR", "MISSING", "UT2"}

ALLOC_BUFFER_STATUSES: set = {29, 30, 32}
REFILL_BUFFER_STATUSES: set = {29, 30}

NEAR_MISS_PCT: float = 0.30  # 30 % över behov

RF_PALLPLATS_EXCLUDE_ARTICLES: set = {
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
    "artikel": ORDER_SCHEMA["artikel"],
    "staplingsbar": [
        "staplingsbar", "staplings bar", "staplbar", "stackable",
        "ej staplingsbar", "ejstaplingsbar", "ej_staplingsbar", "non stackable"
    ]
}

PICK_LOG_SCHEMA: Dict[str, List[str]] = {
    "artikel": ["artikel", "artikelnr", "artnr", "art.nr", "artikelnummer", "sku", "article"],
    "antal":   ["plockat", "antal", "quantity", "qty", "picked", "units"],
    "datum":   ["datum", "datumtid", "timestamp", "date", "tid", "time"]
}

# ---------------------------------------------------------------------------
# Hjälpfunktioner
# ---------------------------------------------------------------------------

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ta bort BOM/whitespace i kolumnnamn."""
    try:
        df.rename(columns=lambda c: str(c).replace("\ufeff", "").strip(), inplace=True)
    except Exception:
        pass
    return df


def smart_to_datetime(s) -> pd.Series:
    """Robust datumtolkning."""
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
        try:
            return pd.to_datetime(s, errors="coerce", dayfirst=True)
        except Exception:
            return pd.to_datetime(s, errors="coerce", dayfirst=False)


def to_num(x) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else 0.0


def find_col(df: pd.DataFrame, candidates: List[str], required: bool = True, default=None) -> str:
    """Hitta en kolumn via exakt eller substring-match (case-insensitive)."""
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    for key, orig in cols.items():
        for cand in candidates:
            if cand.lower() in key:
                return orig
    if required and default is None:
        raise KeyError(f"Hittar inte kolumnerna {candidates} i {list(df.columns)}")
    return default


def _safe_str_series(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("").str.strip()


def _str_to_num(x) -> float:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    s = str(x).replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else 0.0


def _num_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.map(_str_to_num), errors="coerce").fillna(0)


def _sum_not_putaway(not_putaway_df: Optional[pd.DataFrame]) -> pd.Series:
    if not isinstance(not_putaway_df, pd.DataFrame) or not len(not_putaway_df):
        return pd.Series(dtype=float)
    df = not_putaway_df.copy()
    if "Artikel" not in df.columns or "Antal" not in df.columns:
        return pd.Series(dtype=float)
    df["Artikel"] = _safe_str_series(df["Artikel"])
    df["Antal"] = _num_series(df["Antal"])
    return df.groupby("Artikel")["Antal"].sum()


def _collect_exclude_source_ids(allocated_df: Optional[pd.DataFrame]) -> set:
    exclude: set = set()
    if isinstance(allocated_df, pd.DataFrame) and not allocated_df.empty:
        if "Källtyp" in allocated_df.columns and "Källa" in allocated_df.columns:
            mask = _safe_str_series(allocated_df["Källtyp"]) == "HELPALL"
            vals = _safe_str_series(allocated_df.loc[mask, "Källa"]).replace("", pd.NA).dropna().unique().tolist()
            exclude = set(vals)
    return exclude


def _open_df_in_excel(df, label: str = "data") -> str:
    """Skriv DF (eller {blad: DF}) till temporär fil och öppna."""
    if isinstance(df, dict):
        engine = None
        if importlib.util.find_spec("openpyxl"):
            engine = "openpyxl"
        elif importlib.util.find_spec("xlsxwriter"):
            engine = "xlsxwriter"
        else:
            raise RuntimeError("Saknar Excel-skrivare (installera 'openpyxl' eller 'xlsxwriter').")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        path = tmp.name
        tmp.close()
        with pd.ExcelWriter(path, engine=engine) as writer:
            for sheet, d in df.items():
                dd = d if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
                dd.to_excel(writer, sheet_name=str(sheet)[:31] or "Sheet1", index=False)
    else:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.csv")
        path = tmp.name
        tmp.close()
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


def save_df_to_excel(df, label: str = "data", out_path: str = None) -> str:
    """Spara DF (eller {blad: DF}) till xlsx utan att öppna. Returnerar sökvägen."""
    engine = None
    if importlib.util.find_spec("openpyxl"):
        engine = "openpyxl"
    elif importlib.util.find_spec("xlsxwriter"):
        engine = "xlsxwriter"
    else:
        raise RuntimeError("Saknar Excel-skrivare.")

    if out_path is None:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.xlsx")
        out_path = tmp.name
        tmp.close()

    if isinstance(df, dict):
        with pd.ExcelWriter(out_path, engine=engine) as writer:
            for sheet, d in df.items():
                dd = d if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
                sheet_name = str(sheet)[:31] or "Sheet1"
                dd.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(out_path, engine=engine) as writer:
            d = df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)
            d.to_excel(writer, sheet_name="Data", index=False)
    return out_path


# ---------------------------------------------------------------------------
# Filläsning och normalisering
# ---------------------------------------------------------------------------

def read_prognos_xlsx(path: str) -> pd.DataFrame:
    """Läs och normalisera en prognos-XLSX."""
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

    def _pick_col(cols: List[str], candidates: List[str]):
        s_cols = {_ci_match(c): c for c in cols}
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
    """Läs och normalisera en kampanjvolymfil (XLSX)."""
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
    for col_idx in [5, 4, 3, 1, 0]:
        if col_idx in df.columns:
            df = df.drop(columns=[col_idx])
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


def _read_not_putaway_csv(path: str) -> pd.DataFrame:
    """Läs CSV för 'Ej inlagrade'."""
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
    """Normalisera 'Ej inlagrade'-fil."""
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
    for c in ["Namn", "Pall nr", "SSCC"]:
        if c in out.columns:
            out[c] = out[c].fillna("").astype(str).str.strip()
    return out


def normalize_saldo(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normalisera saldofil."""
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
        "Plockplats": (df[plats_col].astype(str).str.strip() if plats_col else pd.Series([""] * len(df))),
    })
    agg = (out.groupby("Artikel", as_index=False)
              .agg({"Plocksaldo": "sum", "Plockplats": lambda s: next((x for x in s if isinstance(x, str) and x.strip()), "")}))
    return agg


def normalize_pick_log(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normalisera plocklogg."""
    df = _clean_columns(df_raw.copy())
    art_col = find_col(df, PICK_LOG_SCHEMA["artikel"], required=True)
    qty_col = find_col(df, PICK_LOG_SCHEMA["antal"], required=True)
    dt_col  = find_col(df, PICK_LOG_SCHEMA["datum"], required=True)
    name_col = None
    for cand in ["artikelnamn", "namn", "benämning", "artikelbenämning", "produktnamn"]:
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


def normalize_items(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normalisera item-fil."""
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


# ---------------------------------------------------------------------------
# Statistik
# ---------------------------------------------------------------------------

def compute_sales_metrics(df_norm: pd.DataFrame, today=None) -> pd.DataFrame:
    """Beräkna sales-mått per Artikelnummer."""
    if df_norm is None or df_norm.empty:
        cols = [
            "Artikelnummer", "Artikel", "Total_7", "Total_30", "Total_90", "ADV_30", "ADV_90",
            "SenastPlockad", "DagarSedanSenast", "UnikaPlockdagar_90", "NollraderPerPlockdag_90", "ABC_klass"
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
    zero_rows = (sub90.assign(IsZero=(sub90["Plockat"] == 0))
                      .groupby(["Artikelnummer", "DatumNorm"])["IsZero"].sum()
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
        out = out.merge(df_norm[["Artikelnummer", "Artikel"]].drop_duplicates(),
                        on="Artikelnummer", how="left")
    else:
        out["Artikel"] = out["Artikelnummer"]
    cols = ["Artikelnummer", "Artikel"] + [c for c in out.columns if c not in ["Artikelnummer", "Artikel"]]
    out = out[cols]
    return out


def annotate_refill(refill_df: pd.DataFrame, df_metrics: pd.DataFrame) -> pd.DataFrame:
    """Lägg på sales-kolumner i refill-blad."""
    if refill_df is None or refill_df.empty or df_metrics is None or df_metrics.empty:
        return refill_df
    cols = ["Artikel", "ADV_90", "ABC_klass", "DagarSedanSenast", "UnikaPlockdagar_90", "NollraderPerPlockdag_90"]
    cols = [c for c in cols if c in df_metrics.columns or c == "Artikel"]
    out = refill_df.merge(df_metrics[cols], on="Artikel", how="left")
    return out


# ---------------------------------------------------------------------------
# Pallplatser
# ---------------------------------------------------------------------------

def compute_pallet_spaces(result_df: pd.DataFrame) -> pd.DataFrame:
    """Beräkna pallplatsbehov per kund."""
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
    records: list = []
    for keys, sub in groups:
        if kund1_col is None:
            kund_val = keys
            kund1_val = ""
        else:
            kund_val, kund1_val = keys
        mask_bottom = (sub[zone_col] == "H") & ((sub[stack_col] == "N") | (sub[stack_col] == ""))
        B = int(mask_bottom.sum())
        rows_A = int((sub[zone_col] == "A").sum())
        top_A = math.ceil(rows_A / 20.0) if rows_A > 0 else 0
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
        record = {
            "Kund": kund_val,
            "Kund1": kund1_val,
            "hellpall stapelbar": B,
            "hellpall ej stapelbar": top_H,
            "Sjö pall": S_rows,
            "Skrymme": top_S,
            "Plockpall": top_A,
            "autostore": top_R,
            "Botten Pallar": B,
            "Topp Pallar": T,
            "Totalt Pallar": total_pallar,
            "Pallplatser": P
        }
        records.append(record)
    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# FIFO-hjälp
# ---------------------------------------------------------------------------

def _fifo_pallar_for_article(buffer_df: Optional[pd.DataFrame], article: str, needed_units: float, exclude_source_ids: Optional[set] = None) -> float:
    """FIFO-baserad beräkning för antal pallar."""
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


# ---------------------------------------------------------------------------
# Prognos vs autoplock
# ---------------------------------------------------------------------------

def build_prognos_vs_autoplock_report(
    prognos_df: pd.DataFrame,
    saldo_norm_df: Optional[pd.DataFrame] = None,
    buffer_df: Optional[pd.DataFrame] = None,
    *,
    exclude_source_ids: Optional[set] = None,
    allocated_df: Optional[pd.DataFrame] = None,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Bygg rapport som jämför prognosbehov med saldo."""
    meta: Dict[str, str] = {"partial": "no", "missing": "", "note": ""}
    missing: List[str] = []
    if not isinstance(prognos_df, pd.DataFrame) or prognos_df.empty:
        empty = pd.DataFrame(columns=[
            "Artikelnummer", "Behov i prognosen (antal styck)", "Saldo i autoplock",
            "Behov efter saldo", "Summa antal i ej inlagrade artiklar", "FIFO-baserad beräkning (antal pall)",
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
        "Artikelnummer", "Behov i prognosen (antal styck)", "Saldo i autoplock",
        "Behov efter saldo", "Buffertsaldo (status 29,30)", "FIFO-baserad beräkning (antal pall)",
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
    """Skriv prognosrapport till Excel och returnera sökväg."""
    sheets: dict = {}
    if isinstance(meta, dict) and (meta.get("partial") == "yes" or meta.get("note")):
        lines: list = []
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


# ---------------------------------------------------------------------------
# HIB-koppling
# ---------------------------------------------------------------------------

def compute_hib_koppling(details_df: pd.DataFrame, overview_df: pd.DataFrame) -> pd.DataFrame:
    """Analysera ordrar och identifiera HIB-ordrar som behöver ändras."""
    details = details_df.copy()
    overview = overview_df.copy()
    details.columns = [str(c).replace("\ufeff", "").strip() for c in details.columns]
    overview.columns = [str(c).replace("\ufeff", "").strip() for c in overview.columns]
    synonyms = {
        "Ordernr": ["Ordernr", "Order nr", "Order number", "Ordernummer"],
        "Ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
        "Kund nr": ["Kund nr", "Kundnr", "Kundnummer", "Customer number", "Kund NR"],
        "Bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
        "Orderdatum": ["Orderdatum", "Order datum", "Order date", "Orderdate"],
        "Sändningsnr": ["Sändningsnr", "Sändnings nr", "Sändningsnummer", "Sendingsnr", "Sändnings number"],
        "Zon": ["Zon", "Zone"],
        "Multi": ["Multi", "Multi nr", "Multinr", "Multi number"],
    }
    for canonical, syns in synonyms.items():
        if canonical in overview.columns:
            continue
        for candidate in syns:
            for col in list(overview.columns):
                if col.strip().lower() == candidate.strip().lower():
                    overview.rename(columns={col: canonical}, inplace=True)
                    break
            if canonical in overview.columns:
                break
    required_overview_cols = {"Ordernr", "Ordertyp", "Kund nr", "Orderdatum", "Sändningsnr", "Zon", "Multi"}
    missing = [c for c in required_overview_cols if c not in overview.columns]
    if missing:
        return pd.DataFrame(columns=["ordernummer", "Orderdatum", "sändningsnummer", "Zon", "Multi"])
    ov = overview.copy()
    ov["Ordertyp"] = ov["Ordertyp"].astype(str).str.strip().str.upper()
    if ov.empty:
        return pd.DataFrame(columns=["ordernummer", "Orderdatum", "sändningsnummer", "Zon", "Multi"])
    details.columns = [c.replace("\ufeff", "").strip() for c in details.columns]
    if "Order nr" not in details.columns or "Status" not in details.columns:
        return pd.DataFrame(columns=["ordernummer", "Orderdatum", "sändningsnummer", "Zon", "Multi"])

    def to_status_numeric(x):
        try:
            return int(float(str(x).strip()))
        except Exception:
            return 9999

    details["_status_num"] = details["Status"].apply(to_status_numeric)
    order_status_max = details.groupby("Order nr")["_status_num"].max().to_dict()
    order_zones = details.groupby("Order nr")["Zon"].apply(lambda x: list(x.dropna().astype(str))).to_dict()
    order_to_kundnamn: dict = {}
    if "Order nr" in details.columns and "Kund.1" in details.columns:
        try:
            order_to_kundnamn = (details.groupby("Order nr")["Kund.1"].first()
                                 .fillna("").astype(str).str.strip().to_dict())
        except Exception:
            order_to_kundnamn = {}
    rows: list = []
    for kund_nr, kund_df in ov.groupby("Kund nr"):
        store_df = kund_df[kund_df["Ordertyp"] == "N"].copy()
        hib_df = kund_df[kund_df["Ordertyp"] == "HIB"].copy()
        if not store_df.empty:
            store_df = store_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
        if not hib_df.empty:
            hib_df = hib_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
        if store_df.empty or hib_df.empty:
            continue
        valid_store_df = store_df[store_df["Ordernr"].map(lambda ordnum: order_status_max.get(ordnum, 0) < 34)].copy()
        if valid_store_df.empty:
            continue

        def _choose_earliest(df: pd.DataFrame) -> pd.Series:
            earliest_row = df.iloc[0]
            earliest_date = str(earliest_row["Orderdatum"]).strip()
            for _, r in df.iterrows():
                date_str = str(r["Orderdatum"]).strip()
                try:
                    d_new = pd.to_datetime(date_str, errors="coerce")
                    d_old = pd.to_datetime(earliest_date, errors="coerce")
                    if (pd.isna(d_old) and not pd.isna(d_new)) or (not pd.isna(d_old) and not pd.isna(d_new) and d_new < d_old):
                        earliest_row = r
                        earliest_date = date_str
                    elif pd.isna(d_new) and pd.isna(d_old) and date_str < earliest_date:
                        earliest_row = r
                        earliest_date = date_str
                except Exception:
                    if date_str < earliest_date:
                        earliest_row = r
                        earliest_date = date_str
            return earliest_row

        fallback_store_row = _choose_earliest(valid_store_df)
        hib_orders: list = []
        for _, hib_row in hib_df.iterrows():
            h_ord = hib_row["Ordernr"]
            maxstatus = order_status_max.get(h_ord, 9999)
            if maxstatus >= 34:
                continue
            hib_orders.append({"row": hib_row, "ordernr": h_ord})
        if not hib_orders:
            continue
        zone_flag = False
        hib_zone_updates = {}
        for hib in hib_orders:
            h_ord = hib["ordernr"]
            zones = [z.strip().upper() for z in order_zones.get(h_ord, []) if str(z).strip()]
            if any(z not in ("F", "H", "R") for z in zones):
                zone_flag = True
                hib_zone_updates[h_ord] = "F"
            else:
                hib_zone_updates[h_ord] = ""
        hib_f_multi: dict = {}
        missing_multi_per_order: dict = {}
        for hib in hib_orders:
            h_ord = hib["ordernr"]
            hib_zone_rows = kund_df[(kund_df["Ordernr"] == h_ord) & (kund_df["Zon"].astype(str).str.strip().str.upper() == "F")]
            mlist: list = []
            if hib_zone_rows.empty:
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
        multi_vals_global: set = set()
        for mlist in hib_f_multi.values():
            for m in mlist:
                if m:
                    multi_vals_global.add(m)
        common_multi_exists = len(multi_vals_global) == 1
        common_multi_value = next(iter(multi_vals_global)) if common_multi_exists else None
        for hib in hib_orders:
            h_row = hib["row"]
            h_ord = hib["ordernr"]
            ship_update = ""
            date_update = ""
            z_update = hib_zone_updates.get(h_ord, "")
            cur_ship = str(h_row["Sändningsnr"]).strip()
            cur_date = str(h_row["Orderdatum"]).strip()
            hib_kundnamn = order_to_kundnamn.get(h_ord, "").strip().lower()

            def _store_kname(ordnr: str) -> str:
                return order_to_kundnamn.get(ordnr, "").strip().lower()

            ship_kname_candidates = store_df[
                (store_df["Sändningsnr"].astype(str).str.strip() == cur_ship)
                & (store_df["Ordernr"].map(lambda x: _store_kname(x) == hib_kundnamn))
            ]
            if not ship_kname_candidates.empty:
                candidate_row = _choose_earliest(ship_kname_candidates)
            else:
                ship_candidates = store_df[store_df["Sändningsnr"].astype(str).str.strip() == cur_ship]
                if not ship_candidates.empty:
                    candidate_row = _choose_earliest(ship_candidates)
                else:
                    candidate_row = fallback_store_row
            ref_ship = str(candidate_row["Sändningsnr"]).strip()
            ref_date = str(candidate_row["Orderdatum"]).strip()
            if cur_ship != ref_ship:
                ship_update = ref_ship
            if cur_date != ref_date:
                date_update = ref_date
            multi_update = ""
            if len(hib_orders) > 1:
                if missing_multi_per_order.get(h_ord, False):
                    multi_update = "MULTI"
                else:
                    if common_multi_exists:
                        if set(hib_f_multi.get(h_ord, [])) != {common_multi_value}:
                            multi_update = "MULTI"
                    else:
                        multi_update = "MULTI"
            if ship_update or date_update or z_update or multi_update:
                rows.append({
                    "ordernummer": h_ord,
                    "kundnamn": order_to_kundnamn.get(h_ord, ""),
                    "Orderdatum": date_update,
                    "sändningsnummer": ship_update,
                    "Zon": z_update,
                    "Multi": multi_update
                })
    if not rows:
        return pd.DataFrame(columns=["ordernummer", "Orderdatum", "sändningsnummer", "Zon", "Multi"])
    result_df = pd.DataFrame(rows)
    if result_df.empty:
        return result_df
    result_df = result_df.sort_values(by=["kundnamn", "ordernummer"]).reset_index(drop=True)
    cols = ["ordernummer", "kundnamn", "Orderdatum", "sändningsnummer", "Zon", "Multi"]
    result_df = result_df[cols]
    return result_df


def compute_missed_departures(details_df: pd.DataFrame, overview_df: pd.DataFrame) -> pd.DataFrame:
    """Identifiera HIB-ordrar som missat sin avgång."""
    try:
        details = details_df.copy()
        overview = overview_df.copy()
        details.columns = [str(c).replace("\ufeff", "").strip() for c in details.columns]
        overview.columns = [str(c).replace("\ufeff", "").strip() for c in overview.columns]
        synonyms = {
            "Ordernr": ["Ordernr", "Order nr", "Order number", "Ordernummer"],
            "Ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
            "Kund nr": ["Kund nr", "Kundnr", "Kundnummer", "Customer number", "Kund NR"],
            "Bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
            "Orderdatum": ["Orderdatum", "Order datum", "Order date", "Orderdate"],
            "Sändningsnr": ["Sändningsnr", "Sändnings nr", "Sändningsnummer", "Sendingsnr", "Sändnings number"],
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
        required_overview_cols = {"Ordernr", "Ordertyp", "Kund nr", "Sändningsnr"}
        if any(c not in overview.columns for c in required_overview_cols):
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        ov = overview.copy()
        ov["Ordertyp"] = ov["Ordertyp"].astype(str).str.strip().str.upper()
        if ov.empty:
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])
        if "Order nr" not in details.columns or "Status" not in details.columns:
            return pd.DataFrame(columns=["ordernummer", "kundnamn", "Missat"])

        def to_status_numeric(x):
            try:
                return int(float(str(x).strip()))
            except Exception:
                return 9999

        details["_status_num"] = details["Status"].apply(to_status_numeric)
        order_status_max = details.groupby("Order nr")["_status_num"].max().to_dict()
        order_to_kundnamn: dict = {}
        if "Order nr" in details.columns and "Kund.1" in details.columns:
            try:
                order_to_kundnamn = (
                    details.groupby("Order nr")["Kund.1"].first()
                    .fillna("").astype(str).str.strip().to_dict()
                )
            except Exception:
                order_to_kundnamn = {}
        rows: list = []
        for kund_nr, kund_df in ov.groupby("Kund nr"):
            store_df = kund_df[kund_df["Ordertyp"] == "N"].copy()
            hib_df = kund_df[kund_df["Ordertyp"] == "HIB"].copy()
            if not store_df.empty:
                store_df = store_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
            if not hib_df.empty:
                hib_df = hib_df.drop_duplicates(subset=["Ordernr"]).reset_index(drop=True)
            if store_df.empty or hib_df.empty:
                continue
            store_ships: set = set()
            for _, row in store_df.iterrows():
                ship = str(row.get("Sändningsnr", "")).strip()
                if ship:
                    store_ships.add(ship)
            for _, hib_row in hib_df.iterrows():
                h_ord = hib_row["Ordernr"]
                maxstatus = order_status_max.get(h_ord, 9999)
                if maxstatus <= 34:
                    continue
                cur_ship = str(hib_row.get("Sändningsnr", "")).strip()
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


# ---------------------------------------------------------------------------
# Allokering
# ---------------------------------------------------------------------------

def _reclassify_skrymmande(result_df: pd.DataFrame, saldo_norm: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Omklassificera rader utifrån orderfilens zonkod."""
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


def allocate(orders_raw: pd.DataFrame, buffer_raw: pd.DataFrame, log=None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Allokera beställningsrader mot buffert (HELPALL→AUTOSTORE→HUVUDPLOCK).
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
        buff_type_col = find_col(buffer_raw, ["palltyp", "pall typ", "pall type"], required=False, default=None)
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
        _log(f"Filtrerar bort {excluded_count} rad(er) från bufferten pga lagerplats-regler.")
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
    near_miss_article_set: set = set()

    def clone_row(orow: pd.Series) -> dict:
        return orow.to_dict()

    def record_near_miss(orow: pd.Series, pal: dict, need: float) -> None:
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


def calculate_refill(
    allocated_df: pd.DataFrame,
    buffer_raw: pd.DataFrame,
    saldo_df: Optional[pd.DataFrame] = None,
    not_putaway_df: Optional[pd.DataFrame] = None
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Beräkna påfyllningspallar."""
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

    used_help_ids: set = set()
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
            grp = npu.groupby(npu[npu_art_col].astype(str).str.strip())[npu_qty_col].apply(
                lambda s: float(pd.to_numeric(s, errors="coerce").fillna(0).sum()))
            npu_sum = {str(k): float(v) for k, v in grp.to_dict().items()}
        except Exception:
            npu_sum = {}

    def fifo_for_art(art_key: str) -> pd.DataFrame:
        d = b[b["_artikel"] == art_key].copy()
        if not d.empty and used_help_ids:
            d = d[~d["_source_id"].astype(str).isin(used_help_ids)].copy()
        return d.sort_values("_received")

    hp_like = result[result.get("Källtyp", pd.Series("", index=result.index)).isin(["HUVUDPLOCK", "SKRYMMANDE", "HIB", "EHANDEL"])].copy()
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
                continue
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
                    continue
                behov_kvar = float(behov_int)
                pall_count = 0
                for q in (fifo_df["_qty"].astype(float) if not fifo_df.empty else []):
                    if behov_kvar <= 0:
                        break
                    pall_count += 1
                    behov_kvar -= float(q)
                rows_hp.append({
                    "Artikel": art_key,
                    "Zon": zone,
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
                    continue
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


# ---------------------------------------------------------------------------
# Filter-hjälp
# ---------------------------------------------------------------------------

def apply_value_filters(df: pd.DataFrame, active_filters: Dict[str, List[str]]) -> pd.DataFrame:
    """
    Filtrera ett DataFrame baserat på aktiva filter.
    active_filters: {"bolag": ["B01","B02"], "ordertyp": ["N","HIB"]}
    """
    if not active_filters or df.empty:
        return df
    filter_column_candidates = {
        "bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
        "ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
    }
    result = df.copy()
    for filter_key, selected_values in active_filters.items():
        if not selected_values:
            continue
        candidates = filter_column_candidates.get(filter_key, [filter_key])
        col = find_col(result, candidates, required=False, default=None)
        if col is None:
            continue
        result = result[result[col].astype(str).str.strip().isin([str(v) for v in selected_values])].copy()
    return result


def scan_filter_values(df: pd.DataFrame) -> Dict[str, List[str]]:
    """Skanna ett DataFrame och returnera unika värden för bolag/ordertyp."""
    filter_column_candidates = {
        "bolag": ["Bolag", "Company", "Bolag nr", "Bol"],
        "ordertyp": ["Ordertyp", "Order typ", "Order type", "Ordertype"],
    }
    result: Dict[str, List[str]] = {}
    for filter_key, candidates in filter_column_candidates.items():
        col = find_col(df, candidates, required=False, default=None)
        if col:
            vals = sorted(df[col].astype(str).str.strip().dropna().unique().tolist())
            if vals:
                result[filter_key] = vals
    return result


def read_csv_auto(path: str) -> pd.DataFrame:
    """Läs CSV med automatisk separator-detektering."""
    try:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        if df.shape[1] == 1 and len(df):
            first = str(df.iloc[0, 0]) if len(df) else ""
            if "\t" in first:
                df = pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig")
        return _clean_columns(df)
    except Exception:
        try:
            return _clean_columns(pd.read_csv(path, dtype=str, sep="\t", engine="python", encoding="utf-8-sig"))
        except Exception:
            return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
