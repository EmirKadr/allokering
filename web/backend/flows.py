"""Flöden: ett API-handtag per CLI-kommando i allokering12.1.py.

Varje handler tar emot:
  files  - dict {input_key: Path till temporär uppladdad fil}
  params - dict {input_key: strängvärde} för text/nummer/textarea-fält

och returnerar en standarddict:
  {
    "summary": {etikett: värde, ...},   # visas som kort
    "tables":  [(key, label, DataFrame), ...],
    "text":    str | None,              # fritext-rapport (eftersök, vecka27)
    "log":     [str, ...],
  }

All domänlogik kommer från motorn - inga beräkningar dupliceras här.
"""
from __future__ import annotations

import tempfile
import uuid
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

from engine import engine as E

NEAR_MISS_COLUMNS = [
    "Artikel", "OrderID", "OrderRad", "PallID", "Kallplats", "Mottagen",
    "Behov_vid_tillfallet", "Pall_kvantitet", "Skillnad",
    "Procentuell skillnad (%)", "Anledning", "Gäller (INSTEAD R/A)",
]


def _read(path: Path) -> pd.DataFrame:
    return E._read_cli_table(str(path))


def _temp(suffix: str) -> Path:
    """En unik temporär sökväg som ännu inte finns (motorn skapar filen)."""
    return Path(tempfile.gettempdir()) / f"allok_{uuid.uuid4().hex}{suffix}"


# --- Floden ------------------------------------------------------------------

def flow_allocate(files: dict, params: dict) -> dict:
    orders_raw = _read(files["orders"])
    buffer_raw = _read(files["buffer"])
    saldo_norm = E.normalize_saldo(_read(files["saldo"])) if "saldo" in files else None
    item_norm = E.normalize_items(_read(files["items"])) if "items" in files else None
    not_putaway_norm = (
        E.normalize_not_putaway(_read(files["not_putaway"])) if "not_putaway" in files else None
    )

    log: list[str] = []
    result_df, near_miss_df = E.allocate(orders_raw, buffer_raw, log=log.append)
    result_df = E.App._reclassify_skrymmande(result_df, saldo_norm)
    result_df = E._merge_item_flags(result_df, item_norm)
    if near_miss_df.empty and len(near_miss_df.columns) == 0:
        near_miss_df = pd.DataFrame(columns=NEAR_MISS_COLUMNS)

    refill_hp_df, refill_autostore_df = E.calculate_refill(
        result_df, buffer_raw, saldo_df=saldo_norm, not_putaway_df=not_putaway_norm,
    )
    pallet_spaces_df = E.compute_pallet_spaces(result_df)

    return {
        "summary": {
            "Resultatrader": len(result_df),
            "Near-miss": len(near_miss_df),
            "Refill Huvudplock": len(refill_hp_df),
            "Refill AutoStore": len(refill_autostore_df),
            "Pallplatser": len(pallet_spaces_df),
        },
        "tables": [
            ("result", "Resultat", result_df),
            ("near_miss", "Near-miss", near_miss_df),
            ("refill_hp", "Refill Huvudplock", refill_hp_df),
            ("refill_autostore", "Refill AutoStore", refill_autostore_df),
            ("pallet_spaces", "Pallplatser", pallet_spaces_df),
        ],
        "log": log,
    }


def flow_ordersaldo(files: dict, params: dict) -> dict:
    orders_df = _read(files["orders"])
    column_names = E._find_ordersaldo_columns(orders_df)
    utbest_map = E.utbest_per_article(_read(files["saldo"])) if "saldo" in files else {}
    complete_orders, shortage_df = E.compute_ordersaldo_data(
        orders_df, utbest_map=utbest_map, column_names=column_names,
    )
    return {
        "summary": {
            "Kompletta ordrar": len(complete_orders),
            "Artiklar med underskott": len(shortage_df),
        },
        "tables": [
            ("complete", "Kompletta ordrar", pd.DataFrame({"Ordernr": complete_orders})),
            ("shortage", "Underskott", E._df_with_named_index(shortage_df, "Artikel")),
        ],
        "log": [],
    }


def flow_lyx(files: dict, params: dict) -> dict:
    saldo_df = _read(files["saldo"])
    max_path = files["max_csv"] if "max_csv" in files else E._resolve_max_csv_path(None)
    max_df = _read(Path(max_path))
    articles, filtered_rows = E.compute_lyx_articles(saldo_df, max_df)
    return {
        "summary": {"LYX-artiklar": len(articles), "Filtrerade rader": filtered_rows},
        "tables": [("articles", "LYX-artiklar", pd.DataFrame({"Artikel": articles}))],
        "log": [],
    }


def flow_pafyllnadsprio(files: dict, params: dict) -> dict:
    orders_df = _read(files["orders"])
    column_names = E._find_ordersaldo_columns(orders_df)
    utbest_map = E.utbest_per_article(_read(files["saldo"])) if "saldo" in files else {}
    _, shortage_df = E.compute_ordersaldo_data(
        orders_df, utbest_map=utbest_map, column_names=column_names,
    )
    max_path = files["max_csv"] if "max_csv" in files else E._resolve_max_csv_path(None)
    max_df = _read(Path(max_path))

    log: list[str] = []
    window_map_df = None
    mode = "fallback"
    if "overview" in files:
        try:
            overview_df = _read(files["overview"])
            report_df, _bold, log, missing_ref, window_map_df = (
                E.build_pafyllnadsprio_lastningsfonster_report(
                    orders_df, shortage_df, overview_df, max_df, column_names=column_names,
                )
            )
            mode = "lastningsfonster"
        except Exception as exc:  # noqa: BLE001
            log = [f"Lastningsfönster-läge misslyckades, faller tillbaka: {exc}"]
            report_df, missing_ref = E.build_pafyllnadsprio_report(shortage_df, max_df)
    else:
        report_df, missing_ref = E.build_pafyllnadsprio_report(shortage_df, max_df)

    tables = [("report", "Påfyllnadsprio", report_df)]
    if isinstance(window_map_df, pd.DataFrame):
        tables.append(("window_map", "Lastningsfönster", window_map_df))
    return {
        "summary": {
            "Läge": "Lastningsfönster" if mode == "lastningsfonster" else "Standard",
            "Rapportrader": len(report_df),
            "Saknad referens": int(missing_ref),
        },
        "tables": tables,
        "log": log,
    }


def flow_hib_koppling(files: dict, params: dict) -> dict:
    details_df = _read(files["details"])
    overview_df = _read(files["overview"])
    changes_df = E.compute_hib_koppling(details_df, overview_df)
    missed_df = E.compute_missed_departures(details_df, overview_df)
    return {
        "summary": {"Ändringar": len(changes_df), "Missade avgångar": len(missed_df)},
        "tables": [
            ("changes", "Ändringar", changes_df),
            ("missed", "Missade avgångar", missed_df),
        ],
        "log": [],
    }


def flow_overview_check(files: dict, params: dict) -> dict:
    overview_df = _read(files["overview"])
    details_df = _read(files["details"]) if "details" in files else None
    result = E.build_overview_check_result(overview_df, details_df=details_df)
    sheets = E._build_overview_check_sheets(result)
    tables = [(key.lower().replace(" ", "_"), key, df) for key, df in sheets.items()]
    return {
        "summary": {
            "Sändningsrader": len(result.shipment_df),
            "HIB-rader": len(result.hib_df),
        },
        "tables": tables,
        "log": list(result.log_lines or []),
    }


def flow_dispatch_check(files: dict, params: dict) -> dict:
    overview_df = _read(files["overview"])
    dispatch_df = _read(files["dispatch"])
    details_df = _read(files["details"]) if "details" in files else None
    result = E.build_dispatch_check_result(overview_df, dispatch_df, details_df=details_df)
    return {
        "summary": {"Avvikelser": len(result.diff_df)},
        "tables": [("diff", "Dispatchavvikelser", result.diff_df)],
        "log": list(result.log_lines or []),
    }


def flow_vecka27_check(files: dict, params: dict) -> dict:
    orders_df = _read(files["orders"])
    result = E.build_vecka27_check_result(orders_df)
    return {
        "summary": {"Avvikelser": len(result.deviations)},
        "tables": [("report", "Avvikelser", result.report_df)],
        "text": result.report_text,
        "log": list(result.log_lines or []),
    }


def flow_eftersok(files: dict, params: dict) -> dict:
    purchase = (params.get("purchase") or "").strip()
    article = (params.get("article") or "").strip()
    if not purchase or not article:
        raise ValueError("Ange både inköpsnummer och artikelnummer.")
    if "wms_receive" not in files:
        raise ValueError("Mottagningslogg (v_ask_receive_log) krävs.")
    wms_paths = {
        key: (str(files[key]) if key in files else None)
        for key in ("wms_receive", "wms_booking", "wms_buffert", "wms_trans", "wms_pick", "wms_correct")
    }
    result = E.build_eftersok_result(purchase, article, wms_paths)
    return {
        "summary": {"Inköp": purchase, "Artikel": article, "Rapportrader": len(result.report_lines)},
        "tables": [("report", "Eftersök", result.report_df)],
        "text": result.report_text,
        "log": [],
    }


def flow_prognos_report(files: dict, params: dict) -> dict:
    if "prognos" not in files and "campaign" not in files:
        raise ValueError("Ange minst en prognosfil eller en kampanjfil.")
    if "saldo" not in files:
        raise ValueError("Saldo/automation krävs - rapporten filtrerar på Robot=Y.")
    prognos_df = E._load_prognos_cli_source(str(files["prognos"])) if "prognos" in files else None
    campaign_df = E._load_campaign_cli_source(str(files["campaign"])) if "campaign" in files else None
    saldo_df = _read(files["saldo"])
    buffer_df = _read(files["buffer"]) if "buffer" in files else None
    result = E.build_prognos_report_result(
        prognos_df=prognos_df, campaign_df=campaign_df, saldo_df=saldo_df, buffer_df=buffer_df,
    )
    meta = result.meta if isinstance(result.meta, dict) else {}
    return {
        "summary": {
            "Rapportrader": len(result.report_df),
            "Kombinerade rader": len(result.combined_df),
            "Partiell": "Ja" if meta.get("partial") == "yes" else "Nej",
        },
        "tables": [
            ("report", "Prognos vs Autoplock", result.report_df),
            ("combined", "Kombinerat underlag", result.combined_df),
        ],
        "log": list(result.log_lines or []),
    }


def flow_observations_update(files: dict, params: dict) -> dict:
    buffer_df = _read(files["buffer"])
    # Skriv till temporära filer - rör aldrig repo-data från demon.
    result = E.build_observations_update_result(
        buffer_df,
        observations_path=str(_temp(".csv.gz")),
        artikel_max_out=str(_temp(".csv")),
        push_to_github=False,
    )
    return {
        "summary": {
            "Nya observationer": result.new_row_count,
            "Artikel-max-rader": result.article_max_rows,
        },
        "tables": [("new_rows", "Nya observationer", result.new_rows_df)],
        "log": [
            "Skrivet till temporära filer (repo-data orörd).",
            f"Observations: {result.observations_path}",
            f"Artikel-max: {result.article_max_path}",
        ],
    }


def flow_observations_sync(files: dict, params: dict) -> dict:
    result = E.build_observations_sync_result(
        observations_path=str(_temp(".csv.gz")),
        artikel_max_out=str(_temp(".csv")),
        remote_file=str(files["remote_file"]) if "remote_file" in files else None,
        push_orphaned=False,
    )
    return {
        "summary": {
            "Hämtade rader": result.fetched_rows,
            "Totalt observationer": result.total_observations,
            "Artikel-max-rader": result.article_max_rows,
        },
        "tables": [],
        "log": ["Synkat till temporära filer (repo-data orörd, ingen push)."],
    }


def flow_split_values(files: dict, params: dict) -> dict:
    if "values_file" in files:
        values = E._read_cli_text_lines(str(files["values_file"]))
    else:
        raw = params.get("values") or ""
        values = [line.strip() for line in raw.splitlines() if line.strip()]
    if not values:
        raise ValueError("Inga värden angivna - klistra in eller ladda upp en textfil.")
    try:
        chunk_size = int(params.get("chunk_size") or 2000)
    except ValueError:
        chunk_size = 2000
    result = E.build_chunked_values_result(values, chunk_size=max(1, chunk_size))
    return {
        "summary": {
            "Antal värden": result.value_count,
            "Antal kolumner": result.chunk_count,
            "Per kolumn": result.chunk_size,
        },
        "tables": [("report", "Delade värden", result.report_df)],
        "log": [],
    }


def flow_update_check(files: dict, params: dict) -> dict:
    result = E.build_update_check_cli_result()
    return {
        "summary": {
            "Ny version finns": "Ja" if result.has_update else "Nej",
            "Nuvarande version": result.current_version,
            "Senaste version": result.latest_version,
        },
        "tables": [],
        "text": (
            f"Release: {result.release_url}\nInstallerare: {result.installer_name}"
            if result.has_update
            else "Appen är uppdaterad."
        ),
        "log": [],
    }


# --- Registry ----------------------------------------------------------------
# Varje post: id, label, category, description, inputs[], handler.
# input.type: file | text | number | textarea
# input.detect: lista av filtyper (fran motorns _detect_file_type) som auto-routas hit.

FLOWS: list[dict] = [
    {
        "id": "allocate", "label": "Allokering", "category": "Allokering",
        "description": "Allokera kundorder mot buffertpallar (Helpall -> AutoStore -> Huvudplock, FIFO) med near-miss-loggning, refill och pallplatsberäkning.",
        "handler": flow_allocate,
        "inputs": [
            {"key": "orders", "label": "Beställningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "buffer", "label": "Buffertpallar", "type": "file", "required": True, "detect": ["buffer"]},
            {"key": "saldo", "label": "Saldo / automation", "type": "file", "required": False, "detect": ["automation"]},
            {"key": "items", "label": "Item option", "type": "file", "required": False, "detect": ["item"]},
            {"key": "not_putaway", "label": "Ej inlagrade", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "ordersaldo", "label": "Ordersaldo", "category": "Order & saldo",
        "description": "Beräkna kompletta ordrar och artiklar med underskott utifrån beställningslinjer.",
        "handler": flow_ordersaldo,
        "inputs": [
            {"key": "orders", "label": "Beställningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "saldo", "label": "Saldo / automation (Utbestallt)", "type": "file", "required": False, "detect": ["automation"]},
        ],
    },
    {
        "id": "lyx", "label": "LYX-artiklar", "category": "Order & saldo",
        "description": "Identifiera LYX-artiklar utifrån en saldofil och artikel_max-referens.",
        "handler": flow_lyx,
        "inputs": [
            {"key": "saldo", "label": "Saldofil", "type": "file", "required": True, "detect": ["automation", "buffer"]},
            {"key": "max_csv", "label": "artikel_max.csv (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "pafyllnadsprio", "label": "Påfyllnadsprio", "category": "Order & saldo",
        "description": "Prioritera påfyllnad utifrån underskott. Med orderöversikt används lastningsfönster-läge.",
        "handler": flow_pafyllnadsprio,
        "inputs": [
            {"key": "orders", "label": "Beställningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "saldo", "label": "Saldo / automation", "type": "file", "required": False, "detect": ["automation"]},
            {"key": "overview", "label": "Orderöversikt (lastningsfönster)", "type": "file", "required": False, "detect": ["overview"]},
            {"key": "max_csv", "label": "artikel_max.csv (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "hib-koppling", "label": "HIB-koppling", "category": "Kontroller",
        "description": "Räkna ut vilka HIB-ordrar som behöver kopplas om samt missade avgångar.",
        "handler": flow_hib_koppling,
        "inputs": [
            {"key": "details", "label": "Beställningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "overview", "label": "Orderöversikt", "type": "file", "required": True, "detect": ["overview"]},
        ],
    },
    {
        "id": "overview-check", "label": "Orderöversiktkontroll", "category": "Kontroller",
        "description": "Hitta sändningsnr med flera kunder/transportörer och HIB utan butikssändning.",
        "handler": flow_overview_check,
        "inputs": [
            {"key": "overview", "label": "Orderöversikt", "type": "file", "required": True, "detect": ["overview"]},
            {"key": "details", "label": "Beställningslinjer (kundnamn)", "type": "file", "required": False, "detect": ["orders"]},
        ],
    },
    {
        "id": "dispatch-check", "label": "Dispatchkontroll", "category": "Kontroller",
        "description": "Jämför orderöversikt mot dispatchpallar och lista avvikelser.",
        "handler": flow_dispatch_check,
        "inputs": [
            {"key": "overview", "label": "Orderöversikt", "type": "file", "required": True, "detect": ["overview"]},
            {"key": "dispatch", "label": "Dispatchpallar", "type": "file", "required": True, "detect": ["dispatch"]},
            {"key": "details", "label": "Beställningslinjer (kundnamn)", "type": "file", "required": False, "detect": ["orders"]},
        ],
    },
    {
        "id": "vecka27-check", "label": "Vecka 27-kontroll", "category": "Kontroller",
        "description": "Kontrollera orderrader mot vecka 27-reglerna.",
        "handler": flow_vecka27_check,
        "inputs": [
            {"key": "orders", "label": "Beställningslinjer", "type": "file", "required": True, "detect": ["orders"]},
        ],
    },
    {
        "id": "eftersok", "label": "Eftersök", "category": "Sökning & prognos",
        "description": "Spåra en artikel/pall genom WMS-loggarna utifrån inköps- och artikelnummer.",
        "handler": flow_eftersok,
        "inputs": [
            {"key": "purchase", "label": "Inköpsnummer", "type": "text", "required": True},
            {"key": "article", "label": "Artikelnummer", "type": "text", "required": True},
            {"key": "wms_receive", "label": "Mottagningslogg", "type": "file", "required": True, "detect": ["wms_receive"]},
            {"key": "wms_booking", "label": "Inlagringslogg", "type": "file", "required": False, "detect": ["wms_booking"]},
            {"key": "wms_buffert", "label": "Buffertpallar", "type": "file", "required": False, "detect": ["buffer"]},
            {"key": "wms_trans", "label": "Transaktionslogg", "type": "file", "required": False, "detect": ["wms_trans"]},
            {"key": "wms_pick", "label": "Plocklogg", "type": "file", "required": False, "detect": ["wms_pick"]},
            {"key": "wms_correct", "label": "Korrigeringslogg", "type": "file", "required": False, "detect": ["wms_correct"]},
        ],
    },
    {
        "id": "prognos-report", "label": "Prognosrapport", "category": "Sökning & prognos",
        "description": "Bygg prognos-/kampanjrapport mot autoplock. Saldo krävs (Robot=Y-filter).",
        "handler": flow_prognos_report,
        "inputs": [
            {"key": "prognos", "label": "Prognosfil", "type": "file", "required": False, "detect": ["prognos"]},
            {"key": "campaign", "label": "Kampanjfil", "type": "file", "required": False, "detect": ["campaign"]},
            {"key": "saldo", "label": "Saldo / automation", "type": "file", "required": True, "detect": ["automation"]},
            {"key": "buffer", "label": "Buffertpallar", "type": "file", "required": False, "detect": ["buffer"]},
        ],
    },
    {
        "id": "observations-update", "label": "Observations-uppdatering", "category": "Data & verktyg",
        "description": "Lägg till nya status-30-pallar i observations och räkna om artikel_max. Skriver till temporära filer.",
        "handler": flow_observations_update,
        "inputs": [
            {"key": "buffer", "label": "Buffertpallar", "type": "file", "required": True, "detect": ["buffer"]},
        ],
    },
    {
        "id": "observations-sync", "label": "Observations-synk", "category": "Data & verktyg",
        "description": "Hämta observations från GitHub (eller en lokal fil). Ingen push, skriver till temporära filer.",
        "handler": flow_observations_sync,
        "inputs": [
            {"key": "remote_file", "label": "Lokal observationsfil (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "split-values", "label": "Dela värden", "category": "Data & verktyg",
        "description": "Dela en lång lista av värden i kolumner med valbar kolumnstorlek.",
        "handler": flow_split_values,
        "inputs": [
            {"key": "values", "label": "Värden (ett per rad)", "type": "textarea", "required": False},
            {"key": "values_file", "label": "...eller ladda upp textfil", "type": "file", "required": False, "detect": []},
            {"key": "chunk_size", "label": "Antal per kolumn", "type": "number", "required": False, "default": "2000"},
        ],
    },
    {
        "id": "update-check", "label": "Uppdateringskoll", "category": "Data & verktyg",
        "description": "Kontrollera om en nyare version av appen finns på GitHub.",
        "handler": flow_update_check,
        "inputs": [],
    },
]

FLOW_BY_ID: dict[str, dict] = {flow["id"]: flow for flow in FLOWS}

# Flöden som visas som egna vyer. Allt övrigt samlas i den kombinerade
# huvudvyn där filerna delas mellan körningarna.
SOLO_FLOWS = {
    "eftersok",
    "observations-update",
    "observations-sync",
    "split-values",
    "update-check",
}


# Gemensam datapool: combined-flöden laddar upp filerna EN gång här, och
# varje flödes filinput mappas till en pool-nyckel. Endast "details" skiljer
# sig från sin pool-nyckel (samma filformat som "orders").
DATA_POOL: list[dict] = [
    {"key": "orders", "label": "Beställningslinjer", "detect": ["orders"]},
    {"key": "buffer", "label": "Buffertpallar", "detect": ["buffer"]},
    {"key": "saldo", "label": "Saldo / automation", "detect": ["automation"]},
    {"key": "overview", "label": "Orderöversikt", "detect": ["overview"]},
    {"key": "dispatch", "label": "Dispatchpallar", "detect": ["dispatch"]},
    {"key": "items", "label": "Item option", "detect": ["item"]},
    {"key": "not_putaway", "label": "Ej inlagrade", "detect": []},
    {"key": "prognos", "label": "Prognosfil", "detect": ["prognos"]},
    {"key": "campaign", "label": "Kampanjfil", "detect": ["campaign"]},
    {"key": "max_csv", "label": "artikel_max.csv", "detect": []},
]

_POOL_KEY_OVERRIDE = {"details": "orders"}


def _pool_key(input_key: str) -> str:
    return _POOL_KEY_OVERRIDE.get(input_key, input_key)


def public_registry() -> list[dict]:
    """Registret utan handler-referenser - sant till frontenden.

    Varje flöde får ett ``view``-fält: ``solo`` (egen vy) eller
    ``combined`` (delar huvudvyn med övriga combined-flöden). Filinputs i
    combined-flöden får en ``pool``-nyckel mot den gemensamma datapoolen.
    """
    result: list[dict] = []
    for flow in FLOWS:
        view = "solo" if flow["id"] in SOLO_FLOWS else "combined"
        inputs: list[dict] = []
        for inp in flow["inputs"]:
            new_inp = dict(inp)
            if view == "combined" and inp.get("type") == "file":
                new_inp["pool"] = _pool_key(inp["key"])
            inputs.append(new_inp)
        result.append({
            **{key: value for key, value in flow.items() if key != "handler"},
            "inputs": inputs,
            "view": view,
        })
    return result


def public_pool() -> list[dict]:
    """Datapoolens slots - sänt till frontenden för den kombinerade vyn."""
    return [dict(slot) for slot in DATA_POOL]
