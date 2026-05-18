"""Floden: ett API-handtag per CLI-kommando i allokering12.1.py.

Varje handler tar emot:
  files  - dict {input_key: Path till temporar uppladdad fil}
  params - dict {input_key: stranvarde} for text/nummer/textarea-falt

och returnerar en standarddict:
  {
    "summary": {etikett: varde, ...},   # visas som kort
    "tables":  [(key, label, DataFrame), ...],
    "text":    str | None,              # fritext-rapport (eftersok, vecka27)
    "log":     [str, ...],
  }

All domanlogik kommer fran motorn - inga berakningar dupliceras har.
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
    "Procentuell skillnad (%)", "Anledning", "Galler (INSTEAD R/A)",
]


def _read(path: Path) -> pd.DataFrame:
    return E._read_cli_table(str(path))


def _temp(suffix: str) -> Path:
    """En unik temporar sokvag som annu inte finns (motorn skapar filen)."""
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
            log = [f"Lastningsfonster-lage misslyckades, faller tillbaka: {exc}"]
            report_df, missing_ref = E.build_pafyllnadsprio_report(shortage_df, max_df)
    else:
        report_df, missing_ref = E.build_pafyllnadsprio_report(shortage_df, max_df)

    tables = [("report", "Pafyllnadsprio", report_df)]
    if isinstance(window_map_df, pd.DataFrame):
        tables.append(("window_map", "Lastningsfonster", window_map_df))
    return {
        "summary": {
            "Lage": "Lastningsfonster" if mode == "lastningsfonster" else "Standard",
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
        "summary": {"Andringar": len(changes_df), "Missade avgangar": len(missed_df)},
        "tables": [
            ("changes", "Andringar", changes_df),
            ("missed", "Missade avgangar", missed_df),
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
            "Sandningsrader": len(result.shipment_df),
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
        raise ValueError("Ange bade inkopsnummer och artikelnummer.")
    if "wms_receive" not in files:
        raise ValueError("Mottagningslogg (v_ask_receive_log) kravs.")
    wms_paths = {
        key: (str(files[key]) if key in files else None)
        for key in ("wms_receive", "wms_booking", "wms_buffert", "wms_trans", "wms_pick", "wms_correct")
    }
    result = E.build_eftersok_result(purchase, article, wms_paths)
    return {
        "summary": {"Inkop": purchase, "Artikel": article, "Rapportrader": len(result.report_lines)},
        "tables": [("report", "Eftersok", result.report_df)],
        "text": result.report_text,
        "log": [],
    }


def flow_prognos_report(files: dict, params: dict) -> dict:
    if "prognos" not in files and "campaign" not in files:
        raise ValueError("Ange minst en prognosfil eller en kampanjfil.")
    if "saldo" not in files:
        raise ValueError("Saldo/automation kravs - rapporten filtrerar pa Robot=Y.")
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
    # Skriv till temporara filer - ror aldrig repo-data fran demon.
    result = E.build_observations_update_result(
        buffer_df,
        observations_path=str(_temp(".csv.gz")),
        artikel_max_out=str(_temp(".csv")),
        push_to_github=False,
    )
    return {
        "summary": {
            "Nya observationer": result.new_row_count,
            "Artikel-max rader": result.article_max_rows,
        },
        "tables": [("new_rows", "Nya observationer", result.new_rows_df)],
        "log": [
            "Skrivet till temporara filer (repo-data orord).",
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
            "Hamtade rader": result.fetched_rows,
            "Totalt observationer": result.total_observations,
            "Artikel-max rader": result.article_max_rows,
        },
        "tables": [],
        "log": ["Synkat till temporara filer (repo-data orord, ingen push)."],
    }


def flow_split_values(files: dict, params: dict) -> dict:
    if "values_file" in files:
        values = E._read_cli_text_lines(str(files["values_file"]))
    else:
        raw = params.get("values") or ""
        values = [line.strip() for line in raw.splitlines() if line.strip()]
    if not values:
        raise ValueError("Inga varden angivna - klistra in eller ladda upp en textfil.")
    try:
        chunk_size = int(params.get("chunk_size") or 2000)
    except ValueError:
        chunk_size = 2000
    result = E.build_chunked_values_result(values, chunk_size=max(1, chunk_size))
    return {
        "summary": {
            "Antal varden": result.value_count,
            "Antal kolumner": result.chunk_count,
            "Per kolumn": result.chunk_size,
        },
        "tables": [("report", "Delade varden", result.report_df)],
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
            else "Appen ar uppdaterad."
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
        "description": "Allokera kundorder mot buffertpallar (Helpall -> AutoStore -> Huvudplock, FIFO) med near-miss-loggning, refill och pallplatsberakning.",
        "handler": flow_allocate,
        "inputs": [
            {"key": "orders", "label": "Bestallningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "buffer", "label": "Buffertpallar", "type": "file", "required": True, "detect": ["buffer"]},
            {"key": "saldo", "label": "Saldo / automation", "type": "file", "required": False, "detect": ["automation"]},
            {"key": "items", "label": "Item option", "type": "file", "required": False, "detect": ["item"]},
            {"key": "not_putaway", "label": "Ej inlagrade", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "ordersaldo", "label": "Ordersaldo", "category": "Order & saldo",
        "description": "Berakna kompletta ordrar och artiklar med underskott utifran bestallningslinjer.",
        "handler": flow_ordersaldo,
        "inputs": [
            {"key": "orders", "label": "Bestallningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "saldo", "label": "Saldo / automation (Utbestallt)", "type": "file", "required": False, "detect": ["automation"]},
        ],
    },
    {
        "id": "lyx", "label": "LYX-artiklar", "category": "Order & saldo",
        "description": "Identifiera LYX-artiklar utifran en saldofil och artikel_max-referens.",
        "handler": flow_lyx,
        "inputs": [
            {"key": "saldo", "label": "Saldofil", "type": "file", "required": True, "detect": ["automation", "buffer"]},
            {"key": "max_csv", "label": "artikel_max.csv (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "pafyllnadsprio", "label": "Pafyllnadsprio", "category": "Order & saldo",
        "description": "Prioritera pafyllnad utifran underskott. Med orderoversikt anvands lastningsfonster-lage.",
        "handler": flow_pafyllnadsprio,
        "inputs": [
            {"key": "orders", "label": "Bestallningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "saldo", "label": "Saldo / automation", "type": "file", "required": False, "detect": ["automation"]},
            {"key": "overview", "label": "Orderoversikt (lastningsfonster)", "type": "file", "required": False, "detect": ["overview"]},
            {"key": "max_csv", "label": "artikel_max.csv (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "hib-koppling", "label": "HIB-koppling", "category": "Kontroller",
        "description": "Rakna ut vilka HIB-ordrar som behover kopplas om samt missade avgangar.",
        "handler": flow_hib_koppling,
        "inputs": [
            {"key": "details", "label": "Bestallningslinjer", "type": "file", "required": True, "detect": ["orders"]},
            {"key": "overview", "label": "Orderoversikt", "type": "file", "required": True, "detect": ["overview"]},
        ],
    },
    {
        "id": "overview-check", "label": "Orderoversiktkontroll", "category": "Kontroller",
        "description": "Hitta sandningsnr med flera kunder/transportorer och HIB utan butikssandning.",
        "handler": flow_overview_check,
        "inputs": [
            {"key": "overview", "label": "Orderoversikt", "type": "file", "required": True, "detect": ["overview"]},
            {"key": "details", "label": "Bestallningslinjer (kundnamn)", "type": "file", "required": False, "detect": ["orders"]},
        ],
    },
    {
        "id": "dispatch-check", "label": "Dispatchkontroll", "category": "Kontroller",
        "description": "Jamfor orderoversikt mot dispatchpallar och lista avvikelser.",
        "handler": flow_dispatch_check,
        "inputs": [
            {"key": "overview", "label": "Orderoversikt", "type": "file", "required": True, "detect": ["overview"]},
            {"key": "dispatch", "label": "Dispatchpallar", "type": "file", "required": True, "detect": ["dispatch"]},
            {"key": "details", "label": "Bestallningslinjer (kundnamn)", "type": "file", "required": False, "detect": ["orders"]},
        ],
    },
    {
        "id": "vecka27-check", "label": "Vecka 27-kontroll", "category": "Kontroller",
        "description": "Kontrollera orderrader mot vecka 27-reglerna.",
        "handler": flow_vecka27_check,
        "inputs": [
            {"key": "orders", "label": "Bestallningslinjer", "type": "file", "required": True, "detect": ["orders"]},
        ],
    },
    {
        "id": "eftersok", "label": "Eftersok", "category": "Sokning & prognos",
        "description": "Spara en artikel/pall genom WMS-loggarna utifran inkops- och artikelnummer.",
        "handler": flow_eftersok,
        "inputs": [
            {"key": "purchase", "label": "Inkopsnummer", "type": "text", "required": True},
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
        "id": "prognos-report", "label": "Prognosrapport", "category": "Sokning & prognos",
        "description": "Bygg prognos-/kampanjrapport mot autoplock. Saldo kravs (Robot=Y-filter).",
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
        "description": "Lagg till nya status-30-pallar i observations och racka om artikel_max. Skriver till temporara filer.",
        "handler": flow_observations_update,
        "inputs": [
            {"key": "buffer", "label": "Buffertpallar", "type": "file", "required": True, "detect": ["buffer"]},
        ],
    },
    {
        "id": "observations-sync", "label": "Observations-synk", "category": "Data & verktyg",
        "description": "Hamta observations fran GitHub (eller en lokal fil). Ingen push, skriver till temporara filer.",
        "handler": flow_observations_sync,
        "inputs": [
            {"key": "remote_file", "label": "Lokal observationsfil (valfri)", "type": "file", "required": False, "detect": []},
        ],
    },
    {
        "id": "split-values", "label": "Dela varden", "category": "Data & verktyg",
        "description": "Dela en lang lista av varden i kolumner med valbar kolumnstorlek.",
        "handler": flow_split_values,
        "inputs": [
            {"key": "values", "label": "Varden (ett per rad)", "type": "textarea", "required": False},
            {"key": "values_file", "label": "...eller ladda upp textfil", "type": "file", "required": False, "detect": []},
            {"key": "chunk_size", "label": "Antal per kolumn", "type": "number", "required": False, "default": "2000"},
        ],
    },
    {
        "id": "update-check", "label": "Uppdateringskoll", "category": "Data & verktyg",
        "description": "Kontrollera om en nyare version av appen finns pa GitHub.",
        "handler": flow_update_check,
        "inputs": [],
    },
]

FLOW_BY_ID: dict[str, dict] = {flow["id"]: flow for flow in FLOWS}

# Floden som visas som egna vyer. Allt ovrigt samlas i den kombinerade
# huvudvyn dar filerna delas mellan korningarna.
SOLO_FLOWS = {
    "eftersok",
    "observations-update",
    "observations-sync",
    "split-values",
    "update-check",
}


def public_registry() -> list[dict]:
    """Registret utan handler-referenser - sant till frontenden.

    Varje flode far ett ``view``-falt: ``solo`` (egen vy) eller
    ``combined`` (delar huvudvyn med ovriga combined-floden).
    """
    return [
        {
            **{key: value for key, value in flow.items() if key != "handler"},
            "view": "solo" if flow["id"] in SOLO_FLOWS else "combined",
        }
        for flow in FLOWS
    ]
