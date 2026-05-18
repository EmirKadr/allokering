"""Laddar allokerings-motorn från allokering12.1.py.

Modulfilen heter ``allokering12.1.py`` vilket inte är ett giltigt
Python-identifierare, så den laddas via importlib från sökvägen. All
domänlogik återexporteras härifrån så att API-lagret aldrig behöver
importera tkinter-widgets eller CLI-parsern direkt.

Detta följer AGENTS.md: GUI/CLI/API delar exakt samma underliggande motor.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

# allokering/web/backend/engine.py  ->  allokering/
PROJECT_ROOT = Path(__file__).resolve().parents[2]
ENGINE_FILE = PROJECT_ROOT / "allokering12.1.py"


def _load_engine():
    if not ENGINE_FILE.exists():
        raise FileNotFoundError(f"Hittar inte motorn: {ENGINE_FILE}")
    # app_info / analytics_store / update_service importeras av motorn.
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    spec = importlib.util.spec_from_file_location("allokering_engine", ENGINE_FILE)
    module = importlib.util.module_from_spec(spec)
    sys.modules["allokering_engine"] = module
    spec.loader.exec_module(module)
    return module


engine = _load_engine()

# --- Domänfunktioner som API-lagret använder ---------------------------------
read_table = engine._read_cli_table
normalize_saldo = engine.normalize_saldo
normalize_items = engine.normalize_items
normalize_not_putaway = engine.normalize_not_putaway
allocate = engine.allocate
calculate_refill = engine.calculate_refill
compute_pallet_spaces = engine.compute_pallet_spaces
reclassify_skrymmande = engine.App._reclassify_skrymmande  # staticmethod
merge_item_flags = engine._merge_item_flags
open_df_in_excel = engine._open_df_in_excel
build_observations_update_result = engine.build_observations_update_result
fetch_observations_from_github = engine.fetch_observations_from_github

APP_VERSION = engine.APP_VERSION
APP_TITLE = engine.APP_TITLE


def detect_file_type(path: str):
    """Återanvänder GUI:ts filtypsdetektering.

    ``App._detect_file_type`` rör aldrig ``self``, så den kan anropas
    obunden med ``None`` som self.
    """
    return engine.App._detect_file_type(None, path)
