from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
MODULE_PATH = ROOT / "allokering12.1.py"
SPEC = importlib.util.spec_from_file_location("allokering12_1_module", MODULE_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"Could not load module from {MODULE_PATH}")
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class LyxLogicTest(unittest.TestCase):
    def test_compute_lyx_articles_uses_utbestallt_and_filters_mg_pick_locations(self) -> None:
        saldo_df = pd.DataFrame(
            [
                {"Artikel": "A100", "Plocksaldo": "5", "Utbestallt": "10", "Plockplats": "P1", "Bolag": "MG"},
                {"Artikel": "A101", "Plocksaldo": "6", "Utbestallt": "16", "Plockplats": "P2", "Bolag": "MG"},
                {"Artikel": "A102", "Plocksaldo": "2", "Utbestallt": "0", "Plockplats": "", "Bolag": "MG"},
                {"Artikel": "A103", "Plocksaldo": "2", "Utbestallt": "0", "Plockplats": "P3", "Bolag": "GG"},
                {"Artikel": "A104", "Plocksaldo": "1", "Utbestallt": "0", "Plockplats": "P4", "Bolag": "MG"},
            ]
        )
        max_df = pd.DataFrame(
            [
                {"artikelnummer": "A100", "max": "100"},
                {"artikelnummer": "A101", "max": "100"},
                {"artikelnummer": "A104", "max": "10"},
            ]
        )

        articles, filtered_rows = MODULE.compute_lyx_articles(saldo_df, max_df)

        self.assertEqual(filtered_rows, 3)
        self.assertEqual(articles, ["A100", "A104"])

    def test_compute_lyx_articles_defaults_utbestallt_to_zero(self) -> None:
        saldo_df = pd.DataFrame(
            [
                {"Artikel": "B200", "Plocksaldo": "4", "Plockplats": "P1", "Bolag": "MG"},
                {"Artikel": "B201", "Plocksaldo": "5", "Plockplats": "P2", "Bolag": "MG"},
            ]
        )
        max_df = pd.DataFrame(
            [
                {"artikelnummer": "B200", "max": "20"},
                {"artikelnummer": "B201", "max": "20"},
            ]
        )

        articles, filtered_rows = MODULE.compute_lyx_articles(saldo_df, max_df)

        self.assertEqual(filtered_rows, 2)
        self.assertEqual(articles, ["B200"])


if __name__ == "__main__":
    unittest.main()
