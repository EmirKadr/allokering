from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
MODULE_PATH = ROOT / "allokering12.1.py"
SPEC = importlib.util.spec_from_file_location("allokering12_1_pafyllnadsprio_module", MODULE_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"Could not load module from {MODULE_PATH}")
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class PafyllnadsprioLogicTest(unittest.TestCase):
    def test_compute_ordersaldo_data_uses_utbestallt_and_max_plocksaldo(self) -> None:
        orders_df = pd.DataFrame(
            [
                {"Ordernr": "O1", "Artikel": "A100", "Beställt": "20", "Plock": "10"},
                {"Ordernr": "O1", "Artikel": "A100", "Beställt": "20", "Plock": "30"},
                {"Ordernr": "O2", "Artikel": "B200", "Beställt": "10", "Plock": "20"},
            ]
        )

        complete_orders, shortage_df = MODULE.compute_ordersaldo_data(
            orders_df,
            utbest_map={"A100": 5.0},
        )

        self.assertEqual(complete_orders, ["O2"])
        self.assertEqual(shortage_df.index.tolist(), ["A100"])
        self.assertEqual(float(shortage_df.loc["A100", "Total beställt"]), 40.0)
        self.assertEqual(float(shortage_df.loc["A100", "Tillgängligt saldo (Plock)"]), 30.0)
        self.assertEqual(float(shortage_df.loc["A100", "Utbeställt"]), 5.0)
        self.assertEqual(float(shortage_df.loc["A100", "Underskott"]), 15.0)

    def test_build_pafyllnadsprio_report_uses_exclusive_priority_buckets(self) -> None:
        shortage_df = pd.DataFrame(
            {
                "Underskott": [25, 40, 55, 70, 71],
            },
            index=["A1", "A2", "A3", "A4", "A5"],
        )
        max_df = pd.DataFrame(
            [
                {"artikelnummer": "A1", "max": "100"},
                {"artikelnummer": "A2", "max": "100"},
                {"artikelnummer": "A3", "max": "100"},
                {"artikelnummer": "A4", "max": "100"},
                {"artikelnummer": "A5", "max": "100"},
            ]
        )

        report_df, missing_reference_count = MODULE.build_pafyllnadsprio_report(shortage_df, max_df)

        self.assertEqual(missing_reference_count, 0)
        self.assertEqual(report_df["ALLA"].tolist(), ["A1", "A2", "A3", "A4", "A5"])
        self.assertEqual(report_df["PRIO 1"].tolist(), ["A1", "", "", "", ""])
        self.assertEqual(report_df["PRIO 2"].tolist(), ["A2", "", "", "", ""])
        self.assertEqual(report_df["PRIO 3"].tolist(), ["A3", "", "", "", ""])
        self.assertEqual(report_df["PRIO 4"].tolist(), ["A4", "", "", "", ""])
        self.assertEqual(report_df["PRIO 5"].tolist(), ["A5", "", "", "", ""])

    def test_build_pafyllnadsprio_report_puts_missing_or_invalid_reference_in_prio5(self) -> None:
        shortage_df = pd.DataFrame(
            {
                "Underskott": [10, 12],
            },
            index=["B1", "B2"],
        )
        max_df = pd.DataFrame(
            [
                {"artikelnummer": "B1", "max": "0"},
            ]
        )

        report_df, missing_reference_count = MODULE.build_pafyllnadsprio_report(shortage_df, max_df)

        self.assertEqual(missing_reference_count, 2)
        self.assertEqual(report_df["ALLA"].tolist(), ["B1", "B2"])
        self.assertEqual(report_df["PRIO 5"].tolist(), ["B1", "B2"])

    def test_build_pafyllnadsprio_report_returns_empty_frame_for_empty_shortage(self) -> None:
        report_df, missing_reference_count = MODULE.build_pafyllnadsprio_report(
            pd.DataFrame(columns=["Underskott"]),
            pd.DataFrame(columns=["artikelnummer", "max"]),
        )

        self.assertEqual(missing_reference_count, 0)
        self.assertEqual(report_df.columns.tolist(), MODULE.PAFYLLNADSPRIO_COLUMNS)
        self.assertTrue(report_df.empty)


if __name__ == "__main__":
    unittest.main()
