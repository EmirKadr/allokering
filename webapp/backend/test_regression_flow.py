from __future__ import annotations

import importlib
import io
import os
import re
import sys
import time
import unicodedata
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Dict

import pandas as pd
import psycopg
from fastapi.testclient import TestClient
from pgembed import get_server


ROOT = Path(__file__).resolve().parents[1]
JOB_TIMEOUT_SECONDS = 45


def _write_csv(path: Path, rows: list[dict], *, columns: list[str]) -> Path:
    pd.DataFrame(rows, columns=columns).to_csv(path, index=False)
    return path


def _write_prognos_xlsx(path: Path) -> Path:
    rows = [
        ["ignore", "", "", "", "", ""],
        ["ignore", "", "", "", "", ""],
        ["ignore", "Product code", "Product name", "Antal styck", "Antal rader", "Antal butiker"],
        ["ignore", "", "", "", "", ""],
        ["ignore", "A100", "Artikel A", 6, 1, 1],
    ]
    pd.DataFrame(rows).to_excel(path, index=False, header=False, engine="openpyxl")
    return path


def _write_campaign_xlsx(path: Path) -> Path:
    rows = [
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "A200", "", "", "", 3],
    ]
    pd.DataFrame(rows).to_excel(path, index=False, header=False, engine="openpyxl")
    return path


def _build_fixture_files(base: Path) -> Dict[str, Path]:
    base.mkdir(parents=True, exist_ok=True)

    files: Dict[str, Path] = {}

    files["orders_allok"] = _write_csv(
        base / "orders_allok.csv",
        [
            {"Status": 33, "Order nr": "ORD-COMP", "Rad": "1", "Artikel": "A100", "Beställt": 5, "Plock": 5, "Bolag": "GG"},
            {"Status": 33, "Order nr": "ORD-REFILL", "Rad": "1", "Artikel": "A200", "Beställt": 4, "Plock": 1, "Bolag": "GG"},
        ],
        columns=["Status", "Order nr", "Rad", "Artikel", "Beställt", "Plock", "Bolag"],
    )

    files["buffer_allok"] = _write_csv(
        base / "buffer_allok.csv",
        [
            {
                "Lagerplats": "BUF-A1",
                "Pallid": "PAL-A100",
                "Artikel": "A100",
                "Antal": 5,
                "Status": 30,
                "Timestamp": "2026-03-17 10:00:00",
                "Bolag": "GG",
                "Palltyp": "EUR",
            },
            {
                "Lagerplats": "AUTOSTORE-A2",
                "Pallid": "PAL-A200",
                "Artikel": "A200",
                "Antal": 2,
                "Status": 30,
                "Timestamp": "2026-03-17 10:05:00",
                "Bolag": "GG",
                "Palltyp": "EUR",
            },
        ],
        columns=["Lagerplats", "Pallid", "Artikel", "Antal", "Status", "Timestamp", "Bolag", "Palltyp"],
    )

    files["automation"] = _write_csv(
        base / "automation.csv",
        [
            {"Artikel": "A100", "Plocksaldo": 5, "Plockplats": "P-A100", "Saldo autoplock": 0, "Robot": "N", "Bolag": "GG"},
            {"Artikel": "A200", "Plocksaldo": 1, "Plockplats": "P-A200", "Saldo autoplock": 1, "Robot": "Y", "Bolag": "GG"},
        ],
        columns=["Artikel", "Plocksaldo", "Plockplats", "Saldo autoplock", "Robot", "Bolag"],
    )

    files["item"] = _write_csv(
        base / "item.csv",
        [
            {"Artikel": "A100", "Ej Staplingsbar": "N", "Bolag": "GG"},
            {"Artikel": "A200", "Ej Staplingsbar": "Y", "Bolag": "GG"},
        ],
        columns=["Artikel", "Ej Staplingsbar", "Bolag"],
    )

    files["prognos"] = _write_prognos_xlsx(base / "prognos.xlsx")
    files["campaign"] = _write_campaign_xlsx(base / "campaign.xlsx")

    files["orders_controls"] = _write_csv(
        base / "orders_controls.csv",
        [
            {"Status": 33, "Order nr": "STORE1", "Kund.1": "Butik A", "Artikel": "X100", "Beställt": 1, "Plock": 0, "Zon": "A", "Bolag": "GG"},
            {"Status": 33, "Order nr": "HIB1", "Kund.1": "Butik A", "Artikel": "X200", "Beställt": 1, "Plock": 0, "Zon": "A", "Bolag": "GG"},
            {"Status": 33, "Order nr": "DP1", "Kund.1": "Butik D", "Artikel": "X300", "Beställt": 1, "Plock": 0, "Zon": "A", "Bolag": "GG"},
        ],
        columns=["Status", "Order nr", "Kund.1", "Artikel", "Beställt", "Plock", "Zon", "Bolag"],
    )

    files["overview"] = _write_csv(
        base / "overview.csv",
        [
            {
                "Ordernr": "STORE1",
                "Status": 31,
                "Transportör": "Carrier A",
                "Orderdatum": "2026-03-17",
                "Ursprungsdatum": "2026-03-16",
                "Zon": "A",
                "Multi": "M1",
                "Ordertyp": "N",
                "Kund nr": "C100",
                "Kund": "Butik A",
                "Sändningsnr": "SHIP-STORE",
                "Bolag": "GG",
            },
            {
                "Ordernr": "HIB1",
                "Status": 33,
                "Transportör": "Carrier A",
                "Orderdatum": "2026-03-18",
                "Ursprungsdatum": "2026-03-16",
                "Zon": "A",
                "Multi": "",
                "Ordertyp": "HIB",
                "Kund nr": "C100",
                "Kund": "Butik A",
                "Sändningsnr": "SHIP-HIB",
                "Bolag": "GG",
            },
            {
                "Ordernr": "DUP1",
                "Status": 31,
                "Transportör": "Carrier A",
                "Orderdatum": "2026-03-17",
                "Ursprungsdatum": "2026-03-16",
                "Zon": "A",
                "Multi": "",
                "Ordertyp": "N",
                "Kund nr": "C200",
                "Kund": "Butik B",
                "Sändningsnr": "SHIP-DUP",
                "Bolag": "GG",
            },
            {
                "Ordernr": "DUP2",
                "Status": 31,
                "Transportör": "Carrier B",
                "Orderdatum": "2026-03-17",
                "Ursprungsdatum": "2026-03-16",
                "Zon": "A",
                "Multi": "",
                "Ordertyp": "N",
                "Kund nr": "C201",
                "Kund": "Butik C",
                "Sändningsnr": "SHIP-DUP",
                "Bolag": "GG",
            },
            {
                "Ordernr": "DP1",
                "Status": 31,
                "Transportör": "Carrier A",
                "Orderdatum": "2026-03-17",
                "Ursprungsdatum": "2026-03-16",
                "Zon": "A",
                "Multi": "",
                "Ordertyp": "N",
                "Kund nr": "C300",
                "Kund": "Butik D",
                "Sändningsnr": "SHIP-EXPECTED",
                "Bolag": "GG",
            },
        ],
        columns=[
            "Ordernr",
            "Status",
            "Transportör",
            "Orderdatum",
            "Ursprungsdatum",
            "Zon",
            "Multi",
            "Ordertyp",
            "Kund nr",
            "Kund",
            "Sändningsnr",
            "Bolag",
        ],
    )

    files["dispatch"] = _write_csv(
        base / "dispatch.csv",
        [{"Plockpallsnr": "PL-1", "Ordernr": "DP1", "Sändningsnr": "SHIP-WRONG", "Bolag": "GG"}],
        columns=["Plockpallsnr", "Ordernr", "Sändningsnr", "Bolag"],
    )

    files["wms_receive"] = _write_csv(
        base / "wms_receive.csv",
        [{"Typ": 61, "Inköpsnr": "P-900", "Artikel": "A200", "Pallid": "REC-1", "Mottaget": 10, "Bolag": "GG"}],
        columns=["Typ", "Inköpsnr", "Artikel", "Pallid", "Mottaget", "Bolag"],
    )

    files["wms_buffer"] = _write_csv(
        base / "wms_buffer.csv",
        [{"Lagerplats": "BUF-EF", "Pallid": "BUF-1", "Artikel": "A200", "Antal": 8, "Status": 30, "Inköpsnr": "P-900", "Bolag": "GG"}],
        columns=["Lagerplats", "Pallid", "Artikel", "Antal", "Status", "Inköpsnr", "Bolag"],
    )

    return files


def _clear_backend_modules() -> None:
    for name in [
        "backend.main",
        "backend.db",
        "backend.auth",
        "backend.job_queue",
        "backend.job_runner",
        "backend.session_store",
        "backend.gg_kontroll",
    ]:
        sys.modules.pop(name, None)


def _norm(value: object) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text)


def _find_column(df: pd.DataFrame, *candidates: str) -> str:
    normalized = {_norm(col): col for col in df.columns}
    for candidate in candidates:
        key = _norm(candidate)
        if key in normalized:
            return normalized[key]
    raise KeyError(f"Kunde inte hitta någon av kolumnerna {candidates} i {list(df.columns)}")


def _find_sheet(book: dict[str, pd.DataFrame], *candidates: str) -> pd.DataFrame:
    normalized = {_norm(name): df for name, df in book.items()}
    for candidate in candidates:
        key = _norm(candidate)
        if key in normalized:
            return normalized[key]
    raise KeyError(f"Kunde inte hitta något av bladen {candidates} i {list(book.keys())}")


class WebAppRegressionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._saved_env = {key: os.environ.get(key) for key in ("ALLOK_DATABASE_URL", "DATABASE_URL", "ALLOK_STATE_DIR")}
        cls._tmp = TemporaryDirectory()
        cls._tmp_path = Path(cls._tmp.name)
        cls._fixture_dir = cls._tmp_path / "fixtures"
        cls.fixtures = _build_fixture_files(cls._fixture_dir)

        cls._server = get_server(cls._tmp_path / "pgdata", cleanup_mode="delete")
        with psycopg.connect(cls._server.get_uri("postgres"), autocommit=True) as conn:
            with conn.cursor() as cur:
                cur.execute("CREATE DATABASE alloktest")

        os.environ["ALLOK_DATABASE_URL"] = cls._server.get_uri("alloktest")
        os.environ["DATABASE_URL"] = cls._server.get_uri("alloktest")
        os.environ["ALLOK_STATE_DIR"] = str(cls._tmp_path / "state")

        if str(ROOT) not in sys.path:
            sys.path.insert(0, str(ROOT))
        _clear_backend_modules()
        cls._main_mod = importlib.import_module("backend.main")
        cls._client_cm = TestClient(cls._main_mod.app)
        cls.client = cls._client_cm.__enter__()
        cls.headers = cls._login()

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "_client_cm"):
            cls._client_cm.__exit__(None, None, None)
        if hasattr(cls, "_server"):
            cls._server.cleanup()
        if hasattr(cls, "_tmp"):
            cls._tmp.cleanup()
        for key, value in cls._saved_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value

    @classmethod
    def _login(cls) -> dict[str, str]:
        resp = cls.client.post("/api/auth/login", json={"identifier": "emikad", "code": "0012"})
        if resp.status_code != 200:
            raise AssertionError(f"Inloggning misslyckades: {resp.status_code} {resp.text}")
        token = resp.json()["token"]
        return {"Authorization": f"Bearer {token}"}

    def _new_session(self) -> str:
        resp = self.client.post("/api/session", headers=self.headers)
        self.assertEqual(resp.status_code, 200, resp.text)
        return resp.json()["session_id"]

    def _upload_file(self, sid: str, file_key: str, path: Path, page_id: str) -> dict:
        with path.open("rb") as fh:
            resp = self.client.post(
                f"/api/upload/{sid}",
                headers=self.headers,
                data={"file_key": file_key, "page_id": page_id},
                files={"file": (path.name, fh, "application/octet-stream")},
            )
        self.assertEqual(resp.status_code, 200, resp.text)
        return resp.json()

    def _wait_for_job(self, sid: str, expected_result: str | None = None) -> dict:
        deadline = time.time() + JOB_TIMEOUT_SECONDS
        last_status: dict = {}
        while time.time() < deadline:
            resp = self.client.get(f"/api/run/status/{sid}", headers=self.headers)
            self.assertEqual(resp.status_code, 200, resp.text)
            last_status = resp.json()
            state = str(last_status.get("current_job_state") or "").lower()
            running = bool(last_status.get("running")) or state in {"queued", "running"}
            if not running:
                if state == "failed":
                    self.fail(f"Jobbet misslyckades: {last_status}")
                if expected_result:
                    self.assertIn(expected_result, last_status.get("results", []), last_status)
                return last_status
            time.sleep(0.25)
        self.fail(f"Timeout medan jobbet kördes: {last_status}")

    def _download_excel(self, sid: str, result_key: str) -> dict[str, pd.DataFrame]:
        resp = self.client.get(f"/api/download/{sid}/{result_key}", headers=self.headers)
        self.assertEqual(resp.status_code, 200, resp.text[:200])
        return pd.read_excel(io.BytesIO(resp.content), sheet_name=None)

    def test_regression_flow(self) -> None:
        self.assertEqual(self.client.get("/").status_code, 200)
        self.assertEqual(self.client.get("/login.html").status_code, 200)

        me_resp = self.client.get("/api/auth/me", headers=self.headers)
        self.assertEqual(me_resp.status_code, 200, me_resp.text)
        self.assertEqual(me_resp.json()["username"], "emikad")

        chunk_resp = self.client.post(
            "/api/chunked-excel",
            json={"values": ["A100", "A200", "A300"], "chunk_size": 2},
            headers=self.headers,
        )
        self.assertEqual(chunk_resp.status_code, 200, chunk_resp.text)
        self.assertGreater(len(chunk_resp.content), 1000)

        allok_sid = self._new_session()
        for key, fixture_key in (
            ("orders", "orders_allok"),
            ("buffer", "buffer_allok"),
            ("automation", "automation"),
            ("item", "item"),
            ("prognos", "prognos"),
            ("campaign", "campaign"),
        ):
            self._upload_file(allok_sid, key, self.fixtures[fixture_key], "allokering-page")

        upload_status = self.client.get(f"/api/upload/{allok_sid}", headers=self.headers)
        self.assertEqual(upload_status.status_code, 200, upload_status.text)
        uploaded_keys = set(upload_status.json()["files"].keys())
        self.assertTrue({"orders", "buffer", "automation", "item", "prognos", "campaign"}.issubset(uploaded_keys))

        filter_resp = self.client.get(f"/api/filters/{allok_sid}", headers=self.headers)
        self.assertEqual(filter_resp.status_code, 200, filter_resp.text)
        self.assertIn("GG", filter_resp.json().get("bolag", []))

        set_filter_resp = self.client.post(
            f"/api/filters/{allok_sid}",
            headers=self.headers,
            json={"bolag": ["GG"], "ordertyp": []},
        )
        self.assertEqual(set_filter_resp.status_code, 200, set_filter_resp.text)

        run_allok = self.client.post(f"/api/run/allokering/{allok_sid}", headers=self.headers)
        self.assertEqual(run_allok.status_code, 202, run_allok.text)
        allok_status = self._wait_for_job(allok_sid, expected_result="allokerade")
        self.assertIn("refill", allok_status["results"])

        allok_book = self._download_excel(allok_sid, "allokerade")
        allok_sheet = next(iter(allok_book.values()))
        kalltyp_col = _find_column(allok_sheet, "Källtyp")
        self.assertEqual(sorted(set(allok_sheet[kalltyp_col].dropna().astype(str))), ["AUTOSTORE", "HELPALL"])
        self.assertEqual(len(allok_sheet), 3)

        refill_book = self._download_excel(allok_sid, "refill")
        self.assertTrue(any(not sheet.empty for sheet in refill_book.values()))

        refresh_resp = self.client.post(f"/api/ordersaldo/refresh/{allok_sid}", headers=self.headers)
        self.assertEqual(refresh_resp.status_code, 200, refresh_resp.text)
        self.assertEqual(refresh_resp.json()["list1_count"], 1)
        self.assertEqual(refresh_resp.json()["list2_count"], 1)

        list1 = self.client.get(f"/api/ordersaldo/list1/{allok_sid}", headers=self.headers).json()["values"]
        list2 = self.client.get(f"/api/ordersaldo/list2/{allok_sid}", headers=self.headers).json()["values"]
        self.assertEqual(list1, ["ORD-COMP"])
        self.assertEqual(list2, ["A200"])

        prognos_resp = self.client.post(f"/api/result/prognos/{allok_sid}", headers=self.headers)
        self.assertEqual(prognos_resp.status_code, 200, prognos_resp.text)
        prognos_meta = prognos_resp.json()
        self.assertEqual(prognos_meta["rows"], 1)
        self.assertFalse(prognos_meta["partial"])

        prognos_book = self._download_excel(allok_sid, "prognos")
        prognos_sheet = _find_sheet(prognos_book, "Prognos vs Autoplock")
        art_col = _find_column(prognos_sheet, "Artikelnummer")
        self.assertEqual(sorted(prognos_sheet[art_col].astype(str).tolist()), ["A200"])

        control_sid = self._new_session()
        for key, fixture_key in (
            ("orders", "orders_controls"),
            ("overview", "overview"),
            ("dispatch", "dispatch"),
        ):
            self._upload_file(control_sid, key, self.fixtures[fixture_key], "allokering-page")

        hib_resp = self.client.post(f"/api/run/hib-koppling/{control_sid}", headers=self.headers)
        self.assertEqual(hib_resp.status_code, 202, hib_resp.text)
        self._wait_for_job(control_sid, expected_result="hib-koppling")
        hib_book = self._download_excel(control_sid, "hib-koppling")
        hib_sheet = _find_sheet(hib_book, "Ändringar")
        order_col = _find_column(hib_sheet, "ordernummer")
        ship_col = _find_column(hib_sheet, "sändningsnummer")
        date_col = _find_column(hib_sheet, "Orderdatum")
        zone_col = _find_column(hib_sheet, "Zon")
        hib_row = hib_sheet.loc[hib_sheet[order_col].astype(str) == "HIB1"].iloc[0]
        self.assertEqual(str(hib_row[ship_col]).strip(), "SHIP-STORE")
        self.assertEqual(str(hib_row[date_col]).strip(), "2026-03-17")
        self.assertEqual(str(hib_row[zone_col]).strip(), "F")

        orderkontroll_resp = self.client.post(f"/api/run/orderkontroll/{control_sid}", headers=self.headers)
        self.assertEqual(orderkontroll_resp.status_code, 202, orderkontroll_resp.text)
        self._wait_for_job(control_sid, expected_result="orderkontroll")
        orderkontroll_book = self._download_excel(control_sid, "orderkontroll")
        orderkontroll_sheet = _find_sheet(orderkontroll_book, "Orderkontroll")
        orderkontroll_text = orderkontroll_sheet.fillna("").astype(str).agg(" ".join, axis=1).str.cat(sep="\n")
        self.assertIn("SHIP-DUP", orderkontroll_text)
        self.assertIn("HIB1", orderkontroll_text)

        dispatch_resp = self.client.post(f"/api/run/dispatchkontroll/{control_sid}", headers=self.headers)
        self.assertEqual(dispatch_resp.status_code, 202, dispatch_resp.text)
        self._wait_for_job(control_sid, expected_result="dispatchkontroll")
        dispatch_book = self._download_excel(control_sid, "dispatchkontroll")
        dispatch_sheet = _find_sheet(dispatch_book, "Dispatchkontroll")
        order_col = _find_column(dispatch_sheet, "Ordernr")
        expected_col = _find_column(dispatch_sheet, "Översikt sändningsnr")
        dispatch_col = _find_column(dispatch_sheet, "Dispatch sändningsnr")
        dispatch_row = dispatch_sheet.loc[dispatch_sheet[order_col].astype(str) == "DP1"].iloc[0]
        self.assertEqual(str(dispatch_row[expected_col]).strip(), "SHIP-EXPECTED")
        self.assertEqual(str(dispatch_row[dispatch_col]).strip(), "SHIP-WRONG")

        eftersok_sid = self._new_session()
        for key, fixture_key in (("wms_receive", "wms_receive"), ("wms_buffer", "wms_buffer")):
            self._upload_file(eftersok_sid, key, self.fixtures[fixture_key], "eftersok-page")

        eftersok_resp = self.client.post(
            f"/api/run/eftersok/{eftersok_sid}",
            headers=self.headers,
            json={"purchase": "P-900", "article": "A200"},
        )
        self.assertEqual(eftersok_resp.status_code, 202, eftersok_resp.text)
        self._wait_for_job(eftersok_sid, expected_result="eftersok")
        eftersok_book = self._download_excel(eftersok_sid, "eftersok")
        self.assertGreaterEqual(len(eftersok_book), 1)
        self.assertTrue(any(not sheet.empty for sheet in eftersok_book.values()))


def main() -> None:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(WebAppRegressionTest)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        raise SystemExit(1)


if __name__ == "__main__":
    main()
