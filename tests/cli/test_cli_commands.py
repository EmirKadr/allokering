from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


def _write_csv(path: Path, rows: list[dict]) -> Path:
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def _write_tsv(path: Path, rows: list[dict], columns: list[str] | None = None) -> Path:
    df = pd.DataFrame(rows)
    if columns is not None:
        df = df.reindex(columns=columns)
    df.to_csv(path, index=False, encoding="utf-8", sep="\t")
    return path


def _write_gzip_csv(path: Path, rows: list[dict]) -> Path:
    pd.DataFrame(rows).to_csv(path, index=False, compression="gzip")
    return path


def _write_text(path: Path, text: str) -> Path:
    path.write_text(text, encoding="utf-8")
    return path


def test_allocate_cli_writes_result_and_near_miss_files(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders.csv",
        [
            {"Artikel": "A1", "Antal": 6, "Ordernr": "O1", "Radnr": "1", "Status": 30, "Zon": "A"},
        ],
    )
    buffer_path = _write_csv(
        tmp_path / "buffer.csv",
        [
            {"Artikel": "A1", "Antal": 4, "Lagerplats": "H1", "Datum/Tid": "2024-01-01 10:00", "PallID": "P1", "Status": 29},
            {"Artikel": "A1", "Antal": 3, "Lagerplats": "AUTOSTORE-1", "Datum/Tid": "2024-01-02 10:00", "PallID": "P2", "Status": 29},
        ],
    )
    result_out = tmp_path / "allocated.csv"
    near_out = tmp_path / "near.csv"

    completed = run_cli_cmd(
        "allocate",
        "--orders",
        str(orders_path),
        "--buffer",
        str(buffer_path),
        "--result-out",
        str(result_out),
        "--near-miss-out",
        str(near_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "allocate"
    assert payload["result_rows"] == 2
    assert payload["near_miss_rows"] == 0

    result_df = pd.read_csv(result_out, dtype=str, encoding="utf-8-sig")
    assert result_df["Källtyp"].tolist() == ["HELPALL", "AUTOSTORE"]
    assert result_df["Källa"].tolist() == ["P1", "P2"]

    near_df = pd.read_csv(near_out, dtype=str, encoding="utf-8-sig")
    assert near_df.empty


def test_allocate_cli_ignores_order_rows_above_status_31(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders_status.csv",
        [
            {"Artikel": "A31", "Antal": 1, "Ordernr": "O31", "Radnr": "1", "Status": 31, "Zon": "A"},
            {"Artikel": "A32", "Antal": 1, "Ordernr": "O32", "Radnr": "1", "Status": 32, "Zon": "A"},
            {"Artikel": "A40", "Antal": 1, "Ordernr": "O40", "Radnr": "1", "Status": 40, "Zon": "A"},
        ],
    )
    buffer_path = _write_csv(
        tmp_path / "buffer_status.csv",
        [
            {"Artikel": "A31", "Antal": 1, "Lagerplats": "H31", "Datum/Tid": "2024-01-01 10:00", "PallID": "P31", "Status": 29},
            {"Artikel": "A32", "Antal": 1, "Lagerplats": "H32", "Datum/Tid": "2024-01-01 10:00", "PallID": "P32", "Status": 29},
            {"Artikel": "A40", "Antal": 1, "Lagerplats": "H40", "Datum/Tid": "2024-01-01 10:00", "PallID": "P40", "Status": 29},
        ],
    )
    result_out = tmp_path / "allocated_status.csv"

    completed = run_cli_cmd(
        "allocate",
        "--orders",
        str(orders_path),
        "--buffer",
        str(buffer_path),
        "--result-out",
        str(result_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["result_rows"] == 1

    result_df = pd.read_csv(result_out, dtype=str, encoding="utf-8-sig")
    assert result_df["Ordernr"].tolist() == ["O31"]
    assert result_df["Artikel"].tolist() == ["A31"]


def test_ordersaldo_cli_writes_shortage_report(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "ordersaldo.csv",
        [
            {"Ordernr": "O1", "Artikel": "A100", "Antal": 20, "Plock": 10},
            {"Ordernr": "O1", "Artikel": "A100", "Antal": 20, "Plock": 30},
            {"Ordernr": "O2", "Artikel": "B200", "Antal": 10, "Plock": 20},
        ],
    )
    saldo_path = _write_csv(
        tmp_path / "saldo.csv",
        [
            {"Artikel": "A100", "Utbestallt": 5},
        ],
    )
    complete_out = tmp_path / "complete.txt"
    shortage_out = tmp_path / "shortage.csv"

    completed = run_cli_cmd(
        "ordersaldo",
        "--orders",
        str(orders_path),
        "--saldo",
        str(saldo_path),
        "--complete-orders-out",
        str(complete_out),
        "--shortage-out",
        str(shortage_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["complete_order_count"] == 1
    assert payload["complete_orders"] == ["O2"]
    assert payload["shortage_article_count"] == 1
    assert payload["shortage_articles"] == ["A100"]

    assert complete_out.read_text(encoding="utf-8").splitlines() == ["O2"]
    shortage_df = pd.read_csv(shortage_out, dtype=str, encoding="utf-8-sig")
    assert shortage_df["Artikel"].tolist() == ["A100"]


def test_lyx_cli_writes_article_list(tmp_path: Path, run_cli_cmd) -> None:
    saldo_path = _write_csv(
        tmp_path / "saldo.csv",
        [
            {"Artikel": "A100", "Plocksaldo": 5, "Utbestallt": 10, "Plockplats": "P1", "Bolag": "MG"},
            {"Artikel": "A101", "Plocksaldo": 6, "Utbestallt": 16, "Plockplats": "P2", "Bolag": "MG"},
            {"Artikel": "A102", "Plocksaldo": 2, "Utbestallt": 0, "Plockplats": "", "Bolag": "MG"},
            {"Artikel": "A103", "Plocksaldo": 2, "Utbestallt": 0, "Plockplats": "P3", "Bolag": "GG"},
            {"Artikel": "A104", "Plocksaldo": 1, "Utbestallt": 0, "Plockplats": "P4", "Bolag": "MG"},
        ],
    )
    max_path = _write_csv(
        tmp_path / "artikel_max.csv",
        [
            {"artikelnummer": "A100", "max": 100},
            {"artikelnummer": "A101", "max": 100},
            {"artikelnummer": "A104", "max": 10},
        ],
    )
    out_path = tmp_path / "lyx.txt"

    completed = run_cli_cmd(
        "lyx",
        "--saldo",
        str(saldo_path),
        "--max-csv",
        str(max_path),
        "--output",
        str(out_path),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["filtered_row_count"] == 3
    assert payload["articles"] == ["A100", "A104"]
    assert out_path.read_text(encoding="utf-8").splitlines() == ["A100", "A104"]


def test_pafyllnadsprio_cli_writes_fallback_report(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders.csv",
        [
            {"Ordernr": "O1", "Artikel": "A100", "Antal": 20, "Plock": 10},
            {"Ordernr": "O1", "Artikel": "A100", "Antal": 20, "Plock": 30},
            {"Ordernr": "O2", "Artikel": "B200", "Antal": 10, "Plock": 20},
        ],
    )
    saldo_path = _write_csv(
        tmp_path / "saldo.csv",
        [
            {"Artikel": "A100", "Utbestallt": 5},
        ],
    )
    max_path = _write_csv(
        tmp_path / "artikel_max.csv",
        [
            {"artikelnummer": "A100", "max": 100},
        ],
    )
    report_out = tmp_path / "pafyllnadsprio.csv"

    completed = run_cli_cmd(
        "pafyllnadsprio",
        "--orders",
        str(orders_path),
        "--saldo",
        str(saldo_path),
        "--max-csv",
        str(max_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["mode"] == "fallback"
    assert payload["shortage_article_count"] == 1
    assert payload["missing_reference_count"] == 0

    report_df = pd.read_csv(report_out, dtype=str, encoding="utf-8-sig").fillna("")
    assert report_df["ALLA"].tolist() == ["A100"]
    assert report_df["PRIO 1"].tolist() == ["A100"]


def test_overview_check_cli_writes_shipment_and_hib_reports(tmp_path: Path, run_cli_cmd) -> None:
    overview_path = _write_csv(
        tmp_path / "overview.csv",
        [
            {"Ordernummer": "O1", "Sandningsnummer": "S1", "Kundnamn": "Butik A", "Transportor": "T1", "Ordertyp": "N", "Status": 30},
            {"Ordernummer": "O2", "Sandningsnummer": "S1", "Kundnamn": "Butik B", "Transportor": "T2", "Ordertyp": "N", "Status": 30},
            {"Ordernummer": "H100", "Sandningsnummer": "H1", "Kundnamn": "Butik C", "Transportor": "T3", "Ordertyp": "HIB", "Status": 32},
        ],
    )
    shipment_out = tmp_path / "shipment.csv"
    hib_out = tmp_path / "hib.csv"

    completed = run_cli_cmd(
        "overview-check",
        "--overview",
        str(overview_path),
        "--shipment-out",
        str(shipment_out),
        "--hib-out",
        str(hib_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "overview-check"
    assert payload["shipment_rows"] == 1
    assert payload["hib_rows"] == 1

    shipment_df = pd.read_csv(shipment_out, dtype=str, encoding="utf-8-sig")
    assert shipment_df["Sändningsnr"].tolist() == ["S1"]
    assert shipment_df["Unika kunder"].tolist() == ["2"]

    hib_df = pd.read_csv(hib_out, dtype=str, encoding="utf-8-sig")
    assert hib_df["Ordernr"].tolist() == ["H100"]
    assert hib_df["Status"].tolist() == ["32"]


def test_dispatch_check_cli_writes_mismatch_report(tmp_path: Path, run_cli_cmd) -> None:
    overview_path = _write_csv(
        tmp_path / "overview.csv",
        [
            {"Ordernummer": "O1", "Sandningsnummer": "S1", "Kundnamn": "Butik A"},
            {"Ordernummer": "O2", "Sandningsnummer": "S2", "Kundnamn": "Butik B"},
        ],
    )
    dispatch_path = _write_csv(
        tmp_path / "dispatch.csv",
        [
            {"Ordernummer": "O1", "Sandningsnummer": "S1", "Plockpallsnr": "P1"},
            {"Ordernummer": "O2", "Sandningsnummer": "WRONG", "Plockpallsnr": "P2"},
        ],
    )
    report_out = tmp_path / "dispatch_report.csv"

    completed = run_cli_cmd(
        "dispatch-check",
        "--overview",
        str(overview_path),
        "--dispatch",
        str(dispatch_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "dispatch-check"
    assert payload["mismatch_rows"] == 1

    diff_df = pd.read_csv(report_out, dtype=str, encoding="utf-8-sig")
    assert diff_df["Ordernr"].tolist() == ["O2"]
    assert diff_df["Översikt sändningsnr"].tolist() == ["S2"]
    assert diff_df["Dispatch sändningsnr"].tolist() == ["WRONG"]


def test_vecka27_check_cli_writes_deviation_report(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders.csv",
        [
            {"Ordernr": "PR100", "Artikel": "2002039", "Antal": 2},
            {"Ordernr": "PR100", "Artikel": "2003511", "Antal": 1},
        ],
    )
    report_out = tmp_path / "vecka27.txt"

    completed = run_cli_cmd(
        "vecka27-check",
        "--orders",
        str(orders_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "vecka27-check"
    assert payload["deviation_count"] == 1
    assert "PR100" in payload["deviations"][0]
    assert "Hej Lina!" in report_out.read_text(encoding="utf-8")


def test_eftersok_cli_writes_report(tmp_path: Path, run_cli_cmd) -> None:
    receive_path = _write_tsv(
        tmp_path / "receive.csv",
        [
            {
                "Ink\u00f6psnr": "PO1",
                "Artikel": "A1",
                "Pallid": "P1",
                "Mottaget": "10",
                "\u00c4ndrad": "2024-01-01 10:00",
            }
        ],
    )
    booking_path = _write_tsv(
        tmp_path / "booking.csv",
        [],
        columns=["Pall nr", "Ink\u00f6psnr", "\u00c4ndrad"],
    )
    buffert_path = _write_tsv(
        tmp_path / "buffert.csv",
        [
            {
                "Pallid": "P1",
                "Lagerplats": "UTE1",
                "Datum/tid": "2024-01-01 11:00",
            }
        ],
    )
    trans_path = _write_tsv(
        tmp_path / "trans.csv",
        [],
        columns=["Pallid", "Till", "Timestamp", "Fr\u00e5n"],
    )
    pick_path = _write_tsv(
        tmp_path / "pick.csv",
        [],
        columns=["Pallid", "Artikelnr", "Plockat", "Ordernr", "Datum"],
    )
    correct_path = _write_tsv(
        tmp_path / "correct.csv",
        [],
        columns=["Pallid", "Antal", "Anledning", "Artikel", "\u00c4ndrad"],
    )
    report_out = tmp_path / "eftersok.txt"

    completed = run_cli_cmd(
        "eftersok",
        "--purchase",
        "PO1",
        "--article",
        "A1",
        "--wms-receive",
        str(receive_path),
        "--wms-booking",
        str(booking_path),
        "--wms-buffert",
        str(buffert_path),
        "--wms-trans",
        str(trans_path),
        "--wms-pick",
        str(pick_path),
        "--wms-correct",
        str(correct_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "eftersok"
    assert payload["report_line_count"] > 0
    report_text = report_out.read_text(encoding="utf-8")
    assert "PO1" in report_text
    assert "A1" in report_text


def test_hib_koppling_cli_writes_change_and_missed_reports(tmp_path: Path, run_cli_cmd) -> None:
    details_path = _write_csv(
        tmp_path / "details.csv",
        [
            {"Order nr": "STORE1", "Status": 30, "Zon": "A", "Kund.1": "Butik A"},
            {"Order nr": "H1", "Status": 30, "Zon": "A", "Kund.1": "Butik A"},
            {"Order nr": "H2", "Status": 35, "Zon": "F", "Kund.1": "Butik A"},
        ],
    )
    overview_path = _write_csv(
        tmp_path / "overview_hib.csv",
        [
            {
                "Ordernr": "STORE1",
                "Ordertyp": "N",
                "Kund nr": "K1",
                "Orderdatum": "2024-01-01",
                "S\u00e4ndningsnr": "S1",
                "Zon": "A",
                "Multi": "",
                "Ursprungsdatum": "2024-01-01",
            },
            {
                "Ordernr": "H1",
                "Ordertyp": "HIB",
                "Kund nr": "K1",
                "Orderdatum": "2024-01-02",
                "S\u00e4ndningsnr": "WRONG",
                "Zon": "A",
                "Multi": "",
                "Ursprungsdatum": "2024-01-02",
            },
            {
                "Ordernr": "H2",
                "Ordertyp": "HIB",
                "Kund nr": "K1",
                "Orderdatum": "2024-01-03",
                "S\u00e4ndningsnr": "MISS",
                "Zon": "F",
                "Multi": "",
                "Ursprungsdatum": "2024-01-03",
            },
        ],
    )
    changes_out = tmp_path / "changes.csv"
    missed_out = tmp_path / "missed.csv"

    completed = run_cli_cmd(
        "hib-koppling",
        "--details",
        str(details_path),
        "--overview",
        str(overview_path),
        "--changes-out",
        str(changes_out),
        "--missed-out",
        str(missed_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "hib-koppling"
    assert payload["change_rows"] == 1
    assert payload["missed_rows"] == 1

    changes_df = pd.read_csv(changes_out, dtype=str, encoding="utf-8-sig")
    assert changes_df["ordernummer"].tolist() == ["H1"]
    assert changes_df["s\u00e4ndningsnummer"].tolist() == ["S1"]
    assert changes_df["Zon"].tolist() == ["F"]

    missed_df = pd.read_csv(missed_out, dtype=str, encoding="utf-8-sig")
    assert missed_df["ordernummer"].tolist() == ["H2"]
    assert missed_df["Missat"].tolist() == ["MISSAT SIN AVG\u00c5NG"]


def test_allocate_cli_writes_near_miss_refill_and_pallet_space_reports(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders_branch.csv",
        [
            {
                "Artikel": "A1",
                "Antal": 10,
                "Ordernr": "O100",
                "Radnr": "1",
                "Kund": "Butik A",
                "Kund1": "Region 1",
            },
        ],
    )
    buffer_path = _write_csv(
        tmp_path / "buffer_branch.csv",
        [
            {
                "Artikel": "A1",
                "Antal": 12,
                "Lagerplats": "H-01",
                "Datum/Tid": "2024-01-01 08:00",
                "PallID": "P1",
                "Status": 29,
                "Palltyp": "EURO",
            },
        ],
    )
    result_out = tmp_path / "allocated_branch.csv"
    near_out = tmp_path / "near_branch.csv"
    refill_hp_out = tmp_path / "refill_hp.csv"
    pallet_out = tmp_path / "pallet_spaces.csv"

    completed = run_cli_cmd(
        "allocate",
        "--orders",
        str(orders_path),
        "--buffer",
        str(buffer_path),
        "--result-out",
        str(result_out),
        "--near-miss-out",
        str(near_out),
        "--refill-hp-out",
        str(refill_hp_out),
        "--pallet-spaces-out",
        str(pallet_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["near_miss_rows"] == 1
    assert payload["refill_hp_rows"] == 1
    assert payload["pallet_space_rows"] == 1

    near_df = pd.read_csv(near_out, dtype=str, encoding="utf-8-sig")
    assert near_df["Artikel"].tolist() == ["A1"]
    assert near_df["PallID"].tolist() == ["P1"]

    refill_df = pd.read_csv(refill_hp_out, dtype=str, encoding="utf-8-sig")
    assert refill_df["Artikel"].tolist() == ["A1"]
    assert refill_df["Zon"].tolist() == ["A"]

    pallet_df = pd.read_csv(pallet_out, dtype=str, encoding="utf-8-sig")
    assert pallet_df["Kund"].tolist() == ["Butik A"]
    assert pallet_df["Pallplatser"].tolist() == ["1"]


def test_allocate_cli_counts_hib_pallet_spaces_separately_from_autostore(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders_hib.csv",
        [
            {
                "Artikel": "F1",
                "Antal": 1,
                "Ordernr": f"OF{i}",
                "Radnr": str(i),
                "Kund": "Butik F",
                "Zon": "F",
            }
            for i in range(21)
        ],
    )
    buffer_path = _write_csv(
        tmp_path / "buffer_hib.csv",
        [
            {
                "Artikel": "B1",
                "Antal": 1,
                "Lagerplats": "H-01",
                "Datum/Tid": "2024-01-01 08:00",
                "PallID": "PB1",
                "Status": 29,
                "Palltyp": "EURO",
            },
        ],
    )
    pallet_out = tmp_path / "pallet_spaces_hib.csv"

    completed = run_cli_cmd(
        "allocate",
        "--orders",
        str(orders_path),
        "--buffer",
        str(buffer_path),
        "--pallet-spaces-out",
        str(pallet_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    pallet_df = pd.read_csv(pallet_out, dtype=str, encoding="utf-8-sig")
    assert pallet_df["Kund"].tolist() == ["Butik F"]
    assert pallet_df["HIB"].tolist() == ["2"]
    assert pallet_df["autostore"].tolist() == ["0"]
    assert pallet_df["Pallplatser"].tolist() == ["2"]


def test_overview_check_cli_handles_clean_input(tmp_path: Path, run_cli_cmd) -> None:
    overview_path = _write_csv(
        tmp_path / "overview_clean.csv",
        [
            {
                "Ordernummer": "O1",
                "Sandningsnummer": "S1",
                "Kundnamn": "Butik A",
                "Transportor": "T1",
                "Ordertyp": "N",
                "Status": 30,
            },
            {
                "Ordernummer": "H1",
                "Sandningsnummer": "S1",
                "Kundnamn": "Butik A",
                "Transportor": "T1",
                "Ordertyp": "HIB",
                "Status": 31,
            },
        ],
    )
    report_out = tmp_path / "overview_clean_report.csv"

    completed = run_cli_cmd(
        "overview-check",
        "--overview",
        str(overview_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["shipment_rows"] == 0
    assert payload["hib_rows"] == 0
    assert "Avvikelsetyp" in report_out.read_text(encoding="utf-8-sig")


def test_dispatch_check_cli_handles_clean_input(tmp_path: Path, run_cli_cmd) -> None:
    overview_path = _write_csv(
        tmp_path / "overview_dispatch_clean.csv",
        [
            {"Ordernummer": "O1", "Sandningsnummer": "S1", "Kundnamn": "Butik A"},
        ],
    )
    dispatch_path = _write_csv(
        tmp_path / "dispatch_clean.csv",
        [
            {"Ordernummer": "O1", "Sandningsnummer": "S1", "Plockpallsnr": "P1"},
        ],
    )
    report_out = tmp_path / "dispatch_clean_report.csv"

    completed = run_cli_cmd(
        "dispatch-check",
        "--overview",
        str(overview_path),
        "--dispatch",
        str(dispatch_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["mismatch_rows"] == 0
    assert report_out.exists()


def test_vecka27_check_cli_handles_no_deviations(tmp_path: Path, run_cli_cmd) -> None:
    orders_path = _write_csv(
        tmp_path / "orders_ok.csv",
        [
            {"Ordernr": "PR100", "Artikel": "2002039", "Antal": 1},
            {"Ordernr": "PR100", "Artikel": "2003511", "Antal": 1},
        ],
    )
    report_out = tmp_path / "vecka27_ok.txt"

    completed = run_cli_cmd(
        "vecka27-check",
        "--orders",
        str(orders_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["deviation_count"] == 0
    assert "inga avvikelser" in report_out.read_text(encoding="utf-8").lower()


def test_eftersok_cli_works_with_only_required_receive_file(tmp_path: Path, run_cli_cmd) -> None:
    receive_path = _write_tsv(
        tmp_path / "receive_only.csv",
        [
            {
                "Ink\u00f6psnr": "PO2",
                "Artikel": "A2",
                "Pallid": "P2",
                "Mottaget": "5",
                "\u00c4ndrad": "2024-01-01 12:00",
            }
        ],
    )

    completed = run_cli_cmd(
        "eftersok",
        "--purchase",
        "PO2",
        "--article",
        "A2",
        "--wms-receive",
        str(receive_path),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "eftersok"
    assert payload["report_line_count"] > 0


def test_prognos_report_cli_writes_report_and_combined_inputs(tmp_path: Path, run_cli_cmd) -> None:
    prognos_path = _write_csv(
        tmp_path / "prognos.csv",
        [
            {
                "Artikelnummer": "A1",
                "Beskrivning": "Artikel 1",
                "Antal styck": 10,
                "Antal rader": 1,
                "Antal butiker": 1,
            }
        ],
    )
    saldo_path = _write_csv(
        tmp_path / "saldo_prognos.csv",
        [
            {"Artikel": "A1", "Robot": "Y", "Saldo autoplock": 3},
        ],
    )
    buffer_path = _write_csv(
        tmp_path / "buffer_prognos.csv",
        [
            {"Artikel": "A1", "Antal": 4, "Lagerplats": "B1", "Datum/Tid": "2024-01-01 08:00", "PallID": "P1", "Status": 29},
            {"Artikel": "A1", "Antal": 5, "Lagerplats": "B2", "Datum/Tid": "2024-01-02 08:00", "PallID": "P2", "Status": 30},
        ],
    )
    report_out = tmp_path / "prognos_report.csv"
    combined_out = tmp_path / "prognos_combined.csv"

    completed = run_cli_cmd(
        "prognos-report",
        "--prognos",
        str(prognos_path),
        "--saldo",
        str(saldo_path),
        "--buffer",
        str(buffer_path),
        "--report-out",
        str(report_out),
        "--combined-out",
        str(combined_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "prognos-report"
    assert payload["combined_rows"] == 1
    assert payload["report_rows"] == 1
    assert payload["partial"] is False

    report_df = pd.read_csv(report_out, dtype=str, encoding="utf-8-sig")
    assert report_df["Artikelnummer"].tolist() == ["A1"]
    assert report_df["Behov efter saldo"].tolist() == ["7.0"]
    assert report_df["FIFO-baserad ber\u00e4kning (antal pall)"].tolist() == ["2.0"]

    combined_df = pd.read_csv(combined_out, dtype=str, encoding="utf-8-sig")
    assert combined_df["Artikelnummer"].tolist() == ["A1"]


def test_prognos_report_cli_requires_saldo(tmp_path: Path, run_cli_cmd) -> None:
    prognos_path = _write_csv(
        tmp_path / "prognos_partial.csv",
        [
            {
                "Artikelnummer": "A1",
                "Beskrivning": "Artikel 1",
                "Antal styck": 10,
                "Antal rader": 1,
                "Antal butiker": 1,
            }
        ],
    )

    completed = run_cli_cmd(
        "prognos-report",
        "--prognos",
        str(prognos_path),
        "--json",
    )

    assert completed.returncode != 0
    assert "Ange --saldo" in completed.stderr


def test_prognos_report_cli_keeps_rows_without_buffer(tmp_path: Path, run_cli_cmd) -> None:
    prognos_path = _write_csv(
        tmp_path / "prognos_partial.csv",
        [
            {
                "Artikelnummer": "A1",
                "Beskrivning": "Artikel 1",
                "Antal styck": 10,
                "Antal rader": 1,
                "Antal butiker": 1,
            }
        ],
    )
    saldo_path = _write_csv(
        tmp_path / "saldo_partial.csv",
        [
            {"Artikel": "A1", "Robot": "Y", "Saldo autoplock": 3},
        ],
    )
    report_out = tmp_path / "prognos_partial_report.csv"

    completed = run_cli_cmd(
        "prognos-report",
        "--prognos",
        str(prognos_path),
        "--saldo",
        str(saldo_path),
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "prognos-report"
    assert payload["combined_rows"] == 1
    assert payload["report_rows"] == 1
    assert payload["partial"] is True
    assert payload["missing"] == "buffert"

    report_df = pd.read_csv(report_out, dtype=str, encoding="utf-8-sig")
    assert report_df["Artikelnummer"].tolist() == ["A1"]
    assert report_df["Saldo i autoplock"].tolist() == ["3.0"]
    assert report_df["Behov efter saldo"].tolist() == ["7.0"]


def test_observations_update_cli_writes_new_rows_and_article_max(tmp_path: Path, run_cli_cmd) -> None:
    buffer_path = _write_csv(
        tmp_path / "buffer_obs.csv",
        [
            {"Artikel": "A1", "Antal": 10, "PallID": "P1", "Status": 30},
            {"Artikel": "A2", "Antal": 4, "PallID": "P2", "Status": 29},
        ],
    )
    observations_path = tmp_path / "observations.csv.gz"
    article_max_path = tmp_path / "artikel_max.csv"
    new_out = tmp_path / "new_rows.csv"

    completed = run_cli_cmd(
        "observations-update",
        "--buffer",
        str(buffer_path),
        "--observations-path",
        str(observations_path),
        "--article-max-out",
        str(article_max_path),
        "--new-out",
        str(new_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "observations-update"
    assert payload["new_rows"] == 1
    assert payload["article_max_rows"] == 1

    new_df = pd.read_csv(new_out, dtype=str, encoding="utf-8-sig")
    assert new_df["pallid"].tolist() == ["P1"]

    obs_df = pd.read_csv(observations_path, compression="gzip", dtype=str)
    assert obs_df["pallid"].tolist() == ["P1"]

    max_df = pd.read_csv(article_max_path, dtype=str, encoding="utf-8-sig")
    assert max_df["artikelnummer"].tolist() == ["A1"]
    assert max_df["pallid"].tolist() == ["P1"]


def test_observations_sync_cli_reads_local_remote_file(tmp_path: Path, run_cli_cmd) -> None:
    observations_path = _write_gzip_csv(
        tmp_path / "local_obs.csv.gz",
        [
            {"artikelnummer": "A1", "pallid": "P1", "antal": "10"},
        ],
    )
    remote_path = _write_gzip_csv(
        tmp_path / "remote_obs.csv.gz",
        [
            {"artikelnummer": "A1", "pallid": "P1", "antal": "10"},
            {"artikelnummer": "A2", "pallid": "P2", "antal": "7"},
        ],
    )
    article_max_path = tmp_path / "artikel_max_sync.csv"

    completed = run_cli_cmd(
        "observations-sync",
        "--observations-path",
        str(observations_path),
        "--article-max-out",
        str(article_max_path),
        "--remote-file",
        str(remote_path),
        "--no-push",
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "observations-sync"
    assert payload["fetched_rows"] == 1
    assert payload["pushed_rows"] == 0
    assert payload["total_observations"] == 2

    obs_df = pd.read_csv(observations_path, compression="gzip", dtype=str)
    assert sorted(obs_df["pallid"].tolist()) == ["P1", "P2"]

    max_df = pd.read_csv(article_max_path, dtype=str, encoding="utf-8-sig")
    assert sorted(max_df["artikelnummer"].tolist()) == ["A1", "A2"]


def test_split_values_cli_writes_chunked_columns(tmp_path: Path, run_cli_cmd) -> None:
    input_path = _write_text(tmp_path / "values.txt", "A\nB\nC\nD\nE\n")
    report_out = tmp_path / "split.csv"

    completed = run_cli_cmd(
        "split-values",
        "--input",
        str(input_path),
        "--chunk-size",
        "2",
        "--report-out",
        str(report_out),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "split-values"
    assert payload["value_count"] == 5
    assert payload["chunk_count"] == 3

    split_df = pd.read_csv(report_out, dtype=str, encoding="utf-8-sig").fillna("")
    assert split_df.columns.tolist() == ["Kolumn 1", "Kolumn 2", "Kolumn 3"]
    assert split_df.iloc[0].tolist() == ["A", "C", "E"]
    assert split_df.iloc[1].tolist() == ["B", "D", ""]


def test_update_check_cli_reads_local_release_payload(tmp_path: Path, run_cli_cmd) -> None:
    release_path = _write_text(
        tmp_path / "release.json",
        json.dumps(
            {
                "tag_name": "v12.9.0",
                "html_url": "https://example.test/releases/v12.9.0",
                "assets": [
                    {
                        "name": "Allokering-12.9.0-Setup.exe",
                        "browser_download_url": "https://example.test/Allokering-12.9.0-Setup.exe",
                    }
                ],
            }
        ),
    )

    completed = run_cli_cmd(
        "update-check",
        "--release-json",
        str(release_path),
        "--json",
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout.strip())
    assert payload["command"] == "update-check"
    assert payload["has_update"] is True
    assert payload["latest_version"] == "12.9.0"
    assert payload["installer_name"] == "Allokering-12.9.0-Setup.exe"
