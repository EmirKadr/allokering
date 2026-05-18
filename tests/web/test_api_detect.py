from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

from fastapi.testclient import TestClient
import pandas as pd
import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BACKEND = PROJECT_ROOT / "web" / "backend"

if str(BACKEND) not in sys.path:
    sys.path.insert(0, str(BACKEND))

import api  # noqa: E402
import flows  # noqa: E402


@pytest.mark.parametrize(
    ("filename", "expected_type"),
    [
        ("v_ask_receive_log-20260518075529.csv", "wms_receive"),
        ("v_ask_booking_putaway-20260518075529.csv", "wms_booking"),
        ("v_ask_article_buffertpallet-20260518075529.csv", "buffer"),
        ("v_ask_trans_log-20260518075529.csv", "wms_trans"),
        ("v_ask_pick_log_full-20260518075529.csv", "wms_pick"),
        ("v_ask_correct_log-20260518075529.csv", "wms_correct"),
        ("buffertpallar-20260518075529.csv", "buffer"),
    ],
)
def test_detect_preserves_uploaded_filename_hints(filename, expected_type):
    client = TestClient(api.app)

    response = client.post(
        "/api/detect",
        files={
            "file": (
                filename,
                b"dummy\n",
                "text/csv",
            )
        },
    )

    assert response.status_code == 200
    assert response.json()["file_type"] == expected_type


def test_wms_pick_routes_only_to_eftersok_plocklogg_input():
    registry = flows.public_registry()
    eftersok = next(flow for flow in registry if flow["id"] == "eftersok")
    eftersok_match = [
        inp["key"]
        for inp in eftersok["inputs"]
        if "wms_pick" in inp.get("detect", [])
    ]

    combined_matches = [
        (flow["id"], inp["key"])
        for flow in registry
        if flow["view"] == "combined"
        for inp in flow["inputs"]
        if "wms_pick" in inp.get("detect", [])
    ]
    orders_matches = [
        (flow["id"], inp["key"])
        for flow in registry
        for inp in flow["inputs"]
        if inp["key"] in {"orders", "details"} and "wms_pick" in inp.get("detect", [])
    ]

    assert eftersok_match == ["wms_pick"]
    assert combined_matches == []
    assert orders_matches == []


def test_observations_update_endpoint_uses_uploaded_buffer(monkeypatch):
    client = TestClient(api.app)
    calls = {}

    def fake_read_table(path: str):
        calls["path"] = Path(path)
        return object()

    def fake_build_observations_update_result(buffer_df, push_to_github=False):
        calls["buffer_df"] = buffer_df
        calls["push_to_github"] = push_to_github
        return SimpleNamespace(
            new_row_count=3,
            article_max_rows=10,
            pushed_to_github=False,
            observations_path="observations.csv.gz",
            article_max_path="artikel_max.csv",
        )

    monkeypatch.setattr(api.engine, "read_table", fake_read_table)
    monkeypatch.setattr(
        api.engine,
        "build_observations_update_result",
        fake_build_observations_update_result,
    )

    response = client.post(
        "/api/observations/update",
        files={"file": ("buffertpallar.csv", b"Artikel,Antal\n1,2\n", "text/csv")},
    )

    assert response.status_code == 200
    assert response.json()["new_rows"] == 3
    assert calls["path"].name.startswith("allok_upload_buffertpallar_")
    assert calls["push_to_github"] is True
    assert not calls["path"].exists()


def test_table_column_endpoint_returns_full_split_values_column():
    client = TestClient(api.app)
    values = [f"V{i}" for i in range(1005)]

    response = client.post(
        "/api/flow/split-values",
        data={"values": "\n".join(values), "chunk_size": "1005"},
    )

    assert response.status_code == 200
    data = response.json()
    table = data["tables"][0]["table"]
    assert table["truncated"] is True
    assert len(table["rows"]) == 1000

    column_response = client.get(f"/api/table-column/{data['session_id']}/report/0")

    assert column_response.status_code == 200
    assert column_response.json()["text"].splitlines() == values


def test_open_excel_for_split_values_uses_headerless_workbook(monkeypatch):
    client = TestClient(api.app)
    captured = {}

    def fake_open_df_in_excel_without_header(df, label):
        captured["df"] = df
        captured["label"] = label
        return f"tmp_{label}.xlsx"

    monkeypatch.setattr(
        api,
        "_open_df_in_excel_without_header",
        fake_open_df_in_excel_without_header,
    )

    response = client.post(
        "/api/flow/split-values",
        data={"values": "A\nB\nC\nD", "chunk_size": "2"},
    )
    data = response.json()

    excel_response = client.post(
        "/api/open-excel",
        json={"session_id": data["session_id"], "key": "report"},
    )

    assert excel_response.status_code == 200
    assert excel_response.json()["path"].endswith("_Delade värden.xlsx")
    assert captured["label"] == "Delade värden"
    assert captured["df"].columns.tolist() == ["Kolumn 1", "Kolumn 2"]


def test_headerless_excel_writer_starts_with_values(monkeypatch):
    opened = []
    monkeypatch.setattr(api, "_open_path", lambda path: opened.append(path))
    df = pd.DataFrame({"Kolumn 1": ["A", "B"], "Kolumn 2": ["C", "D"]})

    path = api._open_df_in_excel_without_header(df, label="Delade värden")

    try:
        written = pd.read_excel(path, header=None, dtype=str)
        assert written.values.tolist() == [["A", "C"], ["B", "D"]]
        assert opened == [path]
    finally:
        Path(path).unlink(missing_ok=True)
