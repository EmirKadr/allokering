from __future__ import annotations

import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

import psycopg
from fastapi.testclient import TestClient
from pgembed import get_server


ROOT = Path(__file__).resolve().parents[1]
TESTDATA = Path(os.environ.get("ALLOK_TESTDATA_DIR", str(Path(r"C:\allokering\testdata")))).resolve()


def _upload_file(client: TestClient, headers: dict, sid: str, file_key: str, path: Path, page_id: str) -> dict:
    with path.open("rb") as fh:
        resp = client.post(
            f"/api/upload/{sid}",
            headers=headers,
            data={"file_key": file_key, "page_id": page_id},
            files={"file": (path.name, fh, "application/octet-stream")},
        )
    if resp.status_code != 200:
        raise RuntimeError(f"Upload av {file_key} misslyckades: {resp.status_code} {resp.text[:500]}")
    return resp.json()


def _process_next_job(worker_id: str = "sandbox-smoke-worker") -> dict:
    from backend.job_queue import claim_next_job
    from backend.worker import _process_job

    job = claim_next_job(worker_id)
    if not job:
        raise RuntimeError("Hittade inget köat jobb att processa")
    asyncio.run(_process_job(job, worker_id))
    return job


def _assert_download(client: TestClient, headers: dict, sid: str, result_key: str) -> None:
    resp = client.get(f"/api/download/{sid}/{result_key}", headers=headers)
    if resp.status_code != 200:
        raise RuntimeError(f"Nedladdning av {result_key} misslyckades: {resp.status_code} {resp.text[:200]}")


def _run_child_restart_check(env: dict, token: str, sid: str, ef_sid: str) -> dict:
    child_env = env.copy()
    child_env["SMOKE_TOKEN"] = token
    child_env["SMOKE_SESSION_ID"] = sid
    child_env["SMOKE_EFTERSOK_SESSION_ID"] = ef_sid
    child_env["SMOKE_WEBAPP_ROOT"] = str(ROOT)
    child_code = r"""
import json
import os
import sys
from fastapi.testclient import TestClient

sys.path.insert(0, os.environ["SMOKE_WEBAPP_ROOT"])
import backend.main as main

with TestClient(main.app) as client:
    headers = {"Authorization": f"Bearer {os.environ['SMOKE_TOKEN']}"}
    payload = {
        "me": client.get("/api/auth/me", headers=headers).json(),
        "files": client.get(f"/api/upload/{os.environ['SMOKE_SESSION_ID']}", headers=headers).json(),
        "status": client.get(f"/api/run/status/{os.environ['SMOKE_SESSION_ID']}", headers=headers).json(),
        "ef_status": client.get(f"/api/run/status/{os.environ['SMOKE_EFTERSOK_SESSION_ID']}", headers=headers).json(),
    }
    print(json.dumps(payload, ensure_ascii=False))
"""
    completed = subprocess.run(
        [sys.executable, "-c", child_code],
        capture_output=True,
        text=True,
        env=child_env,
        check=True,
    )
    return json.loads(completed.stdout.strip().splitlines()[-1])


def main() -> None:
    if not TESTDATA.exists():
        raise SystemExit(f"Testdatamappen finns inte: {TESTDATA}")

    sys.path.insert(0, str(ROOT))

    results: dict = {}
    with TemporaryDirectory() as td:
        td_path = Path(td)
        pgdata = td_path / "pgdata"
        state_dir = td_path / "state"
        server = get_server(pgdata, cleanup_mode="delete")
        try:
            with psycopg.connect(server.get_uri("postgres"), autocommit=True) as conn:
                with conn.cursor() as cur:
                    cur.execute("CREATE DATABASE alloktest")

            os.environ["ALLOK_DATABASE_URL"] = server.get_uri("alloktest")
            os.environ["ALLOK_STATE_DIR"] = str(state_dir)

            import backend.main as main_mod

            with TestClient(main_mod.app) as client:
                login_resp = client.post("/api/auth/login", json={"identifier": "emikad", "code": "0012"})
                login_resp.raise_for_status()
                token = login_resp.json()["token"]
                headers = {"Authorization": f"Bearer {token}"}

                me_resp = client.get("/api/auth/me", headers=headers)
                me_resp.raise_for_status()
                results["auth_me"] = me_resp.json()

                session_resp = client.post("/api/session", headers=headers)
                session_resp.raise_for_status()
                sid = session_resp.json()["session_id"]
                results["allokering_session_id"] = sid

                allok_uploads = {
                    "orders": TESTDATA / "v_ask_customer_order_details_all-20260317145125.csv",
                    "buffer": TESTDATA / "v_ask_article_buffertpallet-20260317145136.csv",
                    "automation": TESTDATA / "v_ask_item_summary_stock_automation-20260317145351.csv",
                    "item": TESTDATA / "item_option-20260317145203.csv",
                    "overview": TESTDATA / "v_ask_order_overview-20260317145114.csv",
                    "dispatch": TESTDATA / "v_ask_dispatch_pallet-20260316130458.csv",
                }
                results["allokering_uploads"] = {
                    key: _upload_file(client, headers, sid, key, path, "allokering-page")
                    for key, path in allok_uploads.items()
                }
                results["allokering_file_list_before"] = client.get(f"/api/upload/{sid}", headers=headers).json()

                run_resp = client.post(f"/api/run/allokering/{sid}", headers=headers)
                if run_resp.status_code != 202:
                    raise RuntimeError(f"Allokering kunde inte köas: {run_resp.status_code} {run_resp.text}")
                results["allokering_queue_response"] = run_resp.json()
                results["allokering_job_claimed"] = _process_next_job()["job_type"]
                results["allokering_status"] = client.get(f"/api/run/status/{sid}", headers=headers).json()
                for result_key in ("allokerade", "nearmiss", "pallplatser", "refill"):
                    if result_key in results["allokering_status"]["results"]:
                        _assert_download(client, headers, sid, result_key)

                ord1 = client.get(f"/api/ordersaldo/list1/{sid}", headers=headers).json()
                ord2 = client.get(f"/api/ordersaldo/list2/{sid}", headers=headers).json()
                results["ordersaldo_counts"] = {
                    "list1": len(ord1.get("values", [])),
                    "list2": len(ord2.get("values", [])),
                }

                for job_name, expected_key in (
                    ("hib-koppling", "hib-koppling"),
                    ("orderkontroll", "orderkontroll"),
                    ("dispatchkontroll", "dispatchkontroll"),
                ):
                    resp = client.post(f"/api/run/{job_name}/{sid}", headers=headers)
                    if resp.status_code != 202:
                        raise RuntimeError(f"{job_name} kunde inte köas: {resp.status_code} {resp.text}")
                    _process_next_job()
                    status_json = client.get(f"/api/run/status/{sid}", headers=headers).json()
                    results[f"{job_name}_status"] = status_json
                    if expected_key not in status_json["results"]:
                        raise RuntimeError(f"{job_name} saknar resultatnyckeln {expected_key}")
                    _assert_download(client, headers, sid, expected_key)

                ef_session_resp = client.post("/api/session", headers=headers)
                ef_session_resp.raise_for_status()
                ef_sid = ef_session_resp.json()["session_id"]
                results["eftersok_session_id"] = ef_sid

                ef_uploads = {
                    "wms_receive": TESTDATA / "v_ask_receive_log-20260317145157.csv",
                    "wms_buffer": TESTDATA / "v_ask_article_buffertpallet-20260317145136.csv",
                    "wms_booking": TESTDATA / "v_ask_booking_putaway-20260317145232.csv",
                    "wms_trans": TESTDATA / "v_ask_trans_log-20260317170854.csv",
                    "wms_pick": TESTDATA / "v_ask_pick_log_full-20260317170910.csv",
                    "wms_correct": TESTDATA / "v_ask_correct_log-20260317145302.csv",
                }
                for key, path in ef_uploads.items():
                    _upload_file(client, headers, ef_sid, key, path, "eftersok-page")

                ef_resp = client.post(
                    f"/api/run/eftersok/{ef_sid}",
                    headers=headers,
                    json={"purchase": "999109661", "article": "200870001"},
                )
                if ef_resp.status_code != 202:
                    raise RuntimeError(f"Eftersök kunde inte köas: {ef_resp.status_code} {ef_resp.text}")
                _process_next_job()
                results["eftersok_status"] = client.get(f"/api/run/status/{ef_sid}", headers=headers).json()
                if "eftersok" not in results["eftersok_status"]["results"]:
                    raise RuntimeError("Eftersök-resultatet skapades inte")
                _assert_download(client, headers, ef_sid, "eftersok")

                results["restart_check"] = _run_child_restart_check(os.environ.copy(), token, sid, ef_sid)

        finally:
            server.cleanup()

    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
