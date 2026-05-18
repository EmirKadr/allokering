"""Snabbt end-to-end-test av API:t med FastAPI TestClient (ingen server behovs)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi.testclient import TestClient

import api

SAMPLE = Path(__file__).resolve().parents[1] / "sample_data"
client = TestClient(api.app)


def main() -> int:
    h = client.get("/api/health").json()
    print("health:", h)

    with open(SAMPLE / "bestallningslinjer.csv", "rb") as f:
        det = client.post("/api/detect", files={"file": ("bestallningslinjer.csv", f, "text/csv")})
    print("detect orders:", det.json())

    with open(SAMPLE / "bestallningslinjer.csv", "rb") as o, open(SAMPLE / "buffertpallar.csv", "rb") as b:
        res = client.post(
            "/api/allocate",
            files={
                "orders": ("bestallningslinjer.csv", o, "text/csv"),
                "buffer": ("buffertpallar.csv", b, "text/csv"),
            },
        )
    if res.status_code != 200:
        print("ALLOCATE FAILED", res.status_code, res.text)
        return 1
    data = res.json()
    print("summary:", data["summary"])
    print("result columns:", data["tables"]["result"]["columns"])
    print("result rows:", data["tables"]["result"]["row_count"])
    print("log lines:", len(data["log"]))

    dl = client.get(f"/api/download/{data['session_id']}/result")
    print("download status:", dl.status_code, "bytes:", len(dl.content))
    print("OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
