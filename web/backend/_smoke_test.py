"""End-to-end-test av API:t med FastAPI TestClient (ingen server behovs)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi.testclient import TestClient

import api

SAMPLE = Path(__file__).resolve().parents[1] / "sample_data"
client = TestClient(api.app)


def _f(name: str):
    return open(SAMPLE / name, "rb")


def main() -> int:
    print("health:", client.get("/api/health").json()["status"])

    reg = client.get("/api/flows").json()["flows"]
    print("flows registrerade:", len(reg))

    with _f("bestallningslinjer.csv") as o:
        print("detect:", client.post("/api/detect", files={"file": ("orders.csv", o, "text/csv")}).json())

    # allocate
    with _f("bestallningslinjer.csv") as o, _f("buffertpallar.csv") as b:
        res = client.post("/api/flow/allocate", files={
            "orders": ("orders.csv", o, "text/csv"),
            "buffer": ("buffer.csv", b, "text/csv"),
        })
    if res.status_code != 200:
        print("ALLOCATE FAILED", res.status_code, res.text)
        return 1
    data = res.json()
    print("allocate summary:", data["summary"])
    print("allocate tabeller:", [t["key"] for t in data["tables"]])

    # ordersaldo
    with _f("bestallningslinjer.csv") as o:
        res = client.post("/api/flow/ordersaldo", files={"orders": ("orders.csv", o, "text/csv")})
    print("ordersaldo:", res.status_code, res.json().get("summary") if res.status_code == 200 else res.text)

    # vecka27
    with _f("bestallningslinjer.csv") as o:
        res = client.post("/api/flow/vecka27-check", files={"orders": ("orders.csv", o, "text/csv")})
    print("vecka27:", res.status_code, res.json().get("summary") if res.status_code == 200 else res.text)

    # split-values (textfalt)
    res = client.post("/api/flow/split-values", data={"values": "A\nB\nC\nD", "chunk_size": "2"})
    print("split-values:", res.status_code, res.json().get("summary") if res.status_code == 200 else res.text)

    # excel-export + download
    excel = client.post("/api/open-excel", json={"session_id": data["session_id"], "key": "result"})
    print("open-excel:", excel.status_code)
    dl = client.get(f"/api/download/{data['session_id']}/result")
    print("download:", dl.status_code, len(dl.content), "bytes")
    print("OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
