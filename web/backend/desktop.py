"""Desktop-launcher: startar API:t och visar React-frontenden i ett pywebview-fonster.

Kor:  python web/backend/desktop.py

Detta ger en "riktig app"-kansla men under huven ar allt samma HTTP-API
som senare kan deployas som webbapp.
"""
from __future__ import annotations

import sys
import threading
import time
import urllib.request
from pathlib import Path

# Sa att "import api" / "import engine" fungerar oavsett varifran scriptet kors.
sys.path.insert(0, str(Path(__file__).resolve().parent))

HOST = "127.0.0.1"
PORT = 8765
URL = f"http://{HOST}:{PORT}"


def _run_server() -> None:
    import uvicorn

    import api

    uvicorn.run(api.app, host=HOST, port=PORT, log_level="warning")


def _wait_for_server(timeout: float = 20.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(f"{URL}/api/health", timeout=1.0)
            return True
        except Exception:
            time.sleep(0.25)
    return False


def main() -> int:
    dist = Path(__file__).resolve().parents[1] / "frontend" / "dist"
    if not dist.exists():
        print("Frontenden ar inte byggd an. Kor forst:")
        print("  cd web/frontend && npm install && npm run build")
        return 1

    threading.Thread(target=_run_server, daemon=True).start()
    if not _wait_for_server():
        print("API-servern startade inte i tid.")
        return 1

    import webview

    try:
        version = webview.__version__  # noqa: F841
    except Exception:
        pass

    webview.create_window(
        "Allokering - Demo",
        URL,
        width=1440,
        height=920,
        min_size=(1100, 720),
    )
    # http_server=False: vi serverar allt sjalva via FastAPI.
    webview.start()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
