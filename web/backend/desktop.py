"""Desktop-launcher: startar API:t och visar React-frontenden i ett pywebview-fonster.

Kör:  python web/backend/desktop.py

Detta ger en "riktig app"-känsla men under huven är allt samma HTTP-API
som senare kan deployas som webbapp.
"""
from __future__ import annotations

import sys
import threading
import time
import urllib.request
from pathlib import Path

# Så att "import api" / "import engine" fungerar oavsett varifrån scriptet körs.
sys.path.insert(0, str(Path(__file__).resolve().parent))

HOST = "127.0.0.1"
PORT = 8765
URL = f"http://{HOST}:{PORT}"
APP_ICON = Path(__file__).resolve().parents[1] / "frontend" / "public" / "app.ico"


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
        print("Frontenden är inte byggd än. Kör först:")
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
        "Allokering",
        URL,
        width=1440,
        height=920,
        min_size=(1100, 720),
    )
    # http_server=False: vi serverar allt själva via FastAPI.
    webview.start(icon=str(APP_ICON) if APP_ICON.exists() else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
