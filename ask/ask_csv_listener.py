#!/usr/bin/env python3
"""
Always-on CSV listener for ASK UI.

This service keeps a browser session alive and listens for requests:
- /init-login: open URL and wait for manual login/2FA
- /fetch-csv: open view, export CSV, and return the file

Run:
  uvicorn ask_csv_listener:app --host 127.0.0.1 --port 8010
"""

from __future__ import annotations

import os
import re
import tempfile
import threading
import time
from urllib.parse import parse_qs, unquote, urlencode, urlparse
from pathlib import Path
from typing import Literal, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import Locator, Page, sync_playwright
from starlette.background import BackgroundTask

DEFAULT_ASK_URL = os.environ.get(
    "ASK_CSV_DEFAULT_URL",
    "https://noeffectui-frey.nowastelogistics.com/desktop",
).strip()


def _click_by_text(page: Page, text: str, timeout_ms: int = 4000) -> bool:
    candidates = [
        page.get_by_role("tab", name=text),
        page.get_by_role("button", name=text),
        page.get_by_role("menuitem", name=text),
        page.get_by_text(text, exact=True),
        page.get_by_text(text),
    ]
    for loc in candidates:
        try:
            loc.first.wait_for(state="visible", timeout=timeout_ms)
            loc.first.click()
            return True
        except Exception:
            continue
    return False


def _first_visible(
    page: Page,
    selectors: list[str],
    timeout_ms: int = 15000,
    max_per_selector: int = 25,
):
    step_ms = 200
    attempts = max(1, timeout_ms // step_ms)
    for _ in range(attempts):
        for sel in selectors:
            try:
                loc = page.locator(sel)
                count = min(loc.count(), max_per_selector)
                for i in range(count):
                    item = loc.nth(i)
                    if item.is_visible():
                        return item
            except Exception:
                continue
        page.wait_for_timeout(step_ms)
    return None


def _grid_box_score(bbox: dict) -> tuple[float, float]:
    area = float(bbox["width"] * bbox["height"])
    top = float(bbox["y"])
    return (top, -area)


def _right_click_in_primary_grid(page: Page, timeout_ms: int) -> bool:
    viewport = page.viewport_size or {"width": 1920, "height": 1080}
    vheight = float(viewport.get("height", 1080))
    upper_limit = vheight * 0.72
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    # First try visible rows in the upper/main grid.
    row = _first_visible(
        page,
        selectors=[
            "tr.x-grid-row",
            ".x-grid-row",
            ".x-grid-item",
            ".ag-row",
            "[role='row']",
            "table tbody tr",
        ],
        timeout_ms=timeout_ms,
        max_per_selector=120,
    )
    if row:
        try:
            bbox = row.bounding_box()
            if bbox:
                center_y = bbox["y"] + (bbox["height"] / 2.0)
                if 90 <= center_y <= upper_limit and bbox["width"] >= 250:
                    x = bbox["x"] + min(220.0, max(70.0, bbox["width"] * 0.2))
                    y = center_y
                    page.mouse.click(x, y, button="left")
                    page.wait_for_timeout(90)
                    page.mouse.click(x, y, button="right")
                    return True
        except Exception:
            pass

    # Fallback to largest/upper grid container.
    candidates: list[tuple[Locator, dict]] = []
    for sel in [
        ".x-grid-view",
        ".x-grid-body",
        ".x-grid-inner-normal",
        ".x-panel-body",
        ".ag-root",
        "[role='grid']",
        "table",
    ]:
        try:
            loc = page.locator(sel)
            count = min(loc.count(), 40)
            for i in range(count):
                item = loc.nth(i)
                if not item.is_visible():
                    continue
                bbox = item.bounding_box()
                if not bbox:
                    continue
                if bbox["width"] < 450 or bbox["height"] < 120:
                    continue
                if bbox["y"] < 80 or bbox["y"] > upper_limit:
                    continue
                candidates.append((item, bbox))
        except Exception:
            continue

    if not candidates:
        return False

    _, best_box = sorted(candidates, key=lambda it: _grid_box_score(it[1]))[0]
    x = best_box["x"] + min(260.0, max(90.0, best_box["width"] * 0.18))
    y = best_box["y"] + min(180.0, max(55.0, best_box["height"] * 0.3))
    page.mouse.click(x, y, button="left")
    page.wait_for_timeout(90)
    page.mouse.click(x, y, button="right")
    return True


def _open_view_from_menu(page: Page, view_name: str, use_shortcut: bool = True) -> bool:
    view_name = (view_name or "").strip()
    if not view_name:
        return False

    if use_shortcut:
        try:
            page.keyboard.press("v")
            page.wait_for_timeout(120)
            page.keyboard.press("o")
            page.wait_for_timeout(250)
        except Exception:
            pass

    search = _first_visible(
        page,
        selectors=[
            "input[placeholder*='SÃ¶k']",
            "input[placeholder*='sok' i]",
            "input[placeholder*='search' i]",
            "input[type='search']",
        ],
        timeout_ms=1800,
    )

    if not search:
        plus = _first_visible(
            page,
            selectors=[
                "button:has-text('+')",
                "[role='button']:has-text('+')",
                "button[aria-label*='plus' i]",
                "button[title*='plus' i]",
            ],
            timeout_ms=5000,
        )
        if not plus:
            return False
        plus.click()

        search = _first_visible(
            page,
            selectors=[
                "input[placeholder*='SÃ¶k']",
                "input[placeholder*='sok' i]",
                "input[placeholder*='search' i]",
                "input[type='search']",
                "input[type='text']",
            ],
            timeout_ms=7000,
        )
        if not search:
            return False

    try:
        search.click()
        search.fill(view_name)
        page.wait_for_timeout(350)
    except Exception:
        return False

    if _click_by_text(page, view_name, timeout_ms=5000):
        return True

    try:
        search.press("Enter")
        return True
    except Exception:
        return False


def _safe_filename(name: str, fallback: str = "export.csv") -> str:
    s = re.sub(r'[\\/:*?"<>|]+', "_", (name or "").strip())
    if not s:
        s = fallback
    if not s.lower().endswith(".csv"):
        s += ".csv"
    return s


def _dedupe_keep_order(values: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for v in values:
        vv = (v or "").strip()
        if not vv or vv in seen:
            continue
        seen.add(vv)
        out.append(vv)
    return out


def _filename_from_content_disposition(header_value: str) -> str:
    if not header_value:
        return ""
    m = re.search(r"filename\*=UTF-8''([^;]+)", header_value, flags=re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    m = re.search(r'filename="([^"]+)"', header_value, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"filename=([^;]+)", header_value, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip().strip('"')
    return ""


def _extract_view_ids_from_url(page_url: str) -> list[str]:
    try:
        parsed = urlparse(page_url or "")
        m = re.search(r"/view/([^/?#]+)", parsed.path or "")
        if not m:
            return []
        seg = (m.group(1) or "").strip()
        if not seg:
            return []
        out = [seg]
        if seg.startswith("v_") and "-" in seg:
            base = seg.split("-", 1)[0].strip()
            if base:
                out.append(base)
        return _dedupe_keep_order(out)
    except Exception:
        return []


def _extract_access_tokens(page: Page) -> list[str]:
    tokens: list[str] = []

    try:
        parsed = urlparse(page.url or "")
        q = parse_qs(parsed.query or "")
        for key in ("access_token", "token"):
            vals = q.get(key, [])
            for v in vals:
                if v and len(v) >= 16:
                    tokens.append(v)
    except Exception:
        pass

    try:
        storages = page.evaluate(
            """() => {
                const ls = {};
                const ss = {};
                try {
                    for (let i = 0; i < localStorage.length; i++) {
                        const k = localStorage.key(i);
                        ls[k] = localStorage.getItem(k);
                    }
                } catch (_) {}
                try {
                    for (let i = 0; i < sessionStorage.length; i++) {
                        const k = sessionStorage.key(i);
                        ss[k] = sessionStorage.getItem(k);
                    }
                } catch (_) {}
                return { ls, ss };
            }"""
        )
        if isinstance(storages, dict):
            for root in ("ls", "ss"):
                bag = storages.get(root) or {}
                if not isinstance(bag, dict):
                    continue
                for k, v in bag.items():
                    key = str(k or "").lower()
                    val = str(v or "")
                    if ("access_token" in key or key == "token") and len(val) >= 16:
                        tokens.append(val)
                    if val:
                        for m in re.finditer(
                            r'(?i)(?:access_token|token)\s*["\':=,\s]+\s*["\']?([A-Za-z0-9._~+/=\-]{16,})',
                            val,
                        ):
                            tokens.append(m.group(1))
    except Exception:
        pass

    try:
        for c in page.context.cookies():
            name = str(c.get("name", "")).lower()
            value = str(c.get("value", ""))
            if ("token" in name or "auth" in name) and len(value) >= 16:
                tokens.append(value)
    except Exception:
        pass

    return _dedupe_keep_order(tokens)


def _looks_like_csv_payload(content_type: str, body: bytes) -> bool:
    if not body:
        return False
    ct = (content_type or "").lower()
    if "text/csv" in ct:
        return True
    head = body[:800].decode("utf-8", errors="ignore")
    low = head.lower()
    if "<html" in low:
        return False
    if "," in head or ";" in head or "\t" in head:
        return True
    if "octet-stream" in ct or "text/plain" in ct:
        return True
    return False


def _try_direct_csv_download(page: Page, output_name: str = "") -> Optional[tuple[str, str]]:
    page_url = page.url or ""
    parsed = urlparse(page_url)
    if not parsed.scheme or not parsed.netloc:
        return None
    base_origin = f"{parsed.scheme}://{parsed.netloc}"
    view_ids = _extract_view_ids_from_url(page_url)
    if not view_ids:
        return None

    tokens = _extract_access_tokens(page)

    urls: list[str] = []
    for view_id in view_ids:
        base = f"{base_origin}/api/views/{view_id}/data/download"
        urls.append(base)
        for tok in tokens[:8]:
            urls.append(f"{base}?{urlencode({'access_token': tok})}")

    for url in _dedupe_keep_order(urls):
        try:
            resp = page.request.get(url, timeout=20000, fail_on_status_code=False)
        except Exception:
            continue
        if not resp.ok:
            continue

        body = resp.body()
        content_type = resp.header_value("content-type") or ""
        if not _looks_like_csv_payload(content_type, body):
            continue

        suggested = (
            output_name.strip()
            or _filename_from_content_disposition(resp.header_value("content-disposition") or "")
            or "export.csv"
        )
        safe_name = _safe_filename(suggested)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", prefix="ask_export_direct_")
        tmp_path = tmp.name
        tmp.close()
        with open(tmp_path, "wb") as f:
            f.write(body)
        return tmp_path, safe_name

    return None


def _looks_logged_in(page: Page) -> bool:
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""

    if "/view/" in url:
        return True

    for sel in ["text=Logga ut", "text=Connected", "text=Rapportera fel"]:
        try:
            loc = page.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                return True
        except Exception:
            continue
    return False


def _wait_for_login_ready(page: Page, wait_sec: int) -> bool:
    deadline = time.time() + max(0, int(wait_sec))
    while time.time() < deadline:
        if _looks_logged_in(page):
            return True
        page.wait_for_timeout(700)
    return _looks_logged_in(page)


class InitBody(BaseModel):
    url: str
    login_wait: int = 60
    headless: bool = False
    slow_mo: int = 80
    goto_timeout: int = 60


class FetchBody(BaseModel):
    view_name: str = "Order\u00f6versikt"
    open_via: Literal["auto", "tab", "menu", "shortcut"] = "auto"
    open_text: str = "Visa"
    export_text: str = "Exportera till CSV"
    grid_wait: int = 30
    download_timeout: int = 120
    output_name: str = ""
    url: Optional[str] = None
    goto_timeout: int = 60


class AskCsvAgent:
    def __init__(self):
        self._lock = threading.RLock()
        self._pw = None
        self._browser = None
        self._context = None
        self._page: Optional[Page] = None
        self._active_url: str = ""
        self._ready: bool = False

    def _ensure_browser(self, headless: bool, slow_mo: int) -> None:
        if self._browser:
            return
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=headless, slow_mo=slow_mo)
        self._context = self._browser.new_context(accept_downloads=True, locale="sv-SE")
        self._page = self._context.new_page()

    def init_login(self, body: InitBody) -> dict:
        with self._lock:
            self._ensure_browser(headless=body.headless, slow_mo=body.slow_mo)
            assert self._page is not None
            page = self._page
            page.goto(body.url, wait_until="domcontentloaded", timeout=body.goto_timeout * 1000)
            self._active_url = body.url
            self._ready = False

        ready = _wait_for_login_ready(page, wait_sec=body.login_wait)
        with self._lock:
            self._ready = bool(ready)
            current_url = self._page.url if self._page else ""

        return {
            "ok": True,
            "ready": self._ready,
            "url": current_url,
            "message": "Login klar" if self._ready else "Login ej klar inom väntetid",
        }

    def warm_open(self, url: str, headless: bool = False, slow_mo: int = 80, goto_timeout: int = 60) -> dict:
        with self._lock:
            self._ensure_browser(headless=headless, slow_mo=slow_mo)
            assert self._page is not None
            self._page.goto(url, wait_until="domcontentloaded", timeout=max(10, int(goto_timeout)) * 1000)
            self._active_url = url
            self._ready = _looks_logged_in(self._page)
            return {
                "ok": True,
                "ready": self._ready,
                "url": self._page.url if self._page else "",
                "message": "Browser öppnad",
            }

    def status(self) -> dict:
        with self._lock:
            return {
                "ok": True,
                "ready": self._ready,
                "active_url": self._active_url,
                "page_url": self._page.url if self._page else "",
                "browser_alive": bool(self._browser),
            }

    def fetch_csv(self, body: FetchBody) -> tuple[str, str]:
        with self._lock:
            if not self._page:
                raise RuntimeError("Agenten Ã¤r inte initierad. KÃ¶r /init-login fÃ¶rst.")
            if not self._ready:
                raise RuntimeError("Agenten Ã¤r inte redo Ã¤nnu. VÃ¤nta tills login/2FA Ã¤r klart.")

            page = self._page

            if body.url:
                page.goto(body.url, wait_until="domcontentloaded", timeout=body.goto_timeout * 1000)
                self._active_url = body.url

            opened = False
            if body.open_via == "tab":
                opened = _click_by_text(page, body.view_name, timeout_ms=7000)
            elif body.open_via == "menu":
                opened = _open_view_from_menu(page, body.view_name, use_shortcut=False)
            elif body.open_via == "shortcut":
                opened = _open_view_from_menu(page, body.view_name, use_shortcut=True)
            else:
                opened = _click_by_text(page, body.view_name, timeout_ms=5000)
                if not opened:
                    opened = _open_view_from_menu(page, body.view_name, use_shortcut=True)
                if not opened:
                    opened = _open_view_from_menu(page, body.view_name, use_shortcut=False)

            if not opened:
                raise RuntimeError(f"Kunde inte Ã¶ppna vy '{body.view_name}'.")

            if body.open_text and body.open_text.strip():
                _click_by_text(page, body.open_text.strip(), timeout_ms=7000)

            # Preferred path: call the same ASK download endpoint directly.
            direct = _try_direct_csv_download(page, output_name=body.output_name.strip())
            if direct:
                return direct

            # Fallback path: UI context-menu export.
            if not _right_click_in_primary_grid(page, timeout_ms=body.grid_wait * 1000):
                raise RuntimeError("Kunde inte hitta tabell/grid att högerklicka i.")
            try:
                with page.expect_download(timeout=body.download_timeout * 1000) as dl_info:
                    ok = _click_by_text(page, body.export_text, timeout_ms=10000)
                    if not ok:
                        raise RuntimeError(f"Kunde inte hitta menyval '{body.export_text}'.")
                download = dl_info.value
            except PlaywrightTimeoutError as e:
                raise RuntimeError("Ingen fil laddades ner efter export-klick.") from e

            suggested = body.output_name.strip() or download.suggested_filename or "export.csv"
            safe_name = _safe_filename(suggested)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", prefix="ask_export_")
            tmp_path = tmp.name
            tmp.close()
            download.save_as(tmp_path)
            return tmp_path, safe_name

    def shutdown(self) -> dict:
        with self._lock:
            try:
                if self._context:
                    self._context.close()
            except Exception:
                pass
            try:
                if self._browser:
                    self._browser.close()
            except Exception:
                pass
            try:
                if self._pw:
                    self._pw.stop()
            except Exception:
                pass
            self._pw = None
            self._browser = None
            self._context = None
            self._page = None
            self._ready = False
            return {"ok": True}


app = FastAPI(title="ASK CSV Listener")
agent = AskCsvAgent()


@app.on_event("startup")
def on_startup():
    # Open ASK immediately so user can log in even before the webapp sends requests.
    if not DEFAULT_ASK_URL:
        return
    try:
        agent.warm_open(url=DEFAULT_ASK_URL, headless=False, slow_mo=80, goto_timeout=60)
    except Exception:
        # Keep service running even if auto-open fails.
        pass


@app.get("/health")
def health():
    return agent.status()


@app.post("/init-login")
def init_login(body: InitBody):
    try:
        return agent.init_login(body)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@app.post("/fetch-csv")
def fetch_csv(body: FetchBody):
    try:
        tmp_path, filename = agent.fetch_csv(body)
        return FileResponse(
            path=tmp_path,
            media_type="text/csv",
            filename=filename,
            background=BackgroundTask(lambda: os.path.exists(tmp_path) and os.remove(tmp_path)),
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@app.post("/shutdown")
def shutdown():
    return agent.shutdown()


