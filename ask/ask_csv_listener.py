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
import unicodedata
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


def _fold_text(value: str) -> str:
    txt = unicodedata.normalize("NFKD", value or "")
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", txt).strip().casefold()


def _view_tokens(view_name: str) -> list[str]:
    base = _fold_text(view_name)
    tokens = [t for t in re.split(r"[^a-z0-9]+", base) if len(t) >= 3]
    return tokens[:8]


def _verify_view_open(page: Page, view_name: str, wait_ms: int = 2200) -> bool:
    tokens = _view_tokens(view_name)
    folded_hint = _fold_text(view_name)
    step_ms = 200
    rounds = max(1, wait_ms // step_ms)

    for _ in range(rounds):
        try:
            folded_url = _fold_text(page.url or "")
            if folded_hint and folded_hint in folded_url:
                return True
            if tokens:
                hit_count = sum(1 for tok in tokens if tok in folded_url)
                if hit_count >= min(2, len(tokens)):
                    return True
        except Exception:
            pass

        try:
            folded_title = _fold_text(page.title() or "")
            if folded_hint and folded_hint in folded_title:
                return True
            if tokens:
                hit_count = sum(1 for tok in tokens if tok in folded_title)
                if hit_count >= min(2, len(tokens)):
                    return True
        except Exception:
            pass

        for sel in ["[role='tab']", ".nav-link", "button", "a"]:
            try:
                loc = page.locator(sel).filter(has_text=view_name)
                count = min(loc.count(), 15)
                for i in range(count):
                    item = loc.nth(i)
                    if not item.is_visible():
                        continue
                    bbox = item.bounding_box()
                    if bbox and bbox["y"] <= 220:
                        return True
            except Exception:
                continue

        page.wait_for_timeout(step_ms)

    return False


def _find_primary_grid_bbox(page: Page) -> Optional[dict]:
    viewport = page.viewport_size or {"width": 1920, "height": 1080}
    vheight = float(viewport.get("height", 1080))
    upper_limit = vheight * 0.72

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
        timeout_ms=900,
        max_per_selector=120,
    )
    if row:
        try:
            bbox = row.bounding_box()
            if bbox and bbox["width"] >= 250 and 90 <= bbox["y"] <= upper_limit:
                return bbox
        except Exception:
            pass

    candidates: list[dict] = []
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
            count = min(loc.count(), 35)
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
                candidates.append(bbox)
        except Exception:
            continue

    if not candidates:
        return None
    return sorted(candidates, key=_grid_box_score)[0]


def _collect_upper_three_dot_buttons(page: Page) -> list[Locator]:
    viewport = page.viewport_size or {"width": 1920, "height": 1080}
    vheight = float(viewport.get("height", 1080))
    upper_limit = vheight * 0.72
    grid_box = _find_primary_grid_bbox(page)
    target_y = float(grid_box["y"] - 25.0) if grid_box else 130.0

    scored: list[tuple[tuple[float, float, float], Locator]] = []
    selectors = [
        "button:has(i.bi.bi-three-dots-vertical)",
        "[role='button']:has(i.bi.bi-three-dots-vertical)",
        "button:has(i.bi-three-dots-vertical)",
        "[role='button']:has(i.bi-three-dots-vertical)",
    ]
    for sel in selectors:
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
                if bbox["y"] < 45 or bbox["y"] > upper_limit:
                    continue
                if bbox["width"] < 14 or bbox["width"] > 85:
                    continue
                if bbox["height"] < 14 or bbox["height"] > 85:
                    continue
                score = (
                    abs((bbox["y"] + (bbox["height"] / 2.0)) - target_y),
                    bbox["y"],
                    bbox["x"],
                )
                scored.append((score, item))
        except Exception:
            continue

    if not scored:
        return []

    deduped: list[Locator] = []
    seen_boxes: set[str] = set()
    for _, item in sorted(scored, key=lambda t: t[0]):
        try:
            bbox = item.bounding_box() or {}
            key = f"{round(float(bbox.get('x', 0.0)), 1)}:{round(float(bbox.get('y', 0.0)), 1)}"
        except Exception:
            key = f"loc-{len(deduped)}"
        if key in seen_boxes:
            continue
        seen_boxes.add(key)
        deduped.append(item)
    return deduped


def _find_open_menu_near_button(page: Page, btn_bbox: Optional[dict], timeout_ms: int = 2600):
    deadline = time.time() + (max(100, timeout_ms) / 1000.0)
    while time.time() < deadline:
        try:
            menus = page.locator("ul.dropdown-menu.show[role='menu']")
            count = min(menus.count(), 8)
            visible: list[tuple[float, Locator]] = []
            for i in range(count):
                menu = menus.nth(i)
                if not menu.is_visible():
                    continue
                bbox = menu.bounding_box()
                if not bbox:
                    continue
                if not btn_bbox:
                    score = float(bbox["y"])
                else:
                    dx = abs(float(bbox["x"]) - float(btn_bbox["x"]))
                    dy = abs(float(bbox["y"]) - float(btn_bbox["y"]))
                    score = (dy * 2.0) + dx
                visible.append((score, menu))
            if visible:
                visible.sort(key=lambda t: t[0])
                return visible[0][1]
        except Exception:
            pass
        page.wait_for_timeout(100)
    return None


def _click_export_in_dropdown(menu: Locator, export_text: str, timeout_ms: int = 2500) -> tuple[bool, str]:
    candidates = [
        ("css_span", menu.locator("a.dropdown-item:has(span:has-text('Exportera till CSV'))")),
        ("css_item", menu.locator("a.dropdown-item", has_text=export_text)),
        ("css_text", menu.locator("a.dropdown-item:has-text('Exportera till CSV')")),
    ]
    last_hits = "none"
    deadline = time.time() + (max(100, timeout_ms) / 1000.0)
    while time.time() < deadline:
        for key, loc in candidates:
            try:
                count = min(loc.count(), 10)
                if count <= 0:
                    continue
                last_hits = f"{key}:{count}"
                for i in range(count):
                    item = loc.nth(i)
                    if not item.is_visible():
                        continue
                    cls = (item.get_attribute("class") or "").lower()
                    if "disabled" in cls:
                        continue
                    item.click()
                    return True, last_hits
            except Exception:
                continue
        time.sleep(0.1)
    return False, last_hits


def _capture_error_screenshot(page: Page, prefix: str) -> str:
    out_dir = Path(tempfile.gettempdir()) / "ask_csv_listener"
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"{prefix}_{ts}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        return str(path)
    except Exception:
        return ""


def _export_via_three_dots(page: Page, export_text: str, download_timeout: int, step_log: list[str]):
    buttons = _collect_upper_three_dot_buttons(page)
    step_log.append(f"three_dot_buttons={len(buttons)}")
    if not buttons:
        raise RuntimeError("no_upper_three_dot_button")

    attempts = min(3, len(buttons))
    last_error = "menu_export_unknown"
    for idx in range(attempts):
        btn = buttons[idx]
        try:
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(80)
            except Exception:
                pass
            btn_bbox = btn.bounding_box()
            step_log.append(f"menu_try_{idx + 1}_btn_y={round(float((btn_bbox or {}).get('y', -1.0)), 1)}")

            with page.expect_download(timeout=max(1, int(download_timeout)) * 1000) as dl_info:
                btn.click()
                menu = _find_open_menu_near_button(page, btn_bbox=btn_bbox, timeout_ms=3200)
                if not menu:
                    raise RuntimeError("menu_not_found")
                ok, hit = _click_export_in_dropdown(menu, export_text=export_text, timeout_ms=2600)
                step_log.append(f"menu_try_{idx + 1}_selector_hits={hit}")
                if not ok:
                    raise RuntimeError("export_item_not_found")
            return dl_info.value
        except PlaywrightTimeoutError:
            last_error = f"menu_try_{idx + 1}_download_timeout"
            step_log.append(last_error)
            continue
        except Exception as e:
            last_error = f"menu_try_{idx + 1}_failed:{e}"
            step_log.append(last_error)
            continue

    raise RuntimeError(last_error)


def _export_via_right_click(
    page: Page,
    export_text: str,
    grid_wait_sec: int,
    download_timeout: int,
    step_log: list[str],
):
    if not _right_click_in_primary_grid(page, timeout_ms=max(1, int(grid_wait_sec)) * 1000):
        raise RuntimeError("grid_not_found_for_right_click")
    step_log.append("right_click_grid_ok")
    try:
        with page.expect_download(timeout=max(1, int(download_timeout)) * 1000) as dl_info:
            ok = _click_by_text(page, export_text, timeout_ms=10000)
            step_log.append(f"right_click_export_click={ok}")
            if not ok:
                raise RuntimeError(f"export_menu_item_not_found:{export_text}")
        return dl_info.value
    except PlaywrightTimeoutError as e:
        raise RuntimeError("right_click_download_timeout") from e


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

    if candidates:
        _, best_box = sorted(candidates, key=lambda it: _grid_box_score(it[1]))[0]
        x = best_box["x"] + min(260.0, max(90.0, best_box["width"] * 0.18))
        y = best_box["y"] + min(180.0, max(55.0, best_box["height"] * 0.3))
    else:
        # Last resort for views where row/grid selectors are unstable:
        # click a safe point in the upper data area.
        vwidth = float(viewport.get("width", 1920))
        x = min(max(140.0, vwidth * 0.22), vwidth - 140.0)
        y = min(max(180.0, vheight * 0.30), upper_limit - 20.0)

    page.mouse.click(x, y, button="left")
    page.wait_for_timeout(90)
    page.mouse.click(x, y, button="right")
    return True


def _open_view_from_menu_once(page: Page, view_name: str, use_shortcut: bool = True) -> bool:
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
            "input[placeholder*='Sök']",
            "input[placeholder*='sök' i]",
            "input[placeholder*='sok' i]",
            "input[placeholder*='search' i]",
            ".dropdown-menu.show input[type='search']",
            ".dropdown-menu.show input[type='text']",
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
                "button:has(i.bi-plus-circle)",
                "[role='button']:has(i.bi-plus-circle)",
            ],
            timeout_ms=5000,
        )
        if not plus:
            return False
        plus.click()

        search = _first_visible(
            page,
            selectors=[
                "input[placeholder*='Sök']",
                "input[placeholder*='sök' i]",
                "input[placeholder*='sok' i]",
                "input[placeholder*='search' i]",
                ".dropdown-menu.show input[type='search']",
                ".dropdown-menu.show input[type='text']",
                "input[type='search']",
                "input[type='text']",
            ],
            timeout_ms=7000,
        )
        if not search:
            return False

    try:
        search.click()
        search.press("ControlOrMeta+A")
        search.press("Backspace")
        search.fill(view_name)
        page.wait_for_timeout(350)
        search.press("Enter")
        page.wait_for_timeout(650)
        return True
    except Exception:
        return False


def _open_view_from_menu(
    page: Page,
    view_name: str,
    use_shortcut: bool = True,
    retries: int = 1,
    step_log: Optional[list[str]] = None,
) -> bool:
    logs = step_log if step_log is not None else []
    for attempt in range(max(0, retries) + 1):
        ok = _open_view_from_menu_once(page, view_name=view_name, use_shortcut=use_shortcut)
        logs.append(f"open_view_try_{attempt + 1}_input={ok}")
        if ok and _verify_view_open(page, view_name=view_name, wait_ms=2500):
            logs.append(f"open_view_try_{attempt + 1}_verify=ok")
            return True
        logs.append(f"open_view_try_{attempt + 1}_verify=failed")
        if attempt < retries:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            page.wait_for_timeout(220)
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
            step_log: list[str] = []

            if body.url:
                page.goto(body.url, wait_until="domcontentloaded", timeout=body.goto_timeout * 1000)
                self._active_url = body.url
                step_log.append("goto_body_url=ok")

            opened = False
            if body.open_via == "tab":
                opened = _click_by_text(page, body.view_name, timeout_ms=7000)
                step_log.append(f"open_via_tab_click={opened}")
                if opened:
                    opened = _verify_view_open(page, view_name=body.view_name, wait_ms=2400)
                    step_log.append(f"open_via_tab_verify={opened}")
            elif body.open_via == "menu":
                opened = _open_view_from_menu(
                    page,
                    body.view_name,
                    use_shortcut=False,
                    retries=1,
                    step_log=step_log,
                )
            elif body.open_via == "shortcut":
                opened = _open_view_from_menu(
                    page,
                    body.view_name,
                    use_shortcut=True,
                    retries=1,
                    step_log=step_log,
                )
            else:
                opened = _click_by_text(page, body.view_name, timeout_ms=5000)
                step_log.append(f"open_auto_tab_click={opened}")
                if opened:
                    opened = _verify_view_open(page, view_name=body.view_name, wait_ms=2000)
                    step_log.append(f"open_auto_tab_verify={opened}")
                if not opened:
                    opened = _open_view_from_menu(
                        page,
                        body.view_name,
                        use_shortcut=True,
                        retries=1,
                        step_log=step_log,
                    )
                if not opened:
                    opened = _open_view_from_menu(
                        page,
                        body.view_name,
                        use_shortcut=False,
                        retries=1,
                        step_log=step_log,
                    )

            if not opened:
                screenshot = _capture_error_screenshot(page, "open_view_failed")
                steps = " | ".join(step_log[-18:])
                raise RuntimeError(
                    f"open_view_failed: kunde inte oppna vy '{body.view_name}'. "
                    f"screenshot: {screenshot or 'n/a'}. steps: {steps}"
                )

            if body.open_text and body.open_text.strip():
                clicked_open = _click_by_text(page, body.open_text.strip(), timeout_ms=7000)
                step_log.append(f"open_text_click={clicked_open}")

            menu_error: Optional[str] = None
            try:
                download = _export_via_three_dots(
                    page=page,
                    export_text=body.export_text,
                    download_timeout=body.download_timeout,
                    step_log=step_log,
                )
                step_log.append("menu_export=ok")
            except Exception as menu_exc:
                menu_error = str(menu_exc)
                step_log.append(f"menu_export_failed:{menu_error}")
                try:
                    download = _export_via_right_click(
                        page=page,
                        export_text=body.export_text,
                        grid_wait_sec=body.grid_wait,
                        download_timeout=body.download_timeout,
                        step_log=step_log,
                    )
                    step_log.append("right_click_fallback=ok")
                except Exception as rc_exc:
                    screenshot = _capture_error_screenshot(page, "csv_export_failed")
                    steps = " | ".join(step_log[-24:])
                    raise RuntimeError(
                        f"menu_export_failed: {menu_error}; "
                        f"right_click_failed: {rc_exc}; "
                        f"screenshot: {screenshot or 'n/a'}; "
                        f"steps: {steps}"
                    ) from rc_exc

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


