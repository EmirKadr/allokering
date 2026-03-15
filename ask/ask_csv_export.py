#!/usr/bin/env python3
"""
Automate CSV export from the Ask web UI after manual 2FA login.

Flow:
1) Open URL.
2) Wait N seconds for manual login/2FA (default 60).
3) Open target view via tab OR via + menu search / keyboard shortcut (v, o).
4) Right-click a row/grid.
5) Click context-menu item "Exportera till CSV".
6) Save the downloaded CSV file.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import Locator, Page, sync_playwright


def _click_by_text(page: Page, text: str, timeout_ms: int = 4000) -> bool:
    """Try several locator strategies to click a visible element by text."""
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
    """Return the first visible locator among selector candidates."""
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
    """
    Open the view picker and select a view by name.

    If use_shortcut=True, tries keyboard sequence v then o first.
    Falls back to clicking the plus button if needed.
    """
    view_name = (view_name or "").strip()
    if not view_name:
        return False

    if use_shortcut:
        try:
            page.keyboard.press("v")
            page.wait_for_timeout(100)
            page.keyboard.press("o")
            page.wait_for_timeout(250)
        except Exception:
            pass

    search = _first_visible(
        page,
        selectors=[
            "input[placeholder*='Sök']",
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
                "input[placeholder*='Sök']",
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

    # Try to click the matching view row/item.
    if _click_by_text(page, view_name, timeout_ms=5000):
        return True

    # Fallback: press Enter to choose top result.
    try:
        search.press("Enter")
        return True
    except Exception:
        return False


def run(args: argparse.Namespace) -> Path:
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=args.slow_mo)
        context = browser.new_context(accept_downloads=True, locale="sv-SE")
        page = context.new_page()

        print(f"[1/6] Opening: {args.url}")
        page.goto(args.url, wait_until="domcontentloaded", timeout=args.goto_timeout * 1000)

        print(f"[2/6] Waiting {args.login_wait}s for manual login / 2FA...")
        page.wait_for_timeout(args.login_wait * 1000)

        target_view = (args.view_name or args.tab_text or "").strip()
        if target_view:
            print(f"[3/6] Opening view: {target_view} (mode: {args.open_via})")

            opened = False
            if args.open_via == "tab":
                opened = _click_by_text(page, target_view, timeout_ms=7000)
            elif args.open_via == "menu":
                opened = _open_view_from_menu(page, target_view, use_shortcut=False)
            elif args.open_via == "shortcut":
                opened = _open_view_from_menu(page, target_view, use_shortcut=True)
            else:  # auto
                opened = _click_by_text(page, target_view, timeout_ms=5000)
                if not opened:
                    opened = _open_view_from_menu(page, target_view, use_shortcut=True)
                if not opened:
                    opened = _open_view_from_menu(page, target_view, use_shortcut=False)

            if not opened:
                raise RuntimeError(
                    f"Could not open view '{target_view}' via mode '{args.open_via}'."
                )

        if args.open_text:
            print(f"[4/6] Clicking: {args.open_text}")
            if not _click_by_text(page, args.open_text, timeout_ms=7000):
                print(f"Warning: could not click '{args.open_text}'. Continuing anyway.")

        print("[5/6] Right-clicking inside top/main grid...")
        if not _right_click_in_primary_grid(page, timeout_ms=args.grid_wait * 1000):
            raise RuntimeError("Could not find top grid/table to right-click.")

        print(f"[6/6] Clicking context action: {args.export_text}")
        try:
            with page.expect_download(timeout=args.download_timeout * 1000) as dl_info:
                ok = _click_by_text(page, args.export_text, timeout_ms=10000)
                if not ok:
                    raise RuntimeError(f"Could not find context menu item: {args.export_text}")
            download = dl_info.value
        except PlaywrightTimeoutError as e:
            raise RuntimeError(
                "No download event detected after export click. "
                "Check selector texts or that export is enabled in the UI."
            ) from e

        filename = args.output_name.strip() if args.output_name else download.suggested_filename
        if not filename:
            filename = "export.csv"
        if not filename.lower().endswith(".csv"):
            filename += ".csv"

        target = out_dir / filename
        download.save_as(str(target))
        print(f"Done. CSV saved to: {target}")

        if args.keep_open:
            print("Browser kept open. Press Ctrl+C in terminal when done.")
            try:
                while True:
                    page.wait_for_timeout(1000)
            except KeyboardInterrupt:
                pass

        browser.close()
        return target


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export CSV from Ask UI after manual login/2FA.")
    parser.add_argument("--url", required=True, help="Target page URL")
    parser.add_argument(
        "--tab-text",
        default="Orderoversikt",
        help="Tab/button text to open before export (default: Orderoversikt)",
    )
    parser.add_argument(
        "--view-name",
        default="",
        help="View name to open via + menu search / shortcut (overrides --tab-text if set).",
    )
    parser.add_argument(
        "--open-via",
        choices=["auto", "tab", "menu", "shortcut"],
        default="auto",
        help="How to open the view: auto|tab|menu|shortcut (default: auto).",
    )
    parser.add_argument(
        "--open-text",
        default="Visa",
        help="Button text to open the view/grid (default: Visa). Use empty string to skip.",
    )
    parser.add_argument(
        "--export-text",
        default="Exportera till CSV",
        help="Context menu text for export (default: Exportera till CSV)",
    )
    parser.add_argument(
        "--login-wait",
        type=int,
        default=60,
        help="Seconds to wait for manual login and 2FA (default: 60)",
    )
    parser.add_argument(
        "--out-dir",
        default=str(Path.cwd() / "exports"),
        help="Output folder for CSV (default: ./exports)",
    )
    parser.add_argument(
        "--output-name",
        default="",
        help="Optional output filename (e.g. orderoversikt.csv)",
    )
    parser.add_argument(
        "--grid-wait",
        type=int,
        default=30,
        help="Seconds to wait for grid/rows before right-click (default: 30)",
    )
    parser.add_argument(
        "--download-timeout",
        type=int,
        default=120,
        help="Seconds to wait for CSV download event (default: 120)",
    )
    parser.add_argument(
        "--goto-timeout",
        type=int,
        default=60,
        help="Navigation timeout in seconds (default: 60)",
    )
    parser.add_argument(
        "--slow-mo",
        type=int,
        default=80,
        help="Playwright slow motion in milliseconds (default: 80)",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run headless (not recommended for manual 2FA login)",
    )
    parser.add_argument(
        "--keep-open",
        action="store_true",
        help="Keep browser open after export for debugging",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    # Allow disabling open-text from CLI with: --open-text ""
    if args.open_text is not None:
        args.open_text = args.open_text.strip()

    try:
        run(args)
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
