#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import importlib.util
import math
import os
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Callable, Dict, Iterable, Optional

import pandas as pd
from pandas.testing import assert_frame_equal


REPO_ROOT = Path(r"C:\allokering")
WEBAPP_BACKEND = REPO_ROOT / "webapp" / "backend"
CURRENT_LOGIC_PATH = WEBAPP_BACKEND / "logic.py"
CURRENT_MAIN_PATH = WEBAPP_BACKEND / "main.py"
REFERENCE_APP_PATH = REPO_ROOT / "allokera12.py"
TESTDATA_DIR = REPO_ROOT / "testdata"


CSV_FILES: Dict[str, str] = {
    "orders": "v_ask_customer_order_details_all",
    "buffer": "v_ask_article_buffertpallet",
    "overview": "v_ask_order_overview",
    "automation": "v_ask_item_summary_stock_automation",
    "items": "item_option",
    "booking_putaway": "v_ask_booking_putaway",
    "receive_log": "v_ask_receive_log",
    "correct_log": "v_ask_correct_log",
}

XLSX_FILES: Dict[str, str] = {
    "prognos": "Prognos idag",
    "campaign": "Granngården prognos kampanjplock +6v",
}


@dataclass
class TimingResult:
    name: str
    baseline_seconds: float
    current_seconds: float

    @property
    def delta_seconds(self) -> float:
        return self.current_seconds - self.baseline_seconds

    @property
    def delta_percent(self) -> float:
        if self.baseline_seconds == 0:
            return 0.0
        return (self.delta_seconds / self.baseline_seconds) * 100.0

    @property
    def speedup(self) -> float:
        if self.current_seconds == 0:
            return math.inf
        return self.baseline_seconds / self.current_seconds


def load_module_from_path(module_name: str, path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Kunde inte ladda modul från {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def load_head_logic_module(temp_dir: Path):
    result = subprocess.run(
        ["git", "show", "HEAD:webapp/backend/logic.py"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=True,
    )
    out_path = temp_dir / "logic_head.py"
    out_path.write_text(result.stdout, encoding="utf-8")
    return load_module_from_path("logic_head_benchmark", out_path)


def find_test_file(prefix: str, suffix: str) -> Path:
    matches = sorted(TESTDATA_DIR.glob(f"{prefix}*{suffix}"))
    if not matches:
        raise FileNotFoundError(f"Hittade ingen testfil för prefix '{prefix}' och suffix '{suffix}'")
    return matches[0]


def test_paths() -> Dict[str, Path]:
    paths: Dict[str, Path] = {}
    for key, prefix in CSV_FILES.items():
        paths[key] = find_test_file(prefix, ".csv")
    for key, prefix in XLSX_FILES.items():
        paths[key] = find_test_file(prefix, ".xlsx")
    return paths


def best_time(callable_fn: Callable[[], Any], repeat: int = 3) -> tuple[Any, float]:
    best_value = None
    best_seconds = None
    for _ in range(max(1, repeat)):
        start = time.perf_counter()
        value = callable_fn()
        elapsed = time.perf_counter() - start
        if best_seconds is None or elapsed < best_seconds:
            best_seconds = elapsed
            best_value = value
    return best_value, float(best_seconds or 0.0)


def normalize_scalar(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        if pd.isna(value):
            return None
        return pd.Timestamp(value).isoformat()
    if isinstance(value, float):
        if pd.isna(value):
            return None
        return value
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    return value


def normalize_df(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    out = out.reset_index(drop=True)
    out.columns = [str(col) for col in out.columns]
    for col in out.columns:
        series = out[col]
        if pd.api.types.is_datetime64_any_dtype(series):
            out[col] = series.map(normalize_scalar)
        else:
            out[col] = series.map(normalize_scalar)
    return out


def frames_match(left: Optional[pd.DataFrame], right: Optional[pd.DataFrame]) -> tuple[bool, str]:
    try:
        assert_frame_equal(
            normalize_df(left),
            normalize_df(right),
            check_dtype=False,
            check_like=False,
        )
        return True, ""
    except AssertionError as exc:
        return False, str(exc).splitlines()[0]


def print_timing(result: TimingResult) -> None:
    faster_slower = "snabbare" if result.delta_seconds < 0 else "långsammare"
    print(
        f"{result.name}: HEAD {result.baseline_seconds:.3f}s -> nu {result.current_seconds:.3f}s | "
        f"delta {result.delta_seconds:+.3f}s ({result.delta_percent:+.1f}%) | {result.speedup:.2f}x | {faster_slower}"
    )


def compare_csv_reader(head_logic, current_logic, paths: Dict[str, Path]) -> list[TimingResult]:
    print("\nCSV-läsning")
    timings: list[TimingResult] = []
    for key, path in paths.items():
        if path.suffix.lower() != ".csv":
            continue
        head_df, head_time = best_time(lambda p=path: head_logic.read_csv_auto(str(p)), repeat=3)
        current_df, current_time = best_time(lambda p=path: current_logic.read_csv_auto(str(p)), repeat=3)
        same, detail = frames_match(head_df, current_df)
        timing = TimingResult(f"read_csv_auto[{key}]", head_time, current_time)
        print_timing(timing)
        print(f"  resultat: {'OK' if same else 'AVVIKER'}{f' | {detail}' if detail else ''}")
        timings.append(timing)
    return timings


def compare_xlsx_readers(reference_app, current_logic, paths: Dict[str, Path]) -> None:
    print("\nXLSX-läsning mot allokera12.py")
    for key, path in paths.items():
        if key == "prognos":
            ref_df, ref_time = best_time(lambda p=path: reference_app.read_prognos_xlsx(str(p)), repeat=2)
            cur_df, cur_time = best_time(lambda p=path: current_logic.read_prognos_xlsx(str(p)), repeat=2)
            same, detail = frames_match(ref_df, cur_df)
            timing = TimingResult("read_prognos_xlsx", ref_time, cur_time)
        elif key == "campaign":
            ref_df, ref_time = best_time(lambda p=path: reference_app.read_campaign_xlsx(str(p)), repeat=2)
            cur_df, cur_time = best_time(lambda p=path: current_logic.read_campaign_xlsx(str(p)), repeat=2)
            same, detail = frames_match(ref_df, cur_df)
            timing = TimingResult("read_campaign_xlsx", ref_time, cur_time)
        else:
            continue
        print_timing(timing)
        print(f"  resultat: {'OK' if same else 'AVVIKER'}{f' | {detail}' if detail else ''}")


def slow_collect_filter_options(paths: Iterable[str], read_csv_auto_fn, scan_filter_values_fn) -> Dict[str, list]:
    combined: Dict[str, list] = {}
    for raw_path in paths:
        if not raw_path or not os.path.exists(raw_path):
            continue
        try:
            df = read_csv_auto_fn(raw_path)
            vals = scan_filter_values_fn(df)
            for fk, fvals in vals.items():
                existing = set(combined.get(fk, []))
                existing.update(fvals)
                combined[fk] = sorted(existing)
        except Exception:
            continue
    return combined


def compare_filter_cache(current_main, current_logic, paths: Dict[str, Path]) -> None:
    print("\nFilter-cache")
    session = SimpleNamespace(
        files={key: str(path) for key, path in paths.items()},
        filter_scan_cache={},
        ordersaldo_list1=[],
        ordersaldo_list2=[],
    )
    slow_result, slow_time = best_time(
        lambda: slow_collect_filter_options(session.files.values(), current_logic.read_csv_auto, current_logic.scan_filter_values),
        repeat=1,
    )
    cold_result, cold_time = best_time(lambda: current_main._collect_filter_options(session), repeat=1)
    warm_result, warm_time = best_time(lambda: current_main._collect_filter_options(session), repeat=3)
    slow_vs_cold, detail_cold = frames_match(pd.DataFrame([slow_result]), pd.DataFrame([cold_result]))
    slow_vs_warm, detail_warm = frames_match(pd.DataFrame([slow_result]), pd.DataFrame([warm_result]))
    print_timing(TimingResult("filter_scan cold-cache", slow_time, cold_time))
    print(f"  resultat: {'OK' if slow_vs_cold else 'AVVIKER'}{f' | {detail_cold}' if detail_cold else ''}")
    print(
        f"filter_scan warm-cache: tidigare {slow_time:.3f}s -> varm cache {warm_time:.3f}s | "
        f"delta {warm_time - slow_time:+.3f}s ({((warm_time - slow_time) / slow_time * 100.0):+.1f}%) | "
        f"{(slow_time / warm_time) if warm_time else math.inf:.2f}x"
    )
    print(f"  resultat: {'OK' if slow_vs_warm else 'AVVIKER'}{f' | {detail_warm}' if detail_warm else ''}")


def workbook_to_frames(path: Path) -> Dict[str, pd.DataFrame]:
    with pd.ExcelFile(path) as workbook:
        return {sheet: normalize_df(pd.read_excel(workbook, sheet_name=sheet, dtype=object)) for sheet in workbook.sheet_names}


def compare_excel_export(head_logic, current_logic, allocated_df: pd.DataFrame, near_df: pd.DataFrame) -> TimingResult:
    print("\nExcel-export")
    sheets = {
        "Allokering": allocated_df,
        "NearMiss": near_df,
    }
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        head_path = tmp_dir / "head_export.xlsx"
        current_path = tmp_dir / "current_export.xlsx"
        _, head_time = best_time(lambda: head_logic.save_df_to_excel(sheets, "bench", str(head_path)), repeat=2)
        _, current_time = best_time(lambda: current_logic.save_df_to_excel(sheets, "bench", str(current_path)), repeat=2)
        same, detail = frames_match_dict(workbook_to_frames(head_path), workbook_to_frames(current_path))
    timing = TimingResult("save_df_to_excel", head_time, current_time)
    print_timing(timing)
    print(f"  resultat: {'OK' if same else 'AVVIKER'}{f' | {detail}' if detail else ''}")
    return timing


def frames_match_dict(left: Dict[str, pd.DataFrame], right: Dict[str, pd.DataFrame]) -> tuple[bool, str]:
    if set(left.keys()) != set(right.keys()):
        return False, f"olika blad: {sorted(left.keys())} vs {sorted(right.keys())}"
    for sheet in left:
        same, detail = frames_match(left[sheet], right[sheet])
        if not same:
            return False, f"{sheet}: {detail}"
    return True, ""


def load_shared_inputs(head_logic, paths: Dict[str, Path]) -> Dict[str, pd.DataFrame]:
    shared: Dict[str, pd.DataFrame] = {}
    for key in ("orders", "buffer", "overview", "automation", "items", "booking_putaway", "receive_log", "correct_log"):
        shared[key] = head_logic.read_csv_auto(str(paths[key]))
    shared["booking_putaway_norm"] = head_logic.normalize_not_putaway(shared["booking_putaway"])
    return shared


def compare_logic_workflows(head_logic, current_logic, reference_app, data: Dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame]:
    print("\nKärnflöden mot allokera12.py")

    ref_alloc, ref_alloc_time = best_time(
        lambda: reference_app.allocate(data["orders"].copy(), data["buffer"].copy()),
        repeat=1,
    )
    head_alloc, head_alloc_time = best_time(
        lambda: head_logic.allocate(data["orders"].copy(), data["buffer"].copy()),
        repeat=1,
    )
    cur_alloc, cur_alloc_time = best_time(
        lambda: current_logic.allocate(data["orders"].copy(), data["buffer"].copy()),
        repeat=1,
    )
    ref_alloc_df, ref_near_df = ref_alloc
    head_alloc_df, head_near_df = head_alloc
    cur_alloc_df, cur_near_df = cur_alloc

    print_timing(TimingResult("allocate", head_alloc_time, cur_alloc_time))
    same_alloc, detail_alloc = frames_match(ref_alloc_df, cur_alloc_df)
    same_near, detail_near = frames_match(ref_near_df, cur_near_df)
    print(f"  allokering mot allokera12.py: {'OK' if same_alloc else 'AVVIKER'}{f' | {detail_alloc}' if detail_alloc else ''}")
    print(f"  near-miss mot allokera12.py: {'OK' if same_near else 'AVVIKER'}{f' | {detail_near}' if detail_near else ''}")
    same_head_alloc, detail_head_alloc = frames_match(ref_alloc_df, head_alloc_df)
    same_head_near, detail_head_near = frames_match(ref_near_df, head_near_df)
    print(f"  HEAD mot allokera12.py: allokering {'OK' if same_head_alloc else 'AVVIKER'}{f' | {detail_head_alloc}' if detail_head_alloc else ''}")
    print(f"  HEAD mot allokera12.py: near-miss {'OK' if same_head_near else 'AVVIKER'}{f' | {detail_head_near}' if detail_head_near else ''}")

    ref_reclass = reference_app.App._reclassify_skrymmande(ref_alloc_df.copy(), None)
    cur_reclass = current_logic._reclassify_skrymmande(cur_alloc_df.copy(), None)
    same_reclass, detail_reclass = frames_match(ref_reclass, cur_reclass)
    print(f"_reclassify_skrymmande mot allokera12.py: {'OK' if same_reclass else 'AVVIKER'}{f' | {detail_reclass}' if detail_reclass else ''}")

    ref_spaces, ref_spaces_time = best_time(lambda: reference_app.compute_pallet_spaces(ref_reclass.copy()), repeat=1)
    cur_spaces, cur_spaces_time = best_time(lambda: current_logic.compute_pallet_spaces(cur_reclass.copy()), repeat=1)
    print_timing(TimingResult("compute_pallet_spaces", ref_spaces_time, cur_spaces_time))
    same_spaces, detail_spaces = frames_match(ref_spaces, cur_spaces)
    print(f"  resultat: {'OK' if same_spaces else 'AVVIKER'}{f' | {detail_spaces}' if detail_spaces else ''}")

    ref_booking_norm = reference_app.normalize_not_putaway(data["booking_putaway"].copy())
    cur_booking_norm = current_logic.normalize_not_putaway(data["booking_putaway"].copy())
    ref_refill, ref_refill_time = best_time(
        lambda: reference_app.calculate_refill(
            ref_reclass.copy(),
            data["buffer"].copy(),
            saldo_df=data["automation"].copy(),
            not_putaway_df=ref_booking_norm.copy(),
        ),
        repeat=1,
    )
    cur_refill, cur_refill_time = best_time(
        lambda: current_logic.calculate_refill(
            cur_reclass.copy(),
            data["buffer"].copy(),
            saldo_df=data["automation"].copy(),
            not_putaway_df=cur_booking_norm.copy(),
        ),
        repeat=1,
    )
    ref_hp_df, ref_as_df = ref_refill
    cur_hp_df, cur_as_df = cur_refill
    print_timing(TimingResult("calculate_refill", ref_refill_time, cur_refill_time))
    same_hp, detail_hp = frames_match(ref_hp_df, cur_hp_df)
    same_as, detail_as = frames_match(ref_as_df, cur_as_df)
    print(f"  refill HP mot allokera12.py: {'OK' if same_hp else 'AVVIKER'}{f' | {detail_hp}' if detail_hp else ''}")
    print(f"  refill AutoStore mot allokera12.py: {'OK' if same_as else 'AVVIKER'}{f' | {detail_as}' if detail_as else ''}")

    ref_hib, ref_hib_time = best_time(
        lambda: reference_app.compute_hib_koppling(data["orders"].copy(), data["overview"].copy()),
        repeat=1,
    )
    cur_hib, cur_hib_time = best_time(
        lambda: current_logic.compute_hib_koppling(data["orders"].copy(), data["overview"].copy()),
        repeat=1,
    )
    print_timing(TimingResult("compute_hib_koppling", ref_hib_time, cur_hib_time))
    same_hib, detail_hib = frames_match(ref_hib, cur_hib)
    print(f"  resultat: {'OK' if same_hib else 'AVVIKER'}{f' | {detail_hib}' if detail_hib else ''}")

    ref_missed, ref_missed_time = best_time(
        lambda: reference_app.compute_missed_departures(data["orders"].copy(), data["overview"].copy()),
        repeat=1,
    )
    cur_missed, cur_missed_time = best_time(
        lambda: current_logic.compute_missed_departures(data["orders"].copy(), data["overview"].copy()),
        repeat=1,
    )
    print_timing(TimingResult("compute_missed_departures", ref_missed_time, cur_missed_time))
    same_missed, detail_missed = frames_match(ref_missed, cur_missed)
    print(f"  resultat: {'OK' if same_missed else 'AVVIKER'}{f' | {detail_missed}' if detail_missed else ''}")

    return cur_alloc_df, cur_near_df


def main() -> None:
    parser = argparse.ArgumentParser(description="Jämför prestanda och resultat mellan HEAD, nuvarande logic.py och allokera12.py.")
    parser.add_argument(
        "--skip-excel",
        action="store_true",
        help="Hoppa över Excel-exportbenchmarken.",
    )
    args = parser.parse_args()

    paths = test_paths()
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        head_logic = load_head_logic_module(tmp_dir)
        current_logic = load_module_from_path("logic_current_benchmark", CURRENT_LOGIC_PATH)
        current_main = load_module_from_path("main_current_benchmark", CURRENT_MAIN_PATH)
        reference_app = load_module_from_path("allokera12_reference_benchmark", REFERENCE_APP_PATH)

        compare_csv_reader(head_logic, current_logic, paths)
        compare_xlsx_readers(reference_app, current_logic, paths)
        compare_filter_cache(current_main, current_logic, paths)
        shared_inputs = load_shared_inputs(head_logic, paths)
        current_alloc_df, current_near_df = compare_logic_workflows(head_logic, current_logic, reference_app, shared_inputs)
        if not args.skip_excel:
            compare_excel_export(head_logic, current_logic, current_alloc_df, current_near_df)


if __name__ == "__main__":
    main()
