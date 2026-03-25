from __future__ import annotations

import json
import os
import shutil
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

import pandas as pd


ALLOKERING_RESULT_KEYS = ("allokerade", "nearmiss", "pallplatser", "refill")
_ARTIFACT_FILENAMES = {
    "result": "result.pkl",
    "near": "near.pkl",
    "buffer_raw": "buffer_raw.pkl",
    "saldo_norm": "saldo_norm.pkl",
    "not_putaway_norm": "not_putaway_norm.pkl",
}
_MANIFEST_FILENAME = "manifest.json"


def _cache_root(session: Any) -> Path:
    return Path(session.temp_dir) / "artifacts" / "allokering"


def ensure_cache_root(session: Any) -> Path:
    root = _cache_root(session)
    root.mkdir(parents=True, exist_ok=True)
    return root


def clear_generated_results(session: Any) -> None:
    for result_key in ALLOKERING_RESULT_KEYS:
        path = session.results.pop(result_key, None)
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def clear_prepared_cache(session: Any) -> None:
    shutil.rmtree(_cache_root(session), ignore_errors=True)


def invalidate_allokering_cache(session: Any) -> None:
    clear_generated_results(session)
    clear_prepared_cache(session)


def _artifact_path(session: Any, artifact_key: str) -> Path:
    filename = _ARTIFACT_FILENAMES[artifact_key]
    return ensure_cache_root(session) / filename


def _manifest_path(session: Any) -> Path:
    return ensure_cache_root(session) / _MANIFEST_FILENAME


def store_dataframe(session: Any, artifact_key: str, df: Optional[pd.DataFrame]) -> None:
    path = _artifact_path(session, artifact_key)
    if df is None:
        try:
            path.unlink(missing_ok=True)
        except Exception:
            pass
        return
    pd.to_pickle(df, path)


def load_dataframe(session: Any, artifact_key: str) -> Optional[pd.DataFrame]:
    path = _artifact_path(session, artifact_key)
    if not path.exists():
        return None
    loaded = pd.read_pickle(path)
    if isinstance(loaded, pd.DataFrame):
        return loaded
    return None


def save_manifest(session: Any, prepared_results: Iterable[str], meta: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    normalized = [key for key in prepared_results if key in ALLOKERING_RESULT_KEYS]
    payload = {
        "prepared_results": normalized,
        "meta": dict(meta or {}),
    }
    _manifest_path(session).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def load_manifest(session: Any) -> Dict[str, Any]:
    path = _manifest_path(session)
    if not path.exists():
        return {"prepared_results": [], "meta": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"prepared_results": [], "meta": {}}
    return {
        "prepared_results": [
            key for key in list(data.get("prepared_results") or []) if key in ALLOKERING_RESULT_KEYS
        ],
        "meta": dict(data.get("meta") or {}),
    }


def get_prepared_results(session: Any) -> list[str]:
    return list(load_manifest(session).get("prepared_results") or [])
