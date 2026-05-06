from __future__ import annotations

import json
import os
import re
from pathlib import Path

from app_info import ANALYTICS_LOCAL_STORAGE_DIR, ANALYTICS_STORAGE_DIR_ENV, APP_NAME


def default_analytics_storage_dir() -> Path:
    appdata = os.environ.get("APPDATA") or str(Path.home())
    app_slug = re.sub(r"[^a-z0-9]+", "-", APP_NAME.lower()).strip("-") or "allokering"
    return Path(appdata) / app_slug / "analytics"


def resolve_analytics_storage_dir(raw_value: str | None = None) -> Path:
    candidate = str(raw_value or "").strip()
    if candidate:
        return Path(candidate).expanduser()

    env_candidate = str(os.environ.get(ANALYTICS_STORAGE_DIR_ENV, "")).strip()
    if env_candidate:
        return Path(env_candidate).expanduser()

    configured_candidate = str(ANALYTICS_LOCAL_STORAGE_DIR or "").strip()
    if configured_candidate:
        return Path(configured_candidate).expanduser()

    return default_analytics_storage_dir()


def ensure_analytics_storage_dir(storage_dir: Path) -> Path:
    storage_dir.mkdir(parents=True, exist_ok=True)
    return storage_dir


def _safe_install_id(install_id: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_-]+", "-", str(install_id or "").strip())
    return text or "unknown-installation"


def analytics_event_file(storage_dir: Path, install_id: str) -> Path:
    return ensure_analytics_storage_dir(storage_dir) / f"{_safe_install_id(install_id)}.jsonl"


def append_analytics_event(storage_dir: Path, install_id: str, payload: dict) -> Path:
    path = analytics_event_file(storage_dir, install_id)
    serialized = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    with path.open("a", encoding="utf-8") as handle:
        handle.write(serialized)
        handle.write("\n")
    return path


def iter_analytics_event_files(storage_dir: Path) -> list[Path]:
    if not storage_dir.exists():
        return []
    return sorted(path for path in storage_dir.glob("*.jsonl") if path.is_file())


def load_analytics_events(storage_dir: Path) -> list[dict]:
    rows: list[dict] = []
    for path in iter_analytics_event_files(storage_dir):
        try:
            with path.open("r", encoding="utf-8") as handle:
                for line_number, line in enumerate(handle, start=1):
                    text = line.strip()
                    if not text:
                        continue
                    try:
                        payload = json.loads(text)
                    except json.JSONDecodeError:
                        continue
                    if not isinstance(payload, dict):
                        continue
                    payload["_source_file"] = str(path)
                    payload["_source_line"] = line_number
                    rows.append(payload)
        except OSError:
            continue
    return rows
