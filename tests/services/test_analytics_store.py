from __future__ import annotations

from pathlib import Path

from analytics_store import (
    analytics_event_file,
    append_analytics_event,
    load_analytics_events,
    resolve_analytics_storage_dir,
)


def test_resolve_analytics_storage_dir_uses_explicit_path(tmp_path: Path) -> None:
    resolved = resolve_analytics_storage_dir(str(tmp_path))
    assert resolved == tmp_path


def test_append_and_load_analytics_events(tmp_path: Path) -> None:
    payload = {
        "event": "app_started",
        "timestamp": "2026-05-06T12:00:00+00:00",
        "properties": {
            "install_id": "abc123",
            "app_version": "12.1.3",
        },
    }

    written_path = append_analytics_event(tmp_path, "abc123", payload)

    assert written_path == analytics_event_file(tmp_path, "abc123")
    assert written_path.exists()

    loaded = load_analytics_events(tmp_path)

    assert len(loaded) == 1
    assert loaded[0]["event"] == "app_started"
    assert loaded[0]["properties"]["install_id"] == "abc123"
    assert loaded[0]["_source_file"].endswith("abc123.jsonl")
