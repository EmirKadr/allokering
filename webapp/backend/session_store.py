from __future__ import annotations

import shutil
import threading
from collections.abc import Iterator, MutableMapping
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Optional, Tuple

from psycopg.types.json import Jsonb

try:
    from .db import (
        connect,
        get_state_dir,
        new_uuid,
        relativize_session_path,
        resolve_session_path,
        session_results_dir,
        session_root_dir,
        session_uploads_dir,
        utcnow,
        work_session_expiry,
    )
except ImportError:
    from db import (
        connect,
        get_state_dir,
        new_uuid,
        relativize_session_path,
        resolve_session_path,
        session_results_dir,
        session_root_dir,
        session_uploads_dir,
        utcnow,
        work_session_expiry,
    )


_runtime_lock = threading.Lock()
_runtime_filter_cache: Dict[str, Dict[str, Tuple[Tuple[int, int], Dict[str, list]]]] = {}


def _normalize_filter_values(raw: Optional[list]) -> Optional[list]:
    if raw is None:
        return None
    return list(raw or [])


def _normalize_filters(value: Optional[Dict[str, list]]) -> Dict[str, Optional[list]]:
    data = value or {}
    return {
        "bolag": _normalize_filter_values(data.get("bolag")),
        "ordertyp": _normalize_filter_values(data.get("ordertyp")),
    }


class SessionLogProxy:
    def __init__(self, session: "SessionData"):
        self.session = session
        self.job_id: Optional[str] = None
        self.attempt: int = 1
        self.last_terminal_marker: Optional[str] = None

    def bind(self, job_id: Optional[str], attempt: int = 1) -> None:
        self.job_id = job_id
        self.attempt = max(1, int(attempt or 1))
        self.last_terminal_marker = None

    def put_nowait(self, message: str) -> None:
        try:
            from .job_queue import append_job_log
        except ImportError:
            from job_queue import append_job_log

        if message in {"__DONE__", "__ERROR__"}:
            self.last_terminal_marker = message
        append_job_log(
            self.session.session_id,
            job_id=self.job_id,
            attempt=self.attempt,
            message=str(message),
        )


class PersistentPathMapping(MutableMapping[str, str]):
    def __init__(self, session: "SessionData", *, table_kind: str):
        self.session = session
        self.table_kind = table_kind
        self._cache: Optional[Dict[str, str]] = None

    @property
    def _table_name(self) -> str:
        return "work_session_files" if self.table_kind == "file" else "work_session_results"

    @property
    def _key_column(self) -> str:
        return "file_key" if self.table_kind == "file" else "result_key"

    def _load(self) -> Dict[str, str]:
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT {self._key_column} AS item_key, path
                    FROM {self._table_name}
                    WHERE session_id = %s
                    ORDER BY {self._key_column}
                    """,
                    (self.session.session_id,),
                )
                rows = cur.fetchall()
        self._cache = {
            row["item_key"]: resolve_session_path(self.session.session_id, row["path"])
            for row in rows
        }
        return dict(self._cache)

    def _data(self) -> Dict[str, str]:
        return self._load() if self._cache is None else self._cache

    def __getitem__(self, key: str) -> str:
        return self._data()[key]

    def __setitem__(self, key: str, value: str) -> None:
        relative_path = relativize_session_path(self.session.session_id, value)
        filename = value.split("\\")[-1].split("/")[-1]
        with connect() as conn:
            with conn.cursor() as cur:
                if self.table_kind == "file":
                    cur.execute(
                        """
                        INSERT INTO work_session_files (session_id, file_key, path, original_filename, uploaded_at)
                        VALUES (%s, %s, %s, %s, %s)
                        ON CONFLICT (session_id, file_key)
                        DO UPDATE SET
                            path = EXCLUDED.path,
                            original_filename = EXCLUDED.original_filename,
                            uploaded_at = EXCLUDED.uploaded_at
                        """,
                        (self.session.session_id, key, relative_path, filename, utcnow()),
                    )
                else:
                    cur.execute(
                        """
                        INSERT INTO work_session_results (session_id, result_key, path, created_at)
                        VALUES (%s, %s, %s, %s)
                        ON CONFLICT (session_id, result_key)
                        DO UPDATE SET
                            path = EXCLUDED.path,
                            created_at = EXCLUDED.created_at
                        """,
                        (self.session.session_id, key, relative_path, utcnow()),
                    )
            conn.commit()
        self._data()[key] = value
        self.session.touch()

    def __delitem__(self, key: str) -> None:
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    f"DELETE FROM {self._table_name} WHERE session_id = %s AND {self._key_column} = %s",
                    (self.session.session_id, key),
                )
            conn.commit()
        self._data().pop(key, None)
        self.session.touch()

    def __iter__(self) -> Iterator[str]:
        return iter(dict(self._data()))

    def __len__(self) -> int:
        return len(self._data())

    def pop(self, key: str, default=None):
        data = self._data()
        if key not in data:
            return default
        value = data[key]
        del self[key]
        return value

    def clear(self) -> None:
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    f"DELETE FROM {self._table_name} WHERE session_id = %s",
                    (self.session.session_id,),
                )
            conn.commit()
        self._cache = {}
        self.session.touch()

    def invalidate(self) -> None:
        self._cache = None


@dataclass
class SessionData:
    session_id: str
    owner_username: str
    created_at: datetime
    _active_filters: Dict[str, Optional[list]] = field(default_factory=dict)
    _ordersaldo_list1: list = field(default_factory=list)
    _ordersaldo_list2: list = field(default_factory=list)
    _running: bool = False
    _current_job_id: Optional[str] = None
    _current_job_type: Optional[str] = None
    _current_job_state: Optional[str] = None

    def __post_init__(self) -> None:
        self.temp_dir = str(session_root_dir(self.session_id))
        self.uploads_dir = str(session_uploads_dir(self.session_id))
        self.results_dir = str(session_results_dir(self.session_id))
        self.files = PersistentPathMapping(self, table_kind="file")
        self.results = PersistentPathMapping(self, table_kind="result")
        with _runtime_lock:
            self.filter_scan_cache = _runtime_filter_cache.setdefault(self.session_id, {})
        self.log_queue = SessionLogProxy(self)

    def touch(self) -> None:
        now = utcnow()
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (now, work_session_expiry(now), self.session_id),
                )
            conn.commit()

    @property
    def active_filters(self) -> Dict[str, Optional[list]]:
        return dict(self._active_filters)

    @active_filters.setter
    def active_filters(self, value: Dict[str, Optional[list]]) -> None:
        normalized = _normalize_filters(value)
        self._active_filters = normalized
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET active_filters_json = %s,
                        updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (Jsonb(normalized), utcnow(), work_session_expiry(), self.session_id),
                )
            conn.commit()

    @property
    def ordersaldo_list1(self) -> list:
        return list(self._ordersaldo_list1)

    @ordersaldo_list1.setter
    def ordersaldo_list1(self, value: list) -> None:
        payload = list(value or [])
        self._ordersaldo_list1 = payload
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET ordersaldo_list1_json = %s,
                        updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (Jsonb(payload), utcnow(), work_session_expiry(), self.session_id),
                )
            conn.commit()

    @property
    def ordersaldo_list2(self) -> list:
        return list(self._ordersaldo_list2)

    @ordersaldo_list2.setter
    def ordersaldo_list2(self, value: list) -> None:
        payload = list(value or [])
        self._ordersaldo_list2 = payload
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET ordersaldo_list2_json = %s,
                        updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (Jsonb(payload), utcnow(), work_session_expiry(), self.session_id),
                )
            conn.commit()

    @property
    def running(self) -> bool:
        return bool(self._running)

    @running.setter
    def running(self, value: bool) -> None:
        running = bool(value)
        self._running = running
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET running = %s,
                        updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (running, utcnow(), work_session_expiry(), self.session_id),
                )
            conn.commit()

    @property
    def current_job_id(self) -> Optional[str]:
        return self._current_job_id

    @current_job_id.setter
    def current_job_id(self, value: Optional[str]) -> None:
        self._current_job_id = value
        self._update_job_state()

    @property
    def current_job_type(self) -> Optional[str]:
        return self._current_job_type

    @current_job_type.setter
    def current_job_type(self, value: Optional[str]) -> None:
        self._current_job_type = value
        self._update_job_state()

    @property
    def current_job_state(self) -> Optional[str]:
        return self._current_job_state

    @current_job_state.setter
    def current_job_state(self, value: Optional[str]) -> None:
        self._current_job_state = value
        self._update_job_state()

    def _update_job_state(self) -> None:
        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE work_sessions
                    SET current_job_id = %s,
                        current_job_type = %s,
                        current_job_state = %s,
                        updated_at = %s,
                        expires_at = %s
                    WHERE session_id = %s
                    """,
                    (
                        self._current_job_id,
                        self._current_job_type,
                        self._current_job_state,
                        utcnow(),
                        work_session_expiry(),
                        self.session_id,
                    ),
                )
            conn.commit()

    def bind_job_log(self, job_id: Optional[str], attempt: int = 1) -> None:
        self.log_queue.bind(job_id, attempt)

    def refresh_paths(self) -> None:
        self.files.invalidate()
        self.results.invalidate()


def _cleanup_session_artifacts(session_id: str) -> None:
    shutil.rmtree(get_state_dir() / "sessions" / session_id, ignore_errors=True)
    with _runtime_lock:
        _runtime_filter_cache.pop(session_id, None)


def cleanup_expired_sessions() -> None:
    now = utcnow()
    expired_ids: list[str] = []
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT session_id::text
                FROM work_sessions
                WHERE expires_at <= %s
                """,
                (now,),
            )
            expired_ids = [row["session_id"] for row in cur.fetchall()]
            for sid in expired_ids:
                cur.execute("DELETE FROM work_sessions WHERE session_id = %s", (sid,))
        conn.commit()
    for sid in expired_ids:
        _cleanup_session_artifacts(sid)


def create_session(owner_username: str) -> SessionData:
    cleanup_expired_sessions()
    sid = new_uuid()
    now = utcnow()
    session_root_dir(sid)
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO work_sessions (
                    session_id,
                    owner_username,
                    active_filters_json,
                    ordersaldo_list1_json,
                    ordersaldo_list2_json,
                    current_job_id,
                    current_job_type,
                    current_job_state,
                    running,
                    created_at,
                    updated_at,
                    expires_at
                )
                VALUES (%s, %s, %s, %s, %s, NULL, NULL, NULL, FALSE, %s, %s, %s)
                """,
                (
                    sid,
                    owner_username,
                    Jsonb({"bolag": None, "ordertyp": None}),
                    Jsonb([]),
                    Jsonb([]),
                    now,
                    now,
                    work_session_expiry(now),
                ),
            )
        conn.commit()
    return get_session(sid)


def get_session(sid: str) -> Optional[SessionData]:
    if not sid:
        return None
    cleanup_expired_sessions()
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    session_id::text,
                    owner_username,
                    active_filters_json,
                    ordersaldo_list1_json,
                    ordersaldo_list2_json,
                    current_job_id::text AS current_job_id,
                    current_job_type,
                    current_job_state,
                    running,
                    created_at
                FROM work_sessions
                WHERE session_id = %s
                  AND expires_at > %s
                """,
                (sid, now),
            )
            row = cur.fetchone()
            if not row:
                conn.rollback()
                return None
            cur.execute(
                """
                UPDATE work_sessions
                SET updated_at = %s,
                    expires_at = %s
                WHERE session_id = %s
                """,
                (now, work_session_expiry(now), sid),
            )
        conn.commit()
    return SessionData(
        session_id=row["session_id"],
        owner_username=row["owner_username"],
        created_at=row["created_at"],
        _active_filters=_normalize_filters(row.get("active_filters_json") or {}),
        _ordersaldo_list1=list(row.get("ordersaldo_list1_json") or []),
        _ordersaldo_list2=list(row.get("ordersaldo_list2_json") or []),
        _running=bool(row.get("running")),
        _current_job_id=row.get("current_job_id"),
        _current_job_type=row.get("current_job_type"),
        _current_job_state=row.get("current_job_state"),
    )


def delete_session(sid: str) -> None:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM work_sessions WHERE session_id = %s", (sid,))
        conn.commit()
    _cleanup_session_artifacts(sid)
