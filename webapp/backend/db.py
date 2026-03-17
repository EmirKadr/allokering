from __future__ import annotations

import os
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterator, Optional

import psycopg
from psycopg.rows import dict_row


DEFAULT_AUTH_TTL_HOURS = int(os.environ.get("ALLOK_AUTH_TTL_HOURS", "24"))
DEFAULT_WORK_SESSION_TTL_HOURS = int(os.environ.get("ALLOK_WORK_SESSION_TTL_HOURS", "72"))
DEFAULT_JOB_LEASE_SECONDS = int(os.environ.get("ALLOK_JOB_LEASE_SECONDS", "300"))
DEFAULT_JOB_HEARTBEAT_SECONDS = int(os.environ.get("ALLOK_JOB_HEARTBEAT_SECONDS", "10"))
DEFAULT_JOB_MAX_ATTEMPTS = int(os.environ.get("ALLOK_JOB_MAX_ATTEMPTS", "3"))


def _normalize_database_url(url: str) -> str:
    normalized = (url or "").strip()
    if normalized.startswith("postgres://"):
        normalized = "postgresql://" + normalized[len("postgres://") :]
    return normalized


def get_database_url() -> str:
    url = _normalize_database_url(
        os.environ.get("ALLOK_DATABASE_URL") or os.environ.get("DATABASE_URL") or ""
    )
    if not url:
        raise RuntimeError(
            "Saknar ALLOK_DATABASE_URL eller DATABASE_URL. "
            "Provisionera PostgreSQL och sätt anslutningssträngen innan webapp/worker startas."
        )
    return url


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


def auth_expiry(now: Optional[datetime] = None) -> datetime:
    base = now or utcnow()
    return base + timedelta(hours=DEFAULT_AUTH_TTL_HOURS)


def work_session_expiry(now: Optional[datetime] = None) -> datetime:
    base = now or utcnow()
    return base + timedelta(hours=DEFAULT_WORK_SESSION_TTL_HOURS)


def job_lease_expiry(now: Optional[datetime] = None) -> datetime:
    base = now or utcnow()
    return base + timedelta(seconds=DEFAULT_JOB_LEASE_SECONDS)


def get_state_dir() -> Path:
    default_dir = Path(__file__).resolve().parents[1] / "state"
    path = Path(os.environ.get("ALLOK_STATE_DIR", str(default_dir))).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    (path / "sessions").mkdir(parents=True, exist_ok=True)
    return path


def session_root_dir(session_id: str) -> Path:
    root = get_state_dir() / "sessions" / session_id
    root.mkdir(parents=True, exist_ok=True)
    (root / "uploads").mkdir(parents=True, exist_ok=True)
    (root / "results").mkdir(parents=True, exist_ok=True)
    return root


def session_uploads_dir(session_id: str) -> Path:
    return session_root_dir(session_id) / "uploads"


def session_results_dir(session_id: str) -> Path:
    return session_root_dir(session_id) / "results"


def relativize_session_path(session_id: str, absolute_path: str) -> str:
    root = session_root_dir(session_id)
    return str(Path(absolute_path).resolve().relative_to(root))


def resolve_session_path(session_id: str, relative_path: str) -> str:
    return str((session_root_dir(session_id) / relative_path).resolve())


@contextmanager
def connect(*, autocommit: bool = False) -> Iterator[psycopg.Connection[Any]]:
    conn = psycopg.connect(get_database_url(), row_factory=dict_row, autocommit=autocommit)
    try:
        yield conn
    finally:
        conn.close()


def init_db() -> None:
    get_state_dir()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS auth_sessions (
                    token_hash TEXT PRIMARY KEY,
                    username TEXT NOT NULL,
                    email TEXT NOT NULL DEFAULT '',
                    role TEXT NOT NULL,
                    lists_json JSONB NOT NULL DEFAULT '[]'::jsonb,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    last_seen_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    expires_at TIMESTAMPTZ NOT NULL,
                    revoked_at TIMESTAMPTZ NULL
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS work_sessions (
                    session_id UUID PRIMARY KEY,
                    owner_username TEXT NOT NULL,
                    active_filters_json JSONB NOT NULL DEFAULT '{}'::jsonb,
                    ordersaldo_list1_json JSONB NOT NULL DEFAULT '[]'::jsonb,
                    ordersaldo_list2_json JSONB NOT NULL DEFAULT '[]'::jsonb,
                    current_job_id UUID NULL,
                    current_job_type TEXT NULL,
                    current_job_state TEXT NULL,
                    running BOOLEAN NOT NULL DEFAULT FALSE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    expires_at TIMESTAMPTZ NOT NULL
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS work_session_files (
                    session_id UUID NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
                    file_key TEXT NOT NULL,
                    path TEXT NOT NULL,
                    original_filename TEXT NOT NULL DEFAULT '',
                    uploaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    PRIMARY KEY (session_id, file_key)
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS work_session_results (
                    session_id UUID NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
                    result_key TEXT NOT NULL,
                    path TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    PRIMARY KEY (session_id, result_key)
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    job_id UUID PRIMARY KEY,
                    session_id UUID NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
                    job_type TEXT NOT NULL,
                    status TEXT NOT NULL,
                    payload_json JSONB NOT NULL DEFAULT '{}'::jsonb,
                    attempt_count INTEGER NOT NULL DEFAULT 0,
                    lease_owner TEXT NULL,
                    lease_expires_at TIMESTAMPTZ NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    started_at TIMESTAMPTZ NULL,
                    finished_at TIMESTAMPTZ NULL,
                    error_text TEXT NULL
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS job_logs (
                    id BIGSERIAL PRIMARY KEY,
                    job_id UUID NULL REFERENCES jobs(job_id) ON DELETE CASCADE,
                    session_id UUID NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
                    attempt INTEGER NOT NULL DEFAULT 1,
                    message TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_auth_sessions_expires_at
                ON auth_sessions (expires_at);
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_work_sessions_owner
                ON work_sessions (owner_username);
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_work_sessions_expires_at
                ON work_sessions (expires_at);
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_jobs_status_created_at
                ON jobs (status, created_at);
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_jobs_session_status
                ON jobs (session_id, status);
                """
            )
            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_job_logs_session_id_id
                ON job_logs (session_id, id);
                """
            )
        conn.commit()


def new_uuid() -> str:
    return str(uuid.uuid4())
