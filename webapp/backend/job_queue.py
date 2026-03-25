from __future__ import annotations

from typing import Any, Dict, List, Optional

from psycopg.types.json import Jsonb

try:
    from .allokering_result_cache import get_prepared_results
    from .db import (
        DEFAULT_JOB_MAX_ATTEMPTS,
        connect,
        job_lease_expiry,
        new_uuid,
        utcnow,
        work_session_expiry,
    )
except ImportError:
    from allokering_result_cache import get_prepared_results
    from db import (
        DEFAULT_JOB_MAX_ATTEMPTS,
        connect,
        job_lease_expiry,
        new_uuid,
        utcnow,
        work_session_expiry,
    )


ACTIVE_JOB_STATES = {"queued", "running"}


def append_job_log(session_id: str, *, job_id: Optional[str], attempt: int, message: str) -> int:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO job_logs (job_id, session_id, attempt, message, created_at)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id
                """,
                (job_id, session_id, max(1, int(attempt or 1)), str(message), utcnow()),
            )
            row = cur.fetchone()
        conn.commit()
    return int(row["id"])


def clear_session_logs(session_id: str) -> None:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM job_logs WHERE session_id = %s", (session_id,))
        conn.commit()


def fetch_session_logs_since(session_id: str, last_log_id: int, limit: int = 500) -> List[Dict[str, Any]]:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, job_id::text AS job_id, attempt, message
                FROM job_logs
                WHERE session_id = %s
                  AND id > %s
                ORDER BY id ASC
                LIMIT %s
                """,
                (session_id, max(0, int(last_log_id or 0)), int(limit)),
            )
            rows = cur.fetchall()
    return rows


def session_has_active_job(session_id: str) -> bool:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT 1
                FROM jobs
                WHERE session_id = %s
                  AND status IN ('queued', 'running')
                LIMIT 1
                """,
                (session_id,),
            )
            row = cur.fetchone()
    return row is not None


def enqueue_job(session_id: str, job_type: str, payload: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    if session_has_active_job(session_id):
        raise RuntimeError("En körning pågår redan")
    now = utcnow()
    job_id = new_uuid()
    payload = payload or {}
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO jobs (
                    job_id,
                    session_id,
                    job_type,
                    status,
                    payload_json,
                    attempt_count,
                    lease_owner,
                    lease_expires_at,
                    created_at,
                    started_at,
                    finished_at,
                    error_text
                )
                VALUES (%s, %s, %s, 'queued', %s, 0, NULL, NULL, %s, NULL, NULL, NULL)
                """,
                (job_id, session_id, job_type, Jsonb(payload), now),
            )
            cur.execute(
                """
                UPDATE work_sessions
                SET current_job_id = %s,
                    current_job_type = %s,
                    current_job_state = 'queued',
                    running = TRUE,
                    updated_at = %s,
                    expires_at = %s
                WHERE session_id = %s
                """,
                (job_id, job_type, now, work_session_expiry(now), session_id),
            )
        conn.commit()
    return {"job_id": job_id, "status": "queued", "job_type": job_type}


def get_session_status(session_id: str) -> Dict[str, Any]:
    try:
        from .session_store import get_session
    except ImportError:
        from session_store import get_session

    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    running,
                    current_job_id::text AS current_job_id,
                    current_job_type,
                    current_job_state
                FROM work_sessions
                WHERE session_id = %s
                """,
                (session_id,),
            )
            session_row = cur.fetchone()
            cur.execute(
                """
                SELECT result_key
                FROM work_session_results
                WHERE session_id = %s
                ORDER BY result_key
                """,
                (session_id,),
            )
            results = [row["result_key"] for row in cur.fetchall()]
    if not session_row:
        raise RuntimeError("Session saknas")
    session = get_session(session_id)
    return {
        "running": bool(session_row.get("running")),
        "current_job_id": session_row.get("current_job_id"),
        "current_job_type": session_row.get("current_job_type"),
        "current_job_state": session_row.get("current_job_state"),
        "results": results,
        "prepared_results": get_prepared_results(session) if session else [],
    }


def requeue_stale_jobs() -> int:
    now = utcnow()
    touched = 0
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT job_id::text AS job_id, session_id::text AS session_id, job_type, attempt_count
                FROM jobs
                WHERE status = 'running'
                  AND lease_expires_at IS NOT NULL
                  AND lease_expires_at <= %s
                FOR UPDATE
                """,
                (now,),
            )
            rows = cur.fetchall()
            for row in rows:
                touched += 1
                if int(row["attempt_count"] or 0) >= DEFAULT_JOB_MAX_ATTEMPTS:
                    cur.execute(
                        """
                        UPDATE jobs
                        SET status = 'failed',
                            finished_at = %s,
                            lease_owner = NULL,
                            lease_expires_at = NULL,
                            error_text = COALESCE(error_text, %s)
                        WHERE job_id = %s
                        """,
                        (now, "Worker-lease gick ut för många gånger", row["job_id"]),
                    )
                    cur.execute(
                        """
                        UPDATE work_sessions
                        SET current_job_id = %s,
                            current_job_type = %s,
                            current_job_state = 'failed',
                            running = FALSE,
                            updated_at = %s,
                            expires_at = %s
                        WHERE session_id = %s
                        """,
                        (
                            row["job_id"],
                            row["job_type"],
                            now,
                            work_session_expiry(now),
                            row["session_id"],
                        ),
                    )
                    cur.execute(
                        """
                        INSERT INTO job_logs (job_id, session_id, attempt, message, created_at)
                        VALUES (%s, %s, %s, %s, %s)
                        """,
                        (
                            row["job_id"],
                            row["session_id"],
                            int(row["attempt_count"] or 0),
                            "__ERROR__",
                            now,
                        ),
                    )
                else:
                    cur.execute(
                        """
                        UPDATE jobs
                        SET status = 'queued',
                            lease_owner = NULL,
                            lease_expires_at = NULL,
                            error_text = %s
                        WHERE job_id = %s
                        """,
                        ("Worker-lease gick ut. Jobbet köas om.", row["job_id"]),
                    )
                    cur.execute(
                        """
                        UPDATE work_sessions
                        SET current_job_id = %s,
                            current_job_type = %s,
                            current_job_state = 'queued',
                            running = TRUE,
                            updated_at = %s,
                            expires_at = %s
                        WHERE session_id = %s
                        """,
                        (
                            row["job_id"],
                            row["job_type"],
                            now,
                            work_session_expiry(now),
                            row["session_id"],
                        ),
                    )
                    cur.execute(
                        """
                        INSERT INTO job_logs (job_id, session_id, attempt, message, created_at)
                        VALUES (%s, %s, %s, %s, %s)
                        """,
                        (
                            row["job_id"],
                            row["session_id"],
                            int(row["attempt_count"] or 0),
                            "Worker-lease gick ut. Jobbet köas om från början.",
                            now,
                        ),
                    )
        conn.commit()
    return touched


def claim_next_job(worker_id: str) -> Optional[Dict[str, Any]]:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    job_id::text AS job_id,
                    session_id::text AS session_id,
                    job_type,
                    payload_json,
                    attempt_count
                FROM jobs
                WHERE status = 'queued'
                ORDER BY created_at ASC
                FOR UPDATE SKIP LOCKED
                LIMIT 1
                """
            )
            row = cur.fetchone()
            if not row:
                conn.rollback()
                return None
            attempt = int(row.get("attempt_count") or 0) + 1
            cur.execute(
                """
                UPDATE jobs
                SET status = 'running',
                    attempt_count = %s,
                    lease_owner = %s,
                    lease_expires_at = %s,
                    started_at = COALESCE(started_at, %s),
                    error_text = NULL
                WHERE job_id = %s
                """,
                (attempt, worker_id, job_lease_expiry(now), now, row["job_id"]),
            )
            cur.execute(
                """
                UPDATE work_sessions
                SET current_job_id = %s,
                    current_job_type = %s,
                    current_job_state = 'running',
                    running = TRUE,
                    updated_at = %s,
                    expires_at = %s
                WHERE session_id = %s
                """,
                (
                    row["job_id"],
                    row["job_type"],
                    now,
                    work_session_expiry(now),
                    row["session_id"],
                ),
            )
        conn.commit()
    row["attempt_count"] = attempt
    row["status"] = "running"
    return row


def heartbeat_job(job_id: str, worker_id: str) -> bool:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE jobs
                SET lease_expires_at = %s
                WHERE job_id = %s
                  AND status = 'running'
                  AND lease_owner = %s
                RETURNING job_id
                """,
                (job_lease_expiry(now), job_id, worker_id),
            )
            row = cur.fetchone()
        conn.commit()
    return row is not None


def complete_job(session_id: str, job_id: str) -> None:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE jobs
                SET status = 'completed',
                    finished_at = %s,
                    lease_owner = NULL,
                    lease_expires_at = NULL
                WHERE job_id = %s
                """,
                (now, job_id),
            )
            cur.execute(
                """
                UPDATE work_sessions
                SET current_job_id = %s,
                    current_job_state = 'completed',
                    running = FALSE,
                    updated_at = %s,
                    expires_at = %s
                WHERE session_id = %s
                """,
                (job_id, now, work_session_expiry(now), session_id),
            )
        conn.commit()


def fail_job(session_id: str, job_id: str, error_text: str) -> None:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE jobs
                SET status = 'failed',
                    finished_at = %s,
                    lease_owner = NULL,
                    lease_expires_at = NULL,
                    error_text = %s
                WHERE job_id = %s
                """,
                (now, error_text, job_id),
            )
            cur.execute(
                """
                UPDATE work_sessions
                SET current_job_id = %s,
                    current_job_state = 'failed',
                    running = FALSE,
                    updated_at = %s,
                    expires_at = %s
                WHERE session_id = %s
                """,
                (job_id, now, work_session_expiry(now), session_id),
            )
        conn.commit()
