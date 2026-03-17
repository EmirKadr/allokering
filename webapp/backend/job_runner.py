from __future__ import annotations

import os
from typing import Any, Dict, Optional

try:
    from .session_store import SessionData, get_session
except ImportError:
    from session_store import SessionData, get_session


JOB_RESULT_KEYS = {
    "allokering": ("allokerade", "nearmiss", "pallplatser", "refill"),
    "hib-koppling": ("hib-koppling",),
    "orderkontroll": ("orderkontroll",),
    "dispatchkontroll": ("dispatchkontroll",),
    "eftersok": ("eftersok",),
}


def _clear_result_key(session: SessionData, result_key: str) -> None:
    path = session.results.pop(result_key, None)
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass


def prepare_job_artifacts(session: SessionData, job_type: str) -> None:
    for result_key in JOB_RESULT_KEYS.get(job_type, ()):
        _clear_result_key(session, result_key)
    if job_type == "allokering":
        session.ordersaldo_list1 = []
        session.ordersaldo_list2 = []


async def run_job(job: Dict[str, Any]) -> Optional[str]:
    try:
        from .main import (
            _job_allokering,
            _job_dispatchkontroll,
            _job_eftersok,
            _job_hib_koppling,
            _job_orderkontroll,
        )
    except ImportError:
        from main import (
            _job_allokering,
            _job_dispatchkontroll,
            _job_eftersok,
            _job_hib_koppling,
            _job_orderkontroll,
        )

    session = get_session(job["session_id"])
    if not session:
        raise RuntimeError("Session saknas")

    session.bind_job_log(job["job_id"], int(job.get("attempt_count") or 1))
    prepare_job_artifacts(session, job["job_type"])
    session.log_queue.put_nowait(f"Worker startade jobb: {job['job_type']}")

    payload = job.get("payload_json") or {}
    job_type = job["job_type"]
    if job_type == "allokering":
        await _job_allokering(session)
    elif job_type == "hib-koppling":
        await _job_hib_koppling(session)
    elif job_type == "orderkontroll":
        await _job_orderkontroll(session)
    elif job_type == "dispatchkontroll":
        await _job_dispatchkontroll(session)
    elif job_type == "eftersok":
        await _job_eftersok(
            session,
            str(payload.get("purchase", "")),
            str(payload.get("article", "")),
        )
    else:
        raise RuntimeError(f"Okänd job_typ: {job_type}")

    return session.log_queue.last_terminal_marker
