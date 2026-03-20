from __future__ import annotations

import asyncio
from typing import Any, Dict, Optional

try:
    from .db import DEFAULT_JOB_HEARTBEAT_SECONDS
    from .job_queue import append_job_log, complete_job, fail_job, heartbeat_job
    from .job_runner import run_job
except ImportError:
    from db import DEFAULT_JOB_HEARTBEAT_SECONDS
    from job_queue import append_job_log, complete_job, fail_job, heartbeat_job
    from job_runner import run_job


async def heartbeat_loop(job_id: str, worker_id: str, stop_event: asyncio.Event) -> None:
    while not stop_event.is_set():
        await asyncio.sleep(DEFAULT_JOB_HEARTBEAT_SECONDS)
        if stop_event.is_set():
            break
        ok = await asyncio.to_thread(heartbeat_job, job_id, worker_id)
        if not ok:
            break


async def process_claimed_job(job: Dict[str, Any], worker_id: str) -> Optional[str]:
    stop_event = asyncio.Event()
    heartbeat_task = asyncio.create_task(heartbeat_loop(job["job_id"], worker_id, stop_event))
    try:
        terminal_marker = await run_job(job)
        if terminal_marker == "__ERROR__":
            await asyncio.to_thread(
                fail_job,
                job["session_id"],
                job["job_id"],
                "Jobbet avslutades med fel",
            )
        else:
            await asyncio.to_thread(complete_job, job["session_id"], job["job_id"])
        return terminal_marker
    except Exception as exc:
        await asyncio.to_thread(
            append_job_log,
            job["session_id"],
            job_id=job["job_id"],
            attempt=int(job.get("attempt_count") or 1),
            message=f"FEL: {exc}",
        )
        await asyncio.to_thread(
            append_job_log,
            job["session_id"],
            job_id=job["job_id"],
            attempt=int(job.get("attempt_count") or 1),
            message="__ERROR__",
        )
        await asyncio.to_thread(fail_job, job["session_id"], job["job_id"], str(exc))
        return "__ERROR__"
    finally:
        stop_event.set()
        await heartbeat_task
