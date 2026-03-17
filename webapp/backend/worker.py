from __future__ import annotations

import asyncio
import os
import socket
import uuid

try:
    from .db import DEFAULT_JOB_HEARTBEAT_SECONDS, init_db
    from .job_queue import (
        append_job_log,
        claim_next_job,
        complete_job,
        fail_job,
        heartbeat_job,
        requeue_stale_jobs,
    )
    from .job_runner import run_job
except ImportError:
    from db import DEFAULT_JOB_HEARTBEAT_SECONDS, init_db
    from job_queue import (
        append_job_log,
        claim_next_job,
        complete_job,
        fail_job,
        heartbeat_job,
        requeue_stale_jobs,
    )
    from job_runner import run_job


POLL_INTERVAL_SECONDS = float(os.environ.get("ALLOK_WORKER_POLL_SECONDS", "1.0"))


def _worker_id() -> str:
    return os.environ.get("ALLOK_WORKER_ID") or f"{socket.gethostname()}-{uuid.uuid4().hex[:8]}"


async def _heartbeat_loop(job_id: str, worker_id: str, stop_event: asyncio.Event) -> None:
    while not stop_event.is_set():
        await asyncio.sleep(DEFAULT_JOB_HEARTBEAT_SECONDS)
        if stop_event.is_set():
            break
        ok = await asyncio.to_thread(heartbeat_job, job_id, worker_id)
        if not ok:
            break


async def _process_job(job: dict, worker_id: str) -> None:
    stop_event = asyncio.Event()
    heartbeat_task = asyncio.create_task(_heartbeat_loop(job["job_id"], worker_id, stop_event))
    try:
        terminal_marker = await run_job(job)
        if terminal_marker == "__ERROR__":
            await asyncio.to_thread(fail_job, job["session_id"], job["job_id"], "Jobbet avslutades med fel")
        else:
            await asyncio.to_thread(complete_job, job["session_id"], job["job_id"])
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
    finally:
        stop_event.set()
        await heartbeat_task


async def worker_loop() -> None:
    init_db()
    worker_id = _worker_id()
    while True:
        await asyncio.to_thread(requeue_stale_jobs)
        job = await asyncio.to_thread(claim_next_job, worker_id)
        if not job:
            await asyncio.sleep(POLL_INTERVAL_SECONDS)
            continue
        await _process_job(job, worker_id)


def main() -> None:
    asyncio.run(worker_loop())


if __name__ == "__main__":
    main()
