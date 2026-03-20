from __future__ import annotations

import asyncio
import os
import socket
import uuid

try:
    from .db import init_db
    from .job_queue import (
        claim_next_job,
        requeue_stale_jobs,
    )
    from .worker_runtime import process_claimed_job
except ImportError:
    from db import init_db
    from job_queue import (
        claim_next_job,
        requeue_stale_jobs,
    )
    from worker_runtime import process_claimed_job


POLL_INTERVAL_SECONDS = float(os.environ.get("ALLOK_WORKER_POLL_SECONDS", "1.0"))


def _worker_id() -> str:
    return os.environ.get("ALLOK_WORKER_ID") or f"{socket.gethostname()}-{uuid.uuid4().hex[:8]}"


async def worker_loop() -> None:
    init_db()
    worker_id = _worker_id()
    while True:
        await asyncio.to_thread(requeue_stale_jobs)
        job = await asyncio.to_thread(claim_next_job, worker_id)
        if not job:
            await asyncio.sleep(POLL_INTERVAL_SECONDS)
            continue
        await process_claimed_job(job, worker_id)


def main() -> None:
    asyncio.run(worker_loop())


if __name__ == "__main__":
    main()
