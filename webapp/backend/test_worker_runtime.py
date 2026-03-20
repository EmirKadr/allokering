from __future__ import annotations

import asyncio
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from backend import worker_runtime  # noqa: E402


class WorkerRuntimeHeartbeatTest(unittest.IsolatedAsyncioTestCase):
    async def test_process_claimed_job_heartbeats_until_completion(self) -> None:
        job = {"job_id": "job-1", "session_id": "session-1", "attempt_count": 1}
        heartbeat_calls: list[tuple[str, str]] = []
        completion_calls: list[tuple[str, str]] = []
        failure_calls: list[tuple[str, str, str]] = []

        async def fake_run_job(_job: dict) -> None:
            await asyncio.sleep(0.03)
            return None

        def fake_heartbeat(job_id: str, worker_id: str) -> bool:
            heartbeat_calls.append((job_id, worker_id))
            return True

        def fake_complete(session_id: str, job_id: str) -> None:
            completion_calls.append((session_id, job_id))

        def fake_fail(session_id: str, job_id: str, message: str) -> None:
            failure_calls.append((session_id, job_id, message))

        with patch.object(worker_runtime, "DEFAULT_JOB_HEARTBEAT_SECONDS", 0.01), \
             patch.object(worker_runtime, "run_job", fake_run_job), \
             patch.object(worker_runtime, "heartbeat_job", fake_heartbeat), \
             patch.object(worker_runtime, "complete_job", fake_complete), \
             patch.object(worker_runtime, "fail_job", fake_fail):
            marker = await worker_runtime.process_claimed_job(job, "worker-1")

        self.assertIsNone(marker)
        self.assertGreaterEqual(len(heartbeat_calls), 1)
        self.assertEqual(completion_calls, [("session-1", "job-1")])
        self.assertEqual(failure_calls, [])
