import asyncio
import uuid
import tempfile
import shutil
from dataclasses import dataclass, field
from typing import Dict, Optional
from datetime import datetime


@dataclass
class SessionData:
    session_id: str
    temp_dir: str
    files: Dict[str, str] = field(default_factory=dict)       # file_key -> abs path
    active_filters: Dict[str, list] = field(default_factory=dict)
    results: Dict[str, str] = field(default_factory=dict)     # result_key -> xlsx path
    log_queue: asyncio.Queue = field(default_factory=asyncio.Queue)
    created_at: datetime = field(default_factory=datetime.utcnow)
    running: bool = False


_sessions: Dict[str, SessionData] = {}


def create_session() -> SessionData:
    sid = str(uuid.uuid4())
    tmp = tempfile.mkdtemp(prefix=f"allok_{sid[:8]}_")
    s = SessionData(session_id=sid, temp_dir=tmp)
    _sessions[sid] = s
    return s


def get_session(sid: str) -> Optional[SessionData]:
    return _sessions.get(sid)


def delete_session(sid: str):
    s = _sessions.pop(sid, None)
    if s:
        shutil.rmtree(s.temp_dir, ignore_errors=True)
