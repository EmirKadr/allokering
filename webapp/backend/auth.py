"""
auth.py - Autentisering och behörighetshantering.
Lagrar användare i users.json och auth-sessioner i PostgreSQL.
"""

from __future__ import annotations

import hashlib
import json
import threading
import uuid
from pathlib import Path
from typing import Dict, List, Optional

from fastapi import HTTPException
from psycopg.types.json import Jsonb

try:
    from .db import auth_expiry, connect, utcnow
except ImportError:
    from db import auth_expiry, connect, utcnow


ADMIN_EMAIL = "emir.k@nowaste.se"
ADMIN_USERNAME = "emikad"
ADMIN_CODE = "0012"

USERS_FILE = Path(__file__).parent / "users.json"

VALID_LISTS = ("gg",)

_users_lock = threading.Lock()


def load_users() -> dict:
    with _users_lock:
        if not USERS_FILE.exists():
            default = {
                "lists": {
                    "gg": {"label": "Granngården LKON", "users": []},
                    "mg": {"label": "Mestergruppen LKON", "users": []},
                    "both": {"label": "Båda", "users": []},
                }
            }
            USERS_FILE.write_text(
                json.dumps(default, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            return default
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))


def save_users(data: dict) -> None:
    with _users_lock:
        USERS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _normalize(s: str) -> str:
    return s.strip().lower()


def _token_hash(token: str) -> str:
    return hashlib.sha256((token or "").encode("utf-8")).hexdigest()


def cleanup_expired_auth_sessions() -> None:
    now = utcnow()
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                DELETE FROM auth_sessions
                WHERE revoked_at IS NOT NULL OR expires_at <= %s
                """,
                (now,),
            )
        conn.commit()


def find_user(identifier: str) -> Optional[dict]:
    ident = _normalize(identifier)

    if ident in (_normalize(ADMIN_EMAIL), _normalize(ADMIN_USERNAME)):
        return {
            "username": ADMIN_USERNAME,
            "email": ADMIN_EMAIL,
            "role": "admin",
            "lists": list(VALID_LISTS),
        }

    data = load_users()
    member_lists: List[str] = []
    found_username: Optional[str] = None
    found_email: Optional[str] = None

    for list_key in VALID_LISTS:
        for user in data["lists"].get(list_key, {}).get("users", []):
            if _normalize(user.get("username", "")) == ident or _normalize(user.get("email", "")) == ident:
                member_lists.append(list_key)
                found_username = found_username or user.get("username")
                found_email = found_email or user.get("email")
                break

    if not member_lists:
        return None

    return {
        "username": found_username or ident,
        "email": found_email or "",
        "role": "user",
        "lists": member_lists,
    }


def create_auth_session(user_info: dict) -> str:
    cleanup_expired_auth_sessions()
    token = str(uuid.uuid4())
    now = utcnow()
    expires_at = auth_expiry(now)
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO auth_sessions (
                    token_hash,
                    username,
                    email,
                    role,
                    lists_json,
                    created_at,
                    last_seen_at,
                    expires_at,
                    revoked_at
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NULL)
                """,
                (
                    _token_hash(token),
                    user_info.get("username", ""),
                    user_info.get("email", ""),
                    user_info.get("role", "user"),
                    Jsonb(list(user_info.get("lists", []))),
                    now,
                    now,
                    expires_at,
                ),
            )
        conn.commit()
    return token


def get_auth_session(token: str) -> Optional[dict]:
    if not token:
        return None
    cleanup_expired_auth_sessions()
    token_hash = _token_hash(token)
    now = utcnow()
    expires_at = auth_expiry(now)
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT username, email, role, lists_json
                FROM auth_sessions
                WHERE token_hash = %s
                  AND revoked_at IS NULL
                  AND expires_at > %s
                """,
                (token_hash, now),
            )
            row = cur.fetchone()
            if not row:
                conn.rollback()
                return None
            cur.execute(
                """
                UPDATE auth_sessions
                SET last_seen_at = %s,
                    expires_at = %s
                WHERE token_hash = %s
                """,
                (now, expires_at, token_hash),
            )
        conn.commit()
    return {
        "username": row["username"],
        "email": row.get("email", ""),
        "role": row["role"],
        "lists": list(row.get("lists_json") or []),
        "token": token,
    }


def delete_auth_session(token: str) -> None:
    if not token:
        return
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE auth_sessions
                SET revoked_at = %s
                WHERE token_hash = %s
                """,
                (utcnow(), _token_hash(token)),
            )
        conn.commit()


def require_auth(token: Optional[str]) -> dict:
    if not token:
        raise HTTPException(status_code=401, detail="Ej inloggad")
    session = get_auth_session(token)
    if not session:
        raise HTTPException(status_code=401, detail="Ogiltig session")
    return session


def require_admin(token: Optional[str]) -> dict:
    user = require_auth(token)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Admin-behörighet krävs")
    return user


def extract_token(authorization: Optional[str] = None, token_param: Optional[str] = None) -> Optional[str]:
    if authorization and authorization.startswith("Bearer "):
        return authorization[7:]
    return token_param
