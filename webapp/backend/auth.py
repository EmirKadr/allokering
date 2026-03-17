"""
auth.py – Autentisering och behörighetshantering.
Lagrar användare i users.json, sessioner i minnet.
"""

from __future__ import annotations

import json
import threading
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import HTTPException

# ---------------------------------------------------------------------------
# Konstanter
# ---------------------------------------------------------------------------

ADMIN_EMAIL = "emir.k@nowaste.se"
ADMIN_USERNAME = "emikad"
ADMIN_CODE = "0012"

USERS_FILE = Path(__file__).parent / "users.json"

VALID_LISTS = ("gg", "mg", "both")

# ---------------------------------------------------------------------------
# In-memory auth sessions:  token -> user info dict
# ---------------------------------------------------------------------------

_auth_sessions: Dict[str, dict] = {}
_users_lock = threading.Lock()

# ---------------------------------------------------------------------------
# JSON-fil I/O
# ---------------------------------------------------------------------------


def load_users() -> dict:
    with _users_lock:
        if not USERS_FILE.exists():
            default = {
                "lists": {
                    "gg":   {"label": "Granngården LKON",   "users": []},
                    "mg":   {"label": "Mestergruppen LKON", "users": []},
                    "both": {"label": "Båda",               "users": []},
                }
            }
            USERS_FILE.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
            return default
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))


def save_users(data: dict) -> None:
    with _users_lock:
        USERS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Uppslagning
# ---------------------------------------------------------------------------


def _normalize(s: str) -> str:
    return s.strip().lower()


def find_user(identifier: str) -> Optional[dict]:
    """Slå upp en användare via e-post eller användarnamn.

    Returnerar dict med nycklar: username, email, role, lists
    eller None om identifieraren inte finns i någon lista.
    """
    ident = _normalize(identifier)

    # Kolla admin
    if ident in (_normalize(ADMIN_EMAIL), _normalize(ADMIN_USERNAME)):
        return {
            "username": ADMIN_USERNAME,
            "email": ADMIN_EMAIL,
            "role": "admin",
            "lists": list(VALID_LISTS),
        }

    # Kolla listor
    data = load_users()
    member_lists: List[str] = []
    found_username: Optional[str] = None
    found_email: Optional[str] = None

    for list_key in VALID_LISTS:
        for u in data["lists"].get(list_key, {}).get("users", []):
            if _normalize(u.get("username", "")) == ident or _normalize(u.get("email", "")) == ident:
                member_lists.append(list_key)
                found_username = found_username or u.get("username")
                found_email = found_email or u.get("email")
                break  # hittad i denna lista, kolla nästa

    if not member_lists:
        return None

    return {
        "username": found_username or ident,
        "email": found_email or "",
        "role": "user",
        "lists": member_lists,
    }


# ---------------------------------------------------------------------------
# Auth-sessioner
# ---------------------------------------------------------------------------


def create_auth_session(user_info: dict) -> str:
    token = str(uuid.uuid4())
    _auth_sessions[token] = {**user_info, "token": token}
    return token


def get_auth_session(token: str) -> Optional[dict]:
    return _auth_sessions.get(token)


def delete_auth_session(token: str) -> None:
    _auth_sessions.pop(token, None)


# ---------------------------------------------------------------------------
# Behörighetskontroll (för endpoints)
# ---------------------------------------------------------------------------


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
    """Extrahera token från Authorization-header eller query-param."""
    if authorization and authorization.startswith("Bearer "):
        return authorization[7:]
    return token_param
