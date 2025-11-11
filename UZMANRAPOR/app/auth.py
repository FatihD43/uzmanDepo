from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional
import json

from app import storage


@dataclass(frozen=True)
class User:
    """Basit kullanıcı modeli (parolasız, yalnızca yetki bilgisi)."""

    username: str
    permissions: frozenset[str]

    def has_permission(self, perm: str) -> bool:
        if not perm:
            return True
        normalized = perm.strip().lower()
        perms = {p.strip().lower() for p in self.permissions}
        return normalized in perms or "admin" in perms

    @classmethod
    def anonymous(cls) -> "User":
        return cls(username="Anonim", permissions=frozenset({"read", "write", "admin"}))


def _build_user(record: dict) -> Optional[User]:
    if not isinstance(record, dict):
        return None
    username = str(record.get("username", "")).strip()
    if not username:
        return None
    perms_raw: Iterable[str] = record.get("permissions", []) or []
    perms = frozenset(str(p).strip() for p in perms_raw if str(p).strip())
    if not perms:
        perms = frozenset({"read"})
    return User(username=username, permissions=perms)


def authenticate(username: str, password: str) -> Optional[User]:
    """Kullanıcı adı/şifre doğrular; başarılıysa User döndürür."""
    username = (username or "").strip()
    if not username:
        return None

    password = password or ""

    load_users_fn = getattr(storage, "load_users", None)
    if callable(load_users_fn):
        records = load_users_fn()
    else:
        records = _load_users_fallback()
    for rec in records:
        rec_username = str(rec.get("username", "")).strip()
        if rec_username.lower() != username.lower():
            continue

        salt = rec.get("salt")
        password_hash = rec.get("password_hash")
        password_plain = rec.get("password")

        valid = False
        if password_hash:
            if salt:
                valid = _check_password(password, str(salt), str(password_hash))
            else:
                # Bazı eski kayıtlar yalnızca şifreyi düz metin saklıyor olabilir.
                valid = password == str(password_hash)
        elif password_plain is not None:
            valid = password == str(password_plain)

        if not valid:
            continue

        user = _build_user(rec)
        if user is None:
            continue
        return User(username=user.username, permissions=user.permissions)
    return None


def list_users() -> list[User]:
    """users.json içindeki tüm kullanıcıları döndürür."""
    out: list[User] = []
    load_users_fn = getattr(storage, "load_users", None)
    if callable(load_users_fn):
        records = load_users_fn()
    else:
        records = _load_users_fallback()

    for rec in records:
        user = _build_user(rec)
        if user:
            out.append(user)
    return out

def _check_password(candidate: str, salt: str, expected_hash: str) -> bool:
    """Compare the given password against the stored hash, tolerating missing helpers."""

    hash_fn = getattr(storage, "hash_password", None)
    if callable(hash_fn):
        try:
            return hash_fn(candidate, salt) == expected_hash
        except Exception:
            return False
    # Fallback: treat stored value as plain text if hashing helper is unavailable.
    return candidate == expected_hash


def _load_users_fallback() -> list[dict]:
    """Minimal users.json reader used when storage.load_users() is absent."""

    ensure_fn = getattr(storage, "ensure_user_db", None)
    if callable(ensure_fn):
        try:
            ensure_fn()
        except Exception:
            pass

    user_db_path = getattr(storage, "USERS_DB_PATH", None)
    if user_db_path is None:
        user_db_path = Path.home() / ".uzman_rapor" / "users.json"
    else:
        user_db_path = Path(user_db_path)

    try:
        with open(user_db_path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except Exception:
        return []

    if isinstance(payload, dict):
        users = payload.get("users")
        if isinstance(users, list):
            return users
    return []