"""
Login e usuários — senhas só como hash PBKDF2 em data/users.json.

Na 1ª execução cria o admin padrão se o arquivo não existir.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
from datetime import datetime

APP_USER_VERSION = 28
# Admin inicial — só usado quando users.json ainda não existe
DEFAULT_EMAIL = "luizhcastro06@gmail.com"
DEFAULT_PASSWORD = "7355608"
DEFAULT_NAME = "Luiz Henrique"

PBKDF2_ITERATIONS = 260000
HASH_PREFIX = "pbkdf2:"
LEGACY_SALT = "logsystem_v7_salt_2025"


def hash_password(password: str) -> str:
    """Gera hash PBKDF2 com salt aleatório."""
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS
    )
    return (
        f"{HASH_PREFIX}{PBKDF2_ITERATIONS}$"
        f"{base64.urlsafe_b64encode(salt).decode().rstrip('=')}$"
        f"{base64.urlsafe_b64encode(digest).decode().rstrip('=')}"
    )


def verify_password(password: str, stored_hash: str) -> bool:
    if not stored_hash:
        return False
    if stored_hash.startswith(HASH_PREFIX):
        try:
            body = stored_hash[len(HASH_PREFIX) :]
            iterations_str, salt_b64, hash_b64 = body.split("$", 2)
            iterations = int(iterations_str)
            pad = lambda s: s + "=" * (-len(s) % 4)
            salt = base64.urlsafe_b64decode(pad(salt_b64))
            expected = base64.urlsafe_b64decode(pad(hash_b64))
            digest = hashlib.pbkdf2_hmac(
                "sha256", password.encode("utf-8"), salt, iterations
            )
            return secrets.compare_digest(digest, expected)
        except Exception:
            return False
    legacy = hashlib.sha256((LEGACY_SALT + password).encode()).hexdigest()
    return secrets.compare_digest(legacy, stored_hash)


def _default_user() -> dict:
    return {
        "name": DEFAULT_NAME,
        "password_hash": hash_password(DEFAULT_PASSWORD),
        "role": "admin",
        "created": datetime.now().isoformat(),
    }


def load_users(users_file: str) -> dict:
    """Carrega users.json; recria admin se versão mudou ou hash legado."""
    users = {}
    if os.path.exists(users_file):
        try:
            with open(users_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                users = raw
        except Exception:
            users = {}

    email = DEFAULT_EMAIL.strip().lower()
    meta_key = "_meta"
    meta = users.get(meta_key, {}) if isinstance(users.get(meta_key), dict) else {}
    needs_reset = (
        email not in users
        or meta.get("version") != APP_USER_VERSION
        or not str(users.get(email, {}).get("password_hash", "")).startswith(HASH_PREFIX)
    )

    if needs_reset:
        users[email] = _default_user()
        users[meta_key] = {"version": APP_USER_VERSION}
        save_users(users_file, users)

    if meta_key in users:
        del users[meta_key]

    return users


def save_users(users_file: str, users: dict):
    os.makedirs(os.path.dirname(users_file) or ".", exist_ok=True)
    payload = {k: v for k, v in users.items() if k != "_meta"}
    payload["_meta"] = {"version": APP_USER_VERSION}
    with open(users_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
