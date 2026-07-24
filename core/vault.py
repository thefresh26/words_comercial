"""Cifrado de credenciales con contraseña (PBKDF2 + Fernet)."""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any

from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

VAULT_FILENAME = "credentials.vault.enc"
PBKDF2_ITERATIONS = 480_000
SALT_BYTES = 16


def pkg_dir() -> Path:
    return Path(__file__).resolve().parent


def default_vault_path() -> Path:
    return pkg_dir().parent / VAULT_FILENAME


def _derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    )
    return base64.urlsafe_b64encode(kdf.derive(password.encode("utf-8")))


def encrypt_payload(payload: dict[str, Any], password: str) -> bytes:
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    salt = os.urandom(SALT_BYTES)
    fernet = Fernet(_derive_key(password, salt))
    return salt + fernet.encrypt(raw)


def decrypt_payload(blob: bytes, password: str) -> dict[str, Any]:
    if len(blob) <= SALT_BYTES:
        raise ValueError("Archivo de credenciales inválido o corrupto.")
    salt, token = blob[:SALT_BYTES], blob[SALT_BYTES:]
    fernet = Fernet(_derive_key(password, salt))
    try:
        raw = fernet.decrypt(token)
    except InvalidToken as e:
        raise ValueError("Contraseña incorrecta o archivo dañado.") from e
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Contenido del vault inválido.")
    return data


def read_vault(path: Path | None = None) -> bytes:
    vault = path or default_vault_path()
    if not vault.is_file():
        raise FileNotFoundError(f"No existe {vault.name} en {vault.parent}")
    return vault.read_bytes()


def write_vault(payload: dict[str, Any], password: str, path: Path | None = None) -> Path:
    vault = path or default_vault_path()
    vault.write_bytes(encrypt_payload(payload, password))
    return vault


def database_section_from_local_conf(conf_path: Path) -> dict[str, Any]:
    with open(conf_path, encoding="utf-8") as f:
        cfg = json.load(f)
    db = cfg.get("database")
    if not isinstance(db, dict):
        raise ValueError(f'"{conf_path}" no tiene sección "database".')
    return {
        "user": db.get("user", ""),
        "password": db.get("password", ""),
        "address": db.get("address", "127.0.0.1:5432"),
        "database": db.get("database", ""),
        "parameters": db.get("parameters", ""),
    }


def dsn_from_database_section(db: dict[str, Any]) -> str:
    from urllib.parse import quote_plus

    user = (db.get("user") or "").strip()
    password = "" if db.get("password") is None else str(db.get("password"))
    database = (db.get("database") or "").strip()
    if not user or not database:
        raise ValueError("Faltan user o database en las credenciales.")
    address = (db.get("address") or "127.0.0.1:5432").strip()
    if ":" in address:
        host, port = address.rsplit(":", 1)
        host, port = host.strip(), port.strip()
    else:
        host, port = address, "5432"
    params = (db.get("parameters") or "").strip().lstrip("?&")
    userinfo = f"{quote_plus(user)}:{quote_plus(password)}"
    url = f"postgresql://{userinfo}@{host}:{port}/{quote_plus(database)}"
    if params:
        url = f"{url}?{params}"
    return url
