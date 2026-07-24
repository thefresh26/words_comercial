#!/usr/bin/env python3
"""
Arranca la interfaz web con un solo comando (venv, deps, credenciales cifradas).

  python run.py

Variables opcionales:
  INFORME_VAULT_PASSWORD  — evita el prompt de contraseña del vault
  INFORME_PORT            — puerto Streamlit (default 8501)
"""

from __future__ import annotations

import getpass
import os
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

PKG = Path(__file__).resolve().parent
CORE_DIR = PKG / "core"
REQUIREMENTS = PKG / "requirements.txt"
APP = CORE_DIR / "app.py"
VENV = PKG / ".venv"


def _say(msg: str = "") -> None:
    print(msg, flush=True)


def _prompt_password(prompt: str) -> str:
    """Prompt de contraseña; en Windows CMD evita bloqueos de getpass."""
    if sys.platform == "win32":
        try:
            import msvcrt

            print(prompt, end="", flush=True)
            chars: list[str] = []
            while True:
                ch = msvcrt.getwch()
                if ch in ("\r", "\n"):
                    _say()
                    break
                if ch == "\x03":
                    raise KeyboardInterrupt
                if ch == "\x08":
                    if chars:
                        chars.pop()
                        sys.stdout.write("\b \b")
                        sys.stdout.flush()
                    continue
                chars.append(ch)
                sys.stdout.write("*")
                sys.stdout.flush()
            return "".join(chars)
        except (ImportError, OSError):
            pass
    return getpass.getpass(prompt)


def _reexec_in_venv(py: Path) -> None:
    env = os.environ.copy()
    env["INFORME_VENV_BOOTSTRAPPED"] = "1"
    script = str(Path(__file__).resolve())
    args = [str(py), script, *sys.argv[1:]]
    if sys.platform == "win32":
        raise SystemExit(subprocess.call(args, cwd=PKG, env=env))
    os.execve(str(py), args, env)


def venv_python() -> Path:
    if sys.platform == "win32":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def ensure_venv() -> Path:
    py = venv_python()
    if not py.is_file():
        _say("Creando entorno virtual (.venv)…")
        subprocess.run([sys.executable, "-m", "venv", str(VENV)], check=True, cwd=PKG)
    _say("Comprobando dependencias (la primera vez puede tardar 1–2 min)…")
    subprocess.run(
        [str(py), "-m", "pip", "install", "-q", "--upgrade", "pip"],
        check=True,
        cwd=PKG,
    )
    subprocess.run(
        [str(py), "-m", "pip", "install", "-q", "-r", str(REQUIREMENTS)],
        check=True,
        cwd=PKG,
    )
    return py


def load_vault_into_env() -> bool:
    from vault import dsn_from_database_section, decrypt_payload, default_vault_path, read_vault

    vault_path = default_vault_path()
    if not vault_path.is_file():
        return False

    password = (os.environ.get("INFORME_VAULT_PASSWORD") or "").strip()
    if not password:
        password = _prompt_password(f"Contraseña del vault ({vault_path.name}): ")
    if not password:
        print("Se requiere contraseña para descifrar credenciales.", file=sys.stderr)
        raise SystemExit(1)

    _say("Verificando contraseña (espere unos segundos)…")
    try:
        payload = decrypt_payload(read_vault(vault_path), password)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(1) from e

    db = payload.get("database")
    if not isinstance(db, dict):
        print('El vault no contiene sección "database".', file=sys.stderr)
        raise SystemExit(1)

    os.environ["DATABASE_URL"] = dsn_from_database_section(db)
    os.environ.setdefault("PATH_CONF", str(PKG))
    os.environ.setdefault("FILE_CONF", "")
    return True


def resolve_connection_hint() -> None:
    if os.environ.get("DATABASE_URL", "").strip():
        return
    from core import dsn_from_local_conf, resolve_local_conf_path

    conf = resolve_local_conf_path(None)
    dsn = dsn_from_local_conf(conf)
    if dsn:
        os.environ["DATABASE_URL"] = dsn
        return

    print(
        "No hay conexión a BD configurada.\n"
        "  • Operadores: coloca credentials.vault.enc aquí y ejecuta de nuevo.\n"
        "  • Desarrollo: local.conf o DATABASE_URL en .env",
        file=sys.stderr,
    )
    raise SystemExit(1)


def main() -> None:
    if str(CORE_DIR) not in sys.path:
        sys.path.insert(0, str(CORE_DIR))

    py = ensure_venv()
    if os.environ.get("INFORME_VENV_BOOTSTRAPPED") != "1":
        if Path(sys.executable).resolve() != py.resolve():
            _reexec_in_venv(py)

    used_vault = load_vault_into_env()
    if not used_vault:
        resolve_connection_hint()

    port = (os.environ.get("INFORME_PORT") or "8501").strip()
    url = f"http://localhost:{port}"
    env = os.environ.copy()
    cmd = [
        str(py),
        "-u",
        "-m",
        "streamlit",
        "run",
        str(APP),
        "--server.port",
        port,
        "--server.headless",
        "true",
        "--server.runOnSave",
        "false",
        "--browser.gatherUsageStats",
        "false",
    ]
    _say(f"Credenciales OK. Iniciando la aplicación en {url}")
    _say("(La primera vez puede tardar 30–60 s; no cierres esta ventana.)")

    def _open_browser() -> None:
        time.sleep(3)
        try:
            webbrowser.open(url)
        except OSError:
            _say(f"Abre manualmente el navegador en {url}")

    threading.Thread(target=_open_browser, daemon=True).start()
    raise SystemExit(subprocess.call(cmd, cwd=PKG, env=env))


if __name__ == "__main__":
    main()
