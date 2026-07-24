#!/usr/bin/env python3
"""
CLI — Informe de subasta Polibid.

  python utils/informe_subasta_polybid.py <UUID> --json

Interfaz web (operadores, un solo comando):

  python run.py
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_core_dir = Path(__file__).resolve().parent.parent / "core"
if str(_core_dir) not in sys.path:
    sys.path.insert(0, str(_core_dir))

from core import (
    fetch_informe,
    format_informe_texto,
    get_connection,
    json_default,
    load_dotenv_files,
    resolve_local_conf_path,
)


def _print_connection_help(conf_path: Path) -> None:
    print(
        "\nCómo conectar:\n"
        f"  1) Credenciales en {conf_path} → sección \"database\".\n"
        "  2) .env en la raíz: DATABASE_URL o PG*.\n"
        "  3) UI: python run.py\n",
        file=sys.stderr,
    )


def print_human(data: dict) -> None:
    print(format_informe_texto(data))


def main() -> None:
    ap = argparse.ArgumentParser(description="Informe subasta Polibid por UUID")
    ap.add_argument("auction_id", help="UUID de polybid.auctions.id")
    ap.add_argument("--json", action="store_true", help="Salida JSON")
    ap.add_argument("--dsn", metavar="URL", help="Cadena de conexión")
    ap.add_argument(
        "--config",
        metavar="RUTA",
        help="Ruta a local.conf",
    )
    args = ap.parse_args()

    load_dotenv_files()
    local_conf = resolve_local_conf_path(args.config)

    try:
        conn = get_connection(args.dsn, local_conf)
    except ImportError as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(1) from e
    except Exception as e:
        print(f"Error de conexión: {e}", file=sys.stderr)
        _print_connection_help(local_conf)
        raise SystemExit(2)

    try:
        data = fetch_informe(conn, args.auction_id.strip())
    except ValueError as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(3)
    finally:
        conn.close()

    if args.json:
        print(json.dumps(data, default=json_default, ensure_ascii=False, indent=2))
    else:
        print_human(data)


if __name__ == "__main__":
    main()
