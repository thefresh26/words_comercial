"""
Consulta en la base de datos (usando el local.conf de este equipo) los 3
datos que faltan para completar las 12 Actas de Alcance de los Paquetes 1
al 13: Código ActiBid del inmueble, oferente ganador (nombre + cédula/NIT)
y el monto/puja ganadora.

Este script se corre DIRECTO en este computador (donde sí hay salida de
red hacia la base de datos), no en la nube. Genera un archivo
"datos_alcance_paquetes_1_13.json" en esta misma carpeta con los
resultados, que luego se le pasa de vuelta a Claude para terminar de
completar los 12 documentos .docx ya generados.

Uso:
    python consultar_datos_bd_alcance.py
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO / "core"))
sys.path.insert(0, str(PROYECTO))

from core import get_connection, load_dotenv_files, fetch_informe  # noqa: E402

# (etiqueta para el resumen, identificador a buscar: FMI/código)
CASOS = [
    ("Paquete 1 - Subasta 1.1 (Acta 015)", "50C-814138"),
    ("Paquete 1 - Subasta 1.2 (Acta 014)", "200-179598"),
    ("Paquete 1 - Subasta 1.3 (Acta 013)", "50C-1257163"),
    ("Paquete 3 (Acta 16 / referenciada No.17)", "060-279174"),
    ("Paquete 4 (Acta 018)", "001-313399"),
    ("Paquete 6 - Subasta 1.1 (Acta 19 / referenciada No.026)", "001-790674"),
    ("Paquete 6 - Subasta 1.2 (Acta 020)", "50C-332612"),
    ("Paquete 9 (Acta 023, 4 inmuebles)", "001-42028"),
    ("Paquete 10 (Acta 22, 3 grupos)", "200-88115"),
    ("Paquete 11 (Acta 024, 4 inmuebles)", "50N-1112089"),
    ("Paquete 12 (Acta Comité Técnico Sociedades, 4 inmuebles)", "01N-5168178"),
    ("Paquete 13 (Acta 025)", "001-911997"),
]


def _conectar():
    load_dotenv_files()
    local_conf_path = PROYECTO / "local.conf"
    if not local_conf_path.is_file():
        raise FileNotFoundError(f"No se encontró local.conf en {PROYECTO}")
    return get_connection(None, local_conf_path)


def _resolver_identificador(conn, identificador: str) -> str:
    identificador = identificador.strip()
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", identificador.lower()):
        return identificador
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1", (identificador,))
        row = cur.fetchone()
        if row:
            return str(row[0])

    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT id, grupo_id, codigo, numero_matricula, codigo_grupo, nombre_grupo
            FROM mst_inmuebles
            WHERE UPPER(numero_matricula) = UPPER(%s)
               OR UPPER(codigo) = UPPER(%s)
               OR UPPER(codigo_grupo) = UPPER(%s)
               OR UPPER(referencia) = UPPER(%s)
            """,
            (identificador, identificador, identificador, identificador),
        )
        rows = cur.fetchall()

    if not rows:
        raise ValueError(f"No se encontró ninguna subasta, código, FMI ni unidad con: {identificador}")

    inmueble_ids = sorted({r[0] for r in rows if r[0] is not None})
    grupo_ids = sorted({r[1] for r in rows if r[1] is not None})

    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT DISTINCT
                COALESCE(a.id, psv.auction_id)      AS auction_id,
                COALESCE(a.code, psv.auction_code)  AS code,
                COALESCE(a.status, psv.estado)      AS status,
                COALESCE(a.start_date, psv.fecha_inicio) AS start_date
            FROM polibid_subastas_v2 psv
            LEFT JOIN polybid.auctions a ON a.id = psv.auction_id
            WHERE psv.inmueble_id = ANY(%s)
               OR (psv.grupo_id IS NOT NULL AND psv.grupo_id = ANY(%s))
            ORDER BY start_date DESC NULLS LAST
            """,
            (inmueble_ids or [-1], grupo_ids or [-1]),
        )
        candidatos = cur.fetchall()

    if not candidatos:
        raise ValueError(f"Se encontró '{identificador}' pero no tiene ninguna subasta asociada.")

    # Toma la más reciente automáticamente (sin preguntar, esto corre sin consola interactiva del usuario)
    return str(candidatos[0][0])


def _fmt_numero(valor) -> str:
    if not valor:
        return "0"
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return str(valor)


def obtener_datos_bd(conn, auction_uuid: str) -> dict:
    data = fetch_informe(conn, auction_uuid)
    subasta = data.get("subasta", {})
    ganador = data.get("ganador") or {}

    codigo_inmueble = "—"
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT psv.inmueble_id, psv.grupo_id
            FROM polibid_subastas_v2 psv
            WHERE psv.auction_id = %s::uuid
            ORDER BY psv.id DESC LIMIT 1
            """,
            (auction_uuid,),
        )
        link = cur.fetchone()
        if link:
            inmueble_id, grupo_id = link
            if inmueble_id:
                cur.execute("SELECT codigo, codigo_grupo FROM mst_inmuebles WHERE id = %s", (inmueble_id,))
            elif grupo_id:
                cur.execute(
                    "SELECT codigo, codigo_grupo FROM mst_inmuebles WHERE grupo_id = %s "
                    "ORDER BY es_padre DESC, id LIMIT 1",
                    (grupo_id,),
                )
            else:
                cur.execute("SELECT NULL, NULL")
            row = cur.fetchone()
            if row:
                codigo_inmueble = row[1] or row[0] or "—"

    nombre_ganador = (ganador.get("nombre_principal") or ganador.get("display_name") or "—").upper()
    id_ganador = "—"
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT ct.identificacion_numero
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid AND p.client_id = %s
            LIMIT 1
            """,
            (auction_uuid, ganador.get("client_id")),
        )
        row = cur.fetchone()
        if row:
            id_ganador = str(row[0] or "—")

    return {
        "codigo_subasta":  subasta.get("code", "—"),
        "codigo_inmueble": codigo_inmueble,
        "nombre_ganador":  nombre_ganador,
        "id_ganador":      id_ganador,
        "monto_ganador":   _fmt_numero(ganador.get("amount", 0)),
    }


def main():
    print(f"\n{'='*60}")
    print("  CONSULTA DE DATOS BD PARA ACTAS DE ALCANCE (Paquetes 1-13)")
    print(f"{'='*60}\n")

    print("Conectando a la base de datos...")
    conn = _conectar()
    print("✓ Conectado\n")

    resultados = {}
    try:
        for etiqueta, identificador in CASOS:
            print(f"[{identificador}] {etiqueta} ...", end=" ", flush=True)
            try:
                auction_uuid = _resolver_identificador(conn, identificador)
                bd = obtener_datos_bd(conn, auction_uuid)
                resultados[identificador] = {"etiqueta": etiqueta, "ok": True, **bd}
                print(f"OK  (Código: {bd['codigo_inmueble']}, Ganador: {bd['nombre_ganador']})")
            except Exception as e:
                resultados[identificador] = {"etiqueta": etiqueta, "ok": False, "error": str(e)}
                print(f"ERROR: {e}")
    finally:
        conn.close()

    salida = Path(__file__).resolve().parent.parent / "datos_alcance_paquetes_1_13.json"
    salida.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n✓ Resultados guardados en: {salida}")


if __name__ == "__main__":
    main()
