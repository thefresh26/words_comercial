"""
GENERADOR MASIVO: corre acta, informe, certificados de DD y juramentadas
para una LISTA de FMIs/códigos/unidades, uno tras otro, sin tener que
escribir el comando manualmente para cada uno.

Uso:
    python generar_lote.py fmis.txt
    python generar_lote.py fmis.txt acta,informe
    python generar_lote.py fmis.txt acta,informe,dd,juramentada

Donde fmis.txt es un archivo de texto con UN identificador por línea
(FMI, código de inmueble, código de unidad UNI-XXXX-AAAA, código de
subasta o UUID). Líneas vacías o que empiecen con # se ignoran.

Ejemplo de fmis.txt:
    # comentario, se ignora
    001-949942
    50C-1591957
    UNI-0090-2025
    100-17480

Si no se indica la lista de documentos (segundo argumento), se generan
LOS 4 tipos para cada identificador: acta, informe, certificados de DD y
juramentadas.

Si un FMI/unidad tiene varias subastas asociadas, este script NO pregunta
nada por consola (no hay quién responda en un lote) — usa automáticamente
la subasta más reciente y lo deja anotado en el resumen final.

Cada documento se genera llamando al script correspondiente
(generar_acta.py, generar_informe.py, generar_certificados_dd.py,
generar_juramentadas.py) con el UUID de subasta ya resuelto, así que la
lógica de scraping/Clarity/plantillas de cada uno no se duplica aquí.

Al final se imprime un resumen con OK/ERROR por cada FMI y documento, y se
guarda ese mismo resumen en output/lote_resultado_<fecha>.txt.
"""
from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
SCRIPTS  = PROYECTO / "scripts"
SALIDA   = PROYECTO / "output"

DOCS_DISPONIBLES = {
    "acta":        "generar_acta.py",
    "informe":     "generar_informe.py",
    "dd":          "generar_certificados_dd.py",
    "juramentada": "generar_juramentadas.py",
}


def _conectar():
    sys.path.insert(0, str(PROYECTO / "core"))
    sys.path.insert(0, str(PROYECTO))
    from core import get_connection, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section

    load_dotenv_files()
    local_conf_path = PROYECTO / "local.conf"
    vault_path = PROYECTO / "credentials.vault.enc"
    dsn = None
    if local_conf_path.is_file():
        dsn = None
    elif vault_path.is_file():
        password = input("Contraseña del vault: ")
        blob = read_vault(vault_path)
        creds = decrypt_payload(blob, password)
        db = creds.get("database") or creds
        dsn = dsn_from_database_section(db)

    conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
    return conn


def _resolver_uuid(conn, identificador: str) -> str:
    import re
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
                identificador.lower()):
        return identificador
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1", (identificador,))
        row = cur.fetchone()
        if row:
            return str(row[0])
    raise ValueError(f"No se encontró subasta: {identificador}")


def _resolver_automatico(conn, identificador: str) -> tuple[str, str]:
    """Igual que el resolver de los otros scripts, pero SIN preguntar nada
    por consola: si hay varias subastas para el mismo FMI/unidad, toma la
    más reciente automáticamente (la consulta ya viene ordenada por
    start_date DESC). Devuelve (auction_uuid, nota) — 'nota' queda vacía si
    no hubo ambigüedad, o explica qué pasó si la hubo."""
    identificador = identificador.strip()
    try:
        return _resolver_uuid(conn, identificador), ""
    except ValueError:
        pass

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
        raise ValueError(
            f"No se encontró ninguna subasta, código, FMI ni unidad inmobiliaria "
            f"con el identificador: {identificador}"
        )

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
        raise ValueError(
            f"Se encontró el inmueble/unidad '{identificador}' pero no tiene "
            f"ninguna subasta asociada."
        )

    auction_id, code = candidatos[0][0], candidatos[0][1]
    if len(candidatos) == 1:
        return str(auction_id), ""

    nota = (f"{len(candidatos)} subastas encontradas para '{identificador}'; "
            f"se usó la más reciente: {code or auction_id}")
    return str(auction_id), nota


def _correr_script(nombre_script: str, auction_uuid: str) -> tuple[bool, str]:
    """Corre uno de los 4 scripts pasándole el UUID ya resuelto (así nunca
    dispara el modo interactivo de elegir subasta, aunque el FMI original
    tuviera varias). Devuelve (ok, resumen_corto)."""
    try:
        resultado = subprocess.run(
            [sys.executable, nombre_script, auction_uuid],
            cwd=str(SCRIPTS),
            capture_output=True,
            text=True,
            timeout=180,
        )
    except subprocess.TimeoutExpired:
        return False, "⏱ Se agotó el tiempo de espera (180s)"

    salida = (resultado.stdout or "") + (resultado.stderr or "")
    if resultado.returncode != 0:
        # Se muestra la última línea no vacía como resumen del error.
        lineas = [l for l in salida.splitlines() if l.strip()]
        ultima = lineas[-1] if lineas else f"código de salida {resultado.returncode}"
        return False, ultima
    if "✓" not in salida and "Listo" not in salida:
        return False, "Terminó sin errores pero no se vio confirmación de éxito"
    return True, "OK"


def main():
    if len(sys.argv) < 2:
        print(f"Uso: python generar_lote.py <archivo_fmis.txt> [{','.join(DOCS_DISPONIBLES)}]")
        sys.exit(1)

    ruta_lista = Path(sys.argv[1])
    if not ruta_lista.is_file():
        ruta_lista = PROYECTO / sys.argv[1]
    if not ruta_lista.is_file():
        print(f"✗ No se encontró el archivo: {sys.argv[1]}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        docs_pedidos = [d.strip().lower() for d in sys.argv[2].split(",") if d.strip()]
        invalidos = [d for d in docs_pedidos if d not in DOCS_DISPONIBLES]
        if invalidos:
            print(f"✗ Documento(s) no reconocido(s): {', '.join(invalidos)}")
            print(f"  Válidos: {', '.join(DOCS_DISPONIBLES)}")
            sys.exit(1)
    else:
        docs_pedidos = list(DOCS_DISPONIBLES)

    identificadores = [
        l.strip() for l in ruta_lista.read_text(encoding="utf-8").splitlines()
        if l.strip() and not l.strip().startswith("#")
    ]
    if not identificadores:
        print("✗ El archivo de lista está vacío.")
        sys.exit(1)

    print(f"\n{'='*60}")
    print("  GENERADOR MASIVO")
    print(f"{'='*60}")
    print(f"  Lista:      {ruta_lista}")
    print(f"  Elementos:  {len(identificadores)}")
    print(f"  Documentos: {', '.join(docs_pedidos)}\n")

    conn = _conectar()
    print("✓ Conectado a la base de datos\n")

    resumen: list[dict] = []

    for idx, identificador in enumerate(identificadores, start=1):
        print(f"[{idx}/{len(identificadores)}] {identificador}")
        fila = {"identificador": identificador, "nota": "", "resultados": {}}
        try:
            auction_uuid, nota = _resolver_automatico(conn, identificador)
            if nota:
                print(f"  ⚠ {nota}")
            fila["nota"] = nota
            fila["auction_uuid"] = auction_uuid
        except Exception as e:
            print(f"  ✗ No se pudo resolver: {e}")
            fila["nota"] = f"NO RESUELTO: {e}"
            for doc in docs_pedidos:
                fila["resultados"][doc] = (False, "no resuelto")
            resumen.append(fila)
            print()
            continue

        for doc in docs_pedidos:
            script_nombre = DOCS_DISPONIBLES[doc]
            print(f"  → {doc}...", end=" ", flush=True)
            ok, detalle = _correr_script(script_nombre, auction_uuid)
            fila["resultados"][doc] = (ok, detalle)
            print("✓" if ok else f"✗ ({detalle})")

        resumen.append(fila)
        print()

    conn.close()

    # ── Resumen final ────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("  RESUMEN")
    print(f"{'='*60}")
    lineas_resumen = []
    for fila in resumen:
        estado = " | ".join(
            f"{doc}: {'OK' if fila['resultados'].get(doc, (False,''))[0] else 'ERROR'}"
            for doc in docs_pedidos
        ) if fila["resultados"] else "NO RESUELTO"
        linea = f"{fila['identificador']:<20} {estado}"
        if fila["nota"]:
            linea += f"   ({fila['nota']})"
        print(linea)
        lineas_resumen.append(linea)

    total_ok = sum(
        1 for fila in resumen for doc in docs_pedidos
        if fila["resultados"].get(doc, (False, ""))[0]
    )
    total_intentos = len(resumen) * len(docs_pedidos)
    print(f"\n{total_ok}/{total_intentos} documentos generados sin error.")

    SALIDA.mkdir(parents=True, exist_ok=True)
    log_path = SALIDA / f"lote_resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    log_path.write_text(
        f"Lista: {ruta_lista}\nDocumentos: {', '.join(docs_pedidos)}\n\n"
        + "\n".join(lineas_resumen)
        + f"\n\n{total_ok}/{total_intentos} documentos generados sin error.\n",
        encoding="utf-8",
    )
    print(f"\nResumen guardado en: {log_path.relative_to(PROYECTO)}")


if __name__ == "__main__":
    main()
