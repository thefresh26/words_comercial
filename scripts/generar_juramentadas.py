"""
Genera una DECLARACIÓN JURAMENTADA (.docx) por cada participante.

Uso:
    python scripts/generar_juramentadas.py <auction_uuid>
    python scripts/generar_juramentadas.py ACTIBID-POLI-46-...
    python scripts/generar_juramentadas.py 41742149          (solo cédula)

Requisito: 003_FORMATO_DECLARACION_JURAMENTADA_FO_GP_008.docx en templates/.
Los archivos se guardan en output/juramentadas/
"""
from __future__ import annotations
import sys, zipfile
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
CORE_DIR = PROYECTO / "core"
PLANTILLA = PROYECTO / "templates" / "003_FORMATO_DECLARACION_JURAMENTADA_FO_GP_008.docx"
SALIDA    = PROYECTO / "output" / "juramentadas"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_fecha(fecha) -> str:
    if not fecha: return "—"
    return f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"


def _formatear_fecha(fecha) -> str:
    if not fecha: return "—"
    return f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"


def _conectar():
    """Abre conexión a la BD usando local.conf o vault."""
    sys.path.insert(0, str(CORE_DIR))
    from core import get_connection, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section

    load_dotenv_files()
    local_conf_path = PROYECTO / "local.conf"
    vault_path      = PROYECTO / "credentials.vault.enc"
    dsn = None
    if local_conf_path.is_file():
        dsn = None
    elif vault_path.is_file():
        password = input("Contraseña del vault: ")
        blob  = read_vault(vault_path)
        creds = decrypt_payload(blob, password)
        db    = creds.get("database") or creds
        dsn   = dsn_from_database_section(db)

    conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
    print("✓ Conectado a la base de datos")
    return conn


# ── Resolver entrada ──────────────────────────────────────────────────────────

def _resolver_uuid(conn, identificador: str) -> str:
    """Acepta UUID o código de subasta y devuelve el UUID."""
    import re
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
                identificador.lower()):
        return identificador
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1",
                    (identificador,))
        row = cur.fetchone()
        if row:
            return str(row[0])
    raise ValueError(f"No se encontró subasta: {identificador}")


# ── Modo cédula: sin subasta ──────────────────────────────────────────────────

def _desde_cedula(cedula: str) -> list[dict]:
    """Genera participante directamente desde cédula sin necesitar subasta."""
    conn = _conectar()
    with conn.cursor() as cur:
        cur.execute("""
            SELECT nombre_principal, identificacion_numero, identificacion_tipo,
                   lugar_expedicion_doc, ciudad, fecha_diligenciamiento
            FROM contact_terceros
            WHERE identificacion_numero = %s LIMIT 1
        """, (cedula,))
        row = cur.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No se encontró persona con cédula: {cedula}")

    nombre, cedula_num, tipo_id, lugar_exp, ciudad, fecha = row
    ciudad_cedula = ciudad if tipo_id == "NIT" else (lugar_exp or "—")

    print(f"✓ Encontrado: {nombre}")
    return [{
        "nombre":        (nombre or "—").upper(),
        "cedula":        str(cedula_num or "—"),
        "ciudad_cedula": ciudad_cedula or "—",
        "fecha":         _formatear_fecha(fecha),
        "id":            cedula,
    }]


# ── Modo subasta ──────────────────────────────────────────────────────────────

def _desde_subasta(identificador: str) -> tuple[dict, list[dict]]:
    """Obtiene datos desde una subasta (UUID o código)."""
    conn = _conectar()
    from core import fetch_informe

    auction_uuid = _resolver_uuid(conn, identificador)

    data = fetch_informe(conn, auction_uuid)
    print(f"✓ Subasta: {data['subasta'].get('code')} — {data['subasta'].get('status')}")
    print(f"  Participantes: {len(data['participantes'])}")

    participantes = _obtener_datos(conn, auction_uuid)
    conn.close()
    return data, participantes


def _obtener_datos(conn, auction_uuid: str) -> list[dict]:
    """Solo SELECT — no modifica nada."""
    with conn.cursor() as cur:
        cur.execute("""
            SELECT ct.nombre_principal, ct.identificacion_numero,
                   ct.identificacion_tipo, ct.lugar_expedicion_doc,
                   ct.ciudad, ct.fecha_diligenciamiento, p.id
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid
            ORDER BY p.created_at
        """, (auction_uuid,))
        resultado = []
        for row in cur.fetchall():
            nombre, cedula, tipo_id, lugar_exp, ciudad, fecha, pid = row
            ciudad_cedula = ciudad if tipo_id == "NIT" else (lugar_exp or "—")
            resultado.append({
                "nombre":        (nombre or "—").upper(),
                "cedula":        str(cedula or "—"),
                "ciudad_cedula": ciudad_cedula or "—",
                "fecha":         _formatear_fecha(fecha),
                "id":            str(pid),
            })
        return resultado


# ── Generar docx ──────────────────────────────────────────────────────────────

def generar_docx(participante: dict, ruta_salida: Path) -> None:
    with zipfile.ZipFile(PLANTILLA, "r") as zin:
        archivos = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = archivos["word/document.xml"].decode("utf-8")
    for ph, val in {
        "##nombre##":        participante["nombre"],
        "##cedula##":        participante["cedula"],
        "##ciudad_cedula##": participante["ciudad_cedula"],
        "##fecha##":         participante["fecha"],
    }.items():
        doc_xml = doc_xml.replace(ph, val)

    archivos["word/document.xml"] = doc_xml.encode("utf-8")
    with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Uso: python scripts/generar_juramentadas.py <uuid|codigo|cedula>")
        sys.exit(1)

    if not PLANTILLA.exists():
        print(f"✗ No se encuentra la plantilla: {PLANTILLA.name}")
        sys.exit(1)

    identificador = sys.argv[1].strip()

    print(f"\n{'='*55}")
    print(f"  GENERADOR DE DECLARACIONES JURAMENTADAS")
    print(f"{'='*55}")
    print(f"  Entrada: {identificador}\n")

    SALIDA.mkdir(exist_ok=True)

    # Modo cédula — sin subasta
    if identificador.isdigit():
        participantes = _desde_cedula(identificador)
        codigo = f"CC-{identificador}"
    else:
        data, participantes = _desde_subasta(identificador)
        codigo = data["subasta"].get("code", identificador[:8])

    if not participantes:
        print("⚠ No se encontraron participantes.")
        sys.exit(0)

    print(f"\nGenerando en: output/juramentadas/\n")
    for i, p in enumerate(participantes, 1):
        nombre_corto = p["nombre"].replace(" ", "_")[:30]
        nombre_archivo = f"JURA_{codigo}_{i:02d}_{nombre_corto}.docx"
        ruta = SALIDA / nombre_archivo
        try:
            generar_docx(p, ruta)
            print(f"  ✓ [{i}/{len(participantes)}] {nombre_archivo}")
            print(f"       Nombre:  {p['nombre']}")
            print(f"       Cédula:  {p['cedula']}")
            print(f"       Ciudad:  {p['ciudad_cedula']}")
            print(f"       Fecha:   {p['fecha']}")
        except Exception as e:
            print(f"  ✗ ERROR — {p['nombre']}: {e}")

    print(f"\n✓ Listo. {len(participantes)} declaración(es) en ./juramentadas/")


if __name__ == "__main__":
    main()
