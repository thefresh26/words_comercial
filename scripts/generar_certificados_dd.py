"""
Genera un CERTIFICADO DE DEBIDA DILIGENCIA (.docx) por cada participante
de una subasta, usando el archivo original como plantilla y reemplazando
los marcadores ##nombre##, ##cedula## y ##id_puja##.

Uso:
    python scripts/generar_certificados_dd.py <auction_uuid>

Requisito: 002_CERTIFICADO_RESULTADO_DD.docx en templates/.
Los archivos se guardan en output/certificados_dd/
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
CORE_DIR = PROYECTO / "core"
PLANTILLA = PROYECTO / "templates" / "002_CERTIFICADO_RESULTADO_DD.docx"
SALIDA    = PROYECTO / "output" / "certificados_dd"


# ── Conexión y datos ──────────────────────────────────────────────────────────


def _resolver_uuid(conn, identificador: str) -> str:
    """Acepta UUID o código de subasta y devuelve el UUID."""
    import re
    # Si ya es UUID
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", identificador.lower()):
        return identificador
    # Buscar por codigo
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1", (identificador,))
        row = cur.fetchone()
        if row:
            return str(row[0])
    raise ValueError(f"No se encontró subasta con código: {identificador}")


def conectar_y_obtener_datos(auction_uuid: str):
    sys.path.insert(0, str(CORE_DIR))
    from core import get_connection, fetch_informe, resolve_local_conf_path, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section

    load_dotenv_files()
    local_conf = resolve_local_conf_path(None)

    dsn = None
    vault_path = PROYECTO / "credentials.vault.enc"
    if vault_path.is_file():
        password = input("Contraseña del vault: ")
        blob = read_vault(vault_path)
        creds = decrypt_payload(blob, password)
        db = creds.get("database") or creds
        dsn = dsn_from_database_section(db)

    local_conf_path = PROYECTO / "local.conf"
    conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
    print(f"✓ Conectado a la base de datos")

    data = fetch_informe(conn, auction_uuid)
    print(f"✓ Subasta: {data['subasta'].get('code')} — {data['subasta'].get('status')}")
    print(f"  Participantes: {len(data['participantes'])}")

    cedulas = _obtener_cedulas(conn, auction_uuid)
    conn.close()
    return data, cedulas


def _obtener_cedulas(conn, auction_uuid: str) -> dict:
    """Solo SELECT — no modifica nada en la base de datos."""
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT p.client_id, ct.identificacion_numero
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid
            """,
            (auction_uuid,),
        )
        return {row[0]: row[1] or "—" for row in cur.fetchall()}


def preparar_participantes(data: dict, cedulas: dict) -> list[dict]:
    resultado = []
    for p in data.get("participantes", []):
        client_id = p.get("client_id")
        resultado.append({
            "nombre":  (p.get("nombre_principal") or p.get("display_name") or "—").upper(),
            "cedula":  str(cedulas.get(client_id, "—")),
            "id_puja": str(p.get("id") or "—"),
        })
    return resultado


# ── Reemplazo robusto en XML ──────────────────────────────────────────────────

def _colapsar_placeholder(xml: str, placeholder: str) -> str:
    """
    Word a veces parte ##texto## en múltiples runs XML.
    Esta función reconstruye el placeholder en un solo run antes de reemplazar.
    Ej: '##cedula#</w:t>...</w:t>#' → '##cedula##'
    """
    # Buscar el patrón partido: ##palabra# ... # (separado por tags XML)
    # Estrategia: colapsar todos los caracteres # que estén fragmentados
    # entre runs consecutivos dentro del mismo párrafo
    tag = placeholder.replace("##", "")  # ej: "cedula"

    # Patrón: ##tag# seguido de tags XML y luego #
    patron = rf'(##\s*{re.escape(tag)}\s*#)(</w:t>.*?<w:t[^>]*>)(#)'
    xml = re.sub(patron, rf'##{tag}##\2', xml, flags=re.DOTALL)
    return xml


def reemplazar_en_xml(xml: str, reemplazos: dict) -> str:
    """Reemplaza ##placeholder## en el XML manejando fragmentación de Word."""
    for placeholder, valor in reemplazos.items():
        xml = _colapsar_placeholder(xml, placeholder.strip("#"))
        xml = xml.replace(placeholder, valor)
    return xml


def generar_docx(participante: dict, ruta_salida: Path) -> None:
    with zipfile.ZipFile(PLANTILLA, "r") as zin:
        archivos = {name: zin.read(name) for name in zin.namelist()}

    reemplazos = {
        "##nombre##":  participante["nombre"],
        "##cedula##":  participante["cedula"],
        "##id_puja##": participante["id_puja"],
    }

    doc_xml = archivos["word/document.xml"].decode("utf-8")
    doc_xml = reemplazar_en_xml(doc_xml, reemplazos)
    archivos["word/document.xml"] = doc_xml.encode("utf-8")

    with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Uso: python generar_certificados_dd.py <auction_uuid>")
        sys.exit(1)

    if not PLANTILLA.exists():
        print(f"✗ No se encuentra la plantilla: {PLANTILLA.name}")
        sys.exit(1)

    auction_uuid = sys.argv[1].strip()

    print(f"\n{'='*55}")
    print(f"  GENERADOR DE CERTIFICADOS DE DEBIDA DILIGENCIA")
    print(f"{'='*55}")
    print(f"  Subasta: {auction_uuid}\n")

    data, cedulas = conectar_y_obtener_datos(auction_uuid)
    participantes = preparar_participantes(data, cedulas)

    if not participantes:
        print("⚠ No se encontraron participantes.")
        sys.exit(0)

    SALIDA.mkdir(exist_ok=True)
    codigo = data["subasta"].get("code", auction_uuid[:8])

    print(f"\nGenerando en: ./certificados_dd/\n")
    for i, p in enumerate(participantes, 1):
        nombre_corto = p["nombre"].replace(" ", "_")[:30]
        nombre_archivo = f"DD_{codigo}_{i:02d}_{nombre_corto}.docx"
        ruta = SALIDA / nombre_archivo
        try:
            generar_docx(p, ruta)
            print(f"  ✓ [{i}/{len(participantes)}] {nombre_archivo}")
        except Exception as e:
            print(f"  ✗ [{i}/{len(participantes)}] ERROR — {p['nombre']}: {e}")

    print(f"\n✓ Listo. {len(participantes)} certificado(s) en ./certificados_dd/")


if __name__ == "__main__":
    main()
