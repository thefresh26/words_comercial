"""
Genera un CERTIFICADO DE DEBIDA DILIGENCIA (.docx) por cada participante.

Uso:
    python generar_certificados_dd.py <identificador>
    python generar_certificados_dd.py 41742149      (solo cédula)
    python generar_certificados_dd.py 901325988     (solo NIT, persona jurídica)

Donde <identificador> puede ser una cédula o NIT (solo dígitos), el UUID de
la subasta, su código (ACTIBID-...), un FMI/número de matrícula, el código
de un inmueble individual, o el código de una unidad inmobiliaria
(UNI-XXXX-AAAA). Si el FMI/unidad tiene varias subastas asociadas, se
listan y se pide elegir. Si se pasa una cédula o NIT, se genera el
certificado para esa persona/empresa directamente, sin necesidad de
subasta.

Estructura esperada:
    AUTOMATIZACION_WORDS/
        core/           core.py, vault.py
        scripts/        este archivo
        templates/      002_CERTIFICADO_RESULTADO_DD.docx
        output/         certificados_dd/
        local.conf
"""
from __future__ import annotations
import re, sys, zipfile
from pathlib import Path

PROYECTO  = Path(__file__).parent
RAIZ      = PROYECTO.parent
CORE      = RAIZ / "core"
PLANTILLA = RAIZ / "templates" / "002_CERTIFICADO_RESULTADO_DD.docx"
SALIDA    = RAIZ / "output" / "certificados_dd"
LOCAL_CONF = RAIZ / "local.conf"
VAULT     = RAIZ / "credentials.vault.enc"


def _conectar():
    sys.path.insert(0, str(CORE))
    sys.path.insert(0, str(RAIZ))
    from core import get_connection, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section
    load_dotenv_files()
    dsn = None
    if not LOCAL_CONF.is_file() and VAULT.is_file():
        password = input("Contraseña del vault: ")
        blob = read_vault(VAULT)
        creds = decrypt_payload(blob, password)
        db = creds.get("database") or creds
        dsn = dsn_from_database_section(db)
    conn = get_connection(dsn, LOCAL_CONF if LOCAL_CONF.is_file() else None)
    print("✓ Conectado a la base de datos")
    return conn


def _resolver_uuid(conn, identificador: str) -> str:
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
                identificador.lower()):
        return identificador
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1", (identificador,))
        row = cur.fetchone()
        if row: return str(row[0])
    raise ValueError(f"No se encontró subasta: {identificador}")


def _resolver_identificador(conn, identificador: str) -> str:
    """Acepta UUID, código de subasta, FMI/número de matrícula, código de
    inmueble o código de unidad inmobiliaria (UNI-XXXX-AAAA), y devuelve el
    UUID de la subasta a usar. Si hay varias subastas asociadas al mismo
    FMI/unidad, se le pide al usuario elegir por consola."""
    identificador = identificador.strip()
    try:
        return _resolver_uuid(conn, identificador)
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
                COALESCE(a.title, psv.titulo)       AS title,
                COALESCE(a.status, psv.estado)      AS status,
                COALESCE(a.start_date, psv.fecha_inicio) AS start_date,
                COALESCE(a.end_date, psv.fecha_fin)      AS end_date
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

    if len(candidatos) == 1:
        auction_id, code = candidatos[0][0], candidatos[0][1]
        print(f"✓ Subasta encontrada automáticamente para '{identificador}': {code or auction_id}")
        return str(auction_id)

    print(f"\n⚠ Se encontraron {len(candidatos)} subastas asociadas a '{identificador}':\n")
    for i, (aid, code, title, status, start, end) in enumerate(candidatos, start=1):
        print(f"  [{i}] {code or aid}  |  estado: {status or '—'}  |  {title or '—'}")
        print(f"       Inicio: {_fmt_fecha_hora(start)}   Fin: {_fmt_fecha_hora(end)}")
    while True:
        seleccion = input(f"\nElige el número de subasta a usar (1-{len(candidatos)}): ").strip()
        if seleccion.isdigit() and 1 <= int(seleccion) <= len(candidatos):
            return str(candidatos[int(seleccion) - 1][0])
        print("  ⚠ Opción inválida, intenta de nuevo.")


def _fmt_fecha_hora(dt):
    if not dt:
        return "—"
    try:
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(dt)


def _desde_cedula(identificacion: str) -> list[dict]:
    """Genera participante directamente desde cédula o NIT, sin necesitar
    subasta. Sirve tanto para personas naturales (cédula) como jurídicas
    (NIT) — compara solo los dígitos, e ignora un posible dígito de
    verificación al final del NIT guardado en la base de datos.
    El id_puja se toma de su participación más reciente (si tiene alguna)."""
    solo_digitos = re.sub(r"\D", "", identificacion)
    conn = _conectar()
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT id, nombre_principal, identificacion_numero, identificacion_tipo
            FROM contact_terceros
            WHERE regexp_replace(identificacion_numero, '\\D', '', 'g') = %s
               OR regexp_replace(identificacion_numero, '\\D', '', 'g') LIKE %s
            ORDER BY (identificacion_tipo = 'NIT') DESC
            LIMIT 1
            """,
            (solo_digitos, solo_digitos + "_"),
        )
        row = cur.fetchone()
        if not row:
            conn.close()
            raise ValueError(f"No se encontró persona ni empresa con cédula/NIT: {identificacion}")
        contact_tercero_id, nombre, id_num, tipo_id = row

        cur.execute(
            """
            SELECT ap.id
            FROM polibid_credentials pc
            JOIN polybid.auction_participants ap ON ap.client_id = pc.client_id
            LEFT JOIN polybid.auctions a ON a.id = ap.auction_id
            WHERE pc.contact_tercero_id = %s
            ORDER BY a.start_date DESC NULLS LAST
            LIMIT 1
            """,
            (contact_tercero_id,),
        )
        puja_row = cur.fetchone()
    conn.close()

    id_puja = str(puja_row[0]) if puja_row and puja_row[0] is not None else "—"
    etiqueta = "NIT" if tipo_id == "NIT" else "cédula"
    print(f"✓ Encontrado ({etiqueta}): {nombre}")
    if puja_row:
        print(f"  Número de puja (participación más reciente): {id_puja}")
    else:
        print("  ⚠ No se encontró ninguna participación en subastas; id_puja queda en —.")
    return [{
        "nombre":  (nombre or "—").upper(),
        "cedula":  str(id_num or "—"),
        "id_puja": id_puja,
    }]


def _obtener_datos(conn, auction_uuid: str):
    from core import fetch_informe
    data = fetch_informe(conn, auction_uuid)
    print(f"✓ Subasta: {data['subasta'].get('code')} — {data['subasta'].get('status')}")
    print(f"  Participantes: {len(data['participantes'])}")

    with conn.cursor() as cur:
        cur.execute("""
            SELECT p.client_id, ct.identificacion_numero
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid
        """, (auction_uuid,))
        cedulas = {row[0]: row[1] or "—" for row in cur.fetchall()}

    participantes = []
    for p in data.get("participantes", []):
        client_id = p.get("client_id")
        participantes.append({
            "nombre":  (p.get("nombre_principal") or p.get("display_name") or "—").upper(),
            "cedula":  str(cedulas.get(client_id, "—")),
            "id_puja": str(p.get("id") or "—"),
        })
    return data, participantes


def _colapsar(xml: str) -> str:
    xml = re.sub(r'<w:t>##</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>([^<#]+)</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>##</w:t></w:r>',
                 r'<w:t>##\1##</w:t></w:r>', xml, flags=re.DOTALL)
    xml = re.sub(r'<w:t>##([^<#]+)#</w:t></w:r><w:proofErr[^/]*/><w:r[^>]*><w:t[^>]*>#\s*</w:t></w:r>',
                 r'<w:t>##\1##</w:t></w:r>', xml, flags=re.DOTALL)
    xml = re.sub(r'<w:t>([^<]+)</w:r>', r'<w:t>\1</w:t></w:r>', xml)
    return xml


def generar_docx(participante: dict, ruta_salida: Path) -> None:
    with zipfile.ZipFile(PLANTILLA, "r") as zin:
        archivos = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = archivos["word/document.xml"].decode("utf-8")
    doc_xml = _colapsar(doc_xml)
    for ph, val in {
        "##nombre##":  participante["nombre"],
        "##cedula##":  participante["cedula"],
        "##id_puja##": participante["id_puja"],
    }.items():
        doc_xml = doc_xml.replace(ph, val)

    archivos["word/document.xml"] = doc_xml.encode("utf-8")
    with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)


def main():
    if len(sys.argv) < 2:
        print("Uso: python generar_certificados_dd.py <cedula_o_nit | auction_uuid | codigo_subasta | FMI | codigo_unidad>")
        sys.exit(1)
    if not PLANTILLA.exists():
        print(f"✗ No se encuentra: {PLANTILLA}")
        sys.exit(1)

    identificador = sys.argv[1].strip()
    print(f"\n{'='*55}\n  GENERADOR DE CERTIFICADOS DE DEBIDA DILIGENCIA\n{'='*55}\n  Identificador: {identificador}\n")

    if identificador.isdigit():
        participantes = _desde_cedula(identificador)
        codigo = f"CC-{identificador}"
    else:
        conn = _conectar()
        auction_uuid = _resolver_identificador(conn, identificador)
        data, participantes = _obtener_datos(conn, auction_uuid)
        conn.close()
        codigo = data["subasta"].get("code", auction_uuid[:8])

    if not participantes:
        print("⚠ No se encontraron participantes.")
        sys.exit(0)

    SALIDA.mkdir(parents=True, exist_ok=True)

    print(f"\nGenerando en: output/certificados_dd/\n")
    for i, p in enumerate(participantes, 1):
        nombre_corto = p["nombre"].replace(" ", "_")[:30]
        nombre_archivo = f"DD_{codigo}_{i:02d}_{nombre_corto}.docx"
        ruta = SALIDA / nombre_archivo
        try:
            generar_docx(p, ruta)
            print(f"  ✓ [{i}/{len(participantes)}] {nombre_archivo}")
        except Exception as e:
            print(f"  ✗ ERROR — {p['nombre']}: {e}")

    print(f"\n✓ Listo. {len(participantes)} certificado(s) en output/certificados_dd/")


if __name__ == "__main__":
    main()
