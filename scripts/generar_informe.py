"""
Genera el INFORME DE SUBASTA ELECTRÓNICA (.docx).
Uso: python generar_informe.py <identificador>

Donde <identificador> puede ser el UUID de la subasta, su código
(ACTIBID-...), un FMI/número de matrícula, el código de un inmueble
individual, o el código de una unidad inmobiliaria (UNI-XXXX-AAAA). Si el
FMI/unidad tiene varias subastas asociadas, se listan y se pide elegir.

Requisito: INFORME_SUBASTA.docx en la misma carpeta.
El archivo se guarda en ./informes/
"""
from __future__ import annotations
import re, sys, time, zipfile
from io import BytesIO
from pathlib import Path

PROYECTO = Path(__file__).parent
PLANTILLA = PROYECTO.parent / "templates" / "INFORME_SUBASTA.docx"
SALIDA    = PROYECTO.parent / "output" / "informes"
MESES_ABR = {"ene":"01","feb":"02","mar":"03","abr":"04","may":"05","jun":"06",
             "jul":"07","ago":"08","sep":"09","sept":"09","oct":"10","nov":"11","dic":"12"}

FILA_FMI  = '<w:tr w:rsidR="004335EC" w14:paraId="7346AD35" w14:textId="77777777" w:rsidTr="001D7A74"><w:tc><w:tcPr><w:tcW w:w="800" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="3405DB84" w14:textId="0D6A24F0" w:rsidR="004335EC" w:rsidRDefault="003F62F4" w:rsidP="001E7040"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##fmi##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="726" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="5FC3148B" w14:textId="0F16A98C" w:rsidR="00711FCF" w:rsidRDefault="003F62F4" w:rsidP="0042080F"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##direccion##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="951" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="59F7FE2C" w14:textId="7B033D7C" w:rsidR="004335EC" w:rsidRDefault="003F62F4"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##ciudad##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="795" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="19DECA7C" w14:textId="344F554D" w:rsidR="00401EB7" w:rsidRPr="004E5742" w:rsidRDefault="003F62F4" w:rsidP="00401EB7"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##departamento##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="889" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="130AAB20" w14:textId="33B82877" w:rsidR="004335EC" w:rsidRDefault="003F62F4"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##tipo_inmueble##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="839" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="1B53A752" w14:textId="01C72209" w:rsidR="004335EC" w:rsidRDefault="003F62F4" w:rsidP="00672C46"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>##area##</w:t></w:r></w:p></w:tc></w:tr>'
FILA_PART = '<w:tr w:rsidR="004335EC" w14:paraId="449D52F6" w14:textId="77777777" w:rsidTr="00625CB7"><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="1B74246B" w14:textId="16E44D49" w:rsidR="004335EC" w:rsidRPr="00625CB7" w:rsidRDefault="003F62F4" w:rsidP="00625CB7"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>##id_partipante##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2420" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="15B3D4D3" w14:textId="229CDDEA" w:rsidR="00287CC0" w:rsidRDefault="003F62F4" w:rsidP="00F357C8"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>##nombre##</w:t></w:r></w:p><w:p w14:paraId="51A5E134" w14:textId="56E2256C" w:rsidR="004335EC" w:rsidRPr="00287CC0" w:rsidRDefault="000E3453" w:rsidP="00F357C8"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>##correo##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="875" w:type="pct"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="63BD7AE7" w14:textId="69A24D8A" w:rsidR="004335EC" w:rsidRPr="00625CB7" w:rsidRDefault="000E3453" w:rsidP="00625CB7"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>##estado_registro##</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:left w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/><w:right w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders><w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w14:paraId="7720E926" w14:textId="1912E6E2" w:rsidR="004335EC" w:rsidRPr="00625CB7" w:rsidRDefault="000E3453" w:rsidP="00625CB7"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>##fecha_registro##</w:t></w:r></w:p></w:tc></w:tr>'


def _resolver_uuid(conn, identificador):
    if re.match(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", identificador.lower()):
        return identificador
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM polybid.auctions WHERE code = %s LIMIT 1", (identificador,))
        row = cur.fetchone()
        if row: return str(row[0])
    raise ValueError(f"No se encontro subasta: {identificador}")


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


def _fmt_fecha(f):
    return f"{f.day:02d}/{f.month:02d}/{f.year}" if f else "—"

def _fmt_numero(v):
    try: return f"{int(v):,}".replace(",",".")
    except: return str(v) if v else "0"

def _slugify(t):
    if not t: return ""
    t = t.lower()
    for k,v in {"a\u0301":"a","e\u0301":"e","i\u0301":"i","o\u0301":"o","u\u0301":"u",
                "\u00e1":"a","\u00e9":"e","\u00ed":"i","\u00f3":"o","\u00fa":"u","\u00f1":"n"}.items():
        t = t.replace(k,v)
    t = re.sub(r"[\s\-]+","-",t); t = re.sub(r"[^a-z0-9\-]","",t)
    return t.strip("-")

def _limpiar(v):
    v = str(v) if v else ""
    v = "".join(c for c in v if ord(c) >= 32)
    return v.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

_RUN_RE = re.compile(
    r'<w:r\b[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>([^<]*)</w:t></w:r>',
    re.DOTALL,
)
_RUN_OPEN_RE = re.compile(r'^(<w:r\b[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>)', re.DOTALL)


def _colapsar_marcadores(xml):
    """Une en un solo <w:t> los marcadores '##nombre##' que Word partió en
    varios <w:r> seguidos (pasa muy seguido al escribir el marcador a mano en
    Word: cada tramo de texto que se escribe/pega puede quedar en su propio
    run). Solo fusiona runs mientras estén EXACTAMENTE pegados (sin ningún
    otro tag ni texto entre ellos) y mientras sigamos "dentro" de un '#' sin
    cerrar — así nunca se fusiona texto normal que no tiene nada que ver."""
    piezas = []
    pos = 0
    fin_anterior = None
    buffer = []        # runs (texto completo del match) acumulados sin cerrar
    buffer_textos = [] # sus contenidos de texto
    dobles_vistos = 0  # cantidad de "##" (el token completo, no el símbolo suelto) acumulados

    def volcar():
        if not buffer:
            return
        if len(buffer) == 1:
            piezas.append(buffer[0])
        else:
            m = _RUN_OPEN_RE.match(buffer[0])
            apertura = m.group(1) if m else "<w:r><w:t>"
            piezas.append(apertura + "".join(buffer_textos) + "</w:t></w:r>")
        buffer.clear()
        buffer_textos.clear()

    for m in _RUN_RE.finditer(xml):
        if fin_anterior is None or m.start() != fin_anterior:
            # hay algo (otro tag, salto de párrafo, etc.) entre este run y el
            # anterior: se cierra lo que hubiera quedado pendiente.
            volcar()
            dobles_vistos = 0
            piezas.append(xml[pos:m.start()])
        texto = m.group(1)
        buffer.append(m.group(0))
        buffer_textos.append(texto)
        # Se cuenta el token "##" completo (no el símbolo "#" suelto): un
        # número IMPAR de "##" acumulados significa que hay un marcador
        # abierto esperando su cierre (##nombre_partido_en_runs##).
        dobles_vistos += texto.count("##")
        if dobles_vistos % 2 == 0:
            volcar()
            dobles_vistos = 0
        pos = fin_anterior = m.end()

    volcar()
    piezas.append(xml[pos:])
    return "".join(piezas)


def _colapsar_xml(xml):
    # OJO: los reemplazos anteriores (basados en regex con ".*?" + DOTALL
    # buscando dos runs sueltos que digan exactamente "##") se quitaron
    # porque, si en el documento hay más de un marcador partido, ese patrón
    # puede "saltar" desde un "##" suelto hasta OTRO "##" suelto muy lejano,
    # tragándose (y corrompiendo) todo el contenido real que había en medio.
    # _colapsar_marcadores hace lo mismo pero solo une runs que están
    # exactamente pegados uno al otro, así que nunca puede saltar por encima
    # de contenido real.
    xml = re.sub(r'<w:t>([^<]+)</w:r>', r'<w:t>\1</w:t></w:r>', xml)
    xml = _colapsar_marcadores(xml)
    return xml


def _scrape(grupo_id, nombre_grupo, inm_id=None):
    res = {f"fecha_cronograma{i}":"—" for i in range(1,17)}
    res.update({"fecha_publicacion":"—","fecha_inicio":"—","fecha_fin":"—"})
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from webdriver_manager.chrome import ChromeDriverManager

        slug = _slugify(nombre_grupo)
        if grupo_id:
            url = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/{slug}"
        else:
            url = f"https://activosporcolombia.com/es/inmueble/{inm_id}/{slug}"

        opts = Options()
        for a in ["--headless","--no-sandbox","--disable-dev-shm-usage","--disable-gpu",
                  "--disable-extensions","--disable-images","--blink-settings=imagesEnabled=false",
                  "--disable-javascript-jit-optimization","--window-size=1600,1000"]:
            opts.add_argument(a)
        opts.page_load_strategy = "eager"  # no espera a que carguen imágenes/recursos secundarios
        opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

        print(f"  → {url}")
        # cache_valid_range evita que webdriver_manager revise por internet si
        # hay una versión nueva del driver en cada corrida — usa la que ya
        # tiene descargada mientras no pasen 30 días, ahorrando varios segundos.
        try:
            driver_path = ChromeDriverManager(cache_valid_range=30).install()
        except TypeError:
            driver_path = ChromeDriverManager().install()
        driver = webdriver.Chrome(service=Service(driver_path), options=opts)
        try:
            driver.set_page_load_timeout(15)
            driver.get(url)

            # El bloque de "Cronograma del proceso" carga sus datos por JS
            # DESPUÉS del render inicial de la página (a veces tarda más que
            # el resto). Un time.sleep() fijo muy corto puede capturar el
            # texto ANTES de que esos datos lleguen, y entonces el cronograma
            # sale vacío aunque sí exista en el sitio. En vez de un sleep
            # fijo, se reintenta leyendo el texto de la página hasta ver
            # "COMPLETADO"/fechas reales en la sección del cronograma, o
            # hasta agotar un máximo de espera.
            espera_max = 6.0
            paso = 0.3
            transcurrido = 0.0
            lineas = []
            while transcurrido < espera_max:
                time.sleep(paso)
                transcurrido += paso
                lineas = [l.strip() for l in driver.find_element(By.TAG_NAME,"body").text.split("\n")]
                texto_actual = "\n".join(lineas)
                if "Cronograma del proceso" in texto_actual and (
                    "COMPLETADO" in texto_actual.upper() or "→" in texto_actual
                ):
                    break
            print(f"  [debug] Esperó {transcurrido:.1f}s a que cargara el cronograma "
                  f"({len(lineas)} líneas capturadas)")

            inicio_cron = next((i for i,l in enumerate(lineas) if "Cronograma del proceso" in l), 0)
            print(f"  [debug] 'Cronograma del proceso' encontrado en la línea {inicio_cron}")
            print("  [debug] Líneas desde el cronograma (índice: texto):")
            for j in range(inicio_cron, min(inicio_cron + 60, len(lineas))):
                print(f"    [{j}] {lineas[j]!r}")

            # El año se busca SOLO dentro de la sección del cronograma (a partir
            # de "Cronograma del proceso"), no en toda la página: si se busca en
            # toda la página se puede agarrar por error un año que no tiene nada
            # que ver (p. ej. un "año de construcción" u otro dato numérico del
            # inmueble que aparezca antes en el texto), dañando TODAS las fechas.
            anio = "2026"
            for l in lineas[inicio_cron:]:
                m = re.search(r"20\d{2}", l)
                if m: anio = m.group(); break

            # El sitio tiene DOS variantes de cronograma según el tipo de página:
            #   - Unidad inmobiliaria (/unidad-inmobiliaria/...): cronograma viejo,
            #     con "Publicación próxima en subasta" como fase 1.
            #   - Inmueble individual (/inmueble/...): el sitio le quitó la fase
            #     de "Publicación"; ahora empieza directo en "Registro y cargue
            #     de documentos".
            # En vez de depender del nombre exacto de la fase 1 (que puede
            # cambiar), ##fecha_publicacion## siempre toma la fecha de inicio
            # de la fase que abre el cronograma, sea cual sea su nombre.
            texto_crono = "\n".join(lineas[inicio_cron:inicio_cron + 60]).lower()
            print(f"  [debug] ¿Detectó 'publicación próxima en subasta'?: "
                  f"{'publicación próxima en subasta' in texto_crono}")
            if "publicación próxima en subasta" in texto_crono:
                fases = [
                    ("Publicación próxima en subasta", "fecha_cronograma1",  "fecha_cronograma2"),
                    ("Registro",                       "fecha_cronograma3",  "fecha_cronograma4"),
                    ("Análisis debida diligencia",     "fecha_cronograma5",  "fecha_cronograma6"),
                    ("Análisis financiero",            "fecha_cronograma7",  "fecha_cronograma8"),
                    ("Expedición y envío de cupones",  "fecha_cronograma9",  "fecha_cronograma10"),
                    ("seriedad",                       "fecha_cronograma11", "fecha_cronograma12"),
                    ("Validación y confirmación",      "fecha_cronograma13", "fecha_cronograma14"),
                    ("Subasta",                        "fecha_cronograma15", "fecha_cronograma16"),
                ]
            else:
                fases = [
                    ("Registro y cargue de documentos",  "fecha_cronograma1",  "fecha_cronograma2"),
                    ("Análisis debida diligencia",       "fecha_cronograma3",  "fecha_cronograma4"),
                    ("Cargue de documentos financieros", "fecha_cronograma5",  "fecha_cronograma6"),
                    ("Análisis financiero",              "fecha_cronograma7",  "fecha_cronograma8"),
                    ("Expedición y envío de cupones",    "fecha_cronograma9",  "fecha_cronograma10"),
                    ("Pago seriedad de la oferta",       "fecha_cronograma11", "fecha_cronograma12"),
                    ("Validación y confirmación",        "fecha_cronograma13", "fecha_cronograma14"),
                    ("Subasta",                          "fecha_cronograma15", "fecha_cronograma16"),
                ]

            def _siguiente_no_vacia(desde):
                """Índice de la siguiente línea no vacía a partir de 'desde'
                (el sitio deja líneas en blanco entre el nombre de la fase,
                el rango de fechas y el mes)."""
                j = desde
                while j < len(lineas) and not lineas[j].strip():
                    j += 1
                return j if j < len(lineas) else None

            # La frase introductoria ("Sigue paso a paso el avance de tu
            # subasta") contiene la palabra "subasta" y aparece ANTES de las
            # tarjetas de fase — si se busca la fase "Subasta" desde
            # inicio_cron, esa frase se encuentra primero y se pisa la
            # fecha real de la última fase. Por eso la búsqueda de fases
            # arranca justo después de la línea "Del X al Y de <año>" (el
            # encabezado con el rango total), que siempre viene después de
            # esa frase y justo antes de la primera tarjeta de fase.
            inicio_fases = inicio_cron
            for i in range(inicio_cron, min(inicio_cron + 15, len(lineas))):
                if lineas[i].startswith("Del ") and "de 20" in lineas[i]:
                    inicio_fases = i + 1
                    break
            print(f"  [debug] Búsqueda de fases arranca en la línea {inicio_fases}")

            print(f"  [debug] Año detectado para el cronograma: {anio}")
            print(f"  [debug] Buscando {len(fases)} fases...")
            # IMPORTANTE: cada fase se busca a partir de donde terminó la
            # ANTERIOR ('cursor'), nunca desde el principio del cronograma.
            # Sin esto, nombres de fase que son substring de otro texto ya
            # visto (p. ej. "Subasta" está contenido dentro de "Publicación
            # próxima en subasta", la fase 1) hacían que una fase posterior
            # "rebotara" hacia atrás y repitiera la fecha de una fase ya
            # leída, en vez de encontrar su propia tarjeta más adelante.
            cursor = inicio_fases
            for fase, ph_ini, ph_fin in fases:
                encontrada = False
                for i in range(cursor, len(lineas)):
                    l = lineas[i]
                    if fase.lower() in l.lower():
                        encontrada = True
                        idx_cand = _siguiente_no_vacia(i+1)
                        cand = lineas[idx_cand] if idx_cand is not None else ""
                        print(f"  [debug]   '{fase}' → encontrada en línea {i} ({l!r}); "
                              f"siguiente línea no vacía [{idx_cand}] = {cand!r}")
                        if cand and cand[0].isdigit():
                            idx_mes = _siguiente_no_vacia(idx_cand+1)
                            mes_l = lineas[idx_mes] if idx_mes is not None else ""
                            mes = MESES_ABR.get(mes_l.split(".")[0].strip(),"00")
                            print(f"  [debug]     línea de mes [{idx_mes}] = {mes_l!r} → mes={mes}")
                            if "→" in cand:
                                partes = cand.split()
                                dia_i, dia_f = partes[0].zfill(2), partes[2].zfill(2)
                            else:
                                dia_i = dia_f = cand.split()[0].zfill(2)
                            res[ph_ini] = f"{dia_i}/{mes}/{anio}"
                            res[ph_fin]  = f"{dia_f}/{mes}/{anio}"
                            print(f"  [debug]     → {ph_ini}={res[ph_ini]}  {ph_fin}={res[ph_fin]}")
                            # La fase que abre el cronograma (sea cual sea su
                            # nombre) define ##fecha_publicacion##.
                            if ph_ini == "fecha_cronograma1":
                                res["fecha_publicacion"] = f"{dia_i}/{mes}/{anio} 10:00 am"
                            cursor = idx_mes + 1 if idx_mes is not None else i + 1
                        else:
                            print(f"  [debug]     ⚠ la línea siguiente no empieza con un dígito, "
                                  f"no se pudo leer la fecha de '{fase}'")
                            cursor = i + 1
                        break
                if not encontrada:
                    print(f"  [debug]   '{fase}' → NO encontrada en el texto del cronograma")

            meses_l = {"enero":"01","febrero":"02","marzo":"03","abril":"04","mayo":"05","junio":"06",
                       "julio":"07","agosto":"08","septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"}
            encontre = False; fechas = []
            for l in lineas:
                if "Estado de la Subasta" in l: encontre = True; continue
                if encontre and "de 20" in l and "a las" in l:
                    try:
                        p = l.lower().split()
                        dia=p[0].zfill(2); mes=meses_l.get(p[2],"00"); yr=p[4]; hr=p[7]
                        ap="".join(p[8:]).replace(".","")
                        fechas.append(f"{dia}/{mes}/{yr} {hr} {ap}")
                    except: pass
                    if len(fechas)==2: break
            if fechas: res["fecha_inicio"] = fechas[0]
            if len(fechas)>1: res["fecha_fin"] = fechas[1]

            # Respaldo: si la sección "Estado de la Subasta" no trajo nada
            # (pasa en subastas ya finalizadas, donde el sitio deja de
            # mostrar esa sección), se usa la fase 8 del cronograma
            # ("Subasta (apertura y cierre)"), que sí quedó bien detectada
            # arriba y trae las mismas fechas reales de apertura/cierre.
            if res["fecha_inicio"] == "—" and res.get("fecha_cronograma15","—") != "—":
                res["fecha_inicio"] = res["fecha_cronograma15"]
                print(f"  [debug] Apertura tomada de fecha_cronograma15 (fase 'Subasta'): {res['fecha_inicio']}")
            if res["fecha_fin"] == "—" and res.get("fecha_cronograma16","—") != "—":
                res["fecha_fin"] = res["fecha_cronograma16"]
                print(f"  [debug] Cierre tomado de fecha_cronograma16 (fase 'Subasta'): {res['fecha_fin']}")

            print(f"  → Publicación: {res['fecha_publicacion']}")
            print(f"  → Apertura:    {res['fecha_inicio']}")
            print(f"  → Cierre:      {res['fecha_fin']}")
        finally:
            driver.quit()
    except Exception as e:
        print(f"  ⚠ Error scraping: {e}")
    return res


def _buscar_inmueble_por_grupo(conn, grupo_id_val):
    """Devuelve (inmueble dict, fmis_lista, nombre_grupo) para un grupo_id."""
    with conn.cursor() as cur:
        cur.execute("""
            SELECT id, codigo, numero_matricula, grupo_id, nombre_grupo,
                   area_lote, area_construida, referencia, codigo_grupo
            FROM mst_inmuebles WHERE grupo_id = %s ORDER BY es_padre DESC, id
        """, (grupo_id_val,))
        rows = cur.fetchall()
    if not rows: return {}, [], ""
    r = rows[0]
    nombre_grupo = r[4] or r[7] or ""
    inmueble = {
        "codigo": r[8] or r[1] or "—",
        "fmi":    ", ".join(x[2] for x in rows if x[2]),
        "area":   str(r[5] or r[6] or 0),
    }
    fmis_lista = [{"fmi": x[2] or "—", "area": str(x[5] or x[6] or 0)} for x in rows]
    return inmueble, fmis_lista, nombre_grupo


def conectar(auction_uuid_input):
    sys.path.insert(0, str(PROYECTO.parent / "core"))
    sys.path.insert(0, str(PROYECTO.parent))
    from core import get_connection, fetch_informe, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section

    load_dotenv_files()
    local_conf_path = PROYECTO.parent / "local.conf"
    vault_path = PROYECTO.parent / "credentials.vault.enc"
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
    print("✓ Conectado a la base de datos")

    auction_uuid = _resolver_identificador(conn, auction_uuid_input)
    data = fetch_informe(conn, auction_uuid)
    print(f"✓ Subasta: {data['subasta'].get('code')} — {data['subasta'].get('status')}")
    datos = _obtener_datos(conn, auction_uuid, data)
    conn.close()
    return data, datos


def _obtener_datos(conn, auction_uuid, data):
    subasta = data["subasta"]

    # ── 1. Buscar inmueble ────────────────────────────────────────────────────
    inmueble = {}; grupo_id = None; nombre_grupo = ""; inm_id = None; fmis_lista = []

    # Primero: polibid_subastas_v2
    with conn.cursor() as cur:
        cur.execute("SELECT inmueble_id, grupo_id FROM polibid_subastas_v2 WHERE auction_id=%s::uuid ORDER BY id DESC LIMIT 1",(auction_uuid,))
        link = cur.fetchone()
        if link and link[0]:
            inm_id = link[0]
            cur.execute("""SELECT codigo, numero_matricula, grupo_id, nombre_grupo,
                                  area_lote, area_construida, referencia, codigo_grupo
                           FROM mst_inmuebles WHERE id=%s""",(inm_id,))
            row = cur.fetchone()
            if row:
                grupo_id     = row[2]
                nombre_grupo = row[3] or row[6] or ""
                if grupo_id:
                    inmueble, fmis_lista, nombre_grupo = _buscar_inmueble_por_grupo(conn, grupo_id)
                else:
                    inmueble = {"codigo": row[7] or row[0] or "—",
                                "fmi":   row[1] or "—",
                                "area":  str(row[4] or row[5] or 0)}
                    fmis_lista = [{"fmi": row[1] or "—", "area": str(row[4] or row[5] or 0)}]
        elif link and link[1]:
            # Subasta de una unidad inmobiliaria agrupada: psv.inmueble_id es
            # NULL en estos casos (es normal, no un error) y el vínculo
            # correcto es psv.grupo_id. Antes de este fix, al no revisar este
            # caso, el código caía al fallback de manifestacion_interes (otra
            # tabla, sin relación garantizada con lo realmente subastado) y
            # podía traer los datos de OTRA propiedad.
            grupo_id = link[1]
            inmueble, fmis_lista, nombre_grupo = _buscar_inmueble_por_grupo(conn, grupo_id)

    # Fallback: manifestacion_interes
    if not inmueble:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT mani.grupo_id, mani.inmueble_id
                FROM polybid.auction_participants p
                JOIN polibid_credentials pc ON pc.client_id = p.client_id
                JOIN manifestacion_interes mani ON mani.contact_tercero_id = pc.contact_tercero_id
                WHERE p.auction_id = %s::uuid
                AND (mani.grupo_id IS NOT NULL OR mani.inmueble_id IS NOT NULL)
                GROUP BY mani.grupo_id, mani.inmueble_id
                ORDER BY COUNT(*) DESC LIMIT 1
            """, (auction_uuid,))
            mani = cur.fetchone()
            if mani:
                if mani[0]:
                    grupo_id = mani[0]
                    inmueble, fmis_lista, nombre_grupo = _buscar_inmueble_por_grupo(conn, grupo_id)
                elif mani[1]:
                    inm_id = mani[1]
                    cur.execute("""SELECT codigo, numero_matricula, grupo_id, nombre_grupo,
                                          area_lote, area_construida, referencia, codigo_grupo
                                   FROM mst_inmuebles WHERE id=%s""",(inm_id,))
                    row = cur.fetchone()
                    if row:
                        grupo_id = row[2]
                        if grupo_id:
                            inmueble, fmis_lista, nombre_grupo = _buscar_inmueble_por_grupo(conn, grupo_id)
                        else:
                            nombre_grupo = row[3] or row[6] or ""
                            inmueble = {"codigo": row[7] or row[0] or "—",
                                        "fmi":   row[1] or "—",
                                        "area":  str(row[4] or row[5] or 0)}
                            fmis_lista = [{"fmi": row[1] or "—", "area": str(row[4] or row[5] or 0)}]

    # ── 2. Ciudad y departamento ──────────────────────────────────────────────
    ciudad = departamento = "—"
    if nombre_grupo and "," in nombre_grupo:
        parte = nombre_grupo.split("-")[-1].strip() if "-" in nombre_grupo else nombre_grupo
        m = re.search(r"([^,]+),\s*([^,]+)$", parte)
        if m: ciudad=m.group(1).strip(); departamento=m.group(2).strip()
    if ciudad == "—" and inm_id:
        with conn.cursor() as cur:
            try:
                cur.execute("""SELECT c.name, d.name FROM mst_inmuebles mi
                    JOIN cities c ON c.codigo=mi.city_id
                    JOIN departments d ON d.id_departamento=c.id_departamento
                    WHERE mi.id=%s""",(inm_id,))
                row = cur.fetchone()
                if row: ciudad=row[0] or "—"; departamento=row[1] or "—"
            except: pass

    # Enriquecer fmis_lista con ciudad, depto, tipo
    tipo_inmueble = "—"
    with conn.cursor() as cur:
        try:
            cur.execute("""SELECT ti.tipo_inmueble FROM polibid_subastas_v2 psv
                JOIN mst_inmuebles mi ON mi.id=psv.inmueble_id
                JOIN mst_tipos_inmueble ti ON ti.id_tipo_inmueble=mi.tipo_inmueble_id
                WHERE psv.auction_id=%s::uuid ORDER BY psv.id DESC LIMIT 1""",(auction_uuid,))
            row = cur.fetchone()
            if row: tipo_inmueble = row[0] or "—"
        except: pass

    for f in fmis_lista:
        f["direccion"]    = nombre_grupo or "—"
        f["ciudad"]       = ciudad
        f["departamento"] = departamento
        f["tipo_inmueble"]= tipo_inmueble

    # ── 3. Ganador ────────────────────────────────────────────────────────────
    ganador = data.get("ganador") or {}
    ganador_nombre = (ganador.get("nombre_principal") or ganador.get("display_name") or "—").upper()
    monto_ganador  = ganador.get("amount",0)
    precio_base    = subasta.get("initial_value",0)
    incremento = "—"
    if precio_base and monto_ganador:
        try: incremento = f"{((float(monto_ganador)-float(precio_base))/float(precio_base))*100:.2f}"
        except: pass

    # ── 4. Participantes ──────────────────────────────────────────────────────
    participantes = []; inscritos = con_ofertas = total_ofertas = 0
    with conn.cursor() as cur:
        cur.execute("""SELECT p.id, ct.nombre_principal, ct.email, p.status, p.created_at
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id=p.client_id
            LEFT JOIN contact_terceros ct ON ct.id=pc.contact_tercero_id
            WHERE p.auction_id=%s::uuid ORDER BY p.created_at""",(auction_uuid,))
        for row in cur.fetchall():
            participantes.append({"id":str(row[0]),"nombre":(row[1] or "—").upper(),
                "correo":(row[2] or "—").lower(),"estado_registro":row[3] or "—",
                "fecha_registro":_fmt_fecha(row[4])})
        inscritos = len(participantes)
        cur.execute("SELECT COUNT(DISTINCT client_id) FROM polybid.auction_bids WHERE auction_id=%s::uuid",(auction_uuid,))
        con_ofertas = cur.fetchone()[0] or 0
        cur.execute("SELECT COUNT(*) FROM polybid.auction_bids WHERE auction_id=%s::uuid",(auction_uuid,))
        total_ofertas = cur.fetchone()[0] or 0

    # ── 5. Pujas ──────────────────────────────────────────────────────────────
    pujas = []
    with conn.cursor() as cur:
        cur.execute("""SELECT b.id, ct.nombre_principal, ct.email, b.amount, b.created_at
            FROM polybid.auction_bids b
            LEFT JOIN polibid_credentials pc ON pc.client_id=b.client_id
            LEFT JOIN contact_terceros ct ON ct.id=pc.contact_tercero_id
            WHERE b.auction_id=%s::uuid ORDER BY b.created_at""",(auction_uuid,))
        for row in cur.fetchall():
            pujas.append({"id":str(row[0]),"nombre":(row[1] or "—").upper(),
                "email":(row[2] or "—").lower(),"monto":_fmt_numero(row[3]),
                "fecha":row[4].strftime("%d/%m/%Y %H:%M") if row[4] else "—"})

    # ── 6. URL pagina ─────────────────────────────────────────────────────────
    if grupo_id:
        url_pagina = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/{_slugify(nombre_grupo)}"
    elif inm_id:
        url_pagina = f"https://activosporcolombia.com/es/inmueble/{inm_id}/{_slugify(nombre_grupo)}"
    else:
        url_pagina = "—"

    # ── 7. Scraping ───────────────────────────────────────────────────────────
    print("\nObteniendo datos de la página web...")
    web = _scrape(grupo_id, nombre_grupo, inm_id)

    # ── 8. Tráfico web (Microsoft Clarity) ───────────────────────────────────
    print("\nObteniendo estadísticas de tráfico (Clarity)...")
    clarity = None
    try:
        sys.path.insert(0, str(PROYECTO.parent / "utils"))
        from clarity_capture import capturar_datos_clarity

        fecha_ini_clarity = web.get("fecha_cronograma1", "—")
        fecha_fin_clarity = web.get("fecha_cronograma16", "—")
        clarity = capturar_datos_clarity(fecha_ini_clarity, fecha_fin_clarity, url_pagina)
    except Exception as e:
        print(f"  ⚠ No se pudo obtener datos de Clarity: {e}")

    def _clarity_num(valor):
        return _fmt_numero(valor) if valor is not None else "—"

    def _clarity_num_o_cero(valor):
        # Para "sesiones de la URL": si no hubo sesiones (o no se pudo leer
        # el dato), se muestra 0 en vez de "—" — así lo pidió el usuario.
        return _fmt_numero(valor) if valor is not None else "0"

    # La fecha de publicación se toma PRIMERO del cronograma scrapeado de la
    # página (fecha real de inicio del proceso). Solo si la página no trae
    # ningún dato de cronograma (pasa con subastas ya finalizadas cuya URL
    # el sitio reutilizó para un ciclo de venta nuevo de la misma propiedad)
    # se usa como respaldo el created_at de la subasta en la base de datos
    # — que no es exactamente lo mismo, pero es mejor que dejarlo en blanco.
    fecha_publicacion = web.get("fecha_publicacion", "—")
    if fecha_publicacion == "—":
        with conn.cursor() as cur:
            cur.execute(
                "SELECT created_at FROM polybid.auctions WHERE id = %s::uuid",
                (auction_uuid,),
            )
            row = cur.fetchone()
            if row and row[0]:
                fecha_publicacion = f"{row[0].day:02d}/{row[0].month:02d}/{row[0].year} " \
                                     f"{row[0].strftime('%I:%M %p').lstrip('0').lower()}"

    # Igual que fecha_publicacion: "fecha_inicio"/"fecha_fin" ya vienen en
    # 'web' con el valor "—" puesto de entrada (ver _scrape), así que
    # web.get(clave, default) NUNCA usaba el default aunque el scraping
    # hubiera fallado — la clave siempre "existía" con valor "—". Se compara
    # explícitamente contra "—" para sí poder caer al valor de la subasta en
    # la base de datos cuando la página no trajo nada.
    fecha_inicio_web = web.get("fecha_inicio", "—")
    fecha_inicio = fecha_inicio_web if fecha_inicio_web != "—" else _fmt_fecha(subasta.get("start_date"))
    fecha_fin_web = web.get("fecha_fin", "—")
    fecha_fin = fecha_fin_web if fecha_fin_web != "—" else _fmt_fecha(subasta.get("end_date"))

    return {
        "codigo_subasta":    subasta.get("code","—"),
        "fecha_publicacion": fecha_publicacion,
        "fecha_inicio":      fecha_inicio,
        "fecha_fin":         fecha_fin,
        "fmi":               inmueble.get("fmi","—"),
        "direccion":         nombre_grupo or "—",
        "ciudad":            ciudad,
        "departamento":      departamento,
        "tipo_inmueble":     tipo_inmueble,
        "area":              inmueble.get("area","—"),
        "codigo_inmueble":   inmueble.get("codigo","—"),
        "precio_base":       _fmt_numero(precio_base),
        "nombre_ganador":    ganador_nombre,
        "oferta_gandora":    _fmt_numero(monto_ganador),
        "incremento":        incremento,
        "porcentaje_subasta": incremento,
        "incritos":          str(inscritos),
        "ofertas":           str(con_ofertas),
        "total_ofertas":     str(total_ofertas),
        "url_pagina":        url_pagina,
        "participantes":     participantes,
        "fmis_grupo":        fmis_lista,
        "pujas":             pujas,
        "clarity_sesiones_totales": _clarity_num(clarity.get("sesiones_totales")) if clarity else "—",
        "clarity_bots_excluidos":   _clarity_num(clarity.get("bots_excluidos")) if clarity else "—",
        "clarity_sesiones_url":     _clarity_num(clarity.get("sesiones_url")) if clarity else "—",
        # Alias con los nombres de marcador que quedaron en la plantilla:
        # ##cantidad## (sesiones totales), ##cantidad_sesiones## (bots
        # excluidos) y ##cantidad_url## (sesiones de la URL del inmueble,
        # 0 si no hubo ninguna).
        "cantidad":          _clarity_num(clarity.get("sesiones_totales")) if clarity else "0",
        "cantidad_sesiones": _clarity_num(clarity.get("bots_excluidos")) if clarity else "0",
        "cantidad_url":      _clarity_num_o_cero(clarity.get("sesiones_url")) if clarity else "0",
        "_clarity_imagenes": clarity,
        **{f"fecha_cronograma{i}": web.get(f"fecha_cronograma{i}","—") for i in range(1,17)},
    }


def _celda_puja(texto, cabecera=False, bg="FFFFFF", ancho="1080"):
    bold = "<w:b/>" if cabecera else ""
    color = '<w:color w:val="FFFFFF"/>' if cabecera else '<w:color w:val="000000"/>'
    fill = "1E3A5F" if cabecera else bg
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    borde = '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/></w:tcBorders>'
    rpr = f'<w:rPr>{bold}{color}<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
    ppr = f'<w:pPr><w:spacing w:before="60" w:after="60"/>{rpr}</w:pPr>'
    lineas = str(texto).split("\n")
    parrafos = "".join(f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{l}</w:t></w:r></w:p>' for l in lineas)
    return f'<w:tc><w:tcPr><w:tcW w:w="{ancho}" w:type="dxa"/>{borde}{shd}</w:tcPr>{parrafos}</w:tc>'

def _generar_tabla_pujas(pujas):
    if not pujas:
        return '<w:p><w:r><w:t>Sin pujas registradas.</w:t></w:r></w:p>'
    anchos = ["2800","4200","2000","2800"]
    cols   = ["ID","Usuario","Monto","Fecha"]
    cabecera = "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + "".join(
        _celda_puja(c, cabecera=True, ancho=a) for c,a in zip(cols,anchos)) + "</w:tr>"
    filas = ""
    for i,p in enumerate(pujas):
        bg = "F0F4F8" if i%2==0 else "FFFFFF"
        usuario = f"{p['nombre']}\n{p['email']}"
        vals = [p["id"], usuario, f"$ {p['monto']}", p["fecha"]]
        filas += "<w:tr>" + "".join(_celda_puja(v,bg=bg,ancho=a) for v,a in zip(vals,anchos)) + "</w:tr>"
    return (f'<w:tbl><w:tblPr><w:tblW w:w="11800" w:type="dxa"/><w:jc w:val="center"/>' +
            f'<w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'<w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'<w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/>' +
            f'</w:tblBorders></w:tblPr>{cabecera}{filas}</w:tbl><w:p/>')


# La plantilla INFORME_SUBASTA.docx ya NO trae las imágenes de Clarity
# embebidas: en su lugar tiene los marcadores de texto ##foto_clarity## y
# ##foto_clarity_url## (así se dejó a propósito, para que este script sea el
# que inserte las fotos reales). Por eso aquí no reutilizamos ninguna
# relación (rId) existente: cada vez que se genera un informe se agrega la
# imagen como una parte nueva del .docx (nuevo archivo en word/media/, nueva
# relación en word/_rels/document.xml.rels) y se reemplaza el texto del
# marcador por el XML del dibujo (<w:drawing>) correspondiente.
_CLARITY_ANCHO_EMU = {
    "foto_clarity": 5619750,       # resumen (Sesiones, Ideas, etc.)
    "foto_clarity_url": 3924754,   # tarjeta "Páginas principales"
}


def _siguiente_rid(rels_xml):
    """Calcula el próximo rId libre (rId<N+1>) mirando los que ya existen en
    word/_rels/document.xml.rels, para no chocar con ninguno."""
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def _agregar_relacion_imagen(archivos, nombre_media):
    """Registra un nuevo archivo de imagen como relación de tipo 'image' en
    word/_rels/document.xml.rels y devuelve el rId asignado."""
    rels_path = "word/_rels/document.xml.rels"
    rels_xml = archivos[rels_path].decode("utf-8")
    nuevo_rid = _siguiente_rid(rels_xml)
    nueva_relacion = (
        f'<Relationship Id="{nuevo_rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{nombre_media}"/>'
    )
    rels_xml = rels_xml.replace("</Relationships>", nueva_relacion + "</Relationships>")
    archivos[rels_path] = rels_xml.encode("utf-8")
    return nuevo_rid


def _drawing_xml(rid, cx, cy, doc_id, nombre):
    """XML mínimo de un dibujo (imagen en línea) para insertar dentro de un
    <w:r>, equivalente al que genera Word al pegar una imagen."""
    anchor = f"{doc_id:08X}"
    return (
        "<w:drawing>"
        f'<wp:inline distT="0" distB="0" distL="0" distR="0" '
        f'wp14:anchorId="{anchor}" wp14:editId="{anchor}">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_id}" name="{nombre}"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_id}" name="{nombre}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}" cstate="print"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
    )


def _insertar_imagen_en_placeholder(archivos, doc_xml, placeholder, pil_img, ancho_emu, doc_id, nombre_media):
    """Reemplaza el texto '##placeholder##' (dentro de su <w:t>) por una
    imagen real: agrega el archivo a word/media/, crea su relación y pone
    el XML del dibujo en el lugar exacto del marcador."""
    marca = f"##{placeholder}##"
    idx = doc_xml.find(marca)
    if idx < 0:
        return doc_xml
    if pil_img is None:
        # No se pudo capturar la imagen (Chrome/Clarity no disponible, etc.):
        # se deja un texto claro en vez del marcador crudo "##...##".
        return doc_xml.replace(marca, "(imagen de Clarity no disponible)")

    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    archivos[f"word/media/{nombre_media}"] = buf.getvalue()
    rid = _agregar_relacion_imagen(archivos, nombre_media)

    ancho_px, alto_px = pil_img.size
    cx = ancho_emu
    cy = int(cx * (alto_px / ancho_px)) if ancho_px else ancho_emu

    inicio_t = doc_xml.rfind("<w:t", 0, idx)
    fin_t = doc_xml.find("</w:t>", idx)
    if inicio_t < 0 or fin_t < 0:
        return doc_xml
    fin_t += len("</w:t>")

    dibujo = _drawing_xml(rid, cx, cy, doc_id, placeholder)
    return doc_xml[:inicio_t] + dibujo + doc_xml[fin_t:]


def _insertar_imagenes_clarity(archivos, doc_xml, clarity):
    if not clarity:
        return doc_xml
    doc_xml = _insertar_imagen_en_placeholder(
        archivos, doc_xml, "foto_clarity", clarity.get("imagen_resumen"),
        _CLARITY_ANCHO_EMU["foto_clarity"], 900001, "clarity_resumen.png",
    )
    doc_xml = _insertar_imagen_en_placeholder(
        archivos, doc_xml, "foto_clarity_url", clarity.get("imagen_paginas"),
        _CLARITY_ANCHO_EMU["foto_clarity_url"], 900002, "clarity_paginas.png",
    )
    return doc_xml


def generar_docx(datos, ruta_salida):
    with zipfile.ZipFile(PLANTILLA,"r") as zin:
        archivos = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = archivos["word/document.xml"].decode("utf-8")
    doc_xml = _colapsar_xml(doc_xml)

    # Duplicar fila FMI por cada inmueble
    fmis_lista = datos.get("fmis_grupo", [])
    if FILA_FMI in doc_xml and fmis_lista:
        filas_fmi = ""
        for item in fmis_lista:
            f = FILA_FMI
            f = f.replace("##fmi##",           _limpiar(item.get("fmi","—")))
            f = f.replace("##direccion##",      _limpiar(item.get("direccion","—")))
            f = f.replace("##ciudad##",         _limpiar(item.get("ciudad","—")))
            f = f.replace("##departamento##",   _limpiar(item.get("departamento","—")))
            f = f.replace("##tipo_inmueble##",  _limpiar(item.get("tipo_inmueble","—")))
            f = f.replace("##area##",           _limpiar(item.get("area","0")))
            filas_fmi += f
        doc_xml = doc_xml.replace(FILA_FMI, filas_fmi)

    # Duplicar fila participantes
    idx = doc_xml.find("##id_partipante##")
    if idx >= 0:
        inicio = doc_xml.rfind("<w:tr ", 0, idx)
        fin = doc_xml.find("</w:tr>", idx) + 7
        fila_plantilla = doc_xml[inicio:fin]
        filas_xml = ""
        for part in datos.get("participantes", []):
            f = fila_plantilla
            f = f.replace("##id_partipante##", _limpiar(part.get("id","—")))
            f = f.replace("##nombre##",         _limpiar(part.get("nombre","—")))
            f = f.replace("##correo##",          _limpiar(part.get("correo","—")))
            f = f.replace("##estado_registro##", _limpiar(part.get("estado_registro","—")))
            f = f.replace("##fecha_registro##",  _limpiar(part.get("fecha_registro","—")))
            filas_xml += f
        doc_xml = doc_xml[:inicio] + filas_xml + doc_xml[fin:]

    reemplazos = {
        "##codigo_subasta##":    datos["codigo_subasta"],
        "##fecha_publicacion##": datos["fecha_publicacion"],
        "##fecha_inicio##":      datos["fecha_inicio"],
        "##fecha_fin##":         datos["fecha_fin"],
        "##fmi##":               datos["fmi"],
        "##direccion##":         datos["direccion"],
        "##ciudad##":            datos["ciudad"],
        "##departamento##":      datos["departamento"],
        "##tipo_inmueble##":     datos["tipo_inmueble"],
        "##area##":              datos["area"],
        "##codigo_inmueble##":   datos["codigo_inmueble"],
        "##precio_base##":       datos["precio_base"],
        "##nombre_ganador##":    datos["nombre_ganador"],
        "##oferta_gandora##":    datos["oferta_gandora"],
        "##incremento##":        datos["incremento"],
        "##porcentaje_subasta##":datos["porcentaje_subasta"],
        "##incritos##":          datos["incritos"],
        "##ofertas##":           datos["ofertas"],
        "##total_ofertas##":     datos["total_ofertas"],
        "##url_pagina##":        datos["url_pagina"],
        "##clarity_sesiones_totales##": datos.get("clarity_sesiones_totales", "—"),
        "##clarity_bots_excluidos##":   datos.get("clarity_bots_excluidos", "—"),
        "##clarity_sesiones_url##":     datos.get("clarity_sesiones_url", "—"),
        "##cantidad##":          datos.get("cantidad", "0"),
        "##cantidad_sesiones##": datos.get("cantidad_sesiones", "0"),
        "##cantidad_url##":      datos.get("cantidad_url", "0"),
    }
    for i in range(1,17):
        reemplazos[f"##fecha_cronograma{i}##"] = datos.get(f"fecha_cronograma{i}","—")

    for ph, val in reemplazos.items():
        doc_xml = doc_xml.replace(ph, _limpiar(val))

    # Tabla de pujas
    idx_tabla = doc_xml.find("##tabla_pujas##")
    if idx_tabla >= 0:
        inicio_p = doc_xml.rfind("<w:p ", 0, idx_tabla)
        fin_p = doc_xml.find("</w:p>", idx_tabla) + 6
        doc_xml = doc_xml[:inicio_p] + _generar_tabla_pujas(datos.get("pujas",[])) + doc_xml[fin_p:]

    # Imágenes de tráfico web (Microsoft Clarity)
    doc_xml = _insertar_imagenes_clarity(archivos, doc_xml, datos.get("_clarity_imagenes"))

    archivos["word/document.xml"] = doc_xml.encode("utf-8")

    if "word/footer1.xml" in archivos:
        footer = archivos["word/footer1.xml"].decode("utf-8")
        footer = footer.replace("##fecha_publicacion##", _limpiar(datos["fecha_publicacion"]))
        footer = footer.replace("##codigo_subasta##",    _limpiar(datos["codigo_subasta"]))
        archivos["word/footer1.xml"] = footer.encode("utf-8")

    with zipfile.ZipFile(ruta_salida,"w",zipfile.ZIP_DEFLATED) as zout:
        for name, d in archivos.items():
            zout.writestr(name, d)


def main():
    if len(sys.argv) < 2:
        print("Uso: python generar_informe.py <auction_uuid | codigo_subasta | FMI | codigo_unidad>"); sys.exit(1)
    if not PLANTILLA.exists():
        print(f"✗ No se encuentra: {PLANTILLA.name}"); sys.exit(1)

    auction_input = sys.argv[1].strip()
    print(f"\n{'='*55}\n  GENERADOR DE INFORME DE SUBASTA\n{'='*55}\n  Subasta: {auction_input}\n")

    data, datos = conectar(auction_input)
    SALIDA.mkdir(exist_ok=True)
    codigo_corto = datos["codigo_subasta"].replace("ACTIBID-","")[:20]
    nombre_archivo = f"INFORME_{codigo_corto}.docx"
    ruta = SALIDA / nombre_archivo

    try:
        generar_docx(datos, ruta)
        print(f"\n✓ Informe generado: ./informes/{nombre_archivo}")
        print(f"  Código:      {datos['codigo_inmueble']}")
        print(f"  FMI:         {datos['fmi']}")
        print(f"  Ciudad:      {datos['ciudad']}")
        print(f"  Tipo:        {datos['tipo_inmueble']}")
        print(f"  Precio base: $ {datos['precio_base']}")
        print(f"  Ganador:     {datos['nombre_ganador']}")
        print(f"  Inscritos:   {datos['incritos']}")
        print(f"  Con ofertas: {datos['ofertas']}")
        print(f"  Total pujas: {datos['total_ofertas']}")
        print(f"  FMIs grupo:  {len(datos['fmis_grupo'])}")
    except Exception as e:
        print(f"✗ ERROR: {e}")
        import traceback; traceback.print_exc()

if __name__ == "__main__":
    main()
