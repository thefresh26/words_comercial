"""
Genera un ACTA DE ALCANCE (.docx) que complementa una Acta de Comité de
Adjudicación ANTIGUA con la información que el "Modelo Vigente" exige
agregar: Código ActiBid, valor catastral, resultado económico de la
subasta, identificación del oferente ganador, etc.

Uso:
    python generar_acta_alcance.py <acta_inicial.docx|.pdf> [identificador]

Donde:
  - <acta_inicial.docx|.pdf> es el Acta de Comité de Adjudicación ORIGINAL
    que se va a complementar. Puede ser el .docx o el .pdf (el formato en
    que realmente se guardan las actas viejas en el SharePoint/OneDrive).
    El script la LEE DIRECTAMENTE (no hay que transcribir nada a mano) y
    saca de ahí: el número y la fecha del acta vieja, el número de
    paquete, y la tabla de ID / FMI / Tipología / Dirección / Municipio /
    Estado de ocupación / Estado jurídico / Fecha subasta.
  - [identificador] (opcional): FMI, código de inmueble, código de subasta
    o UUID a usar para traer de la base de datos al oferente ganador y su
    identificación. Si no se indica, el script usa el primer FMI que haya
    encontrado dentro del acta inicial.

Requisitos (se buscan siempre relativos a la carpeta del proyecto):
  - templates/ACTA_DE_ALCANCE.docx          → la plantilla con ##marcadores##
  - templates/firmas/firma_heidy.png
  - templates/firmas/firma_jairo.png
  - templates/firmas/firma_jose.png
  - templates/firmas/firma_julio.png        → las 4 firmas a insertar
  - state/consecutivo_alcance.txt           → próximo número de acta de
    alcance (si no existe, se crea empezando en 81, tal como se pidió:
    "el consecutivo desde el paquete 1 comienza desde el número 81").
    Solo se avanza este contador cuando el documento se generó sin error.

IMPORTANTE — campos que este script NO puede completar automáticamente
todavía (ni el acta inicial ni la base de datos conocida los tienen) y que
quedan marcados como "[POR COMPLETAR]" en el .docx generado:
  - Valor catastral vigente
  - Precio Mínimo de Venta (PMV) y si fue superado
  - "Cronograma" del inmueble (ej. "Cronograma 30.1")
  - Fecha límite del compromiso de la Secretaría Técnica (queda con la que
    ya trae la plantilla, revisar si aplica para esta acta)
Al final de la ejecución el script imprime la lista de lo que quedó
pendiente de completar a mano para esa acta puntual.
"""
from __future__ import annotations

import copy
import re
import sys
from pathlib import Path

import docx
from docx.shared import Mm

PROYECTO       = Path(__file__).resolve().parent.parent
PLANTILLA      = PROYECTO / "templates" / "ACTA_DE_ALCANCE.docx"
CARPETA_FIRMAS = PROYECTO / "templates" / "firmas"
SALIDA         = PROYECTO / "output" / "actas_alcance"
ARCHIVO_CONSECUTIVO = PROYECTO / "state" / "consecutivo_alcance.txt"
CONSECUTIVO_INICIAL = 81

FIRMAS = {
    "heidy": CARPETA_FIRMAS / "firma_heidy.png",
    "jairo": CARPETA_FIRMAS / "firma_jairo.png",
    "jose":  CARPETA_FIRMAS / "firma_jose.png",
    "julio": CARPETA_FIRMAS / "firma_julio.png",
}

VALOR_POR_COMPLETAR = "[POR COMPLETAR]"


# ── Consecutivo ──────────────────────────────────────────────────────────

def _leer_consecutivo() -> int:
    if ARCHIVO_CONSECUTIVO.is_file():
        try:
            return int(ARCHIVO_CONSECUTIVO.read_text(encoding="utf-8").strip())
        except Exception:
            pass
    return CONSECUTIVO_INICIAL


def _guardar_consecutivo(siguiente: int) -> None:
    ARCHIVO_CONSECUTIVO.parent.mkdir(parents=True, exist_ok=True)
    ARCHIVO_CONSECUTIVO.write_text(str(siguiente), encoding="utf-8")


# ── Lectura del Acta inicial (vieja) ─────────────────────────────────────

def _normalizar_fecha_texto(texto: str) -> str:
    """'15 de mayo 2026' -> '15 de mayo de 2026' (agrega el 'de' antes del
    año si el acta vieja no lo trae, para que combine con el estilo del
    Acta de Alcance)."""
    m = re.match(r"(\d{1,2})\s+de\s+(\w+)\s+(?:de\s+)?(\d{4})", texto.strip(), re.IGNORECASE)
    if m:
        dia, mes, anio = m.groups()
        return f"{dia} de {mes} de {anio}"
    return texto.strip()


def _fecha_hoy_es() -> str:
    from datetime import date
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio",
             "agosto","septiembre","octubre","noviembre","diciembre"]
    hoy = date.today()
    return f"{hoy.day:02d} de {meses[hoy.month - 1]} de {hoy.year}"


def _listar_con_y(items: list[str]) -> str:
    items = [i for i in items if i and i != "—"]
    if not items:
        return "—"
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " y " + items[-1]


def _extraer_campos_comunes(texto_parrafos: str, tablas: list) -> dict:
    """Lógica de extracción compartida entre Word y PDF. 'tablas' es una
    lista de tablas, cada una una lista de filas, cada fila una lista de
    strings de celda (ya puede traer saltos de línea internos)."""

    # Número del acta vieja (ej. "No.025" o "No. 025")
    m_num = re.search(r"ADJUDICAC\w*.*?No\.?\s*0*(\d+)", texto_parrafos, re.IGNORECASE | re.DOTALL)
    if not m_num:
        m_num = re.search(r"\bNo\.?\s*0*(\d{2,4})\b", texto_parrafos)
    numero_acta_vieja = m_num.group(1) if m_num else "—"

    # Fecha (fila "Fecha | <fecha> | Hora | ...")
    fecha_acta_vieja = "—"
    for tabla in tablas:
        for fila in tabla:
            celdas = [(c or "").strip() for c in fila]
            if celdas and celdas[0].strip().lower() == "fecha" and len(celdas) > 1:
                candidato = celdas[1].strip()
                if candidato and candidato.lower() != "fecha":
                    fecha_acta_vieja = candidato
                    break
        if fecha_acta_vieja != "—":
            break

    # Número de paquete
    m_paq = re.search(r"Paquete\s*No\.?\s*(\d+)", texto_parrafos, re.IGNORECASE)
    numero_paquete = m_paq.group(1) if m_paq else "—"

    # Tabla de inmuebles (encabezado con ID / FMI / Tipología / ...). En PDF
    # el encabezado a veces viene partido en 2 filas (texto envuelto), así
    # que se busca entre las primeras filas de cada tabla, no solo la [0].
    inmuebles = []
    for tabla in tablas:
        if not tabla:
            continue
        idx_encabezado = None
        fila_encabezado = None
        for i, fila in enumerate(tabla[:3]):
            celdas_enc = [(c or "").strip().lower() for c in fila]
            if "fmi" in celdas_enc and any("tipolog" in c or c == "id" for c in celdas_enc):
                idx_encabezado = i
                fila_encabezado = fila
                break
        if idx_encabezado is None:
            continue

        # El encabezado a veces viene partido en varias filas (texto
        # envuelto en el PDF, ej. "Estado de" / "ocupación" en filas
        # separadas), seguidas de filas de relleno vacías, antes de llegar
        # a la fila real de datos. Se van fusionando esas filas de
        # encabezado hasta encontrar la primera fila con datos (contiene
        # dígitos, ej. el ID o el FMI).
        filas_encabezado = [fila_encabezado]
        j = idx_encabezado + 1
        while j < len(tabla) and (j - idx_encabezado) <= 5:
            fila_j = tabla[j]
            celdas_j = [(c or "").strip() for c in fila_j]
            no_vacias = [c for c in celdas_j if c]
            if not no_vacias:
                filas_encabezado.append(fila_j)
                j += 1
                continue
            tiene_digitos = any(ch.isdigit() for ch in " ".join(no_vacias))
            if not tiene_digitos and all(len(c) <= 20 for c in no_vacias):
                filas_encabezado.append(fila_j)
                j += 1
                continue
            break
        idx_datos_inicio = j

        num_columnas = max(len(f) for f in filas_encabezado)
        encabezado = []
        for col_idx in range(num_columnas):
            partes = []
            for fila_h in filas_encabezado:
                if col_idx < len(fila_h) and fila_h[col_idx]:
                    partes.append(str(fila_h[col_idx]).strip())
            encabezado.append(" ".join(partes).strip().lower())

        idx = {nombre: i for i, nombre in enumerate(encabezado) if nombre}

        def _col(claves):
            for clave in claves:
                for nombre_col, i in idx.items():
                    if clave in nombre_col:
                        return i
            return None

        col_id      = _col(["id"])
        col_fmi     = _col(["fmi"])
        col_tipo    = _col(["tipolog"])
        col_dir     = _col(["direcci"])
        col_mun     = _col(["municipio"])
        col_ocup    = _col(["ocupaci"])
        col_jur     = _col(["jur"])
        col_fsub    = _col(["subasta"])

        for fila in tabla[idx_datos_inicio:]:
            celdas = [(c or "").replace("\n", " ").strip() for c in fila]
            if not any(celdas):
                continue
            # descartar filas residuales del encabezado partido (repiten
            # "ocupación"/"jurídico"/"subasta" pero sin datos reales de ID/FMI)
            valor_id = celdas[col_id] if col_id is not None and col_id < len(celdas) else ""
            valor_fmi = celdas[col_fmi] if col_fmi is not None and col_fmi < len(celdas) else ""
            if not valor_id and not valor_fmi:
                continue

            def _val(col, quitar_espacios=False):
                if col is None or col >= len(celdas) or not celdas[col]:
                    return "—"
                valor = celdas[col]
                return valor.replace(" ", "") if quitar_espacios else valor

            inmuebles.append({
                "id":               _val(col_id),
                "fmi":              _val(col_fmi, quitar_espacios=True),
                "tipologia":        _val(col_tipo),
                "direccion":        _val(col_dir),
                "municipio":        _val(col_mun),
                "estado_ocupacion": _val(col_ocup),
                "estado_juridico":  _val(col_jur),
                "fecha_subasta":    _val(col_fsub),
            })
        if inmuebles:
            break

    return {
        "numero_acta_vieja": numero_acta_vieja,
        "fecha_acta_vieja":  fecha_acta_vieja,
        "numero_paquete":    numero_paquete,
        "inmuebles":         inmuebles,
    }


def leer_acta_inicial(ruta: Path) -> dict:
    """Lee el Acta inicial desde un .docx."""
    doc = docx.Document(str(ruta))
    texto_parrafos = "\n".join(p.text for p in doc.paragraphs)
    tablas = [[[c.text for c in fila.cells] for fila in tabla.rows] for tabla in doc.tables]
    return _extraer_campos_comunes(texto_parrafos, tablas)


def leer_acta_inicial_pdf(ruta: Path) -> dict:
    """Lee el Acta inicial desde un .pdf (el formato real en que están
    guardadas las actas en el SharePoint/OneDrive)."""
    import pdfplumber

    texto_partes = []
    tablas = []
    with pdfplumber.open(str(ruta)) as pdf:
        for pagina in pdf.pages:
            texto_partes.append(pagina.extract_text() or "")
            tablas.extend(pagina.extract_tables())
    texto_parrafos = "\n".join(texto_partes)
    return _extraer_campos_comunes(texto_parrafos, tablas)


def leer_acta_inicial_archivo(ruta: Path) -> dict:
    """Detecta la extensión y usa el lector correcto (.docx o .pdf)."""
    if ruta.suffix.lower() == ".pdf":
        return leer_acta_inicial_pdf(ruta)
    return leer_acta_inicial(ruta)


# ── Conexión y datos de base de datos ────────────────────────────────────

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
        password = input("Contraseña del vault: ").strip()
        if password:
            blob = read_vault(vault_path)
            creds = decrypt_payload(blob, password)
            db = creds.get("database") or creds
            dsn = dsn_from_database_section(db)

    conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
    return conn


def _resolver_identificador(conn, identificador: str) -> str:
    """Igual patrón que en los demás scripts: UUID / código de subasta /
    FMI / código de inmueble / unidad. Si hay varias subastas, pregunta."""
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

    if len(candidatos) == 1:
        return str(candidatos[0][0])

    print(f"\n⚠ Se encontraron {len(candidatos)} subastas asociadas a '{identificador}':\n")
    for i, (aid, code, status, start) in enumerate(candidatos, start=1):
        print(f"  [{i}] {code or aid}  |  estado: {status or '—'}  |  inicio: {start}")
    while True:
        seleccion = input(f"\nElige el número de subasta a usar (1-{len(candidatos)}): ").strip()
        if seleccion.isdigit() and 1 <= int(seleccion) <= len(candidatos):
            return str(candidatos[int(seleccion) - 1][0])
        print("  ⚠ Opción inválida, intenta de nuevo.")


def obtener_datos_bd(conn, auction_uuid: str) -> dict:
    """Trae de la base de datos SOLO lo que el usuario pidió que viniera de
    ahí: código ActiBid del inmueble y la identificación del oferente
    ganador con su puja. El resto de datos del inmueble (tipología,
    dirección, municipio, estado...) se toman del Acta inicial, porque
    esas columnas no existen hoy en ninguna tabla ya usada por los otros
    scripts de este proyecto."""
    sys.path.insert(0, str(PROYECTO / "core"))
    sys.path.insert(0, str(PROYECTO))
    from core import fetch_informe

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


def _fmt_numero(valor) -> str:
    if not valor:
        return "0"
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return str(valor)


# ── Generación del .docx ─────────────────────────────────────────────────

def _duplicar_fila(tabla, indice: int, veces: int):
    """Duplica tabla.rows[indice] 'veces' veces más, insertándolas justo
    después de la original. Devuelve la lista de filas (incluida la
    original) en orden."""
    tr_referencia = tabla.rows[indice]._tr
    ultimo_tr = tr_referencia
    for _ in range(veces):
        nuevo_tr = copy.deepcopy(tr_referencia)
        ultimo_tr.addnext(nuevo_tr)
        ultimo_tr = nuevo_tr
    return list(tabla.rows[indice: indice + 1 + veces])


def _reemplazar_texto_parrafo(parrafo, mapa: dict) -> None:
    texto_completo = parrafo.text
    if not texto_completo or not any(marcador in texto_completo for marcador in mapa):
        return
    nuevo_texto = texto_completo
    for marcador, valor in mapa.items():
        nuevo_texto = nuevo_texto.replace(marcador, str(valor))
    if parrafo.runs:
        parrafo.runs[0].text = nuevo_texto
        for run in parrafo.runs[1:]:
            run.text = ""
    else:
        parrafo.add_run(nuevo_texto)


def _reemplazar_en_celda(celda, mapa: dict) -> None:
    for p in celda.paragraphs:
        _reemplazar_texto_parrafo(p, mapa)


def _insertar_firma(celda, ruta_imagen: Path, ancho_mm: float = 32) -> bool:
    if not ruta_imagen.is_file():
        return False
    for p in celda.paragraphs:
        if "##firma_" in p.text:
            for run in p.runs:
                run.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            run.add_picture(str(ruta_imagen), width=Mm(ancho_mm))
            return True
    return False


def generar_docx(acta_vieja: dict, bd: dict, num_consecutivo_actual: int, ruta_salida: Path,
                  excel: dict | None = None) -> list[str]:
    """Genera el .docx final. Devuelve la lista de mensajes de "pendiente
    de completar a mano" para esta acta puntual.

    'excel' (opcional) trae datos del Excel "FORMATO_Paquete_N_2026..." de
    la biblioteca 2026_Paquetes: valor_catastral, vigencia_catastral, pmv.
    Si no se pasa (o el inmueble no se encontró en el Excel), esos campos
    quedan como [POR COMPLETAR], igual que antes."""
    doc = docx.Document(str(PLANTILLA))
    pendientes = []
    excel = excel or {}

    inmuebles = acta_vieja["inmuebles"] or [{
        "id": "—", "fmi": "—", "tipologia": "—", "direccion": "—", "municipio": "—",
        "estado_ocupacion": "—", "estado_juridico": "—", "fecha_subasta": "—",
    }]
    fecha_normalizada = _normalizar_fecha_texto(acta_vieja["fecha_acta_vieja"])

    if len(inmuebles) == 1:
        descripcion = ", ".join(
            x for x in [inmuebles[0]["tipologia"], inmuebles[0]["direccion"], inmuebles[0]["municipio"]]
            if x and x != "—"
        )
    else:
        descripcion = "; ".join(
            f"{i['tipologia']} {i['direccion']}".strip()
            for i in inmuebles if i.get("direccion") and i["direccion"] != "—"
        )
    paquete_y_direccion = f"Paquete No. {acta_vieja['numero_paquete']}"
    if descripcion:
        paquete_y_direccion += f" ({descripcion})"

    # Valor catastral vigente y PMV: vienen del Excel "FORMATO_Paquete_N_2026..."
    # cuando se encontró el inmueble ahí; si no, quedan pendientes.
    if excel.get("valor_catastral") not in (None, ""):
        vigencia = excel.get("vigencia_catastral")
        valor_catastral_txt = f"$ {_fmt_numero(excel['valor_catastral'])}"
        if vigencia:
            valor_catastral_txt += f" (vigencia {vigencia})"
    else:
        valor_catastral_txt = VALOR_POR_COMPLETAR

    pmv_crudo = excel.get("pmv")
    if pmv_crudo not in (None, ""):
        try:
            pmv_numero = float(str(pmv_crudo).replace("$", "").replace(",", "").strip())
            precio_minimo_txt = f"$ {_fmt_numero(pmv_numero)}"
        except (ValueError, TypeError):
            pmv_numero = None
            precio_minimo_txt = str(pmv_crudo).strip()
    else:
        pmv_numero = None
        precio_minimo_txt = VALOR_POR_COMPLETAR

    try:
        monto_ganador_num = float(str(bd.get("monto_ganador", "0")).replace(".", "").replace(",", "."))
    except (ValueError, TypeError):
        monto_ganador_num = None

    if pmv_numero is not None and monto_ganador_num is not None:
        supero_txt = "Sí" if monto_ganador_num >= pmv_numero else "No"
    else:
        supero_txt = VALOR_POR_COMPLETAR

    # Cronograma: no existe un dato de texto por inmueble individual (solo
    # existen imágenes tipo Gantt por paquete completo), así que se usa una
    # referencia genérica por PAQUETE en vez de dejarlo pendiente.
    numero_paquete = acta_vieja.get("numero_paquete", "—")
    if numero_paquete and numero_paquete != "—":
        cronograma_txt = f"Cronograma Paquete No. {numero_paquete}"
    else:
        cronograma_txt = VALOR_POR_COMPLETAR

    mapa_texto = {
        "##num_consecutivo_actual##":        str(num_consecutivo_actual),
        "##num_consecutivo_acta_existente##": f"No. {acta_vieja['numero_acta_vieja']}",
        "##fecha_acta_vieja##":              f"del {fecha_normalizada}" if fecha_normalizada != "—" else "—",
        "##fmis_acta_aterior##":             _listar_con_y([i["fmi"] for i in inmuebles]),
        "##paquete_y_direccion##":           paquete_y_direccion,
        "##fecha_hoy##":                     _fecha_hoy_es(),
        "##codigo_inmueble##":               bd["codigo_inmueble"],
        "##fmi##":                           _listar_con_y([i["fmi"] for i in inmuebles]),
        "##fmi_catastral##":                 valor_catastral_txt,
        "##fmi_catastral ##":                valor_catastral_txt,
        "##cronograma_inmueble##":           cronograma_txt,
        "##tipologia##":                     inmuebles[0]["tipologia"],
        "##direccion_fmi##":                 inmuebles[0]["direccion"],
        "##municipio_fmi##":                 inmuebles[0]["municipio"],
        "##estado##":                        inmuebles[0]["estado_ocupacion"],
        "##estado_juridico##":               inmuebles[0]["estado_juridico"],
        "##fecha_subasta##":                 inmuebles[0]["fecha_subasta"],
        "##precio_minimo##":                 precio_minimo_txt,
        "##oferta_ganadora##":               f"$ {bd['monto_ganador']}",
        "##¿supero_precio_base?##":          supero_txt,
        "##nombre_participante##":           bd["nombre_ganador"],
        "##nit/cedula##":                    bd["id_ganador"],
        "##puja_ganadora##":                 f"$ {bd['monto_ganador']}",
        "##valor_cerrado_subasta##":         f"$ {bd['monto_ganador']}",
    }

    if bd["codigo_inmueble"] == "—":
        pendientes.append("Código ActiBid: no se encontró en la base de datos, revisar a mano.")
    if bd["nombre_ganador"] == "—":
        pendientes.append("Oferente ganador: no se encontró en la base de datos, revisar a mano.")
    if valor_catastral_txt == VALOR_POR_COMPLETAR:
        pendientes.append("Valor catastral vigente: no se encontró el inmueble en el Excel del paquete, revisar a mano.")
    if precio_minimo_txt == VALOR_POR_COMPLETAR:
        pendientes.append("Precio Mínimo de Venta (PMV): no se encontró el inmueble en el Excel del paquete, revisar a mano.")
    if supero_txt == VALOR_POR_COMPLETAR:
        pendientes.append("¿Superó el precio base?: no se pudo calcular (falta PMV o monto adjudicado), revisar a mano.")
    if cronograma_txt == VALOR_POR_COMPLETAR:
        pendientes.append("Cronograma del inmueble: no se encontró el número de paquete, revisar a mano.")
    else:
        pendientes.append(f"Cronograma del inmueble: se puso '{cronograma_txt}' (referencia genérica por paquete, no hay dato por inmueble individual); verificar si aplica o si hay una fase específica.")
    pendientes.append("Fecha límite del compromiso de la Secretaría Técnica: revisar si la de la plantilla aplica.")

    # Ajustes puntuales de texto literal (no son ##marcadores##) que trae la
    # plantilla del usuario, hardcodeados con los datos del ejemplo "019":
    #  - "Acta No. ##marcador##" -> evitar que quede "No. No. 25" cuando el
    #    valor del marcador ya incluye su propio "No.".
    #  - "acta inicial No. 019" -> reemplazar el 019 por el número real.
    def _ajustes_texto_literal(texto: str) -> str:
        texto = texto.replace(
            "Acta No. ##num_consecutivo_acta_existente##",
            "Acta ##num_consecutivo_acta_existente##",
        )
        texto = texto.replace(
            "acta inicial No. 019", f"acta inicial No. {acta_vieja['numero_acta_vieja']}"
        )
        texto = texto.replace(
            "Acta inicial No. 019", f"Acta inicial No. {acta_vieja['numero_acta_vieja']}"
        )
        return texto

    def _aplicar_ajustes_parrafo(p) -> None:
        texto_original = p.text
        texto_nuevo = _ajustes_texto_literal(texto_original)
        if texto_nuevo != texto_original and p.runs:
            p.runs[0].text = texto_nuevo
            for r in p.runs[1:]:
                r.text = ""

    for p in doc.paragraphs:
        _aplicar_ajustes_parrafo(p)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    _aplicar_ajustes_parrafo(p)

    # Tabla de inmuebles (ID/FMI/Tipología/...): duplicar filas si hay más de un inmueble.
    tabla_inmuebles = None
    for tabla in doc.tables:
        encabezado = [c.text.strip().lower() for c in tabla.rows[0].cells] if tabla.rows else []
        if "fmi" in encabezado and "tipología" in encabezado:
            tabla_inmuebles = tabla
            break

    if tabla_inmuebles is not None and len(inmuebles) > 1:
        filas = _duplicar_fila(tabla_inmuebles, 1, len(inmuebles) - 1)
        for fila, inm in zip(filas, inmuebles):
            fila.cells[0].text = inm["id"]
            fila.cells[1].text = inm["fmi"]
            fila.cells[2].text = inm["tipologia"]
            fila.cells[3].text = inm["direccion"]
            fila.cells[4].text = inm["municipio"]
            fila.cells[5].text = inm["estado_ocupacion"]
            fila.cells[6].text = inm["estado_juridico"]
            fila.cells[7].text = inm["fecha_subasta"]
    elif tabla_inmuebles is not None:
        fila = tabla_inmuebles.rows[1]
        fila.cells[0].text = inmuebles[0]["id"]
        # el resto (FMI/tipología/etc.) se llena vía mapa_texto más abajo,
        # ya que esa fila sigue teniendo los ##marcadores## originales.

    # Reemplazo general de texto en todo el documento (excepto celdas de firma).
    for p in doc.paragraphs:
        _reemplazar_texto_parrafo(p, mapa_texto)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "##firma_" in celda.text:
                    continue
                _reemplazar_en_celda(celda, mapa_texto)

    # Firmas (imagen) en la tabla de "FIRMA DE ASISTENTES".
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for persona, ruta_img in FIRMAS.items():
                    marcador = f"##firma_{persona}##"
                    if marcador in celda.text:
                        if not _insertar_firma(celda, ruta_img):
                            pendientes.append(f"Firma de {persona}: no se pudo insertar la imagen, revisar a mano.")

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ruta_salida))
    return pendientes


# ── Main ───────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Uso: python generar_acta_alcance.py <acta_inicial.docx|.pdf> [identificador]")
        sys.exit(1)

    ruta_acta_vieja = Path(sys.argv[1])
    if not ruta_acta_vieja.is_file():
        print(f"✗ No se encontró el archivo: {ruta_acta_vieja}")
        sys.exit(1)

    if ruta_acta_vieja.suffix.lower() not in (".docx", ".pdf"):
        print(f"✗ Formato no soportado ({ruta_acta_vieja.suffix}). Usa .docx o .pdf")
        sys.exit(1)

    if not PLANTILLA.is_file():
        print(f"✗ No se encuentra la plantilla: {PLANTILLA}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print("  GENERADOR DE ACTA DE ALCANCE")
    print(f"{'='*60}")
    print(f"  Acta inicial: {ruta_acta_vieja.name}\n")

    print("Leyendo el acta inicial...")
    acta_vieja = leer_acta_inicial_archivo(ruta_acta_vieja)
    print(f"  Número acta vieja: {acta_vieja['numero_acta_vieja']}")
    print(f"  Fecha acta vieja:  {acta_vieja['fecha_acta_vieja']}")
    print(f"  Paquete No.:       {acta_vieja['numero_paquete']}")
    print(f"  Inmuebles encontrados: {len(acta_vieja['inmuebles'])}")
    for inm in acta_vieja["inmuebles"]:
        print(f"    - ID {inm['id']} | FMI {inm['fmi']} | {inm['tipologia']} | {inm['direccion']} | {inm['municipio']}")

    if not acta_vieja["inmuebles"]:
        print("\n⚠ No se pudo encontrar la tabla de inmuebles dentro del acta inicial.")
        print("  Revisa que el archivo tenga una tabla con columnas ID/FMI/Tipología/...")

    identificador = sys.argv[2].strip() if len(sys.argv) > 2 else None
    if not identificador:
        if acta_vieja["inmuebles"] and acta_vieja["inmuebles"][0]["fmi"] != "—":
            identificador = acta_vieja["inmuebles"][0]["fmi"]
            print(f"\n  (usando el FMI del acta inicial para buscar en la base de datos: {identificador})")
        else:
            print("\n✗ No se pudo determinar un FMI/identificador para buscar en la base de datos.")
            print("  Indícalo como segundo argumento: python generar_acta_alcance.py <acta.docx> <FMI>")
            sys.exit(1)

    print("\nConectando a la base de datos...")
    conn = _conectar()
    print("✓ Conectado")
    try:
        auction_uuid = _resolver_identificador(conn, identificador)
        bd = obtener_datos_bd(conn, auction_uuid)
    finally:
        conn.close()

    print(f"\n  Subasta:          {bd['codigo_subasta']}")
    print(f"  Código ActiBid:   {bd['codigo_inmueble']}")
    print(f"  Oferente ganador: {bd['nombre_ganador']} ({bd['id_ganador']})")
    print(f"  Valor adjudicado: $ {bd['monto_ganador']}")

    consecutivo = _leer_consecutivo()
    print(f"\nConsecutivo de esta acta de alcance: {consecutivo}")

    codigo_archivo = re.sub(r"[^A-Za-z0-9]+", "_", identificador)
    SALIDA.mkdir(parents=True, exist_ok=True)
    ruta_salida = SALIDA / f"ACTA_DE_ALCANCE_{consecutivo}_{codigo_archivo}.docx"

    try:
        pendientes = generar_docx(acta_vieja, bd, consecutivo, ruta_salida)
    except Exception as e:
        print(f"\n✗ ERROR generando el documento: {e}")
        sys.exit(1)

    _guardar_consecutivo(consecutivo + 1)

    print(f"\n✓ Acta de alcance generada: {ruta_salida}")
    print(f"  (el próximo consecutivo a usar será {consecutivo + 1})")
    if pendientes:
        print("\n⚠ Pendiente de revisar/completar a mano en este documento:")
        for msg in pendientes:
            print(f"    - {msg}")


if __name__ == "__main__":
    main()
