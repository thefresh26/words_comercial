"""
Genera un ACTA DE CERTIFICACIÓN DE SUBASTA ELECTRÓNICA (.docx)
usando el archivo original como plantilla.

Uso:
    python generar_acta.py <identificador>

Donde <identificador> puede ser:
    - El UUID de la subasta.
    - El código de la subasta (ej. ACTIBID-POLI-1-...).
    - El FMI / número de matrícula de un inmueble individual.
    - El código de una unidad inmobiliaria (ej. UNI-0090-2025).

Si el FMI o la unidad tienen más de una subasta asociada, el script las
lista y pide elegir cuál usar.

Requisito: ACTA_DE_CERTIFICACIÓN_DE_SUBASTA_ELECTRÓNICA.docx en la misma carpeta.
El archivo se guarda en ./actas/
"""

from __future__ import annotations

import re
import sys
import time
import zipfile
from pathlib import Path

PROYECTO = Path(__file__).parent
PLANTILLA = PROYECTO.parent / "templates" / "ACTA_DE_CERTIFICACIÓN_DE_SUBASTA_ELECTRÓNICA.docx"
SALIDA    = PROYECTO.parent / "output" / "actas"
# ── Slug y scraping ───────────────────────────────────────────────────────────


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


def _codigo_unidad_de_grupo(conn, grupo_id_val) -> str | None:
    """Dado un grupo_id, devuelve el código de la unidad inmobiliaria
    (formato "UNI-XXXX-AAAA") tomado de mst_inmuebles.codigo_grupo, o None
    si el grupo no existe o ningún registro tiene ese código."""
    if not grupo_id_val:
        return None
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT codigo_grupo
            FROM mst_inmuebles
            WHERE grupo_id = %s AND codigo_grupo IS NOT NULL
            ORDER BY es_padre DESC, id
            LIMIT 1
            """,
            (grupo_id_val,),
        )
        row = cur.fetchone()
        return row[0] if row and row[0] else None


def _resolver_identificador(conn, identificador: str) -> str:
    """Acepta UUID de subasta, código de subasta, FMI/número de matrícula
    de un inmueble individual, código de un inmueble, o código de una
    unidad inmobiliaria (UNI-XXXX-AAAA), y devuelve el UUID de la subasta
    a usar.

    Si el identificador corresponde a un FMI/unidad con varias subastas
    asociadas, se le muestran las opciones al usuario y se le pide elegir
    (por consola), ya que puede haber más de una subasta para el mismo
    inmueble o unidad.
    """
    identificador = identificador.strip()

    # 1. UUID directo o código de subasta (comportamiento original)
    try:
        return _resolver_uuid(conn, identificador)
    except ValueError:
        pass

    # 2. FMI / código de inmueble / código de unidad inmobiliaria
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


def _fmt_fecha_hora(dt) -> str:
    if not dt:
        return "—"
    try:
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(dt)


def _slugify(texto: str) -> str:
    texto = texto.lower()
    for k, v in {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ü":"u"}.items():
        texto = texto.replace(k, v)
    texto = re.sub(r"[\s\-]+", "-", texto)
    texto = re.sub(r"[^a-z0-9\-]", "", texto)
    return texto.strip("-")


def _scrape_fechas(grupo_id, nombre_grupo: str, inm_id=None) -> dict:
    """Extrae fecha de publicación, apertura y cierre desde la página web."""
    resultado = {"publicacion": "—", "apertura": "—", "cierre": "—", "direccion": "—"}
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from webdriver_manager.chrome import ChromeDriverManager
        import re as _re

        if grupo_id:
            slug = _slugify(nombre_grupo) if nombre_grupo else ""
            url = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/{slug}"
        else:
            slug = _slugify(nombre_grupo) if nombre_grupo else ""
            url = f"https://activosporcolombia.com/es/inmueble/{inm_id}/{slug}"

        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-images")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.add_argument("--window-size=1600,1000")
        options.page_load_strategy = "eager"  # no espera a que carguen imágenes/recursos secundarios
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        print(f"  → Consultando: {url}")
        # cache_valid_range evita que webdriver_manager revise por internet si
        # hay una versión nueva del driver en cada corrida — usa la que ya
        # tiene descargada mientras no pasen 30 días, ahorrando varios segundos.
        try:
            driver_path = ChromeDriverManager(cache_valid_range=30).install()
        except TypeError:
            driver_path = ChromeDriverManager().install()
        driver = webdriver.Chrome(service=Service(driver_path), options=options)
        try:
            driver.set_page_load_timeout(15)
            driver.get(url)

            # Igual que en generar_informe.py: el cronograma carga sus datos
            # por JS después del render inicial, así que se reintenta leer el
            # texto de la página hasta ver el cronograma con datos reales
            # ("COMPLETADO"/fechas), en vez de confiar en un sleep fijo corto
            # que puede capturar el texto antes de que lleguen esos datos.
            espera_max = 6.0
            paso = 0.3
            transcurrido = 0.0
            lineas = []
            while transcurrido < espera_max:
                time.sleep(paso)
                transcurrido += paso
                lineas = driver.find_element(By.TAG_NAME, "body").text.split("\n")
                texto_actual = "\n".join(lineas)
                if "Cronograma del proceso" in texto_actual and (
                    "COMPLETADO" in texto_actual.upper() or "→" in texto_actual
                ):
                    break
            print(f"  [debug] Esperó {transcurrido:.1f}s a que cargara el cronograma "
                  f"({len(lineas)} líneas capturadas)")

            meses_abr = {
                "ene":"01","feb":"02","mar":"03","abr":"04","may":"05","jun":"06",
                "jul":"07","ago":"08","sep":"09","sept":"09","oct":"10","nov":"11","dic":"12"
            }

            # Índice donde empieza el cronograma (para no buscar el año ni las
            # fechas en cualquier parte de la página, solo en esa sección).
            inicio_cron = next(
                (i for i, l in enumerate(lineas) if "Cronograma del proceso" in l), 0
            )
            print(f"  [debug] 'Cronograma del proceso' encontrado en la línea {inicio_cron}")
            print("  [debug] Líneas desde el cronograma (índice: texto):")
            for j in range(inicio_cron, min(inicio_cron + 60, len(lineas))):
                print(f"    [{j}] {lineas[j].strip()!r}")

            # Buscar año SOLO dentro de la sección del cronograma: si se busca
            # en toda la página se puede agarrar por error un año que no tiene
            # nada que ver (p. ej. un año de construcción u otro dato numérico
            # del inmueble que aparezca antes en el texto).
            anio = "2026"
            for l in lineas[inicio_cron:]:
                m = _re.search(r"20\d{2}", l)
                if m:
                    anio = m.group()
                    break

            def _siguiente_no_vacia(desde: int) -> int | None:
                """Índice de la siguiente línea no vacía a partir de 'desde'
                (el sitio deja líneas en blanco entre el nombre de la fase,
                el rango de fechas y el mes)."""
                j = desde
                while j < len(lineas) and not lineas[j].strip():
                    j += 1
                return j if j < len(lineas) else None

            # Fecha publicación: fase 1 del cronograma. El sitio tiene dos
            # variantes: las páginas de unidad inmobiliaria todavía muestran
            # "Publicación próxima en subasta" como fase 1; las páginas de
            # inmueble individual le quitaron esa fase y ahora empiezan
            # directo en "Registro y cargue de documentos". En vez de atarse
            # al nombre exacto de la fase 1, se usa siempre la fecha de la
            # fase que abre el cronograma, sea cual sea su nombre.
            idx_fase1_nombre = None
            for i, linea in enumerate(lineas):
                if i < inicio_cron:
                    continue
                if "Publicación próxima en subasta" in linea or "Registro y cargue de documentos" in linea:
                    idx_fase1_nombre = i
                    break

            if idx_fase1_nombre is not None:
                try:
                    idx_rango = _siguiente_no_vacia(idx_fase1_nombre + 1)
                    idx_mes = _siguiente_no_vacia(idx_rango + 1) if idx_rango is not None else None
                    rango = lineas[idx_rango].strip()
                    mes_l = lineas[idx_mes].strip()
                    dia = rango.split()[0].zfill(2)
                    mes = meses_abr.get(mes_l.split(".")[0].strip(), "00")
                    resultado["publicacion"] = f"{dia}/{mes}/{anio} 10:00 am"
                    print(f"  → Publicación: {resultado['publicacion']}")
                except Exception:
                    pass

            # Fechas apertura y cierre: después de "Estado de la Subasta"
            encontre_estado = False
            fechas_subasta = []
            for linea in lineas:
                linea = linea.strip()
                if "Estado de la Subasta" in linea:
                    encontre_estado = True
                    continue
                if encontre_estado and "de 20" in linea and "a las" in linea:
                    fecha_fmt = _convertir_fecha_espanol(linea)
                    if fecha_fmt:
                        fechas_subasta.append(fecha_fmt)
                    if len(fechas_subasta) == 2:
                        break

            if len(fechas_subasta) >= 1:
                resultado["apertura"] = fechas_subasta[0]
                print(f"  → Apertura: {resultado['apertura']}")
            if len(fechas_subasta) >= 2:
                resultado["cierre"] = fechas_subasta[1]
                print(f"  → Cierre: {resultado['cierre']}")

            # Respaldo: si no hay sección "Estado de la Subasta" (pasa en
            # subastas ya finalizadas, el sitio deja de mostrarla), se usa la
            # última tarjeta del cronograma, "Subasta (apertura y cierre)"
            # — se busca la frase completa (no solo "Subasta") para no
            # confundirla con la fase 1 ("Publicación próxima en subasta").
            if resultado["apertura"] == "—" or resultado["cierre"] == "—":
                for i, linea in enumerate(lineas):
                    if i < inicio_cron:
                        continue
                    if "subasta (apertura y cierre)" in linea.lower():
                        try:
                            idx_rango = _siguiente_no_vacia(i + 1)
                            idx_mes = _siguiente_no_vacia(idx_rango + 1) if idx_rango is not None else None
                            rango = lineas[idx_rango].strip()
                            mes_l = lineas[idx_mes].strip()
                            mes = meses_abr.get(mes_l.split(".")[0].strip(), "00")
                            if "→" in rango:
                                partes = rango.split()
                                dia_i, dia_f = partes[0].zfill(2), partes[2].zfill(2)
                            else:
                                dia_i = dia_f = rango.split()[0].zfill(2)
                            if resultado["apertura"] == "—":
                                resultado["apertura"] = f"{dia_i}/{mes}/{anio}"
                                print(f"  [debug] Apertura tomada de la fase 'Subasta (apertura y cierre)': {resultado['apertura']}")
                            if resultado["cierre"] == "—":
                                resultado["cierre"] = f"{dia_f}/{mes}/{anio}"
                                print(f"  [debug] Cierre tomado de la fase 'Subasta (apertura y cierre)': {resultado['cierre']}")
                        except Exception:
                            pass
                        break

            # Extraer dirección física de la sección Ubicación
            for i, linea in enumerate(lineas):
                if "Dirección:" in linea:
                    dir_texto = linea.replace("Dirección:", "").strip()
                    if dir_texto:
                        resultado["direccion"] = dir_texto
                        print(f"  → Dirección: {dir_texto}")
                        break

        finally:
            driver.quit()
    except Exception as e:
        print(f"  ⚠ No se pudo obtener fechas: {e}")
    return resultado


def _convertir_fecha_espanol(texto: str) -> str:
    meses = {
        "enero":"01","febrero":"02","marzo":"03","abril":"04",
        "mayo":"05","junio":"06","julio":"07","agosto":"08",
        "septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"
    }
    try:
        partes = texto.lower().split()
        dia  = partes[0].zfill(2)
        mes  = meses.get(partes[2], "00")
        anio = partes[4]
        hora = partes[7] if len(partes) > 7 else "10:00"
        ampm = "".join(partes[8:]).replace(".", "") if len(partes) > 8 else "am"
        return f"{dia}/{mes}/{anio} {hora} {ampm}"
    except Exception:
        return None


# ── Conexión ──────────────────────────────────────────────────────────────────

def conectar(auction_uuid: str):
    sys.path.insert(0, str(PROYECTO.parent / "core"))
    sys.path.insert(0, str(PROYECTO.parent))
    from core import get_connection, fetch_informe, load_dotenv_files
    from vault import read_vault, decrypt_payload, dsn_from_database_section

    load_dotenv_files()

    dsn = None
    vault_path = PROYECTO.parent / "credentials.vault.enc"
    local_conf_path = PROYECTO.parent / "local.conf"
    if local_conf_path.is_file():
        # Si existe local.conf, se usa directamente y no se pregunta nada.
        dsn = None
    elif vault_path.is_file():
        try:
            password = input("Contraseña del vault (Enter para omitir): ").strip()
            if password:
                blob = read_vault(vault_path)
                creds = decrypt_payload(blob, password)
                db = creds.get("database") or creds
                dsn = dsn_from_database_section(db)
        except Exception:
            print("  ⚠ Vault falló, usando local.conf")
            dsn = None

    conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
    print("✓ Conectado a la base de datos")
    auction_uuid = _resolver_identificador(conn, auction_uuid)

    data = fetch_informe(conn, auction_uuid)
    print(f"✓ Subasta: {data['subasta'].get('code')} — {data['subasta'].get('status')}")

    datos = _obtener_datos_acta(conn, auction_uuid, data)
    conn.close()
    return data, datos


# ── Queries ───────────────────────────────────────────────────────────────────

def _obtener_datos_acta(conn, auction_uuid: str, data: dict) -> dict:
    subasta = data["subasta"]

    # 1. Inmueble
    inmueble = {}
    grupo_id = None
    nombre_grupo = ""
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT psv.contact_tercero_id, psv.inmueble_id, psv.grupo_id,
                   ct.direccion_principal
            FROM polibid_subastas_v2 psv
            LEFT JOIN contact_terceros ct ON ct.id = psv.contact_tercero_id
            WHERE psv.auction_id = %s::uuid
            ORDER BY psv.id DESC LIMIT 1
            """,
            (auction_uuid,),
        )
        link = cur.fetchone()
        direccion = "—"
        if link:
            direccion = link[3] or "—"
            if link[1]:
                # Subasta de un inmueble individual (no agrupado).
                cur.execute(
                    """
                    SELECT codigo, numero_matricula, referencia, grupo_id, nombre_grupo, referencia
                    FROM mst_inmuebles WHERE id = %s
                    """,
                    (link[1],),
                )
                row = cur.fetchone()
                if row:
                    grupo_id    = row[3]
                    nombre_grupo = row[4] or row[2] or ""
                    inmueble = {
                        "codigo":      row[0] or "—",
                        "fmi":         row[1] or "—",
                        "nombre_grupo": row[4] or row[2] or "—",
                    }
            elif link[2]:
                # Subasta de una unidad inmobiliaria agrupada: psv.inmueble_id
                # es NULL en estos casos (es normal, no un error) y el vínculo
                # correcto es psv.grupo_id. Antes de este fix, al no revisar
                # este caso, el código caía al fallback de manifestacion_interes
                # (otra tabla, sin relación garantizada con lo realmente
                # subastado) y podía traer los datos de OTRA propiedad.
                cur.execute(
                    """
                    SELECT id, codigo, numero_matricula, grupo_id, nombre_grupo,
                           referencia, codigo_grupo
                    FROM mst_inmuebles WHERE grupo_id = %s ORDER BY es_padre DESC, id
                    """,
                    (link[2],),
                )
                rows_g = cur.fetchall()
                if rows_g:
                    r = rows_g[0]
                    grupo_id = r[3]
                    nombre_grupo = r[4] or r[5] or ""
                    codigo_inm = r[6] or r[1] or "—"
                    fmi_inm = ", ".join(x[2] for x in rows_g if x[2])
                    inmueble = {
                        "codigo": codigo_inm,
                        "fmi": fmi_inm,
                        "nombre_grupo": nombre_grupo,
                    }
        inmueble["direccion"] = inmueble.get("nombre_grupo", direccion)

        # Fallback: buscar por manifestacion_interes si no se encontro inmueble
        if not inmueble.get("fmi") or inmueble.get("fmi") == "—":
            with conn.cursor() as cur2:
                cur2.execute("""
                    SELECT DISTINCT mani.grupo_id, mani.inmueble_id
                    FROM polybid.auction_participants p
                    JOIN polibid_credentials pc ON pc.client_id = p.client_id
                    JOIN manifestacion_interes mani ON mani.contact_tercero_id = pc.contact_tercero_id
                    WHERE p.auction_id = %s::uuid
                    AND (mani.grupo_id IS NOT NULL OR mani.inmueble_id IS NOT NULL)
                    ORDER BY mani.grupo_id DESC NULLS LAST
                    LIMIT 1
                """, (auction_uuid,))
                mani = cur2.fetchone()
                if mani:
                    if mani[0]:
                        cur2.execute("""
                            SELECT id, codigo, numero_matricula, grupo_id, nombre_grupo,
                                   referencia, codigo_grupo
                            FROM mst_inmuebles WHERE grupo_id = %s ORDER BY es_padre DESC, id
                        """, (mani[0],))
                        rows_g = cur2.fetchall()
                        if rows_g:
                            r = rows_g[0]
                            inm_id = r[0]; grupo_id = r[3]
                            nombre_grupo = r[4] or r[5] or ""
                            codigo_inm = r[6] or r[1] or "—"
                            fmi_inm = ", ".join(x[2] for x in rows_g if x[2])
                            inmueble = {"codigo": codigo_inm, "fmi": fmi_inm,
                                        "nombre_grupo": nombre_grupo, "direccion": nombre_grupo}
                    elif mani[1]:
                        cur2.execute("""
                            SELECT id, codigo, numero_matricula, grupo_id, nombre_grupo, referencia
                            FROM mst_inmuebles WHERE id = %s
                        """, (mani[1],))
                        r = cur2.fetchone()
                        if r:
                            inm_id = r[0]; grupo_id = r[3]
                            nombre_grupo = r[4] or r[5] or ""
                            inmueble = {"codigo": r[1] or "—", "fmi": r[2] or "—",
                                        "nombre_grupo": nombre_grupo, "direccion": nombre_grupo}

    # 1.b Unidad inmobiliaria: si el inmueble pertenece a un grupo (varios
    # FMIs agrupados bajo un grupo_id), el campo ##fmi## del acta debe
    # mostrar el CÓDIGO DE LA UNIDAD (formato "UNI-XXXX-AAAA"), no la lista
    # de folios de matrícula individuales — así lo pidió el usuario.
    if grupo_id:
        codigo_unidad = _codigo_unidad_de_grupo(conn, grupo_id)
        if codigo_unidad:
            inmueble["fmi"] = codigo_unidad

    # 2. Fechas via scraping
    print(f"\nObteniendo fechas desde la página...")
    fechas = _scrape_fechas(grupo_id, nombre_grupo, inm_id=link[1] if link else None)
    fecha_apertura = fechas.get("apertura", "—")
    fecha_cierre = fechas.get("cierre", "—")

    # La fecha de publicación se toma PRIMERO del cronograma scrapeado de la
    # página (fecha real de inicio del proceso). Solo si la página no trae
    # ningún dato de cronograma (pasa con subastas ya finalizadas cuya URL
    # el sitio reutilizó para un ciclo de venta nuevo de la misma propiedad)
    # se usa como respaldo el created_at de la subasta en la base de datos
    # — que no es exactamente lo mismo, pero es mejor que dejarlo en blanco.
    fecha_publicacion = fechas.get("publicacion", "—")
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

    # 3. Participantes
    participantes = []
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT
                ct.nombre_principal,
                ct.identificacion_numero,
                ct.identificacion_tipo,
                ct.lugar_expedicion_doc,
                ct.ciudad,
                (SELECT b.amount
                 FROM polybid.auction_bids b
                 WHERE b.auction_id = p.auction_id
                   AND b.client_id = p.client_id
                 ORDER BY b.created_at DESC
                 LIMIT 1) AS ultima_puja
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid
            ORDER BY p.created_at
            """,
            (auction_uuid,),
        )
        for row in cur.fetchall():
            nombre, cedula, tipo_id, lugar_exp, ciudad, monto = row
            # Solo incluir si tiene al menos una puja
            if not monto:
                continue
            ciudad_cedula = ciudad if tipo_id == "NIT" else (lugar_exp or "—")
            participantes.append({
                "nombre":        (nombre or "—").upper(),
                "cedula":        str(cedula or "—"),
                "ciudad_cedula": ciudad_cedula or "—",
                "monto":         _fmt_numero(monto),
            })

    # 4. Ganador
    ganador = data.get("ganador") or {}
    ganador_nombre = (ganador.get("nombre_principal") or ganador.get("display_name") or "—").upper()
    cedula_ganador = "—"
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
            cedula_ganador = str(row[0] or "—")

    return {
        "codigo_subasta":    subasta.get("code", "—"),
        "fecha_publicacion": fecha_publicacion,
        "fecha_inicio":      fecha_apertura if fecha_apertura != "—" else _fmt_fecha(subasta.get("start_date")),
        "fecha_fin":         fecha_cierre if fecha_cierre != "—" else _fmt_fecha(subasta.get("end_date")),
        "fmi":               inmueble.get("fmi", "—"),
        "direcion":          fechas.get("direccion", inmueble.get("direccion", "—")),
        "direccion":         fechas.get("direccion", inmueble.get("direccion", "—")),
        "codigo_inmueble":   inmueble.get("codigo", "—"),
        "participantes":     participantes,
        "nombre_ganador":    ganador_nombre,
        "cedula_ganador":    cedula_ganador,
        "monto_ganador":     _fmt_numero(ganador.get("amount", 0)),
    }


# ── Formato ───────────────────────────────────────────────────────────────────

def _fmt_fecha(fecha) -> str:
    if not fecha:
        return "—"
    return f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"


def _fmt_numero(valor) -> str:
    if not valor:
        return "0"
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return str(valor)


# ── Generación docx ───────────────────────────────────────────────────────────

def generar_docx(datos: dict, ruta_salida: Path) -> None:
    import re as _re

    with zipfile.ZipFile(PLANTILLA, "r") as zin:
        archivos = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = archivos["word/document.xml"].decode("utf-8")

    # 1. Reemplazos simples
    reemplazos = {
        "##codigo_subasta##":    datos["codigo_subasta"],
        "##fecha_publicacion##": datos["fecha_publicacion"],
        "##fecha_inicio##":      datos["fecha_inicio"],
        "##fecha_fin##":         datos["fecha_fin"],
        "##fmi##":               datos["fmi"],
        "##direcion##":          datos["direcion"],
        "##direccion##":         datos["direccion"],
        "##codigo_inmueble##":   datos["codigo_inmueble"],
        "##nombre_ganador##":    datos["nombre_ganador"],
        "##cedula_ganador##":    datos["cedula_ganador"],
        "##monto_ganador##":     datos["monto_ganador"],
    }
    # Colapsar placeholders fragmentados en multiples runs
    import re as _re_col
    # Patron: ##fecha_X## partido en runs
    doc_xml = _re_col.sub(
        r'<w:t>##fecha_</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>([^<]+)</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>##</w:t></w:r>',
        r'<w:t>##fecha_\1##</w:t></w:r>',
        doc_xml, flags=_re_col.DOTALL
    )
    # Patron: ## + nombre + ## partido en runs
    doc_xml = _re_col.sub(
        r'<w:t>##</w:t></w:r><w:r[^>]*><w:t>([^<#]+)</w:t></w:r><w:r[^>]*><w:t>##</w:t></w:r>',
        r'<w:t>##\1##</w:t></w:r>',
        doc_xml, flags=_re_col.DOTALL
    )
    # Patron: ##nombre# + <w:proofErr/> + # (con posible espacio) partido
    doc_xml = _re_col.sub(
        r'<w:t>##([^<#]+)#</w:t></w:r><w:proofErr[^/]*/><w:r[^>]*><w:t[^>]*>#\s*</w:t></w:r>',
        r'<w:t>##\1##</w:t></w:r>',
        doc_xml, flags=_re_col.DOTALL
    )
    # Reparar tags mal cerrados
    doc_xml = _re_col.sub(r'<w:t>([^<]+)</w:r>', r'<w:t>\1</w:t></w:r>', doc_xml)

    for placeholder, valor in reemplazos.items():
        doc_xml = doc_xml.replace(placeholder, valor)

    # 2. Bloque de participantes — duplicar el parrafo plantilla por cada participante
    PARRAFO_PLANTILLA = (
        '<w:p w14:paraId="28701AD4" w14:textId="1843D5B8" w:rsidR="00A46A55" '
        'w:rsidRDefault="007F604A" w:rsidP="001141D7"><w:pPr><w:ind w:left="360" '
        'w:right="704"/></w:pPr><w:r><w:t>##nombre##</w:t></w:r>'
        '<w:r w:rsidR="00511C95"><w:t xml:space="preserve"> </w:t></w:r>'
        '<w:r w:rsidR="005F4ECB" w:rsidRPr="005F4ECB"><w:rPr><w:lang w:val="es-ES"/>'
        '</w:rPr><w:t xml:space="preserve">– C.C: </w:t></w:r>'
        '<w:r><w:t xml:space="preserve">##cedula##, </w:t></w:r>'
        '<w:r><w:rPr><w:lang w:val="es-ES"/></w:rPr>'
        '<w:t xml:space="preserve">##ciudad_cedula## </w:t></w:r>'
        '<w:r w:rsidR="005F4ECB" w:rsidRPr="005F4ECB"><w:rPr><w:lang w:val="es-ES"/>'
        '</w:rPr><w:t xml:space="preserve">- </w:t></w:r>'
        '<w:r w:rsidR="00511C95" w:rsidRPr="00511C95">'
        '<w:t xml:space="preserve">$ </w:t></w:r>'
        '<w:r><w:t>##monto##</w:t></w:r></w:p>'
    )

    if datos["participantes"]:
        parrafos_participantes = ""
        for p in datos["participantes"]:
            parrafo = PARRAFO_PLANTILLA
            parrafo = parrafo.replace("##nombre##", p["nombre"])
            parrafo = parrafo.replace("##cedula##", p["cedula"])
            parrafo = parrafo.replace("##ciudad_cedula##", p["ciudad_cedula"])
            parrafo = parrafo.replace("##monto##", p["monto"])
            parrafos_participantes += parrafo
        doc_xml = doc_xml.replace(PARRAFO_PLANTILLA, parrafos_participantes)
    else:
        doc_xml = doc_xml.replace(PARRAFO_PLANTILLA, "")

    archivos["word/document.xml"] = doc_xml.encode("utf-8")

    with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Uso: python generar_acta.py <auction_uuid | codigo_subasta | FMI | codigo_unidad>")
        sys.exit(1)

    if not PLANTILLA.exists():
        print(f"✗ No se encuentra la plantilla: {PLANTILLA.name}")
        sys.exit(1)

    identificador = sys.argv[1].strip()

    print(f"\n{'='*55}")
    print(f"  GENERADOR DE ACTA DE CERTIFICACIÓN")
    print(f"{'='*55}")
    print(f"  Identificador: {identificador}\n")

    data, datos = conectar(identificador)

    SALIDA.mkdir(exist_ok=True)
    codigo = datos["codigo_subasta"]
    codigo_corto = codigo.replace("ACTIBID-", "")[:20]
    nombre_archivo = f"ACTA_{codigo_corto}.docx"
    ruta = SALIDA / nombre_archivo

    try:
        generar_docx(datos, ruta)
        print(f"\n✓ Acta generada: ./actas/{nombre_archivo}")
        print(f"\n  Subasta:     {datos['codigo_subasta']}")
        print(f"  Publicación: {datos['fecha_publicacion']}")
        print(f"  Apertura:    {datos['fecha_inicio']}")
        print(f"  Cierre:      {datos['fecha_fin']}")
        print(f"  FMI:         {datos['fmi']}")
        print(f"  Dirección:   {datos['direccion']}")
        print(f"  Ganador:     {datos['nombre_ganador']}")
        print(f"  Cédula:      {datos['cedula_ganador']}")
        print(f"  Monto:       $ {datos['monto_ganador']}")
        print(f"  Participantes:")
        for p in datos["participantes"]:
            print(f"    - {p['nombre']} | {p['cedula']} | $ {p['monto']}")
    except Exception as e:
        print(f"✗ ERROR: {e}")


if __name__ == "__main__":
    main()
