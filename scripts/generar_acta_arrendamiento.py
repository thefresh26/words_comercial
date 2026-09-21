"""
Genera Actas de Comité de Arrendamiento (.docx) a partir de la plantilla
templates/ACTA_ARRIENDOS.docx, sacando los datos AUTOMÁTICAMENTE (sin que haya
que leerlos ni transcribirlos a mano) de los 2 soportes que siempre traen
esa información en texto (no como imagen escaneada):

  - Estimado de Renta (PDF)  -> FMI, dirección, ciudad, dirección
    territorial, tipo de bien (inferido de la descripción).
  - Aprobado de garantía/aseguradora (PDF) -> nombre e identificación del
    arrendatario (y codeudor, si tiene).

Los 2 soportes que SÍ son imagen/captura (SAGRILAFT y, si se adjunta, la
cédula) no se "leen" -- se insertan directamente como fotos en el Word,
que es lo único que se necesita de ellos.

Campos que NUNCA vienen en ningún soporte y quedan tal cual "(HUMANO)" en
la plantilla (así lo pidió el usuario, no se tocan):
  - Fecha de inicio, fecha final, duración
  - Valor canon mensual, valor canon total

Para agregar un caso nuevo: agrega una entrada a la lista CASOS más abajo
con las rutas de sus soportes (no hace falta que un humano/LLM transcriba
ningún dato -- el script los saca solo). Al final imprime, por cada acta,
qué quedó pendiente de revisar a mano (incluye siempre la ciudad de
expedición de la cédula, porque esa sí requiere mirar la cédula física).
"""
from __future__ import annotations

import os
import re
import sys
import unicodedata
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*getdata.*")

import docx
import pdfplumber
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

PROYECTO = Path(__file__).resolve().parent.parent
PLANTILLA = PROYECTO / "templates" / "ACTA_ARRIENDOS.docx"
SALIDA = PROYECTO / "output" / "actas_arrendamiento"
SCRATCH = PROYECTO / "scratch_arrendamiento"

FMI_PLANTILLA_VIEJO = "060-243312"  # hardcodeado en la plantilla, se reemplaza siempre.

# Regla de negocio dada por el usuario:
#   COMERCIAL = local, bodega, lote (incl. "lote de construcción"), y
#               apartamento SOLO SI es de uso turístico.
#   VIVIENDA  = casa, y apartamento (por defecto, si no es turístico).
# "apartamento" se resuelve aparte (ver _clasificar_tipo_bien) porque su
# categoría depende de si el texto menciona uso turístico o no.
TIPOS_BIEN = {
    "local": ("LOCAL COMERCIAL", "COMERCIAL"),
    "bodega": ("BODEGA", "COMERCIAL"),
    "lote de construcción": ("LOTE DE CONSTRUCCIÓN", "COMERCIAL"),
    "lote": ("LOTE", "COMERCIAL"),
    "parqueadero": ("PARQUEADERO", "COMERCIAL"),  # regla confirmada por el usuario
    "casa": ("CASA", "VIVIENDA"),
    "apartamento": ("APARTAMENTO", "VIVIENDA"),  # categoría se ajusta abajo si es turístico
}

PALABRAS_TURISTICO = ["turístic", "turistic"]  # cubre "turístico/a" y sin tilde


def _clasificar_tipo_bien(texto_lower: str) -> tuple[str, str]:
    """Devuelve (etiqueta, categoria_contrato) según la regla de negocio."""
    for candidato, (etiqueta, categoria) in TIPOS_BIEN.items():
        if candidato in texto_lower:
            if candidato == "apartamento" and any(p in texto_lower for p in PALABRAS_TURISTICO):
                return "APARTAMENTO TURÍSTICO", "COMERCIAL"
            return etiqueta, categoria
    return "—", "—"


# ── Extracción automática desde los PDFs de soporte ──────────────────────

def _ocr_pdf(ruta: Path) -> str:
    """Respaldo para PDFs escaneados (sin capa de texto real): convierte
    cada página a imagen con PyMuPDF y le pasa OCR con Tesseract. Se usa
    solo cuando pdfplumber no extrae nada, porque el OCR es más lento y
    menos preciso que leer el texto real cuando este existe."""
    try:
        import pymupdf
        import pytesseract
        from PIL import Image
        import io

        partes = []
        with pymupdf.open(str(ruta)) as pdf:
            for pagina in pdf:
                pix = pagina.get_pixmap(dpi=250)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                partes.append(pytesseract.image_to_string(img, lang="spa"))
        return "\n".join(partes)
    except Exception:
        return ""


def _texto_pdf(ruta: Path) -> str:
    # x_tolerance=1: el "Estimado de Renta" viene con texto justificado y,
    # con la tolerancia por defecto de pdfplumber (3), las palabras quedan
    # pegadas sin espacio (ej. "Localmedianeroubicado..."). Con tolerancia 1
    # sí separa bien las palabras.
    with pdfplumber.open(str(ruta)) as pdf:
        texto = "\n".join(p.extract_text(x_tolerance=1) or "" for p in pdf.pages)
    if len(texto.strip()) < 20:
        # Probablemente un PDF escaneado (imagen) sin capa de texto -- se
        # intenta OCR antes de devolver vacío, para no perder datos que sí
        # están en el documento, solo que como imagen.
        texto_ocr = _ocr_pdf(ruta)
        if len(texto_ocr.strip()) > len(texto.strip()):
            return texto_ocr
    return texto


def extraer_descripcion_estimado_renta(ruta_pdf: Path) -> str | None:
    """Saca el párrafo completo de la fila 'DESCRIPCIÓN' de la tabla del
    Estimado de Renta (incluye las Notas), con los espacios corregidos.
    Se hace recortando la celda por sus coordenadas reales (no por texto
    plano), porque en texto plano la etiqueta 'DESCRIPCIÓN' queda intercalada
    en medio del párrafo (por cómo pdfplumber ordena el texto en Y) y no
    sirve como separador confiable."""
    with pdfplumber.open(str(ruta_pdf)) as pdf:
        page = pdf.pages[0]
        tablas = page.find_tables()
        if not tablas:
            return None
        for fila in tablas[0].rows:
            celdas = [c for c in fila.cells if c]
            if len(celdas) < 2:
                continue
            etiqueta = (page.crop(celdas[0]).extract_text(x_tolerance=1) or "").strip().upper()
            if etiqueta == "DESCRIPCIÓN" or etiqueta == "DESCRIPCION":
                contenido = page.crop(celdas[1]).extract_text(x_tolerance=1) or ""
                texto = " ".join(l.strip() for l in contenido.split("\n") if l.strip())
                return _limpiar_descripcion_cadastral(texto)
    return None


def _limpiar_descripcion_cadastral(texto: str) -> str:
    """Algunos Estimados de Renta traen, pegado al final del mismo párrafo de
    descripción, un bloque de nomenclatura catastral (p. ej. 'Nomenclatura
    Nro 000000000...') que en el PDF está escrito con cada letra separada
    por espacios -- pdfplumber lo extrae como 'N o m e n c l a ...' y queda
    mezclado con la descripción real. Se detecta ese patrón (varias letras
    sueltas seguidas, separadas por espacios) y se corta el texto ahí, para
    no incluir ese bloque ilegible en el Acta."""
    m = re.search(r"(?:\b[A-Za-zÀ-ÿ]\s){5,}[A-Za-zÀ-ÿ]\b", texto)
    if m and m.start() > 50:
        return texto[: m.start()].strip()
    return texto


# Umbral de área (en puntos PDF²) para distinguir una foto real del
# inmueble de un logo o una firma escaneada (que también quedan como
# "imagen embebida" en el PDF, pero son mucho más chicos). 4,500 pt² deja
# fuera los logos (79x30 ~ 2,400 pt²) y firmas (96x17 ~ 1,600 pt², 70x40 ~
# 2,800 pt²), pero SÍ deja pasar las fotos en miniatura que trae el formato
# nuevo de reporte automatizado (~72px de alto, áreas de 5,000-13,000 pt²),
# además de las fotos/collage grandes del formato viejo (~67,800 pt²).
AREA_MINIMA_FOTO = 4500

# Si una "imagen" ocupa más de este porcentaje del área de la página, no es
# una foto del inmueble -- es la PÁGINA ENTERA escaneada como una sola
# imagen (pasa con Estimados de Renta que llegan como escaneo, no como PDF
# de texto). Insertar eso como "foto del inmueble" mostraría el documento
# completo (con tablas, firmas, etc.) en vez de una foto real, así que se
# descarta y se deja como pendiente en vez de insertar algo engañoso.
PROPORCION_MAXIMA_PAGINA_COMPLETA = 0.5


# Umbral de proporción de píxeles "casi blancos" (fondo de tabla/página) en
# una imagen candidata para considerarla "texto/tabla" y no una foto real.
# Medido contra casos reales: las fotos de inmuebles (interiores, fachadas,
# etc.) dan como mucho ~0.14; las tablas de avalúo/renta y los párrafos de
# texto capturados como imagen (pasa en algunos formatos de reporte, p. ej.
# fincas rurales) dan 0.45 o más -- se deja un margen amplio en 0.30. No
# depende de ningún programa externo (solo PIL, que ya se usa en todo el
# script), a diferencia de un OCR que requeriría instalar Tesseract aparte
# en cada computador donde se corra esto.
PROPORCION_MAXIMA_PIXELES_BLANCOS = 0.30


def _es_imagen_de_texto_o_tabla(ruta_imagen: Path) -> bool:
    """True si la imagen recortada parece ser una tabla o un bloque de texto
    (capturado como imagen en el PDF) en vez de una foto real del inmueble."""
    from PIL import Image

    with Image.open(ruta_imagen) as img:
        # Se reduce a una miniatura antes de contar píxeles -- de sobra para
        # medir la proporción de blanco, y evita recorrer imágenes grandes
        # píxel por píxel (más rápido, sin perder precisión para esto).
        muestra = img.convert("RGB").resize((150, 150))
        total = muestra.width * muestra.height
        pixeles = list(muestra.getdata())
    casi_blancos = sum(1 for r, g, b in pixeles if r > 235 and g > 235 and b > 235)
    return (casi_blancos / total) >= PROPORCION_MAXIMA_PIXELES_BLANCOS


def extraer_fotos_inmueble(ruta_pdf: Path, nombre_salida: str) -> list[Path]:
    """Recorta del Estimado de Renta TODAS las imágenes que parezcan fotos
    reales del inmueble (en todas las páginas, no solo la 1), filtrando por
    tamaño para descartar logos y firmas escaneadas, y descartando también
    cualquier imagen que sea básicamente la página completa escaneada (eso
    no es "una foto del inmueble", es el documento entero). Devuelve la
    lista de rutas (puede tener 0, 1 o más elementos), ordenadas de más
    grande a más chica."""
    fotos = []
    with pdfplumber.open(str(ruta_pdf)) as pdf:
        candidatas = []
        for num_pagina, page in enumerate(pdf.pages):
            area_pagina = page.width * page.height
            for im in page.images:
                area = (im["x1"] - im["x0"]) * (im["bottom"] - im["top"])
                if area < AREA_MINIMA_FOTO:
                    continue
                if area >= PROPORCION_MAXIMA_PAGINA_COMPLETA * area_pagina:
                    continue  # es la página completa escaneada, no una foto
                candidatas.append((area, num_pagina, im))
        candidatas.sort(key=lambda c: c[0], reverse=True)

        SCRATCH.mkdir(parents=True, exist_ok=True)
        for i, (area, num_pagina, im) in enumerate(candidatas):
            page = pdf.pages[num_pagina]
            bbox = (
                max(im["x0"], 0), max(im["top"], 0),
                min(im["x1"], page.width), min(im["bottom"], page.height),
            )
            ruta_salida = SCRATCH / f"{nombre_salida}-foto{i + 1}.png"
            page.crop(bbox).to_image(resolution=200).save(str(ruta_salida))
            if _es_imagen_de_texto_o_tabla(ruta_salida):
                # No es una foto real -- es una tabla (avalúo, renta, etc.)
                # o un párrafo de texto capturado como imagen; se descarta.
                ruta_salida.unlink(missing_ok=True)
                continue
            fotos.append(ruta_salida)
    return fotos


def extraer_de_estimado_renta(ruta_pdf: Path, fmi_conocido: str | None = None, territorial_conocida: str | None = None) -> dict:
    """fmi_conocido: FMI que ya se conoce de forma independiente (p. ej. por
    el nombre de la carpeta del caso) -- se usa como respaldo cuando el PDF
    no trae la matrícula en ningún formato reconocido (como el "Dictamen
    Comercial y Financiero", que no incluye ese dato).

    territorial_conocida: territorial que ya se conoce por la carpeta donde
    está el caso (ej. ".../02_TERRITORIAL OCCIDENTE/...") -- mismo tipo de
    respaldo que fmi_conocido, para cuando el documento tampoco trae ese
    dato (pasa con el mismo "Dictamen Comercial y Financiero")."""
    texto = _texto_pdf(ruta_pdf)

    def _buscar(patron, default="—"):
        m = re.search(patron, texto, re.IGNORECASE)
        return m.group(1).strip() if m else default

    # Se prueba primero la variante narrativa "Folio de Matrícula
    # Inmobiliaria No. 140-16649 corresponde a..." porque es más estricta
    # (corta justo en el número/código, sin arrastrar el resto de la
    # oración); si no aparece, se prueba el formato viejo de tabla
    # "MATRÍCULA INMOBILIARIA <valor>" que sí puede tomar toda la línea.
    fmi = _buscar(r"Folio\s+de\s+Matr[íi]cula\s+Inmobiliaria\s+No\.?\s*([0-9A-Za-z\-]+)")
    if fmi == "—":
        fmi = _buscar(r"MATR[ÍI]CULA INMOBILIARIA\s+([^\n]+)")
    if fmi == "—" and fmi_conocido:
        fmi = fmi_conocido
    # (?!General\b) evita que esto agarre el pie de página corporativo que
    # trae TODO Estimado de Renta/Dictamen ("Direccion General: Carrera 7 #
    # 32-42 Centro Comercial San Martin Local 107 / PBX: ..."), que no es la
    # dirección del inmueble sino la de las oficinas de la SAE -- si no se
    # excluye, en los documentos que no traen una "DIRECCIÓN" real (formato
    # "Dictamen Comercial y Financiero", que usa "Ubicacion:" en su lugar)
    # esto se colaba como si fuera la dirección del inmueble.
    direccion = _buscar(r"DIRECCI[ÓO]N\s+(?!General\b)([^\n]+)")
    m_ciudad = re.search(r"CIUDAD\s+([A-Za-zÀ-ÿ]+)\s+\d+\s+BARRIO\s+([^\n]+)", texto, re.IGNORECASE)
    ciudad = m_ciudad.group(1).strip() if m_ciudad else "—"
    barrio = m_ciudad.group(2).strip() if m_ciudad else "—"
    if direccion == "—":
        # Formato nuevo del "Dictamen Comercial y Financiero": no trae
        # "DIRECCIÓN"/"CIUDAD"/"BARRIO" por separado, sino una sola línea
        # "Ubicacion: Barrio X, Sector Y, Ciudad".
        m_ubic = re.search(r"Ubicacion:\s*([^\n]+)", texto, re.IGNORECASE)
        if m_ubic:
            direccion = m_ubic.group(1).strip()
            partes_ubic = [p.strip() for p in direccion.split(",")]
            ciudad = partes_ubic[-1] if partes_ubic else "—"
    territorial = _buscar(r"DIRECCION TERRITORIAL\s+(\w+)\s+FECHA")
    if territorial == "—":
        # Variante real encontrada en varios Estimados de Renta: no dice
        # "DIRECCION TERRITORIAL", sino "GERENCIA / REGIONAL (...) TERRITORIAL <ZONA>".
        m_terr2 = re.search(
            r"TERRITORIAL\s+(CARIBE|OCCIDENTE|SUR|CENTRO ORIENTE|NORTE|CENTRO)",
            texto, re.IGNORECASE,
        )
        territorial = m_terr2.group(1) if m_terr2 else "—"
    if territorial == "—" and territorial_conocida:
        territorial = territorial_conocida

    # Tipo de bien: primero se busca una etiqueta explícita -- "Tipo de
    # Inmueble:" o simplemente "Inmueble:" (formatos nuevos automatizados
    # tipo "Dictamen Comercial y Financiero" / "Informe Técnico"); si no
    # aparece ninguna, se infiere buscando palabras clave en todo el texto
    # (el formato viejo del Estimado de Renta, donde la descripción suele
    # venir sin espacios por el layout, ej. "Localmedianeroubicado...").
    # Buscar primero la etiqueta explícita evita falsos positivos como el
    # pie de página "...Centro Comercial San Martin Local 107", que de otro
    # modo clasificaría como LOCAL COMERCIAL cualquier inmueble.
    m_tipo_explicito = re.search(r"(?:Tipo de )?Inmueble:\s*([A-Za-zÀ-ÿ ]+)", texto, re.IGNORECASE)
    if m_tipo_explicito:
        tipo_bien, categoria_contrato = _clasificar_tipo_bien(m_tipo_explicito.group(1).strip().lower())
    else:
        tipo_bien, categoria_contrato = "—", "—"
    if tipo_bien == "—":
        tipo_bien, categoria_contrato = _clasificar_tipo_bien(texto.lower())

    # El FMI de la matrícula a veces trae la unidad/local pegado (ej. "140-
    # 103759 UE 2 (Local 2)"); se separa la matrícula "limpia" (para nombre
    # de archivo y referencias cortas) del detalle completo (para la tabla).
    m_fmi_limpio = re.match(r"([\d]{2,3}[A-Za-z]?-\d+)", fmi)
    fmi_limpio = m_fmi_limpio.group(1) if m_fmi_limpio else fmi

    return {
        "fmi": fmi,
        "fmi_limpio": fmi_limpio,
        "direccion": f"{direccion}, Barrio {barrio}, {ciudad}" if barrio != "—" else direccion,
        "ciudad": ciudad,
        "direccion_territorial": territorial.upper() if territorial != "—" else territorial,
        "tipo_bien": tipo_bien,
        "categoria_contrato": categoria_contrato,
    }


def extraer_de_aprobado(ruta_pdf: Path) -> dict:
    """Lee el documento de "aprobado" (póliza/garantía). El formato del
    identificador varía entre casos: personas naturales traen "CC:" y
    empresas traen "NIT" (a veces con dígito de verificación separado por
    guion, ej. "901851608-7"), así que se aceptan ambos."""
    texto = _texto_pdf(ruta_pdf)
    m = re.search(
        r"DATOS DEL ARRENDATARIO\s*\n([A-ZÑÁÉÍÓÚ0-9&.,\- ]+?)\s*\n"
        r"(NIT|CC|C\.C\.?)\.?:?\s*([\d.\-]+)",
        texto, re.IGNORECASE,
    )
    nombre = m.group(1).strip() if m else "—"
    cedula = m.group(3).strip() if m else "—"
    id_siglas = _normalizar_siglas_id(m.group(2)) if m else _normalizar_siglas_id(None, cedula)
    m_codeudor = re.search(
        r"CODEUDOR SOLIDARIO\s*\n([A-ZÑÁÉÍÓÚ0-9&.,\- ]+?)\s*\n"
        r"(?:NIT|CC|C\.C\.?)\.?:?\s*([\d.\-]+)",
        texto, re.IGNORECASE,
    )
    codeudor_nombre = m_codeudor.group(1).strip() if m_codeudor else None
    codeudor_cedula = m_codeudor.group(2).strip() if m_codeudor else None
    return {
        "arrendatario_nombre": nombre,
        "id_numero": cedula,
        "id_siglas": id_siglas,
        "codeudor_nombre": codeudor_nombre,
        "codeudor_cedula": codeudor_cedula,
    }


def _normalizar_siglas_id(etiqueta: str | None, numero: str = "") -> str:
    """Decide si el identificador del arrendatario es "NIT" o "C.C.". Prioriza
    la etiqueta encontrada en el documento (NIT/CC/C.C.); si no hay etiqueta,
    usa como respaldo el formato del número: un NIT colombiano casi siempre
    trae el dígito de verificación separado por un guion al final (ej.
    "900.456.885-3"), mientras que una cédula de persona natural no."""
    if etiqueta and etiqueta.strip().upper().replace(".", "").startswith("NIT"):
        return "NIT"
    if etiqueta:
        return "C.C."
    if numero and re.search(r"-\d\s*$", numero.strip()):
        return "NIT"
    return "C.C."


def extraer_de_solicitud_arrendamiento(ruta_pdf: Path) -> dict | None:
    """Respaldo para cuando la carpeta NO tiene el documento de Aprobado/
    Póliza (p. ej. Cindy Murillo, FMI 001-389120, que solo tiene la
    "SOLICITUD DE ARRENDAMIENTO" dentro de la subcarpeta "ARRENDATARIO ...").
    Ese formato trae el nombre, número de documento y ciudad de expedición
    en frases tipo "Yo <NOMBRE> identificado(a) con el documento de
    identidad C.C. X , C.E. , NIT No. <NUMERO> expedido en la ciudad de
    <CIUDAD>", repetidas varias veces en el PDF (autorización de centrales
    de riesgo, declaración de origen de fondos, etc.) -- se toma la primera
    coincidencia. Devuelve None si el documento no trae ese formato (para
    no insertar datos incorrectos a la fuerza)."""
    texto = _texto_pdf(ruta_pdf)
    m = re.search(
        r"Yo\s+([A-ZÑÁÉÍÓÚ\s]+?)\s+identificad[oa]\s*\(?a?\)?\s*con\s+el\s+documento\s+de\s+identidad\s+"
        r"C\.C\.\s*(X)?\s*,\s*C\.E\.\s*(X)?\s*,?\s*NIT\s*No\.\s*([\d.\-]+)\s+expedido\s+en\s*"
        r"(?:la\s+ciudad\s+de\s*)?([A-ZÑÁÉÍÓÚ,.\s]+?),\s*actuando",
        texto, re.IGNORECASE,
    )
    if not m:
        return None
    nombre, marca_cc, marca_ce, numero, ciudad = m.groups()
    if marca_ce and not marca_cc:
        etiqueta = "C.E."
    elif marca_cc:
        etiqueta = "C.C."
    else:
        etiqueta = None
    return {
        "arrendatario_nombre": nombre.strip(),
        "id_numero": numero.strip(),
        "id_siglas": _normalizar_siglas_id(etiqueta, numero),
        "id_ciudad": re.sub(r"\s+", " ", ciudad).strip(" ,."),
        "codeudor_nombre": None,
        "codeudor_cedula": None,
    }


def extraer_de_carta_juramentada(ruta_pdf: Path) -> dict | None:
    """Segundo respaldo (después de extraer_de_solicitud_arrendamiento) para
    cuando la carpeta NO tiene el documento de Aprobado/Póliza NI la
    "SOLICITUD DE ARRENDAMIENTO" -- cubre varias variantes de la carta/
    formato de declaración juramentada vistas en distintas carpetas:
      - Eyda Tenorio (FMI 024-20085): "Yo, <NOMBRE>, identificado(a) con
        Cedula de ciudadanía No. <NUMERO> de <CIUDAD> actuando..."
      - María Camila Hoyos (FMI 290-225030): el "FORMATO DECLARACIÓN
        JURAMENTADA..." institucional, que trae "Yo, <NOMBRE> identificado
        (a) con documento de identidad número <NUMERO> de <CIUDAD>,
        actuando...", pero como es un formato de líneas para llenar a
        mano, el texto sale con guiones bajos intercalados dentro de cada
        palabra (ej. "M__a_ri_a_ C_a_m__ila_") -- por eso se le quitan los
        "_" antes de buscar el patrón.
    A diferencia del formato de la Solicitud de Arrendamiento, el nombre
    aquí NO viene en mayúsculas fijas. Devuelve None si el documento no
    trae ninguna de estas frases (para no insertar datos incorrectos a la
    fuerza) -- pasa, por ejemplo, cuando el PDF es un escaneo sin texto
    (imagen), que _texto_pdf() no puede leer sin OCR."""
    texto = _texto_pdf(ruta_pdf).replace("_", "")
    m = re.search(
        r"Yo,?\s+([A-Za-zÀ-ÿ\s]+?),?\s+identificad[oa]\s*\(?a?\)?\s*con\s+"
        r"(Cedula de ciudadan[íi]a\]?|documento de identidad n[uú]mero|C\.C\.?|CC|NIT)\.?\s*(?:No\.?)?\s*"
        r"([\d.\-]+)\s+de\s+([A-Za-zÀ-ÿ,\s]+?)\s*,?\s*actuando",
        texto, re.IGNORECASE,
    )
    if not m:
        return None
    nombre, tipo_doc, numero, ciudad = m.groups()
    etiqueta = "NIT" if "NIT" in tipo_doc.upper() else "C.C."
    return {
        "arrendatario_nombre": re.sub(r"\s+", " ", nombre).strip().upper(),
        "id_numero": numero.strip(),
        "id_siglas": _normalizar_siglas_id(etiqueta, numero),
        "id_ciudad": re.sub(r"\s+", " ", ciudad).strip(" ,."),
        "codeudor_nombre": None,
        "codeudor_cedula": None,
    }


# ── Búsqueda de archivos por palabra clave (carpetas reales de OneDrive) ──
# Los nombres de archivo NO son consistentes entre casos (p.ej. "ESTIMADO
# DE RENTA...", "ACTI RENTA...", "ER FMI...", con o sin espacios/typos), así
# que en vez de exigir un nombre exacto se busca por palabras clave.

def _buscar_archivo(carpeta: Path, *palabras_clave: str, recursivo: bool = False) -> Path | None:
    """Primer archivo dentro de `carpeta` cuyo nombre contenga (sin
    importar mayúsculas/acentos) alguna de las `palabras_clave`. Si
    `recursivo` es True también busca dentro de subcarpetas."""
    import unicodedata

    def _normalizar(s: str) -> str:
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper()
        return s

    claves = [_normalizar(p) for p in palabras_clave]
    iterador = carpeta.rglob("*") if recursivo else carpeta.iterdir()
    candidatos = sorted(f for f in iterador if f.is_file())
    for f in candidatos:
        nombre_norm = _normalizar(f.name)
        if any(clave in nombre_norm for clave in claves):
            return f
    return None


def _buscar_subcarpeta(carpeta: Path, *palabras_clave: str) -> Path | None:
    import unicodedata

    def _normalizar(s: str) -> str:
        return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper()

    claves = [_normalizar(p) for p in palabras_clave]
    for f in sorted(p for p in carpeta.iterdir() if p.is_dir()):
        nombre_norm = _normalizar(f.name)
        if any(clave in nombre_norm for clave in claves):
            return f
    return None


def _pdf_a_png(ruta_pdf: Path, nombre_salida: str) -> Path | None:
    """Convierte la primera página del PDF a PNG usando PyMuPDF (pip install
    pymupdf) -- NO requiere instalar ningún programa aparte (a diferencia de
    Poppler/pdftoppm), por eso se usa esta librería y no subprocess."""
    if not ruta_pdf or not ruta_pdf.is_file():
        return None
    import pymupdf  # pip install pymupdf

    SCRATCH.mkdir(parents=True, exist_ok=True)
    candidato = SCRATCH / f"{nombre_salida}-1.png"
    with pymupdf.open(str(ruta_pdf)) as pdf:
        pagina = pdf[0]
        pix = pagina.get_pixmap(dpi=150)
        pix.save(str(candidato))
    return candidato if candidato.is_file() else None


# ── Generación del .docx ─────────────────────────────────────────────────

def _pdf_a_png_todas_paginas(ruta_pdf: Path, nombre_salida: str) -> list[Path]:
    """Convierte TODAS las páginas del PDF a PNG (una imagen por página) --
    se usa para mostrar el documento completo del Estimado de Renta bajo la
    sección "ESTIMADO DE RENTA" de la plantilla (distinta de las fotos
    reales del inmueble, que van en "REGISTRO FOTOGRÁFICO DEL INMUEBLE")."""
    if not ruta_pdf or not ruta_pdf.is_file():
        return []
    import pymupdf  # pip install pymupdf

    SCRATCH.mkdir(parents=True, exist_ok=True)
    paginas = []
    with pymupdf.open(str(ruta_pdf)) as pdf:
        for i, pagina in enumerate(pdf):
            candidato = SCRATCH / f"{nombre_salida}-pag{i + 1}.png"
            pix = pagina.get_pixmap(dpi=150)
            pix.save(str(candidato))
            if candidato.is_file():
                paginas.append(candidato)
    return paginas


def _insertar_imagen_en_marcador(doc, marcador: str, ruta_imagen: Path | None, ancho_mm: float = 140) -> bool:
    if not ruta_imagen or not ruta_imagen.is_file():
        return False
    for p in doc.paragraphs:
        if marcador in p.text:
            for run in list(p.runs):
                run.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            try:
                run.add_picture(str(ruta_imagen), width=Mm(ancho_mm))
            except Exception:
                # Archivo de imagen corrupto/ilegible: se trata igual que si
                # no se hubiera encontrado, para que el llamador use el
                # texto SIN_IMAGEN en lugar de que se caiga todo el script.
                return False
            return True
    return False


def _insertar_imagenes_en_marcador(doc, marcador: str, rutas_imagenes: list[Path], ancho_mm: float = 140) -> bool:
    """Igual que _insertar_imagen_en_marcador pero para varias imágenes en el
    mismo párrafo (una debajo de otra, con un salto de línea entre cada
    una)."""
    rutas_imagenes = [r for r in rutas_imagenes if r and r.is_file()]
    if not rutas_imagenes:
        return False
    for p in doc.paragraphs:
        if marcador in p.text:
            for run in list(p.runs):
                run.text = ""
            primero = True
            alguna_insertada = False
            for ruta in rutas_imagenes:
                run = p.add_run()
                if not primero:
                    run.add_break()
                try:
                    run.add_picture(str(ruta), width=Mm(ancho_mm))
                except Exception:
                    # Imagen corrupta/ilegible: se omite y se sigue con las
                    # demás en vez de interrumpir todo el script.
                    continue
                primero = False
                alguna_insertada = True
            return alguna_insertada
    return False


def _quitar_bordes_tabla(tabla) -> None:
    """Quita todos los bordes de una tabla (para que la cuadrícula de fotos
    no se vea como una tabla de verdad)."""
    tbl = tabla._tbl
    tblPr = tbl.tblPr
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = docx.oxml.OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elemento = docx.oxml.OxmlElement(f"w:{borde}")
        elemento.set(qn("w:val"), "nil")
        tblBorders.append(elemento)


def _insertar_imagenes_en_marcador_grid(
    doc,
    marcador: str,
    rutas_imagenes: list[Path],
    columnas: int = 4,
    ancho_fila_mm: float = 150,
    ancho_maximo_mm: float = 70,
) -> bool:
    """Igual que _insertar_imagenes_en_marcador pero acomoda las fotos en una
    cuadrícula (varias por fila, una al lado de la otra) usando una tabla sin
    bordes, en vez de apilarlas una debajo de otra.

    El ancho de cada foto NO es fijo: se calcula repartiendo `ancho_fila_mm`
    entre las columnas que realmente se usan (si hay menos fotos que
    `columnas`, cada una queda más grande en vez de dejar espacio vacío en
    la fila) -- con un tope de `ancho_maximo_mm` para que no se vean
    gigantes cuando hay solo 1 o 2 fotos."""
    rutas_imagenes = [r for r in rutas_imagenes if r and r.is_file()]
    if not rutas_imagenes:
        return False

    parrafo_marcador = None
    for p in doc.paragraphs:
        if marcador in p.text:
            parrafo_marcador = p
            break
    if parrafo_marcador is None:
        return False

    for run in list(parrafo_marcador.runs):
        run.text = ""

    columnas_usadas = min(columnas, len(rutas_imagenes))
    ancho_mm = min(ancho_maximo_mm, ancho_fila_mm / columnas_usadas)

    filas_rutas = [rutas_imagenes[i : i + columnas] for i in range(0, len(rutas_imagenes), columnas)]
    tabla = doc.add_table(rows=len(filas_rutas), cols=columnas)
    tabla.autofit = True
    _quitar_bordes_tabla(tabla)

    for fila_idx, fila_rutas in enumerate(filas_rutas):
        for col_idx in range(columnas):
            celda = tabla.cell(fila_idx, col_idx)
            parrafo_celda = celda.paragraphs[0]
            parrafo_celda.paragraph_format.space_after = Pt(6)
            if col_idx < len(fila_rutas):
                run = parrafo_celda.add_run()
                try:
                    run.add_picture(str(fila_rutas[col_idx]), width=Mm(ancho_mm))
                except Exception:
                    # Foto corrupta/ilegible: se deja la celda vacía en vez
                    # de tumbar todo el script.
                    parrafo_celda.add_run("(foto no disponible)")

    # Mueve la tabla justo después del párrafo donde estaba el marcador
    # (add_table() la agrega al final del documento por defecto).
    parrafo_marcador._p.addnext(tabla._tbl)
    return True


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


def _reemplazar_en_todo_el_documento(doc, reemplazos: dict) -> None:
    """Reemplazo "de respaldo" que recorre TODOS los nodos de texto del
    .docx, incluidos los que viven dentro del bloque especial de la Tabla
    de Contenido (Word la guarda en un "content control" que
    doc.paragraphs/doc.tables NO recorren, así que un reemplazo hecho solo
    ahí deja intacto el texto cacheado del índice -- por eso el título del
    anexo se actualizaba en el cuerpo del Acta pero no en el "4. ANEXOS"
    del índice). Solo sirve para textos que quedan completos dentro de un
    mismo nodo <w:t> (como el FMI viejo hardcodeado); los marcadores
    ##...## que pueden quedar repartidos en varios runs de un párrafo
    normal se siguen manejando con _reemplazar_texto_parrafo."""
    for nodo in doc.element.body.iter(qn("w:t")):
        texto = nodo.text or ""
        nuevo = texto
        for buscar, valor in reemplazos.items():
            if buscar in nuevo:
                nuevo = nuevo.replace(buscar, str(valor))
        if nuevo != texto:
            nodo.text = nuevo


def generar_caso(caso: dict, identificador: str | int) -> Path:
    """caso trae las RUTAS de los soportes; todo el resto se extrae solo.

    `identificador` ya NO es un consecutivo de Acta -- ese número lo asigna
    la Coordinación/Comité por fuera de este script (así lo pidió el
    usuario). Aquí solo se usa como sufijo interno para que los archivos
    temporales de imágenes de cada caso no se pisen entre sí (puede ser
    cualquier texto/número único, ej. el FMI o el índice del caso en el
    lote)."""
    datos_estimado = extraer_de_estimado_renta(
        caso["estimado_renta"],
        fmi_conocido=caso.get("fmi_conocido"),
        territorial_conocida=caso.get("territorial_conocida"),
    )

    # "fmi_forzado": a diferencia de "fmi_conocido" (que solo se usa si la
    # extracción del texto falla), este SIEMPRE reemplaza el FMI extraído.
    # Se usa cuando una misma carpeta/documento cubre varios FMI y se está
    # generando una Acta separada para cada uno -- así cada Acta queda con
    # el FMI que corresponde, aunque el documento mencione varios.
    if caso.get("fmi_forzado"):
        datos_estimado["fmi"] = caso["fmi_forzado"]
        m_fmi_limpio = re.match(r"([\d]{2,3}[A-Za-z]?-\d+)", caso["fmi_forzado"])
        datos_estimado["fmi_limpio"] = m_fmi_limpio.group(1) if m_fmi_limpio else caso["fmi_forzado"]

    pendientes = []

    if caso.get("aprobado"):
        datos_aprobado = extraer_de_aprobado(caso["aprobado"])
    else:
        # No hay Aprobado/Póliza -- se prueban, EN ORDEN, todos los demás
        # documentos de la carpeta que puedan traer nombre/cédula/ciudad de
        # expedición (ver extraer_de_solicitud_arrendamiento y
        # extraer_de_carta_juramentada). No basta con que el archivo
        # exista: puede estar escaneado sin texto (sin OCR no se puede
        # leer) y entonces la función de extracción devuelve None -- en
        # ese caso se sigue probando con el siguiente respaldo antes de
        # rendirse y dejarlo en "—".
        datos_aprobado = None
        fuente_respaldo = None
        for candidato, extractor, nombre_fuente in (
            (caso.get("solicitud_arrendamiento"), extraer_de_solicitud_arrendamiento, "SOLICITUD DE ARRENDAMIENTO"),
            (caso.get("carta_juramentada"), extraer_de_carta_juramentada, "carta/declaración juramentada"),
        ):
            if candidato:
                datos_aprobado = extractor(candidato)
                if datos_aprobado:
                    fuente_respaldo = nombre_fuente
                    break
        if datos_aprobado:
            datos_aprobado.setdefault("codeudor_nombre", None)
            datos_aprobado.setdefault("codeudor_cedula", None)
            pendientes.append(
                "No se encontró el documento de Aprobado/Póliza del arrendatario en la carpeta; el nombre, "
                f"cédula/NIT y ciudad de expedición se tomaron de la '{fuente_respaldo}' en vez de eso -- "
                "verificar que coincidan."
            )
        else:
            datos_aprobado = {
                "arrendatario_nombre": "—",
                "id_numero": "—",
                "id_siglas": "C.C.",
                "codeudor_nombre": None,
                "codeudor_cedula": None,
            }
            pendientes.append(
                "No se encontró el documento de Aprobado/Póliza del arrendatario en la carpeta (ni otro "
                "documento legible con esos datos, como la Solicitud de Arrendamiento o la carta/declaración "
                "juramentada -- puede que existan pero estén escaneados como imagen sin texto); el nombre y "
                "número de cédula/NIT del arrendatario quedaron en '—', revisar y completar a mano."
            )

    tipo_contrato = caso.get("tipo_contrato_forzado") or f"CONTRATO {datos_estimado['categoria_contrato']}".strip()

    descripcion = caso.get("descripcion_forzada") or extraer_descripcion_estimado_renta(caso["estimado_renta"])
    if not descripcion:
        descripcion = (
            "No se pudo extraer automáticamente la descripción del Estimado de Renta; "
            "revisar y pegar a mano el párrafo de 'DESCRIPCIÓN' del PDF."
        )
        pendientes.append("Descripción del inmueble: no se pudo extraer automáticamente, revisar a mano.")

    # Fotos reales del inmueble (recortadas del "REGISTRO FOTOGRÁFICO" del
    # Estimado de Renta, en TODAS sus páginas, filtrando logos/firmas y la
    # página completa escaneada -- ver AREA_MINIMA_FOTO y
    # PROPORCION_MAXIMA_PAGINA_COMPLETA). TODAS se insertan juntas en
    # "##foto_estimado_renta##" (una debajo de otra), así haya 1 o varias.
    # Si no se encuentra ninguna, se deja el campo sin imagen y con un
    # pendiente -- YA NO se usa como respaldo una captura de la página
    # completa, porque eso no es "una foto del inmueble" sino el documento
    # entero (tablas, firmas, etc.), y sería engañoso insertarlo como si lo
    # fuera.
    fotos_inmueble = extraer_fotos_inmueble(caso["estimado_renta"], f"estimado_{identificador}")

    if caso.get("aprobado"):
        imagen_aprobado = _pdf_a_png(caso["aprobado"], f"aprobado_{identificador}")
    else:
        imagen_aprobado = None
    imagen_sagrilaft = caso.get("sagrilaft_imagen")  # ya viene como imagen (png/jpg)

    doc = docx.Document(str(PLANTILLA))

    def _ajustar_texto_literal(texto: str) -> str:
        return texto.replace(FMI_PLANTILLA_VIEJO, datos_estimado["fmi"])

    for p in doc.paragraphs:
        nuevo = _ajustar_texto_literal(p.text)
        if nuevo != p.text and p.runs:
            p.runs[0].text = nuevo
            for r in p.runs[1:]:
                r.text = ""
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    nuevo = _ajustar_texto_literal(p.text)
                    if nuevo != p.text and p.runs:
                        p.runs[0].text = nuevo
                        for r in p.runs[1:]:
                            r.text = ""

    # Respaldo: la entrada del FMI viejo que queda cacheada dentro de la
    # Tabla de Contenido (Word la guarda en un bloque que los dos loops de
    # arriba no alcanzan) -- ver _reemplazar_en_todo_el_documento.
    _reemplazar_en_todo_el_documento(doc, {FMI_PLANTILLA_VIEJO: datos_estimado["fmi"]})

    id_ciudad = caso.get("id_ciudad_forzada") or datos_aprobado.get("id_ciudad") or datos_estimado["ciudad"]
    if not caso.get("id_ciudad_forzada") and not datos_aprobado.get("id_ciudad"):
        pendientes.append(
            f"Ciudad de expedición de la cédula: se asumió '{id_ciudad}' (la misma del inmueble) "
            "porque esa fecha solo aparece en la cédula física escaneada; verificar contra la cédula."
        )

    mapa_texto = {
        # Ya NO se pone un número aquí -- el consecutivo del Acta lo asigna
        # la Coordinación/Comité por fuera de este script (así lo pidió el
        # usuario), así que el "ACTA No. ___ - 2026" queda en blanco para
        # que ellos lo completen a mano.
        "##numero_consecutivo##": "____",
        "##descripcion_estimado_renta##": descripcion,
        "##fmi_estimado_renta##": datos_estimado["fmi"],
        "##tipo_bien_estimado_renta##": datos_estimado["tipo_bien"],
        "##contrato##": tipo_contrato,
        "##carpeta_arrendatario_nombre##": datos_aprobado["arrendatario_nombre"],
        "##cedula/nit_siglas##": caso.get("id_siglas") or datos_aprobado.get("id_siglas", "C.C."),
        "##cedula/nit_numero##": datos_aprobado["id_numero"],
        "##cedula/nit_ciudad##": id_ciudad,
    }

    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = fila.cells
            if len(celdas) < 2:
                continue
            etiqueta = celdas[0].text.strip().lower()
            if etiqueta.startswith("direccion territorial"):
                _reemplazar_texto_parrafo(celdas[1].paragraphs[0], {"##direccion_estimado_renta##": datos_estimado["direccion_territorial"]})
            elif etiqueta.startswith("direccion:"):
                _reemplazar_texto_parrafo(celdas[1].paragraphs[0], {"##direccion_estimado_renta##": datos_estimado["direccion"]})

    for p in doc.paragraphs:
        _reemplazar_texto_parrafo(p, mapa_texto)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    _reemplazar_texto_parrafo(p, mapa_texto)

    # Texto que se deja en vez de la variable "##...##" cuando no se
    # encuentra ninguna imagen para ese marcador -- así nunca queda visible
    # en el Acta el nombre literal de la variable.
    SIN_IMAGEN = "No hay imagen para este documento"

    for marcador, ruta in {
        "##carpeta_foto_aprobado_sagrilaft##": imagen_sagrilaft,
        "##archivo_foto_preaprobado_poliza##": imagen_aprobado,
    }.items():
        if not _insertar_imagen_en_marcador(doc, marcador, ruta):
            for p in doc.paragraphs:
                _reemplazar_texto_parrafo(p, {marcador: SIN_IMAGEN})
            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        for p in celda.paragraphs:
                            _reemplazar_texto_parrafo(p, {marcador: SIN_IMAGEN})
            pendientes.append(f"No se pudo insertar la imagen de '{marcador}', revisar a mano.")

    # TODAS las fotos del inmueble (1 o varias) van juntas en este marcador.
    # Si no se encontró ninguna foto real, se deja el texto SIN_IMAGEN (NO
    # el texto literal "##foto_estimado_renta##") con un pendiente para
    # agregarla a mano -- ver nota arriba sobre por qué ya no se usa una
    # captura de la página completa como respaldo.
    marcador_foto = "##foto_estimado_renta##"
    if not _insertar_imagenes_en_marcador_grid(doc, marcador_foto, fotos_inmueble, columnas=6):
        for p in doc.paragraphs:
            _reemplazar_texto_parrafo(p, {marcador_foto: SIN_IMAGEN})
        pendientes.append(
            "Foto del inmueble: no se encontró ninguna foto real del 'REGISTRO FOTOGRÁFICO' en el "
            "Estimado de Renta (o el documento viene escaneado como una sola imagen de página completa, "
            "que no se usa por no ser una foto real); agregar manualmente."
        )

    # Este marcador vive bajo el encabezado "ESTIMADO DE RENTA" (sección
    # aparte de "REGISTRO FOTOGRÁFICO DEL INMUEBLE"): aquí va el documento
    # completo del Estimado de Renta (todas sus páginas, como imagen) --
    # confirmado por el usuario que así está bien.
    marcador_adicionales = "##foto_estimado_renta_(si hay mas se colocan todas)##"
    paginas_estimado_renta = _pdf_a_png_todas_paginas(caso["estimado_renta"], f"docestimado_{identificador}")
    if not _insertar_imagenes_en_marcador(doc, marcador_adicionales, paginas_estimado_renta):
        for p in doc.paragraphs:
            _reemplazar_texto_parrafo(p, {marcador_adicionales: SIN_IMAGEN})
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    for p in celda.paragraphs:
                        _reemplazar_texto_parrafo(p, {marcador_adicionales: SIN_IMAGEN})
        pendientes.append(
            "Documento del Estimado de Renta: no se pudo convertir a imagen para la sección "
            "'ESTIMADO DE RENTA'; agregar manualmente."
        )

    if datos_aprobado.get("codeudor_nombre"):
        pendientes.append(
            f"Este caso tiene codeudor solidario ({datos_aprobado['codeudor_nombre']}, "
            f"CC {datos_aprobado['codeudor_cedula']}); la plantilla actual no tiene un campo para "
            "codeudor, revisar si hay que agregarlo a mano."
        )
    pendientes.append(f"Tipo de contrato: se puso '{tipo_contrato}' (inferido del tipo de bien); confirmar.")
    pendientes.append("Fecha de inicio / fecha final / duración / canon mensual / canon total: quedan en '(HUMANO)'.")
    pendientes.append("Número de Acta (\"ACTA No. ___ - 2026\"): queda en blanco -- lo asigna la Coordinación/Comité.")
    for nota in caso.get("pendientes_extra", []):
        pendientes.append(nota)

    nombre_archivo = f"ACTA_ARRENDAMIENTO_{datos_estimado['fmi_limpio'].replace('/', '_').replace(' ', '_')}.docx"
    ruta_salida = SALIDA / nombre_archivo
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ruta_salida))

    print(f"[{identificador}] {nombre_archivo}")
    for msg in pendientes:
        print(f"    - {msg}")
    return ruta_salida


# ── Búsqueda interactiva por FMI (un solo script, sin archivo aparte) ──────
# Antes esta parte (buscar la carpeta del caso a partir del FMI, pedirlo por
# consola, etc.) vivía en un segundo archivo ("generar_actas_arriendo_lote1.py")
# que importaba de este -- se fusionó todo aquí para tener un solo script que
# usar (así no hay que acordarse cuál de los dos correr).

_NOMBRE_CARPETA = "actas_arriendo"


def _detectar_base() -> Path:
    candidatos = []
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        valor = os.environ.get(var)
        if valor:
            candidatos.append(Path(valor) / "Documentos" / _NOMBRE_CARPETA)
            candidatos.append(Path(valor) / _NOMBRE_CARPETA)
    candidatos.append(Path(
        "/mnt/user-data/uploads/OneDrive - Activos por Colombia/Documentos/" + _NOMBRE_CARPETA
    ))
    for c in candidatos:
        if c.is_dir():
            return c
    raise FileNotFoundError(
        f"No se encontró la carpeta '{_NOMBRE_CARPETA}'. Verifica que el OneDrive "
        f"de Activos por Colombia esté sincronizado en este computador."
    )


def _normalizar_fmi(s: str) -> str:
    """Deja solo dígitos y letras (sin guiones, espacios, etc.) para poder
    comparar FMI escritos con formatos distintos."""
    return re.sub(r"[^0-9A-Za-z]", "", s).upper()


def _normalizar_texto(s: str) -> str:
    """Igual que _normalizar_fmi pero conservando letras (sin acentos,
    en mayúsculas) -- para comparar nombres de arrendatario/empresa."""
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper()


def _separar_items(entrada: str) -> list[str]:
    """Separa varias búsquedas escritas en una sola línea, por coma o
    punto y coma (ej. "294-91489 Juan Rios, 290-158476 Juan Rios")."""
    return [p.strip() for p in re.split(r"[,;]+", entrada.strip()) if p.strip()]


def _parsear_item(item: str) -> tuple[str, str | None]:
    """Separa un ítem en (FMI, nombre opcional): el FMI es la primera
    palabra, y todo lo que sigue (si hay algo) se toma como el nombre del
    arrendatario/empresa para filtrar cuando el FMI por sí solo da muchas
    coincidencias (ej. "294-91489 Juan Rios" -> ("294-91489", "Juan Rios"))."""
    partes = item.strip().split(None, 1)
    if not partes:
        return "", None
    fmi = partes[0]
    nombre = partes[1].strip() if len(partes) > 1 else None
    return fmi, nombre


def buscar_carpetas_por_fmi(fmi_buscado: str, nombre_filtro: str | None = None) -> list[Path]:
    """Busca, en TODAS las carpetas bajo las territoriales de
    actas_arriendo, aquellas cuyo nombre contenga el FMI dado (comparando
    versiones normalizadas, sin importar guiones/espacios). Descarta
    coincidencias que sean subcarpetas de otra coincidencia ya encontrada
    (para no listar dos veces el mismo caso).

    Si el FMI por sí solo da varias coincidencias (pasa con matrículas
    cortas, que quedan contenidas dentro de otras más largas) y se dio un
    `nombre_filtro`, se reduce la lista a las carpetas cuyo nombre también
    contenga ese nombre -- así no hay que elegir a mano cada vez."""
    clave = _normalizar_fmi(fmi_buscado)
    if not clave:
        return []
    encontradas = []
    for p in BASE.rglob("*"):
        if p.is_dir() and clave in _normalizar_fmi(p.name):
            encontradas.append(p)
    # Si una carpeta encontrada es subcarpeta de otra también encontrada,
    # se queda solo la de más arriba (el caso, no una sub-subcarpeta suya).
    resultado = [p for p in encontradas if not any(otra != p and otra in p.parents for otra in encontradas)]
    if nombre_filtro and len(resultado) > 1:
        clave_nombre = _normalizar_texto(nombre_filtro)
        filtrado = [p for p in resultado if clave_nombre in _normalizar_texto(p.name)]
        if filtrado:
            resultado = filtrado
    return resultado


def _tokens_significativos(nombre_carpeta: str) -> set[str]:
    """Extrae las palabras "importantes" del nombre de una carpeta de
    caso (el nombre del arrendatario/empresa), quitando el prefijo
    numérico, la palabra FMI y los números de matrícula -- para poder
    comparar si dos carpetas distintas son en realidad la misma unidad/
    negocio (ej. "04_FMI 001-315387 FUNDALUVA" y "13_FMI 020-70965
    FUNDALUVA" comparten el token "FUNDALUVA")."""
    normal = unicodedata.normalize("NFKD", nombre_carpeta).encode("ascii", "ignore").decode().upper()
    normal = re.sub(r"^\d+_?\s*", "", normal)  # quita el prefijo "07_"
    normal = re.sub(r"\bFMI\b", " ", normal)
    normal = re.sub(r"[0-9\-]+", " ", normal)  # quita números/guiones (matrículas)
    return {w for w in re.split(r"[^A-Z]+", normal) if len(w) > 3}


def _agrupar_por_unidad(resueltos: list[tuple[str, Path]]) -> list[dict]:
    """Agrupa los (fmi, carpeta) resueltos: si dos terminan en la MISMA
    carpeta, o en carpetas distintas que comparten el nombre del
    arrendatario/empresa, se tratan como una sola unidad/negocio -- una
    sola Acta, con todos los FMI que cubre anotados en pendientes.

    (No se usa por defecto en procesar_entrada -- el usuario pidió que cada
    FMI genere su propia Acta -- pero se deja disponible por si se necesita
    de nuevo más adelante.)"""
    grupos: list[dict] = []
    for fmi, carpeta in resueltos:
        tokens = _tokens_significativos(carpeta.name)
        grupo_encontrado = None
        for g in grupos:
            if carpeta in g["carpetas"]:
                grupo_encontrado = g
                break
            if tokens & g["tokens"]:
                grupo_encontrado = g
                break
        if grupo_encontrado:
            if carpeta not in grupo_encontrado["carpetas"]:
                grupo_encontrado["carpetas"].append(carpeta)
            grupo_encontrado["fmis"].append(fmi)
            grupo_encontrado["tokens"] |= tokens
        else:
            grupos.append({"carpetas": [carpeta], "fmis": [fmi], "tokens": tokens})
    return grupos


# ── Búsqueda de archivos dentro de la carpeta del caso ────────────────────

def _archivo_estimado_renta(carpeta: Path) -> Path | None:
    for claves in (
        ("ESTIMADO DE RENTA",),
        ("ESTRIMADO DE RENTA",),  # typo real visto en un caso
        ("ACTI RENTA",),
        ("ACTIRENTA",),  # mismo documento, sin espacio ("ACTIRENTA...")
        ("ER FMI",),
    ):
        f = _buscar_archivo(carpeta, *claves)
        if f:
            return f
    return None


def _archivo_aprobado_poliza(carpeta: Path) -> Path | None:
    carpetas_donde_buscar = [carpeta]
    # "ARRENDTARIO" (sin la "A") es un typo real visto en una carpeta (Eyda
    # Tenorio, FMI 024-20085) -- se acepta también para no dejar de
    # encontrar la subcarpeta por eso.
    carpeta_arrendatario = _buscar_subcarpeta(carpeta, "ARRENDATARIO", "ARRENDTARIO")
    if carpeta_arrendatario:
        carpetas_donde_buscar.append(carpeta_arrendatario)

    for c in carpetas_donde_buscar:
        for claves in (
            ("APROBADO POLIZA",),
            ("PREAPROBADO DE POLIZA",),
            ("PREAPROBADO POLIZA",),
        ):
            f = _buscar_archivo(c, *claves)
            if f:
                return f
        # Respaldo: cualquier archivo "APROBADO ..." que no sea el de
        # SAGRILAFT (ese es la aprobación de lista restrictiva, no trae los
        # datos del arrendatario en el formato que se necesita aquí).
        candidatos = sorted(p for p in c.iterdir() if p.is_file())
        for f in candidatos:
            n = f.name.upper()
            if "APROBADO" in n and "SAGRILAFT" not in n:
                return f
    return None


def _archivo_solicitud_arrendamiento(carpeta: Path) -> Path | None:
    """Respaldo cuando no hay documento de Aprobado/Póliza: la "SOLICITUD DE
    ARRENDAMIENTO" (normalmente dentro de la subcarpeta "ARRENDATARIO ...")
    trae también el nombre/cédula/ciudad de expedición -- ver
    extraer_de_solicitud_arrendamiento()."""
    # "ARRENDTARIO" (sin la "A") es un typo real visto en una carpeta (Eyda
    # Tenorio, FMI 024-20085) -- se acepta también para no dejar de
    # encontrar la subcarpeta por eso.
    carpeta_arrendatario = _buscar_subcarpeta(carpeta, "ARRENDATARIO", "ARRENDTARIO")
    if carpeta_arrendatario:
        f = _buscar_archivo(carpeta_arrendatario, "SOLICITUD DE ARRENDAMIENTO")
        if f:
            return f
    return _buscar_archivo(carpeta, "SOLICITUD DE ARRENDAMIENTO")


def _archivo_carta_juramentada(carpeta: Path) -> Path | None:
    """Otro respaldo, además de la Solicitud de Arrendamiento: cualquier
    archivo con "JURAMENTADA" en el nombre (cubre "CARTA DECLARACIÓN
    JURAMENTADA", "FORMATO Declara Juramentada", etc. -- se usa esa única
    palabra por ser lo bastante distintiva y porque el resto del nombre
    varía mucho de una carpeta a otra) -- ver extraer_de_carta_juramentada()."""
    carpeta_arrendatario = _buscar_subcarpeta(carpeta, "ARRENDATARIO", "ARRENDTARIO")
    if carpeta_arrendatario:
        f = _buscar_archivo(carpeta_arrendatario, "JURAMENTADA")
        if f:
            return f
    return _buscar_archivo(carpeta, "JURAMENTADA")


def _imagen_sagrilaft(carpeta: Path, nombre_salida: str) -> Path | None:
    """Busca el soporte de SAGRILAFT aprobado. A veces está como imagen
    (png/jpg) y a veces como PDF ("APROBADO SAGRILAFT.pdf") -- en ese
    segundo caso se convierte la primera página a PNG (igual que se hace
    con el Aprobado/Póliza), para poder insertarla como imagen en el Acta."""
    # "ARRENDTARIO" (sin la "A") es un typo real visto en una carpeta (Eyda
    # Tenorio, FMI 024-20085) -- se acepta también para no dejar de
    # encontrar la subcarpeta por eso.
    carpeta_arrendatario = _buscar_subcarpeta(carpeta, "ARRENDATARIO", "ARRENDTARIO")
    encontrado = None
    if carpeta_arrendatario:
        encontrado = _buscar_archivo(carpeta_arrendatario, "SAGRILAFT")
    if not encontrado:
        encontrado = _buscar_archivo(carpeta, "SAGRILAFT")
    if not encontrado:
        return None
    if encontrado.suffix.lower() == ".pdf":
        return _pdf_a_png(encontrado, nombre_salida)
    return encontrado


def _territorial_de_ruta(carpeta: Path) -> str | None:
    """Respaldo para cuando el Estimado de Renta no trae la territorial en
    su texto (pasa con el formato "Dictamen Comercial y Financiero"): las
    carpetas de actas_arriendo están organizadas por territorial en el
    nombre de una carpeta ancestro (ej. ".../02_TERRITORIAL OCCIDENTE/..."),
    así que se busca ahí en vez de dejarlo en '—'."""
    for ancestro in carpeta.parents:
        m = re.search(
            r"TERRITORIAL\s+(CARIBE|OCCIDENTE|SUR|CENTRO ORIENTE|NORTE|CENTRO)",
            ancestro.name, re.IGNORECASE,
        )
        if m:
            return m.group(1).upper()
    return None


def _construir_caso_desde_carpeta(carpeta: Path, fmi_conocido: str, indice: int, fmis_extra: list[str] | None = None, carpetas_extra: list[Path] | None = None, fmi_forzado: str | None = None) -> dict:
    """`indice` es solo un identificador interno para que los archivos
    temporales de imágenes de cada caso no se pisen entre sí -- ya NO es
    un consecutivo de Acta (eso lo asigna la Coordinación/Comité).

    `fmi_forzado`: cuando se pide una Acta separada por cada FMI (aunque
    compartan carpeta/documento), este es el FMI que debe quedar en ESA
    Acta -- reemplaza lo que se extraiga del texto del documento."""
    estimado = _archivo_estimado_renta(carpeta)
    aprobado = _archivo_aprobado_poliza(carpeta)
    sagrilaft = _imagen_sagrilaft(carpeta, f"sagrilaft_{indice}")
    # Respaldos cuando no hay documento de Aprobado/Póliza (p. ej. Cindy
    # Murillo, FMI 001-389120, o Eyda Tenorio, FMI 024-20085): otros
    # documentos de la carpeta pueden traer el nombre/cédula/ciudad de
    # expedición en formatos distintos -- ver extraer_de_solicitud_
    # arrendamiento() y extraer_de_carta_juramentada(). Se buscan los DOS
    # (no solo el primero que aparezca) porque un archivo puede existir
    # pero venir escaneado sin texto (sin OCR no se puede leer) -- en ese
    # caso generar_caso() sigue probando con el otro en vez de rendirse.
    solicitud_arrendamiento = None if aprobado else _archivo_solicitud_arrendamiento(carpeta)
    carta_juramentada = None if aprobado else _archivo_carta_juramentada(carpeta)
    territorial_conocida = _territorial_de_ruta(carpeta)

    if estimado is None:
        # El Estimado de Renta es indispensable (de ahí sale FMI, tipo de bien,
        # dirección, descripción y fotos) -- sin eso no hay caso que generar.
        raise FileNotFoundError(f"No se encontró el Estimado de Renta en: {carpeta}")
    # El documento de Aprobado/Póliza puede faltar (p. ej. Juan Gregory Blandón,
    # que solo tiene el soporte de SAGRILAFT en su carpeta). En ese caso se
    # genera igual el Acta, con el nombre/cédula del arrendatario en "—" y un
    # pendiente para completar a mano, en vez de saltarse el caso entero.

    caso = dict(
        estimado_renta=estimado,
        aprobado=aprobado,
        solicitud_arrendamiento=solicitud_arrendamiento,
        carta_juramentada=carta_juramentada,
        sagrilaft_imagen=sagrilaft,
        fmi_conocido=fmi_conocido,
        territorial_conocida=territorial_conocida,
    )
    if fmi_forzado:
        caso["fmi_forzado"] = fmi_forzado
    if fmis_extra:
        nombres_extra = ", ".join(f"{fmi} ({c.name})" for fmi, c in zip(fmis_extra, carpetas_extra or []))
        caso["pendientes_extra"] = [
            f"Este caso también cubre el/los FMI adicional(es) {nombres_extra}, detectado(s) como la "
            "misma unidad/negocio por compartir nombre de arrendatario/empresa en la carpeta; verificar "
            "que no se necesite mencionarlos en el cuerpo del Acta."
        ]
    return caso


def procesar_entrada(entrada: str, indice: int) -> int:
    """Procesa un texto de entrada con uno o varios FMI (separados por coma,
    opcionalmente con nombre) -- busca cada carpeta, genera su Acta, y
    devuelve el índice actualizado."""
    items = _separar_items(entrada)
    resueltos: list[tuple[str, Path]] = []
    for item in items:
        fmi, nombre_filtro = _parsear_item(item)
        candidatos = buscar_carpetas_por_fmi(fmi, nombre_filtro)
        if not candidatos:
            sufijo_nombre = f" y el nombre '{nombre_filtro}'" if nombre_filtro else ""
            print(f"  ✗ No se encontró ninguna carpeta con el FMI '{fmi}'{sufijo_nombre}.")
            continue
        if len(candidatos) > 1:
            print(f"  Se encontró más de una carpeta para '{item}':")
            for i, c in enumerate(candidatos, 1):
                print(f"    {i}. {c.relative_to(BASE)}")
            if not nombre_filtro:
                print("  (Tip: si le agregas el nombre del arrendatario/empresa después del FMI, "
                      "por lo general encuentra el caso exacto sin tener que elegir.)")
            seleccion = input("  ¿Cuál es? (número): ").strip()
            try:
                carpeta = candidatos[int(seleccion) - 1]
            except (ValueError, IndexError):
                print("  Selección inválida, se omite este FMI.")
                continue
        else:
            carpeta = candidatos[0]
        resueltos.append((fmi, carpeta))

    if not resueltos:
        return indice

    # Se genera UNA Acta por cada FMI ingresado, aunque dos o más
    # compartan la misma carpeta/documento (p. ej. un mismo predio con
    # dos matrículas) -- así lo pidió el usuario explícitamente. Cada
    # Acta queda con su propio FMI forzado (fmi_forzado), aunque el
    # documento del Estimado de Renta mencione varios.
    for fmi, carpeta in resueltos:
        print(f"\n── {carpeta.name}  (FMI {fmi}) ──")

        try:
            caso = _construir_caso_desde_carpeta(
                carpeta, fmi, indice, fmi_forzado=fmi,
            )
        except FileNotFoundError as e:
            print(f"  ✗ {e}")
            indice += 1
            continue

        print(f"  Estimado de Renta: {caso['estimado_renta'].name}")
        print(f"  Aprobado/Póliza:   {caso['aprobado'].name if caso['aprobado'] else '(NO ENCONTRADO -- queda pendiente a mano)'}")
        print(f"  SAGRILAFT imagen:  {caso['sagrilaft_imagen'].name if caso['sagrilaft_imagen'] else '(no encontrada)'}")

        ruta_salida = generar_caso(caso, indice)
        print(f"  -> {ruta_salida.name}")
        indice += 1

    return indice


def main():
    print(f"Buscando dentro de: {BASE}\n")

    # Si se pasan argumentos por línea de comandos (ej.
    # "python generar_acta_arrendamiento.py 294-91489, 294-91407"), se
    # procesan de una vez y el script termina.
    if len(sys.argv) > 1:
        procesar_entrada(" ".join(sys.argv[1:]), indice=1)
        print("\nListo.")
        return

    indice = 1
    while True:
        entrada = input(
            "FMI a buscar -- opcionalmente con el nombre después, para que no dé varias "
            "coincidencias (ej. \"294-91489 Juan Rios\"); varios separados por coma; "
            "Enter/'salir' para terminar: "
        ).strip()
        if not entrada or entrada.lower() in ("salir", "exit", "q"):
            break

        indice = procesar_entrada(entrada, indice)

    print("\nListo.")


BASE = _detectar_base()


if __name__ == "__main__":
    main()
