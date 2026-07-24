"""
Genera el INFORME DE SUBASTA ELECTRÓNICA (.docx).
Uso: python scripts/generar_informe.py <auction_uuid_o_codigo>
Requisito: INFORME_SUBASTA.docx en templates/.
El archivo se guarda en output/informes/
"""
from __future__ import annotations
import re, sys, time, zipfile
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
CORE_DIR = PROYECTO / "core"
PLANTILLA = PROYECTO / "templates" / "INFORME_SUBASTA.docx"
SALIDA    = PROYECTO / "output" / "informes"

MESES_ABR = {"ene":"01","feb":"02","mar":"03","abr":"04","may":"05","jun":"06",
             "jul":"07","ago":"08","sep":"09","oct":"10","nov":"11","dic":"12"}

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

def _colapsar_xml(xml):
    xml = re.sub(
        r'##fecha_cronograma</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>(\d+)</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>##</w:t></w:r>',
        r'##fecha_cronograma\1##', xml, flags=re.DOTALL)
    xml = re.sub(
        r'<w:t>##</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>([^<#]+)</w:t></w:r><w:r[^>]*><w:rPr>.*?</w:rPr><w:t>##</w:t></w:r>',
        r'<w:t>##\1##</w:t></w:r>', xml, flags=re.DOTALL)
    xml = re.sub(r'<w:t>([^<]+)</w:r>', r'<w:t>\1</w:t></w:r>', xml)
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
        for a in ["--headless","--no-sandbox","--disable-dev-shm-usage","--disable-gpu"]:
            opts.add_argument(a)
        opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

        print(f"  → {url}")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        try:
            driver.get(url); time.sleep(3)
            lineas = [l.strip() for l in driver.find_element(By.TAG_NAME,"body").text.split("\n")]

            anio = "2026"
            for l in lineas:
                m = re.search(r"20\d{2}", l)
                if m: anio = m.group(); break

            inicio_cron = next((i for i,l in enumerate(lineas) if "Cronograma del proceso" in l), 0)

            fases = [
                ("Publicación próxima en subasta", "fecha_cronograma1",  "fecha_cronograma2"),
                ("Registro",                       "fecha_cronograma3",  "fecha_cronograma4"),
                ("Análisis debida diligencia",     "fecha_cronograma5",  "fecha_cronograma6"),
                ("Análisis financiero",            "fecha_cronograma7",  "fecha_cronograma8"),
                ("Expedición y envío de cupones",  "fecha_cronograma9",  "fecha_cronograma10"),
                ("seriedad",                       "fecha_cronograma11", "fecha_cronograma12"),
                ("Validación y confirmación",      "fecha_cronograma13", "fecha_cronograma14"),
                ("Subasta (apertura y cierre)",    "fecha_cronograma15", "fecha_cronograma16"),
            ]

            for fase, ph_ini, ph_fin in fases:
                for i,l in enumerate(lineas):
                    if i < inicio_cron: continue
                    if fase.lower() in l.lower():
                        for offset in [1,2,3]:
                            cand = lineas[i+offset] if i+offset < len(lineas) else ""
                            if cand and cand[0].isdigit():
                                mes_l = lineas[i+offset+1] if i+offset+1 < len(lineas) else ""
                                mes = MESES_ABR.get(mes_l.split(".")[0].strip(),"00")
                                if "→" in cand:
                                    partes = cand.split()
                                    dia_i, dia_f = partes[0].zfill(2), partes[2].zfill(2)
                                else:
                                    dia_i = dia_f = cand.split()[0].zfill(2)
                                res[ph_ini] = f"{dia_i}/{mes}/{anio}"
                                res[ph_fin]  = f"{dia_f}/{mes}/{anio}"
                                if fase == "Publicación próxima en subasta":
                                    res["fecha_publicacion"] = f"{dia_i}/{mes}/{anio} 10:00 am"
                                break
                        break

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
    sys.path.insert(0, str(CORE_DIR))
    from core import get_connection, fetch_informe, load_dotenv_files
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
    print("✓ Conectado a la base de datos")

    auction_uuid = _resolver_uuid(conn, auction_uuid_input)
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
        cur.execute("SELECT inmueble_id FROM polibid_subastas_v2 WHERE auction_id=%s::uuid ORDER BY id DESC LIMIT 1",(auction_uuid,))
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

    return {
        "codigo_subasta":    subasta.get("code","—"),
        "fecha_publicacion": web.get("fecha_publicacion","—"),
        "fecha_inicio":      web.get("fecha_inicio", _fmt_fecha(subasta.get("start_date"))),
        "fecha_fin":         web.get("fecha_fin", _fmt_fecha(subasta.get("end_date"))),
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
        print("Uso: python generar_informe.py <auction_uuid_o_codigo>"); sys.exit(1)
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
