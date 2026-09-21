"""
Captura automática de datos de Microsoft Clarity para el informe de subasta.

Cómo funciona (confirmado navegando el dashboard real de Clarity)
------------------------------------------------------------------
Clarity guarda el filtro de fechas Y el filtro de ruta (URL) directamente
en la URL del dashboard, así que NO hace falta simular clics en calendarios
ni menús — basta con construir el enlace correcto y navegar a él:

    https://clarity.microsoft.com/projects/view/<PROJECT_ID>/dashboard
        ?date=Custom&start=<epoch_ms>&end=<epoch_ms>
        &URL=2;2;<ruta del inmueble>          (filtro "Ruta > Dirección URL > contiene")

Las fechas van en milisegundos "epoch", calculadas como si el día
empezara/terminara en hora de Bogotá (UTC-5, Colombia no tiene horario de
verano) — así es como las arma el propio selector de fechas de Clarity.

Se hacen 2 cargas de la página:
  1. Solo con fecha (sin filtro de URL): de ahí sale la imagen de resumen
     (Sesiones / Páginas por sesión / etc.) y la tarjeta "Páginas
     principales" tal como se ve siempre en el informe.
  2. Con fecha + filtro de Ruta (URL que contiene la del inmueble): de ahí
     se lee el número de "Sesiones" ya filtrado, que es la cifra que va en
     el texto "obtuvo N sesiones únicas a través de nuestro portal web".

Requisito: una ventana de Chrome YA ABIERTA con tu sesión de Clarity
iniciada, en modo de depuración remota (puerto 9222). Usa el acceso
directo "abrir_chrome_clarity.bat" antes de generar el informe.

Modo debug (para revisar capturas sin tocar ningún informe real):

    python utils/clarity_capture.py "11/05/2026" "29/05/2026" \\
        "https://activosporcolombia.com/es/unidad-inmobiliaria/20/apartamento-y-garaje-en-venta-cali-valle-del-cauca"

Guarda las capturas y el texto de la página en output/_debug_clarity/.
"""
from __future__ import annotations

import re
import sys
import time
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import quote

PROJECT_ID = "t4qu8mqi5b"
BASE_URL = f"https://clarity.microsoft.com/projects/view/{PROJECT_ID}/dashboard"
DEBUG_PORT = 9222

PROYECTO = Path(__file__).resolve().parent.parent
DEBUG_DIR = PROYECTO / "output" / "_debug_clarity"

_EPOCH = datetime(1970, 1, 1)
_OFFSET_BOGOTA = timedelta(hours=5)  # Bogotá es UTC-5 todo el año (sin horario de verano)


def _log(msg: str) -> None:
    print(f"  [clarity] {msg}", flush=True)


def _epoch_ms_bogota(fecha_ddmmyyyy: str, fin_de_dia: bool) -> int | None:
    """'dd/mm/YYYY' (hora de Bogotá) -> epoch en milisegundos, igual a como
    lo arma el selector de fechas del dashboard de Clarity."""
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", (fecha_ddmmyyyy or "").strip())
    if not m:
        return None
    dia, mes, anio = (int(x) for x in m.groups())
    try:
        if fin_de_dia:
            local = datetime(anio, mes, dia, 23, 59, 59, 999000)
        else:
            local = datetime(anio, mes, dia, 0, 0, 0)
    except ValueError:
        return None
    return int(((local + _OFFSET_BOGOTA) - _EPOCH).total_seconds() * 1000)


def _ruta_desde_url_pagina(url_pagina: str | None) -> str | None:
    """De 'https://activosporcolombia.com/es/unidad-inmobiliaria/20/xxx' saca
    'unidad-inmobiliaria/20/xxx' (lo que Clarity espera en el filtro de Ruta)."""
    if not url_pagina or url_pagina == "—":
        return None
    m = re.search(r"activosporcolombia\.com/es/(.+?)/?$", url_pagina.strip())
    if m:
        return m.group(1)
    return url_pagina.split("//")[-1].split("/", 1)[-1] or None


def _url_dashboard(fecha_ini: str, fecha_fin: str, ruta_filtro: str | None = None) -> str | None:
    start_ms = _epoch_ms_bogota(fecha_ini, fin_de_dia=False)
    end_ms = _epoch_ms_bogota(fecha_fin, fin_de_dia=True)
    if start_ms is None or end_ms is None:
        return None
    url = f"{BASE_URL}?date=Custom&start={start_ms}&end={end_ms}"
    if ruta_filtro:
        # "2;2;" = filtro "Dirección URL visitada" + operador "contiene"
        # (confirmado navegando Clarity: así arma el link el propio selector de Filtros > Ruta).
        url += f"&URL=2%3B2%3B{quote(ruta_filtro, safe='')}"
    return url


def conectar_chrome_debug(puerto: int = DEBUG_PORT):
    """Se engancha a una ventana de Chrome ya abierta en modo debug.
    NO abre una ventana nueva ni pide login: reutiliza tu sesión."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    # Primero se revisa si el puerto de depuración siquiera responde, para
    # poder distinguir "Chrome no está abierto en modo debug" (el .bat no se
    # corrió, o se cerró esa ventana) de otros errores menos obvios de
    # Selenium al engancharse — y para mostrar qué pestañas/ventanas ve.
    import json
    import urllib.request

    _log(f"Verificando Chrome en modo depuración (127.0.0.1:{puerto})...")
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{puerto}/json/version", timeout=3) as resp:
            info = json.loads(resp.read().decode("utf-8"))
        _log(f"  ✓ Chrome respondió: {info.get('Browser', '—')}")
        with urllib.request.urlopen(f"http://127.0.0.1:{puerto}/json/list", timeout=3) as resp:
            pestanas = json.loads(resp.read().decode("utf-8"))
        _log(f"  Pestañas abiertas ({len(pestanas)}):")
        for p in pestanas:
            _log(f"    - {p.get('title','—')!r} → {p.get('url','—')}")
    except Exception as e:
        _log(f"  ⚠ El puerto {puerto} no respondió: {e}")
        _log("  → Probablemente 'abrir_chrome_clarity.bat' no está abierto, o se cerró esa ventana.")

    opts = Options()
    opts.debugger_address = f"127.0.0.1:{puerto}"
    try:
        driver = webdriver.Chrome(options=opts)
        _log(f"  ✓ Selenium se enganchó a Chrome. URL actual: {driver.current_url}")
    except Exception as e:
        raise RuntimeError(
            "No se pudo conectar a Chrome en modo depuración (puerto "
            f"{puerto}). Abre 'abrir_chrome_clarity.bat' y deja la sesión de "
            "Clarity iniciada antes de generar el informe."
        ) from e

    # Tamaño de ventana fijo y grande: evita que las capturas salgan angostas
    # o recortadas si la ventana de Chrome del usuario quedó pequeña.
    try:
        driver.set_window_size(1600, 1000)
    except Exception:
        pass

    return driver


def _numero(txt: str) -> int | None:
    txt = (txt or "").strip().replace(".", "").replace(",", "")
    m = re.match(r"^(\d+)", txt)
    return int(m.group(1)) if m else None


def _leer_sesiones_y_bots(driver) -> tuple[int | None, int | None]:
    """Lee el número grande de 'Sesiones' y el de 'sesiones de bot excluidas'
    tal como se ven actualmente en pantalla (con el filtro que esté activo)."""
    from selenium.webdriver.common.by import By

    lineas = [l.strip() for l in driver.find_element(By.TAG_NAME, "body").text.split("\n")]
    for i, l in enumerate(lineas):
        if l == "Sesiones":
            # El ícono de ayuda (ⓘ) junto a "Sesiones" suele generar una línea
            # vacía antes del número real — se saltan las líneas en blanco.
            j = i + 1
            while j < len(lineas) and lineas[j] == "":
                j += 1
            sesiones = _numero(lineas[j]) if j < len(lineas) else None

            bots = None
            k = j + 1
            while k < len(lineas) and lineas[k] == "":
                k += 1
            if k < len(lineas):
                m = re.search(r"([\d.,]+)\s+sesiones? de bot", lineas[k], re.IGNORECASE)
                if m:
                    bots = _numero(m.group(1))
            return sesiones, bots
    return None, None


def _leer_sesiones_pagina_principal(driver, ruta):
    """Busca, dentro de la tarjeta 'Páginas principales', la fila cuya URL es
    exactamente la del inmueble y devuelve su número de sesiones.

    Esto es más preciso que el número del bloque 'Sesiones' (que usa el
    filtro de Ruta con 'contiene', y por eso puede sumar variantes parecidas
    de la misma URL) porque acá se compara la URL completa de esa fila
    puntual, tal como aparece en la tarjeta."""
    if not ruta:
        return None
    from selenium.webdriver.common.by import By

    lineas = [l.strip() for l in driver.find_element(By.TAG_NAME, "body").text.split("\n")]
    for i, l in enumerate(lineas):
        if l.startswith("http") and ruta in l:
            j = i + 1
            while j < len(lineas) and lineas[j] == "":
                j += 1
            if j < len(lineas):
                numero = _numero(lineas[j])
                if numero is not None:
                    return numero
    return None


def _asegurar_ventana_principal(driver, ventana_principal: str) -> None:
    """Si al interactuar con la página se abrió alguna pestaña/ventana nueva
    (p. ej. un enlace de ayuda o de 'compartir' de Clarity), la cierra y
    devuelve el foco a la ventana principal del dashboard. Sin esto, una
    captura de pantalla podría terminar tomándose sobre la pestaña nueva
    (normalmente mucho más angosta), produciendo una imagen rota."""
    try:
        handles = driver.window_handles
    except Exception:
        return
    if len(handles) > 1:
        for h in handles:
            if h != ventana_principal:
                try:
                    driver.switch_to.window(h)
                    driver.close()
                except Exception:
                    pass
    try:
        driver.switch_to.window(ventana_principal)
    except Exception:
        pass


def _screenshot_recorte(driver, y0: int, y1: int, ancho_minimo: int = 800):
    """Recorta la captura de pantalla actual entre las coordenadas y0..y1 (según el DOM),
    ajustando por el devicePixelRatio real de la ventana.

    Si la captura resulta sospechosamente angosta (por ejemplo porque el
    foco terminó en una pestaña/ventana distinta a la del dashboard), se
    reintenta una vez tras confirmar que estamos en la ventana correcta."""
    from PIL import Image

    def _capturar():
        png = driver.get_screenshot_as_png()
        img = Image.open(BytesIO(png))
        ancho_ventana = driver.execute_script("return window.innerWidth") or img.width
        escala = img.width / float(ancho_ventana)
        caja = (0, max(int(y0 * escala), 0), img.width, int(y1 * escala))
        return img, img.crop(caja)

    img, recorte = _capturar()
    if img.width < ancho_minimo:
        _log(
            f"⚠ La captura salió angosta ({img.width}px) — probablemente el foco "
            "quedó en una pestaña/ventana distinta. Reintentando..."
        )
        try:
            handles = driver.window_handles
            principal = handles[0] if handles else None
            if principal:
                _asegurar_ventana_principal(driver, principal)
        except Exception:
            pass
        time.sleep(0.5)
        img, recorte = _capturar()
        if img.width < ancho_minimo:
            _log(f"⚠ La captura sigue angosta ({img.width}px) tras reintentar.")
    return recorte


def capturar_resumen(driver) -> dict[str, Any]:
    """Bloque de arriba del dashboard (Filtros/fecha + 4 tarjetas + Información
    general/Ideas/Mi lista de reproducción), tal como se ve sin filtro de URL."""
    from selenium.webdriver.common.by import By

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)

    sesiones, bots = _leer_sesiones_y_bots(driver)

    # Límite inferior generoso por defecto: mejor que sobre un poco de fondo
    # blanco a que la foto quede corta y se pierda información (Ideas,
    # usuarios, etc.), que es justo lo que se pedía mostrar.
    y0, y1 = 0, 950
    try:
        filtros = driver.find_elements(By.XPATH, "//button[contains(text(), 'Filtros')]")
        if filtros:
            y0 = max(filtros[0].rect["y"] - 15, 0)
        y1 = y0 + 950

        # Se busca el título de la SIGUIENTE sección (más abajo de "Mi lista
        # de reproducción"/"Eventos inteligentes") para recortar justo ahí,
        # en vez de usar la altura del propio título "Mi lista de
        # reproducción" (que es solo el texto, muy angosto, y cortaba la
        # foto apenas empezando esa fila de tarjetas).
        for texto_siguiente in ("Embudos", "Origen de referencia", "Exploradores"):
            candidatos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{texto_siguiente}')]")
            for el in candidatos:
                try:
                    y = driver.execute_script("return arguments[0].getBoundingClientRect().y;", el)
                except Exception:
                    continue
                if y > y0 + 200:
                    y1 = max(y - 15, y0 + 500)
                    break
            else:
                continue
            break
    except Exception as e:
        _log(f"⚠ No se pudo calcular el recorte exacto del resumen, usando aproximado: {e}")

    imagen = _screenshot_recorte(driver, y0, y1)
    return {"imagen": imagen, "sesiones_totales": sesiones, "bots_excluidos": bots}


def capturar_tarjeta_paginas(driver):
    """Screenshot de la tarjeta 'Páginas principales' completa (vista sin
    filtrar), igual a como se ve siempre en el informe.

    Se recorta por coordenadas (como el resumen) en vez de usar
    element.screenshot_as_png, porque el contenedor real de la tarjeta queda
    varios niveles por encima del título y es difícil de ubicar de forma
    confiable — recortar entre "donde empieza este título" y "donde empieza
    el siguiente" es más robusto."""
    from selenium.webdriver.common.by import By

    titulos = driver.find_elements(By.XPATH, "//*[contains(text(), 'Páginas principales')]")
    if not titulos:
        _log("⚠ No se encontró la tarjeta 'Páginas principales'.")
        return None
    titulo = titulos[0]

    driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", titulo)
    time.sleep(0.6)

    rect_top = driver.execute_script(
        "return arguments[0].getBoundingClientRect().y;", titulo
    )
    y0 = max(rect_top - 10, 0)

    # Límite inferior por defecto (si no se encuentra la siguiente tarjeta).
    y1 = y0 + 480

    for texto_siguiente in ("Tráfico de bots", "Errores de JavaScript", "Embudos"):
        candidatos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{texto_siguiente}')]")
        for el in candidatos:
            try:
                y = driver.execute_script("return arguments[0].getBoundingClientRect().y;", el)
            except Exception:
                continue
            if y > rect_top + 30:
                y1 = max(y - 15, y0 + 150)
                break
        else:
            continue
        break

    return _screenshot_recorte(driver, y0, y1)


def capturar_datos_clarity(
    fecha_ini: str, fecha_fin: str, url_pagina: str, debug: bool = False
) -> dict[str, Any] | None:
    """Punto de entrada principal. Nunca lanza excepción hacia afuera: si algo
    falla devuelve None y el informe sigue generándose sin estos datos."""
    url_base = _url_dashboard(fecha_ini, fecha_fin)
    if not url_base:
        _log(f"⚠ Fechas inválidas para Clarity: {fecha_ini} / {fecha_fin}")
        return None

    try:
        driver = conectar_chrome_debug()
    except Exception as e:
        _log(str(e))
        return None

    ventana_principal = driver.current_window_handle

    try:
        _log(f"Abriendo dashboard (sin filtro de URL): {url_base}")
        driver.get(url_base)
        time.sleep(2.5)

        resumen = capturar_resumen(driver)
        _asegurar_ventana_principal(driver, ventana_principal)

        # La foto de "Páginas principales" se toma DESPUÉS de aplicar el
        # filtro de Ruta (Filtros > Ruta > URL del inmueble > Aplicar
        # filtros) — así lo confirmó el usuario: sin ese filtro la tarjeta
        # muestra las páginas más visitadas de TODO el sitio, no la del
        # inmueble. Por eso esta captura se hace en el mismo paso donde se
        # navega a la URL filtrada, no en el dashboard general.
        sesiones_url = None
        imagen_paginas = None
        ruta = _ruta_desde_url_pagina(url_pagina)
        if ruta:
            url_filtrada = _url_dashboard(fecha_ini, fecha_fin, ruta_filtro=ruta)
            _log(f"Abriendo dashboard filtrado por la URL del inmueble: {ruta}")
            driver.get(url_filtrada)
            time.sleep(2.5)
            sesiones_url, _ = _leer_sesiones_y_bots(driver)
            if sesiones_url is None:
                _log("⚠ No hubo sesiones para esa URL en el rango de fechas (o no se pudo leer).")
            imagen_paginas = capturar_tarjeta_paginas(driver)
            _asegurar_ventana_principal(driver, ventana_principal)

            # El número del bloque "Sesiones" (arriba) usa el filtro de Ruta
            # con "contiene", así que puede sumar variantes parecidas de la
            # misma URL. Si en la tarjeta "Páginas principales" aparece la
            # fila con la URL EXACTA del inmueble, ese número es más preciso
            # y es el mismo que se ve en la foto — se usa ese en su lugar.
            sesiones_url_exacta = _leer_sesiones_pagina_principal(driver, ruta)
            if sesiones_url_exacta is not None:
                if sesiones_url_exacta != sesiones_url:
                    _log(
                        f"  (bloque Sesiones filtrado: {sesiones_url} — "
                        f"fila exacta en Páginas principales: {sesiones_url_exacta}; "
                        "se usa el de la fila exacta)"
                    )
                sesiones_url = sesiones_url_exacta
        else:
            _log("⚠ No se pudo determinar la ruta del inmueble a partir de url_pagina.")

        if debug:
            DEBUG_DIR.mkdir(parents=True, exist_ok=True)
            if resumen.get("imagen"):
                resumen["imagen"].save(DEBUG_DIR / "resumen.png")
            if imagen_paginas:
                imagen_paginas.save(DEBUG_DIR / "paginas.png")
            (DEBUG_DIR / "pagina_completa.png").write_bytes(driver.get_screenshot_as_png())
            (DEBUG_DIR / "texto.txt").write_text(
                driver.find_element("tag name", "body").text, encoding="utf-8"
            )
            _log(f"Debug guardado en: {DEBUG_DIR}")

        return {
            "sesiones_totales": resumen.get("sesiones_totales"),
            "bots_excluidos": resumen.get("bots_excluidos"),
            "sesiones_url": sesiones_url,
            "imagen_resumen": resumen.get("imagen"),
            "imagen_paginas": imagen_paginas,
        }
    except Exception as e:
        _log(f"⚠ Error capturando datos de Clarity: {e}")
        import traceback

        traceback.print_exc()
        return None


def main():
    if len(sys.argv) < 4:
        print('Uso: python utils/clarity_capture.py "dd/mm/YYYY" "dd/mm/YYYY" "<url_pagina>"')
        sys.exit(1)
    fecha_ini, fecha_fin, url_pagina = sys.argv[1], sys.argv[2], sys.argv[3]
    resultado = capturar_datos_clarity(fecha_ini, fecha_fin, url_pagina, debug=True)
    if resultado:
        print("\n✓ Resultado:")
        print(f"  Sesiones totales:   {resultado['sesiones_totales']}")
        print(f"  Bots excluidos:     {resultado['bots_excluidos']}")
        print(f"  Sesiones de la URL: {resultado['sesiones_url']}")
    else:
        print("\n✗ No se pudo capturar información de Clarity.")


if __name__ == "__main__":
    main()
