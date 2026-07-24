from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time, re

def slugify(t):
    t = t.lower()
    for k,v in {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}.items(): t=t.replace(k,v)
    t = re.sub(r"[\s\-]+", "-", t)
    t = re.sub(r"[^a-z0-9\-]", "", t)
    return t.strip("-")

MESES_ABR = {
    "ene":"01","feb":"02","mar":"03","abr":"04","may":"05","jun":"06",
    "jul":"07","ago":"08","sep":"09","oct":"10","nov":"11","dic":"12"
}

grupo_id = "399"
nombre = "Finca en venta - La Dorada, Caldas"
url = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/{slugify(nombre)}"

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
try:
    driver.get(url)
    time.sleep(3)
    lineas = [l.strip() for l in driver.find_element(By.TAG_NAME, "body").text.split("\n")]

    anio = "2026"
    for l in lineas:
        m = re.search(r"20\d{2}", l)
        if m:
            anio = m.group()
            break

    fases = [
        ("Publicación próxima en subasta", "f1", "f2"),
        ("Registro",                       "f3", "f4"),
        ("Análisis debida diligencia",     "f5", "f6"),
        ("Análisis financiero",            "f7", "f8"),
        ("Expedición y envío de cupones",  "f9", "f10"),
        ("Link pago seriedad",             "f11", "f12"),
        ("Validación y confirmación",      "f13", "f14"),
        ("Subasta (apertura y cierre)",    "f15", "f16"),
    ]

    for fase_nombre, ph_ini, ph_fin in fases:
        encontrado = False
        for i, linea in enumerate(lineas):
            if fase_nombre.lower() in linea.lower():
                encontrado = True
                print(f"\n✓ '{fase_nombre}' en línea [{i}]: '{linea}'")
                # Buscar rango
                for offset in [1, 2, 3]:
                    candidato = lineas[i+offset] if i+offset < len(lineas) else ""
                    print(f"  [{i+offset}] '{candidato}' — empieza con digito: {bool(candidato and candidato[0].isdigit())}")
                    if candidato and candidato[0].isdigit():
                        mes_l = lineas[i+offset+1] if i+offset+1 < len(lineas) else ""
                        mes = MESES_ABR.get(mes_l.split(".")[0].strip(), "00")
                        if "→" in candidato:
                            partes = candidato.split()
                            dia_ini = partes[0].zfill(2)
                            dia_fin = partes[2].zfill(2)
                        else:
                            dia_ini = candidato.split()[0].zfill(2)
                            dia_fin = dia_ini
                        print(f"  → RESULTADO: {dia_ini}/{mes}/{anio} — {dia_fin}/{mes}/{anio}")
                        break
                break
        if not encontrado:
            print(f"\n✗ '{fase_nombre}' NO encontrado")
finally:
    driver.quit()
