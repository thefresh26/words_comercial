from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

grupo_id = input("grupo_id (ej: 399): ").strip()
nombre = input("nombre_grupo (ej: Finca en venta - La Dorada, Caldas): ").strip()

import re
def slugify(t):
    t = t.lower()
    for k,v in {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}.items(): t=t.replace(k,v)
    t = re.sub(r"[\s\-]+", "-", t)
    t = re.sub(r"[^a-z0-9\-]", "", t)
    return t.strip("-")

url = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/{slugify(nombre)}"
print(f"\nURL: {url}\n")

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

    fases = ["Publicación","Registro","Análisis","financiero","Expedición","Link","Validación","Subasta"]
    
    for i, l in enumerate(lineas):
        if any(f.lower() in l.lower() for f in fases) or "→" in l or ("may" in l.lower() and "día" in l.lower()):
            print(f"[{i:03d}] {l}")
finally:
    driver.quit()
    print("\nListo.")
