"""
Mapea todo el contenido de la página del inmueble en activosporcolombia.com
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

grupo_id = input("Ingresa el grupo_id del inmueble (ej: 399): ").strip()
url = f"https://activosporcolombia.com/es/unidad-inmobiliaria/{grupo_id}/"

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

print(f"\nCargando: {url}")
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

try:
    driver.get(url)
    time.sleep(3)
    lineas = driver.find_element(By.TAG_NAME, "body").text.split("\n")

    print(f"\n{'='*60}")
    print(f"  CONTENIDO COMPLETO DE LA PÁGINA ({len(lineas)} líneas)")
    print(f"{'='*60}\n")

    for i, linea in enumerate(lineas):
        linea = linea.strip()
        if linea:
            print(f"[{i:03d}] {linea}")

finally:
    driver.quit()
    print("\nChrome cerrado.")
