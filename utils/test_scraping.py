"""
Prueba de scraping de fecha de publicación desde activosporcolombia.com
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

url = "https://activosporcolombia.com/es/unidad-inmobiliaria/399/finca-en-venta-la-dorada-caldas"

options = Options()
options.add_argument("--headless")          # Sin abrir ventana
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

print("Iniciando Chrome...")
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

try:
    print(f"Cargando: {url}")
    driver.get(url)
    time.sleep(3)

    # Mostrar todo el texto de la página para encontrar la fecha
    texto = driver.find_element(By.TAG_NAME, "body").text
    lineas = texto.split("\n")
    
    print("\nBuscando fechas en la página...\n")
    for i, linea in enumerate(lineas):
        if any(x in linea.lower() for x in ["public", "apertur", "cierr", "subasta", "may", "jun", "jul"]):
            print(f"[{i}] {linea.strip()}")

finally:
    driver.quit()
    print("\nChrome cerrado.")
