from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from PIL import Image
from io import BytesIO
import time
import datetime

# Configuración de Chrome
chrome_options = Options()
chrome_options.add_argument('--headless')  # Ejecutar Chrome en modo headless
chrome_options.add_argument('--disable-gpu')

service = Service('ruta_al_driver_de_chrome')  # Ruta al driver de Chrome
driver = webdriver.Chrome(service=service, options=chrome_options)

# Lista de RUCs
rucs = ['20603449143']  # Ejemplo, añade todos los RUCs necesarios

for ruc in rucs:
    driver.get('https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias')

    # Ingresar el RUC
    ruc_input = driver.find_element(By.ID, 'txtRuc')
    ruc_input.send_keys(ruc)
    ruc_input.send_keys(Keys.RETURN)

    time.sleep(3)  # Esperar a que la página cargue completamente

    # Captura de pantalla
    fecha = datetime.datetime.now().strftime('%Y%m%d')
    screenshot = driver.get_screenshot_as_png()
    image = Image.open(BytesIO(screenshot))
    image.save(f'{ruc}_{fecha}.png')

    # Guardar página como PDF (Ctrl+P)
    driver.execute_script('window.print();')
    time.sleep(3)  # Esperar para asegurarse que el diálogo de impresión se haya abierto

    # Aquí necesitarías un paso adicional para automatizar el guardado como PDF en la ubicación deseada,
    # lo cual puede implicar configuraciones avanzadas o el uso de herramientas adicionales.

driver.quit()
