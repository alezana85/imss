import os
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

def extract_sheet_id(url):
    """Extrae el ID de la hoja de cálculo de la URL de Google Sheets."""
    if '/d/' in url:
        start = url.find('/d/') + 3
        end = url.find('/', start)
        if end == -1:
            end = url.find('?', start)
        if end == -1:
            end = len(url)
        return url[start:end]
    return None

def get_sheet_data(sheet_url):
    """Obtiene los datos de Google Sheets usando solo pandas."""
    try:
        sheet_id = extract_sheet_id(sheet_url)
        if not sheet_id:
            raise ValueError("No se pudo extraer el ID de la hoja de la URL proporcionada")
        
        # Construir URL para descarga directa como CSV, especificando la hoja 'empresa'
        url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=empresa'
        df = pd.read_csv(url)
        return df
    except Exception as e:
        raise Exception(f"Error al acceder a Google Sheets: {str(e)}")

def ingresar_al_buzon(rfc, certificado, key, contrasena, nombre_corto, ruta_descarga=None):
    # Si no se especifica ruta, pedir al usuario
    if ruta_descarga is None:
        ruta_descarga = input("Ingrese la ruta donde desea guardar el archivo: ")
    
    # Asegurar que la ruta existe
    ruta_descarga = os.path.abspath(ruta_descarga)
    os.makedirs(ruta_descarga, exist_ok=True)
    
    # Nombre del archivo
    nombre_archivo = f'Mi_Opinion_IMSS_{nombre_corto}.pdf'
    
    chrome_options = Options()
    # Configurar preferencias de descarga
    prefs = {
        "download.default_directory": ruta_descarga,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        # Add these to suppress console output
        "excludeSwitches": ['enable-logging', 'enable-automation'],
        "debuggerAddress": "",
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_argument('window-size=1920,1080')
    chrome_options.add_argument('--force-device-scale-factor=0.5')
    chrome_options.add_argument('--headless')
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Add these new options
    chrome_options.add_argument('--log-level=3')  # Only show fatal errors
    chrome_options.add_argument('--silent')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-software-rasterizer')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    service = Service(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        # Establecer timeout global para la página
        driver.set_page_load_timeout(10)
        driver.implicitly_wait(10)
        
        driver.get('https://buzon.imss.gob.mx/buzonimss/login')

        # Switchear al iFrame del login con timeout
        frame = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'formFirmaDigital'))
        )
        driver.switch_to.frame(frame)

        # ingresar el RFC con timeout
        input_rfc = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'inputRFC'))
        )
        input_rfc.send_keys(rfc)

        # Resto de los campos del formulario
        input_certificado = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'inputCertificado'))
        )
        input_certificado.send_keys(certificado)

        input_key = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'inputKey'))
        )
        input_key.send_keys(key)

        input_contrasena = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'inputPassword'))
        )
        input_contrasena.send_keys(contrasena)

        # validar certificado con timeout
        validar_certificado = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'botonValidarCert'))
        )
        validar_certificado.click()

        # salir del iFrame
        driver.switch_to.default_content()

        # Establecer un tiempo inicial para controlar el tiempo total del proceso
        start_time = time.time()

        # boton Cobranza con timeout
        boton_cobranza = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/main/div[1]/div/nav/div/div[2]/ul/li[11]/a'))
        )
        boton_cobranza.click()

        # boton 32D con timeout
        boton_32d = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/main/div[1]/div/nav/div/div[2]/ul/li[11]/ul/li[4]/a'))
        )
        boton_32d.click()

        # boton descargar con timeout
        descargar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/main/div[3]/div/form/div/div/div[2]/div/a/i'))
        )
        descargar.click()

        # Esperar la descarga del archivo con timeout mejorado
        tiempo_espera = 0
        archivo_encontrado = False
        nombre_archivo = f'Mi_Opinion_IMSS_{nombre_corto}.pdf'
        ruta_archivo_final = os.path.join(ruta_descarga, nombre_archivo)
        
        # Obtener lista de archivos antes de la descarga
        archivos_antes = set(os.listdir(ruta_descarga))
        
        while tiempo_espera < 30 and not archivo_encontrado:
            time.sleep(1)
            archivos_actuales = set(os.listdir(ruta_descarga))
            nuevos_archivos = archivos_actuales - archivos_antes
            
            # Buscar archivos temporales de descarga
            archivos_descarga = [f for f in nuevos_archivos if f.endswith('.pdf') or f.endswith('.crdownload')]
            
            if archivos_descarga:
                for archivo in archivos_descarga:
                    ruta_archivo = os.path.join(ruta_descarga, archivo)
                    if archivo.endswith('.pdf') and not archivo.endswith('.crdownload'):
                        try:
                            # Intentar abrir el archivo para verificar que la descarga está completa
                            with open(ruta_archivo, 'rb') as f:
                                pass
                            # Si llegamos aquí, el archivo está completo
                            if os.path.exists(ruta_archivo_final):
                                os.remove(ruta_archivo_final)  # Eliminar archivo existente si es necesario
                            os.rename(ruta_archivo, ruta_archivo_final)
                            archivo_encontrado = True
                            print(f"Archivo guardado como: {nombre_archivo}")
                            break
                        except (PermissionError, OSError):
                            continue
            
            tiempo_espera += 1
            
            # Verificar tiempo total
            if time.time() - start_time > 120:
                raise TimeoutError("El proceso excedió el tiempo límite de 120 segundos")

        if not archivo_encontrado:
            raise Exception("No se encontró el archivo PDF descargado o la descarga no se completó")

    except Exception as e:
        # Tomar screenshot en caso de error
        screenshot_name = f"{nombre_corto}_error_al_descargar.png"
        screenshot_path = os.path.join(ruta_descarga, screenshot_name)
        driver.save_screenshot(screenshot_path)
        print(f"Se ha guardado una captura de pantalla del error en: {screenshot_name}")
        raise Exception(f"Error durante la descarga: {str(e)}")
    finally:
        # cerrar el navegador
        driver.quit()


if __name__ == '__main__':
    # Solicitar URL del Google Sheet
    sheet_url = input("Ingrese la URL del Google Sheet: ")
    ruta_descarga = input("Ingrese la ruta donde desea guardar los archivos: ")
    
    # Obtener datos del Google Sheet
    df = get_sheet_data(sheet_url)
    
    # Filtrar solo las filas donde descargar_opinion es TRUE
    df_filtrado = df[df['descargar_opinion'] == True]
    
    # Iterar sobre las filas filtradas
    for index, row in df_filtrado.iterrows():
        print(f"\nProcesando empresa: {row['empresa']}")
        try:
            ingresar_al_buzon(
                rfc=row['rfc'],
                certificado=row['ruta_certificado'],
                key=row['ruta_llave'],
                contrasena=row['password'],
                nombre_corto=row['nombre_corto'],
                ruta_descarga=ruta_descarga
            )
            print(f"Proceso completado para {row['empresa']}")
        except Exception as e:
            print(f"Error procesando {row['empresa']}: {str(e)}")
        
        # Esperar un tiempo aleatorio entre procesos para evitar sobrecarga
        time.sleep(random.uniform(2, 5))