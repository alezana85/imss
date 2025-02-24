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
        
        # Construir URL para descarga directa como CSV, especificando la hoja 'infonavit'
        url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=infonavit'
        df = pd.read_csv(url)
        return df
    except Exception as e:
        raise Exception(f"Error al acceder a Google Sheets: {str(e)}")
    
def intentar_generar_constancia(driver, nombre_corto, output_dir, max_intentos=10):
    """
    Intenta generar la constancia siguiendo el proceso específico:
    1. Verificar si existe botón de descarga
    2. Si no, click en botón width:211px
    3. Click en Generar constancia
    4. Click en advertencia
    5. Click en botón width:211px nuevamente
    6. Verificar botón de descarga
    7. En caso de no existir el boton de descarga repetir el proceso desde el paso 3 hasta el 6
    """
    for intento in range(max_intentos):
        print(f"Intento {intento + 1} de {max_intentos}")
        
        try:
            # Paso 1: Verificar si existe botón de descarga
            boton_descarga_existe = False
            try:
                descargar = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[alt="Descargar"]'))
                )
                print("Botón de descarga encontrado directamente!")
                return True
            except:
                print("Botón de descarga no encontrado, iniciando proceso de generación")
            
            # Paso 2: Click en botón width:211px
            sin_constancia = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[style="width: 211px;"]'))
            )
            sin_constancia.click()
            time.sleep(2)

            # Iniciar bucle interno para los pasos 3-6
            max_intentos_internos = 3
            for intento_interno in range(max_intentos_internos):
                try:
                    # Paso 3: Click en Generar constancia
                    generar = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="Generar constancia"]'))
                    )
                    generar.click()
                    time.sleep(2)

                    # Paso 4: Click en advertencia
                    advertencia = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/div/div[2]/button'))
                    )
                    advertencia.click()
                    time.sleep(2)

                    # Paso 5: Click nuevamente en botón width:211px
                    sin_constancia_2 = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, '[style="width: 211px;"]'))
                    )
                    sin_constancia_2.click()
                    time.sleep(2)

                    # Paso 6: Verificar botón de descarga
                    try:
                        descargar = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, '[alt="Descargar"]'))
                        )
                        print("Botón de descarga encontrado después del proceso!")
                        return True
                    except:
                        print(f"No se encontró el botón de descarga, reintento interno {intento_interno + 1}")
                        # Paso 7: Si no existe el botón, continuar con el siguiente intento interno
                        continue

                except Exception as e:
                    print(f"Error en intento interno {intento_interno + 1}: {str(e)}")
                    continue

        except Exception as e:
            print(f"Error en el intento principal {intento + 1}: {str(e)}")
            time.sleep(2)
            continue
    
    print(f"No se pudo generar la constancia después de {max_intentos} intentos")
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    screenshot_path = os.path.join(output_dir, f"{nombre_corto}_error_al_descargar_{timestamp}.png")
    driver.save_screenshot(screenshot_path)
    return False

def ingresar_al_portal(rp, mail, contrasena, certificado, llave, password, nombre_corto, output_dir):
    driver = None
    try:
        # Configurar opciones de Chrome
        chrome_options = Options()
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36')
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('window-size=1920,1080')
        chrome_options.add_argument('--force-device-scale-factor=0.5')
        
        # Configurar preferencias de descarga
        prefs = {
            "download.default_directory": output_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Inicializar el driver de Chrome
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        # Ingresar al portal del Infonavit
        driver.get("https://empresarios.infonavit.org.mx/")
        
        # Ingresar Número de Registro Patronal
        rp_input = driver.find_element(By.ID, 'nrp')
        rp_input.send_keys(rp)

        # Ingresar Correo Electrónico
        mail_input = driver.find_element(By.ID, 'correo')
        mail_input.send_keys(mail)

        # Ingresar Contraseña
        pass_input = driver.find_element(By.ID, 'pwd')
        pass_input.send_keys(contrasena)

        # Ingresar captcha (versión actualizada)
        captcha = driver.find_element(By.CSS_SELECTOR, '[class^="captcha-styles_component_textLine"]').text
        captcha_input = driver.find_element(By.CSS_SELECTOR, '[data-testid="captchaInput"]')
        captcha_input.send_keys(captcha)

        # Iniciar sesión
        iniciar_sesion = driver.find_element(By.CSS_SELECTOR, '[class="css-1qtwbji"]')
        iniciar_sesion.click()

        # Ingresar a Comprobante de Situación Fiscal
        situacion_fiscal = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[2]/div[4]/div[3]/a/p[1]')))
        situacion_fiscal.click()

        # Elegir certificado (versión actualizada)
        certificado_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]'))
        )
        certificado_input.send_keys(certificado)

        # Elegir llave privada (versión actualizada)
        llave_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]:last-of-type'))
        )
        llave_input.send_keys(llave)

        # Ingresar contraseña del certificado
        pass_certificado = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'plainPwd')))
        pass_certificado.send_keys(password)

        # Validar certificado
        validar_certificado = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[class="css-1qtwbji"]')))
        validar_certificado.click()

        # Aceptar terminos y condiciones
        terminos = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[role="checkbox"]')))
        terminos.click()

        # Continuar de los terminos
        continuar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[class="css-1qtwbji"]')))
        continuar.click()

        # Intentar generar la constancia
        if intentar_generar_constancia(driver, nombre_corto, output_dir):
            # Descargar constancia
            descargar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[alt="Descargar"]'))
            )
            descargar.click()
            
            # Esperar la descarga y renombrar el archivo
            time.sleep(10)
            archivos = [f for f in os.listdir(output_dir) if f.endswith('.pdf')]
            if archivos:
                ultimo_archivo = max([os.path.join(output_dir, f) for f in archivos], key=os.path.getctime)
                nuevo_nombre = os.path.join(output_dir, f'Mi_Opinion_INFONAVIT_{nombre_corto}.pdf')
                os.rename(ultimo_archivo, nuevo_nombre)
            else:
                raise Exception("No se encontró el archivo PDF descargado")
        else:
            raise Exception("No se pudo generar la constancia")

    except Exception as e:
        print(f"Error en el proceso para {nombre_corto}: {str(e)}")
        if driver:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            screenshot_path = os.path.join(output_dir, f"{nombre_corto}_error_proceso_{timestamp}.png")
            driver.save_screenshot(screenshot_path)
    finally:
        if driver:
            driver.quit()

def main():
    # Solicitar la URL de Google Sheets y el directorio de salida
    sheet_url = input("Ingrese la URL de Google Sheets: ")
    output_dir = input("Ingrese la ruta de la carpeta donde desea guardar los archivos: ")
    
    # Crear el directorio si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Obtener datos de Google Sheets
        df = get_sheet_data(sheet_url)
        
        # Filtrar solo las filas donde descargar_opinion es True
        df_filtrado = df[df['descargar_opinion'] == True]
        
        # Procesar cada fila filtrada
        for index, row in df_filtrado.iterrows():
            try:
                print(f"Procesando registro para: {row['nombre_corto']}")
                ingresar_al_portal(
                    rp=row['rp'],
                    mail=row['correo_electronico'],
                    contrasena=row['password'],
                    certificado=row['ruta_certificado'],
                    llave=row['ruta_llave'],
                    password=row['password_certificado'],
                    nombre_corto=row['nombre_corto'],
                    output_dir=output_dir
                )
            except Exception as e:
                print(f"Error procesando {row['nombre_corto']}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error al procesar la hoja de cálculo: {str(e)}")

if __name__ == "__main__":
    main()
