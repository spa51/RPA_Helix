from robocorp.tasks import task
from robocorp import browser
import os
import tkinter as tk
from tkinter import simpledialog
import oracledb

def conectar_bd():
    print("\n=== Conectando a Base de Datos Oracle ===")
    try:
        dsn = oracledb.makedsn("172.16.14.17", 1521, service_name="ORCL")
        conn = oracledb.connect(
            user="datasoft",
            password="data2001",
            dsn=dsn
        )
        print("Conexión exitosa a la base de datos Oracle.")
        return conn
    except Exception as e:
        print(f"Error al conectar a la base de datos: {e}")
        return None

def get_input_popup(prompt_text, is_password=False):
    root = tk.Tk()
    root.withdraw() # Ocultar la ventana principal
    root.attributes('-topmost', 1) # Asegurar que el popup salga al frente
    if is_password:
        result = simpledialog.askstring("Ingreso requerido", prompt_text, parent=root, show='*')
    else:
        result = simpledialog.askstring("Ingreso requerido", prompt_text, parent=root)
    root.destroy()
    return result

def validar_items(page):
    print("=== Validando existencia de items en la tabla ===")
    try:
        # Esperamos a que aparezca la tabla vacía o la tabla con datos
        page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
        
        # Validamos cuál elemento cargó
        if page.locator(".tc__list-placeholder-text").is_visible():
            print("no hay items")
        elif page.locator(".ngViewport").is_visible():
            contar_items(page)
        else:
            # Por si algo raro pasó con el DOM
            print("no se pudo determinar el estado de la tabla")
            
    except Exception as e:
        print(f"Aviso: Timeout al esperar la tabla. Detalle: {str(e)[:50]}")

def contar_items(page):
    # Contar items iterando por fila
    count = 0
    while True:
        selector = f".ng-scope:nth-child({count + 1}) > .col2 .ngCellText"
        if page.locator(selector).count() > 0:
            count += 1
        else:
            break
    print(f"hay items: {count} item(s) encontrado(s)")

    # Obtener detalle de todos los ítems encontrados
    for i in range(1, count + 1):
        obtener_detalle_item(page, item_index=i)
        
        # Si no es el último ítem, volver a la consola de tickets para poder abrir el siguiente
        if i < count:
            print("Regresando a la consola de tickets...")
            page.go_back()
            # Esperar a que recargue la tabla original antes de buscar el siguiente
            page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
            page.wait_for_timeout(2000) # Pausa mínima para estabilizar la tabla


def obtener_detalle_item(page, item_index=1):
    """
    Hace clic en el ítem de la lista, entra al iframe (#pwa-frame) 
    y extrae el valor específico de 'No. Cédula:'.
    """
    print(f"\n=== Extrayendo detalle del ítem #{item_index} ===")
    try:
        # 1. Hacer clic en la fila para entrar al ítem
        fila_selector = f".ng-scope:nth-child({item_index}) > .col2 .ngCellText"
        page.locator(fila_selector).first.click()
        print("Clic realizado, entrando a la vista del ticket...")

        # 2. Entrar al iframe que contiene la vista progresiva de Helix
        # Hemos quitado las pausas largas! Dejamos que Playwright detecte el iframe rápido.
        frame = page.frame_locator("#pwa-frame")
        
        # 3. Esperar y ubicar el cuadro de texto exacto
        cuadro_datos = frame.locator("#ar1000000151_data")
        # El programa avanzará en el instante microscópico en que detecte que el texto existe
        cuadro_datos.first.wait_for(state="attached", timeout=20000)
        
        # Extraemos el texto completo (del atributo title que contiene la data estructurada)
        texto = cuadro_datos.first.get_attribute("title")
        if not texto:  # Respaldo por si acaso
            texto = cuadro_datos.first.inner_text()

        # 4. Aislar todos los campos solicitados
        campos_buscados = [
            "No. Cédula:",
            "Tipo solicitud:",
            "Usuario de Datasoft:",
            "Información del servicio:",
            "Usuario de referencia DATASOFT:",
            "Línea de Negocio:"
        ]
        
        datos_extraidos = {}
        
        for linea in texto.splitlines():
            for campo in campos_buscados:
                if campo in linea:
                    partes = linea.split(campo)
                    if len(partes) > 1:
                        # Guardamos el valor limpio sin los dos puntos
                        nombre_amigable = campo.replace(":", "").strip()
                        datos_extraidos[nombre_amigable] = partes[1].strip()
                        break

        # Extraer Nombre Completo
        try:
            nombre_loc = frame.locator("#ar301395400_data")
            if nombre_loc.count() == 0:
                nombre_loc = page.locator("#ar301395400_data")
            if nombre_loc.count() > 0:
                datos_extraidos["Nombre Completo"] = nombre_loc.first.inner_text().strip()
        except Exception as e:
            print(f"Aviso: No se pudo extraer Nombre Completo {e}")

        # Extraer Correo
        try:
            correo_loc = frame.locator("#ar1000000048_data")
            if correo_loc.count() == 0:
                correo_loc = page.locator("#ar1000000048_data")
            if correo_loc.count() > 0:
                datos_extraidos["Correo"] = correo_loc.first.inner_text().strip()
        except Exception as e:
            print(f"Aviso: No se pudo extraer Correo {e}")

        # 5. Imprimir el resultado sin caracteres especiales/emojis
        print("\n=================================")
        if datos_extraidos:
            print("DATOS EXTRAIDOS DEL ITEM:")
            # Se imprime sin acentos extraños en los prints para cuidar que la consola de Windows no falle
            for clave, valor in datos_extraidos.items():
                # Reemplazamos acentos en las llaves solo para la impresion segura en CMD
                clave_segura = clave.replace('é', 'e').replace('í', 'i')
                print(f"  {clave_segura}: {valor}")
        else:
            print("Fallo: No se vio ningun texto de los solicitados.")
            print("Texto extraido para depurar:\n", texto)
        print("=================================\n")

    except Exception as e:
        print(f"Error al extraer los datos: {str(e)[:150]}")
    

@task
def login_smartit():
    print("--- Inicio de Sesión Smart IT ---")
    # Credenciales por defecto
    email = "spinerez@bancolombia.com.co"
    password = "D4taf1l3$M3d26#"

    # Configurar el navegador
    browser.configure(
        browser_engine="chromium",
        headless=True, # Necesitamos ver el navegador
        isolated=False,  # Modo Incógnito / Contexto limpio
    )
    
    print("Abriendo el navegador...")
    page = browser.goto("https://bancolombia-smartit.onbmc.com/smartit/app/#/ticket-console")
    
    # 1. Ingresar Correo
    print("Esperando y llenando campo de correo (las redirecciones SSO pueden tomar algo de tiempo)...")
    page.wait_for_selector("#i0116", state="visible", timeout=60000)
    page.locator("#i0116").fill(email)
    
    # Clic en "Siguiente"
    page.locator("#idSIButton9").click()
    
    # 2. Ingresar Contraseña
    print("Esperando y llenando contraseña...")
    page.wait_for_selector("#i0118", state="visible", timeout=30000)
    page.locator("#i0118").fill(password)
    
    # Clic en "Iniciar sesión" / Siguiente
    page.locator("#idSIButton9").click()
    
    # 3. Código del autenticador
    print("\nSe ha enviado la contraseña...")
    
    # Pedimos el código con popup mientras carga la página
    auth_code = get_input_popup("Por favor ingrese el código del autenticador que recibió:")
    
    if auth_code:
        print("Ingresando el código de autenticación...")
        try:
            page.wait_for_selector("#idTxtBx_SAOTCC_OTC", state="visible", timeout=30000)
            page.locator("#idTxtBx_SAOTCC_OTC").fill(auth_code)
            
            # Clic en "Confirmar" / "Verificar"
            page.locator("#idSubmit_SAOTCC_Continue").click()
        except Exception as e:
            print(f"Aviso: Ocurrió un error al intentar ingresar el código. Detalle: {str(e)[:50]}")
    
    # 4. Seleccionar cuenta si aparece el picker de cuentas de Microsoft (Falta verificar aveces aparece esta  raro)
    try:
        account_tile_selector = f"div[data-test-id='{email}']"
        page.wait_for_selector(account_tile_selector, state="visible", timeout=5000)
        print("Selector de cuenta detectado, haciendo click...")
        page.locator(account_tile_selector).click()
    except Exception:
        pass  # No apareció el selector de cuenta, se continúa normalmente

    # 5. Confirmar "¿Mantener la sesión iniciada?"
    try:
        page.wait_for_selector("#idBtn_Back", state="visible", timeout=5000)
        page.locator("#idBtn_Back").click()
    except Exception:
        pass 
    
    # 5. Esperar a que cargue la consola de Smart IT
    print("Esperando a que cargue completamente la consola de Smart IT...")
    page.wait_for_timeout(10000) 
    
    # 6. Validar existencia de ítems
    validar_items(page)
    
    # Tomar captura de pantalla
    os.makedirs("output/img", exist_ok=True)
    page.screenshot(path="output/img/smartit_login.png")
    print("Captura de pantalla guardada en 'output/img/smartit_login.png'")
    print("--- Fin de la prueba de RPA ---")

