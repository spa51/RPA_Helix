from robocorp.tasks import task
from robocorp import browser
import os
from db_utils import conectar_bd
from helix_actions import validar_items
from utils import generar_codigo_totp

@task
def login_smartit():
    print("--- Inicio de Sesión Smart IT ---")
    # Credenciales por defecto
    email = "spinerez@bancolombia.com.co"
    password = "D4taf1l3$M3d27#"

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
    
    # Generamos el código automáticamente
    auth_code = generar_codigo_totp()
    print(f"Código TOTP generado: {auth_code}")
    
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
    
    # 5. Esperar a que cargue la consola de Smart IT
    print("Esperando a que cargue completamente la consola de Smart IT...")
    page.wait_for_timeout(10000) 
    
    # 5.5 Conectar a la base de datos Oracle
    conn = conectar_bd()
    
    # 6. Validar existencia de ítems
    validar_items(page, conn)
    
    # Cerrar conexion a BD al finalizar
    if conn:
        try:
            conn.close()
            print("Conexión a BD finalizada.")
        except Exception:
            pass
    
    # Tomar captura de pantalla
    os.makedirs("output/img", exist_ok=True)
    page.screenshot(path="output/img/smartit_login.png")
    print("Captura de pantalla guardada en 'output/img/smartit_login.png'")
    print("--- Fin de la prueba de RPA ---")
