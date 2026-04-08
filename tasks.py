from robocorp.tasks import task
from robocorp import browser
import tkinter as tk
from tkinter import simpledialog

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

@task
def login_smartit():
    print("--- Inicio de Sesión Smart IT ---")
    # Usamos popups porque la extensión de Sema4ai/Robocorp bloquea la consola (EOFError)
    email = get_input_popup("Por favor ingrese su correo electrónico:")
    password = get_input_popup("Por favor ingrese su contraseña:", is_password=True)

    if not email or not password:
        print("Credenciales canceladas por el usuario.")
        return

    # Configurar el navegador
    browser.configure(
        browser_engine="chromium",
        headless=False, # Necesitamos ver el navegador
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
    
    # 4. Confirmar "¿Mantener la sesión iniciada?"
    try:
        page.wait_for_selector("#idBtn_Back", state="visible", timeout=5000)
        page.locator("#idBtn_Back").click()
    except Exception:
        pass 
    
    # 5. Esperar a que cargue la consola de Smart IT
    print("Esperando a que cargue completamente la consola de Smart IT...")
    page.wait_for_timeout(10000) 
    
    # Tomar captura de pantalla
    page.screenshot(path="output/smartit_login.png")
    print("Captura de pantalla guardada en 'output/smartit_login.png'")
    print("--- Fin de la prueba de RPA ---")

