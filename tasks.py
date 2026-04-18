from robocorp.tasks import task
from robocorp import browser
import os
import tkinter as tk
from tkinter import simpledialog
import oracledb
import subprocess
import tempfile

def conectar_bd():
    print("\n=== Conectando a Base de Datos Oracle (Modo Grueso) ===")
    try:
        oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_0")
        conn = oracledb.connect(
            user="datasoft",
            password="data2001",
            dsn="172.16.14.17:1521/ORCL"
        )
        print("Conexión exitosa a la BD.")
        return conn
    except Exception as e:
        print(f"Error al conectar a BD: {e}")
        return None

def ejecutar_consulta_db(query, conn):
    if not conn:
        print("Aviso: No hay conexión a BD para ejecutar la consulta.")
        return "0"
    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            res = cursor.fetchone()
            if res:
                return str(res[0])
            return "0"
    except Exception as e:
        print(f"Error ejecutando consulta en BD: {e}")
        return "0"

def ejecutar_actualizacion_db(query, conn):
    if not conn:
        print("Aviso: No hay conexión a BD para ejecutar la actualización.")
        return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            conn.commit()
            return True
    except Exception as e:
        print(f"Error ejecutando actualización en BD: {e}")
        return False

def generar_informe_glpi(page, datos_extraidos=None, login_generado=None, password_generado=None):
    print("\n=== Iniciando creación de ticket en GLPI ===")
    try:
        context = page.context
        glpi_page = context.new_page()
        
        glpi_page.goto("http://172.0.0.60/Mesa_Servicios_Tecnologicos/index.php")
        
        # Comprobar si pide login o si ya hay sesión guardada en memoria
        try:
            # Esperamos 3 segundos a ver si aparece el campo de usuario
            glpi_page.locator("#login_name").wait_for(state="visible", timeout=3000)
            necesita_login = True
        except:
            necesita_login = False

        if necesita_login:
            print("Ingresando credenciales en GLPI...")
            glpi_page.locator("#login_name").fill("santiago.pinerez")
            # Selector genérico para evitar fallos por names dinámicos
            glpi_page.locator("input[type='password']").fill("Spadata55/")
            glpi_page.locator("[name=submit]").click()
            # Espera breve para asegurar que el inicio de sesión termine
            glpi_page.wait_for_timeout(2000)
        else:
            print("Sesión de GLPI detectada en caché. Omitiendo login...")
        
        print("Navegando al menú Crear caso...")
        glpi_page.locator("a", has_text="Soporte").click()
        glpi_page.locator("a", has_text="Crear caso").click()
        
        print("Esperando la carga del formulario de ticket...")
        # 1. Seleccionar Categoría (id dinámico)
        categoria_selector = "[id^='select2-dropdown_itilcategories_id']"
        glpi_page.locator(categoria_selector).wait_for(state="visible", timeout=15000)
        glpi_page.locator(categoria_selector).click()
        
        search_field = glpi_page.locator(".select2-search--dropdown > .select2-search__field")
        search_field.wait_for(state="visible", timeout=5000)
        search_field.fill("Aplicaciones > Administración de usuarios Bancolombia Banco")
        
        glpi_page.locator("li.select2-results__option", has_text="Aplicaciones > Administración de usuarios Bancolombia Banco").first.click()
        
        # 2. Asignado a
        print("Asignando usuario: Jose Luis Echavarria Ochoa...")
        
        # Esperamos a que apareza al menos un input de select2
        glpi_page.locator("input.select2-search__field").first.wait_for(state="visible", timeout=10000)
        
        # En el formulario estándar de GLPI hay 3 campos de actores: Solicitante, Observador y Asignado.
        # Seleccionamos explícitamente el tercero.
        inputs = glpi_page.locator("input.select2-search__field")
        count = inputs.count()
        
        if count >= 3:
            search_assign = inputs.nth(2)
            print("Campo de asignación identificado (Tercer select2).")
        else:
            search_assign = inputs.last
            print(f"Fallback: Utilizando el último campo select2 de {count} disponibles.")
            
        search_assign.scroll_into_view_if_needed()
        search_assign.click()
        search_assign.fill("Jose")
        
        # Esperamos que cargue la lista flotante y damos clic
        glpi_page.wait_for_timeout(2000) # dar un instante para que reaccione GLPI
        opcion_busqueda = glpi_page.locator("li.select2-results__option", has_text="Jose Luis Echavarria Ochoa").first
        opcion_busqueda.wait_for(state="visible", timeout=10000)
        opcion_busqueda.click()
        
        # 3. Título (id dinámico)
        print("Ingresando título...")
        titulo_input = glpi_page.locator("input[id^='name_']")
        titulo_input.fill("ACTIVACION USUARIO BANCO")
        
        # 4. Descripción (dentro del iframe de TinyMCE)
        print("Ingresando descripción...")
        iframe_locator = glpi_page.frame_locator("iframe.tox-edit-area__iframe")
        body_locator = iframe_locator.locator("body#tinymce")
        body_locator.wait_for(state="visible", timeout=10000)
        
        # Construir la descripción de manera dinámica
        descripcion_glpi = "Buen día,\nMe colaboran gestionando este ticket del Helix para activación de usuario de Bancolombia.\nadjunto evidencia:\n\n"
        
        if datos_extraidos:
            descripcion_glpi += "--- Datos del usuario de Helix ---\n"
            for clave, valor in datos_extraidos.items():
                descripcion_glpi += f"{clave} {valor}\n"
                
        if login_generado and password_generado:
            descripcion_glpi += f"\n--- Usuario generado ---\nUsuario: {login_generado}\nContraseña: {password_generado}\n"
        
        body_locator.click()
        body_locator.fill(descripcion_glpi)
        # Presionamos un espacio y lo borramos para forzar a TinyMCE a registrar el evento de teclado
        # Esto elimina el estado nativo de error 'required' (rojo inactivo)
        body_locator.press("Space")
        body_locator.press("Backspace")
        
        # 5. Guardar (Agregar)
        print("Guardando caso en GLPI...")
        glpi_page.locator("[name=add]").click()
        
        # Dar tiempo al sistema para crear
        glpi_page.wait_for_timeout(3000)
        print("Caso creado en GLPI exitosamente.")
        
        # Se cierra la pestaña para no estorbar a Helix
        glpi_page.close()
    except Exception as e:
        print(f"Aviso: Error al intentar generar informe en GLPI: {e}")
        try:
            glpi_page.close() # Limpieza en caso de error
        except:
            pass

def validar_datos_condicionales(datos_extraidos, conn, page=None, frame=None):
    print("\n--- Validando en Base de Datos vía Terminal ---")

    info_servicio = datos_extraidos.get("Información del servicio", "").lower()
    tipo_solicitud = datos_extraidos.get("Tipo solicitud", "")
    usuario_datasoft = datos_extraidos.get("Usuario de Datasoft", "")
    cedula = datos_extraidos.get("No. Cédula", "")

    if "masiva" in info_servicio:
        print("(Solicitud masiva omitida)")
        return

    try:
        if tipo_solicitud == "Activación/desbloqueo":
            print(f"Validando existencia de '{usuario_datasoft}' o '{cedula}'...")
            existe = False
            login_final = usuario_datasoft
            
            if usuario_datasoft:
                res = ejecutar_consulta_db(f"SELECT COUNT(*) FROM a_usuario WHERE LOGIN = '{usuario_datasoft}'", conn)
                if res and res.strip() != "0":
                    existe = True
            
            if not existe and cedula:
                res_login = ejecutar_consulta_db(f"SELECT LOGIN FROM a_usuario WHERE NO_IDENTIFICACION = '{cedula}'", conn)
                if res_login and res_login.strip() != "0":
                    existe = True
                    login_final = res_login
                    print(f"Se comprobó existencia por cédula. Usando LOGIN de la base: '{login_final}'")
                    
            if existe:
                print("existe")
                if login_final:
                    print(f"Actualizando usuario {login_final}...")
                    update_query = f"UPDATE a_usuario SET CONTRASENA = NO_IDENTIFICACION, ESTADO = 'A', CAMBIO_CONTRASENA = 'S' WHERE LOGIN = '{login_final}'"
                    exito = ejecutar_actualizacion_db(update_query, conn)
                    if exito:
                        print("Actualización exitosa. Ingresando nota al ticket...")
                        mensaje = f"""POR FAVOR LEER MUY DESPACIO Y SEGUIR EL  PASO A PASO CON LAS INSTRUCCIONES QUE SE DETALLAN A CONTINUACIÓN

Buen día.

Tu usuario fue activado;
Usuario: {login_final}
Contraseña: {cedula}

Debes realizar los siguientes pasos para restablecer la contraseña o cambiarla:

Digita en la página inicial de Datasoft http://172.17.21.14/Datasoft/login.php los campos: Usuario, 
compañía, sin darle Clave y código de seguridad y luego das clic en Restablecer Contraseña.
IMPORTANTE: Se debe ingresar al sistema antes de 24 horas para Activar el Usuario y evitar que se vuelva a 
inactivar. Tener en cuenta al cambiar la contraseña:
• Longitud mínima 8 caracteres
• Longitud máxima 12 caracteres
• Debe contener como mínimo (2) dos letras
• Debe contener como mínimo (2) dos números
• Debe contener como mínimo (2) dos caracteres especiales de la siguiente lista: $ - _ =
• No se pueden repetir contraseñas anteriores.
• No usar el asterisco * (asterisco)

Saludos,"""
                        try:
                            # Intentar escribir la nota en el textarea dentro del frame o la página
                            textarea_selector = 'textarea[data-testid="304247080"], textarea[name="ar304247080"]'
                            loc = frame.locator(textarea_selector).first if frame else page.locator(textarea_selector).first
                            loc.wait_for(state="visible", timeout=10000)
                            loc.fill(mensaje)
                            
                            # Clic en el botón "Publicación"
                            btn_selector = 'button[name="ar304268430"], button[id="304268430"], button[title="Publicación"]'
                            btn_loc = frame.locator(btn_selector).first if frame else page.locator(btn_selector).first
                            btn_loc.wait_for(state="visible", timeout=5000)
                            btn_loc.click()
                            print("Nota ingresada y publicada correctamente en el ticket.")
                            
                            # Clic en "Asignarme a mí"
                            print("Asignando el ticket a mí mismo...")
                            page.wait_for_timeout(2000) # Esperar a que consolide la nota
                            btn_asignar_sel = 'button[name="ar304421551"], button[id="304421551"], button[title="Asignarme a mí"]'
                            btn_asignar_loc = frame.locator(btn_asignar_sel).first if frame else page.locator(btn_asignar_sel).first
                            btn_asignar_loc.wait_for(state="visible", timeout=10000)
                            btn_asignar_loc.click()
                            print("Ticket asignado correctamente.")
                            
                            # 1. Clic en "Editar"
                            print("Cambiando el estado a Finalizado...")
                            page.wait_for_timeout(2000) # Esperar a que consolide la asignación
                            btn_editar_sel = 'button[name="ar304420591"], button[title="Editar"]'
                            btn_editar_loc = frame.locator(btn_editar_sel).first if frame else page.locator(btn_editar_sel).first
                            btn_editar_loc.wait_for(state="visible", timeout=10000)
                            btn_editar_loc.click()
                            
                            # 2. Clic en select "Estado"
                            page.wait_for_timeout(1000)
                            estado_sel = 'button[name="ar7"], button[aria-label="Estado"]'
                            estado_loc = frame.locator(estado_sel).first if frame else page.locator(estado_sel).first
                            estado_loc.wait_for(state="visible", timeout=10000)
                            estado_loc.click()
                            
                            # 3. Seleccionar "Finalizado"
                            page.wait_for_timeout(1000)
                            finalizado_sel = 'button.rx-select__option:has-text("Finalizado"), button[role="option"]:has-text("Finalizado")'
                            finalizado_loc = frame.locator(finalizado_sel).first if frame else page.locator(finalizado_sel).first
                            finalizado_loc.wait_for(state="visible", timeout=5000)
                            finalizado_loc.click()
                            
                            # 4. Clic en "Guardar ticket"
                            page.wait_for_timeout(1000)
                            guardar_sel = 'button[name="ar304440891"], button[title="Guardar ticket"]'
                            guardar_loc = frame.locator(guardar_sel).first if frame else page.locator(guardar_sel).first
                            guardar_loc.wait_for(state="visible", timeout=5000)
                            guardar_loc.click()
                            print("Ticket finalizado y guardado exitosamente.")
                        except Exception as e:
                            print(f"Aviso: No se pudo ingresar o publicar la nota en el textarea. Detalle: {e}")
                    else:
                        print("Fallo la actualización.")
                else:
                    print("Aviso: No se pudo actualizar porque falta el Usuario de Datasoft.")
            else:
                print("no existe")
                print("Ingresando nota de rechazo al ticket...")
                mensaje_rechazo = """Buen dia.
Tras validar el ticket, confirmamos que la cuenta solicitada no existe. 
Por este motivo, procedemos con el rechazo de la solicitud. 
Por favor, genere un nuevo requerimiento para la creacion del usuario."""
                try:
                    # 1. Escribir nota
                    textarea_selector = 'textarea[data-testid="304247080"], textarea[name="ar304247080"]'
                    loc = frame.locator(textarea_selector).first if frame else page.locator(textarea_selector).first
                    loc.wait_for(state="visible", timeout=10000)
                    loc.fill(mensaje_rechazo)
                    
                    # 2. Clic en "Publicación"
                    btn_selector = 'button[name="ar304268430"], button[id="304268430"], button[title="Publicación"]'
                    btn_loc = frame.locator(btn_selector).first if frame else page.locator(btn_selector).first
                    btn_loc.wait_for(state="visible", timeout=5000)
                    btn_loc.click()
                    print("Nota de rechazo ingresada y publicada.")
                    
                    # Clic en "Asignarme a mí"
                    print("Asignando el ticket a mí mismo...")
                    page.wait_for_timeout(2000)
                    btn_asignar_sel = 'button[name="ar304421551"], button[id="304421551"], button[title="Asignarme a mí"]'
                    btn_asignar_loc = frame.locator(btn_asignar_sel).first if frame else page.locator(btn_asignar_sel).first
                    btn_asignar_loc.wait_for(state="visible", timeout=10000)
                    btn_asignar_loc.click()
                    print("Ticket asignado correctamente.")
                    
                    # 3. Clic en "Editar"
                    print("Cambiando el estado a Rechazado...")
                    page.wait_for_timeout(2000)
                    btn_editar_sel = 'button[name="ar304420591"], button[title="Editar"]'
                    btn_editar_loc = frame.locator(btn_editar_sel).first if frame else page.locator(btn_editar_sel).first
                    btn_editar_loc.wait_for(state="visible", timeout=10000)
                    btn_editar_loc.click()
                    
                    # 4. Clic en select "Estado"
                    page.wait_for_timeout(1000)
                    estado_sel = 'button[name="ar7"], button[aria-label="Estado"]'
                    estado_loc = frame.locator(estado_sel).first if frame else page.locator(estado_sel).first
                    estado_loc.wait_for(state="visible", timeout=10000)
                    estado_loc.click()
                    
                    # 5. Seleccionar "Rechazado"
                    page.wait_for_timeout(1000)
                    rechazado_sel = 'button.rx-select__option:has-text("Rechazado"), button[role="option"]:has-text("Rechazado")'
                    rechazado_loc = frame.locator(rechazado_sel).first if frame else page.locator(rechazado_sel).first
                    rechazado_loc.wait_for(state="visible", timeout=5000)
                    rechazado_loc.click()
                    
                    # 6. Clic en "Guardar ticket"
                    page.wait_for_timeout(1000)
                    guardar_sel = 'button[name="ar304440891"], button[title="Guardar ticket"]'
                    guardar_loc = frame.locator(guardar_sel).first if frame else page.locator(guardar_sel).first
                    guardar_loc.wait_for(state="visible", timeout=5000)
                    guardar_loc.click()
                    print("Ticket rechazado y guardado exitosamente.")
                except Exception as e:
                    print(f"Aviso: No se pudo ingresar la nota o registrar el rechazo. Detalle: {e}")
                    
            # Generar el registro en GLPI al finalizar independientemente de si existía o no
            if page:
                login_a_enviar = login_final if existe else None
                pass_a_enviar = cedula if existe else None
                generar_informe_glpi(page, datos_extraidos, login_a_enviar, pass_a_enviar)
                
        elif "Creación" in tipo_solicitud or "Creacion" in tipo_solicitud:
            print(f"Validando creacion de '{cedula}'...")
            if cedula:
                res = ejecutar_consulta_db(f"SELECT COUNT(*) FROM a_usuario WHERE NO_IDENTIFICACION = '{cedula}'", conn)
                if res and res.strip() != "0":
                    print("existe")
                else:
                    print("no existe")
            else:
                print("no existe") # Si no hay cédula, asumimos que no existe
    except Exception as e:
        print(f"Error en la consulta: {e}")

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

def validar_items(page, conn=None):
    print("=== Validando existencia de items en la tabla ===")
    try:
        # Esperamos a que aparezca la tabla vacía o la tabla con datos
        page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
        
        # Validamos cuál elemento cargó
        if page.locator(".tc__list-placeholder-text").is_visible():
            print("no hay items")
        elif page.locator(".ngViewport").is_visible():
            contar_items(page, conn)
        else:
            # Por si algo raro pasó con el DOM
            print("no se pudo determinar el estado de la tabla")
            
    except Exception as e:
        print(f"Aviso: Timeout al esperar la tabla. Detalle: {str(e)[:50]}")

def contar_items(page, conn=None):
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
        obtener_detalle_item(page, item_index=i, conn=conn)
        
        # Si no es el último ítem, volver a la consola de tickets para poder abrir el siguiente
        if i < count:
            print("Regresando a la consola de tickets...")
            page.go_back()
            # Esperar a que recargue la tabla original antes de buscar el siguiente
            page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
            page.wait_for_timeout(2000) # Pausa mínima para estabilizar la tabla


def obtener_detalle_item(page, item_index=1, conn=None):
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
                
            validar_datos_condicionales(datos_extraidos, conn, page, frame)
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
        headless=False, # Necesitamos ver el navegador
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

