def generar_informe_glpi(page, datos_extraidos=None, login_generado=None, password_generado=None, estado_helix=None, titulo_glpi="ACTIVACION USUARIO BANCO", id_ticket=None):
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
        print("Asignando usuario: Cristian Mejia Moreno...")
        
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
        search_assign.fill("Cristian")
        
        # Esperamos que cargue la lista flotante y damos clic
        glpi_page.wait_for_timeout(2000) # dar un instante para que reaccione GLPI
        opcion_busqueda = glpi_page.locator("li.select2-results__option", has_text="Cristian Mejia Moreno").first
        opcion_busqueda.wait_for(state="visible", timeout=10000)
        opcion_busqueda.click()
        
        # 3. Título (id dinámico)
        print("Ingresando título...")
        titulo_input = glpi_page.locator("input[id^='name_']")
        titulo_input.fill(titulo_glpi)
        
        # 4. Descripción (dentro del iframe de TinyMCE)
        print("Ingresando descripción...")
        iframe_locator = glpi_page.frame_locator("iframe.tox-edit-area__iframe")
        body_locator = iframe_locator.locator("body#tinymce")
        body_locator.wait_for(state="visible", timeout=10000)
        
        # Construir la descripción de manera dinámica
        descripcion_glpi = f"Buen día,\nMe colaboran gestionando este ticket del Helix ({id_ticket if id_ticket else ''}) para activación de usuario de Bancolombia.\nadjunto evidencia:\n\n"
        
        if id_ticket:
            descripcion_glpi += f"ID Item Helix: {id_ticket}\n"
        
        if estado_helix:
            descripcion_glpi += f"Estado fijado en Helix: {estado_helix}\n\n"
        
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
