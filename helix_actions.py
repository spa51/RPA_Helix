import os
import tempfile
import glob

def descargar_adjunto_excel(page, frame):
    """
    Descarga el archivo Excel adjunto del ticket en Helix.
    """
    print("\n=== Descargando Excel adjunto del ticket ===")
    try:
        adjunto_selector = (
            'a[href*=".xlsx"], a[href*=".xls"], '
            'a[title*=".xlsx"], a[title*=".xls"], '
            'span:has-text(".xlsx"), span:has-text(".xls"), '
            'button:has-text(".xlsx"), button:has-text(".xls")'
        )
        loc = frame.locator(adjunto_selector).first if frame else page.locator(adjunto_selector).first
        
        try:
            loc.wait_for(state="visible", timeout=5000)
        except:
            print("Buscando adjunto con selectores alternativos...")
            alt_selector = (
                '[class*="attachment"] a, '
                '[class*="Attachment"] a, '
                '[id*="attachment"] a, '
                'a[class*="file"], '
                'img[class*="attachment-icon"]'
            )
            loc = frame.locator(alt_selector).first if frame else page.locator(alt_selector).first
            loc.wait_for(state="visible", timeout=10000)
        
        download_dir = os.path.join(tempfile.gettempdir(), "rpa_helix_downloads")
        os.makedirs(download_dir, exist_ok=True)
        
        try:
            for f in glob.glob(os.path.join(download_dir, "*")):
                os.remove(f)
        except Exception as e:
            print(f"Aviso: No se pudo limpiar la carpeta temporal: {e}")
        
        with page.expect_download(timeout=30000) as download_info:
            loc.click()
        
        download = download_info.value
        archivo_descargado = os.path.join(download_dir, download.suggested_filename)
        download.save_as(archivo_descargado)
        
        print(f"Archivo descargado: {archivo_descargado}")
        return archivo_descargado
        
    except Exception as e:
        print(f"Error al descargar el adjunto Excel: {e}")
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        archivos_excel = glob.glob(os.path.join(downloads_path, "*.xlsx")) + glob.glob(os.path.join(downloads_path, "*.xls"))
        if archivos_excel:
            mas_reciente = max(archivos_excel, key=os.path.getmtime)
            print(f"Fallback: Usando archivo más reciente encontrado en Downloads: {mas_reciente}")
            return mas_reciente
        return None

def publicar_nota_ticket(page, frame, mensaje):
    """Escribe una nota en el textarea del ticket y hace clic en publicar."""
    try:
        textarea_selector = 'textarea[data-testid="304247080"], textarea[name="ar304247080"]'
        loc = frame.locator(textarea_selector).first if frame else page.locator(textarea_selector).first
        loc.wait_for(state="visible", timeout=10000)
        loc.fill(mensaje)
        
        btn_selector = 'button[name="ar304268430"], button[id="304268430"], button[title="Publicación"]'
        btn_loc = frame.locator(btn_selector).first if frame else page.locator(btn_selector).first
        btn_loc.wait_for(state="visible", timeout=5000)
        btn_loc.click()
        print("Nota ingresada y publicada correctamente en el ticket.")
        return True
    except Exception as e:
        print(f"Aviso: No se pudo ingresar o publicar la nota en el textarea. Detalle: {e}")
        return False

def asignar_ticket_a_mi(page, frame):
    """Asigna el ticket al usuario actual."""
    try:
        print("Asignando el ticket a mí mismo...")
        page.wait_for_timeout(2000)
        btn_asignar_sel = 'button[name="ar304421551"], button[id="304421551"], button[title="Asignarme a mí"]'
        btn_asignar_loc = frame.locator(btn_asignar_sel).first if frame else page.locator(btn_asignar_sel).first
        btn_asignar_loc.wait_for(state="visible", timeout=10000)
        btn_asignar_loc.click()
        print("Ticket asignado correctamente.")
        return True
    except Exception as e:
        print(f"Aviso: Error al asignar ticket: {e}")
        return False

def cambiar_estado_ticket(page, frame, estado):
    """Cambia el estado del ticket (e.g., 'Finalizado' o 'Rechazado')."""
    try:
        print(f"Cambiando el estado a {estado}...")
        page.wait_for_timeout(2000)
        btn_editar_sel = 'button[name="ar304420591"], button[title="Editar"]'
        btn_editar_loc = frame.locator(btn_editar_sel).first if frame else page.locator(btn_editar_sel).first
        btn_editar_loc.wait_for(state="visible", timeout=10000)
        btn_editar_loc.click()
        
        page.wait_for_timeout(1000)
        estado_sel = 'button[name="ar7"], button[aria-label="Estado"]'
        estado_loc = frame.locator(estado_sel).first if frame else page.locator(estado_sel).first
        estado_loc.wait_for(state="visible", timeout=10000)
        estado_loc.click()
        
        page.wait_for_timeout(1000)
        
        if estado == "Finalizado":
            estado_btn_sel = (
                'button.rx-select__option:has-text("Finalizado"), '
                'button[role="option"]:has-text("Finalizado"), '
                'button.rx-select__option:has-text("Completed"), '
                'button[role="option"]:has-text("Completed")'
            )
        elif estado == "Rechazado":
            estado_btn_sel = (
                'button.rx-select__option:has-text("Rechazado"), '
                'button[role="option"]:has-text("Rechazado"), '
                'button.rx-select__option:has-text("Rejected"), '
                'button[role="option"]:has-text("Rejected")'
            )
        else:
            raise ValueError(f"Estado '{estado}' no soportado.")
            
        estado_btn_loc = frame.locator(estado_btn_sel).first if frame else page.locator(estado_btn_sel).first
        estado_btn_loc.wait_for(state="visible", timeout=5000)
        estado_btn_loc.click()
        return True
    except Exception as e:
        print(f"Aviso: Error al cambiar estado del ticket a {estado}: {e}")
        return False

def guardar_ticket(page, frame):
    """Guarda los cambios del ticket."""
    try:
        page.wait_for_timeout(1000)
        guardar_sel = 'button[name="ar304440891"], button[title="Guardar ticket"]'
        guardar_loc = frame.locator(guardar_sel).first if frame else page.locator(guardar_sel).first
        guardar_loc.wait_for(state="visible", timeout=5000)
        guardar_loc.click()
        print("Ticket guardado exitosamente.")
        return True
    except Exception as e:
        print(f"Aviso: Error al guardar ticket: {e}")
        return False

def procesar_ticket_completo_helix(page, frame, mensaje, estado):
    """Helper que agrupa: publicar nota -> asignar -> cambiar estado -> guardar."""
    if mensaje:
        publicar_nota_ticket(page, frame, mensaje)
    asignar_ticket_a_mi(page, frame)
    if estado:
        cambiar_estado_ticket(page, frame, estado)
        guardar_ticket(page, frame)

def validar_items(page, conn=None):
    from validaciones_gen import contar_items # Importamos localmente para evitar circular
    print("=== Validando existencia de items en la tabla ===")
    try:
        page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
        if page.locator(".tc__list-placeholder-text").is_visible():
            print("no hay items")
        elif page.locator(".ngViewport").is_visible():
            contar_items(page, conn)
        else:
            print("no se pudo determinar el estado de la tabla")
    except Exception as e:
        print(f"Aviso: Timeout al esperar la tabla. Detalle: {str(e)[:50]}")

def obtener_detalle_item(page, item_index=1, conn=None):
    from validaciones_gen import validar_datos_condicionales
    print(f"\n=== Extrayendo detalle del ítem #{item_index} ===")
    try:
        id_ticket_selector = f".ng-scope:nth-child({item_index}) > .col2 .ngCellText"
        id_ticket = "No encontrado"
        try:
            page.locator(id_ticket_selector).first.wait_for(state="visible", timeout=10000)
            id_ticket = page.locator(id_ticket_selector).first.inner_text().strip()
            print(f"ID del Ticket Helix detectado: {id_ticket}")
        except Exception as e:
            print(f"Aviso: No se pudo extraer el ID del ticket: {e}")

        fila_selector = f".ng-scope:nth-child({item_index}) > .col2 .ngCellText"
        page.locator(fila_selector).first.click()
        print("Clic realizado, entrando a la vista del ticket...")

        frame = page.frame_locator("#pwa-frame")
        cuadro_datos = frame.locator("#ar1000000151_data")
        cuadro_datos.first.wait_for(state="attached", timeout=20000)
        
        texto = cuadro_datos.first.get_attribute("title")
        if not texto:
            texto = cuadro_datos.first.inner_text()

        campos_buscados = [
            "No. Cédula:", "Tipo solicitud:", "Usuario de Datasoft:",
            "Información del servicio:", "Usuario de referencia DATASOFT:", "Línea de Negocio:"
        ]
        
        datos_extraidos = {}
        for linea in texto.splitlines():
            for campo in campos_buscados:
                if campo in linea:
                    partes = linea.split(campo)
                    if len(partes) > 1:
                        nombre_amigable = campo.replace(":", "").strip()
                        datos_extraidos[nombre_amigable] = partes[1].strip()
                        break

        try:
            nombre_loc = frame.locator("#ar301395400_data")
            if nombre_loc.count() == 0: nombre_loc = page.locator("#ar301395400_data")
            if nombre_loc.count() > 0: datos_extraidos["Nombre Completo"] = nombre_loc.first.inner_text().strip()
        except: pass

        try:
            correo_loc = frame.locator("#ar1000000048_data")
            if correo_loc.count() == 0: correo_loc = page.locator("#ar1000000048_data")
            if correo_loc.count() > 0: datos_extraidos["Correo"] = correo_loc.first.inner_text().strip()
        except: pass

        print("\n=================================")
        if datos_extraidos:
            print("DATOS EXTRAIDOS DEL ITEM:")
            for clave, valor in datos_extraidos.items():
                clave_segura = clave.replace('é', 'e').replace('í', 'i')
                print(f"  {clave_segura}: {valor}")
            
            # Llamamos a validaciones
            validar_datos_condicionales(datos_extraidos, conn, page, frame, id_ticket=id_ticket)
        else:
            print("Fallo: No se vio ningun texto de los solicitados.")
            print("Texto extraido para depurar:\n", texto)
        print("=================================\n")

    except Exception as e:
        print(f"Error al extraer los datos: {str(e)[:150]}")
