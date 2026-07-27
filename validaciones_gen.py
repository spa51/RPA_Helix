import os
from db_utils import ejecutar_consulta_db, ejecutar_actualizacion_db, ejecutar_consulta_fila_db
from glpi_utils import generar_informe_glpi
from utils import leer_excel_masivo, lineas_desde_helix
from helix_actions import descargar_adjunto_excel, procesar_ticket_completo_helix, obtener_detalle_item

def generar_login(nombre_completo, conn):
    """
    Genera un LOGIN tomando las 2 primeras letras de cada palabra,
    hasta un máximo de 4 palabras (8 caracteres). Si el LOGIN ya existe,
    agrega un dígito incremental al final.
    """
    palabras = nombre_completo.upper().split()
    palabras = palabras[:4]  # máximo 4 palabras
    base = "".join([p[:2] for p in palabras if len(p) >= 2])
    login = base
    sufijo = 1
    while True:
        res = ejecutar_consulta_db(f"SELECT COUNT(*) FROM a_usuario WHERE LOGIN = '{login}'", conn)
        if res and res.strip() == "0":
            return login
        login = base + str(sufijo)
        sufijo += 1

def validar_fila_activacion(cedula, usuario_ref, conn):
    print(f"\n  >> Validando activación para cédula: {cedula}, usuario: {usuario_ref}")
    existe = False
    login_final = usuario_ref
    
    if usuario_ref:
        res = ejecutar_consulta_db(f"SELECT COUNT(*) FROM a_usuario WHERE LOGIN = '{usuario_ref}'", conn)
        if res and res.strip() != "0":
            existe = True
            print(f"  Usuario encontrado por login: {login_final}")
            
    if not existe and cedula:
        res_login = ejecutar_consulta_db(f"SELECT LOGIN FROM a_usuario WHERE NO_IDENTIFICACION = '{cedula}'", conn)
        if res_login and res_login.strip() not in ("0", ""):
            existe = True
            login_final = res_login.strip()
            print(f"  Usuario encontrado por cédula. LOGIN en BD: {login_final}")
            
    if existe and login_final:
        update_query = (
            f"UPDATE a_usuario SET CONTRASENA = NO_IDENTIFICACION, ESTADO = 'A', "
            f"CAMBIO_CONTRASENA = 'S' WHERE LOGIN = '{login_final}'"
        )
        exito = ejecutar_actualizacion_db(update_query, conn)
        if exito:
            return {"exito": True, "login": login_final, "mensaje": f"Cédula {cedula} - Usuario {login_final} - Contraseña {cedula}: Activado exitosamente."}
        else:
            return {"exito": False, "login": login_final, "mensaje": f"Cédula {cedula} - Usuario {login_final}: Error al actualizar."}
    else:
        return {"exito": False, "login": None, "mensaje": f"Cédula {cedula} - Usuario {usuario_ref}: No existe el usuario. Requiere creación."}

def validar_fila_creacion(cedula, nombre_completo, email_usuario, usuario_ref, conn):
    print(f"\n  >> Validando creación para cédula: {cedula}, nombre: {nombre_completo}")
    usuario_ref_upper = str(usuario_ref).strip().upper() if usuario_ref else ""
    ref_fila = ejecutar_consulta_fila_db(
        f"SELECT LOGIN, CMPN_CODIGO, ESOR_CODIGO, NOMBRE_USUARIO FROM a_usuario WHERE LOGIN = '{usuario_ref_upper}'", conn
    )
    if not ref_fila:
        return {"exito": False, "login": None, "mensaje": f"Usuario de referencia '{usuario_ref}' no existe."}
        
    ref_cmpn = ref_fila.get("CMPN_CODIGO", "BANCOLOMBI")
    ref_esor = ref_fila.get("ESOR_CODIGO", "PIC")
    
    login_existente = ejecutar_consulta_db(f"SELECT LOGIN FROM a_usuario WHERE NO_IDENTIFICACION = '{cedula}'", conn)
    usuario_ya_existe = login_existente and login_existente.strip() not in ("0", "")
    lineas_requeridas = [(ref_cmpn, ref_esor)]
    
    if usuario_ya_existe:
        login_existente = login_existente.strip()
        lineas_faltantes = []
        for cmpn, esor in lineas_requeridas:
            res_autr = ejecutar_consulta_db(
                f"SELECT COUNT(*) FROM AUTORIZADO WHERE AUTR_CODIGO = '{login_existente}' AND AUTR_CMPN_CODIGO = '{cmpn}'", conn
            )
            if res_autr and res_autr.strip() == "0":
                lineas_faltantes.append((cmpn, esor))
        
        if not lineas_faltantes:
            return {"exito": False, "login": login_existente, "mensaje": f"Cédula {cedula} - Usuario {login_existente}: Ya existe con todas las autorizaciones."}
        else:
            for cmpn, esor in lineas_faltantes:
                q_autr = f"INSERT INTO AUTORIZADO (AUTR_CMPN_CODIGO, AUTR_ESOR_CODIGO, AUTR_CODIGO, AUTR_NOMBRE) VALUES ('{cmpn}', '{esor}', '{login_existente}', '{nombre_completo}')"
                ejecutar_actualizacion_db(q_autr, conn)
                q_serie = (
                    f"INSERT INTO autorizado_serie (AUSR_CMPN_CODIGO, AUSR_ESOR_CODIGO, AUSR_SRDC_CODIGO, AUSR_DCMT_CODIGO, AUSR_AUTR_CODIGO, AUTR_NOMBRE) "
                    f"SELECT a.AUSR_CMPN_CODIGO, a.AUSR_ESOR_CODIGO, a.AUSR_SRDC_CODIGO, a.AUSR_DCMT_CODIGO, '{login_existente}', '{nombre_completo}' "
                    f"FROM autorizado_serie a WHERE a.AUSR_AUTR_CODIGO = '{usuario_ref}' "
                    f"AND NOT EXISTS (SELECT 1 FROM autorizado_serie b WHERE b.AUSR_CMPN_CODIGO = a.AUSR_CMPN_CODIGO "
                    f"AND b.AUSR_ESOR_CODIGO = a.AUSR_ESOR_CODIGO AND b.AUSR_SRDC_CODIGO = a.AUSR_SRDC_CODIGO "
                    f"AND b.AUSR_DCMT_CODIGO = a.AUSR_DCMT_CODIGO AND b.AUSR_AUTR_CODIGO = '{login_existente}')"
                )
                ejecutar_actualizacion_db(q_serie, conn)
            return {"exito": True, "login": login_existente, "mensaje": f"Cédula {cedula} - Usuario {login_existente}: Autorizaciones añadidas exitosamente."}
    else:
        nuevo_login = generar_login(nombre_completo, conn)
        q_usuario = (
            f"INSERT INTO a_usuario (ID_USUARIO, LOGIN, CONTRASENA, ID_GRUPO_USUARIO, NOMBRE_USUARIO, "
            f"ESTADO, LOGIN_USUARIO, FECHA_ULTIMA_ACT, CMPN_CODIGO, ESOR_CODIGO, E_MAIL, TIPO_USUARIO, "
            f"CAMBIO_CONTRASENA, FECHA_CAMBIO_CONTRASENA, NO_INTENTOS, RESTABLECE_CONTRASENA, NO_IDENTIFICACION, FECHA_CREACION) "
            f"VALUES (seq_a_usuario.nextval, '{nuevo_login}', '{cedula}', 462, '{nombre_completo}', 'A', 'SPINE', SYSDATE, "
            f"'{ref_cmpn}', '{ref_esor}', '{email_usuario}', 'E', 'N', SYSDATE, 0, 'S', '{cedula}', SYSDATE)"
        )
        if not ejecutar_actualizacion_db(q_usuario, conn):
            return {"exito": False, "login": None, "mensaje": f"Cédula {cedula}: Error al crear usuario ."}
            
        for cmpn, esor in lineas_requeridas:
            q_autr = f"INSERT INTO AUTORIZADO (AUTR_CMPN_CODIGO, AUTR_ESOR_CODIGO, AUTR_CODIGO, AUTR_NOMBRE) VALUES ('{cmpn}', '{esor}', '{nuevo_login}', '{nombre_completo}')"
            ejecutar_actualizacion_db(q_autr, conn)
            
        q_serie = (
            f"INSERT INTO autorizado_serie (AUSR_CMPN_CODIGO, AUSR_ESOR_CODIGO, AUSR_SRDC_CODIGO, AUSR_DCMT_CODIGO, AUSR_AUTR_CODIGO, AUTR_NOMBRE) "
            f"SELECT a.AUSR_CMPN_CODIGO, a.AUSR_ESOR_CODIGO, a.AUSR_SRDC_CODIGO, a.AUSR_DCMT_CODIGO, '{nuevo_login}', '{nombre_completo}' "
            f"FROM autorizado_serie a WHERE a.AUSR_AUTR_CODIGO = '{usuario_ref}' "
            f"AND NOT EXISTS (SELECT 1 FROM autorizado_serie b WHERE b.AUSR_CMPN_CODIGO = a.AUSR_CMPN_CODIGO "
            f"AND b.AUSR_ESOR_CODIGO = a.AUSR_ESOR_CODIGO AND b.AUSR_SRDC_CODIGO = a.AUSR_SRDC_CODIGO "
            f"AND b.AUSR_DCMT_CODIGO = a.AUSR_DCMT_CODIGO AND b.AUSR_AUTR_CODIGO = '{nuevo_login}')"
        )
        ejecutar_actualizacion_db(q_serie, conn)
        return {"exito": True, "login": nuevo_login, "mensaje": f"Cédula {cedula} - Usuario {nuevo_login}: Creado exitosamente. Contraseña: {cedula}"}

def procesar_masiva(datos_extraidos, conn, page, frame, tipo_solicitud, id_ticket=None):
    print("\n========================================\n  PROCESAMIENTO MASIVO INICIADO\n========================================")
    ruta_excel = descargar_adjunto_excel(page, frame)
    if not ruta_excel or not os.path.exists(ruta_excel):
        msg_error = ("Buen día.\nNo fue posible descargar o localizar el archivo Excel adjunto al ticket.\n"
                     "Por favor, verifique que el archivo esté adjunto correctamente y genere un nuevo requerimiento.")
        procesar_ticket_completo_helix(page, frame, msg_error, "Rechazado")
        return
        
    usuarios = leer_excel_masivo(ruta_excel)
    if not usuarios:
        msg_vacio = ("Buen día.\nEl archivo Excel adjunto no contiene datos de usuarios a partir de la fila 13.\n"
                     "Por favor, verifique el formato del archivo y genere un nuevo requerimiento.")
        procesar_ticket_completo_helix(page, frame, msg_vacio, "Rechazado")
        return

    resultados = []
    exitosos = 0
    fallidos = 0
    es_activacion = tipo_solicitud == "Activación/desbloqueo"
    es_creacion = "Creación" in tipo_solicitud or "Creacion" in tipo_solicitud
    
    for idx, usuario in enumerate(usuarios, start=1):
        if es_activacion:
            res = validar_fila_activacion(usuario["cedula"], "", conn)
        elif es_creacion:
            res = validar_fila_creacion(usuario["cedula"], usuario["nombre"], usuario["correo"], usuario["usuario_referencia"], conn)
        else:
            res = validar_fila_activacion(usuario["cedula"], "", conn)
        resultados.append(res)
        if res["exito"]: exitosos += 1
        else: fallidos += 1
        
    tipo_texto = "ACTIVACIÓN" if es_activacion else "CREACIÓN"
    resumen = f"PROCESAMIENTO MASIVO DE {tipo_texto}\n\nDETALLE POR USUARIO:\n" + "-"*50 + "\n"
    for idx, r in enumerate(resultados, start=1):
        resumen += f"{idx}. {r['mensaje']}\n"
    resumen += "-"*50 + "\n"
    
    instrucciones = (
        "Digita en la página inicial de Datasoft http://172.17.21.14/Datasoft/login.php los campos: Usuario, "
        "compañía, sin darle Clave y código de seguridad y luego das clic en Restablecer Contraseña.\n"
        "IMPORTANTE: Se debe ingresar al sistema antes de 24 horas para Activar el Usuario.\n"
        "Tener en cuenta al cambiar la contraseña:\n"
        "• Longitud mínima 8 caracteres\n• Longitud máxima 12 caracteres\n"
        "• Debe contener como mínimo (2) dos letras\n• Debe contener como mínimo (2) dos números\n"
        "• Debe contener como mínimo (2) dos caracteres especiales de la siguiente lista: $ - _ =\n"
        "• No se pueden repetir contraseñas anteriores.\n• No usar el asterisco * (asterisco)\n"
    )
    
    if es_activacion and exitosos > 0:
        resumen += "\nINSTRUCCIONES PARA LOS USUARIOS ACTIVADOS:\n" + instrucciones
    if es_creacion and exitosos > 0:
        resumen += "\nINSTRUCCIONES PARA LOS USUARIOS CREADOS:\nPOR FAVOR LEER MUY DESPACIO...\n\n" + instrucciones
    resumen += "\nSaludos,"
    
    procesar_ticket_completo_helix(page, frame, resumen, "Finalizado")
    
    if page:
        titulo_glpi = f"{tipo_texto} USUARIO BANCO - MASIVO ({exitosos}/{len(usuarios)} exitosos)"
        generar_informe_glpi(page, datos_extraidos, None, None, "Finalizado", titulo_glpi, id_ticket=id_ticket)
        
    try:
        if ruta_excel and os.path.exists(ruta_excel) and ("rpa_helix_downloads" in ruta_excel or "Downloads" in ruta_excel):
            os.remove(ruta_excel)
    except: pass
    print("\n========================================\n  PROCESAMIENTO MASIVO COMPLETADO\n========================================")

def validar_datos_condicionales(datos_extraidos, conn, page=None, frame=None, id_ticket=None):
    info_servicio = datos_extraidos.get("Información del servicio", "").lower()
    tipo_solicitud = datos_extraidos.get("Tipo solicitud", "")
    usuario_datasoft = datos_extraidos.get("Usuario de Datasoft", "")
    cedula = datos_extraidos.get("No. Cédula", "")
    
    if "masiva" in info_servicio:
        procesar_masiva(datos_extraidos, conn, page, frame, tipo_solicitud, id_ticket=id_ticket)
        return

    if tipo_solicitud == "Activación/desbloqueo":
        res = validar_fila_activacion(cedula, usuario_datasoft, conn)
        if res["exito"]:
            mensaje = (
                f"POR FAVOR LEER MUY DESPACIO Y SEGUIR EL PASO A PASO...\n\nBuen día.\nTu usuario fue activado;\n"
                f"Usuario: {res['login']}\nContraseña: {cedula}\n\nDigita en la página inicial de Datasoft "
                "http://172.17.21.14/Datasoft/login.php los campos: Usuario, compañía, sin darle Clave y código de seguridad y luego das clic en Restablecer Contraseña.\n"
                "IMPORTANTE: Se debe ingresar al sistema antes de 24 horas para Activar el Usuario.\n"
                "• Longitud mínima 8 caracteres\n• Longitud máxima 12 caracteres\n• Debe contener como mínimo (2) dos letras\n"
                "• Debe contener como mínimo (2) dos números\n• Debe contener como mínimo (2) dos caracteres especiales: $ - _ =\n"
                "• No se pueden repetir contraseñas.\n• No usar el asterisco * (asterisco)\n\nSaludos,"
            )
            procesar_ticket_completo_helix(page, frame, mensaje, "Finalizado")
        else:
            msg = "Buen dia.\nTras validar el ticket, confirmamos que la cuenta solicitada no existe.\nPor este motivo, procedemos con el rechazo."
            procesar_ticket_completo_helix(page, frame, msg, "Rechazado")
            
        if page:
            generar_informe_glpi(page, datos_extraidos, res["login"], cedula if res["exito"] else None, "Finalizado" if res["exito"] else "Rechazado", id_ticket=id_ticket)

    elif "Creación" in tipo_solicitud or "Creacion" in tipo_solicitud:
        nombre = datos_extraidos.get("Nombre Completo", "")
        correo = datos_extraidos.get("Correo", "")
        usuario_ref = datos_extraidos.get("Usuario de referencia DATASOFT", "")
        linea = datos_extraidos.get("Línea de Negocio", "")
        
        res = validar_fila_creacion(cedula, nombre, correo, usuario_ref, conn)
        
        if "no existe" in res["mensaje"]:
            procesar_ticket_completo_helix(page, frame, res["mensaje"] + "\nProcedemos con el rechazo.", "Rechazado")
            estado = "Rechazado"
        elif "Ya existe con todas las autorizaciones" in res["mensaje"]:
            procesar_ticket_completo_helix(page, frame, res["mensaje"] + "\nProcedemos con el rechazo. Generar requerimiento de Activación.", "Rechazado")
            estado = "Rechazado"
        elif "Autorizaciones añadidas" in res["mensaje"]:
            procesar_ticket_completo_helix(page, frame, res["mensaje"] + "\nEl usuario puede iniciar sesión con su contraseña actual.", "Finalizado")
            estado = "Finalizado"
        elif res["exito"]:
            msg = (
                f"POR FAVOR LEER MUY DESPACIO...\n\nBuen día.\nTu usuario ha sido creado;\n"
                f"Usuario: {res['login']}\nContraseña: {cedula}\n\nDigita en la página inicial de Datasoft "
                "http://172.17.21.14/Datasoft/login.php los campos: Usuario, compañía, sin darle Clave y código de seguridad y luego das clic en Restablecer Contraseña.\n"
                "IMPORTANTE: Se debe ingresar al sistema antes de 24 horas para Activar el Usuario.\n"
                "• Longitud mínima 8 caracteres\n• Longitud máxima 12 caracteres\n• Debe contener como mínimo (2) dos letras\n"
                "• Debe contener como mínimo (2) dos números\n• Debe contener como mínimo (2) dos caracteres especiales: $ - _ =\n"
                "• No se pueden repetir contraseñas.\n• No usar el asterisco * (asterisco)\n\nSaludos,"
            )
            procesar_ticket_completo_helix(page, frame, msg, "Finalizado")
            estado = "Finalizado"
        else:
            procesar_ticket_completo_helix(page, frame, res["mensaje"], "Rechazado")
            estado = "Rechazado"
            
        if page:
            generar_informe_glpi(page, datos_extraidos, res["login"], cedula if res["exito"] else None, estado, id_ticket=id_ticket)

def contar_items(page, conn=None):
    count = 0
    while True:
        selector = f".ng-scope:nth-child({count + 1}) > .col2 .ngCellText"
        if page.locator(selector).count() > 0:
            count += 1
        else:
            break
    print(f"hay items: {count} item(s) encontrado(s)")
    for i in range(count):
        print(f"Procesando ítem {i + 1} de {count} (siempre se toma el primero de la lista)...")
        obtener_detalle_item(page, item_index=1, conn=conn)
        if i < count - 1:
            print("Regresando a la consola de tickets...")
            page.go_back()
            page.wait_for_selector(".tc__list-placeholder-text, .ngViewport", state="visible", timeout=30000)
            page.wait_for_timeout(2000)
