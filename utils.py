import os
import glob
import openpyxl
import pyotp
import tkinter as tk
from tkinter import simpledialog

# Mapeo de Línea de Negocio (Helix) → (CMPN_CODIGO, ESOR_CODIGO)
LINEAS_NEGOCIO = {
    "bancolombia": [("BANCOLOMBI", "PIC")],
    "sufi":        [("SUFI",       "PIC")],
    "leasing":     [("LEASINGBAN", "ARCHIVO")],
    "factoring":   [("FACTBANCOL", "GARANTIAS")],
    "todos":       [
        ("BANCOLOMBI", "PIC"),
        ("LEASINGBAN", "ARCHIVO"),
        ("SUFI",       "PIC"),
        ("FACTBANCOL", "GARANTIAS"),
    ],
}

def lineas_desde_helix(linea_negocio_texto):
    """Convierte el texto de Línea de Negocio de Helix a lista de (CMPN, ESOR)."""
    texto = linea_negocio_texto.strip().lower()
    for clave, valor in LINEAS_NEGOCIO.items():
        if clave in texto:
            return valor
    return [("BANCOLOMBI", "PIC")]  # Default

def generar_codigo_totp():
    """Genera el código TOTP automáticamente a partir del secret proporcionado."""
    totp = pyotp.TOTP('trrj2jhckpdgjzry')
    return totp.now()

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

def leer_excel_masivo(ruta_archivo):
    """
    Lee el archivo Excel de solicitud masiva.
    Formato esperado:
      Fila 12: Encabezados (CEDULA, NOMBRE, CARGO, ÁREA, CORREO ELECTRONICO, USUARIO REFERENCIA)
      Columnas: B=Cédula, C=Nombre, D=Cargo, E=Área, F=Correo, G=Usuario referencia
      Desde fila 13 en adelante: datos de usuarios
    Retorna una lista de diccionarios con los datos de cada usuario.
    """
    print(f"\n=== Leyendo Excel masivo: {ruta_archivo} ===")
    usuarios = []
    
    try:
        wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
        ws = wb.active  # Tomar la primera hoja
        
        # Verificar encabezados en fila 12
        encabezados = {
            'B': ws['B12'].value,
            'C': ws['C12'].value,
            'D': ws['D12'].value,
            'E': ws['E12'].value,
            'F': ws['F12'].value,
            'G': ws['G12'].value,
        }
        print(f"Encabezados detectados: {encabezados}")
        
        # Leer datos desde fila 13 en adelante
        fila = 13
        while True:
            cedula = ws[f'B{fila}'].value
            
            # Si la cédula está vacía, terminamos de leer
            if cedula is None or str(cedula).strip() == "":
                break
            
            nombre = ws[f'C{fila}'].value or ""
            correo = ws[f'F{fila}'].value or ""
            usuario_ref = ws[f'G{fila}'].value or ""
            
            usuario = {
                "cedula": str(cedula).strip(),
                "nombre": str(nombre).strip(),
                "correo": str(correo).strip(),
                "usuario_referencia": str(usuario_ref).strip().upper(),
            }
            usuarios.append(usuario)
            fila += 1
        
        wb.close()
        print(f"Se leyeron {len(usuarios)} usuario(s) del Excel.")
        
    except Exception as e:
        print(f"Error al leer el Excel: {e}")
    
    return usuarios
