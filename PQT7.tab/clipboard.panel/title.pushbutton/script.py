# -*- coding: utf-8 -*-
__title__ = "Copy Title"

# --- IMPORTS ---
from System.Windows.Forms import Clipboard # <--- IMPORTANTE: Necesario para usar el portapapeles
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# Tus extensiones personalizadas
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from pyrevit import revit, script

doc, uidoc = revit.doc, revit.uidoc
output = script.get_output()

# --- 1. VALIDACIÓN INICIAL ---
nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# --- 2. OBTENER Y COPIAR ---
try:
    # Obtenemos el título del documento activo
    title = doc.Title
    
    # Verificamos que no esté vacío antes de copiar (SetText falla si es null)
    if title:
        Clipboard.SetText(title)
        
        # Opcional: Imprimir un mensaje de confirmación para que sepas que funcionó
        output.print_md("### ✅ Título copiado al portapapeles:")
        output.print_md("**{}**".format(title))
    else:
        output.print_md("⚠️ El documento no tiene título.")

except Exception as e:
    output.print_md("❌ Error al intentar copiar: {}".format(e))