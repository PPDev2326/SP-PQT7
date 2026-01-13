# -*- coding: utf-8 -*-
__title__ = "Copy Title"

# --- IMPORTS ---
import clr
# Agregamos la referencia a la DLL de Windows Forms
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Clipboard

from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# Tus extensiones
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
    title = doc.Title
    
    if title:
        Clipboard.SetText(title)
        # Mensaje opcional (puedes borrarlo si no quieres que salga la ventana)
        output.print_md("### Copiado: **{}**".format(title))
    else:
        output.print_md("⚠️ El documento no tiene título.")

except Exception as e:
    output.print_md("❌ Error: {}".format(e))