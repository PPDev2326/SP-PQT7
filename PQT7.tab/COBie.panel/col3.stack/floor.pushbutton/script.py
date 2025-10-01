# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms
import logging

# ==== obtenemos el documento y el uidocument del modelo activo ====
doc = revit.doc
uidoc = revit.uidoc

# ==== Instanciamos la salida output y logger ====
output = script.get_output()
logger = script.get_logger()
# logger.setLevel(logging.DEBUG)

# ==== Obtenemos los niveles del modelo con el FilteredElementCollector ====
fec = FilteredElementCollector(doc)
list_levels_object = fec.OfClass(Level).WhereElementIsNotElementType().ToElements()

list_level_name = []

for level in list_levels_object:
    list_level_name.append(level.Name)

# Log para debug
logger.info("Inicio del script")

# Mostrar tabla en el output
output.print_table(
    table_data=[["ID", "Nombre"], [1, "Muro"], [2, "Puerta"]],
    title="Elementos procesados"
)

# Mensaje con formato
output.print_md("### Procesamiento completado ✅")

logger.warning("Se ignoraron 2 elementos")