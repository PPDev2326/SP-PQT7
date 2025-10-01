# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms

# ==== obtenemos el documento y el uidocument del modelo activo ====
doc = revit.doc
uidoc = revit.uidoc

# ==== Obtenemos los niveles del modelo con el FilteredElementCollector ====
fec = FilteredElementCollector(doc)
list_levels_object = fec.OfClass(Level).WhereElementIsNotElementType().ToElements()

list_level_name = []

for level in list_levels_object:
    list_level_name.append(level.Name)

print(list_level_name)