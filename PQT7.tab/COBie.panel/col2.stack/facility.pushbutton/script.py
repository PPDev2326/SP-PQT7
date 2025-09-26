# -*- coding: utf-8 -*-
__title__ = "COBie Facility"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, script, forms
from Extensions._RevitAPI import getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

uidoc = revit.uidoc
doc = revit.doc

# ==== Obtenemnos la instancia del Project Information ====
fec = FilteredElementCollector(doc)
project_information_element = fec.OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
print("los elementos encontrados son : {}".format(len(project_information_element)))

# ==== Seleccionamos los elementos del modelo activo ====
try:
    list_elements = uidoc.Selection.PickElementsByRectangle("Selecciona los elementos para el COBie.Facility")

except OperationCanceledException:
    forms.alert("Operación cancelada: no se seleccionaron elementos para procesar COBie.Facility",
    title="Cancelación")

