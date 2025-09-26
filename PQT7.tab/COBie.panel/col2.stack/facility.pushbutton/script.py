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

try:
    list_elements = uidoc.Selection.PickElementsByRectangle("Selecciona los elementos para el COBie.Facility")

except OperationCanceledException as Oce:
    forms.alert("No se seleccionaron elementos operacion cancelada {}".format(Oce), "Cancelación")