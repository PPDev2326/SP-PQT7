# -*- coding: utf-8 -*-
__title__ = "IFC Room"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import *

doc = revit.doc
uidoc = revit.uidoc

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# ==== Diccionario de parametros IFC ====
parameters_ifc = {
    BuiltInParameter.IFC_EXPORT_ELEMENT_AS: "IfcSpace",
    BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE: "USERDEFINED"
}

# ==== Obtenemos las habitaciones del modelo ====
fec = FilteredElementCollector(doc)
list_element_room = fec.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

# ==== Contadores ====
total_params = 0
success_count = 0
failed_count = 0

# ==== Iniciamos la transaccion ====
with revit.Transaction("Asignamos IFC"):
    
    for element_Room in list_element_room:
        for builtparam, value in parameters_ifc.items():
            total_params += 1
            
            object_param = GetParameterAPI(element_Room, builtparam)
            
            # Solo necesitas verificar si el SetParameter fue exitoso
            if SetParameter(object_param, value):
                success_count += 1
            else:
                failed_count += 1
                # Opcional: registrar qué falló
                print("Fallo en Room Id {}: parametro {}".format(
                    element_Room.Id, builtparam))

# Mensaje de resumen
mensaje = "Proceso completado:\n"
mensaje += "Habitaciones procesadas: {}\n".format(len(list_element_room))
mensaje += "Parametros asignados: {} / {}\n".format(success_count, total_params)
if failed_count > 0:
    mensaje += "Fallos: {} (ver consola para detalles)".format(failed_count)

TaskDialog.Show("Resultado", mensaje)