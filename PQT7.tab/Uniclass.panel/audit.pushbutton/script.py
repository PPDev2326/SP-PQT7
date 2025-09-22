# -*- coding: utf-8 -*-
__title__ = "Uniclass Transfer"
__doc__ = """
Version = 1.2
Date = 22.09.2025
------------------------------------------------------------------
Description:
Nos permitira unir parametros Uniclass en el ItemReference.
------------------------------------------------------------------
¿Cómo hacerlo?
-> Click en el boton
-> Seleccionamos los elementos de la vista
-> Click en finalizar en la parte superior
------------------------------------------------------------------
Última actualización:
- [22.09.2025] - 1.2 UPDATE - Combinar Number y Description con espacio
- [22.09.2025] - 1.1 UPDATE - New Feature
------------------------------------------------------------------
Autor: Paolo Perez
"""
from pyrevit import forms, script, revit
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import ElementId, FamilyInstance
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

doc = revit.doc
uidoc = revit.uidoc

# Selección de elementos
try:
    selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecciona elementos en la vista")
except Exception as e:
    forms.alert("No se seleccionaron elementos o se produjo un error:\n\n" + str(e), exitscript=True)

if not selected_refs:
    forms.alert("No se seleccionaron elementos.", exitscript=True)

selected_elements = [doc.GetElement(ref) for ref in selected_refs]

# Para evitar procesar el mismo tipo varias veces
tipos_procesados = set()

def transferir_parametros(elem_type):
    if not elem_type or elem_type.Id in tipos_procesados:
        return
    
    tipos_procesados.add(elem_type.Id)
    
    # Obtener los parámetros origen
    param_number = elem_type.LookupParameter("Classification.Uniclass.Ss.Number")
    param_description = elem_type.LookupParameter("Classification.Uniclass.Ss.Description")
    param_destino = elem_type.LookupParameter("ClassificationReference.ItemReference")
    
    if param_destino:
        # Obtener valores (usar cadena vacía si no hay valor)
        number_value = param_number.AsString() if param_number and param_number.HasValue else ""
        description_value = param_description.AsString() if param_description and param_description.HasValue else ""
        
        # Combinar los valores con un espacio
        # Solo agregar espacio si ambos valores existen
        if number_value and description_value:
            combined_value = number_value + " " + description_value
        elif number_value:
            combined_value = number_value
        elif description_value:
            combined_value = description_value
        else:
            combined_value = ""
        
        # Asignar el valor combinado al parámetro destino
        param_destino.Set(combined_value)

with revit.Transaction("Transferir parámetros de tipo"):
    for elem in selected_elements:
        # Tipo principal del elemento
        transferir_parametros(doc.GetElement(elem.GetTypeId()))
        
        # Verificar si el elemento es una instancia de familia (que puede tener subcomponentes)
        if isinstance(elem, FamilyInstance):
            sub_ids = elem.GetSubComponentIds()
            for sid in sub_ids:
                sub_elem = doc.GetElement(sid)
                if sub_elem and sub_elem.GetTypeId() != ElementId.InvalidElementId:
                    transferir_parametros(doc.GetElement(sub_elem.GetTypeId()))

forms.alert("Parámetros transferidos correctamente.")