# -*- coding: utf-8 -*-
__title__ = "Uniclass Transfer"
__doc__ = """
Version = 1.2
Date = 22.09.2025
------------------------------------------------------------------
Description:
Nos permitira unir parametros Uniclass en el ItemReference.
IMPORTANTE: Requiere que el parámetro 'ClassificationReference.ItemReference' 
exista en los tipos de elemento.
------------------------------------------------------------------
¿Cómo hacerlo?
-> Click en el boton
-> Seleccionamos los elementos de la vista
-> Click en finalizar en la parte superior
------------------------------------------------------------------
Última actualización:
- [22.09.2025] - 1.2 UPDATE - Validación de parámetros y manejo de errores mejorado
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

# Para evitar procesar el mismo tipo varias veces y contar correctamente
tipos_procesados = set()
tipos_exitosos = set()
tipos_con_error = set()

def transferir_parametros(elem_type):
    if not elem_type or elem_type.Id in tipos_procesados:
        return None  # Ya fue procesado
    
    tipos_procesados.add(elem_type.Id)
    
    # Obtener los parámetros origen
    param_number = elem_type.LookupParameter("Classification.Uniclass.Ss.Number")
    param_description = elem_type.LookupParameter("Classification.Uniclass.Ss.Description")
    param_destino = elem_type.LookupParameter("ClassificationReference.ItemReference")
    
    # Verificar si existe el parámetro destino
    if not param_destino:
        tipos_con_error.add(elem_type.Id)
        return False
    
    # Verificar que el parámetro destino no sea de solo lectura
    if param_destino.IsReadOnly:
        tipos_con_error.add(elem_type.Id)
        return False
    
    # Obtener valores (usar cadena vacía si no hay valor)
    number_value = param_number.AsString() if param_number and param_number.HasValue else ""
    description_value = param_description.AsString() if param_description and param_description.HasValue else ""
    
    # Verificar que al menos uno de los parámetros origen tenga valor
    if not number_value and not description_value:
        tipos_exitosos.add(elem_type.Id)
        return True  # No hay datos que transferir, pero no es un error
    
    # Combinar los valores con un espacio
    if number_value and description_value:
        combined_value = number_value + " " + description_value
    elif number_value:
        combined_value = number_value
    elif description_value:
        combined_value = description_value
    else:
        combined_value = ""
    
    try:
        # Asignar el valor combinado al parámetro destino
        param_destino.Set(combined_value)
        tipos_exitosos.add(elem_type.Id)
        return True
    except Exception as e:
        tipos_con_error.add(elem_type.Id)
        return False

elementos_procesados = 0
elementos_con_error = 0

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

# Mostrar resultado con conteo de tipos únicos
total_tipos = len(tipos_procesados)
tipos_actualizados = len(tipos_exitosos)
tipos_no_actualizados = len(tipos_con_error)

mensaje = "RESULTADO DE LA TRANSFERENCIA:\n\n"
mensaje += "Total de tipos únicos procesados: " + str(total_tipos) + "\n"
mensaje += "Tipos actualizados exitosamente: " + str(tipos_actualizados) + "\n"
mensaje += "Tipos que NO se pudieron actualizar: " + str(tipos_no_actualizados)

if tipos_no_actualizados > 0:
    mensaje += "\n\nNOTA: Los tipos no actualizados probablemente no tienen el parámetro"
    mensaje += "\n'ClassificationReference.ItemReference' o este es de solo lectura."

forms.alert(mensaje, title="Resultado de Transferencia Uniclass")