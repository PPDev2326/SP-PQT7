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

# Para evitar procesar el mismo tipo varias veces
tipos_procesados = set()

# Lista para almacenar elementos que no tienen el parámetro destino
elementos_sin_parametro = []

def transferir_parametros(elem_type, elem_name=""):
    if not elem_type or elem_type.Id in tipos_procesados:
        return True
    
    tipos_procesados.add(elem_type.Id)
    
    # Obtener los parámetros origen
    param_number = elem_type.LookupParameter("Classification.Uniclass.Ss.Number")
    param_description = elem_type.LookupParameter("Classification.Uniclass.Ss.Description")
    param_destino = elem_type.LookupParameter("ClassificationReference.ItemReference")
    
    # Verificar si existe el parámetro destino
    if not param_destino:
        tipo_nombre = elem_type.FamilyName if hasattr(elem_type, 'FamilyName') else 'Desconocido'
        elementos_sin_parametro.append("Tipo: " + tipo_nombre + " - ID: " + str(elem_type.Id))
        return False
    
    # Verificar que el parámetro destino no sea de solo lectura
    if param_destino.IsReadOnly:
        tipo_nombre = elem_type.FamilyName if hasattr(elem_type, 'FamilyName') else 'Desconocido'
        elementos_sin_parametro.append("Tipo: " + tipo_nombre + " - ID: " + str(elem_type.Id) + " (Solo lectura)")
        return False
    
    # Obtener valores (usar cadena vacía si no hay valor)
    number_value = param_number.AsString() if param_number and param_number.HasValue else ""
    description_value = param_description.AsString() if param_description and param_description.HasValue else ""
    
    # Verificar que al menos uno de los parámetros origen tenga valor
    if not number_value and not description_value:
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
        return True
    except Exception as e:
        elementos_sin_parametro.append("Error al escribir en Tipo ID: " + str(elem_type.Id) + " - " + str(e))
        return False

elementos_procesados = 0
elementos_con_error = 0

with revit.Transaction("Transferir parámetros de tipo"):
    for elem in selected_elements:
        # Tipo principal del elemento
        if transferir_parametros(doc.GetElement(elem.GetTypeId())):
            elementos_procesados += 1
        else:
            elementos_con_error += 1
        
        # Verificar si el elemento es una instancia de familia (que puede tener subcomponentes)
        if isinstance(elem, FamilyInstance):
            sub_ids = elem.GetSubComponentIds()
            for sid in sub_ids:
                sub_elem = doc.GetElement(sid)
                if sub_elem and sub_elem.GetTypeId() != ElementId.InvalidElementId:
                    if transferir_parametros(doc.GetElement(sub_elem.GetTypeId())):
                        elementos_procesados += 1
                    else:
                        elementos_con_error += 1

# Mostrar resultado detallado
if elementos_sin_parametro:
    mensaje_error = "ATENCIÓN: Algunos elementos no tienen el parámetro 'ClassificationReference.ItemReference' o no se puede escribir en él:\n\n"
    mensaje_error += "\n".join(elementos_sin_parametro[:10])  # Mostrar máximo 10
    if len(elementos_sin_parametro) > 10:
        mensaje_error += "\n... y " + str(len(elementos_sin_parametro) - 10) + " más."
    
    mensaje_error += "\n\nElementos procesados exitosamente: " + str(elementos_procesados)
    mensaje_error += "\nElementos con problemas: " + str(elementos_con_error)
    mensaje_error += "\n\nSOLUCIÓN: Necesitas crear el parámetro 'ClassificationReference.ItemReference' en tus familias o tipos de elemento."
    
    forms.alert(mensaje_error, title="Resultado de la transferencia")
else:
    forms.alert("¡Parámetros transferidos correctamente!\n\nElementos procesados: " + str(elementos_procesados), title="Transferencia exitosa")