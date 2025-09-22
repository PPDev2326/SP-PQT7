# -*- coding: utf-8 -*-
__title__ = "Uniclass Transfer"
__doc__ = """
Version = 1.0
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

# Mapeo de parámetros: origen -> destino
param_map = {
    "Classification.Uniclass.Ss.Number": "ClassificationReference.ItemReference",
    "Classification.Uniclass.Ss.Description": "ClassificationReference.Name"
}

# Para evitar procesar el mismo tipo varias veces
tipos_procesados = set()

def transferir_parametros(elem_type):
    if not elem_type or elem_type.Id in tipos_procesados:
        return
    tipos_procesados.add(elem_type.Id)

    for origen, destino in param_map.items():
        param_origen = elem_type.LookupParameter(origen)
        param_destino = elem_type.LookupParameter(destino)

        if param_origen and param_destino:
            valor = param_origen.AsString() or ""
            param_destino.Set(valor)

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