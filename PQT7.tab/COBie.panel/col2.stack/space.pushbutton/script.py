# -*- coding: utf-8 -*-
__title__ = "COBie Space"
__doc__ = """
Version = 1.1
Date = 23.09.2025
------------------------------------------------------------------
Description:
Llenara los parametros del COBieSpace.
------------------------------------------------------------------
¿Cómo hacerlo?
-> Click en el boton.
-> Esperamos que se complete la actualizacion de información.
------------------------------------------------------------------
Última actualización:
- [23.09.2025] - 1.1 UPDATE - New Feature
------------------------------------------------------------------
Autor: Paolo Perez
"""

from pyrevit import forms, script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    BuiltInParameter
)
from collections import defaultdict
from Extensions._RevitAPI import GetParameterAPI, getParameter, get_param_value
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Helper._Dictionary import get_formatted_string, find_mapped_number, ROOM_NAME_MAPPING

output = script.get_output()

def set_param(param, valor, elemento, nombre_param):
    """Setea un parámetro y devuelve True si fue exitoso."""
    if not param:
        output.print_md("⚠️ No se encontró el parámetro **{}** en el elemento `{}`".format(
            nombre_param, elemento.Id.IntegerValue))
        return False
    if param.IsReadOnly:
        output.print_md("🔒 El parámetro **{}** es de solo lectura en el elemento `{}`".format(
            param.Definition.Name, elemento.Id.IntegerValue))
        return False
    try:
        param.Set(valor)
        output.print_md("✅ Asignado **{} = {}** al elemento `{}`".format(
            nombre_param, valor, elemento.Id.IntegerValue))
        return True
    except Exception as e:
        output.print_md("❌ Error al asignar **{}** en el elemento `{}`: {}".format(
            nombre_param, elemento.Id.IntegerValue, e))
        return False

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

doc = revit.doc

# Datos del usuario
CORREO = "jtiburcio@syp.com.pe"
FECHA = "2023-04-07T16:38:56"

# Obtener Rooms y Spaces
spaces = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_MEPSpaces)
    .WhereElementIsNotElementType()
    .ToElements()
)

rooms = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Rooms)
    .WhereElementIsNotElementType()
    .ToElements()
)

elementos = spaces + rooms
output.print_md("### 🔍 Iniciando transferencia COBie")
output.print_md("- Rooms encontrados: **{}**".format(len(rooms)))
output.print_md("- Spaces encontrados: **{}**".format(len(spaces)))
output.print_md("- Total de elementos a procesar: **{}**".format(len(elementos)))

abreviaciones = defaultdict(int)
asignados_ok = 0
asignados_fail = 0
elementos_procesados = 0
elementos_omitidos = 0

with revit.Transaction("Transfiere datos a Parametros COBieSpace"):
    for i, elemento in enumerate(elementos, start=1):
        output.print_md("---")
        output.print_md("➡️ Procesando elemento `{}` ({}/{})".format(
            elemento.Id.IntegerValue, i, len(elementos)))

        room_param_name = GetParameterAPI(elemento, BuiltInParameter.ROOM_NAME)
        room_param_number = GetParameterAPI(elemento, BuiltInParameter.ROOM_NUMBER)
        cl_param_description = getParameter(elemento, "Classification.Space.Description")
        cl_param_number = getParameter(elemento, "Classification.Space.Number")
        
        value_room_name = get_param_value(room_param_name, "Sin nombre")
        value_room_number = get_param_value(room_param_number, "Sin numero")
        value_cl_description = get_param_value(cl_param_description, "Sin nombre")
        value_cl_number = get_param_value(cl_param_number, "Sin nombre")
        name_full = get_formatted_string(value_room_number, value_room_name)
        
        # CORRECCIÓN: Buscar en las claves del diccionario, no en los valores
        if name_full in ROOM_NAME_MAPPING:  # o alternativamente: if name_full in ROOM_NAME_MAPPING.keys():
            elementos_procesados += 1
            output.print_md("✅ Elemento encontrado en mapping: **{}**".format(name_full))
            
            categoria = get_formatted_string(value_cl_number, value_cl_description)
            room_tag = find_mapped_number(name_full)

            height = GetParameterAPI(elemento, BuiltInParameter.ROOM_UPPER_OFFSET)
            area = GetParameterAPI(elemento, BuiltInParameter.ROOM_AREA)

            height_val = get_param_value(height, 0)
            area_val = get_param_value(area, 0)

            parametros = {
                "COBie.Space.Name": name_full,
                "COBie.CreatedBy": CORREO,
                "COBie.CreatedOn": FECHA,
                "COBie.Space.Category": categoria,
                "COBie.Space.Description": value_room_name,
                "COBie.Space.RoomTag": room_tag,
                "COBie.Space.UsableHeight": height_val,
                "COBie.Space.GrossArea": area_val,
                "COBie.Space.NetArea": area_val,
                "COBie": 1,
            }

            for nombre_param, valor in parametros.items():
                if set_param(elemento.LookupParameter(nombre_param), valor, elemento, nombre_param):
                    asignados_ok += 1
                else:
                    asignados_fail += 1
        else:
            elementos_omitidos += 1
            output.print_md("⚠️ Elemento omitido (no encontrado en mapping): **{}**".format(name_full))

output.print_md("---")
output.print_md("### 📊 Resumen final")
output.print_md("- Total elementos analizados: **{}**".format(len(elementos)))
output.print_md("- Elementos procesados (con mapping): **{}**".format(elementos_procesados))
output.print_md("- Elementos omitidos (sin mapping): **{}**".format(elementos_omitidos))
output.print_md("- Parámetros asignados correctamente: **{}**".format(asignados_ok))
output.print_md("- Parámetros con error: **{}**".format(asignados_fail))
output.print_md("✅ Proceso finalizado.")