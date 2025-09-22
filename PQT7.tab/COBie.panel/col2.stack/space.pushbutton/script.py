# -*- coding: utf-8 -*-
__title__ = "COBie Space"
__doc__ = """Version = 1.3
Date = 06.05.2025
------------------------------------------------------------------
Description:
Transfiere datos Space y Room a parámetros COBie.
------------------------------------------------------------------
Author: Paolo Perez"""

from pyrevit import forms, script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    BuiltInParameter
)
from collections import defaultdict
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._Nombre import (
    obtener_nombre_corto, generar_abreviacion,
    obtener_nombre_base_para_contador,
    capitalizar_respetando_mayusculas
)

def set_param(param, valor):
    if param and not param.IsReadOnly:
        try:
            param.Set(valor)
        except Exception:
            pass

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

doc = __revit__.ActiveUIDocument.Document

# Datos del usuario
correo = "pruiz@cgeb.com.pe"
fecha = "2025-04-18T16:45:10"
categoria = forms.ask_for_one_item(
    ["SL_50_20 Categoria ejemplo", "SL_25_10 Educational spaces"],
    default="SL_50_20 Categoria ejemplo",
    prompt="Seleccione una categoria Uniclass",
    title="Agregar a COBie.Space.Category"
)

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


abreviaciones = defaultdict(int)

with revit.Transaction("Transfiere datos a Parametros COBieSpace"):
    for i, elemento in enumerate(elementos, start=1):
        name_param = elemento.get_Parameter(BuiltInParameter.ROOM_NAME)
        if not name_param:
            continue

        name = name_param.AsString() or ""
        name_cap = capitalizar_respetando_mayusculas(name)
        name_full = "{:02}_{}".format(i, name_cap)
        nombre_corto = obtener_nombre_corto(name)
        nombre_base = obtener_nombre_base_para_contador(nombre_corto)
        abreviaciones[nombre_base] += 1
        abreviacion = generar_abreviacion(nombre_base, abreviaciones[nombre_base])

        height = elemento.get_Parameter(BuiltInParameter.ROOM_HEIGHT)
        area = elemento.get_Parameter(BuiltInParameter.ROOM_AREA)

        height_val = height.AsDouble() if height else 0
        area_val = area.AsDouble() if area else 0

        parametros = {
            "COBie.Space.Name": name_full,
            "COBie.CreatedBy": correo,
            "COBie.CreatedOn": fecha,
            "COBie.Space.Category": categoria,
            "COBie.Space.Description": nombre_corto,
            "COBie.Space.RoomTag": abreviacion,
            "COBie.Space.UsableHeight": height_val,
            "COBie.Space.GrossArea": area_val,
            "COBie.Space.NetArea": area_val,
            "COBie": 1,
        }

        for nombre_param, valor in parametros.items():
            set_param(elemento.LookupParameter(nombre_param), valor)
