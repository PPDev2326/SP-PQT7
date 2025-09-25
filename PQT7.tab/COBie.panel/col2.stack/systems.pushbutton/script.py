# -*- coding: utf-8 -*-
__title__ = "COBie System"

from Autodesk.Revit.DB import (
    Document, FilteredElementCollector, BuiltInCategory,
    FamilyInstance, StorageType, BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms, script
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository


uidoc = revit.uidoc
doc = revit.doc


def divide_string(text, idx, character_divider=None):
    """
    Divide un string en partes usando un separador y devuelve el elemento en la posición idx.
    """
    if not text:
        return ""
    parts = text.split(character_divider) if character_divider else text.split()
    if idx < 0 or idx >= len(parts):
        return ""
    return parts[idx]


specialty_instance = SpecialtiesRepository()
specialty_object = specialty_instance.get_specialty_by_document(doc)
specialty = specialty_object.name

if specialty in ["INSTALACIONES SANITARIAS", "INSTALACIONES MECANICAS"]:

    # ==== Variables constantes ====
    CREATED_BY = "jtiburcio@syp.com.pe"
    CREATED_ON = "2025-04-18T16:45:10"

    # ==== Categoria de COBie.System.Category ====
    categoria = {
        "INSTALACIONES SANITARIAS": "Ss_55_15 Water extraction, treatment and storage systems",
        "INSTALACIONES MECANICAS": "Ss_65_60 Specialist ventilation systems"
    }

    ductwork_system = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
    pipe_system = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()

    project_info = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ProjectInformation).FirstElement()
    param = getParameter(project_info, "SiteObjectType")
    param_activo = param.AsString() if (param and param.StorageType == StorageType.String) else None
    activo = divide_string(param_activo, 1)

    parametros_estaticos = {
        "COBie.CreatedBy": CREATED_BY,
        "COBie.CreatedOn": CREATED_ON,
        "COBie.System.Category": categoria.get(specialty, "")
    }

    def procesar_sistema(elementos, active):
        resultado = {}
        tipos_sistema = {}

        for element in elementos:
            element_id = element.Id
            cobie_param = getParameter(element, "COBie")

            # Tipo de sistema
            param_type = GetParameterAPI(element, BuiltInParameter.ELEM_TYPE_PARAM)
            param_tipo_sistema = param_type.AsValueString() if (param_type and param_type.HasValue) else "Sin valor"

            # Nombre del sistema
            param_name_system = GetParameterAPI(element, BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
            name_system_value = param_name_system.AsString() if (param_name_system and param_name_system.HasValue) else "Sin valor"

            # Para nombrado: segunda parte si hay "-", si no queda vacío
            name_system_divide = divide_string(param_tipo_sistema, 1, "-")

            tipo = param_tipo_sistema if param_tipo_sistema else "Sin nombre"

            # Contador de ocurrencias por tipo
            if tipo not in tipos_sistema:
                tipos_sistema[tipo] = 1
            else:
                tipos_sistema[tipo] += 1

            contador = tipos_sistema[tipo]
            sufijo = str(contador).zfill(2)  # 01, 02, 03...

            system_name = "{} - {} {}".format(
                active.strip(),
                name_system_divide if name_system_divide else tipo,
                sufijo
            ).strip()

            parametros_dinamicos = {
                "COBie.System.Name": system_name,
                "COBie.System.Description": "Sistema {}".format(name_system_value)
            }
            parametros_dinamicos.update(parametros_estaticos)

            # Escribir parámetros
            if cobie_param:
                SetParameter(cobie_param, 1)

            for k, v in parametros_dinamicos.items():
                I_param = getParameter(element, k)
                if I_param:
                    SetParameter(I_param, v)

            resultado[element_id.IntegerValue] = param_tipo_sistema

        return resultado

    with revit.Transaction("Completar COBie System"):
        ductos = procesar_sistema(ductwork_system, "Ductos")
        tuberias = procesar_sistema(pipe_system, "Tuberías")

    output = script.get_output()
    if ductos or tuberias:
        output.print_md("## Se completó satisfactoriamente")

    if ductos:
        output.print_md("### Sistema de Ductos")
        for k, v in ductos.items():
            output.print_md("- `{}` → `{}`".format(k, v))

    if tuberias:
        output.print_md("### Sistema de Tuberías")
        for k, v in tuberias.items():
            output.print_md("- `{}` → `{}`".format(k, v))

else:
    script.exit()
