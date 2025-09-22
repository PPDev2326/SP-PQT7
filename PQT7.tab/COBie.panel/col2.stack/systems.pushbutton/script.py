# -*- coding: utf-8 -*-
__title__ = "COBie System"

from Autodesk.Revit.DB import Document, FilteredElementCollector, BuiltInCategory, FamilyInstance, StorageType, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms, script
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter
from Extensions._Dictionary import obtener_especialidad

uidoc = revit.uidoc
doc = revit.doc

especialidad = obtener_especialidad()
creado_por = "pruiz@cgeb.com.pe"
hora = "2025-04-18T16:45:10"
categoria = [
    "Ss_55_15 Water extraction, treatment and storage systems",
    "Ss_65_60 Specialist ventilation systems"
]

ductwork_system = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
pipe_system = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()

project_info = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
param = getParameter(project_info, "SiteObjectType")
param_activo = param.AsString() if param and param.StorageType == StorageType.String else None

activo = None
if param_activo:
    partes = param_activo.split()
    if len(partes) > 1:
        activo = partes[1]

parametros_estaticos = {
    "COBie.CreatedBy": creado_por,
    "COBie.CreatedOn": hora,
    "COBie.System.Category": categoria[1] if especialidad == "Instalaciones mecánicas"
                             else categoria[0] if especialidad == "Instalaciones sanitarias"
                             else ""
}

def procesar_sistema(elementos, nombre_sistema):
    resultado = {}
    tipos_sistema = {}
    
    for element in elementos:
        element_id = element.Id
        cobie = getParameter(element, "COBie")
        param = GetParameterAPI(element, BuiltInParameter.ELEM_TYPE_PARAM)
        param_tipo_sistema = param.AsValueString() if param and param.HasValue else "Sin valor"
        
        tipo = param_tipo_sistema if param_tipo_sistema else "Sin nombre"
        
        if tipo not in tipos_sistema:
            tipos_sistema[tipo] = 1
        else:
            tipos_sistema[tipo] += 1
        
        numero = tipos_sistema[tipo]

        parametros_dinamicos = {
            "COBie.System.Name": "{}_{} {:02}".format(activo, param_tipo_sistema, numero),
            "COBie.System.Description": "Sistema {}".format(param_tipo_sistema)
        }

        parametros_dinamicos.update(parametros_estaticos)

        if cobie:
            SetParameter(cobie, 1)

        for k, v in parametros_dinamicos.items():
            I_param_name = getParameter(element, k)
            SetParameter(I_param_name, v)

        resultado[element_id] = param_tipo_sistema

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