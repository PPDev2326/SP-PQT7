# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms
from Extensions._RevitAPI import get_param_value, GetParameterAPI, getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

# ==== obtenemos el documento y el uidocument del modelo activo ====
doc = revit.doc
uidoc = revit.uidoc

# ==== Instanciamos el colegio correspondiente al modelo actual ====
schools_repositories = ColegiosRepository()
ob_school = schools_repositories.codigo_colegio(doc)
school = ob_school.name
created_by = ob_school.created_by

# ==== Instanciamos la salida output y logger ====
output = script.get_output()
logger = script.get_logger()

# ==== variables estaticas ====
CREATED_ON = "2024-12-12T13:29:49"

# ==== Variables ====
floor_category = {
    "Floor": "NIVEL", 
    "Roof": ["Techo", "Cobertura"], 
    "Site": "Sitio"
}

# ==== Inicializamos un diccionario con los parametros a utilizar como claves
parameters_static = {
            "COBie.CreatedBy": created_by,
            "COBie.CreatedOn": CREATED_ON,
            "COBie.Floor.Category": "",
            "COBie.Floor.Description": "",
            "COBie.Floor.Elevation": "",
            "COBie.Floor.Height": ""
        }

# ==== Log para debug ====
logger.info("Inicio del script")

# ==== Obtenemos los niveles del modelo con el FilteredElementCollector ====
fec = FilteredElementCollector(doc)
list_levels_object = fec.OfClass(Level).WhereElementIsNotElementType().ToElements()

list_level_name = []

for level in list_levels_object:
    ob_param_buildingplan = GetParameterAPI(level, BuiltInParameter.LEVEL_IS_BUILDING_STORY)
    param_buildingplan_value = get_param_value(ob_param_buildingplan)
    
    if param_buildingplan_value == 1:
        level_name = level.Name
        list_level_name.append(level_name)
        
        if level_name in floor_category:
            floot_category_value = floor_category.get()
        
        parameters= {
            "COBie.Floor.Name": level_name,
            "COBie.Floor.Category": "",
            "COBie.Floor.Description": "",
            "COBie.Floor.Elevation": "",
            "COBie.Floor.Height": ""
        }

# ==== Mostrar tabla en el output ====
output.print_table(
    table_data=[["ID", "Nombre"], [1, "Muro"], [2, "Puerta"]],
    title="Elementos procesados"
)

# ==== Mensaje con formato ====
output.print_md("### Procesamiento completado ✅")

logger.warning("Se ignoraron 2 elementos")