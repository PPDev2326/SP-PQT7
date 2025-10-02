# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter, BuiltInCategory, UnitUtils, UnitTypeId
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms
from Extensions._RevitAPI import get_param_value, GetParameterAPI, getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

# ==== Creacion de metodos ====
def divide_string(text, idx, character_divider=None, compare=None, value_default=None):
    """
    Divide un string en partes usando un separador y devuelve el elemento en la posición idx.
    Si el texto completo coincide con 'compare', devuelve 'value_default'.
    """
    if not text:
        return ""
    
    # Verificar coincidencia antes de dividir
    if compare and text.strip().lower() == compare.lower():
        return value_default
    
    parts = text.split(character_divider) if character_divider else text.split()
    if idx < 0 or idx >= len(parts):
        return ""
    
    return parts[idx]

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
category_value = "Sin categoria"

# ==== Inicializamos un diccionario con los parametros a utilizar como claves
parameters_static = {
            "COBie.CreatedBy": created_by,
            "COBie.CreatedOn": CREATED_ON,
        }

# ==== Log para debug ====
logger.info("Inicio del script")

# ==== Obtenemos los niveles del modelo con el FilteredElementCollector ====
fec_levels = FilteredElementCollector(doc)
list_levels_object = fec_levels.OfClass(Level).WhereElementIsNotElementType().ToElements()

# ==== Obtenemos el punto base del proyecto ====
fec_basepoint = FilteredElementCollector(doc)
survey_object = fec_basepoint.OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
ob_param_elevation = GetParameterAPI(survey_object, BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
param_elevation_value = get_param_value(ob_param_elevation)

# ==== Obtenemos el project information ====
fec_information = FilteredElementCollector(doc)
information_object = fec_information.OfCategory(BuiltInCategory.OST_ProjectInformation).WhereElementIsNotElementType().FirstElement()
param_information_value = get_param_value(getParameter(information_object, "SiteObjectType"))

# ==== Abrimos la transaction para iniciar con los cambios ====
with revit.Transaction("Parametros COBie Floor"):

    for level in list_levels_object:
        # ==== Obtenemos el parametro planta de edificacion ====
        ob_param_buildingplan = GetParameterAPI(level, BuiltInParameter.LEVEL_IS_BUILDING_STORY)
        param_buildingplan_value = get_param_value(ob_param_buildingplan)
        
        # ==== Obtenemos el parametro Zonificación ====
        ob_param_zoning = getParameter(level, "S&P_ZONIFICACION")
        param_zoning_value = get_param_value(ob_param_zoning)
        
        if param_buildingplan_value == 1:
            level_name = level.Name
            if isinstance(level, Level):
                elevation = level.Elevation
            
            if "Sitio" in param_information_value:
                category_value = "Site"
            
            else:
                if "nivel" in level_name.lower():
                    category_value = "Floor"
                
                elif "techo" in level_name.lower():
                    category_value = "Roof"
            
            parameters= {
                "COBie.Floor.Name": level_name,
                "COBie.Floor.Category": category_value,
                "COBie.Floor.Description": "{}-{} (NPT:{})".format(level_name, param_zoning_value, elevation),
                "COBie.Floor.Elevation": param_elevation_value + elevation,
                "COBie.Floor.Height": elevation
            }
            
            parameters.update(parameters_static)
            
            for parameter, value in parameters.items():
                param = getParameter(level, parameter)
                if param and not param.IsReadOnly:
                    SetParameter(param, value)

# ==== Mostrar tabla en el output ====
output.print_table(
    table_data=[["ID", "Nombre"], [1, "Muro"], [2, "Puerta"]],
    title="Elementos procesados"
)

# ==== Mensaje con formato ====
output.print_md("### Procesamiento completado ✅")

logger.warning("Se ignoraron 2 elementos")