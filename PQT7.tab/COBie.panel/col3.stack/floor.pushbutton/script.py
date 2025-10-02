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

# ==== Variables para recopilar datos de salida ====
processed_levels = []
skipped_levels = []

last_elevation = None  # Para calcular diferencias de altura

# ==== Abrimos la transaction para iniciar con los cambios ====
with revit.Transaction("Parametros COBie Floor"):

    for level in list_levels_object:
        ob_param_buildingplan = GetParameterAPI(level, BuiltInParameter.LEVEL_IS_BUILDING_STORY)
        param_buildingplan_value = get_param_value(ob_param_buildingplan)
        
        ob_param_zoning = getParameter(level, "S&P_ZONIFICACION")
        param_zoning_value = get_param_value(ob_param_zoning)
        
        if param_buildingplan_value == 1:
            level_name = level.Name
            elevation = level.Elevation if isinstance(level, Level) else 0.0

            # Determinar categoría
            if "Sitio" in param_information_value:
                category_value = "Site"
            elif "nivel" in level_name.lower() or "piso" in level_name.lower():
                category_value = "Floor"
            elif "techo" in level_name.lower() or "cobertura" in level_name.lower():
                category_value = "Roof"
            else:
                category_value = "Sin categoria"

            # Calcular altura (Height)
            if last_elevation is None:
                floor_height = elevation  # Primer nivel: su elevación respecto al base
            else:
                floor_height = elevation - last_elevation  # Diferencia con el nivel anterior

            # Guardar elevación actual como referencia
            last_elevation = elevation

            parameters= {
                "COBie.Floor.Name": level_name,
                "COBie.Floor.Category": category_value,
                "COBie.Floor.Description": "{}-{} (NPT:{})".format(level_name, param_zoning_value, UnitUtils.ConvertFromInternalUnits(elevation, UnitTypeId.Meters)),
                "COBie.Floor.Elevation": param_elevation_value + elevation,
                "COBie.Floor.Height": floor_height
            }
            parameters.update(parameters_static)
            
            for parameter, value in parameters.items():
                param = getParameter(level, parameter)
                if param and not param.IsReadOnly:
                    SetParameter(param, value)

            processed_levels.append([level.Id.IntegerValue, level_name, category_value, round(floor_height, 2)])
        else:
            skipped_levels.append(level.Name)

# ==== Salidas profesionales ====
if processed_levels:
    output.print_md("## ✅ Procesamiento COBie Floor completado")
    output.print_md("**Niveles procesados:** {0}".format(len(processed_levels)))

    for level_id, level_name, category, height in processed_levels:
        output.print_md("- **{0}** | {1} | {2} | Altura: {3} m".format(level_id, level_name, category, height))

if skipped_levels:
    output.print_md("### ⚠️ Niveles ignorados (no son plantas de edificación):")
    for name in skipped_levels:
        output.print_md("- {0}".format(name))