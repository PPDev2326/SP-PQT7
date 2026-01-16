# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter, BuiltInCategory, UnitUtils, UnitTypeId, StorageType
from pyrevit import script, revit
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import get_param_value, GetParameterAPI, getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

# ==== obtenemos el documento y el uidocument del modelo activo ====
doc = revit.doc
uidoc = revit.uidoc

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

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

# ==== Inicializamos un diccionario con los parametros estaticos ====
parameters_static = {
    "COBie.CreatedBy": created_by,
    "COBie.CreatedOn": CREATED_ON,
}

# ==== Log para debug ====
logger.info("Inicio del script")

# ==== Obtenemos el punto base del proyecto ====
fec_basepoint = FilteredElementCollector(doc)
survey_object = fec_basepoint.OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
ob_param_elevation = GetParameterAPI(survey_object, BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
# Aseguramos que sea float
try:
    base_elevation_value = float(get_param_value(ob_param_elevation))
except:
    base_elevation_value = 0.0

# ==== Obtenemos el project information ====
fec_information = FilteredElementCollector(doc)
information_object = fec_information.OfCategory(BuiltInCategory.OST_ProjectInformation).WhereElementIsNotElementType().FirstElement()
param_information_value = get_param_value(getParameter(information_object, "SiteObjectType"))
if not param_information_value: param_information_value = ""

# ==== Obtenemos y Procesamos los Niveles ====
# 1. Obtener todos los niveles
all_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()

# 2. Ordenar por elevación (CRÍTICO para calcular alturas)
sorted_levels = sorted(all_levels, key=lambda lvl: lvl.Elevation)

# 3. Filtrar solo los que son "Building Story" para la lógica de pisos
valid_levels = []
skipped_levels_names = []

for lvl in sorted_levels:
    is_building_story = lvl.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY).AsInteger() == 1
    if is_building_story:
        valid_levels.append(lvl)
    else:
        skipped_levels_names.append(lvl.Name)

# ==== Variables para recopilar datos de salida ====
processed_levels_data = []

# ==== Abrimos la transaction para iniciar con los cambios ====
with revit.Transaction("Parametros COBie Floor"):

    total_levels = len(valid_levels)

    # Iteramos por índice para poder acceder al "siguiente" nivel
    for i in range(total_levels):
        level = valid_levels[i]
        level_name = level.Name
        elevation = level.Elevation # Esto está en PIES (Internal Units)
        
        # Obtener zonificación
        ob_param_zoning = getParameter(level, "S&P_ZONIFICACION")
        param_zoning_value = get_param_value(ob_param_zoning)
        if not param_zoning_value: param_zoning_value = ""

        # --- LÓGICA DE ALTURA (HEIGHT) ---
        # Si no es el último nivel, la altura es (Elevación Siguiente - Elevación Actual)
        if i < total_levels - 1:
            next_level = valid_levels[i+1]
            floor_height_feet = next_level.Elevation - elevation
        else:
            # Es el último nivel (azotea/techo), altura 0 o estandar
            floor_height_feet = 0.0

        # --- LÓGICA DE CATEGORÍA ---
        if "Sitio" in param_information_value:
            category_value = "Site"
        elif "nivel" in level_name.lower() or "piso" in level_name.lower():
            category_value = "Floor"
        elif "techo" in level_name.lower() or "cobertura" in level_name.lower():
            category_value = "Roof"
        else:
            category_value = "Sin categoria"

        # --- PREPARAR PARÁMETROS ---
        parameters = {
            "COBie.Floor.Name": level_name,
            "COBie.Floor.Category": category_value,
            "COBie.Floor.Description": "{}-{} (NPT:{:+.2f})".format(
                level_name, 
                param_zoning_value, 
                UnitUtils.ConvertFromInternalUnits(elevation, UnitTypeId.Meters)
            ),
            # Elevation es Longitud -> Pasamos Pies (Internal + Base)
            "COBie.Floor.Elevation": base_elevation_value + elevation,
        }
        parameters.update(parameters_static)
        
        # Llenamos los parámetros generales
        for parameter_name, value in parameters.items():
            param = getParameter(level, parameter_name)
            if param and not param.IsReadOnly:
                SetParameter(param, value)

        # --- ESCRITURA ESPECÍFICA PARA HEIGHT (LONGITUD/DOUBLE) ---
        param_height = getParameter(level, "COBie.Floor.Height")
        if param_height and not param_height.IsReadOnly:
            # Como es tipo Longitud, pasamos el float directo (PIES)
            # NO convertir a texto ni a metros aquí.
            param_height.Set(float(floor_height_feet))

        # --- DATOS PARA REPORTE (Visualización en Metros) ---
        height_m = UnitUtils.ConvertFromInternalUnits(floor_height_feet, UnitTypeId.Meters)
        processed_levels_data.append([level.Id.IntegerValue, level_name, category_value, round(height_m, 2)])

# ==== Salidas profesionales ====
if processed_levels_data:
    output.print_md("## ✅ Procesamiento COBie Floor completado")
    output.print_md("**Niveles procesados:** {0}".format(len(processed_levels_data)))

    for level_id, level_name, category, height in processed_levels_data:
        output.print_md("- **{0}** | {1} | {2} | Altura: {3} m".format(level_id, level_name, category, height))

if skipped_levels_names:
    output.print_md("### ⚠️ Niveles ignorados (no son plantas de edificación):")
    for name in skipped_levels_names:
        output.print_md("- {0}".format(name))