# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

# ========== Obtenemos las librerias necesarias ==========
from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter, BuiltInCategory, UnitUtils, UnitTypeId, StorageType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms

# Importamos tus modulos personalizados (Asumiendo que funcionan correctamente)
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

# ==== Inicializamos un diccionario con los parametros estáticos ====
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
# Aseguramos que sea float, a veces get_param_value devuelve string si no se controla
try:
    param_elevation_value = float(get_param_value(ob_param_elevation))
except:
    param_elevation_value = 0.0

# ==== Obtenemos el project information ====
fec_information = FilteredElementCollector(doc)
information_object = fec_information.OfCategory(BuiltInCategory.OST_ProjectInformation).WhereElementIsNotElementType().FirstElement()
param_information_value = get_param_value(getParameter(information_object, "SiteObjectType"))
if not param_information_value: 
    param_information_value = "" # Evitar error si es None

# ==== Obtenemos y filtramos los niveles ====
fec_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()

# Paso 1: Ordenar por elevación
raw_levels = sorted(fec_levels, key=lambda lvl: lvl.Elevation)

# Paso 2: Crear una lista limpia solo con niveles que son "Building Story" para calcular alturas reales
valid_levels = []
skipped_levels = []

for lvl in raw_levels:
    is_building_story = lvl.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY).AsInteger() == 1
    if is_building_story:
        valid_levels.append(lvl)
    else:
        skipped_levels.append(lvl.Name)

# ==== Variables para reporte ====
processed_levels_data = []

# ==== Abrimos la transaction ====
with revit.Transaction("Parametros COBie Floor"):
    
    # Iteramos usando índice para poder "mirar al siguiente nivel"
    total_valid = len(valid_levels)
    
    for i in range(total_valid):
        level = valid_levels[i]
        
        # Datos básicos
        level_name = level.Name
        elevation = level.Elevation
        
        # Obtener zonificación
        ob_param_zoning = getParameter(level, "S&P_ZONIFICACION")
        param_zoning_value = get_param_value(ob_param_zoning)
        if not param_zoning_value: param_zoning_value = ""

        # --- LÓGICA DE CATEGORÍA ---
        if "Sitio" in param_information_value:
            category_value = "Site"
        elif "nivel" in level_name.lower() or "piso" in level_name.lower():
            category_value = "Floor"
        elif "techo" in level_name.lower() or "cobertura" in level_name.lower():
            category_value = "Roof"
        else:
            category_value = "Sin categoria"

        # --- LÓGICA DE ALTURA (HEIGHT) ---
        # La altura es: Elevación del siguiente nivel - Elevación del actual
        if i < total_valid - 1:
            next_level = valid_levels[i+1]
            floor_height_internal = next_level.Elevation - elevation
        else:
            # Es el último nivel (azotea), asumimos 0 o una altura estándar
            floor_height_internal = 0.0 

        # --- PREPARAR PARÁMETROS GENERALES ---
        parameters = {
            "COBie.Floor.Name": level_name,
            "COBie.Floor.Category": category_value,
            "COBie.Floor.Description": "{}-{} (NPT:{:+.2f})".format(
                level_name, 
                param_zoning_value, 
                UnitUtils.ConvertFromInternalUnits(elevation, UnitTypeId.Meters)
            ),
            "COBie.Floor.Elevation": param_elevation_value + elevation,
        }
        parameters.update(parameters_static)

        # Escribir parámetros generales usando tu función helper
        for parameter_name, value in parameters.items():
            param = getParameter(level, parameter_name)
            if param and not param.IsReadOnly:
                SetParameter(param, value)

        # --- ESCRITURA ESPECÍFICA PARA HEIGHT (Solución del error) ---
        # Lo hacemos manual para asegurar compatibilidad de tipos
        param_height = getParameter(level, "COBie.Floor.Height")
        
        if param_height and not param_height.IsReadOnly:
            # Opción A: Es un parámetro de LONGITUD (Double) -> Revit espera pies
            if param_height.StorageType == StorageType.Double:
                param_height.Set(float(floor_height_internal))
            
            # Opción B: Es un parámetro de TEXTO (String) -> Revit espera texto en Metros
            elif param_height.StorageType == StorageType.String:
                val_meters = UnitUtils.ConvertFromInternalUnits(floor_height_internal, UnitTypeId.Meters)
                param_height.Set("{:.2f}".format(val_meters))
                
            # Opción C: Es Integer (Raro para altura, pero por si acaso)
            elif param_height.StorageType == StorageType.Integer:
                 param_height.Set(int(floor_height_internal))

        # Guardar datos para el reporte (Convertimos a metros para visualizar)
        height_m = UnitUtils.ConvertFromInternalUnits(floor_height_internal, UnitTypeId.Meters)
        processed_levels_data.append([level.Id.IntegerValue, level_name, category_value, round(height_m, 2)])


# ==== Salidas profesionales ====
if processed_levels_data:
    output.print_md("## ✅ Procesamiento COBie Floor completado")
    output.print_md("**Niveles procesados:** {0}".format(len(processed_levels_data)))

    for level_id, level_name, category, height in processed_levels_data:
        output.print_md("- **{0}** | {1} | {2} | Altura: {3} m".format(level_id, level_name, category, height))

if skipped_levels:
    output.print_md("### ⚠️ Niveles ignorados (no son plantas de edificación):")
    for name in skipped_levels:
        output.print_md("- {0}".format(name))