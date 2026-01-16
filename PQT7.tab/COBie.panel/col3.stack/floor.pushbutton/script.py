# -*- coding: utf-8 -*-
__title__ = "COBie Floor"

from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter, BuiltInCategory, UnitUtils, UnitTypeId, StorageType
from pyrevit import script, revit
# Mantenemos tus importaciones para lo demás
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import get_param_value, GetParameterAPI, getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

doc = revit.doc
output = script.get_output()
logger = script.get_logger()

# ... (Tu código de validación inicial se mantiene igual) ...
nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo): script.exit()

schools_repositories = ColegiosRepository()
ob_school = schools_repositories.codigo_colegio(doc)
school = ob_school.name
created_by = ob_school.created_by
CREATED_ON = "2024-12-12T13:29:49"
parameters_static = { "COBie.CreatedBy": created_by, "COBie.CreatedOn": CREATED_ON }

logger.info("Inicio del script")

# Datos auxiliares
fec_basepoint = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
bp_elev = 0.0
if fec_basepoint:
    p_bp = fec_basepoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
    if p_bp: bp_elev = p_bp.AsDouble()

fec_info = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
info_site_type = ""
if fec_info:
    p_info = fec_info.LookupParameter("SiteObjectType")
    if p_info: info_site_type = p_info.AsString() or ""

# ==== Procesar Niveles ====
all_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
sorted_levels = sorted(all_levels, key=lambda lvl: lvl.Elevation)

valid_levels = []
skipped_levels = []

for lvl in sorted_levels:
    p_story = lvl.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY)
    if p_story and p_story.AsInteger() == 1:
        valid_levels.append(lvl)
    else:
        skipped_levels.append(lvl.Name)

processed_data = []

# ==== TRANSACTION ====
with revit.Transaction("Parametros COBie Floor"):
    total = len(valid_levels)
    
    for i in range(total):
        level = valid_levels[i]
        
        # 1. Calculo de Altura (Pies)
        if i < total - 1:
            raw_height = valid_levels[i+1].Elevation - level.Elevation
        else:
            raw_height = 0.0 # Ultimo nivel
            
        # 2. Categoría
        cat = "Floor"
        if "Sitio" in info_site_type: cat = "Site"
        elif "techo" in level.Name.lower(): cat = "Roof"
        
        # 3. Llenar parámetros Generales (Usando tu funcion helper)
        params = {
            "COBie.Floor.Name": level.Name,
            "COBie.Floor.Category": cat,
            "COBie.Floor.Elevation": bp_elev + level.Elevation
        }
        params.update(parameters_static)
        
        for k, v in params.items():
            p = getParameter(level, k)
            if p and not p.IsReadOnly: SetParameter(p, v)

        # ==========================================================
        # 4. ESCRITURA CRÍTICA (BYPASS DE LIBRERÍA EXTERNA)
        # ==========================================================
        # Usamos .LookupParameter directamente para evitar errores del helper
        p_height = level.LookupParameter("COBie.Floor.Height")
        
        debug_status = "NO ENCONTRADO"
        
        if p_height:
            if not p_height.IsReadOnly:
                # Verificamos tipo de almacenamiento real
                if p_height.StorageType == StorageType.Double:
                    # ESCRITURA DIRECTA DE DOUBLE
                    exito = p_height.Set(float(raw_height))
                    debug_status = "EXITO" if exito else "FALLO (Set devolvio False)"
                else:
                    debug_status = "ERROR TIPO (No es Longitud/Double)"
            else:
                debug_status = "READ ONLY"
        
        # Guardamos para reporte
        h_meters = UnitUtils.ConvertFromInternalUnits(raw_height, UnitTypeId.Meters)
        processed_data.append([level.Name, h_meters, debug_status])


# ==== REPORTE DE DIAGNÓSTICO ====
if processed_data:
    output.print_md("## Reporte de Ejecución")
    print("{:<20} | {:<10} | {:<30}".format("Nivel", "Altura(m)", "Estado Escritura"))
    print("-" * 70)
    for name, h, status in processed_data:
        print("{:<20} | {:<10.2f} | {:<30}".format(name, h, status))