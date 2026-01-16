# -*- coding: utf-8 -*-
__title__ = "COBie Floor - NUCLEAR FIX"

from Autodesk.Revit.DB import FilteredElementCollector, Level, BuiltInParameter, BuiltInCategory, UnitUtils, UnitTypeId, StorageType
from pyrevit import script, revit
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import get_param_value, GetParameterAPI, getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

doc = revit.doc
output = script.get_output()
logger = script.get_logger()

# --- VALIDACIONES INICIALES ---
nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo): script.exit()

schools_repositories = ColegiosRepository()
ob_school = schools_repositories.codigo_colegio(doc)
school = ob_school.name
created_by = ob_school.created_by
CREATED_ON = "2024-12-12T13:29:49"
parameters_static = { "COBie.CreatedBy": created_by, "COBie.CreatedOn": CREATED_ON }

# --- PREPARACIÓN DE DATOS ---
# Base Point Elevation
fec_bp = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).FirstElement()
bp_elev = 0.0
if fec_bp:
    p = fec_bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
    if p: bp_elev = p.AsDouble()

# Site Type
fec_info = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
info_site = ""
if fec_info:
    p = fec_info.LookupParameter("SiteObjectType")
    if p: info_site = p.AsString() or ""

# Niveles
all_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
sorted_levels = sorted(all_levels, key=lambda l: l.Elevation)
valid_levels = [l for l in sorted_levels if l.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY).AsInteger() == 1]

processed_data = []

# =======================================================
#               INICIO DE TRANSACCIÓN
# =======================================================
with revit.Transaction("COBie Floor - Fix Nuclear"):
    total = len(valid_levels)
    
    for i in range(total):
        lvl = valid_levels[i]
        
        # 1. CALCULO DE ALTURA (PIES)
        if i < total - 1:
            raw_height = valid_levels[i+1].Elevation - lvl.Elevation
        else:
            raw_height = 0.0
            
        # 2. CATEGORIA
        cat = "Floor"
        if "Sitio" in info_site: cat = "Site"
        elif "techo" in lvl.Name.lower(): cat = "Roof"
        
        # 3. LLENADO PARAMETROS TEXTO
        params = {
            "COBie.Floor.Name": lvl.Name,
            "COBie.Floor.Category": cat,
            "COBie.Floor.Elevation": bp_elev + lvl.Elevation
        }
        params.update(parameters_static)
        for k, v in params.items():
            p = getParameter(lvl, k)
            if p and not p.IsReadOnly: SetParameter(p, v)
            
        # =======================================================
        # 4. ESCRITURA DE HEIGHT (METODO MULTIPLE)
        # =======================================================
        # Buscamos TODOS los parametros con ese nombre
        found_params = lvl.GetParameters("COBie.Floor.Height")
        
        write_status = "No encontrado"
        count_found = len(found_params)
        
        if count_found > 0:
            success_count = 0
            for param in found_params:
                if not param.IsReadOnly:
                    # Chequeo estricto de tipo
                    if param.StorageType == StorageType.Double:
                        check = param.Set(float(raw_height)) # Escribe PIES
                        if check: success_count += 1
                    elif param.StorageType == StorageType.String:
                        # Por si acaso tienes uno de texto colado
                        m_val = UnitUtils.ConvertFromInternalUnits(raw_height, UnitTypeId.Meters)
                        param.Set("{:.2f}".format(m_val))
            
            write_status = "Escrito en {} de {} parametros".format(success_count, count_found)
        
        # Guardar Log
        h_m = UnitUtils.ConvertFromInternalUnits(raw_height, UnitTypeId.Meters)
        processed_data.append([lvl.Id.IntegerValue, lvl.Name, h_m, write_status])

# --- REPORTE ---
if processed_data:
    output.print_md("## REPORTE DIAGNÓSTICO FINAL")
    print("{:<10} | {:<20} | {:<10} | {:<30}".format("ID", "Nivel", "Alt(m)", "Estado"))
    print("-" * 80)
    for pid, name, h, stat in processed_data:
        print("{:<10} | {:<20} | {:<10.2f} | {:<30}".format(pid, name, h, stat))