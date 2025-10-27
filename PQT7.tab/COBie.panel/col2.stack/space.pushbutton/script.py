# -*- coding: utf-8 -*-
__title__ = "COBie Space"
__doc__ = """
Version = 1.1
Date = 23.09.2025
------------------------------------------------------------------
Description:
Llenara los parametros del COBieSpace como complemento pedira el excel
Matriz COBie donde se encuentra ESTANDAR COBie SPACE.
------------------------------------------------------------------
¿Cómo hacerlo?
-> Click en el boton.
-> Esperamos que se complete la actualizacion de información.
------------------------------------------------------------------
Última actualización:
- [23.09.2025] - 1.1 UPDATE - New Feature
------------------------------------------------------------------
Autor: Paolo Perez
"""

from pyrevit import forms, script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    BuiltInParameter
)
from collections import defaultdict
from Extensions._RevitAPI import GetParameterAPI, getParameter, get_param_value
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Helper._Dictionary import get_formatted_string
from Helper._Excel import Excel
from DBRepositories.SchoolRepository import ColegiosRepository

output = script.get_output()

def set_param(param, valor, elemento, nombre_param):
    """Setea un parámetro y devuelve True si fue exitoso."""
    if not param:
        output.print_md("⚠️ No se encontró el parámetro **{}** en el elemento `{}`".format(
            nombre_param, elemento.Id.IntegerValue))
        return False
    if param.IsReadOnly:
        output.print_md("🔒 El parámetro **{}** es de solo lectura en el elemento `{}`".format(
            param.Definition.Name, elemento.Id.IntegerValue))
        return False
    try:
        param.Set(valor)
        output.print_md("✅ Asignado **{} = {}** al elemento `{}`".format(
            nombre_param, valor, elemento.Id.IntegerValue))
        return True
    except Exception as e:
        output.print_md("❌ Error al asignar **{}** en el elemento `{}`: {}".format(
            nombre_param, elemento.Id.IntegerValue, e))
        return False

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

doc = revit.doc

# ==== Instanciamos el colegio correspondiente del modelo activo ====
school_repo_object = ColegiosRepository()
school_object = school_repo_object.codigo_colegio(doc)
school = None
created_by = None

# ==== Verificamos si el colegio existe
if school_object:
    school = school_object.name
    created_by = school_object.created_by

# Datos del usuario
FECHA = "2023-04-07T16:38:56"

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
output.print_md("### 🔍 Iniciando transferencia COBie")
output.print_md("- Rooms encontrados: **{}**".format(len(rooms)))
output.print_md("- Spaces encontrados: **{}**".format(len(spaces)))
output.print_md("- Total de elementos a procesar: **{}**".format(len(elementos)))

abreviaciones = defaultdict(int)
asignados_ok = 0
asignados_fail = 0
elementos_procesados = 0
elementos_omitidos = 0

# Columna requeridas
columns_space = [
    "COBie.Space.Name",
    "COBie.Space.RoomTag",
    "Classification.Space.Number",
    "Classification.Space.Description"
]

# Crear instancia de Excel (esto pedirá el archivo UNA SOLA VEZ)
excel_instance = Excel()

# Cargar datos de SPACE (del mismo archivo Excel ya abierto - NO PEDIRÁ EL ARCHIVO DE NUEVO)
print("[INFO] Cargando hoja 'ESTANDAR COBie SPACE' del mismo archivo...")
space_rows = excel_instance.read_excel('ESTANDAR COBie SPACE ')
if not space_rows:
    forms.alert("No se encontró hoja de Excel 'ESTANDAR COBie SPACE'.", exitscript=True)

space_headers = excel_instance.get_headers(space_rows, 2)
space_headers_required = excel_instance.headers_required(space_headers, columns_space)
space_data = excel_instance.get_data_by_headers_required(space_rows, space_headers_required, 3)

print(space_data)

with revit.Transaction("Transfiere datos a Parametros COBieSpace"):
    for i, elemento in enumerate(elementos, start=1):
        output.print_md("---")
        output.print_md("➡️ Procesando elemento `{}` ({}/{})".format(
            elemento.Id.IntegerValue, i, len(elementos)))

        room_param_name = GetParameterAPI(elemento, BuiltInParameter.ROOM_NAME)
        room_param_number = GetParameterAPI(elemento, BuiltInParameter.ROOM_NUMBER)
        cl_param_description = getParameter(elemento, "Classification.Space.Description")
        cl_param_number = getParameter(elemento, "Classification.Space.Number")
        
        value_room_name = get_param_value(room_param_name, "Sin nombre")
        value_room_number = get_param_value(room_param_number, "Sin numero")
        value_cl_description = get_param_value(cl_param_description, "Sin nombre")
        value_cl_number = get_param_value(cl_param_number, "Sin nombre")
        # Nombre clave del elemento de Revit (Ej: "101 : PASILLO")
        name_full = get_formatted_string(value_room_number, value_room_name)
        
        
        # 🟢 BÚSQUEDA: Buscar la fila en space_data donde la columna 'COBie.Space.Name' coincide con name_full
        fila_excel = next((
            row for row in space_data 
            if row.get("COBie.Space.Name") == name_full
        ), None) # Retorna el diccionario de la fila si lo encuentra, sino None

        
        if fila_excel:
            # ✅ El elemento se encontró en los datos de Excel.
            elementos_procesados += 1
            output.print_md("✅ Elemento encontrado en datos de Excel: **{}**".format(name_full))
            
            # ----------------------------------------------------
            # 1. VERIFICACIÓN Y ASIGNACIÓN DE CLASIFICACIÓN
            # ----------------------------------------------------
            
            # Obtener los parámetros de Classification del elemento de Revit
            cl_param_number_revit = elemento.LookupParameter("Classification.Space.Number")
            cl_param_description_revit = elemento.LookupParameter("Classification.Space.Description")
            
            # Extraer valores del EXCEL
            cl_number_excel = fila_excel.get("Classification.Space.Number")
            cl_description_excel = fila_excel.get("Classification.Space.Description")
            
            # Lógica de verificación para Classification.Space.Number
            # Leemos el valor actual en Revit. Usamos get_param_value para manejar None/null/vacío.
            is_number_revit_empty = not get_param_value(cl_param_number_revit, "")
            
            if is_number_revit_empty: 
                if cl_number_excel: # Si Revit está vacío Y Excel tiene valor
                    set_param(cl_param_number_revit, cl_number_excel, elemento, "Classification.Space.Number")
                    final_cl_number = cl_number_excel # ⬅️ USAMOS EL VALOR DEL EXCEL
                else:
                    output.print_md("ℹ️ Valor de 'Classification.Space.Number' de Excel vacío. Omitido.")
                    final_cl_number = value_cl_number # ⬅️ USAMOS EL VALOR ORIGINAL DE REVIT
            else:
                output.print_md("🔒 'Classification.Space.Number' ya tiene un valor. Omitido.")
                final_cl_number = value_cl_number # ⬅️ USAMOS EL VALOR ORIGINAL DE REVIT

            # Lógica de verificación para Classification.Space.Description
            is_description_revit_empty = not get_param_value(cl_param_description_revit, "")
            
            if is_description_revit_empty: 
                if cl_description_excel: # Si Revit está vacío Y Excel tiene valor
                    set_param(cl_param_description_revit, cl_description_excel, elemento, "Classification.Space.Description")
                    final_cl_description = cl_description_excel # ⬅️ USAMOS EL VALOR DEL EXCEL
                else:
                    output.print_md("ℹ️ Valor de 'Classification.Space.Description' de Excel vacío. Omitido.")
                    final_cl_description = value_cl_description # ⬅️ USAMOS EL VALOR ORIGINAL DE REVIT
            else:
                output.print_md("🔒 'Classification.Space.Description' ya tiene un valor. Omitido.")
                final_cl_description = value_cl_description # ⬅️ USAMOS EL VALOR ORIGINAL DE REVIT
            
            
            # ----------------------------------------------------
            # 2. GENERAR CATEGORY CON LOS VALORES FINALES
            # ----------------------------------------------------
            
            # Generar CATEGORY con los valores finales (siempre usando el valor que DEBERÍA estar en el parámetro)
            categoria = get_formatted_string(final_cl_number, final_cl_description)
            
            # Extraer el RoomTag (o cualquier otra columna) directamente de la fila de Excel
            room_tag = fila_excel.get("COBie.Space.RoomTag")

            # --- OBTENCIÓN DE DATOS DE REVIT (ÁREA/ALTURA) ---
            height = GetParameterAPI(elemento, BuiltInParameter.ROOM_UPPER_OFFSET)
            area = GetParameterAPI(elemento, BuiltInParameter.ROOM_AREA)

            height_val = get_param_value(height, 0)
            area_val = get_param_value(area, 0)

            # ----------------------------------------------------
            # 3. ASIGNACIÓN DEL RESTO DE PARÁMETROS 
            # ----------------------------------------------------
            
            parametros = {
                "COBie.Space.Name": name_full,
                "COBie.CreatedBy": created_by,
                "COBie.CreatedOn": FECHA,
                "COBie.Space.Category": categoria, # ⬅️ USA EL VALOR GENERADO CORRECTAMENTE
                "COBie.Space.Description": value_room_name,
                "COBie.Space.RoomTag": room_tag,
                "COBie.Space.UsableHeight": height_val,
                "COBie.Space.GrossArea": area_val,
                "COBie.Space.NetArea": area_val,
                "COBie": 1,
            }

            for nombre_param, valor in parametros.items():
                if set_param(elemento.LookupParameter(nombre_param), valor, elemento, nombre_param):
                    asignados_ok += 1
                else:
                    asignados_fail += 1
        else:
            # ⚠️ El elemento NO se encontró en los datos de Excel.
            elementos_omitidos += 1
            output.print_md("⚠️ Elemento omitido (no encontrado en datos de Excel): **{}**".format(name_full))

output.print_md("---")

output.print_md("---")
output.print_md("### 📊 Resumen final")
output.print_md("- Total elementos analizados: **{}**".format(len(elementos)))
output.print_md("- Elementos procesados (con mapping): **{}**".format(elementos_procesados))
output.print_md("- Elementos omitidos (sin mapping): **{}**".format(elementos_omitidos))
output.print_md("- Parámetros asignados correctamente: **{}**".format(asignados_ok))
output.print_md("- Parámetros con error: **{}**".format(asignados_fail))
output.print_md("✅ Proceso finalizado.")