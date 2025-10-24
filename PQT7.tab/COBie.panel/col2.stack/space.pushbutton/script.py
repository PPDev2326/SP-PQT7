# -*- coding: utf-8 -*-
__title__ = "COBie Space"
__doc__ = """
Version = 1.1
Date = 23.09.2025
------------------------------------------------------------------
Description:
Llenara los parametros del COBieSpace.
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

# Datos del usuario
CORREO = "jtiburcio@syp.com.pe"
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
            if not get_param_value(cl_param_number_revit, ""): # Verifica si el parámetro de Revit está vacío
                if cl_number_excel: # Verifica que el valor del Excel no esté vacío
                    set_param(cl_param_number_revit, cl_number_excel, elemento, "Classification.Space.Number")
                else:
                    output.print_md("ℹ️ Valor de 'Classification.Space.Number' de Excel vacío. Omitido.")
            else:
                output.print_md("🔒 'Classification.Space.Number' ya tiene un valor. Omitido.")

            # Lógica de verificación para Classification.Space.Description
            if not get_param_value(cl_param_description_revit, ""): # Verifica si el parámetro de Revit está vacío
                if cl_description_excel: # Verifica que el valor del Excel no esté vacío
                    set_param(cl_param_description_revit, cl_description_excel, elemento, "Classification.Space.Description")
                else:
                    output.print_md("ℹ️ Valor de 'Classification.Space.Description' de Excel vacío. Omitido.")
            else:
                output.print_md("🔒 'Classification.Space.Description' ya tiene un valor. Omitido.")

            
            # --- OBTENCIÓN DE DATOS NECESARIOS PARA CATEGORY (POSTERIOR A LA ASIGNACIÓN) ---
            
            # Debemos RE-LEER los valores después del SET para asegurar que 'categoria' usa el nuevo valor.
            # Sin embargo, como estamos en la misma transacción, los valores de getParameter no se actualizarán
            # inmediatamente. La forma más segura es usar los valores del EXCEL para generar la categoría
            # si los parámetros en Revit estaban vacíos.
            
            # Usaremos los valores del Excel si los parámetros de Revit estaban vacíos (basado en la lógica anterior)
            # Si el parámetro de Revit NO estaba vacío, value_cl_number y value_cl_description 
            # ya contienen el valor original de Revit (leído antes del loop), lo cual es lo que necesitamos.
            
            # Si se asignaron valores de Excel, usamos esos valores para generar la categoría
            if not get_param_value(cl_param_number_revit, ""):
                final_cl_number = cl_number_excel
            else:
                final_cl_number = value_cl_number
                
            if not get_param_value(cl_param_description_revit, ""):
                final_cl_description = cl_description_excel
            else:
                final_cl_description = value_cl_description
            
            # Generar CATEGORY con los valores finales (originales de Revit o asignados del Excel)
            categoria = get_formatted_string(final_cl_number, final_cl_description)
            
            # Extraer el RoomTag (o cualquier otra columna) directamente de la fila de Excel
            room_tag = fila_excel.get("COBie.Space.RoomTag")

            # --- OBTENCIÓN DE DATOS DE REVIT (ÁREA/ALTURA) ---
            height = GetParameterAPI(elemento, BuiltInParameter.ROOM_UPPER_OFFSET)
            area = GetParameterAPI(elemento, BuiltInParameter.ROOM_AREA)

            height_val = get_param_value(height, 0)
            area_val = get_param_value(area, 0)

            # ----------------------------------------------------
            # 2. ASIGNACIÓN DEL RESTO DE PARÁMETROS (incluyendo Category)
            # ----------------------------------------------------
            
            parametros = {
                "COBie.Space.Name": name_full,
                "COBie.CreatedBy": CORREO,
                "COBie.CreatedOn": FECHA,
                "COBie.Space.Category": categoria, # ⬅️ AHORA USA EL VALOR GENERADO
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