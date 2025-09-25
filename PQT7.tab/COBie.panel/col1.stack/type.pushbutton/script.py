# -*- coding: utf-8 -*-
__title__ = "COBie Type"

import re
from Autodesk.Revit.DB import BuiltInParameter, StorageType, UnitUtils, UnitTypeId, FamilyInstance, ElementType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository
from DBRepositories.SchoolRepository import ColegiosRepository
from Helper._Excel import Excel

parametros_cobie = [
        "COBie.Type.Manufacturer",
        "COBie.Type.ModelNumber", 
        "COBie.Type.WarrantyDurationParts",
        "COBie.Type.WarrantyDurationLabor",
        "COBie.Type.ReplacementCost",
        "COBie.Type.ExpectedLife",
        "COBie.Type.NominalLength",
        "COBie.Type.NominalWidth",
        "COBie.Type.NominalHeight",
        "COBie.Type.Color",
        "COBie.Type.Finish",
        "COBie.Type.Constituents",
        "CODIGO"  # Columna identificadora
    ]

def extraer_medida(tipo_name):
    """
    Extrae el contenido que está dentro de paréntesis de un string
    
    Args:
        tipo_name: String que contiene texto con paréntesis
        
    Returns:
        String con el contenido dentro de paréntesis, o None si no encuentra
    """
    if not tipo_name:
        return None
    match = re.search(r'\(([^)]+)\)', tipo_name)
    return match.group(1) if match else None

uidoc = revit.uidoc
doc = revit.doc

# ==== Obtenemos la especialidad del modelo activo y sus datos ====
repo_specialties = SpecialtiesRepository()
specialty_object = repo_specialties.get_specialty_by_document(doc)
specialty = None
sp_accesibility = None
sp_code = None
sp_sustainability = None
sp_feature = None

if specialty_object:
    specialty = specialty_object.name
    sp_accesibility = specialty_object.accessibility_performance
    sp_code = specialty_object.code_perfomance
    sp_sustainability = specialty_object.sustainability
    sp_feature = specialty_object.feature

# ==== Obtenemos el colegio correspondiente segun modelo y sus datos necesarios ====
repo_schools = ColegiosRepository()
school_object = repo_schools.codigo_colegio(doc)
school = None
created_by_value = None
warranty_description_value = None

if school_object:
    school = school_object.name
    created_by_value = school_object.created_by
    warranty_description_value = school_object.warranty_description

# Validación de datos críticos
if not created_by_value or not warranty_description_value:
    forms.alert("No se pudieron obtener los datos del colegio necesarios.", exitscript=True)

# ==== Segun especialidad ====
if specialty in ["ARQUITECTURA", "ESTRUCTURAS"]:
    asset_type = "Semi-fixed"
else:
    asset_type = "Fixed"

# ==== Datos estáticos ====
CREATED_ON = "2025-08-04T11:59:30"
DURATION_UNIT = "AÑO"
SHAPE = "Poligonal"
GRADE = "Grado Estándar"

parameters_static = {
    "COBie.Type.CreatedBy": created_by_value,
    "COBie.Type.CreatedOn": CREATED_ON,
    "COBie.Type.AssetType": asset_type,
    "COBie.Type.WarrantyGuarantorParts": created_by_value,       
    "COBie.Type.WarrantyGuarantorLabor": created_by_value,
    "COBie.Type.WarrantyDurationUnit": DURATION_UNIT,
    "COBie.Type.DurationUnit": DURATION_UNIT,
    "COBie.Type.WarrantyDescription": warranty_description_value,
    "COBie.Type.ModelReference": warranty_description_value,
    "COBie.Type.Shape": SHAPE,
    "COBie.Type.Grade": GRADE,
    "COBie.Type.Features": sp_feature,
    "COBie.Type.AccessibilityPerformance": sp_accesibility,
    "COBie.Type.CodePerformance": sp_code,
    "COBie.Type.SustainabilityPerformance": sp_sustainability,
}

# ==== Selección y preparación ====
try:
    selection = uidoc.Selection.PickElementsByRectangle()
    if not selection:
        forms.alert("No se seleccionaron elementos.", exitscript=True)
except Exception as e:
    forms.alert("No se seleccionaron elementos o se produjo un error:\n\n" + str(e), exitscript=True)

# ==== Obtenemos la hoja excel de acuerdo a la especialidad ====
data_headers = None

if specialty == "ARQUITECTURA":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -AR')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    # Convertir lista de diccionarios a diccionario de listas para compatibilidad
    data_headers = {col: [row[col] for row in data_list] for col in parametros_cobie}
    print("Datos de Arquitectura cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES SANITARIAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - PL')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    data_headers = {col: [row[col] for row in data_list] for col in parametros_cobie}
    print("Datos de Sanitarias cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES ELECTRICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -EE')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    data_headers = {col: [row[col] for row in data_list] for col in parametros_cobie}
    print("Datos de Eléctricas cargados:", len(data_list), "filas")

elif specialty == "COMUNICACIONES":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - IICC')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    data_headers = {col: [row[col] for row in data_list] for col in parametros_cobie}
    print("Datos de Comunicaciones cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES MECANICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - ME')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    data_headers = {col: [row[col] for row in data_list] for col in parametros_cobie}
    print("Datos de Mecánicas cargados:", len(data_list), "filas")

else:
    forms.alert("Especialidad '{}' no reconocida para cargar datos Excel.".format(specialty), exitscript=True)

if not data_headers:
    forms.alert("No se pudieron cargar los datos del Excel.", exitscript=True)

# ==== Mapeo de parámetros Excel -> Revit ====
param_mapping = {
    "CODIGO": "S&P_CODIGO DE ELEMENTO",  # Mapeo especial para comparación
    # Los demás parámetros tienen el mismo nombre
    "COBie.Type.Manufacturer": "COBie.Type.Manufacturer",
    "COBie.Type.ModelNumber": "COBie.Type.ModelNumber",
    "COBie.Type.WarrantyDurationParts": "COBie.Type.WarrantyDurationParts",
    "COBie.Type.WarrantyDurationLabor": "COBie.Type.WarrantyDurationLabor",
    "COBie.Type.ReplacementCost": "COBie.Type.ReplacementCost",
    "COBie.Type.ExpectedLife": "COBie.Type.ExpectedLife",
    "COBie.Type.NominalLength": "COBie.Type.NominalLength",
    "COBie.Type.NominalWidth": "COBie.Type.NominalWidth",
    "COBie.Type.NominalHeight": "COBie.Type.NominalHeight",
    "COBie.Type.Color": "COBie.Type.Color",
    "COBie.Type.Finish": "COBie.Type.Finish",
    "COBie.Type.Constituents": "COBie.Type.Constituents"
}

# ==== Función para buscar datos del Excel por código ====
def buscar_datos_por_codigo(data_headers, codigo_elemento):
    """
    Busca los datos del Excel que coincidan con el código del elemento.
    
    :param data_headers: Datos del Excel organizados por headers {columna: [valores]}.
    :type data_headers: dict
    :param codigo_elemento: Código del elemento de Revit.
    :type codigo_elemento: str
    :return: Diccionario con los datos encontrados o None si no encuentra.
    :rtype: dict or None
    """
    if not data_headers or "CODIGO" not in data_headers:
        return None
    
    # Buscar en qué fila está el código
    codigos_excel = data_headers["CODIGO"]
    for row_index, codigo_excel in enumerate(codigos_excel):
        if codigo_excel and str(codigo_excel).strip() == str(codigo_elemento).strip():
            # Encontrado, extraer todos los datos de esa fila
            datos_encontrados = {}
            for header, values in data_headers.items():
                if row_index < len(values):
                    datos_encontrados[header] = values[row_index]
                else:
                    datos_encontrados[header] = None
            return datos_encontrados
    
    return None

element_types = dict()  # type_id_int -> element_type

for element in selection:
    # Procesar el tipo del elemento principal
    type_elem = doc.GetElement(element.GetTypeId())
    if type_elem:
        element_types[type_elem.Id.IntegerValue] = type_elem

    # Procesar tipos de subcomponentes (si existen)
    if isinstance(element, FamilyInstance):
        try:
            for sc_id in element.GetSubComponentIds():
                sub = doc.GetElement(sc_id)
                if sub:
                    type_sub = doc.GetElement(sub.GetTypeId())
                    if type_sub:
                        element_types[type_sub.Id.IntegerValue] = type_sub
        except Exception as e:
            print("Error procesando subcomponentes: " + str(e))

if not element_types:
    forms.alert("No se encontraron tipos de elementos válidos en la selección.", exitscript=True)

# ==== Proceso COBie.Type optimizado ====
conteo = 0
elementos_omitidos = 0

with revit.Transaction("Transferencia COBie Type con Excel"):
    for type_id, element_type in element_types.items():
        
        # Verificar si el elemento debe procesarse
        param_cobie_type = getParameter(element_type, "COBie.Type")
        if not (param_cobie_type and param_cobie_type.StorageType == StorageType.Integer and param_cobie_type.AsInteger() == 1):
            elementos_omitidos += 1
            continue
        
        try:
            # ==== Obtener el código del elemento ====
            codigo_elemento = None
            param_codigo = getParameter(element_type, "S&P_CODIGO DE ELEMENTO")
            if param_codigo and param_codigo.HasValue:
                codigo_elemento = param_codigo.AsString()
            
            if not codigo_elemento:
                print("Elemento tipo {} no tiene código, se omite".format(element_type.Id))
                elementos_omitidos += 1
                continue
            
            # ==== Buscar datos en Excel ====
            datos_excel = buscar_datos_por_codigo(data_headers, codigo_elemento)
            
            if not datos_excel:
                print("No se encontraron datos en Excel para código: {}".format(codigo_elemento))
                elementos_omitidos += 1
                continue
            
            print("Procesando elemento con código: {}".format(codigo_elemento))
            
            # ==== Obtener datos del elemento ====
            category_object = element_type.Category
            category_name = category_object.Name if category_object else "Sin Categoría"
            
            fam_name = "Sin Familia"
            if isinstance(element_type, ElementType):
                fam_name = element_type.FamilyName
            
            param_name_value = "Sin Nombre"
            object_param_name = GetParameterAPI(element_type, BuiltInParameter.SYMBOL_NAME_PARAM)
            if object_param_name:
                param_name_value = object_param_name.AsString() or "Sin Nombre"
            
            param_desc_value = "Sin Descripción"
            object_param_desc = getParameter(element_type, "Descripción")
            if object_param_desc and object_param_desc.HasValue:
                param_desc_value = object_param_desc.AsString() or "Sin Descripción"
            
            param_name_material = "Sin Material"
            object_param_material = getParameter(element_type, "S&P_MATERIAL DE ELEMENTO")
            if object_param_material and object_param_material.HasValue:
                param_name_material = object_param_material.AsString() or "Sin Material"
            
            medidas = extraer_medida(param_name_value)
            
            param_pr_number = getParameter(element_type, "Classification.Uniclass.Pr.Number")
            param_pr_desc = getParameter(element_type, "Classification.Uniclass.Pr.Description")
            
            pr_number = ""
            pr_desc = ""
            if param_pr_number and param_pr_number.HasValue:
                pr_number = param_pr_number.AsString() or ""
            if param_pr_desc and param_pr_desc.HasValue:
                pr_desc = param_pr_desc.AsString() or ""

            # ==== Parámetros compartidos ====
            parameters_shared = {
                "COBie.Type.Name": "{} : {} : {}".format(category_name, fam_name, param_name_value),
                "COBie.Type.Category": "{} : {}".format(pr_number, pr_desc),
                "COBie.Type.Description": param_desc_value,
                "COBie.Type.Size": medidas,
                "COBie.Type.Material": param_name_material
            }

            # ==== Agregar parámetros estáticos ====
            parameters_shared.update(parameters_static)
            
            # ==== Agregar parámetros del Excel ====
            for excel_param, revit_param in param_mapping.items():
                if excel_param == "CODIGO":
                    continue  # Ya se usó para la búsqueda
                
                if excel_param in datos_excel and datos_excel[excel_param] is not None:
                    valor_excel = datos_excel[excel_param]
                    
                    # Convertir valores numéricos si es necesario
                    if excel_param in ["COBie.Type.WarrantyDurationParts", 
                                     "COBie.Type.WarrantyDurationLabor", 
                                     "COBie.Type.ExpectedLife"]:
                        try:
                            valor_excel = int(float(valor_excel)) if valor_excel else None
                        except (ValueError, TypeError):
                            valor_excel = None
                    
                    elif excel_param in ["COBie.Type.ReplacementCost",
                                       "COBie.Type.NominalLength",
                                       "COBie.Type.NominalWidth", 
                                       "COBie.Type.NominalHeight"]:
                        try:
                            if excel_param.startswith("COBie.Type.Nominal"):
                                # Convertir a unidades internas de Revit (pies)
                                valor_excel = UnitUtils.ConvertToInternalUnits(float(valor_excel), UnitTypeId.Meters) if valor_excel else None
                            else:
                                valor_excel = float(valor_excel) if valor_excel else None
                        except (ValueError, TypeError):
                            valor_excel = None
                    
                    if valor_excel is not None:
                        parameters_shared[revit_param] = valor_excel

            # ==== Aplicar todos los parámetros ====
            for param_name, value in parameters_shared.items():
                if value is not None:
                    try:
                        param = getParameter(element_type, param_name)
                        if param:
                            SetParameter(param, value)
                    except Exception as e:
                        print("Error estableciendo parámetro {}: {}".format(param_name, str(e)))

            conteo += 1
            
        except Exception as e:
            print("Error procesando elemento tipo {}: {}".format(element_type.Id, str(e)))
            elementos_omitidos += 1

# Mostrar resultados detallados
total_tipos = len(element_types)
mensaje = "Procesamiento completado:\n"
mensaje += "• Total de tipos encontrados: {}\n".format(total_tipos)
mensaje += "• Tipos procesados exitosamente: {}\n".format(conteo)
mensaje += "• Tipos omitidos: {}".format(elementos_omitidos)

TaskDialog.Show("Resultado del Proceso", mensaje)