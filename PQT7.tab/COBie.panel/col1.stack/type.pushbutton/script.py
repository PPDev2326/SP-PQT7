# -*- coding: utf-8 -*-
__title__ = "COBie Type"

import re
from Autodesk.Revit.DB import BuiltInParameter, StorageType, UnitUtils, UnitTypeId, FamilyInstance, ElementType, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit, forms
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter, get_param_value
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository
from DBRepositories.SchoolRepository import ColegiosRepository
from Helper._Excel import Excel
from Helper._HSpecialties import get_current_specialty

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

parametros_cobie = [
        "COBie.Type.Manufacturer",
        "COBie.Type.ModelNumber",
        "COBie.Type.WarrantyGuarantorParts",
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
        "COBie.Type.Description",  # Agregado aquí
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

# ===== OBTENER ESPECIALIDAD USANDO EL HELPER CENTRALIZADO =====
specialty_object = get_current_specialty(doc)

# Validar que se obtuvo la especialidad
if not specialty_object:
    forms.alert("No se pudo obtener la especialidad desde Información de Proyecto.\n"
                "Verifique que el parámetro S&P_ESPECIALIDAD esté configurado correctamente.", 
                exitscript=True)

# Extraer los datos de la especialidad
specialty = specialty_object.name
sp_accesibility = specialty_object.accessibility_performance
sp_code = specialty_object.code_perfomance
sp_sustainability = specialty_object.sustainability
sp_feature = specialty_object.feature

print("Especialidad detectada: {}".format(specialty))

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

# ==== Opciones de procesamiento ====
opciones = [
    "Sobreescribir todos los parámetros",
    "Solo llenar parámetros vacíos",
    "Cancelar"
]

opcion_seleccionada = forms.CommandSwitchWindow.show(
    opciones,
    message="¿Cómo deseas manejar los parámetros que ya tienen información?"
)

if not opcion_seleccionada or opcion_seleccionada == "Cancelar":
    script.exit()

modo_sobreescribir = (opcion_seleccionada == "Sobreescribir todos los parámetros")

# ==== Selección y preparación ====
try:
    selection = uidoc.Selection.PickElementsByRectangle()
    if not selection:
        forms.alert("No se seleccionaron elementos.", exitscript=True)
except Exception as e:
    forms.alert("No se seleccionaron elementos o se produjo un error:\n\n" + str(e), exitscript=True)

# ==== Obtenemos la hoja excel de acuerdo a la especialidad ====
data_list = None

if specialty == "Arquitectura":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -AR')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Arquitectura cargados:", len(data_list), "filas")

elif specialty == "Instalaciones Sanitarias":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - PL')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Sanitarias cargados:", len(data_list), "filas")

elif specialty == "Instalaciones Electricas":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -EE')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Eléctricas cargados:", len(data_list), "filas")

elif specialty == "Instalaciones de Comunicacion":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - IICC')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Comunicaciones cargados:", len(data_list), "filas")

elif specialty == "Instalaciones Mecanicas":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - ME')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, parametros_cobie)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Mecánicas cargados:", len(data_list), "filas")

else:
    forms.alert("Especialidad '{}' no reconocida para cargar datos Excel.".format(specialty), exitscript=True)

if not data_list:
    forms.alert("No se pudieron cargar los datos del Excel.", exitscript=True)

# Función para verificar si un parámetro está vacío
def parametro_esta_vacio(element, param_name):
    """
    Verifica si un parámetro está vacío o no tiene valor
    
    Args:
        element: Elemento de Revit
        param_name: Nombre del parámetro
        
    Returns:
        True si el parámetro está vacío, False si tiene valor
    """
    try:
        param = getParameter(element, param_name)
        if not param or not param.HasValue:
            return True
        
        # Verificar según el tipo de almacenamiento
        if param.StorageType == StorageType.String:
            valor = param.AsString()
            return not valor or valor.strip() == ""
        elif param.StorageType == StorageType.Double:
            return param.AsDouble() == 0.0
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger() == 0
        else:
            return True
    except:
        return True

# ==== Función para buscar datos del Excel por código ====
def buscar_datos_por_codigo(data_list, codigo_elemento):
    """
    Busca los datos del Excel que coincidan con el código del elemento.
    
    :param data_list: Lista de diccionarios con los datos del Excel.
    :type data_list: list
    :param codigo_elemento: Código del elemento de Revit.
    :type codigo_elemento: str
    :return: Diccionario con los datos encontrados o None si no encuentra.
    :rtype: dict or None
    """
    if not data_list:
        return None
    
    for row_data in data_list:
        codigo_excel = row_data.get("CODIGO")
        if codigo_excel and str(codigo_excel).strip() == str(codigo_elemento).strip():
            return row_data
    
    return None

# ==== Mapeo de parámetros Excel -> Revit ====
param_mapping = {
    "COBie.Type.Manufacturer": "COBie.Type.Manufacturer",
    "COBie.Type.ModelNumber": "COBie.Type.ModelNumber",
    "COBie.Type.WarrantyGuarantorParts": "COBie.Type.WarrantyGuarantorParts",
    "COBie.Type.WarrantyDurationParts": "COBie.Type.WarrantyDurationParts",
    "COBie.Type.WarrantyDurationLabor": "COBie.Type.WarrantyDurationLabor",
    "COBie.Type.ReplacementCost": "COBie.Type.ReplacementCost",
    "COBie.Type.ExpectedLife": "COBie.Type.ExpectedLife",
    "COBie.Type.NominalLength": "COBie.Type.NominalLength",
    "COBie.Type.NominalWidth": "COBie.Type.NominalWidth",
    "COBie.Type.NominalHeight": "COBie.Type.NominalHeight",
    "COBie.Type.Color": "COBie.Type.Color",
    "COBie.Type.Finish": "COBie.Type.Finish",
    "COBie.Type.Constituents": "COBie.Type.Constituents",
    "COBie.Type.Description": "COBie.Type.Description"  # Agregado aquí
}

# ==== Procesamiento: Instancia → Tipo ====
# Diccionario para almacenar: {type_id: {codigo, element_type, instancias}}
element_types_data = {}

for element in selection:
    # ==== Obtener código de la INSTANCIA ====
    codigo_elemento = None
    param_codigo = getParameter(element, "S&P_CODIGO DE ELEMENTO")
    if param_codigo and param_codigo.HasValue:
        codigo_elemento = param_codigo.AsString()
    
    if not codigo_elemento or not codigo_elemento.strip():
        continue
    
    # ==== Obtener el TIPO del elemento ====
    type_elem = doc.GetElement(element.GetTypeId())
    if not type_elem:
        print("No se pudo obtener el tipo para instancia {}".format(element.Id))
        continue
    
    type_id = type_elem.Id.IntegerValue
    
    # Almacenar o actualizar información del tipo
    if type_id not in element_types_data:
        element_types_data[type_id] = {
            "codigo": codigo_elemento,
            "element_type": type_elem,
            "instancias": []
        }
    
    element_types_data[type_id]["instancias"].append(element.Id)
    
    # Procesar subcomponentes si es FamilyInstance
    if isinstance(element, FamilyInstance):
        try:
            for sc_id in element.GetSubComponentIds():
                sub = doc.GetElement(sc_id)
                if sub:
                    type_sub = doc.GetElement(sub.GetTypeId())
                    if type_sub:
                        sub_type_id = type_sub.Id.IntegerValue
                        if sub_type_id not in element_types_data:
                            element_types_data[sub_type_id] = {
                                "codigo": codigo_elemento,  # Usar el mismo código de la instancia padre
                                "element_type": type_sub,
                                "instancias": []
                            }
                        element_types_data[sub_type_id]["instancias"].append(sub.Id)
        except Exception as e:
            print("Error procesando subcomponentes de {}: {}".format(element.Id, str(e)))

if not element_types_data:
    forms.alert("No se encontraron elementos válidos con códigos.", exitscript=True)

print("Elementos agrupados por tipo:", len(element_types_data))

# ==== Proceso COBie.Type con datos del Excel ====
conteo = 0
elementos_omitidos = 0
codigos_no_encontrados = []

# Preparar datos para transaction en masa
elementos_a_procesar = []

# Fase 1: Preparación de datos sin transaction
print("Preparando datos para procesamiento en masa...")

current_step = 0
total_types = len(element_types_data)

for type_id, type_data in element_types_data.items():
    element_type = type_data["element_type"]
    codigo_elemento = type_data["codigo"]
    instancias = type_data["instancias"]
    
    # Mostrar progreso cada 10 elementos
    current_step += 1
    if current_step % 10 == 0 or current_step == total_types:
        print("Preparando: {} de {}".format(current_step, total_types))
    
    # Verificar si el elemento debe procesarse
    param_cobie_type = getParameter(element_type, "COBie.Type")
    if not (param_cobie_type and param_cobie_type.StorageType == StorageType.Integer and param_cobie_type.AsInteger() == 1):
        elementos_omitidos += 1
        continue
    
    try:
        # ==== Buscar datos en Excel ====
        datos_excel = buscar_datos_por_codigo(data_list, codigo_elemento)
        
        if not datos_excel:
            if codigo_elemento not in codigos_no_encontrados:
                codigos_no_encontrados.append(codigo_elemento)
                print("No se encontraron datos en Excel para código: {}".format(codigo_elemento))
            elementos_omitidos += 1
            continue
        
        # ==== Preparar datos del elemento tipo ====
        category_object = element_type.Category
        category_name = category_object.Name if category_object else "Sin Categoría"
        
        fam_name = "Sin Familia"
        if isinstance(element_type, ElementType):
            fam_name = element_type.FamilyName
        
        param_name_value = "Sin Nombre"
        object_param_name = GetParameterAPI(element_type, BuiltInParameter.SYMBOL_NAME_PARAM)
        if object_param_name:
            param_name_value = object_param_name.AsString() or "Sin Nombre"
        
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

        # ==== Parámetros compartidos (base) ====
        parameters_shared = {
            "COBie.Type.Name": "{} : {} : {}".format(category_name, fam_name, param_name_value),
            "COBie.Type.Category": "{} : {}".format(pr_number, pr_desc),
            "COBie.Type.Size": medidas,
            "COBie.Type.Material": param_name_material
        }

        # ==== Agregar parámetros estáticos ====
        parameters_shared.update(parameters_static)
        
        # ==== Agregar parámetros del Excel ====
        for excel_param, revit_param in param_mapping.items():
            if excel_param in datos_excel and datos_excel[excel_param] is not None:
                valor_excel = datos_excel[excel_param]
                
                # Parámetros de longitud - convertir de metros a unidades internas de Revit
                if excel_param in ["COBie.Type.NominalLength",
                                   "COBie.Type.NominalWidth", 
                                   "COBie.Type.NominalHeight"]:
                    try:
                        if str(valor_excel).strip() != "":
                            valor_metros = float(valor_excel)
                            valor_excel = UnitUtils.ConvertToInternalUnits(valor_metros, UnitTypeId.Meters)
                        else:
                            valor_excel = 0
                    except (ValueError, TypeError):
                        valor_excel = 0
                
                elif excel_param == "COBie.Type.ReplacementCost":
                    try:
                        if str(valor_excel).strip() != "":
                            valor_excel = float(valor_excel)
                        else:
                            valor_excel = None
                    except (ValueError, TypeError):
                        valor_excel = None
                
                # Todos los demás parámetros son texto
                else:
                    if str(valor_excel).strip() != "":
                        valor_excel = str(valor_excel).strip()
                    else:
                        valor_excel = None
                
                if valor_excel is not None:
                    parameters_shared[revit_param] = valor_excel

        # Agregar elemento preparado a la lista
        elementos_a_procesar.append({
            "element_type": element_type,
            "parameters": parameters_shared,
            "codigo": codigo_elemento,
            "instancias": len(instancias)
        })
        
    except Exception as e:
        print("Error preparando elemento tipo {}: {}".format(element_type.Id, str(e)))
        elementos_omitidos += 1

print("Elementos preparados para procesamiento: {}".format(len(elementos_a_procesar)))

# Fase 2: Transaction en masa para aplicar parámetros
print("Iniciando transaction en masa...")

with revit.Transaction("Transferencia COBie Type Masiva"):
    current_element = 0
    total_elements = len(elementos_a_procesar)
    
    for elemento_data in elementos_a_procesar:
        current_element += 1
        if current_element % 10 == 0 or current_element == total_elements:
            print("Aplicando: {} de {}".format(current_element, total_elements))
        
        element_type = elemento_data["element_type"]
        parameters_shared = elemento_data["parameters"]
        
        try:
            # ==== Aplicar todos los parámetros al TIPO ====
            parametros_aplicados = 0
            parametros_omitidos = 0
            
            for param_name, value in parameters_shared.items():
                if value is not None:
                    try:
                        param = getParameter(element_type, param_name)
                        if param:
                            # Verificar si debemos aplicar el parámetro según el modo seleccionado
                            debe_aplicar = modo_sobreescribir or parametro_esta_vacio(element_type, param_name)
                            
                            if debe_aplicar:
                                SetParameter(param, value)
                                parametros_aplicados += 1
                            else:
                                parametros_omitidos += 1
                    except Exception as e:
                        print("Error estableciendo parámetro {} en tipo {}: {}".format(
                            param_name, element_type.Id, str(e)))

            conteo += 1
            if parametros_omitidos > 0:
                print("Procesado tipo {} con código: {} ({} instancias) - Aplicados: {}, Omitidos: {}".format(
                    element_type.Id, elemento_data["codigo"], elemento_data["instancias"], 
                    parametros_aplicados, parametros_omitidos))
            else:
                print("Procesado tipo {} con código: {} ({} instancias) - Aplicados: {}".format(
                    element_type.Id, elemento_data["codigo"], elemento_data["instancias"], parametros_aplicados))
            
        except Exception as e:
            print("Error procesando elemento tipo {}: {}".format(element_type.Id, str(e)))
            elementos_omitidos += 1

# Mostrar resultados detallados
total_tipos = len(element_types_data)
mensaje = "Procesamiento completado:\n"
mensaje += "• Modo: {}\n".format(opcion_seleccionada)
mensaje += "• Total de tipos encontrados: {}\n".format(total_tipos)
mensaje += "• Tipos procesados exitosamente: {}\n".format(conteo)
mensaje += "• Tipos omitidos: {}\n".format(elementos_omitidos)
if codigos_no_encontrados:
    mensaje += "• Códigos no encontrados en Excel: {}\n".format(len(set(codigos_no_encontrados)))

TaskDialog.Show("Resultado del Proceso", mensaje)