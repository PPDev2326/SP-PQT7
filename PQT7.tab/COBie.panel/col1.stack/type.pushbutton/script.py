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

# ==== Obtenemos la hoja excel de acuerdo a la especialidad ====
if specialty == "ARQUITECTURA":
    excel = Excel().read_excel('ESTANDAR COBIE  -AR')
    print(excel)

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

# ==== Datos a revisar ====
# NOMINAL_LENGTH = NOMINAL_WIDTH = NOMINAL_HEIGHT = UnitUtils.ConvertToInternalUnits(1.00, UnitTypeId.Meters)
# ==== Fin de datos a revisar ====

parameters_static = {
    "COBie.Type.CreatedBy": created_by_value,
    "COBie.Type.CreatedOn": CREATED_ON,
    "COBie.Type.AssetType": asset_type,
    # "COBie.Type.Manufacturer": MANUFACTURER,                      # Nombre de la empresa ejecutora
    # "COBie.Type.ModelNumber": model_number,                       # Codificación de acuerdo al standard de cada SC
    "COBie.Type.WarrantyGuarantorParts": created_by_value,       
    # "COBie.Type.WarrantyDurationParts": 1,                        # Varia de acuerdo al excel
    "COBie.Type.WarrantyGuarantorLabor": created_by_value,
    # "COBie.Type.WarrantyDurationLabor": DURATION_LABOR,           # Varia de acuerdo al excel
    "COBie.Type.WarrantyDurationUnit": DURATION_UNIT,
    # "COBie.Type.ReplacementCost": REPLACEMENT_COST,               # Varia de acuerdo al excel
    # "COBie.Type.ExpectedLife": EXPECTED_LIFE,                     # Varia de acuerdo al excel
    "COBie.Type.DurationUnit": DURATION_UNIT,
    "COBie.Type.WarrantyDescription": warranty_description_value,
    # "COBie.Type.NominalLength": NOMINAL_LENGTH,                   # Varia de acuerdo al excel 
    # "COBie.Type.NominalWidth": NOMINAL_WIDTH,                     # Varia de acuerdo al excel
    # "COBie.Type.NominalHeight": NOMINAL_HEIGHT,                   # Varia de acuerdo al excel
    "COBie.Type.ModelReference": warranty_description_value,
    "COBie.Type.Shape": SHAPE,
    # "COBie.Type.Color": COLOR,                                    # Varia de acuerdo al excel
    # "COBie.Type.Finish": finish,                                  # Varia de acuerdo al excel
    "COBie.Type.Grade": GRADE,
    "COBie.Type.Features": sp_feature,
    "COBie.Type.AccessibilityPerformance": sp_accesibility,
    "COBie.Type.CodePerformance": sp_code,
    "COBie.Type.SustainabilityPerformance": sp_sustainability,
    # "COBie.Type.Constituents": constituents                       # Varia de acuerdo al excel
}

# ==== Selección y preparación ====
try:
    selection = uidoc.Selection.PickElementsByRectangle()
    if not selection:
        forms.alert("No se seleccionaron elementos.", exitscript=True)
except Exception as e:
    forms.alert("No se seleccionaron elementos o se produjo un error:\n\n" + str(e), exitscript=True)

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

with revit.Transaction("Transferencia COBie Type Optimizada"):
    for type_id, element_type in element_types.items():
        
        # Verificar si el elemento debe procesarse
        param_cobie_type = getParameter(element_type, "COBie.Type")
        if not (param_cobie_type and param_cobie_type.StorageType == StorageType.Integer and param_cobie_type.AsInteger() == 1):
            elementos_omitidos += 1
            continue
        
        try:
            # ==== Obtenemos la categoria del elemento de tipo ====
            category_object = element_type.Category
            category_name = category_object.Name if category_object else "Sin Categoría"
            
            # ==== Obtenemos el nombre de la familia del elemento ====
            fam_name = "Sin Familia"
            if isinstance(element_type, ElementType):
                fam_name = element_type.FamilyName
            
            # ==== Obtenemos el tipo del elemento seleccionado ====
            param_name_value = "Sin Nombre"
            object_param_name = GetParameterAPI(element_type, BuiltInParameter.SYMBOL_NAME_PARAM)
            if object_param_name:
                param_name_value = object_param_name.AsString() or "Sin Nombre"
            
            # ==== el parametro Descripción ====
            param_desc_value = "Sin Descripción"
            object_param_desc = getParameter(element_type, "Descripción")
            if object_param_desc and object_param_desc.HasValue:
                param_desc_value = object_param_desc.AsString() or "Sin Descripción"
            
            # ==== Obtenemos el parámetro material ====
            param_name_material = "Sin Material"
            object_param_material = getParameter(element_type, "S&P_MATERIAL DE ELEMENTO")
            if object_param_material and object_param_material.HasValue:
                param_name_material = object_param_material.AsString() or "Sin Material"
            
            # ==== Obtenemos las medidas de las familias o tipos de los elementos seleccionados ====
            medidas = extraer_medida(param_name_value)
            
            # ==== Obtenemos clasificación Uniclass ====
            param_pr_number = getParameter(element_type, "Classification.Uniclass.Pr.Number")
            param_pr_desc = getParameter(element_type, "Classification.Uniclass.Pr.Description")
            
            pr_number = ""
            pr_desc = ""
            if param_pr_number and param_pr_number.HasValue:
                pr_number = param_pr_number.AsString() or ""
            if param_pr_desc and param_pr_desc.HasValue:
                pr_desc = param_pr_desc.AsString() or ""

            parameters_shared = {
                "COBie.Type.Name": "{} : {} : {}".format(category_name, fam_name, param_name_value),
                "COBie.Type.Category": "{} : {}".format(pr_number, pr_desc),
                "COBie.Type.Description": param_desc_value,
                "COBie.Type.Size": medidas,
                "COBie.Type.Material": param_name_material
            }

            parameters_shared.update(parameters_static)

            # Aplicar parámetros
            for param_n, value in parameters_shared.items():
                if value is not None:  # Solo aplicar si hay valor
                    try:
                        param = getParameter(element_type, param_n)
                        if param:
                            SetParameter(param, value)
                    except Exception as e:
                        print("Error estableciendo parámetro {}: {}".format(param_n, str(e)))

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