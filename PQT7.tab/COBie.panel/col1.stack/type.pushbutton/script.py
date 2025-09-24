# -*- coding: utf-8 -*-
__title__ = "COBie Type"

from Autodesk.Revit.DB import BuiltInParameter, StorageType, UnitUtils, UnitTypeId, FamilyInstance, ElementType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository
from DBRepositories.SchoolRepository import ColegiosRepository

uidoc = revit.uidoc
doc = revit.doc

# ==== Obtenemos la especialidad del modelo activo y sus datos ====
specialty_object = SpecialtiesRepository().get_specialty_by_document(doc)
if specialty_object:
    specialty = specialty_object.name
    sp_accesibility = specialty_object.accessibility_performance
    sp_code = specialty_object.code_perfomance
    sp_sustainability = specialty_object.sustainability

# ==== Obtenemos el colegio correspondiente segun modelo y sus datos necesarios ====
school_object = ColegiosRepository.codigo_colegio(doc)
if school_object:
    school = school_object.name
    created_by_value = school_object.created_by

# ==== Datos estáticos ====
CREATED_ON = "2025-08-04T11:59:30"

# ==== Datos a revisar ====
DURATION_PARTS = DURATION_LABOR = EXPECTED_LIFE = "5"
DURATION_UNIT = "Años"
REPLACEMENT_COST = 1.00
NOMINAL_LENGTH = NOMINAL_WIDTH = NOMINAL_HEIGHT = UnitUtils.ConvertToInternalUnits(1.00, UnitTypeId.Meters)
warranty_description = "{}-CGC01-MM-GR-000010 | Manual de operacion y mantenimiento".format(ObtenerCodigoColegio(doc))
SHAPE = "Poligonal - Segun modelo"
COLOR = "Basicos"
GRADE = "Grado Estandar"
MATERIAL = "Varios"
# ==== Fin de datos a revisar ====

parameters_static = {
    "COBie.Type.CreatedBy": created_by_value,
    "COBie.Type.CreatedOn": CREATED_ON,
    "COBie.Type.AssetType": assettype,
    "COBie.Type.Manufacturer": MANUFACTURER,
    "COBie.Type.ModelNumber": model_number,
    "COBie.Type.WarrantyGuarantorParts": GUARANTOR_PARTS,
    "COBie.Type.WarrantyDurationParts": DURATION_PARTS,
    "COBie.Type.WarrantyGuarantorLabor": GUARANTOR_LABOR,
    "COBie.Type.WarrantyDurationLabor": DURATION_LABOR,
    "COBie.Type.WarrantyDurationUnit": DURATION_UNIT,
    "COBie.Type.ReplacementCost": REPLACEMENT_COST,
    "COBie.Type.ExpectedLife": EXPECTED_LIFE,
    "COBie.Type.DurationUnit": DURATION_UNIT,
    "COBie.Type.WarrantyDescription": warranty_description,
    "COBie.Type.NominalLength": NOMINAL_LENGTH,
    "COBie.Type.NominalWidth": NOMINAL_WIDTH,
    "COBie.Type.NominalHeight": NOMINAL_HEIGHT,
    "COBie.Type.ModelReference": modelref,
    "COBie.Type.Shape": SHAPE,
    "COBie.Type.Size": size,
    "COBie.Type.Color": COLOR,
    "COBie.Type.Finish": finish,
    "COBie.Type.Grade": GRADE,
    "COBie.Type.Material": MATERIAL,
    "COBie.Type.Features": features,
    "COBie.Type.AccessibilityPerformance": sp_accesibility,
    "COBie.Type.CodePerformance": sp_code,
    "COBie.Type.SustainabilityPerformance": sp_sustainability,
    "COBie.Type.Constituents": constituents
}

# ==== Selección y preparación ====
selection = uidoc.Selection.PickElementsByRectangle()
element_types = dict()  # type_id_int -> element_type

for element in selection:
    # Procesar el tipo del elemento principal
    type_elem = doc.GetElement(element.GetTypeId())
    if type_elem:
        element_types[type_elem.Id.IntegerValue] = type_elem

    # Procesar tipos de subcomponentes (si existen)
    if isinstance(element, FamilyInstance):
        for sc_id in element.GetSubComponentIds():
            sub = doc.GetElement(sc_id)
            if sub:
                type_sub = doc.GetElement(sub.GetTypeId())
                if type_sub:
                    element_types[type_sub.Id.IntegerValue] = type_sub

# ==== Proceso COBie.Type optimizado ====
conteo = 0
with revit.Transaction("Transferencia COBie Type Optimizada"):
    for type_id, element_type in element_types.items():
        
        # ==== Obtenemos la categoria del elemento de tipo ====
        category_object = element_type.Category
        category_name = category_object.Name
        
        # ==== Obtenemos el nombre de la familia del elemento ====
        if isinstance(element_type, ElementType):
            fam_name = element_type.FamilyName
        
        # ==== Obtenemos el tipo del elemento seleccionado ====
        object_param_name = GetParameterAPI(element_type, BuiltInParameter.SYMBOL_NAME_PARAM)
        if object_param_name:
            param_name_value = object_param_name.AsString()
        
        # ==== el parametro Descripción ====
        object_param_desc = getParameter(element_type, "Descripción")
        if object_param_desc:
            param_desc_value = object_param_desc.AsString()
        
        param_cobie_type = getParameter(element_type, "COBie.Type")
        if not (param_cobie_type and param_cobie_type.StorageType == StorageType.Integer and param_cobie_type.AsInteger() == 1):
            continue
        
        conteo += 1

        param_code = GetParameterAPI(element_type, BuiltInParameter.UNIFORMAT_CODE, es_editable=True)
        param_name = GetParameterAPI(element_type, BuiltInParameter.ALL_MODEL_TYPE_NAME)
        param_pr_number = getParameter(element_type, "Classification.Uniclass.Pr.Number")
        param_pr_desc = getParameter(element_type, "Classification.Uniclass.Pr.Description")

        parameters_shared = {
            "COBie.Type.Name": "{} : {} : {}".format(
                category_name,
                fam_name,
                param_name_value
            ),
            "COBie.Type.Category": "{} : {}".format(
                param_pr_number.AsString() if param_pr_number else "",
                param_pr_desc.AsString() if param_pr_desc else ""
            ),
            "COBie.Type.Description": "{}".format(
                param_desc_value
            )
        }

        parameters_shared.update(parameters_static)

        for param_n, value in parameters_shared.items():
            param = getParameter(element_type, param_n)
            SetParameter(param, value)

TaskDialog.Show("Informativo", "Total: {} tipos procesados.".format(conteo))