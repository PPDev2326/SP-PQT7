# -*- coding: utf-8 -*-
__title__ = "COBie Type"

from Autodesk.Revit.DB import BuiltInParameter, StorageType, UnitUtils, UnitTypeId, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit
from Extensions._RevitAPI import getParameter, GetParameterAPI, SetParameter
from Extensions._Dictionary import obtener_especialidad, obtener_model_number, ObtenerCodigoColegio, ObtenerCaracteristicas

uidoc = revit.uidoc
doc = revit.doc

# ==== Datos estáticos ====
especialidad = obtener_especialidad()
caracteristicas = ObtenerCaracteristicas(especialidad)
creado_por = MANUFACTURER = GUARANTOR_PARTS = GUARANTOR_LABOR = "pruiz@cgeb.com.pe"
hora_creacion = "2025-04-18T16:45:10"

assettype = "Semi-fijo" if especialidad in ["Arquitectura, Señaletica y evacuacion, Equipamiento y mobiliario"] else "Fijo"
model_number = modelref = size = finish = obtener_model_number()
DURATION_PARTS = DURATION_LABOR = EXPECTED_LIFE = "5"
DURATION_UNIT = "Años"
REPLACEMENT_COST = 1.00
NOMINAL_LENGTH = NOMINAL_WIDTH = NOMINAL_HEIGHT = UnitUtils.ConvertToInternalUnits(1.00, UnitTypeId.Meters)
warranty_description = "{}-CGC01-MM-GR-000010 | Manual de operacion y mantenimiento".format(ObtenerCodigoColegio(doc))
SHAPE = "Poligonal - Segun modelo"
COLOR = "Basicos"
GRADE = "Grado Estandar"
MATERIAL = "Varios"
features, access_performance, code_performance, sustainability_performance, constituents = caracteristicas

parameters_static = {
    "COBie.Type.CreatedBy": creado_por,
    "COBie.Type.CreatedOn": hora_creacion,
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
    "COBie.Type.AccessibilityPerformance": access_performance,
    "COBie.Type.CodePerformance": code_performance,
    "COBie.Type.SustainabilityPerformance": sustainability_performance,
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
        param_cobie_type = getParameter(element_type, "COBie.Type")
        if not (param_cobie_type and param_cobie_type.StorageType == StorageType.Integer and param_cobie_type.AsInteger() == 1):
            continue

        conteo += 1

        param_code = GetParameterAPI(element_type, BuiltInParameter.UNIFORMAT_CODE, es_editable=True)
        param_name = GetParameterAPI(element_type, BuiltInParameter.ALL_MODEL_TYPE_NAME)
        param_ss_number = getParameter(element_type, "Classification.Uniclass.Ss.Number")
        param_ss_desc = getParameter(element_type, "Classification.Uniclass.Ss.Description")

        parameters_shared = {
            "COBie.Type.Name": "{}_{}".format(
                param_code.AsString() if param_code else "",
                param_name.AsString() if param_name else ""
            ),
            "COBie.Type.Category": "{} {}".format(
                param_ss_number.AsString() if param_ss_number else "",
                param_ss_desc.AsString() if param_ss_desc else ""
            ),
            "COBie.Type.Description": "{}".format(
                param_name.AsString() if param_name else ""
            )
        }

        parameters_shared.update(parameters_static)

        for param_n, value in parameters_shared.items():
            param = getParameter(element_type, param_n)
            SetParameter(param, value)

TaskDialog.Show("Informativo", "Total: {} tipos procesados.".format(conteo))