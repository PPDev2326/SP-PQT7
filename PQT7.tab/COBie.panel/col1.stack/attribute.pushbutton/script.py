# -*- coding: utf-8 -*-
__title__ = "COBie Attribute"

from Autodesk.Revit.DB import Document, FilteredElementCollector, BuiltInCategory, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms, script
from Extensions._RevitAPI import getParameter, SetParameter

uidoc = revit.uidoc
doc = revit.doc

# ==== Obtenemos la categoria projectinformation ====
project_info = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()

# ==== Datos estáticos ====
CLIENTE = "PAQUETE 07 - ANIN"
MP = "Manual de Operación y Mantenimiento"
SUPPLIER = "CONSORCIO SYP"
TEST_SHEET = "Anexos del manual de operacion y mantenimiento"
SUBSECTOR = "Educacion"

codigo_colegio = ObtenerCodigoColegio(doc)



contract_code = ObtenerCodigoContrato(doc)
doc_reference = "28554300_Peru-Bic_Schools_Report_{}1".format(codigo_colegio)
code_inst = ObtenerCodigoInstitucion(doc)
LAND_CODE = "E1"
plot_code = ObtenerCodigoTrama(doc)

CLASS_DESC = "Educational complexes"
CLASS_NUMBER = "Co_25_10"

parametros_estaticos = {
    "COBie.Attribute.AssetOwner" : CLIENTE,
    "COBie.Attribute.MaintenanceProcedure" : MP,
    "COBie.Attribute.OperationsAndMaintenanceManual" : doc_omm,
    "COBie.Attribute.ReplacementDate" : RPM_DATE,
    "COBie.Attribute.Supplier" : SUPPLIER,
    "COBie.Attribute.TestSheet" : TEST_SHEET,
    "COBie.Attribute.ContractCode" : contract_code,
    "COBie.Attribute.DocumentReference" : doc_reference,
    "COBie.Attribute.InstitutionCode" : code_inst,
    "COBie.Attribute.LandUseCode" : LAND_CODE,
    "COBie.Attribute.PlotCode" : plot_code,
    "COBie.Attribute.ProjectCode" : codigo_colegio,
    "COBie.Attribute.SubSector" : SUBSECTOR,
    "Classification.Facility.Description" : CLASS_DESC,
    "Classification.Facility.Number" : CLASS_NUMBER
}

Selections_elements = uidoc.Selection.PickElementsByRectangle()
errores = []

with revit.Transaction("Selecciona los elementos para completar COBie Attribute"):
    for element in Selections_elements:
        elementos_a_procesar = [element]

        # Solo considerar subcomponentes si es FamilyInstance
        if isinstance(element, FamilyInstance):
            try:
                sub_ids = element.GetSubComponentIds()
                if sub_ids:
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass  # Seguridad adicional si el método no está disponible en algún tipo

        # Procesar elemento principal y subcomponentes
        for elem in elementos_a_procesar:
            param_nivel = getParameter(elem, "Nivel de elemento")
            nivel = param_nivel.AsString() if param_nivel else ""

            parametros_API = {
                "COBie.Attribute.Level" : nivel
            }

            parametros = dict(parametros_estaticos)
            parametros.update(parametros_API)

            for param, value in parametros.items():
                p = getParameter(elem, param)
                if not SetParameter(p, value):
                    errores.append((elem.Id, param))

    # También aplicar a la Project Information
    for project_param, value in parametros.items():
        parameter = getParameter(project_info, project_param)
        if not SetParameter(parameter, value):
            errores.append((project_info.Id, project_param))

# Mostrar resultados
if errores:
    output = script.get_output()
    output.print_md("### ❌ Errores en parámetros:")
    for eid, pname in errores:
        output.print_md("- ElementId `{}` | Parámetro: `{}`".format(eid, pname))
    TaskDialog.Show("Errores", "Algunos parámetros no se pudieron completar. Revisa la consola de PyRevit.")
else:
    TaskDialog.Show("Éxito", "Todos los parámetros se completaron correctamente.")