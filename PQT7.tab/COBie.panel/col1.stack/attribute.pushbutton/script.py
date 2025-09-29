# -*- coding: utf-8 -*-
__title__ = "COBie Attribute"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository
import re

uidoc = revit.uidoc
doc = revit.doc

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()


# ---------------------- Utilidades ----------------------
def extract_number_nivel(nivel):
    """Extrae el número del nivel o devuelve TECHO si corresponde.
    Si no encuentra nada devuelve '0' como valor seguro.
    Si el nivel es 'n/a', se devuelve tal cual.
    """
    if not nivel:
        return "0"
    
    nivel = nivel.strip().upper()
    
    # Si explícitamente es n/a → devolver n/a
    if nivel == "N/A":
        return "n/a"
    
    # Si es TECHO → devolver TECHO
    if "TECHO" in nivel:
        return "TECHO"
    
    # Si hay número → devolver número
    match = re.search(r"-?\d+", nivel)
    return match.group(0) if match else "0"


def safe_value(value, default="n/a"):
    """Convierte None o vacío en un valor seguro para parámetros."""
    return value if value not in (None, "", []) else default


def build_parametros(nivel):
    """Construye el diccionario de parámetros combinando estáticos y dinámicos."""
    return dict(parametros_estaticos, **{
        "COBie.Attribute.Level": extract_number_nivel(nivel)
    })


# ---------------------- Project Information ----------------------
project_info = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ProjectInformation
).FirstElement()


# ---------------------- Datos estáticos ----------------------
CLIENTE = "PAQUETE 07 - ANIN"
MP = "Manual de Operación y Mantenimiento"
SUPPLIER = "CONSORCIO SYP"
TEST_SHEET = "Anexos del manual de operacion y mantenimiento"
SUBSECTOR = "Educacion"
CONTRACT_CODE = "30.145"
LAND_USE_CODE = "E1"
NO_APLICA = "n/a"


# ---------------------- Datos dinámicos de colegio ----------------------
school_instance = ColegiosRepository()
school_object = school_instance.codigo_colegio(doc)

maintenance_manual_value = safe_value(getattr(school_object, "operation_and_maintenance", None))
replacement_date_value = safe_value(getattr(school_object, "replacement_date", None))
document_reference_value = safe_value(getattr(school_object, "document_reference", None))
institution_code_value = safe_value(getattr(school_object, "code", None))
plot_code_value = safe_value(getattr(school_object, "plot_code", None))
project_code_value = institution_code_value  # ⚠️ Verificar si realmente deben ser iguales


# ---------------------- Parámetros estáticos ----------------------
parametros_estaticos = {
    "COBie.Attribute.AssetOwner": CLIENTE,
    "COBie.Attribute.MaintenanceProcedure": MP,
    "COBie.Attribute.OperationsAndMaintenanceManual": maintenance_manual_value,
    "COBie.Attribute.ReplacementDate": replacement_date_value,
    "COBie.Attribute.Supplier": SUPPLIER,
    "COBie.Attribute.TestSheet": TEST_SHEET,
    "COBie.Attribute.ContractCode": CONTRACT_CODE,
    "COBie.Attribute.DocumentReference": document_reference_value,
    "COBie.Attribute.InstitutionCode": institution_code_value,
    "COBie.Attribute.LandUseCode": LAND_USE_CODE,
    "COBie.Attribute.PlotCode": plot_code_value,
    "COBie.Attribute.ProjectCode": project_code_value,
    "COBie.Attribute.SubSector": SUBSECTOR,
    "COBie.ExternalIdentifier": NO_APLICA
}


# ---------------------- Selección y procesamiento ----------------------
Selections_elements = uidoc.Selection.PickElementsByRectangle()
errores = []

with revit.Transaction("Asignar COBie Attribute"):
    for element in Selections_elements:
        elementos_a_procesar = [element]

        # Si es un FamilyInstance, añadir también sus subcomponentes
        if isinstance(element, FamilyInstance):
            try:
                sub_elems = [doc.GetElement(sid) for sid in element.GetSubComponentIds()]
                elementos_a_procesar.extend([se for se in sub_elems if se])
            except Exception:
                pass  # Seguridad adicional

        # Procesar elemento principal y subcomponentes
        for elem in elementos_a_procesar:
            param_nivel = getParameter(elem, "S&P_NIVEL DE ELEMENTO")
            nivel = param_nivel.AsString() if param_nivel else ""
            parametros = build_parametros(nivel)

            for param, value in parametros.items():
                if value is None:
                    value = NO_APLICA
                p = getParameter(elem, param)
                if not SetParameter(p, value):
                    errores.append((elem.Id, param))

    # Aplicar solo parámetros estáticos al Project Information
    for project_param, value in parametros_estaticos.items():
        parameter = getParameter(project_info, project_param)
        if not SetParameter(parameter, value):
            errores.append((project_info.Id, project_param))


# ---------------------- Reporte final ----------------------
total = len(Selections_elements)
output = script.get_output()

if errores:
    output.print_md("### ❌ Errores en parámetros:")
    for eid, pname in errores:
        output.print_md("- ElementId `{}` | Parámetro: `{}`".format(eid, pname))
    TaskDialog.Show("Errores",
        "Se procesaron {} elementos.\nHubo {} errores. Revisa la consola de PyRevit.".format(total, len(errores)))
else:
    TaskDialog.Show("Éxito", "Se procesaron {} elementos sin errores.".format(total))

# 📊 Mostrar solo datos generales
output.print_md("### 📊 Resumen general de la operación:")
output.print_md("- Total de elementos procesados: **{}**".format(total))
output.print_md("- Cliente: **{}**".format(CLIENTE))
output.print_md("- Contrato: **{}**".format(CONTRACT_CODE))
output.print_md("- Proveedor: **{}**".format(SUPPLIER))
output.print_md("- Proyecto: **{}**".format(project_code_value))
output.print_md("- Colegio (código): **{}**".format(institution_code_value))
