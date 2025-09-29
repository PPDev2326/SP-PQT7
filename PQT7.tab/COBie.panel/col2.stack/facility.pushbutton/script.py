# -*- coding: utf-8 -*-
__title__ = "COBie Facility"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from pyrevit import revit, script, forms
from Extensions._RevitAPI import getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository

uidoc = revit.uidoc
doc = revit.doc

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

try:
    # ==== Obtenemos el colegio correspondiente al modelo ====
    school_instance = ColegiosRepository()
    object_school = school_instance.codigo_colegio(doc)
    code_arcc = None
    code_cui = None
    code_local = None
    province_sc = None
    district_sc = None
    populated = None
    created_by = None

    if object_school:
        school = object_school.name
        school_classification_desc = object_school.classification.facility_description
        school_classification_number = object_school.classification.facility_number
        
        # ==== Variables ====
        code_arcc = object_school.location.code_arcc
        code_cui = object_school.location.code_cui
        code_local = object_school.location.code_local
        province_sc = object_school.location.province
        district_sc = object_school.location.district
        populated = object_school.location.populated_center
        created_by = object_school.created_by
        
        category = "{} : {}".format(school_classification_number, school_classification_desc)
        
        facility_descripcion = object_school.cobie.facility_description
        facility_project_descripcion = object_school.cobie.project_description
        facility_site_descripcion = object_school.cobie.site_description

    # ==== Obtenemnos la instancia del Project Information ====
    fec = FilteredElementCollector(doc)
    project_information_element = fec.OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
    
    # Validación: si no existe Project Information, detener
    if not project_information_element:
        forms.alert("No se encontró el Project Information en este modelo.", title="Error")
        script.exit()

    # ==== Seleccionamos los elementos del modelo activo ====
    # try:
    #     list_elements = uidoc.Selection.PickElementsByRectangle("Selecciona los elementos para el COBie.Facility")

    # except OperationCanceledException:
    #     forms.alert("Operación cancelada: no se seleccionaron elementos para procesar COBie.Facility",
    #     title="Cancelación")

    # ==== Variables constantes ====
    CENSUS = "Urbana"
    DEPARTMENT = "Piura"
    CREATED_ON = "2023-04-07T16:38:56"
    PROJECT_NAME = "11 Intervenciones (Instituciones Educativas) en el departamento de Piura"
    SITE_NAME = "{} - {}".format(DEPARTMENT, district_sc)
    LINEAR_UNITS = "Metros"
    AREA_UNITS = "Metros cuadrados"
    VOLUMEN_UNITS = "Metros Cúbicos"
    CURRENCY_UNIT = "Sol"
    AREA_MEASUREMENT = "Método de cálculo de área predeterminada de Revit"
    PHASE = "En uso"

    conteo = 0
    
    # ==== Parametros de revit a utlizar ====
    parametros = {
        "S&P_CODIGO ARCC": code_arcc,
        "S&P_CODIGO CUI": code_cui,
        "S&P_CODIGO LOCAL": code_local,
        "S&P_AREA CENSAL": CENSUS,
        "S&P_DEPARTAMENTO": DEPARTMENT,
        "S&P_PROVINCIA": province_sc,
        "S&P_DISTRITO": district_sc,
        "S&P_CENTRO POBLADO": populated,
        "Classification.Facility.Description": school_classification_desc,
        "Classification.Facility.Number": school_classification_number,
        "COBie.Facility.Name": school,
        "COBie.CreatedBy": created_by,
        "COBie.CreatedOn": CREATED_ON,
        "COBie.Facility.Category": category,
        "COBie.Facility.ProjectName": PROJECT_NAME,
        "COBie.Facility.SiteName": SITE_NAME,
        "COBie.Facility.LinearUnits": LINEAR_UNITS,
        "COBie.Facility.AreaUnits": AREA_UNITS,
        "COBie.Facility.VolumeUnits": VOLUMEN_UNITS,
        "COBie.Facility.CurrencyUnit": CURRENCY_UNIT,
        "COBie.Facility.AreaMeasurement": AREA_MEASUREMENT,
        "COBie.Facility.Description": facility_descripcion,
        "COBie.Facility.ProjectDescription": facility_project_descripcion,
        "COBie.Facility.SiteDescription": facility_site_descripcion,
        "COBie.Facility.Phase": PHASE,
    }

    with revit.Transaction("Completar datos COBie.Facility"):
    # ==== Recorremos los elementos de mi variable parametros para asiganarlo al project information ====
        for param, value in parametros.items():
            
            param = getParameter(project_information_element, param)
            set_param = SetParameter(param, value)
            if set_param:
                conteo += 1
        if conteo > 0:
            TaskDialog.Show("Informativo", "Datos cargados correctamente.\n\t{}/{} procesados exitosamente".format(conteo, len(parametros)))

except OperationCanceledException:
    forms.alert("ERROR: comando cancelado por el usuario.", "Cancelación")

output = script.get_output()

output.print_md("## 📊 Resumen general de la operación")
output.print_md("### 🏫 Colegio procesado:")
output.print_md("- **Nombre:** {}".format(school))
output.print_md("- **Código ARCC:** {}".format(code_arcc))
output.print_md("- **Código CUI:** {}".format(code_cui))
output.print_md("- **Distrito:** {}, {}".format(district_sc, province_sc))

output.print_md("### 📝 Parámetros asignados a Project Information")
output.print_md("- **Parámetros procesados:** {}".format(conteo))

for key, value in parametros.items():
    output.print_md("| {} | {} |".format(key, value if value else "-"))