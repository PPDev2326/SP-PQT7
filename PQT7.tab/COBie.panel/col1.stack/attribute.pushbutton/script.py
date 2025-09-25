# -*- coding: utf-8 -*-
__title__ = "COBie Attribute"

from Autodesk.Revit.DB import Document, FilteredElementCollector, BuiltInCategory, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms, script
from Extensions._RevitAPI import getParameter, SetParameter
from DBRepositories.SchoolRepository import ColegiosRepository
import time
from datetime import datetime

uidoc = revit.uidoc
doc = revit.doc
output = script.get_output()

def extract_number_nivel(nivel):
    if not nivel:
        return "0"
    try:
        parts = nivel.split()
        return parts[-1].strip() if parts else "0"
    except:
        return "0"

def print_header():
    """Imprime el encabezado profesional del reporte"""
    output.print_md("# 🏗️ COBie Attribute Assignment Report")
    output.print_md("---")
    output.print_md("**📅 Fecha de ejecución:** {}".format(datetime.now().strftime('%d/%m/%Y %H:%M:%S')))
    output.print_md("**📄 Documento:** {}".format(doc.Title))
    output.print_md("**👤 Usuario:** {}".format(doc.Application.Username))
    output.print_md("---")

def print_project_info(school_object):
    """Imprime información del proyecto"""
    output.print_md("## 📋 Información del Proyecto")
    
    if school_object:
        output.print_md("**🏫 Código de Institución:** `{}`".format(school_object.code if school_object.code else 'N/A'))
        output.print_md("**📍 Código de Parcela:** `{}`".format(school_object.plot_code if school_object.plot_code else 'N/A'))
        output.print_md("**📚 Manual de O&M:** `{}`".format(school_object.operation_and_maintenance if school_object.operation_and_maintenance else 'N/A'))
        output.print_md("**🔄 Fecha de Reemplazo:** `{}`".format(school_object.replacement_date if school_object.replacement_date else 'N/A'))
    else:
        output.print_md("⚠️ **No se encontró información del colegio en la base de datos**")
    
    output.print_md("---")

def print_processing_start(total_elements):
    """Imprime el inicio del procesamiento"""
    output.print_md("## ⚙️ Procesamiento de Elementos")
    output.print_md("**📊 Total de elementos seleccionados:** `{}`".format(total_elements))
    output.print_md("**🔧 Iniciando asignación de parámetros COBie...**")
    output.print_md("")

def print_element_processing(elem, nivel, sub_count=0):
    """Imprime información de procesamiento por elemento"""
    element_type = elem.GetType().Name
    element_id = elem.Id
    
    if sub_count > 0:
        output.print_md("🔸 **ID:** `{}` | **Tipo:** `{}` | **Nivel:** `{}` | **Subcomponentes:** `{}`".format(element_id, element_type, nivel, sub_count))
    else:
        output.print_md("🔸 **ID:** `{}` | **Tipo:** `{}` | **Nivel:** `{}`".format(element_id, element_type, nivel))

def print_parameters_summary(parametros):
    """Imprime resumen de parámetros a asignar"""
    output.print_md("### 📝 Parámetros COBie a asignar:")
    output.print_md("| Parámetro | Valor |")
    output.print_md("|-----------|-------|")
    
    for param, value in parametros.items():
        # Truncar valores muy largos
        display_value = str(value)[:50] + "..." if len(str(value)) > 50 else str(value)
        output.print_md("| `{}` | `{}` |".format(param, display_value))
    
    output.print_md("")

def print_results_summary(total_processed, total_parameters, errores, processing_time):
    """Imprime el resumen final de resultados"""
    output.print_md("---")
    output.print_md("# 📊 Resumen de Resultados")
    
    # Estadísticas generales
    output.print_md("## 📈 Estadísticas Generales")
    output.print_md("**✅ Elementos procesados:** `{}`".format(total_processed))
    output.print_md("**🏷️ Parámetros totales asignados:** `{}`".format(total_parameters))
    output.print_md("**⏱️ Tiempo de procesamiento:** `{:.2f} segundos`".format(processing_time))
    avg_time = processing_time/total_processed if total_processed > 0 else 0
    output.print_md("**📊 Promedio por elemento:** `{:.3f} seg/elemento`".format(avg_time))
    
    # Resultados
    if errores:
        error_count = len(errores)
        success_rate = ((total_parameters - error_count) / total_parameters * 100) if total_parameters > 0 else 0
        
        output.print_md("**❌ Errores encontrados:** `{}`".format(error_count))
        output.print_md("**✅ Tasa de éxito:** `{:.1f}%`".format(success_rate))
        
        output.print_md("## ❌ Detalle de Errores")
        output.print_md("| ID Elemento | Parámetro | Tipo de Error |")
        output.print_md("|-------------|-----------|---------------|")
        
        for error in errores:
            elem_id, param_name = error
            output.print_md("| `{}` | `{}` | Parámetro no encontrado o fallo en asignación |".format(elem_id, param_name))
        
        output.print_md("")
        output.print_md("### 💡 Recomendaciones:")
        output.print_md("- Verificar que los parámetros existen en las familias")
        output.print_md("- Comprobar permisos de escritura en los parámetros")
        output.print_md("- Revisar que las familias estén cargadas correctamente")
    else:
        output.print_md("**✅ Estado:** `COMPLETADO SIN ERRORES`")
        output.print_md("**🎉 Todos los parámetros COBie se asignaron exitosamente**")

def print_footer():
    """Imprime el pie del reporte"""
    output.print_md("---")
    output.print_md("## 🔧 Información Técnica")
    output.print_md("**🏷️ Script:** COBie Attribute Assignment")
    output.print_md("**📦 PyRevit Version:** {}".format(script.get_envvar('pyRevitVersion')))
    output.print_md("**🔌 Revit API:** {}".format(doc.Application.VersionNumber))
    output.print_md("**👨‍💻 Desarrollado por:** Consorcio SYP")
    output.print_md("---")
    output.print_md("*Reporte generado automáticamente por PyRevit*")

# ==================== INICIO DEL SCRIPT ====================

start_time = time.time()

# Imprimir encabezado
print_header()

# ==== Obtenemos la categoria projectinformation ====
project_info = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()

# ==== Datos estáticos ====
CLIENTE = "PAQUETE 07 - ANIN"
MP = "Manual de Operación y Mantenimiento"
SUPPLIER = "CONSORCIO SYP"
TEST_SHEET = "Anexos del manual de operacion y mantenimiento"
SUBSECTOR = "Educacion"
CONTRACT_CODE = "30.145"
LAND_USE_CODE = "E1"
NO_APLICA = "n/a"

# ==== Instanciamos el tipo ColegiosRepository y sus propiedades ====
school_instance = ColegiosRepository()
school_object = school_instance.codigo_colegio(doc)
maintenance_manual_value = None
replacement_date_value = None
document_reference_value = None
institution_code_value = None
plot_code_value = None
project_code_value = None

# ==== Verificamos si school_object existe ====
if school_object:
    maintenance_manual_value = school_object.operation_and_maintenance
    replacement_date_value = school_object.replacement_date
    document_reference_value = school_object.document_reference
    institution_code_value = school_object.code
    plot_code_value = school_object.plot_code
    project_code_value = school_object.code

# Imprimir información del proyecto
print_project_info(school_object)

parametros_estaticos = {
    "COBie.Attribute.AssetOwner" : CLIENTE,
    "COBie.Attribute.MaintenanceProcedure" : MP,
    "COBie.Attribute.OperationsAndMaintenanceManual" : maintenance_manual_value,
    "COBie.Attribute.ReplacementDate" : replacement_date_value,
    "COBie.Attribute.Supplier" : SUPPLIER,
    "COBie.Attribute.TestSheet" : TEST_SHEET,
    "COBie.Attribute.ContractCode" : CONTRACT_CODE,
    "COBie.Attribute.DocumentReference" : document_reference_value,
    "COBie.Attribute.InstitutionCode" : institution_code_value,
    "COBie.Attribute.LandUseCode" : LAND_USE_CODE,
    "COBie.Attribute.PlotCode" : plot_code_value,
    "COBie.Attribute.ProjectCode" : project_code_value,
    "COBie.Attribute.SubSector" : SUBSECTOR,
    "COBie.ExternalIdentifier" : NO_APLICA
}

Selections_elements = uidoc.Selection.PickElementsByRectangle()
errores = []

# Validar selección
if not Selections_elements:
    output.print_md("⚠️ **No se seleccionaron elementos. El script se detendrá.**")
    TaskDialog.Show("Aviso", "No se seleccionaron elementos.")
    script.exit()

# Imprimir inicio de procesamiento
print_processing_start(len(Selections_elements))

# Mostrar parámetros que se van a asignar (solo una vez)
sample_element = Selections_elements[0]
param_nivel = getParameter(sample_element, "S&P_NIVEL DE ELEMENTO")
nivel = param_nivel.AsString() if param_nivel else ""
nivel_number_value = extract_number_nivel(nivel)

parametros_API = {"COBie.Attribute.Level" : nivel_number_value}
parametros_sample = dict(parametros_estaticos)
parametros_sample.update(parametros_API)
print_parameters_summary(parametros_sample)

# Contadores
total_processed = 0
total_parameters_assigned = 0

with revit.Transaction("Selecciona los elementos para completar COBie Attribute"):
    for element in Selections_elements:
        elementos_a_procesar = [element]
        sub_count = 0

        # Solo considerar subcomponentes si es FamilyInstance
        if isinstance(element, FamilyInstance):
            try:
                sub_ids = element.GetSubComponentIds()
                if sub_ids:
                    sub_count = len(sub_ids)
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass  # Seguridad adicional si el método no está disponible en algún tipo

        # Procesar elemento principal y subcomponentes
        for elem in elementos_a_procesar:
            param_nivel = getParameter(elem, "S&P_NIVEL DE ELEMENTO")
            nivel = param_nivel.AsString() if param_nivel else ""
            # ==== Obtenemos el valor numerico del nivel en formato string ====
            nivel_number_value = extract_number_nivel(nivel)

            # Imprimir información del elemento (solo para el principal)
            if elem == element:
                print_element_processing(elem, nivel, sub_count)

            parametros_API = {
                "COBie.Attribute.Level" : nivel_number_value
            }

            parametros = dict(parametros_estaticos)
            parametros.update(parametros_API)

            for param, value in parametros.items():
                p = getParameter(elem, param)
                if SetParameter(p, value):
                    total_parameters_assigned += 1
                else:
                    errores.append((elem.Id, param))

            total_processed += 1

    # También aplicar a la Project Information
    output.print_md("🔸 **Procesando Project Information...**")
    for project_param, value in parametros_estaticos.items():
        parameter = getParameter(project_info, project_param)
        if SetParameter(parameter, value):
            total_parameters_assigned += 1
        else:
            errores.append((project_info.Id, project_param))

# Calcular tiempo de procesamiento
end_time = time.time()
processing_time = end_time - start_time

# Imprimir resultados finales
print_results_summary(total_processed, total_parameters_assigned, errores, processing_time)
print_footer()

# Mostrar diálogo final
if errores:
    TaskDialog.Show("⚠️ Completado con errores", 
                   "Proceso completado.\n\n✅ Elementos procesados: {}\n❌ Errores: {}\n\n📋 Consulta el reporte detallado en la consola de PyRevit.".format(total_processed, len(errores)))
else:
    TaskDialog.Show("✅ Proceso exitoso", 
                   "¡Todos los parámetros COBie se completaron correctamente!\n\n📊 Elementos procesados: {}\n🏷️ Parámetros asignados: {}\n⏱️ Tiempo: {:.2f}s".format(total_processed, total_parameters_assigned, processing_time))