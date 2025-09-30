# -*- coding: utf-8 -*-
__title__ = "COBie\nComponent"

# ==== Obtenemos la librerias necesarias ====
from Autodesk.Revit.DB import Transaction, ElementId, StorageType, FamilyInstance, ElementType, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import script, revit, forms
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import *
from DBRepositories.SchoolRepository import ColegiosRepository
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository
from Helper._Excel import Excel
from Helper._Dictionary import find_mapped_number
from datetime import datetime, timedelta

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# ==== Metodo para validar valor de parametro vacío o n/a ====
def get_first_valid_parameter(elem, names):
    """Devuelve el primer valor válido encontrado en la lista de parámetros."""
    invalid_values = {None, ""}
    for name in names:
        param = getParameter(elem, name)
        if param:
            value = get_param_value(param)
            if isinstance(value, str):
                if value.strip().lower() in {"", "n/a"}:
                    continue
            if value not in invalid_values:
                return value
    return None


# ==== Metodo para dividir una cadena ====
def divide_string(text, idx, character_divider=None, compare=None, value_default=None):
    """
    Divide un string en partes usando un separador y devuelve el elemento en la posición idx.
    Si el texto completo coincide con 'compare', devuelve 'value_default'.
    """
    if not text:
        return ""
    
    # Verificar coincidencia antes de dividir
    if compare and text.strip().lower() == compare.lower():
        return value_default
    
    parts = text.split(character_divider) if character_divider else text.split()
    if idx < 0 or idx >= len(parts):
        return ""
    
    return parts[idx]

columns_headers = [
    "COBie.Component.InstallationDate",
    "COBie.Component.Description",
    "CODIGO"
]

def find_data_by_code(data_list, codigo_elemento):
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

# ==== Obtener el documento activo ====
doc = revit.doc
ui_doc = revit.uidoc

# ==== Seleccionar elementos y obtenemos tipo Reference de los elementos ====
try:
    references = ui_doc.Selection.PickObjects(ObjectType.Element)
except OperationCanceledException:
    forms.alert("Operación cancelada: no se seleccionaron elementos para procesar COBie.Component",
                title="Cancelación")
    script.exit()

# ==== Datos fijos ====
CREATED_ON = "2025-08-04T11:59:30"


# ==== Instanciamos el colegio correspondiente del modelo activo ====
school_repo_object = ColegiosRepository()
school_object = school_repo_object.codigo_colegio(doc)
school = None
created_by = None
warranty_start_date = None

# ==== Verificamos si el colegio existe
if school_object:
    school = school_object.name
    created_by = school_object.created_by
    warranty_start_date = school_object.warranty_start_date.component_warranty

# ==== Instanciamos la especialidad correspondiente al modelo ====
specialty_repo = SpecialtiesRepository()
specialty_object = specialty_repo.get_specialty_by_document(doc)
specialty = None

if specialty_object:
    specialty = specialty_object.name

# ==== Obtenemos la hoja excel de acuerdo a la especialidad ====
data_list = None

print("\n" + "="*70)
print("PROCESAMIENTO COBie COMPONENT - {}".format(specialty))
print("="*70)

if specialty == "ARQUITECTURA":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -AR')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

elif specialty == "INSTALACIONES SANITARIAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - PL')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

elif specialty == "INSTALACIONES ELECTRICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -EE')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

elif specialty == "COMUNICACIONES":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - IICC')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

elif specialty == "INSTALACIONES MECANICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - ME')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

else:
    forms.alert("Especialidad '{}' no reconocida para cargar datos Excel.".format(specialty), exitscript=True)

if not data_list:
    forms.alert("No se pudieron cargar los datos del Excel.", exitscript=True)

print("[OK] Excel cargado: {} registros disponibles".format(len(data_list)))

# ==== Creamos nuevo diccionario que contendra el codigo y sus valores ====
dict_codigos = {}
for row in data_list:
    code = row["CODIGO"]
    dict_codigos[code] = row

# ==== Contadores para estadísticas ====
count = 0
fechas_actualizadas = 0
serial_actualizados = 0
elementos_sin_codigo = 0
errores = []

print("\nIniciando procesamiento de elementos...")
print("-"*70)

with revit.Transaction("Transfiere datos a Parametros COBieComponent"):

    for reference in references:
        element_object = doc.GetElement(reference)
        elementos_a_procesar = [element_object]

        if isinstance(element_object, FamilyInstance):
            try:
                sub_ids = element_object.GetSubComponentIds()
                if sub_ids:
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass

        for elem in elementos_a_procesar:
            uid_elem = elem.UniqueId
            
            cobie = elem.LookupParameter("COBie")
            if not (cobie and not cobie.IsReadOnly and cobie.StorageType == StorageType.Integer and cobie.AsInteger() == 1):
                continue
            
            level_param_value = get_param_value(getParameter(elem, "S&P_NIVEL DE ELEMENTO"))
            level = divide_string(level_param_value, 1)
            
            elem_category_object = elem.Category
            name_category = elem_category_object.Name if elem_category_object else "Sin categoria"
            
            id_elem = elem.Id.IntegerValue
            
            element_type_object_id = elem.GetTypeId()
            if element_type_object_id == ElementId.InvalidElementId:
                errores.append("Elemento ID {} sin tipo asociado".format(id_elem))
                continue
            
            el_type_object = doc.GetElement(element_type_object_id)
            param_object_type = GetParameterAPI(el_type_object, BuiltInParameter.SYMBOL_NAME_PARAM)
            name_type = get_param_value(param_object_type)
            pr_number = get_param_value(getParameter(el_type_object, "Classification.Uniclass.Pr.Number"))

            family_name = el_type_object.FamilyName if isinstance(el_type_object, ElementType) else "Sin familia"

            zonification_value = get_param_value(getParameter(elem, "S&P_ZONIFICACION"))
            mbr_value = divide_string(zonification_value, 1, compare="sitio", value_default="000")
            
            # ==== Obtenemos el ambiente en el component space
            tag_number = 0
            space_component = getParameter(elem, "COBie.Component.Space")
            space = get_param_value(space_component)
            if "," in space:
                tag_number_separate = divide_string(space, 0, ",")
                tag_number = find_mapped_number(tag_number_separate)
            else:
                tag_number = find_mapped_number(space)
            
            code_elem = get_param_value(getParameter(elem, "S&P_CODIGO DE ELEMENTO"))
            if code_elem not in (None, "", "n/a"):
                code_elem
            else:
                elementos_sin_codigo += 1
            
            if specialty in ["INSTALACIONES SANITARIAS", "COMUNICACIONES"]:
                description = get_first_valid_parameter(
                    elem,
                    ["S&P_DESCRIPCION PARTIDA N°2", "S&P_DESCRIPCION PARTIDA N°1"]
                )
            else:
                description = get_first_valid_parameter(
                    elem,
                    ["S&P_DESCRIPCION PARTIDA N°1"]
                )
            
            # ==== CONVERSION DE FECHA DEL EXCEL ====
            if code_elem in dict_codigos:
                data_row = dict_codigos[code_elem]
                
                if "COBie.Component.InstallationDate" in data_row:
                    fecha_excel = data_row["COBie.Component.InstallationDate"]
                    
                    if fecha_excel and fecha_excel not in ("", "n/a", None):
                        param_installation_date = getParameter(elem, "COBie.Component.InstallationDate")
                        
                        if param_installation_date and not param_installation_date.IsReadOnly:
                            try:
                                if isinstance(fecha_excel, (int, float)):
                                    fecha_base = datetime(1899, 12, 30)
                                    fecha_convertida = fecha_base + timedelta(days=float(fecha_excel))
                                    fecha_formateada = fecha_convertida.strftime("%Y-%m-%d")
                                    
                                elif hasattr(fecha_excel, 'strftime'):
                                    fecha_formateada = fecha_excel.strftime("%Y-%m-%d")
                                    
                                elif isinstance(fecha_excel, str):
                                    fecha_formateada = fecha_excel
                                    
                                else:
                                    fecha_formateada = str(fecha_excel)
                                
                                SetParameter(param_installation_date, fecha_formateada)
                                fechas_actualizadas += 1
                                
                            except Exception as e:
                                errores.append("Elemento {}: Error en fecha - {}".format(id_elem, str(e)))

            # ==== Description desde Excel ====
            if "COBie.Component.Description" in data_row:
                desc_excel = data_row["COBie.Component.Description"]
                if desc_excel and str(desc_excel).strip().lower() not in ("", "n/a"):
                    param_desc = getParameter(elem, "COBie.Component.Description")
                    if param_desc and not param_desc.IsReadOnly:
                        try:
                            SetParameter(param_desc, str(desc_excel))
                        except Exception as e:
                            errores.append("Elemento {}: Error en descripción - {}".format(id_elem, str(e)))

            # ==== Verificar si SerialNumber está vacío antes de asignar ====
            param_serial = elem.LookupParameter("COBie.Component.SerialNumber")
            serial_value = get_param_value(param_serial) if param_serial else None
            
            # Solo asignar si está vacío o es None
            if not serial_value or serial_value.strip() in ("", "n/a"):
                serial_number_value = "{} {}".format(code_elem, id_elem)
                serial_actualizados += 1
            else:
                serial_number_value = serial_value  # Mantener el valor existente

            parametros = {
                "COBie.Component.Name": "{} : {} : {} : {}".format(name_category, family_name, name_type, id_elem),
                "COBie.CreatedOn": CREATED_ON,
                # "COBie.Component.Space": ambiente,
                "COBie.CreatedBy": created_by,
                "COBie.Component.SerialNumber": serial_number_value,
                "COBie.Component.WarrantyStartDate": warranty_start_date,
                "COBie.Component.TagNumber": tag_number,
                "COBie.Component.BarCode": "{}{}".format(mbr_value, id_elem),
                "COBie.Component.AssetIdentifier": "{}-ZZ-{}-{}-{}-{}".format(mbr_value, level, tag_number,pr_number,mbr_value+str(id_elem) )
            }

            for param_name, value in parametros.items():
                param = elem.LookupParameter(param_name)
                if param:
                    SetParameter(param, value)
            count += 1

# ==== RESUMEN DE PROCESAMIENTO ====
print("\n" + "="*70)
print("RESUMEN DE PROCESAMIENTO")
print("="*70)
print("Elementos procesados:        {}".format(count))
print("Fechas actualizadas:         {}".format(fechas_actualizadas))
print("SerialNumbers actualizados:  {}".format(serial_actualizados))
if elementos_sin_codigo > 0:
    print("Elementos sin codigo:        {} (ADVERTENCIA)".format(elementos_sin_codigo))
if errores:
    print("\nERRORES ENCONTRADOS ({})".format(len(errores)))
    print("-"*70)
    for error in errores[:10]:  # Mostrar máximo 10 errores
        print("  - {}".format(error))
    if len(errores) > 10:
        print("  ... y {} errores más".format(len(errores) - 10))
else:
    print("Errores:                     0")
print("="*70 + "\n")

TaskDialog.Show("COBie Component", 
    "Procesamiento completado\n\n" +
    "Elementos procesados: {}\n".format(count) +
    "Fechas actualizadas: {}\n".format(fechas_actualizadas) +
    "SerialNumbers actualizados: {}".format(serial_actualizados))