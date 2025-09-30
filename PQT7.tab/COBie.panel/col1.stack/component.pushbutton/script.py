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
    "COBie.Component.BarCode",
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

if specialty == "ARQUITECTURA":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -AR')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Arquitectura cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES SANITARIAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - PL')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Sanitarias cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES ELECTRICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  -EE')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Eléctricas cargados:", len(data_list), "filas")

elif specialty == "COMUNICACIONES":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - IICC')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Comunicaciones cargados:", len(data_list), "filas")

elif specialty == "INSTALACIONES MECANICAS":
    excel_instance = Excel()
    excel_rows = excel_instance.read_excel('ESTANDAR COBIE  - ME')
    headers = excel_instance.get_headers(excel_rows, 2)
    headers_required = excel_instance.headers_required(headers, columns_headers)
    data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)
    print("Datos de Mecánicas cargados:", len(data_list), "filas")

else:
    forms.alert("Especialidad '{}' no reconocida para cargar datos Excel.".format(specialty), exitscript=True)

if not data_list:
    forms.alert("No se pudieron cargar los datos del Excel.", exitscript=True)
# ==== Termino de la lectura de Excel ====

count = 0

with revit.Transaction("Transfiere datos a Parametros COBieComponent"):

    for reference in references:
        # ==== Obtenemos el elemento seleccionado ====
        element_object = doc.GetElement(reference)

        elementos_a_procesar = [element_object]

        # ==== Verificar si tiene subcomponentes (solo si es FamilyInstance) ====
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
            
            # ==== Validar el parámetro COBie ====
            cobie = elem.LookupParameter("COBie")
            if not (cobie and not cobie.IsReadOnly and cobie.StorageType == StorageType.Integer and cobie.AsInteger() == 1):
                continue
            
            # ==== Obtenemos el nivel del elemento ====
            level_param_value = get_param_value(getParameter(elem, "S&P_NIVEL DE ELEMENTO"))
            level = divide_string(level_param_value, 1)
            
            # ==== Obtenemos la categoria del elemento y Verificamos que el elemento tenga categoria ====
            elem_category_object = elem.Category
            name_category = elem_category_object.Name if elem_category_object else "Sin categoria"
            

            id_elem = elem.Id.IntegerValue
            
            # ==== Obtenemos el tipo del elemento seleccionado y su nombre ====
            element_type_object_id = elem.GetTypeId()             # => Obtenemos el ElementId del elemento
            if element_type_object_id == ElementId.InvalidElementId:
                print("El elemento id" + str(id_elem) + " no tiene un tipo asociado.")
                continue
            el_type_object = doc.GetElement(element_type_object_id)     # => Obtenemos el Tipo de elemento por medio de su ID
            param_object_type = GetParameterAPI(el_type_object, BuiltInParameter.SYMBOL_NAME_PARAM)
            name_type = get_param_value(param_object_type)
            pr_number = get_param_value(getParameter(elem, "Classification.Uniclass.Pr.Number"))

            # ==== Obtenemos la familia del elemento y Verificamos que exista ====
            family_name = el_type_object.FamilyName if isinstance(el_type_object, ElementType) else "Sin familia"

            # ==== Obtenemos la zonificacion del MBR ====
            zonification_value = get_param_value(getParameter(elem, "S&P_ZONIFICACION"))
            mbr_value = divide_string(zonification_value, 1, compare="sitio", value_default="000")
            
            # ==== Obtenemos el ambiente del elemento ====
            ambiente_object = getParameter(elem, "S&P_AMBIENTE")
            ambiente = get_param_value(ambiente_object)
            
            # ==== Obtenemos el codigo del elemento ====
            code_elem = get_param_value(getParameter(elem, "S&P_CODIGO DE ELEMENTO"))
            if code_elem not in (None, "", "n/a"):
                code_elem
            
            # ==== Obtenemos el parametro partida N°xx de acuerdo a la especialidad ====
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


            parametros = {
                "COBie.Component.Name": "{} : {} : {} : {}".format(name_category, family_name, name_type, id_elem),
                # "COBie.CreatedBy": created_by,
                "COBie.CreatedOn": CREATED_ON,
                "COBie.Component.Space": ambiente,
                "COBie.Component.Description": description,
                "COBie.Component.SerialNumber": "{} {}".format(code_elem, id_elem),
                "COBie.Component.InstallationDate": "",
                "COBie.Component.WarrantyStartDate": warranty_start_date,
                "COBie.Component.TagNumber": "",
                "COBie.Component.BarCode": "",
                "COBie.Component.AssetIdentifier": "{}-ZZ-{}-".format(mbr_value, level)
            }

            for param_name, value in parametros.items():
                param = elem.LookupParameter(param_name)
                if param:
                    SetParameter(param, value)
            count += 1

    TaskDialog.Show("Informativo", "Son {} procesados correctamente\npara COBie component".format(count))

print(data_list)