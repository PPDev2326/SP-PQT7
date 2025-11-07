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

# ==== Variable global para controlar modo sobrescritura ====
SOBRESCRIBIR_GLOBAL = True

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


def esta_vacio(param):
    """
    Verifica si un parámetro está vacío o tiene valores no válidos.
    """
    if not param:
        return True
    
    value = get_param_value(param)
    
    if value is None:
        return True
    
    if isinstance(value, str):
        value_clean = value.strip().lower()
        if value_clean in ("", "n/a", "none"):
            return True
    
    return False


def asignar_parametro_seguro(param, value):
    """
    Asigna un valor a un parámetro de forma segura (recibe el parámetro ya obtenido).
    Retorna True si se asignó correctamente, False si no.
    """
    if value is None or (isinstance(value, str) and value.strip() in ("", "n/a")):
        return False
    
    if not param or param.IsReadOnly:
        return False
    
    # Verificar si debe sobrescribir o solo llenar vacíos
    if not SOBRESCRIBIR_GLOBAL:
        # Solo asignar si está vacío
        if not esta_vacio(param):
            return False
    
    try:
        SetParameter(param, value)
        return True
    except Exception:
        return False


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

columns_space = [
    "COBie.Space.Name",
    "COBie.Space.RoomTag",
]

def get_roomtag_from_cobie_space(cobie_space_value, space_data_dict):
    """
    Obtiene el RoomTag desde los datos de SPACE ya cargados.
    Recibe el valor directamente en lugar del elemento.
    """
    if not cobie_space_value:
        return "0"
    
    # Normalizar el valor
    cobie_space_value_clean = str(cobie_space_value).strip()
    
    if not cobie_space_value_clean or cobie_space_value_clean == "0":
        return "0"
    
    # Si hay múltiples espacios separados por coma, tomar solo el primero
    if "," in cobie_space_value_clean:
        cobie_space_value_clean = cobie_space_value_clean.split(",")[0].strip()
    
    # Buscar en el diccionario
    if cobie_space_value_clean in space_data_dict:
        room_tag = space_data_dict[cobie_space_value_clean].get("COBie.Space.RoomTag", "0")
        return room_tag if room_tag else "0"
    
    return "0"

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

# ==== PREGUNTA: SOBRESCRIBIR ====================
sobrescribir = forms.alert(
    "¿Deseas SOBRESCRIBIR los parámetros que ya tienen valores?\n\n"
    "SI: Sobrescribir todos (actualizar valores existentes)\n"
    "NO: Solo llenar parámetros vacíos (cada parámetro se evalúa individualmente)",
    title="Modo de sobrescritura COBie Component",
    yes=True,
    no=True
)

# Configurar variable global
SOBRESCRIBIR_GLOBAL = sobrescribir

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

print("\n" + "="*70)
print("PROCESAMIENTO COBie COMPONENT - {}".format(specialty))
print("="*70)
if sobrescribir:
    print("MODO: Sobrescribir valores existentes")
else:
    print("MODO: Solo llenar parametros vacios (individualmente)")
print("="*70)

# ==== CARGAR EXCEL UNA SOLA VEZ ====
print("\n[INFO] Preparando carga de datos...")

# Determinar nombre de hoja según especialidad
specialty_to_sheet = {
    "ARQUITECTURA": "ESTANDAR COBIE  -AR",
    "INSTALACIONES SANITARIAS": "ESTANDAR COBIE  - PL",
    "INSTALACIONES ELECTRICAS": "ESTANDAR COBIE  -EE",
    "COMUNICACIONES": "ESTANDAR COBIE  - IICC",
    "INSTALACIONES MECANICAS": "ESTANDAR COBIE  - ME"
}

sheet_name = specialty_to_sheet.get(specialty)
if not sheet_name:
    forms.alert("Especialidad '{}' no reconocida para cargar datos Excel.".format(specialty), exitscript=True)

# Crear instancia de Excel (esto pedirá el archivo UNA SOLA VEZ)
excel_instance = Excel()

# Cargar datos de COMPONENT
print("[INFO] Seleccione el archivo Excel...")
print("[INFO] Cargando hoja '{}'...".format(sheet_name))
excel_rows = excel_instance.read_excel(sheet_name)
headers = excel_instance.get_headers(excel_rows, 2)
headers_required = excel_instance.headers_required(headers, columns_headers)
data_list = excel_instance.get_data_by_headers_required(excel_rows, headers_required, 3)

if not data_list:
    forms.alert("No se pudieron cargar los datos del Excel.", exitscript=True)

print("[OK] Excel cargado: {} registros de COMPONENT disponibles".format(len(data_list)))

# Crear diccionario de códigos para COMPONENT
dict_codigos = {}
for row in data_list:
    code = row["CODIGO"]
    dict_codigos[code] = row

# Cargar datos de SPACE (del mismo archivo Excel ya abierto - NO PEDIRÁ EL ARCHIVO DE NUEVO)
print("[INFO] Cargando hoja 'ESTANDAR COBie SPACE' del mismo archivo...")
space_rows = excel_instance.read_excel('ESTANDAR COBie SPACE ')
if not space_rows:
    forms.alert("No se encontró hoja de Excel 'ESTANDAR COBie SPACE'.", exitscript=True)

space_headers = excel_instance.get_headers(space_rows, 2)
space_headers_required = excel_instance.headers_required(space_headers, columns_space)
space_data = excel_instance.get_data_by_headers_required(space_rows, space_headers_required, 3)

if not space_data:
    forms.alert("No se pudieron cargar los datos de SPACE del Excel.", exitscript=True)

# Crear diccionario de SPACE
dict_space = {}
for row in space_data:
    code = row.get("COBie.Space.Name")
    if code:
        code_clean = str(code).strip()
        dict_space[code_clean] = row

print("[OK] Datos de SPACE cargados: {} registros disponibles".format(len(dict_space)))

# ==== PRE-COMPUTAR valores que se necesitarán (OPTIMIZACIÓN) ====
SPECIALTY_USES_PARTIDA2 = specialty in ["INSTALACIONES SANITARIAS", "COMUNICACIONES"]

# Nombres de parámetros para descripción según especialidad
if SPECIALTY_USES_PARTIDA2:
    DESCRIPCION_PARAMS = ["S&P_DESCRIPCION PARTIDA N°2", "S&P_DESCRIPCION PARTIDA N°1"]
else:
    DESCRIPCION_PARAMS = ["S&P_DESCRIPCION PARTIDA N°1"]

# Set de niveles de techo
NIVELES_TECHO = {"TECHO", "Techo", "techo", "Cubierta", "CUBIERTA", "cubierta"}

# ==== Contadores para estadísticas ====
count = 0
stats = {
    "InstallationDate": 0,
    "Description": 0,
    "SerialNumber": 0,
    "Name": 0,
    "CreatedOn": 0,
    "CreatedBy": 0,
    "WarrantyStartDate": 0,
    "TagNumber": 0,
    "BarCode": 0,
    "AssetIdentifier": 0
}
elementos_sin_codigo = 0
elementos_ignorados = 0
errores = []

print("\nIniciando procesamiento de elementos...")
print("-"*70)

# ==== CACHE DE TIPOS (OPTIMIZACIÓN) ====
type_cache = {}

def get_type_cached(type_id):
    """Obtiene tipo desde cache para evitar llamadas repetidas a GetElement"""
    if type_id not in type_cache:
        type_cache[type_id] = doc.GetElement(type_id)
    return type_cache[type_id]

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
            # Verificar COBie primero (rápido)
            cobie = elem.LookupParameter("COBie")
            if not (cobie and not cobie.IsReadOnly and cobie.StorageType == StorageType.Integer and cobie.AsInteger() == 1):
                elementos_ignorados += 1
                continue
            
            # ==== OBTENER PARÁMETROS UNA SOLA VEZ (OPTIMIZACIÓN) ====
            param_nivel = getParameter(elem, "S&P_NIVEL DE ELEMENTO")
            param_zonificacion = getParameter(elem, "S&P_ZONIFICACION")
            param_codigo = getParameter(elem, "S&P_CODIGO DE ELEMENTO")
            param_cobie_space = getParameter(elem, "COBie.Component.Space")
            
            # Parámetros COBie (obtener referencias, no valores)
            param_installation_date = elem.LookupParameter("COBie.Component.InstallationDate")
            param_description = elem.LookupParameter("COBie.Component.Description")
            param_serial = elem.LookupParameter("COBie.Component.SerialNumber")
            param_name = elem.LookupParameter("COBie.Component.Name")
            param_created_on = elem.LookupParameter("COBie.CreatedOn")
            param_created_by = elem.LookupParameter("COBie.CreatedBy")
            param_warranty = elem.LookupParameter("COBie.Component.WarrantyStartDate")
            param_tag = elem.LookupParameter("COBie.Component.TagNumber")
            param_barcode = elem.LookupParameter("COBie.Component.BarCode")
            param_asset = elem.LookupParameter("COBie.Component.AssetIdentifier")
            
            # ==== OBTENER VALORES ====
            level_param_value = get_param_value(param_nivel)
            level = "RF" if level_param_value in NIVELES_TECHO else divide_string(level_param_value, 1)
            
            elem_category_object = elem.Category
            name_category = elem_category_object.Name if elem_category_object else "Sin categoria"
            
            id_elem = elem.Id.IntegerValue
            
            # Obtener tipo (usando cache)
            element_type_object_id = elem.GetTypeId()
            if element_type_object_id == ElementId.InvalidElementId:
                errores.append("Elemento ID {} sin tipo asociado".format(id_elem))
                continue
            
            el_type_object = get_type_cached(element_type_object_id)
            param_object_type = GetParameterAPI(el_type_object, BuiltInParameter.SYMBOL_NAME_PARAM)
            name_type = get_param_value(param_object_type)
            pr_number = get_param_value(getParameter(el_type_object, "Classification.Uniclass.Pr.Number"))

            family_name = el_type_object.FamilyName if isinstance(el_type_object, ElementType) else "Sin familia"

            zonification_value = get_param_value(param_zonificacion)
            mbr_value = divide_string(zonification_value, 1, compare="sitio", value_default="000")
            
            # ==== Obtenemos el ambiente (optimizado - pasa valor directamente) ====
            cobie_space_value = get_param_value(param_cobie_space)
            tag_number = get_roomtag_from_cobie_space(cobie_space_value, dict_space)
            
            code_elem = get_param_value(param_codigo)
            if code_elem in (None, "", "n/a"):
                elementos_sin_codigo += 1
                code_elem = ""
            
            # ==== CONVERSION DE FECHA DEL EXCEL ====
            if code_elem and code_elem in dict_codigos:
                data_row = dict_codigos[code_elem]
                
                if "COBie.Component.InstallationDate" in data_row:
                    fecha_excel = data_row["COBie.Component.InstallationDate"]
                    
                    if fecha_excel and fecha_excel not in ("", "n/a", None):
                        try:
                            if isinstance(fecha_excel, (int, float)):
                                fecha_base = datetime(1899, 12, 30)
                                fecha_convertida = fecha_base + timedelta(days=float(fecha_excel))
                                fecha_formateada = fecha_convertida.strftime("%Y-%m-%d")
                            elif hasattr(fecha_excel, 'strftime'):
                                fecha_formateada = fecha_excel.strftime("%Y-%m-%d")
                            else:
                                fecha_formateada = str(fecha_excel)
                            
                            if asignar_parametro_seguro(param_installation_date, fecha_formateada):
                                stats["InstallationDate"] += 1
                        except Exception as e:
                            errores.append("Elemento {}: Error en fecha - {}".format(id_elem, str(e)))
                
                # Description desde Excel
                if "COBie.Component.Description" in data_row:
                    desc_excel = data_row["COBie.Component.Description"]
                    if desc_excel and str(desc_excel).strip().lower() not in ("", "n/a"):
                        if asignar_parametro_seguro(param_description, str(desc_excel)):
                            stats["Description"] += 1

            # ==== SerialNumber ====
            serial_number_value = "{} {}".format(code_elem, id_elem)
            if asignar_parametro_seguro(param_serial, serial_number_value):
                stats["SerialNumber"] += 1

            # ==== Asignar parámetros restantes (usando referencias ya obtenidas) ====
            name_value = "{} : {} : {} : {}".format(name_category, family_name, name_type, id_elem)
            if asignar_parametro_seguro(param_name, name_value):
                stats["Name"] += 1
            
            if asignar_parametro_seguro(param_created_on, CREATED_ON):
                stats["CreatedOn"] += 1
            
            if asignar_parametro_seguro(param_created_by, created_by):
                stats["CreatedBy"] += 1
            
            if asignar_parametro_seguro(param_warranty, warranty_start_date):
                stats["WarrantyStartDate"] += 1
            
            if asignar_parametro_seguro(param_tag, tag_number):
                stats["TagNumber"] += 1
            
            barcode_value = "{}{}".format(mbr_value, id_elem)
            if asignar_parametro_seguro(param_barcode, barcode_value):
                stats["BarCode"] += 1
            
            asset_value = "{}-ZZ-{}-{}-{}-{}".format(mbr_value, level, tag_number, pr_number, mbr_value+str(id_elem))
            if asignar_parametro_seguro(param_asset, asset_value):
                stats["AssetIdentifier"] += 1
            
            count += 1

# ==== RESUMEN DE PROCESAMIENTO ====
print("\n" + "="*70)
print("RESUMEN DE PROCESAMIENTO")
print("="*70)
print("Elementos procesados:          {}".format(count))
print("Elementos ignorados (COBie=0): {}".format(elementos_ignorados))
print("")
print("PARAMETROS ACTUALIZADOS:")
for param_name, cantidad in sorted(stats.items()):
    if cantidad > 0:
        print("  - {:<25} {}".format(param_name + ":", cantidad))

if elementos_sin_codigo > 0:
    print("\nADVERTENCIAS:")
    print("  - Elementos sin codigo:      {} (ADVERTENCIA)".format(elementos_sin_codigo))

if errores:
    print("\nERRORES ENCONTRADOS ({})".format(len(errores)))
    print("-"*70)
    for error in errores[:10]:
        print("  - {}".format(error))
    if len(errores) > 10:
        print("  ... y {} errores más".format(len(errores) - 10))
else:
    print("\nErrores:                       0")
print("="*70 + "\n")

# Mensaje final
mensaje_final = "Procesamiento completado\n\n"
mensaje_final += "Elementos procesados: {}\n".format(count)
mensaje_final += "\nParametros actualizados:\n"
for param_name, cantidad in sorted(stats.items()):
    if cantidad > 0:
        mensaje_final += "  {}: {}\n".format(param_name, cantidad)

TaskDialog.Show("COBie Component", mensaje_final)