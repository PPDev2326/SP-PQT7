# -- coding: utf-8 --
__title__ = "COBie\nComponent"

# ==== Librerías necesarias ====
from Autodesk.Revit.DB import Transaction, ElementId, StorageType, FamilyInstance, ElementType, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import script, revit, forms
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._RevitAPI import GetParameterAPI, get_param_value, getParameter
from DBRepositories.SchoolRepository import ColegiosRepository
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository

# ==== Configuración del logger ====
logger = script.get_logger()
logger.set_level('DEBUG')  # Opciones: DEBUG, INFO, WARNING, ERROR

# ==== Validar nombre del archivo ====
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
    if compare and text.strip().lower() == compare.lower():
        return value_default
    parts = text.split(character_divider) if character_divider else text.split()
    if idx < 0 or idx >= len(parts):
        return ""
    return parts[idx]

# ==== Obtener el documento activo ====
doc = revit.doc
ui_doc = revit.uidoc

# ==== Seleccionar elementos ====
try:
    references = ui_doc.Selection.PickObjects(ObjectType.Element)
except OperationCanceledException:
    forms.alert("Operación cancelada: no se seleccionaron elementos para procesar COBie.Component",
                title="Cancelación")
    script.exit()

# ==== Datos fijos ====
CREATED_ON = "2025-08-04T11:59:30"

# ==== Colegio del modelo ====
school_repo_object = ColegiosRepository()
school_object = school_repo_object.codigo_colegio(doc)
school = created_by = warranty_start_date = None
if school_object:
    school = school_object.name
    created_by = school_object.created_by
    warranty_start_date = school_object.warranty_start_date.component_warranty

# ==== Especialidad del modelo ====
specialty_repo = SpecialtiesRepository()
specialty_object = specialty_repo.get_specialty_by_document(doc)
specialty = specialty_object.name if specialty_object else None

# ==== Contadores ====
count_ok = 0
count_skipped = 0
count_failed = 0
errores = []

with revit.Transaction("Transfiere datos a Parametros COBieComponent"):

    for reference in references:
        element_object = doc.GetElement(reference)
        elementos_a_procesar = [element_object]

        # ==== Verificar subcomponentes ====
        if isinstance(element_object, FamilyInstance):
            try:
                sub_ids = element_object.GetSubComponentIds()
                if sub_ids:
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except Exception:
                pass

        for elem in elementos_a_procesar:
            id_elem = elem.Id.IntegerValue
            logger.info("=== Procesando elemento ID {} ===".format(id_elem))

            # ==== Validar el parámetro COBie ====
            cobie = elem.LookupParameter("COBie")
            if not (cobie and not cobie.IsReadOnly and
                    cobie.StorageType == StorageType.Integer and
                    cobie.AsInteger() == 1):
                logger.debug("Elemento ID {} omitido: no cumple condición COBie.".format(id_elem))
                count_skipped += 1
                continue

            # ==== Nivel del elemento ====
            level_param_value = get_param_value(getParameter(elem, "NIVEL DE ELEMENTO"))
            level = divide_string(level_param_value, 1)

            # ==== Tipo del elemento ====
            element_type_object_id = elem.GetTypeId()
            if element_type_object_id == ElementId.InvalidElementId:
                logger.warning("Elemento ID {} no tiene tipo asociado.".format(id_elem))
                count_failed += 1
                continue

            el_type_object = doc.GetElement(element_type_object_id)
            name_type = get_param_value(GetParameterAPI(el_type_object, BuiltInParameter.SYMBOL_NAME_PARAM))
            family_name = el_type_object.FamilyName if isinstance(el_type_object, ElementType) else "Sin familia"

            # ==== Zonificación MBR ====
            zonification_value = get_param_value(getParameter(elem, "S&P_ZONIFICACION"))
            mbr_value = divide_string(zonification_value, 1, compare="sitio", value_default="000")

            # ==== Ambiente ====
            ambiente = get_param_value(getParameter(elem, "S&P_AMBIENTE"))

            # ==== Código ====
            code_elem = get_param_value(getParameter(elem, "S&P_CODIGO DE ELEMENTO"))

            # ==== Descripción ====
            if specialty in ["INSTALACIONES SANITARIAS", "COMUNICACIONES"]:
                description = get_first_valid_parameter(
                    elem, ["S&P_DESCRIPCION PARTIDA N°2", "S&P_DESCRIPCION PARTIDA N°1"]
                )
            else:
                description = get_first_valid_parameter(elem, ["S&P_DESCRIPCION PARTIDA N°1"])

            # ==== Parámetros a transferir ====
            parametros = {
                "COBie.Component.Name": "{} : {} : {} : {}".format(
                    elem.Category.Name if elem.Category else "Sin categoría",
                    family_name,
                    name_type,
                    id_elem
                ),
                "COBie.CreatedBy": created_by,
                "COBie.CreatedOn": CREATED_ON,
                "COBie.Component.Space": ambiente,
                "COBie.Component.Description": description,
                "COBie.Component.SerialNumber": "{} {}".format(code_elem, id_elem),
                "COBie.Component.WarrantyStartDate": warranty_start_date,
                "COBie.Component.AssetIdentifier": "{}-ZZ-{}-".format(mbr_value, level),
            }

            # ==== Setear parámetros ====
            for param_name, value in parametros.items():
                param = elem.LookupParameter(param_name)
                if not param:
                    errores.append("Elemento ID {} no tiene parámetro '{}'".format(id_elem, param_name))
                    logger.error(errores[-1])
                    count_failed += 1
                    continue
                if param.IsReadOnly:
                    errores.append("Parámetro '{}' en elemento ID {} es de solo lectura".format(param_name, id_elem))
                    logger.warning(errores[-1])
                    count_failed += 1
                    continue
                if value in (None, ""):
                    logger.debug("Parámetro '{}' vacío en elemento ID {}".format(param_name, id_elem))
                    continue
                try:
                    param.Set(value)
                    logger.info("✔ '{}' = '{}' en elemento ID {}".format(param_name, value, id_elem))
                except Exception as ex:
                    errores.append("Error al asignar '{}' en elemento ID {}: {}".format(param_name, id_elem, ex))
                    logger.error(errores[-1])
                    count_failed += 1
            count_ok += 1

# ==== Resumen final ====
logger.info("==== RESUMEN FINAL ====")
logger.info("Procesados correctamente: {}".format(count_ok))
logger.info("Omitidos (sin COBie): {}".format(count_skipped))
logger.info("Fallidos (errores o solo lectura): {}".format(count_failed))

if errores:
    logger.warning("Lista de errores:")
    for err in errores:
        logger.warning(" - {}".format(err))

TaskDialog.Show("Informativo", "COBie Component\nProcesados: {}\nOmitidos: {}\nFallidos: {}".format(
    count_ok, count_skipped, count_failed
))
