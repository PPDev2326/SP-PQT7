# -- coding: utf-8 --
__title__ = "COBie\nComponent"

# ==== Obtenemos la librerias necesarias ====
from Autodesk.Revit.DB import Transaction, ElementId, StorageType, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import script, revit, forms
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from DBRepositories.SchoolRepository import ColegiosRepository

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# ==== Obtener el documento activo ====
doc = revit.doc
ui_doc = revit.uidoc

# ==== Seleccionar referencias de elementos ====
try:
    references = ui_doc.Selection.PickObjects(ObjectType.Element)
except OperationCanceledException:
    forms.alert("Operación cancelada: no se seleccionaron elementos para procesar COBie.Component",
                title="Cancelación")

# ==== Datos fijos ====
CREATED_ON = "2025-08-04T11:59:30"


# ==== Instanciamos el colegio correspondiente del modelo activo ====
school_repo_object = ColegiosRepository()
school_object = school_repo_object.codigo_colegio(doc)
school = None
created_by = None

if school_object:
    school = school_object.name
    created_by = school_object.created_by

count = 0

with revit.Transaction("Transfiere datos a Parametros COBieComponent"):

    for reference in references:
        element = doc.GetElement(reference)
        uniqueId = element.UniqueId
        id_param = element.Id.IntegerValue

        # Validar el parámetro COBie
        cobie = element.LookupParameter("COBie")
        if not (cobie and not cobie.IsReadOnly and cobie.StorageType == StorageType.Integer and cobie.AsInteger() == 1):
            continue

        # Obtener tipo de elemento
        element_type_id = element.GetTypeId()
        if element_type_id == ElementId.InvalidElementId:
            print("El elemento " + str(id_param) + " no tiene un tipo asociado.")
            continue
        element_type = doc.GetElement(element_type_id)

        # Obtener valores de parámetros del elemento y del tipo
        def obtener_valor(elem, nombre_parametro):
            param = elem.LookupParameter(nombre_parametro)
            if param and param.StorageType == StorageType.String:
                return param.AsString() or ""
            return ""

        elementos_a_procesar = [element]

        # Verificar si tiene subcomponentes (solo si es FamilyInstance)
        if isinstance(element, FamilyInstance):
            try:
                sub_ids = element.GetSubComponentIds()
                if sub_ids:
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass

        for elem in elementos_a_procesar:
            id_elem = elem.Id.IntegerValue
            uid_elem = elem.UniqueId
            nrm = obtener_valor(elem, "Descripción de partida NRM")
            ambiente = obtener_valor(elem, "Ambiente")
            descripcion = obtener_valor(elem, "Descripción de elemento")
            instalacion = obtener_valor(elem, "Fecha ejecutada")
            activo = obtener_valor(elem, "Activo")
            serial_number = obtener_serial_number(activo, str(id_elem))

            parametros = {
                "COBie.Component.Name": str(id_elem) + "_" + str(nrm),
                "COBie.CreatedBy": str(CREATED_ON),
                "COBie.CreatedOn": str(DATE),
                "COBie.Component.Space": str(ambiente),
                "COBie.Component.Description": str(descripcion),
                "COBie.Component.SerialNumber": str(serial_number),
                "COBie.Component.InstallationDate": str(instalacion) + str(EXECUTION_TIME),
                "COBie.Component.WarrantyStartDate": str(WARRANTY_DATE),
                "COBie.Component.TagNumber": str(id_elem),
                "COBie.Component.BarCode": str(numero_conts) + str(id_elem),
                "COBie.Component.AssetIdentifier": str(uid_elem)
            }

            for param_name, value in parametros.items():
                param = elem.LookupParameter(param_name)
                if param and not param.IsReadOnly:
                    param.Set(value)
            count += 1

    TaskDialog.Show("Informativo", "Son {} procesados correctamente\npara COBie component".format(count))