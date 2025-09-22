# -- coding: utf-8 --
__title__ = "COBie\nComponent"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import Transaction, ElementId, StorageType, FamilyInstance
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import script, revit

from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._Dictionary import obtener_mantenimiento, obtener_codigo_colegio, obtener_serial_number

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# Obtener el documento activo
doc = revit.doc
ui_doc = revit.uidoc

# Seleccionar referencias de elementos
references = ui_doc.Selection.PickObjects(ObjectType.Element)

# Datos fijos
creado_por = "pruiz@cgeb.com.pe"
fecha = "2025-04-18T16:45:10"
hora_de_ejecucion = "T13:29:49"
fecha_garantia = obtener_mantenimiento()
numero_conts = "775"

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
                "COBie.CreatedBy": str(creado_por),
                "COBie.CreatedOn": str(fecha),
                "COBie.Component.Space": str(ambiente),
                "COBie.Component.Description": str(descripcion),
                "COBie.Component.SerialNumber": str(serial_number),
                "COBie.Component.InstallationDate": str(instalacion) + str(hora_de_ejecucion),
                "COBie.Component.WarrantyStartDate": str(fecha_garantia),
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