# -*- coding: utf-8 -*-
__title__ = "Model number"
__doc__ = """Version = 1.0
Date = 03.04.2025
------------------------------------------------------------------
Description:
Este script permite seleccionar elementos visibles en la vista actual
y extraer valores de distintos parámetros. Luego, unifica información
en los parámetros COBie.Type.ModelNumber y COBie.Type.ModelReference,
asegurando que cada tipo de elemento reciba un código único basado en 
la configuración del proyecto.

------------------------------------------------------------------
How-to:
-> Click en el botón
-> Seleccionar elementos
-> Dar click en finalizar

------------------------------------------------------------------
Last update:
- [03.04.2025] - 1.0 UPDATE - New Feature

------------------------------------------------------------------
Author: Paolo Perez"""   

from pyrevit import forms, script
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import ObjectType
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre

nombre_archivo = obtener_nombre_archivo()

if not validar_nombre(nombre_archivo):
    script.exit()

else:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    # Obtener la categoría "Información del Proyecto"
    project_information = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
    
    if not project_information:
        forms.alert("No se encontró la información del proyecto.", exitscript=True)
    
    # Leer parámetros del proyecto
    site_name_param = project_information.LookupParameter("SiteName")
    site_object_param = project_information.LookupParameter("SiteObjectType")
    
    # Validación: Asegurar que los parámetros existan y tengan valor
    if not site_name_param or not site_object_param:
        forms.alert("Los parámetros 'SiteName' y 'SiteObjectType' no existen en la información del proyecto.", exitscript=True)
    
    param_sitename = site_name_param.AsString() if site_name_param.AsString() else None
    param_activo = site_object_param.AsString() if site_object_param.AsString() else None
    
    if not param_sitename or not param_activo:
        forms.alert("Los parámetros 'SiteName' y/o 'SiteObjectType' no tienen información. Completar antes de ejecutar.", exitscript=True)
    
    # Extraer solo el número del parámetro "SiteObjectType"
    new_param_activo = param_activo.split()[-1]
    param_Pr = "Pr_40_50"

    # Seleccionar elementos
    references = uidoc.Selection.PickObjects(ObjectType.Element, "Selecciona los elementos en la vista")
    sel_element = [doc.GetElement(reference) for reference in references]

    # Filtrar tipos únicos (usando un conjunto para evitar duplicados)
    tipos_unicos = set(elemento.GetTypeId() for elemento in sel_element)
    
    # Mostrar el mensaje flotante con la cantidad de tipos únicos
    types_count = len(tipos_unicos)
    forms.alert("Total de tipos únicos seleccionados: {}".format(types_count), title="Información de Tipos")

    t = Transaction(doc, "Actualizar parámetros COBie")
    t.Start()

    try:
        for tipo_id in tipos_unicos:
            # Obtener el tipo de elemento único
            element_type = doc.GetElement(tipo_id)
            if element_type is None:
                print("El tipo de elemento con ID", tipo_id.IntegerValue, "no existe.")
                continue  # Si el tipo no existe, pasa al siguiente tipo
            
            # Obtener parámetros del tipo
            param_type_number = element_type.LookupParameter("COBie.Type.ModelNumber")
            param_type_mreference = element_type.LookupParameter("COBie.Type.ModelReference")
            
            # Crear código
            code_furniture = param_sitename + "." + param_Pr + "." + new_param_activo + "." + str(element_type.Id.IntegerValue)
            
            # Verificar si los parámetros existen y no son de solo lectura antes de asignar valores
            if param_type_number and not param_type_number.IsReadOnly:
                param_type_number.Set(code_furniture)
                
            if param_type_mreference and not param_type_mreference.IsReadOnly:
                param_type_mreference.Set(code_furniture)
                
            # Imprimir solo si el parámetro existe
            if param_type_number:
                print("Tipo ID", element_type.Id.IntegerValue, ":", param_type_number.AsString())
                
        t.Commit()
    
    except Exception as e:
        print("Error:", e)
        t.RollBack()
