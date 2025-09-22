# -*- coding: utf-8 -*-
__title__ = "On/Off COBie"

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import BuiltInParameter, StorageType, FamilyInstance
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._Ignore import leer_excel_filtrado, ignorar_categorias
from Extensions._Dictionary import obtener_especialidad


def set_param(param, val):
    """Establece un valor entero en el parámetro si es modificable."""
    if param and not param.IsReadOnly and param.StorageType == StorageType.Integer:
        param.Set(val)


def compute_value(code, excel_set, especialidad, categoria, ignore_cats, activate):
    """
    Determina el valor (0 o 1) según el modo (activar/desactivar) y la lógica de especialidad.
    """
    if not activate:
        return 0
    if especialidad in ("Arquitectura", "Equipamiento y mobiliario"):
        return 0 if (not code or code in excel_set) else 1
    if especialidad == "Estructuras":
        return 1
    return 1 if categoria not in ignore_cats else 0

# Validar nombre de archivo y salir si no es válido
if not validar_nombre(obtener_nombre_archivo()):
    script.exit()

# Diálogo nativo: Yes = Activar, No = Desactivar
button_flags = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
result = TaskDialog.Show(
    "Modo de operación",          # título
    "¿Deseas ACTIVAR COBie?",     # instrucción
    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
)
if result == TaskDialogResult.Yes:
    modo_activar = True
elif result == TaskDialogResult.No:
    modo_activar = False
else:
    script.exit()

# Selección de elementos y configuración inicial
try:
    ui_doc     = __revit__.ActiveUIDocument
    refs       = ui_doc.Selection.PickObjects(ObjectType.Element)
    esp        = obtener_especialidad()
    ign_cats   = ignorar_categorias()
    excel_vals = leer_excel_filtrado() if esp in ("Arquitectura", "Equipamiento y mobiliario") else set()
except OperationCanceledException:
    script.exit()

# Preparar caché y tracker de tipos
tipo_codes = {}
processed_types = set()
doc = ui_doc.Document

# Ejecutar transacción general
with revit.Transaction("Toggle COBieType & COBie Component"):
    for ref in refs:
        elem = doc.GetElement(ref)
        if not elem or not elem.Category:
            continue

        elementos_a_procesar = [elem]

        # Solo verificar subcomponentes si es instancia de familia
        if isinstance(elem, FamilyInstance):
            try:
                sub_ids = elem.GetSubComponentIds()
                if sub_ids:
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass  # Por si algún caso lanza excepción

        # Procesar el elemento y sus subcomponentes si los hay
        for el in elementos_a_procesar:
            if not el.Category:
                continue

            cat = el.Category.Name
            t_id = el.GetTypeId()
            tipo = doc.GetElement(t_id)
            if not tipo:
                continue

            if t_id not in tipo_codes:
                p_u = tipo.get_Parameter(BuiltInParameter.UNIFORMAT_CODE)
                tipo_codes[t_id] = p_u.AsString() if p_u and p_u.HasValue else None
            code = tipo_codes[t_id]

            v = compute_value(code, excel_vals, esp, cat, ign_cats, modo_activar)

            if t_id not in processed_types:
                p_type = tipo.LookupParameter("COBie.Type")
                set_param(p_type, v)
                processed_types.add(t_id)

            p_inst = el.LookupParameter("COBie")
            set_param(p_inst, v)

# Mostrar resultado final
title = "Resultado"
message = "Procesados: {0} elementos, {1} tipos.\nModo: {2}".format(
    len(refs), len(processed_types),
    "ACTIVAR" if modo_activar else "DESACTIVAR"
)
forms.alert(message, title=title, exitscript=False)