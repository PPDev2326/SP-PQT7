# -*- coding: utf-8 -*-
__title__ = "Assign ID"

from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException # Importante para cuando das ESC
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from pyrevit import revit, script

doc, uidoc = revit.doc, revit.uidoc
output = script.get_output()

# --- 1. VALIDACIÓN INICIAL ---
nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

param_name = "S&P_ID DE ELEMENTO"

# --- 2. DEFINIR ESTRATEGIA DE VERSIÓN ---
try:
    # Probamos con un elemento dummy o simplemente verificamos el atributo en la clase
    # En este caso, usaremos una función lambda para definir el comportamiento
    dummy_id = doc.GetElement(doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element)).Id).Id.Value
    # Si la linea de arriba no falla, es Revit 2024+
    get_id_val = lambda elem_id: elem_id.Value
except:
    # Si falla, es Revit 2023 o anterior
    get_id_val = lambda elem_id: elem_id.IntegerValue

# --- 3. SELECCIÓN ---
try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecciona elementos para asignar ID")
except OperationCanceledException:
    # Si el usuario presiona ESC, salimos silenciosamente
    script.exit()

# Listas para seguimiento
asignados = []
faltantes = []
errores_tipo = []

# --- 4. PROCESAMIENTO ---
with revit.Transaction("Asignar Element ID"):
    for r in refs:
        e = doc.GetElement(r.ElementId)
        p = e.LookupParameter(param_name)
        
        if p and not p.IsReadOnly:
            try:
                # Usamos la función definida al inicio (más rápido)
                raw_val = get_id_val(e.Id)
                
                # Forzamos a int() por seguridad (Revit 2024 usa Int64)
                val_to_set = int(raw_val)
                
                if p.Set(val_to_set):
                    asignados.append(e.Id)
                else:
                    # Si falla, puede ser porque el parámetro NO es Integer (ej. es Texto)
                    faltantes.append(e.Id)
            except Exception as ex:
                faltantes.append(e.Id)
        else:
            faltantes.append(e.Id)

# --- 5. REPORTE ---
if asignados:
    output.print_md("### ✅ IDs asignados ({})".format(len(asignados)))
    # Opcional: No imprimir miles de IDs si son muchos
    if len(asignados) < 50:
        output.print_md(", ".join(str(i) for i in asignados))
    else:
        output.print_md("_Se han asignado {} elementos (lista oculta por longitud)_".format(len(asignados)))

if faltantes:
    output.print_md("\n### ⚠️ No se pudo asignar ({}):".format(len(faltantes)))
    output.print_md("> Posibles causas: El parámetro no existe, es de solo lectura, o no es de tipo Entero.")
    for eid in faltantes:
        output.print_md("- {}".format(output.linkify(eid)))