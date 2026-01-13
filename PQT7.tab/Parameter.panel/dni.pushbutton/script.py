# -*- coding: utf-8 -*-
__title__ = "Assign ID"

# --- IMPORTS ---
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
# Importamos WorksharingUtils y ElementId para el checkout
from Autodesk.Revit.DB import WorksharingUtils, ElementId 
# Importamos List de .NET para poder pasar la lista de IDs a Revit
from System.Collections.Generic import List 

from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from pyrevit import revit, script

doc, uidoc = revit.doc, revit.uidoc
output = script.get_output()

# --- 1. VALIDACIÓN INICIAL ---
nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

param_name = "S&P_ID DE ELEMENTO"

# --- 2. SELECCIÓN DE ELEMENTOS ---
# Movemos la selección arriba para usarla tanto en la detección de versión como en el checkout
try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecciona elementos para asignar ID")
except OperationCanceledException:
    script.exit()

if not refs:
    script.exit()

# --- 3. DEFINIR ESTRATEGIA DE VERSIÓN (OPTIMIZADO) ---
# Usamos el primer elemento seleccionado para probar la versión de Revit
try:
    # Intentamos obtener el .Value (Revit 2024+) del primer elemento
    first_elem = doc.GetElement(refs[0].ElementId)
    test_val = first_elem.Id.Value
    
    # Si no falló la línea anterior, definimos la función para 2024
    get_id_val = lambda elem_id: elem_id.Value
except:
    # Si falló, es Revit 2023 o anterior
    get_id_val = lambda elem_id: elem_id.IntegerValue

# --- 4. PREPARACIÓN Y CHECKOUT (SOLUCIÓN AL POPUP) ---
# Convertimos la selección a una lista .NET de ElementIds
ids_para_procesar = List[ElementId]()
for r in refs:
    ids_para_procesar.Add(r.ElementId)

# Si el modelo es colaborativo, hacemos Checkout explícito
if doc.IsWorkshared:
    try:
        # Esto evita la ventana de "Desea aplicar checkout..."
        WorksharingUtils.CheckoutElements(doc, ids_para_procesar)
    except Exception as e:
        # Si falla el checkout (ej. alguien más tiene el objeto), lo registramos pero seguimos
        output.print_md("⚠️ **Advertencia de Worksharing:** Algunos elementos no se pudieron reservar. " + str(e))

# --- 5. PROCESAMIENTO ---
asignados = []
faltantes = []

with revit.Transaction("Asignar Element ID"):
    # Iteramos sobre la lista de IDs que ya preparamos
    for eid in ids_para_procesar:
        e = doc.GetElement(eid)
        p = e.LookupParameter(param_name)
        
        if p and not p.IsReadOnly:
            try:
                # 1. Obtener ID puro
                raw_val = get_id_val(e.Id)
                
                # 2. Forzar a int (seguridad de tipos)
                val_to_set = int(raw_val)
                
                # 3. Asignar
                if p.Set(val_to_set):
                    asignados.append(e.Id)
                else:
                    faltantes.append(e.Id)
            except Exception as ex:
                faltantes.append(e.Id)
        else:
            faltantes.append(e.Id)

# --- 6. REPORTE ---
if asignados:
    output.print_md("### ✅ IDs asignados ({})".format(len(asignados)))
    if len(asignados) < 50:
        output.print_md(", ".join(str(i) for i in asignados))
    else:
        output.print_md("_Se han asignado {} elementos (lista oculta por longitud)_".format(len(asignados)))

if faltantes:
    output.print_md("\n### ⚠️ No se pudo asignar ({}):".format(len(faltantes)))
    output.print_md("> Posibles causas: El elemento pertenece a otro usuario, el parámetro no existe o no es entero.")
    for eid in faltantes:
        output.print_md("- {}".format(output.linkify(eid)))