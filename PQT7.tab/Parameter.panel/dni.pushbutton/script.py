# -*- coding: utf-8 -*-
__title__ = "Assign ID"

from Autodesk.Revit.UI.Selection import ObjectType
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from pyrevit import revit, script

doc, uidoc = revit.doc, revit.uidoc

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

output = script.get_output()
param_name = "S&P_ID DE ELEMENTO"

# Seleccionar elementos manualmente
refs = uidoc.Selection.PickObjects(ObjectType.Element)

# Listas para seguimiento
asignados = []
faltantes = []

# Iniciar transacción para asignar el ID
with revit.Transaction("Asignar Element ID"):
    for r in refs:
        e = doc.GetElement(r.ElementId)
        p = e.LookupParameter(param_name)
        if p and not p.IsReadOnly:
            if p.Set(str(e.Id)):
                asignados.append(e.Id)
            else:
                faltantes.append(e.Id)
        else:
            faltantes.append(e.Id)

# Mostrar resultados en el panel de pyRevit
output.print_md("### ✅ IDs asignados:")
output.print_md(", ".join(str(i) for i in asignados) or "_Ninguno_")

# Mostrar elementos que no pudieron asignarse, con enlaces clicables
if faltantes:
    output.print_md("\n### ⚠️ Elementos sin parámetro `{}`:".format(param_name))
    for eid in faltantes:
        output.print_md("- {}".format(output.linkify(eid)))
