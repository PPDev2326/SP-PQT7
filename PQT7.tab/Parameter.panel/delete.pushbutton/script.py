# -*- coding: utf-8 -*-
__title__ = "Delete\nParameters"

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SharedParameterElement,
    InstanceBinding
)
from pyrevit import script, forms, revit
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre

doc = revit.doc
output = script.get_output()

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

# Obtener todos los parámetros compartidos
shared_params = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()
if not shared_params:
    forms.alert("No hay parámetros compartidos en este proyecto.", title="Aviso")
    script.exit()

# Crear diccionario para mostrar nombre + GUID + tipo
param_labels = []
param_map = {}

# Mapa de bindings (para saber si son de tipo o instancia)
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
iterator.Reset()
bindings_dict = {}

while iterator.MoveNext():
    definition = iterator.Key
    binding = iterator.Current
    tipo = "Instancia" if isinstance(binding, InstanceBinding) else "Tipo"
    bindings_dict[definition.Name] = tipo

for param in shared_params:
    name = param.Name
    guid = param.GuidValue
    guid_str = str(guid) if guid else "No disponible"
    tipo = bindings_dict.get(name, "Desconocido")

    label = "{} | GUID: {} | {}".format(name, guid_str, tipo)
    param_labels.append(label)
    param_map[label] = param

# Mostrar selección al usuario
selected_labels = forms.SelectFromList.show(
    param_labels,
    title="Seleccionar parámetros compartidos a eliminar",
    multiselect=True,
    button_name="Eliminar"
)

if selected_labels:
    confirm = forms.alert(
        "¿Estás seguro que deseas eliminar los siguientes parámetros?\n\n{}".format(
            "\n".join(selected_labels)
        ),
        title="Confirmar eliminación",
        ok=True,
        cancel=True
    )

    if not confirm:
        forms.alert("Eliminación cancelada.", title="Cancelado")
        script.exit()

    to_delete_ids = []

    with revit.Transaction("Eliminar bindings"):
        for label in selected_labels:
            param = param_map[label]
            definition = param.GetDefinition()

            try:
                if definition and binding_map.Remove(definition):
                    output.print_md("🔗 Binding eliminado de **{}**.".format(param.Name))
            except Exception as e:
                output.print_md("⚠️ Error al quitar binding de {}: {}".format(param.Name, str(e)))

            to_delete_ids.append(param.Id)

    with revit.Transaction("Eliminar parámetros"):
        for eid in to_delete_ids:
            try:
                doc.Delete(eid)
                output.print_md("> ✅ Se eliminó el parámetro con ID **{}**.".format(eid.IntegerValue))
            except Exception as e:
                output.print_md("❌ No se pudo eliminar el parámetro con ID {}: {}".format(eid.IntegerValue, str(e)))

    forms.alert("Eliminación completada.", title="Finalizado")
else:
    forms.alert("No se seleccionaron parámetros.", title="Cancelado")