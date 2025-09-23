# -*- coding: utf-8 -*-
__title__ = "On/Off COBie"

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import StorageType, FamilyInstance
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._Ignore import leer_excel_filtrado
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository

def set_param(param, val):
    """Establece un valor entero en el parámetro si es modificable."""
    if param and not param.IsReadOnly and param.StorageType == StorageType.Integer:
        param.Set(val)

def get_codigo_partida(elemento, documento):
    """
    Obtiene el código de partida del elemento según especialidad.
    Para INSTALACIONES ELECTRICAS/COMUNICACIONES: prioriza N°2, fallback a N°1.
    Otras especialidades: usa N°1. Retorna None si está vacío.
    """
    specialty_obj = SpecialtiesRepository().get_specialty_by_document(documento)
    specialty = specialty_obj.name if specialty_obj else None
    
    # DEBUG: Imprimir especialidad
    print("Especialidad detectada: {}".format(specialty))
    
    if specialty in ["INSTALACIONES ELECTRICAS", "COMUNICACIONES"]:
        # Intentar primero con CODIGO PARTIDA N°2
        param_n2 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°2")
        if param_n2 and param_n2.HasValue:
            valor_n2 = param_n2.AsString()
            if valor_n2 and valor_n2.strip():
                print("Código N°2 encontrado: {}".format(valor_n2.strip()))
                return valor_n2.strip()
        
        # Fallback a N°1 si N°2 no está disponible o está vacío
        param_n1 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°1")
        if param_n1 and param_n1.HasValue:
            valor_n1 = param_n1.AsString()
            if valor_n1 and valor_n1.strip():
                print("Código N°1 encontrado (fallback): {}".format(valor_n1.strip()))
                return valor_n1.strip()
    else:
        # Para otras especialidades, usar directamente N°1
        param_n1 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°1")
        if param_n1 and param_n1.HasValue:
            valor_n1 = param_n1.AsString()
            if valor_n1 and valor_n1.strip():
                print("Código N°1 encontrado: {}".format(valor_n1.strip()))
                return valor_n1.strip()
    
    print("No se encontró código de partida válido")
    return None

def compute_value(code, sin_cobie_set, con_cobie_set, activate):
    """
    Determina el valor (0 o 1) según el modo y si el código está en las listas del Excel.
    """
    if not activate:
        return 0
    # Si el código está en la lista COBie (Y), devolver 1
    if code and code in con_cobie_set:
        print("Código {} está en lista COBie (Y) - valor: 1".format(code))
        return 1
    # Si el código está en la lista sin COBie (N), devolver 0
    if code and code in sin_cobie_set:
        print("Código {} está en lista sin COBie (N) - valor: 0".format(code))
        return 0
    # Si no está en ninguna lista, comportamiento por defecto
    print("Código {} no está en ninguna lista - valor por defecto: 0".format(code))
    return 0

if not validar_nombre(obtener_nombre_archivo()):
    script.exit()

# TaskDialog más simple y intuitivo
result = TaskDialog.Show(
    "COBie Manager",
    "Selecciona la acción:",
    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
)

# Configurar texto de los botones más claro
if result == TaskDialogResult.Yes:
    modo_activar = True
    print("=== MODO: ACTIVAR COBie ===")
elif result == TaskDialogResult.No:
    modo_activar = False
    print("=== MODO: DESACTIVAR COBie ===")
else:
    script.exit()

# Selección de elementos y configuración inicial
try:
    ui_doc = revit.uidoc
    doc = revit.doc
    print("Selecciona los elementos...")
    refs = ui_doc.Selection.PickObjects(ObjectType.Element)
    print("Elementos seleccionados: {}".format(len(refs)))
    
    print("Cargando datos del Excel...")
    sin_cobie, con_cobie = leer_excel_filtrado()
    print("Códigos sin COBie: {}, Códigos con COBie: {}".format(len(sin_cobie), len(con_cobie)))
    
except OperationCanceledException:
    print("Operación cancelada por el usuario")
    script.exit()

# Convertir listas a sets para búsquedas más rápidas
sin_cobie_set = set(sin_cobie)
con_cobie_set = set(con_cobie)

# Tracker de tipos procesados por código de partida
processed_codes_types = {}
elementos_procesados = 0

print("\n=== INICIANDO PROCESAMIENTO ===")

# Ejecutar transacción general
with revit.Transaction("Toggle COBieType & COBie Component"):
    for ref in refs:
        elem = doc.GetElement(ref)
        if not elem or not elem.Category:
            print("Elemento sin categoría - saltando")
            continue

        elementos_a_procesar = [elem]

        # Solo verificar subcomponentes si es instancia de familia
        if isinstance(elem, FamilyInstance):
            try:
                sub_ids = elem.GetSubComponentIds()
                if sub_ids:
                    print("Elemento con {} subcomponentes".format(len(sub_ids)))
                    for sid in sub_ids:
                        sub_elem = doc.GetElement(sid)
                        if sub_elem:
                            elementos_a_procesar.append(sub_elem)
            except:
                pass

        # Procesar el elemento y sus subcomponentes si los hay
        for el in elementos_a_procesar:
            if not el.Category:
                continue

            print("\n--- Procesando elemento ID: {} ---".format(el.Id))
            
            # Obtener código de partida del elemento de instancia
            code = get_codigo_partida(el, doc)
            
            # Calcular valor según lógica COBie
            v = compute_value(code, sin_cobie_set, con_cobie_set, modo_activar)

            # Procesar tipo (COBie.Type) - solo una vez por código de partida único
            t_id = el.GetTypeId()
            if code and code not in processed_codes_types:
                tipo = doc.GetElement(t_id)
                if tipo:
                    p_type = tipo.LookupParameter("COBie.Type")
                    if p_type:
                        set_param(p_type, v)
                        print("COBie.Type actualizado en tipo: {}".format(v))
                        processed_codes_types[code] = t_id
                    else:
                        print("Parámetro COBie.Type no encontrado en el tipo")

            # Procesar instancia (COBie)
            p_inst = el.LookupParameter("COBie")
            if p_inst:
                set_param(p_inst, v)
                print("COBie actualizado en instancia: {}".format(v))
            else:
                print("Parámetro COBie no encontrado en la instancia")
            
            elementos_procesados += 1

print("\n=== PROCESAMIENTO COMPLETADO ===")
print("Total elementos procesados: {}".format(elementos_procesados))
print("Total códigos únicos: {}".format(len(processed_codes_types)))
print("Modo aplicado: {}".format("ACTIVAR" if modo_activar else "DESACTIVAR"))

# Mensaje final en terminal (sin popup)
print("\n¡Proceso completado exitosamente!")