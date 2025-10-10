# -*- coding: utf-8 -*-
__title__ = "On/Off COBie"
__doc__ = """
Version = 1.1
Date = 22.09.2025
------------------------------------------------------------------
Description:
Mediante una matriz excel nos activara el parametro de COBie y COBieType.
------------------------------------------------------------------
¿Cómo hacerlo?
-> Click en el boton
-> Elegir entre 'SI' y 'NO' segun cada usuario
-> Seleccionamos los elementos de la vista
-> Click en finalizar en la parte superior
-> Elegimos matriz Excel COBie
------------------------------------------------------------------
Última actualización:
- [22.09.2025] - 1.1 UPDATE - New Feature
------------------------------------------------------------------
Autor: Paolo Perez
"""

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
    
    if specialty in ["INSTALACIONES ELECTRICAS", "COMUNICACIONES"]:
        # Intentar primero con CODIGO PARTIDA N°2
        param_n2 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°2")
        if param_n2 and param_n2.HasValue:
            valor_n2 = param_n2.AsString()
            if valor_n2 and valor_n2.strip():
                if valor_n2.strip() != "n/a":
                    return valor_n2.strip()
        
        # Fallback a N°1 si N°2 no está disponible o está vacío
        param_n1 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°1")
        if param_n1 and param_n1.HasValue:
            valor_n1 = param_n1.AsString()
            if valor_n1 and valor_n1.strip():
                return valor_n1.strip()
    else:
        # Para otras especialidades, usar directamente N°1
        param_n1 = elemento.LookupParameter("S&P_CODIGO PARTIDA N°1")
        if param_n1 and param_n1.HasValue:
            valor_n1 = param_n1.AsString()
            if valor_n1 and valor_n1.strip():
                return valor_n1.strip()
    
    return None

def compute_value(code, sin_cobie_set, con_cobie_set, activate):
    """Determina el valor (0 o 1) según el modo y si el código está en las listas del Excel."""
    if not activate:
        return 0
    if code and code in con_cobie_set:
        return 1
    if code and code in sin_cobie_set:
        return 0
    return 0

nombre_archivo = obtener_nombre_archivo
if not validar_nombre(nombre_archivo):
    forms.alert("ERROR: Se tiene error en el archivo {}".format(nombre_archivo), "ERROR")
    script.exit()

# TaskDialog más intuitivo con mensaje claro
result = TaskDialog.Show(
    "COBie Manager",
    "¿Qué acción deseas realizar?\n\nSI = ACTIVAR COBie (valor 1)\nNO = DESACTIVAR COBie (valor 0)",
    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
)

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
    print("Seleccionando elementos...")
    refs = ui_doc.Selection.PickObjects(ObjectType.Element)
    print("Elementos seleccionados: {}".format(len(refs)))
    
    # Solo leer Excel si se va a ACTIVAR COBie
    if modo_activar:
        print("Cargando datos del archivo Excel...")
        sin_cobie, con_cobie = leer_excel_filtrado()
        print("Codigos sin COBie: {}".format(len(sin_cobie)))
        print("Codigos con COBie: {}".format(len(con_cobie)))
    else:
        print("Modo DESACTIVAR: omitiendo lectura de Excel")
        sin_cobie, con_cobie = [], []
        
except OperationCanceledException:
    print("Operacion cancelada por el usuario")
    script.exit()

# Convertir listas a sets para búsquedas más rápidas
sin_cobie_set = set(sin_cobie)
con_cobie_set = set(con_cobie)

# Obtener especialidad del documento
specialty_obj = SpecialtiesRepository().get_specialty_by_document(doc)
specialty_name = specialty_obj.name if specialty_obj else "No detectada"
print("Especialidad del documento: {}".format(specialty_name))

# Tracker de tipos procesados por código de partida
processed_codes_types = {}
elementos_procesados = 0

# Contadores para estadísticas detalladas
instancias_valor_1 = 0
instancias_valor_0 = 0
tipos_valor_1 = 0
tipos_valor_0 = 0

# Calcular total de elementos a procesar para la barra de progreso
total_elementos = 0
elementos_con_subcomponentes = 0

for ref in refs:
    elem = doc.GetElement(ref)
    if elem and elem.Category:
        total_elementos += 1
        if isinstance(elem, FamilyInstance):
            try:
                sub_ids = elem.GetSubComponentIds()
                if sub_ids:
                    subcomponentes_validos = [sid for sid in sub_ids if doc.GetElement(sid)]
                    if subcomponentes_validos:
                        elementos_con_subcomponentes += 1
                        total_elementos += len(subcomponentes_validos)
            except:
                pass

print("Total elementos a procesar: {}".format(total_elementos))
print("Elementos principales con subcomponentes: {}".format(elementos_con_subcomponentes))
print("\n=== INICIANDO PROCESAMIENTO ===")

# Ejecutar transacción general con barra de progreso
with revit.Transaction("Toggle COBieType & COBie Component"):
    with forms.ProgressBar(title='Procesando elementos COBie...', 
                          step=1, 
                          cancellable=True) as pb:
        
        for ref in refs:
            if pb.cancelled:
                script.exit()
                
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
                    pass

            # Procesar el elemento y sus subcomponentes si los hay
            for el in elementos_a_procesar:
                if not el.Category:
                    continue

                # Actualizar barra de progreso
                pb.update_progress(elementos_procesados + 1, total_elementos)
                
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
                        set_param(p_type, v)
                        processed_codes_types[code] = t_id
                        # Contar tipos según valor
                        if v == 1:
                            tipos_valor_1 += 1
                        else:
                            tipos_valor_0 += 1

                # Procesar instancia (COBie)
                p_inst = el.LookupParameter("COBie")
                set_param(p_inst, v)
                
                # Contar instancias según valor
                if v == 1:
                    instancias_valor_1 += 1
                else:
                    instancias_valor_0 += 1
                
                elementos_procesados += 1

# TaskDialog de finalización exitosa
TaskDialog.Show(
    "Proceso Completado",
    "¡Proceso completado exitosamente!\n\n" +
    "Elementos procesados: {}\n".format(elementos_procesados) +
    "Códigos únicos: {}\n".format(len(processed_codes_types)) +
    "Modo aplicado: {}".format("ACTIVAR" if modo_activar else "DESACTIVAR"),
    TaskDialogCommonButtons.Ok
)

print("\n=== RESUMEN FINAL ===")
print("Elementos procesados: {}".format(elementos_procesados))
print("Codigos unicos de partida: {}".format(len(processed_codes_types)))
print("Modo aplicado: {}".format("ACTIVAR" if modo_activar else "DESACTIVAR"))
print("Especialidad: {}".format(specialty_name))

print("\n--- ESTADISTICAS DETALLADAS ---")
print("INSTANCIAS:")
print("  - Con valor 1 (activadas): {}".format(instancias_valor_1))
print("  - Con valor 0 (desactivadas): {}".format(instancias_valor_0))
print("TIPOS:")
print("  - Con valor 1 (activados): {}".format(tipos_valor_1))
print("  - Con valor 0 (desactivados): {}".format(tipos_valor_0))

if modo_activar:
    print("\n--- DATOS DEL EXCEL ---")
    print("Codigos definidos con COBie (Y): {}".format(len(con_cobie)))
    print("Codigos definidos sin COBie (N): {}".format(len(sin_cobie)))
    
print("¡Proceso completado exitosamente!")