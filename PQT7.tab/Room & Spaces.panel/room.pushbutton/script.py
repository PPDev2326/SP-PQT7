# -*- coding: utf-8 -*-
__title__ = "Assign room"

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import (
    StorageType, RevitLinkInstance, FilteredElementCollector, BuiltInParameter
)
from Extensions._utils import (
    obtener_mapeo_nombres_categorias,
    obtener_elementos_de_categorias,
    obtener_habitaciones_y_espacios,
    obtener_habitaciones_de_vinculos_seleccionados,
    get_room_name,
    puntos_representativos,
    is_point_inside,
    obtener_room_mas_cercano
)

doc = revit.doc
output = script.get_output()

# Constantes
TOLERANCIA_PIES = 0.410105  # 12.5 cm en pies
PARAM_NAME = "S&P_AMBIENTE"
PARAM_COBIE = "COBie.Component.Space"
PARAM_COBIE_BOOL = "COBie"
FALLBACK_VALUE = "Activo"


def get_room_number(room):
    """Obtiene el número de la habitación."""
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if param and param.HasValue:
            numero = param.AsString()
            return numero if numero else ""
    except:
        pass
    return ""


def verificar_parametros_vacios(elemento):
    """
    Verifica si los parámetros S&P_AMBIENTE y COBie.Component.Space están vacíos.
    Retorna True si ambos están vacíos, False si al menos uno tiene valor.
    """
    prm_ambiente = elemento.LookupParameter(PARAM_NAME)
    prm_cobie = elemento.LookupParameter(PARAM_COBIE)
    
    ambiente_vacio = True
    cobie_vacio = True
    
    if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
        valor = prm_ambiente.AsString()
        if valor and valor.strip():
            ambiente_vacio = False
    
    if prm_cobie and prm_cobie.StorageType == StorageType.String:
        valor = prm_cobie.AsString()
        if valor and valor.strip():
            cobie_vacio = False
    
    # Solo procesamos si AMBOS están vacíos
    return ambiente_vacio and cobie_vacio


def verificar_cobie_activo(elemento):
    """
    Verifica si el parámetro COBie (bool) está activado (True/Si/1).
    Retorna True si está activado, False en caso contrario.
    """
    try:
        prm_cobie_bool = elemento.LookupParameter(PARAM_COBIE_BOOL)
        if prm_cobie_bool and prm_cobie_bool.StorageType == StorageType.Integer:
            return prm_cobie_bool.AsInteger() == 1
    except:
        pass
    return False


def asignar_ambiente(elemento, nombre_ambiente, numero_ambiente, failed_list):
    """
    Asigna valores a los parámetros 'S&P_AMBIENTE' y 'COBie.Component.Space'.
    - S&P_AMBIENTE: solo el nombre
    - COBie.Component.Space: formato "numero : nombre"
    Retorna True si se asignó correctamente al menos uno, False si ambos fallan.
    """
    exito = False
    
    try:
        # Asignar a S&P_AMBIENTE (solo nombre)
        prm_ambiente = elemento.LookupParameter(PARAM_NAME)
        if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
            prm_ambiente.Set(nombre_ambiente.capitalize())
            exito = True
        
        # Asignar a COBie.Component.Space (numero : nombre)
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
            if numero_ambiente:
                valor_cobie = "{} : {}".format(numero_ambiente, nombre_ambiente.capitalize())
            else:
                valor_cobie = nombre_ambiente.capitalize()
            prm_cobie.Set(valor_cobie)
            exito = True
        
        if not exito:
            failed_list.append(elemento.Id)
            
    except Exception as e:
        output.print_md("**Error en elemento {}: {}**".format(elemento.Id, str(e)))
        failed_list.append(elemento.Id)
        return False
    
    return exito


def procesar_elemento_fase1(elemento, habitaciones, failed_list):
    """Fase 1: Busca si algún punto del elemento está dentro de una habitación."""
    pts = puntos_representativos(elemento) or []
    for p in pts:
        if not p:
            continue
        try:
            room_hit = next((r for r in habitaciones if is_point_inside(r, p)), None)
            if room_hit:
                nombre = get_room_name(room_hit)
                numero = get_room_number(room_hit)
                return asignar_ambiente(elemento, nombre, numero, failed_list)
        except Exception:
            continue
    return False


def procesar_elemento_fase2(elemento, habitaciones, failed_list):
    """Fase 2: Busca la habitación más cercana dentro de la tolerancia."""
    pts = puntos_representativos(elemento) or []
    for p in pts:
        if not p:
            continue
        try:
            room_cercana = obtener_room_mas_cercano(habitaciones, p, tolerancia=TOLERANCIA_PIES)
            if room_cercana:
                nombre = get_room_name(room_cercana)
                numero = get_room_number(room_cercana)
                return asignar_ambiente(elemento, nombre, numero, failed_list)
        except Exception:
            continue
    return False


def obtener_habitaciones(documento):
    """Obtiene habitaciones del documento actual o de vínculos seleccionados."""
    habs = obtener_habitaciones_y_espacios(documento)
    if habs:
        return habs
    
    # Si no hay habitaciones, buscar en vínculos
    links = FilteredElementCollector(documento).OfClass(RevitLinkInstance).ToElements()
    nom_links = sorted({l.Name for l in links})
    
    if not nom_links:
        forms.alert("No hay habitaciones ni vínculos disponibles.", exitscript=True)
    
    sel_links = forms.SelectFromList.show(
        nom_links, 
        multiselect=True,
        title="Selecciona vínculos con habitaciones"
    )
    
    if not sel_links:
        forms.alert("No se seleccionaron vínculos.", exitscript=True)
    
    habs = obtener_habitaciones_de_vinculos_seleccionados(documento, sel_links)
    
    if not habs:
        forms.alert("No se encontraron habitaciones en los vínculos.", exitscript=True)
    
    return habs


# ==================== INICIO DEL SCRIPT ====================

# 1) Obtener categorías y preparar opciones
mapeo = obtener_mapeo_nombres_categorias(doc)
todas_categorias = sorted(mapeo.keys())

# Obtener categorías de la vista activa
vista_activa = doc.ActiveView
categorias_vista = set()
collector = FilteredElementCollector(doc, vista_activa.Id).WhereElementIsNotElementType()
for elem in collector:
    try:
        if elem.Category and elem.Category.Name in mapeo:
            categorias_vista.add(elem.Category.Name)
    except:
        continue

categorias_vista_ordenadas = sorted(categorias_vista)

# Crear opciones combinadas
opciones = []
opciones.append("--- SOLO VISTA ACTIVA ({}) ---".format(vista_activa.Name))
opciones.extend(categorias_vista_ordenadas)
opciones.append("--- TODAS LAS CATEGORÍAS DEL PROYECTO ---")
opciones.extend(todas_categorias)

# 2) Selección de categorías
seleccion = forms.SelectFromList.show(
    opciones, 
    multiselect=True,
    title="Selecciona categorías a procesar"
)

if not seleccion:
    forms.alert("No se seleccionaron categorías.", exitscript=True)

# Filtrar separadores y determinar categorías seleccionadas
seleccion_final = [s for s in seleccion if not s.startswith("---")]

if not seleccion_final:
    forms.alert("No se seleccionaron categorías válidas.", exitscript=True)

sels_cats = [mapeo[n] for n in seleccion_final]
output.print_md("### Categorías seleccionadas: {}".format(", ".join(seleccion_final)))

# 2) Obtener habitaciones
habs = obtener_habitaciones(doc)
output.print_md("### Habitaciones disponibles: {}".format(len(habs)))

# 3) Obtener elementos a procesar
elems = obtener_elementos_de_categorias(doc, sels_cats)
output.print_md("### Elementos encontrados: {}".format(len(elems)))

# Tracking
elems_asignados = set()
elems_ignorados_cobie = []  # Elementos sin COBie activo
elems_ignorados_llenos = []  # Elementos con parámetros ya llenos
failed_param = []
asignados_fase1 = 0
asignados_fase2 = 0
asignados_fase3 = 0

# ==================== PROCESAMIENTO EN UNA SOLA TRANSACCIÓN ====================

with revit.Transaction("Asignar Ambiente"):
    
    # Fase 1: Elementos dentro de habitaciones
    output.print_md("#### Procesando Fase 1: Elementos dentro de habitaciones...")
    for e in elems:
        # Verificar COBie activo
        if not verificar_cobie_activo(e):
            elems_ignorados_cobie.append(e.Id)
            continue
        
        # Verificar si parámetros están vacíos
        if not verificar_parametros_vacios(e):
            elems_ignorados_llenos.append(e.Id)
            continue
        
        if procesar_elemento_fase1(e, habs, failed_param):
            elems_asignados.add(e.Id)
            asignados_fase1 += 1
    
    # Fase 2: Proximidad
    output.print_md("#### Procesando Fase 2: Elementos por proximidad...")
    for e in elems:
        if e.Id in elems_asignados or e.Id in elems_ignorados_cobie or e.Id in elems_ignorados_llenos:
            continue
        
        if procesar_elemento_fase2(e, habs, failed_param):
            elems_asignados.add(e.Id)
            asignados_fase2 += 1
    
    # Fase 3: Resto como "Activo"
    output.print_md("#### Procesando Fase 3: Elementos restantes como '{}'...".format(FALLBACK_VALUE))
    for e in elems:
        if e.Id in elems_asignados or e.Id in elems_ignorados_cobie or e.Id in elems_ignorados_llenos:
            continue
        
        if asignar_ambiente(e, FALLBACK_VALUE, "", failed_param):
            elems_asignados.add(e.Id)
            asignados_fase3 += 1

# ==================== RESULTADOS ====================

total_asignados = asignados_fase1 + asignados_fase2 + asignados_fase3
total_ignorados = len(elems_ignorados_cobie) + len(elems_ignorados_llenos)

output.print_md("---")
output.print_md("### 📊 Resumen de Resultados")
output.print_md("- **Fase 1** (dentro de habitación): {}".format(asignados_fase1))
output.print_md("- **Fase 2** (por proximidad): {}".format(asignados_fase2))
output.print_md("- **Fase 3** (como '{}'): {}".format(FALLBACK_VALUE, asignados_fase3))
output.print_md("- **Total asignados**: {}".format(total_asignados))
output.print_md("- **Ignorados** (COBie inactivo): {}".format(len(elems_ignorados_cobie)))
output.print_md("- **Ignorados** (parámetros ya llenos): {}".format(len(elems_ignorados_llenos)))
output.print_md("- **Sin asignar**: {}".format(len(failed_param)))

if elems_ignorados_cobie:
    sample = elems_ignorados_cobie[:10]
    lines = ["- Id {}".format(i) for i in sample]
    extra = " (mostrando 10 de {})".format(len(elems_ignorados_cobie)) if len(elems_ignorados_cobie) > 10 else ""
    output.print_md("#### ℹ️ Ignorados por COBie inactivo{}:\n{}".format(extra, "\n".join(lines)))

if elems_ignorados_llenos:
    sample = elems_ignorados_llenos[:10]
    lines = ["- Id {}".format(i) for i in sample]
    extra = " (mostrando 10 de {})".format(len(elems_ignorados_llenos)) if len(elems_ignorados_llenos) > 10 else ""
    output.print_md("#### ℹ️ Ignorados por parámetros ya llenos{}:\n{}".format(extra, "\n".join(lines)))

if failed_param:
    sample = failed_param[:15]
    lines = ["- Id {}".format(i) for i in sample]
    extra = " (mostrando 15 de {})".format(len(failed_param)) if len(failed_param) > 15 else ""
    output.print_md("#### ⚠️ Sin asignar (parámetros '{}' o '{}' faltantes o inválidos){}:\n{}".format(
        PARAM_NAME, PARAM_COBIE, extra, "\n".join(lines)
    ))

forms.alert(
    "Proceso terminado:\n\n"
    "✅ {} elementos asignados\n"
    "⏭️ {} elementos ignorados\n"
    "❌ {} sin asignar (sin parámetros válidos)".format(total_asignados, total_ignorados, len(failed_param)),
    title="Asignación Completada"
)