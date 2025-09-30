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


def extraer_numero_para_ordenar(numero_str):
    """
    Extrae el primer número encontrado en la cadena para ordenar.
    Ejemplo: "A-101" -> 101, "205B" -> 205
    """
    import re
    match = re.search(r'\d+', numero_str)
    if match:
        try:
            return int(match.group())
        except:
            pass
    return float('inf')  # Si no hay número, va al final


def obtener_dos_rooms_mas_cercanos(habitaciones, punto):
    """
    Obtiene las 2 habitaciones más cercanas a un punto.
    Retorna una lista de tuplas (distancia, room) ordenadas por distancia.
    """
    distancias = []
    
    for room in habitaciones:
        try:
            location = room.Location
            if not location:
                continue
            
            room_point = location.Point
            dx = room_point.X - punto.X
            dy = room_point.Y - punto.Y
            dz = room_point.Z - punto.Z
            distancia = (dx*dx + dy*dy + dz*dz) ** 0.5
            
            distancias.append((distancia, room))
        except:
            continue
    
    # Ordenar por distancia y tomar los 2 primeros
    distancias.sort(key=lambda x: x[0])
    return distancias[:2]


def procesar_puerta_ventana(elemento, habitaciones, failed_list):
    """
    Procesamiento especial para puertas y ventanas:
    Asigna los nombres de los 2 ambientes más cercanos ordenados por número de room.
    """
    pts = puntos_representativos(elemento) or []
    if not pts:
        return False
    
    # Usar el primer punto representativo
    punto = next((p for p in pts if p), None)
    if not punto:
        return False
    
    try:
        dos_cercanos = obtener_dos_rooms_mas_cercanos(habitaciones, punto)
        
        if not dos_cercanos:
            return False
        
        # Obtener nombres y números
        rooms_con_datos = []
        for distancia, room in dos_cercanos:
            nombre = get_room_name(room)
            numero = get_room_number(room)
            rooms_con_datos.append((numero, nombre, room))
        
        # Ordenar por número (de menor a mayor)
        rooms_con_datos.sort(key=lambda x: extraer_numero_para_ordenar(x[0]))
        
        # Construir el nombre combinado
        if len(rooms_con_datos) == 2:
            nombre_combinado = "{}, {}".format(
                rooms_con_datos[0][1].capitalize(),
                rooms_con_datos[1][1].capitalize()
            )
            # Para COBie, usar el número del primer ambiente
            numero_cobie = rooms_con_datos[0][0]
        elif len(rooms_con_datos) == 1:
            nombre_combinado = rooms_con_datos[0][1].capitalize()
            numero_cobie = rooms_con_datos[0][0]
        else:
            return False
        
        return asignar_ambiente(elemento, nombre_combinado, numero_cobie, failed_list)
        
    except Exception as e:
        output.print_md("**Error procesando puerta/ventana {}: {}**".format(elemento.Id, str(e)))
        return False


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
            prm_ambiente.Set(nombre_ambiente)
            exito = True
        
        # Asignar a COBie.Component.Space (numero : nombre)
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
            if numero_ambiente:
                valor_cobie = "{} : {}".format(numero_ambiente, nombre_ambiente)
            else:
                valor_cobie = nombre_ambiente
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
                return asignar_ambiente(elemento, nombre.capitalize(), numero, failed_list)
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
                return asignar_ambiente(elemento, nombre.capitalize(), numero, failed_list)
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

# Preguntar si solo quiere categorías de la vista activa
usar_solo_vista = forms.alert(
    "¿Deseas trabajar SOLO con las categorías de la vista activa '{}'?\n\n"
    "SI: Solo categorías visibles en esta vista\n"
    "NO: Todas las categorías del proyecto".format(vista_activa.Name),
    title="Filtrar por Vista Activa",
    yes=True,
    no=True
)

# Determinar qué categorías mostrar
if usar_solo_vista:
    opciones_categorias = categorias_vista_ordenadas
    titulo_seleccion = "Selecciona categorías de la vista '{}'".format(vista_activa.Name)
else:
    opciones_categorias = todas_categorias
    titulo_seleccion = "Selecciona categorías del proyecto"

if not opciones_categorias:
    forms.alert("No hay categorías disponibles.", exitscript=True)

# 2) Selección de categorías
seleccion = forms.SelectFromList.show(
    opciones_categorias, 
    multiselect=True,
    title=titulo_seleccion
)

if not seleccion:
    forms.alert("No se seleccionaron categorías.", exitscript=True)

sels_cats = [mapeo[n] for n in seleccion]
output.print_md("### Categorías seleccionadas: {}".format(", ".join(seleccion)))

# Verificar si hay puertas o ventanas seleccionadas
tiene_puertas = "Puertas" in seleccion or "Doors" in seleccion
tiene_ventanas = "Ventanas" in seleccion or "Windows" in seleccion
procesamiento_especial = tiene_puertas or tiene_ventanas

if procesamiento_especial:
    output.print_md("### ℹ️ Se aplicará procesamiento especial para puertas/ventanas (2 ambientes)")

# 3) Obtener habitaciones
habs = obtener_habitaciones(doc)
output.print_md("### Habitaciones disponibles: {}".format(len(habs)))

# 4) Obtener elementos a procesar
elems = obtener_elementos_de_categorias(doc, sels_cats)
output.print_md("### Elementos encontrados: {}".format(len(elems)))

# Tracking
elems_asignados = set()
elems_ignorados_cobie = []
elems_ignorados_llenos = []
failed_param = []
asignados_fase1 = 0
asignados_fase2 = 0
asignados_fase3 = 0
asignados_especial = 0

# ==================== PROCESAMIENTO EN UNA SOLA TRANSACCIÓN ====================

with revit.Transaction("Asignar Ambiente"):
    
    # Procesamiento especial para puertas y ventanas (si aplica)
    if procesamiento_especial:
        output.print_md("#### Procesando puertas y ventanas (2 ambientes)...")
        for e in elems:
            # Solo procesar si es puerta o ventana
            try:
                cat_name = e.Category.Name if e.Category else ""
                es_puerta_ventana = cat_name in ["Puertas", "Doors", "Ventanas", "Windows"]
                
                if not es_puerta_ventana:
                    continue
                
                # Verificar COBie activo
                if not verificar_cobie_activo(e):
                    elems_ignorados_cobie.append(e.Id)
                    continue
                
                # Verificar si parámetros están vacíos
                if not verificar_parametros_vacios(e):
                    elems_ignorados_llenos.append(e.Id)
                    continue
                
                if procesar_puerta_ventana(e, habs, failed_param):
                    elems_asignados.add(e.Id)
                    asignados_especial += 1
            except:
                continue
    
    # Fase 1: Elementos dentro de habitaciones (excluir ya procesados)
    output.print_md("#### Procesando Fase 1: Elementos dentro de habitaciones...")
    for e in elems:
        if e.Id in elems_asignados:
            continue
            
        # Verificar COBie activo
        if not verificar_cobie_activo(e):
            if e.Id not in elems_ignorados_cobie:
                elems_ignorados_cobie.append(e.Id)
            continue
        
        # Verificar si parámetros están vacíos
        if not verificar_parametros_vacios(e):
            if e.Id not in elems_ignorados_llenos:
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

total_asignados = asignados_fase1 + asignados_fase2 + asignados_fase3 + asignados_especial
total_ignorados = len(elems_ignorados_cobie) + len(elems_ignorados_llenos)

output.print_md("---")
output.print_md("### 📊 Resumen de Resultados")
if asignados_especial > 0:
    output.print_md("- **Puertas/Ventanas** (2 ambientes): {}".format(asignados_especial))
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