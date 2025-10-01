# -*- coding: utf-8 -*-
__title__ = "Assign room"

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import (
    StorageType, RevitLinkInstance, FilteredElementCollector, BuiltInParameter, ElementId
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


def procesar_puerta_ventana(elemento, habitaciones, failed_list):
    try:
        rooms_unicos = []
        ids_usados = set()
        numeros_usados = set()
        
        # Obtener el nivel del elemento para filtrar habitaciones
        elemento_nivel_id = None
        try:
            elemento_nivel_id = elemento.LevelId if hasattr(elemento, 'LevelId') else None
        except:
            pass
        
        # Filtrar habitaciones del mismo nivel si es posible
        habitaciones_filtradas = habitaciones
        if elemento_nivel_id and elemento_nivel_id.IntegerValue != -1:
            habs_mismo_nivel = [h for h in habitaciones if h.LevelId == elemento_nivel_id]
            if habs_mismo_nivel:
                habitaciones_filtradas = habs_mismo_nivel
        
        # PASO 1: Intentar FromRoom y ToRoom
        from_room = None
        to_room = None
        try:
            from_room = elemento.FromRoom[doc.Phase] if hasattr(elemento, "FromRoom") else None
            to_room = elemento.ToRoom[doc.Phase] if hasattr(elemento, "ToRoom") else None
        except:
            pass

        # Agregar FromRoom si es válido y del mismo nivel
        if from_room:
            # Verificar que sea del mismo nivel
            if not elemento_nivel_id or from_room.LevelId == elemento_nivel_id:
                nombre = get_room_name(from_room)
                numero = get_room_number(from_room)
                if nombre and numero:
                    nombre_upper = nombre.upper()
                    rooms_unicos.append((numero, nombre_upper, from_room.Id))
                    ids_usados.add(from_room.Id.IntegerValue)
                    numeros_usados.add(numero)

        # Agregar ToRoom si es válido, diferente y del mismo nivel
        if to_room and to_room.Id.IntegerValue not in ids_usados:
            # Verificar que sea del mismo nivel
            if not elemento_nivel_id or to_room.LevelId == elemento_nivel_id:
                nombre = get_room_name(to_room)
                numero = get_room_number(to_room)
                if nombre and numero:
                    nombre_upper = nombre.upper()
                    # Solo agregar si el NÚMERO es diferente
                    if numero not in numeros_usados:
                        rooms_unicos.append((numero, nombre_upper, to_room.Id))
                        ids_usados.add(to_room.Id.IntegerValue)
                        numeros_usados.add(numero)

        # PASO 2: Si tenemos menos de 2 ambientes, buscar por proximidad
        if len(rooms_unicos) < 2:
            pts = puntos_representativos(elemento) or []
            
            # Obtener TODOS los rooms cercanos de todos los puntos (del mismo nivel)
            candidatos = []
            for punto in pts:
                if not punto:
                    continue
                    
                cercanos = obtener_dos_rooms_mas_cercanos(habitaciones_filtradas, punto)
                for distancia, room in cercanos:
                    # Solo agregar si no está en candidatos ya
                    if room.Id.IntegerValue not in [r.Id.IntegerValue for d, r in candidatos]:
                        candidatos.append((distancia, room))
            
            # Ordenar candidatos por distancia
            candidatos.sort(key=lambda x: x[0])
            
            # Intentar agregar hasta tener 2 rooms diferentes
            for distancia, room in candidatos:
                if len(rooms_unicos) >= 2:
                    break
                
                # Saltar si ya usamos este room
                if room.Id.IntegerValue in ids_usados:
                    continue
                
                nombre = get_room_name(room)
                numero = get_room_number(room)
                
                if not nombre or not numero:
                    continue
                
                nombre_upper = nombre.upper()
                
                # Solo agregar si el NÚMERO es diferente
                if numero not in numeros_usados:
                    rooms_unicos.append((numero, nombre_upper, room.Id))
                    ids_usados.add(room.Id.IntegerValue)
                    numeros_usados.add(numero)

        # Verificar que tengamos al menos un ambiente
        if not rooms_unicos:
            return False

        # PASO 3: Ordenar por número y construir cadena
        rooms_unicos.sort(key=lambda x: extraer_numero_para_ordenar(x[0]))

        if len(rooms_unicos) >= 2:
            nombre_combinado = "{} : {}, {} : {}".format(
                rooms_unicos[0][0], rooms_unicos[0][1],
                rooms_unicos[1][0], rooms_unicos[1][1]
            )
        else:
            # Solo un ambiente encontrado
            nombre_combinado = "{} : {}".format(
                rooms_unicos[0][0], rooms_unicos[0][1]
            )

        return asignar_ambiente_puerta_ventana(elemento, nombre_combinado, nombre_combinado, failed_list)

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
    - S&P_AMBIENTE: solo el nombre en MAYÚSCULAS
    - COBie.Component.Space: formato "numero : NOMBRE" en MAYÚSCULAS
    Retorna True si se asignó correctamente al menos uno, False si ambos fallan.
    """
    exito = False
    
    try:
        # Asignar a S&P_AMBIENTE (solo nombre en MAYÚSCULAS)
        prm_ambiente = elemento.LookupParameter(PARAM_NAME)
        if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
            prm_ambiente.Set(nombre_ambiente.upper())
            exito = True
        
        # Asignar a COBie.Component.Space (numero : NOMBRE en MAYÚSCULAS)
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
            if numero_ambiente:
                valor_cobie = "{} : {}".format(numero_ambiente, nombre_ambiente.upper())
            else:
                valor_cobie = nombre_ambiente.upper()
            prm_cobie.Set(valor_cobie)
            exito = True
        
        if not exito:
            failed_list.append(elemento.Id)
            
    except Exception as e:
        output.print_md("**Error en elemento {}: {}**".format(elemento.Id, str(e)))
        failed_list.append(elemento.Id)
        return False
    
    return exito


def asignar_ambiente_puerta_ventana(elemento, nombre_combinado, valor_cobie, failed_list):
    """
    Asignación especial para puertas y ventanas con formato:
    - S&P_AMBIENTE: "23 : SALA, 26 : COMEDOR"
    - COBie.Component.Space: "23 : SALA, 26 : COMEDOR"
    """
    exito = False
    
    try:
        # Asignar a S&P_AMBIENTE
        prm_ambiente = elemento.LookupParameter(PARAM_NAME)
        if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
            prm_ambiente.Set(nombre_combinado)
            exito = True
        
        # Asignar a COBie.Component.Space (mismo valor)
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
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
        output.print_md("### ✅ Habitaciones encontradas en el modelo actual: {}".format(len(habs)))
        return habs
    
    # Si no hay habitaciones en el modelo actual, buscar en vínculos
    output.print_md("### ⚠️ No se encontraron habitaciones en el modelo actual")
    output.print_md("### 🔗 Buscando habitaciones en vínculos...")
    
    links = FilteredElementCollector(documento).OfClass(RevitLinkInstance).ToElements()
    
    if not links:
        forms.alert(
            "No se encontraron habitaciones en el modelo actual\n"
            "y no hay vínculos disponibles en el proyecto.",
            title="Sin habitaciones",
            exitscript=True
        )
    
    # Filtrar solo vínculos cargados
    links_cargados = [l for l in links if l.GetLinkDocument() is not None]
    
    if not links_cargados:
        forms.alert(
            "No se encontraron habitaciones en el modelo actual\n"
            "y no hay vínculos CARGADOS en el proyecto.\n\n"
            "Carga al menos un vínculo con habitaciones.",
            title="Sin vínculos cargados",
            exitscript=True
        )
    
    # Preparar lista de vínculos con información
    nom_links_info = []
    for link in links_cargados:
        link_doc = link.GetLinkDocument()
        if link_doc:
            # Contar habitaciones en el vínculo
            link_rooms = FilteredElementCollector(link_doc).OfCategory(
                doc.Settings.Categories.get_Item("Rooms")
            ).WhereElementIsNotElementType().ToElements()
            
            num_rooms = len([r for r in link_rooms if r.Area > 0])
            nom_links_info.append("{} ({} habitaciones)".format(link.Name, num_rooms))
    
    if not nom_links_info:
        forms.alert(
            "No se encontraron habitaciones en ningún vínculo cargado.",
            title="Sin habitaciones en vínculos",
            exitscript=True
        )
    
    # Selección de vínculos
    sel_links = forms.SelectFromList.show(
        sorted(nom_links_info), 
        multiselect=True,
        title="Selecciona vínculos con habitaciones",
        button_name="Usar estos vínculos"
    )
    
    if not sel_links:
        forms.alert("No se seleccionaron vínculos.", title="Cancelado", exitscript=True)
    
    # Extraer nombres originales (sin el conteo)
    nombres_originales = [s.split(" (")[0] for s in sel_links]
    
    habs = obtener_habitaciones_de_vinculos_seleccionados(documento, nombres_originales)
    
    if not habs:
        forms.alert(
            "No se pudieron obtener habitaciones de los vínculos seleccionados.",
            title="Error",
            exitscript=True
        )
    
    output.print_md("### ✅ Habitaciones encontradas en vínculos: {}".format(len(habs)))
    output.print_md("Vínculos utilizados: {}".format(", ".join(nombres_originales)))
    
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

# Verificar si hay puertas, ventanas o modelos genéricos seleccionados
tiene_puertas = "Puertas" in seleccion or "Doors" in seleccion
tiene_ventanas = "Ventanas" in seleccion or "Windows" in seleccion
tiene_genericos = "Modelos genéricos" in seleccion or "Generic Models" in seleccion
procesamiento_especial = tiene_puertas or tiene_ventanas or tiene_genericos

if procesamiento_especial:
    output.print_md("### ℹ️ Se aplicará procesamiento especial para puertas/ventanas/modelos genéricos (2 ambientes)")

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
    
    # Procesamiento especial para puertas, ventanas y modelos genéricos (si aplica)
    if procesamiento_especial:
        output.print_md("#### Procesando puertas, ventanas y modelos genéricos (2 ambientes)...")
        for e in elems:
            # Solo procesar si es puerta, ventana o modelo genérico
            try:
                cat_name = e.Category.Name if e.Category else ""
                es_elemento_especial = cat_name in ["Puertas", "Doors", "Ventanas", "Windows", "Modelos genéricos", "Generic Models"]
                
                if not es_elemento_especial:
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
    output.print_md("- **Puertas/Ventanas/Genéricos** (2 ambientes): {}".format(asignados_especial))
output.print_md("- **Fase 1** (dentro de habitación): {}".format(asignados_fase1))
output.print_md("- **Fase 2** (por proximidad): {}".format(asignados_fase2))
output.print_md("- **Fase 3** (como '{}'): {}".format(FALLBACK_VALUE, asignados_fase3))
output.print_md("- **Total asignados**: {}".format(total_asignados))
output.print_md("- **Ignorados** (COBie inactivo o parámetros llenos): {}".format(total_ignorados))
output.print_md("- **Sin asignar** (sin parámetros válidos): {}".format(len(failed_param)))

# Mostrar elementos asignados como "Activo" (para revisión manual)
if asignados_fase3 > 0:
    output.print_md("---")
    output.print_md("### ⚠️ Elementos asignados como '{}' (requieren revisión manual)".format(FALLBACK_VALUE))
    output.print_md("Total: {} elementos".format(asignados_fase3))
    output.print_md("Haz clic en los IDs para seleccionarlos en Revit:")
    
    # Recolectar elementos completos asignados como "Activo"
    elementos_activo = []
    for e in elems:
        if e.Id in elems_asignados:
            try:
                prm = e.LookupParameter(PARAM_NAME)
                if prm and prm.AsString() == FALLBACK_VALUE:
                    elementos_activo.append(e)
            except:
                continue
    
    # Mostrar con links seleccionables (máximo 50 para no saturar)
    muestra = elementos_activo[:50]
    output.print_md("")  # Línea en blanco para separación
    for elem in muestra:
        try:
            cat_name = elem.Category.Name if elem.Category else "Sin categoría"
            elem_name = elem.Name if hasattr(elem, 'Name') else "Sin nombre"
            
            # Crear link seleccionable
            output.print_md("- **{}** - {} - ID: {}".format(
                cat_name,
                elem_name,
                output.linkify(elem.Id)
            ))
        except Exception as e:
            output.print_md("- Error con elemento ID: {}".format(elem.Id.IntegerValue))

if failed_param:
    output.print_md("---")
    output.print_md("### ❌ Elementos sin parámetros válidos")
    sample = failed_param[:15]
    for elem_id in sample:
        # ❌ Esto tampoco funciona:
        # output.print_element(elem_id)
        
        # ✅ Usa esto:
        output.print_md("- ID: {}".format(output.linkify(elem_id)))
    
    if len(failed_param) > 15:
        output.print_md("*Mostrando 15 de {} elementos*".format(len(failed_param)))

forms.alert(
    "Proceso terminado:\n\n"
    "✅ {} elementos asignados correctamente\n"
    "⚠️ {} elementos asignados como '{}' (revisar manualmente)\n"
    "⏭️ {} elementos ignorados (COBie inactivo o ya llenos)\n"
    "❌ {} sin parámetros válidos\n\n"
    "Revisa la terminal para IDs seleccionables.".format(
        asignados_fase1 + asignados_fase2 + asignados_especial, 
        asignados_fase3,
        FALLBACK_VALUE,
        total_ignorados, 
        len(failed_param)
    ),
    title="Asignación Completada"
)