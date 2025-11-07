# -*- coding: utf-8 -*-
__title__ = "Assign room"

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import (
    StorageType, RevitLinkInstance, FilteredElementCollector, BuiltInParameter, ElementId
)
from Extensions._Modulo import obtener_nombre_archivo, validar_nombre
from Extensions._utils import (
    obtener_mapeo_nombres_categorias,
    obtener_elementos_de_categorias,
    obtener_habitaciones_y_espacios,
    get_room_name,
    puntos_representativos,
    is_point_inside,
    obtener_room_mas_cercano
)

doc = revit.doc
output = script.get_output()

# Constantes
TOLERANCIA_PIES = 0.410105  # 12.5 cm en pies
TOLERANCIA_NIVEL_PIES = 1.0  # Tolerancia vertical para considerar mismo nivel (30 cm)
TOLERANCIA_PUERTA_VENTANA_PIES = 16.404  # 5 metros - distancia máxima razonable al centro del room
PARAM_NAME = "S&P_AMBIENTE"
PARAM_COBIE = "COBie.Component.Space"
PARAM_COBIE_BOOL = "COBie"
FALLBACK_VALUE = "Activo"

# Variable global para controlar modo sobrescritura
SOBRESCRIBIR_GLOBAL = True


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


def obtener_nivel_z(elemento):
    """
    Obtiene la coordenada Z del nivel del elemento.
    Para puertas y ventanas, usa el punto de inserción.
    """
    try:
        # Intentar obtener el nivel del elemento
        if hasattr(elemento, 'LevelId'):
            nivel_id = elemento.LevelId
            if nivel_id and nivel_id != ElementId.InvalidElementId:
                nivel = doc.GetElement(nivel_id)
                if nivel:
                    return nivel.Elevation
        
        # Si no se puede obtener el nivel, usar el punto de ubicación
        if hasattr(elemento, 'Location'):
            location = elemento.Location
            if location and hasattr(location, 'Point'):
                return location.Point.Z
    except:
        pass
    
    return None


def obtener_habitaciones_de_vinculos_transformadas(documento, nombres_vinculos):
    """
    Obtiene habitaciones de vínculos seleccionados y las transforma al sistema
    de coordenadas del documento actual.
    Retorna una lista de tuplas (room, transform) para usar en cálculos.
    """
    from Autodesk.Revit.DB import BuiltInCategory
    
    links = FilteredElementCollector(documento).OfClass(RevitLinkInstance).ToElements()
    links_dict = {l.Name: l for l in links if l.GetLinkDocument() is not None}
    
    habitaciones_transformadas = []
    
    for nombre_vinculo in nombres_vinculos:
        if nombre_vinculo not in links_dict:
            continue
        
        link_instance = links_dict[nombre_vinculo]
        link_doc = link_instance.GetLinkDocument()
        transform = link_instance.GetTotalTransform()
        
        if not link_doc:
            continue
        
        try:
            # Intentar primero con categoría por nombre
            rooms = FilteredElementCollector(link_doc).OfCategory(
                link_doc.Settings.Categories.get_Item("Rooms")
            ).WhereElementIsNotElementType().ToElements()
        except:
            # Si falla, usar BuiltInCategory
            try:
                rooms = FilteredElementCollector(link_doc).OfCategory(
                    BuiltInCategory.OST_Rooms
                ).WhereElementIsNotElementType().ToElements()
            except:
                continue
        
        # Filtrar rooms con área > 0 y agregar con su transform
        for room in rooms:
            try:
                if room.Area > 0:
                    habitaciones_transformadas.append((room, transform))
            except:
                continue
    
    return habitaciones_transformadas


def obtener_punto_transformado(room, transform):
    """Obtiene el punto de ubicación de un room transformado al documento actual."""
    try:
        location = room.Location
        if location:
            room_point = location.Point
            # Transformar el punto al sistema de coordenadas del documento actual
            return transform.OfPoint(room_point)
    except:
        pass
    return None


def is_point_inside_transformed(room, transform, punto):
    """
    Verifica si un punto está dentro de un room considerando la transformación.
    Transforma el punto del documento actual al sistema del vínculo para la verificación.
    """
    try:
        # Invertir la transformación para llevar el punto al sistema del vínculo
        inverse_transform = transform.Inverse
        punto_en_vinculo = inverse_transform.OfPoint(punto)
        
        # Ahora usar is_point_inside con el punto transformado
        return is_point_inside(room, punto_en_vinculo)
    except:
        return False


def obtener_rooms_cercanos_mismo_nivel_con_limite(habitaciones_transformadas, punto, nivel_z_elemento, max_distancia_pies):
    """
    Obtiene habitaciones cercanas EN EL MISMO NIVEL y dentro de una distancia máxima.
    Solo considera la distancia en X e Y (horizontal).
    """
    distancias = []
    
    for room, transform in habitaciones_transformadas:
        try:
            punto_room = obtener_punto_transformado(room, transform)
            if not punto_room:
                continue
            
            # Verificar que el room esté en el mismo nivel (tolerancia vertical)
            if nivel_z_elemento is not None:
                diferencia_z = abs(punto_room.Z - nivel_z_elemento)
                if diferencia_z > TOLERANCIA_NIVEL_PIES:
                    continue
            
            # Calcular distancia SOLO en horizontal (X, Y)
            dx = punto_room.X - punto.X
            dy = punto_room.Y - punto.Y
            distancia_horizontal = (dx*dx + dy*dy) ** 0.5
            
            # Solo agregar si está dentro de la distancia máxima
            if distancia_horizontal <= max_distancia_pies:
                distancias.append((distancia_horizontal, room, transform))
        except:
            continue
    
    # Ordenar por distancia horizontal
    distancias.sort(key=lambda x: x[0])
    return distancias


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


def procesar_puerta_ventana(elemento, habitaciones_transformadas, failed_list):
    """
    Procesa puertas y ventanas para asignar hasta 2 habitaciones colindantes.
    Estrategia mejorada:
    1. FromRoom y ToRoom (si están en doc actual)
    2. Verificar si puntos están DENTRO de rooms
    3. Si encontró 1, buscar la segunda con tolerancia más generosa (10m)
    4. Si no encontró ninguna, buscar por proximidad normal (5m)
    """
    try:
        rooms_unicos = []
        ids_usados = set()
        numeros_usados = set()
        
        # Obtener el nivel Z del elemento
        nivel_z_elemento = obtener_nivel_z(elemento)
        
        # PASO 1: Intentar FromRoom y ToRoom
        from_room = None
        to_room = None
        try:
            from_room = elemento.FromRoom[doc.Phase] if hasattr(elemento, "FromRoom") else None
            to_room = elemento.ToRoom[doc.Phase] if hasattr(elemento, "ToRoom") else None
        except:
            pass

        if from_room:
            nombre = get_room_name(from_room)
            numero = get_room_number(from_room)
            if nombre and numero:
                nombre_upper = nombre.upper()
                rooms_unicos.append((numero, nombre_upper, from_room.Id))
                ids_usados.add(from_room.Id.IntegerValue)
                numeros_usados.add(numero)

        if to_room and to_room.Id.IntegerValue not in ids_usados:
            nombre = get_room_name(to_room)
            numero = get_room_number(to_room)
            if nombre and numero:
                nombre_upper = nombre.upper()
                if numero not in numeros_usados:
                    rooms_unicos.append((numero, nombre_upper, to_room.Id))
                    ids_usados.add(to_room.Id.IntegerValue)
                    numeros_usados.add(numero)

        # PASO 2: Si tenemos menos de 2, verificar si puntos están DENTRO de rooms
        if len(rooms_unicos) < 2:
            pts = puntos_representativos(elemento) or []
            
            for punto in pts:
                if not punto:
                    continue
                
                if len(rooms_unicos) >= 2:
                    break
                
                z_ref = nivel_z_elemento if nivel_z_elemento is not None else punto.Z
                
                # Verificar en cada habitación del mismo nivel
                for room, transform in habitaciones_transformadas:
                    if len(rooms_unicos) >= 2:
                        break
                    
                    # Saltar si ya lo tenemos
                    if room.Id.IntegerValue in ids_usados:
                        continue
                    
                    # Verificar mismo nivel
                    punto_room = obtener_punto_transformado(room, transform)
                    if not punto_room:
                        continue
                    
                    if z_ref is not None:
                        diferencia_z = abs(punto_room.Z - z_ref)
                        if diferencia_z > TOLERANCIA_NIVEL_PIES:
                            continue
                    
                    # Verificar si el punto está DENTRO
                    if is_point_inside_transformed(room, transform, punto):
                        nombre = get_room_name(room)
                        numero = get_room_number(room)
                        
                        if nombre and numero:
                            nombre_upper = nombre.upper()
                            if numero not in numeros_usados:
                                rooms_unicos.append((numero, nombre_upper, room.Id))
                                ids_usados.add(room.Id.IntegerValue)
                                numeros_usados.add(numero)
        
        # PASO 3: Si tenemos exactamente 1, buscar la segunda con tolerancia GENEROSA (10m)
        # Si no tenemos ninguna, buscar con tolerancia normal (5m)
        if len(rooms_unicos) < 2:
            pts = puntos_representativos(elemento) or []
            
            # CLAVE: Si ya encontró 1, usar tolerancia mayor para la segunda
            tolerancia_busqueda = 32.8084 if len(rooms_unicos) == 1 else TOLERANCIA_PUERTA_VENTANA_PIES  # 10m si ya tiene 1, sino 5m
            
            candidatos = []
            for punto in pts:
                if not punto:
                    continue
                
                z_ref = nivel_z_elemento if nivel_z_elemento is not None else punto.Z
                
                # Buscar habitaciones cercanas
                cercanos = obtener_rooms_cercanos_mismo_nivel_con_limite(
                    habitaciones_transformadas, 
                    punto,
                    z_ref,
                    tolerancia_busqueda
                )
                
                for distancia, room, transform in cercanos:
                    if room.Id.IntegerValue not in [r.Id.IntegerValue for d, r, t in candidatos]:
                        candidatos.append((distancia, room, transform))
            
            # Ordenar por distancia
            candidatos.sort(key=lambda x: x[0])
            
            # Agregar hasta tener 2 rooms diferentes
            for distancia, room, transform in candidatos:
                if len(rooms_unicos) >= 2:
                    break
                
                if room.Id.IntegerValue in ids_usados:
                    continue
                
                nombre = get_room_name(room)
                numero = get_room_number(room)
                
                if not nombre or not numero:
                    continue
                
                nombre_upper = nombre.upper()
                
                if numero not in numeros_usados:
                    rooms_unicos.append((numero, nombre_upper, room.Id))
                    ids_usados.add(room.Id.IntegerValue)
                    numeros_usados.add(numero)

        # Verificar que tengamos al menos un ambiente
        if not rooms_unicos:
            return False

        # PASO 4: Ordenar por número y construir cadena
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
    
    Si SOBRESCRIBIR_GLOBAL es False, solo asigna a parámetros vacíos individualmente.
    Retorna True si se asignó al menos uno, False si no se pudo asignar ninguno.
    """
    asignado_ambiente = False
    asignado_cobie = False
    
    try:
        # PARÁMETRO 1: S&P_AMBIENTE
        prm_ambiente = elemento.LookupParameter(PARAM_NAME)
        if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
            debe_asignar = SOBRESCRIBIR_GLOBAL
            
            if not SOBRESCRIBIR_GLOBAL:
                # Solo asignar si está vacío
                valor_actual = prm_ambiente.AsString()
                if valor_actual is None or valor_actual.strip() == "":
                    debe_asignar = True
            
            if debe_asignar:
                prm_ambiente.Set(nombre_ambiente.upper())
                asignado_ambiente = True
        
        # PARÁMETRO 2: COBie.Component.Space
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
            debe_asignar = SOBRESCRIBIR_GLOBAL
            
            if not SOBRESCRIBIR_GLOBAL:
                # Solo asignar si está vacío
                valor_actual = prm_cobie.AsString()
                if valor_actual is None or valor_actual.strip() == "":
                    debe_asignar = True
            
            if debe_asignar:
                if numero_ambiente:
                    valor_cobie = "{} : {}".format(numero_ambiente, nombre_ambiente.upper())
                else:
                    valor_cobie = nombre_ambiente.upper()
                prm_cobie.Set(valor_cobie)
                asignado_cobie = True
        
        # Retornar True si se asignó AL MENOS UNO
        if asignado_ambiente or asignado_cobie:
            return True
        else:
            # No se pudo asignar ninguno
            failed_list.append(elemento.Id)
            return False
            
    except Exception as e:
        output.print_md("**Error en elemento {}: {}**".format(elemento.Id, str(e)))
        failed_list.append(elemento.Id)
        return False


def asignar_ambiente_puerta_ventana(elemento, nombre_combinado, valor_cobie, failed_list):
    """
    Asignación especial para puertas y ventanas con formato:
    - S&P_AMBIENTE: "23 : SALA, 26 : COMEDOR"
    - COBie.Component.Space: "23 : SALA, 26 : COMEDOR"
    
    Si SOBRESCRIBIR_GLOBAL es False, solo asigna a parámetros vacíos individualmente.
    """
    asignado_ambiente = False
    asignado_cobie = False
    
    try:
        # PARÁMETRO 1: S&P_AMBIENTE
        prm_ambiente = elemento.LookupParameter(PARAM_NAME)
        if prm_ambiente and prm_ambiente.StorageType == StorageType.String:
            debe_asignar = SOBRESCRIBIR_GLOBAL
            
            if not SOBRESCRIBIR_GLOBAL:
                # Solo asignar si está vacío
                valor_actual = prm_ambiente.AsString()
                if valor_actual is None or valor_actual.strip() == "":
                    debe_asignar = True
            
            if debe_asignar:
                prm_ambiente.Set(nombre_combinado)
                asignado_ambiente = True
        
        # PARÁMETRO 2: COBie.Component.Space
        prm_cobie = elemento.LookupParameter(PARAM_COBIE)
        if prm_cobie and prm_cobie.StorageType == StorageType.String:
            debe_asignar = SOBRESCRIBIR_GLOBAL
            
            if not SOBRESCRIBIR_GLOBAL:
                # Solo asignar si está vacío
                valor_actual = prm_cobie.AsString()
                if valor_actual is None or valor_actual.strip() == "":
                    debe_asignar = True
            
            if debe_asignar:
                prm_cobie.Set(valor_cobie)
                asignado_cobie = True
        
        # Retornar True si se asignó AL MENOS UNO
        if asignado_ambiente or asignado_cobie:
            return True
        else:
            # No se pudo asignar ninguno
            failed_list.append(elemento.Id)
            return False
            
    except Exception as e:
        output.print_md("**Error en elemento {}: {}**".format(elemento.Id, str(e)))
        failed_list.append(elemento.Id)
        return False


def procesar_elemento_fase1(elemento, habitaciones_transformadas, failed_list):
    """Fase 1: Busca si algún punto del elemento está dentro de una habitación."""
    pts = puntos_representativos(elemento) or []
    for p in pts:
        if not p:
            continue
        try:
            # Buscar en habitaciones transformadas
            for room, transform in habitaciones_transformadas:
                if is_point_inside_transformed(room, transform, p):
                    nombre = get_room_name(room)
                    numero = get_room_number(room)
                    return asignar_ambiente(elemento, nombre, numero, failed_list)
        except Exception:
            continue
    return False


def procesar_elemento_fase2(elemento, habitaciones_transformadas, failed_list):
    """Fase 2: Busca la habitación más cercana dentro de la tolerancia."""
    pts = puntos_representativos(elemento) or []
    for p in pts:
        if not p:
            continue
        try:
            # Buscar el room más cercano con transformación
            distancia_minima = float('inf')
            room_cercana = None
            
            for room, transform in habitaciones_transformadas:
                punto_room = obtener_punto_transformado(room, transform)
                if not punto_room:
                    continue
                
                dx = punto_room.X - p.X
                dy = punto_room.Y - p.Y
                dz = punto_room.Z - p.Z
                distancia = (dx*dx + dy*dy + dz*dz) ** 0.5
                
                if distancia < distancia_minima and distancia <= TOLERANCIA_PIES:
                    distancia_minima = distancia
                    room_cercana = room
            
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
        # Convertir a formato (room, transform_identidad) para compatibilidad
        from Autodesk.Revit.DB import Transform
        return [(h, Transform.Identity) for h in habs]
    
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
    from Autodesk.Revit.DB import BuiltInCategory
    nom_links_info = []
    for link in links_cargados:
        link_doc = link.GetLinkDocument()
        if link_doc:
            try:
                # Intentar con categoría por nombre
                link_rooms = FilteredElementCollector(link_doc).OfCategory(
                    link_doc.Settings.Categories.get_Item("Rooms")
                ).WhereElementIsNotElementType().ToElements()
                
                num_rooms = len([r for r in link_rooms if r.Area > 0])
                nom_links_info.append("{} ({} habitaciones)".format(link.Name, num_rooms))
            except:
                # Si falla, usar BuiltInCategory
                try:
                    link_rooms = FilteredElementCollector(link_doc).OfCategory(
                        BuiltInCategory.OST_Rooms
                    ).WhereElementIsNotElementType().ToElements()
                    
                    num_rooms = len([r for r in link_rooms if r.Area > 0])
                    nom_links_info.append("{} ({} habitaciones)".format(link.Name, num_rooms))
                except:
                    nom_links_info.append("{} (0 habitaciones)".format(link.Name))
    
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
    
    habs = obtener_habitaciones_de_vinculos_transformadas(documento, nombres_originales)
    
    if not habs:
        forms.alert(
            "No se pudieron obtener habitaciones de los vínculos seleccionados.",
            title="Error",
            exitscript=True
        )
    
    output.print_md("### ✅ Habitaciones encontradas en vínculos: {}".format(len(habs)))
    output.print_md("Vínculos utilizados: {}".format(", ".join(nombres_originales)))
    
    return habs

nombre_archivo = obtener_nombre_archivo()
if not validar_nombre(nombre_archivo):
    script.exit()

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

# ==================== PREGUNTA: SOBRESCRIBIR ====================
sobrescribir = forms.alert(
    "¿Deseas SOBRESCRIBIR los parámetros que ya tienen valores?\n\n"
    "SI: Sobrescribir todos (actualizar valores existentes)\n"
    "NO: Solo llenar parámetros vacíos (cada parámetro se evalúa individualmente)",
    title="Modo de sobrescritura",
    yes=True,
    no=True
)

# Configurar variable global
SOBRESCRIBIR_GLOBAL = sobrescribir

if sobrescribir:
    output.print_md("### ⚠️ MODO: Sobrescribir valores existentes")
else:
    output.print_md("### ℹ️ MODO: Solo llenar parámetros vacíos (cada uno individualmente)")

# Verificar si hay puertas, ventanas o modelos genéricos seleccionados
tiene_puertas = "Puertas" in seleccion or "Doors" in seleccion
tiene_ventanas = "Ventanas" in seleccion or "Windows" in seleccion
tiene_genericos = "Modelos genéricos" in seleccion or "Generic Models" in seleccion
procesamiento_especial = tiene_puertas or tiene_ventanas or tiene_genericos

if procesamiento_especial:
    output.print_md("### ℹ️ Puertas/Ventanas: Detectan hasta 2 ambientes colindantes (máximo {} m de distancia)".format(
        round(TOLERANCIA_PUERTA_VENTANA_PIES * 0.3048, 1)))

# 3) Obtener habitaciones (ahora como lista de tuplas (room, transform))
habs = obtener_habitaciones(doc)
output.print_md("### Habitaciones disponibles: {}".format(len(habs)))

# 4) Obtener elementos a procesar
elems = obtener_elementos_de_categorias(doc, sels_cats)
output.print_md("### Elementos encontrados: {}".format(len(elems)))

# Tracking
elems_asignados = set()
elems_ignorados_cobie = []
failed_param = []
asignados_fase1 = 0
asignados_fase2 = 0
asignados_fase3 = 0
asignados_especial = 0
elementos_fase3_ids = []

# ==================== PROCESAMIENTO EN UNA SOLA TRANSACCIÓN ====================

with revit.Transaction("Asignar Ambiente"):
    
    # Procesamiento especial para puertas, ventanas y modelos genéricos (si aplica)
    if procesamiento_especial:
        output.print_md("#### Procesando puertas, ventanas y modelos genéricos...")
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
                
                # SIEMPRE intenta procesar el elemento
                # La función asignar_ambiente_puerta_ventana maneja individualmente cada parámetro
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
        
        # SIEMPRE intenta procesar el elemento
        # La función asignar_ambiente maneja individualmente cada parámetro
        if procesar_elemento_fase1(e, habs, failed_param):
            elems_asignados.add(e.Id)
            asignados_fase1 += 1
    
    # Fase 2: Proximidad
    output.print_md("#### Procesando Fase 2: Elementos por proximidad...")
    for e in elems:
        if e.Id in elems_asignados or e.Id in elems_ignorados_cobie:
            continue
        
        if procesar_elemento_fase2(e, habs, failed_param):
            elems_asignados.add(e.Id)
            asignados_fase2 += 1
    
    # Fase 3: Resto como "Activo"
    output.print_md("#### Procesando Fase 3: Elementos restantes como '{}'...".format(FALLBACK_VALUE))
    for e in elems:
        if e.Id in elems_asignados or e.Id in elems_ignorados_cobie:
            continue
        
        if asignar_ambiente(e, FALLBACK_VALUE, "", failed_param):
            elems_asignados.add(e.Id)
            asignados_fase3 += 1
            elementos_fase3_ids.append(e.Id)

# ==================== RESULTADOS ====================

total_asignados = asignados_fase1 + asignados_fase2 + asignados_fase3 + asignados_especial

output.print_md("---")
output.print_md("### 📊 Resumen de Resultados")
if asignados_especial > 0:
    output.print_md("- **Puertas/Ventanas/Genéricos** (hasta 2 ambientes): {}".format(asignados_especial))
output.print_md("- **Fase 1** (dentro de habitación): {}".format(asignados_fase1))
output.print_md("- **Fase 2** (por proximidad): {}".format(asignados_fase2))
output.print_md("- **Fase 3** (como '{}'): {}".format(FALLBACK_VALUE, asignados_fase3))
output.print_md("- **Total asignados**: {}".format(total_asignados))
output.print_md("- **Ignorados** (COBie inactivo): {}".format(len(elems_ignorados_cobie)))
output.print_md("- **Sin asignar** (sin parámetros válidos): {}".format(len(failed_param)))

# Mostrar elementos asignados como "Activo" (para revisión manual)
if asignados_fase3 > 0:
    output.print_md("---")
    output.print_md("### ⚠️ Elementos asignados como '{}' (requieren revisión manual)".format(FALLBACK_VALUE))
    output.print_md("Total: {} elementos".format(asignados_fase3))
    output.print_md("Haz clic en los IDs para seleccionarlos en Revit:")
    output.print_md("")
    
    # Mostrar con links seleccionables (máximo 50 para no saturar)
    muestra_ids = elementos_fase3_ids[:50]
    
    for elem_id in muestra_ids:
        try:
            elem = doc.GetElement(elem_id)
            if elem:
                cat_name = elem.Category.Name if elem.Category else "Sin categoría"
                
                # Intentar obtener nombre del tipo
                elem_name = ""
                try:
                    tipo = doc.GetElement(elem.GetTypeId())
                    if tipo:
                        elem_name = tipo.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                except:
                    pass
                
                if not elem_name:
                    elem_name = "Sin nombre"
                
                # Crear link seleccionable
                output.print_md("- **{}** | {} | ID: {}".format(
                    cat_name,
                    elem_name,
                    output.linkify(elem_id)
                ))
            else:
                output.print_md("- ID: {}".format(output.linkify(elem_id)))
        except Exception as ex:
            output.print_md("- ID: {} (Error: {})".format(output.linkify(elem_id), str(ex)))
    
    if len(elementos_fase3_ids) > 50:
        output.print_md("")
        output.print_md("*Mostrando 50 de {} elementos*".format(len(elementos_fase3_ids)))

if failed_param:
    output.print_md("---")
    output.print_md("### ❌ Elementos sin parámetros válidos")
    output.print_md("")
    sample = failed_param[:15]
    
    for elem_id in sample:
        try:
            elem = doc.GetElement(elem_id)
            if elem and elem.Category:
                cat_name = elem.Category.Name
                output.print_md("- **{}** | ID: {}".format(cat_name, output.linkify(elem_id)))
            else:
                output.print_md("- ID: {}".format(output.linkify(elem_id)))
        except:
            output.print_md("- ID: {}".format(output.linkify(elem_id)))
    
    if len(failed_param) > 15:
        output.print_md("")
        output.print_md("*Mostrando 15 de {} elementos*".format(len(failed_param)))

# Mensaje final personalizado
mensaje_final = "Proceso terminado:\n\n"
mensaje_final += "✅ {} elementos asignados correctamente\n".format(asignados_fase1 + asignados_fase2 + asignados_especial)

if asignados_fase3 > 0:
    mensaje_final += "⚠️ {} elementos asignados como '{}' (revisar manualmente)\n".format(asignados_fase3, FALLBACK_VALUE)

if len(elems_ignorados_cobie) > 0:
    mensaje_final += "⏭️ {} elementos ignorados (COBie inactivo)\n".format(len(elems_ignorados_cobie))

if len(failed_param) > 0:
    mensaje_final += "❌ {} sin parámetros válidos\n".format(len(failed_param))

mensaje_final += "\nRevisa la terminal para IDs seleccionables."

forms.alert(mensaje_final, title="Asignación Completada")