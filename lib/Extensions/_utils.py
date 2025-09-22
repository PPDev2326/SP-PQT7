# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import SpatialElementBoundaryOptions, BoundingBoxXYZ
from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, FilteredElementCollector, 
    StorageType, SpatialElementBoundaryOptions, ElementCategoryFilter,
    RevitLinkInstance
)

def obtener_mapeo_nombres_categorias(doc):
    mapeo = {}
    for bic in BuiltInCategory.GetValues(BuiltInCategory):
        try:
            cat = doc.Settings.Categories.get_Item(bic)
            if cat:
                mapeo[cat.Name] = bic
        except:
            continue
    return mapeo

def obtener_elementos_de_categorias(doc, categorias):
    elementos = []
    for bic in categorias:
        filtro = ElementCategoryFilter(bic)
        elems = FilteredElementCollector(doc).WherePasses(filtro).WhereElementIsNotElementType().ToElements()
        elementos.extend(elems)
    return elementos

def obtener_habitaciones_y_espacios(doc):
    elementos = []
    elementos.extend(FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements())
    elementos.extend(FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_MEPSpaces)
        .WhereElementIsNotElementType()
        .ToElements())
    return elementos

def obtener_habitaciones_de_vinculos_seleccionados(doc, nombres_vinculos):
    habitaciones = []
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    for link in link_instances:
        try:
            if link.Name in nombres_vinculos:
                link_doc = link.GetLinkDocument()
                if link_doc:
                    rooms = FilteredElementCollector(link_doc)\
                        .OfCategory(BuiltInCategory.OST_Rooms)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
                    habitaciones.extend(rooms)
        except:
            continue
    return habitaciones

def get_room_name(room):
    try:
        if hasattr(room, 'Name') and room.Name:
            return room.Name
        param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        if param and param.HasValue:
            return param.AsString()
    except:
        return None

def obtener_room_mas_cercano(habitaciones, punto, tolerancia=0.328084):
    """Devuelve la habitación más cercana al punto si está dentro del rango de tolerancia (en pies)."""
    room_mas_cercano = None
    distancia_minima = float('inf')

    for room in habitaciones:
        try:
            boundaries = room.GetBoundarySegments(SpatialElementBoundaryOptions())
            if not boundaries:
                continue

            for segments in boundaries:
                for segment in segments:
                    curva = segment.GetCurve()
                    if curva:
                        proyeccion = curva.Project(punto)
                        if proyeccion:
                            punto_proyectado = proyeccion.XYZPoint
                            distancia = punto.DistanceTo(punto_proyectado)

                            # Asegurarse de que esté dentro de la tolerancia
                            if distancia <= tolerancia and distancia < distancia_minima:
                                distancia_minima = distancia
                                room_mas_cercano = room
        except:
            continue

    return room_mas_cercano

def puntos_representativos(elem):
    puntos = []
    if hasattr(elem.Location, 'Point') and elem.Location.Point:
        puntos.append(elem.Location.Point)
    elif hasattr(elem.Location, 'Curve') and elem.Location.Curve:
        curva = elem.Location.Curve
        puntos.append(curva.Evaluate(0.5, True))  # punto medio
        puntos.append(curva.GetEndPoint(0))
        puntos.append(curva.GetEndPoint(1))

    bbox = elem.get_BoundingBox(None)
    if bbox:
        centro = (bbox.Min + bbox.Max) * 0.5
        puntos.append(centro)
    return puntos


def punto_dentro_rango_z(room, punto):
    bbox = room.get_BoundingBox(None)
    if bbox:
        return bbox.Min.Z <= punto.Z <= bbox.Max.Z
    return True

from Autodesk.Revit.DB import SpatialElementBoundaryOptions, XYZ

def punto_en_poligono_2d(pt, poly):
    """
    Ray‑casting: devuelve True si el punto 2D (X,Y) está dentro del polígono.
    `poly` es una lista de XYZ (solo usamos X,Y).
    """
    x, y = pt.X, pt.Y
    inside = False
    n = len(poly)
    for i in range(n):
        j = (i + n - 1) % n
        xi, yi = poly[i].X, poly[i].Y
        xj, yj = poly[j].X, poly[j].Y
        # ¿Cruza la línea horizontal?
        intersect = ((yi > y) != (yj > y)) and \
                    (x < (xj - xi) * (y - yi) / (yj - yi + 1e-9) + xi)
        if intersect:
            inside = not inside
    return inside

def is_point_in_room_2d(room, punto):
    """Devuelve True si (X,Y) de `punto` está dentro del contorno 2D del room."""
    opts = SpatialElementBoundaryOptions()
    loops = room.GetBoundarySegments(opts)
    if not loops:
        return False
    # Tomamos solo el primer loop exterior
    exterior = loops[0]
    poly = []
    for seg in exterior:
        curve = seg.GetCurve()
        # Tomamos ambos extremos para asegurarnos de cerrar correctamente
        poly.append(XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, 0))
    return punto_en_poligono_2d(XYZ(punto.X, punto.Y, 0), poly)

from Autodesk.Revit.DB import SpatialElementBoundaryOptions, XYZ

def punto_en_poligono_2d(pt, poly):
    x, y = pt.X, pt.Y
    inside = False
    n = len(poly)
    for i in range(n):
        j = (i + n - 1) % n
        xi, yi = poly[i].X, poly[i].Y
        xj, yj = poly[j].X, poly[j].Y
        intersect = ((yi > y) != (yj > y)) and \
                    (x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-9) + xi)
        if intersect:
            inside = not inside
    return inside

def is_point_in_room_2d(room, punto):
    opts = SpatialElementBoundaryOptions()
    loops = room.GetBoundarySegments(opts)
    if not loops: 
        return False
    exterior = loops[0]
    poly = []
    for seg in exterior:
        curve = seg.GetCurve()
        # muestro los extremos y el punto medio de cada segmento
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        pm = curve.Evaluate(0.5, True)
        poly.extend([XYZ(p0.X, p0.Y, 0), XYZ(pm.X, pm.Y, 0), XYZ(p1.X, p1.Y, 0)])
    return punto_en_poligono_2d(punto, poly)

def punto_en_bounding_box(punto, bbox):
    if not bbox:
        return False
    return (bbox.Min.X <= punto.X <= bbox.Max.X and
            bbox.Min.Y <= punto.Y <= bbox.Max.Y and
            bbox.Min.Z <= punto.Z <= bbox.Max.Z)

def is_point_inside(room, punto):
    # 1) rango Z
    bbox = room.get_BoundingBox(None)
    if bbox and not punto_en_bounding_box(punto, bbox):
        return False

    # 2) si es local, intento IsPointInRoom
    try:
        if not room.Document.IsLinked and room.IsPointInRoom(punto):
            return True
    except:
        pass

    # 3) fallback 2D (funciona para locales y vinculados)
    return is_point_in_room_2d(room, punto)