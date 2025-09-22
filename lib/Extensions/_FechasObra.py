# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import FamilyInstance
from datetime import datetime, timedelta
import random, re

# Generar calendario de semanas desde el lunes 01/04/2024
def generar_semanas():
    semanas = []
    inicio = datetime.strptime("01/04/2024", "%d/%m/%Y")
    for i in range(60):
        ini = inicio + timedelta(weeks=i)
        fin = ini + timedelta(days=6)
        semanas.append({
            "semana": "SEM" + str(i + 1).zfill(2),
            "inicio": ini,
            "fin": fin
        })
    return semanas

def buscar_semana(fecha, semanas):
    for s in semanas:
        if s["inicio"] <= fecha <= s["fin"]:
            return s["semana"]
    return "Fuera de rango"

# Obtener una fecha aleatoria dentro de una semana
def buscar_fecha_por_semana(nombre_semana, semanas):
    for s in semanas:
        if s["semana"] == nombre_semana:
            fecha_random = s["inicio"] + timedelta(days=random.randint(0, 6))
            return fecha_random.strftime("%Y-%m-%d")
    return None

# Incluir familias anidadas compartidas
def obtener_todos_los_elementos(elementos):
    todos = []
    for elem in elementos:
        todos.append(elem)
        if isinstance(elem, FamilyInstance):
            try:
                sub_ids = elem.GetSubComponentIds()
                for sid in sub_ids:
                    sub_elem = doc.GetElement(sid)
                    if sub_elem:
                        todos.append(sub_elem)
            except:
                pass
    return todos

# Función para corregir el formato de fecha y devolver string con formato "yyyy-mm-dd"
def corregir_formato_fecha(fecha_str):
    if not fecha_str or not isinstance(fecha_str, str):
        raise ValueError("Fecha vacía o no es cadena")

    # Reemplaza '/' por '-' para unificar
    fecha_normalizada = fecha_str.replace('/', '-').strip()

    # Patrones y formatos más comunes
    patrones = {
        r'^\d{2}-\d{2}-\d{4}$': '%d-%m-%Y',
        r'^\d{4}-\d{2}-\d{2}$': '%Y-%m-%d',
    }

    for patron, formato in patrones.items():
        if re.match(patron, fecha_normalizada):
            try:
                fecha_obj = datetime.strptime(fecha_normalizada, formato)
                return fecha_obj.strftime('%Y-%m-%d')
            except ValueError:
                pass

    # Intentar formatos comunes adicionales
    formatos_posibles = ['%d-%m-%Y', '%Y-%m-%d', '%d-%m-%y', '%y-%m-%d']
    for formato in formatos_posibles:
        try:
            fecha_obj = datetime.strptime(fecha_normalizada, formato)
            return fecha_obj.strftime('%Y-%m-%d')
        except ValueError:
            continue

    raise ValueError("Formato de fecha no reconocido o inválido: '" + fecha_str + "'")