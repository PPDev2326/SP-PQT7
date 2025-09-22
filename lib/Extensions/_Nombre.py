# -*- coding: utf-8 -*-

import re

def obtener_nombre_corto(nombre):
    # Elimina todo lo que venga después de "1er", "2do", "3ro" o "-"
    nombre_limpio = re.sub(r'\s*(1er|2do|3ro|3er|4to|5to|\(|-).*$', '', nombre, flags=re.IGNORECASE).strip()
    return capitalizar_respetando_mayusculas(nombre_limpio)


def obtener_nombre_base_para_contador(nombre):
    """Devuelve el nombre base ignorando cualquier número final, manteniendo mayúsculas originales y evitando problemas de agrupación."""
    nombre = nombre.strip()
    nombre_sin_numero = re.sub(r'\s*\d+$', '', nombre)
    nombre_normalizado = re.sub(r'\s+', ' ', nombre_sin_numero).strip()
    return capitalizar_respetando_mayusculas(nombre_normalizado)


def generar_abreviacion(nombre_base, contador):
    """Genera una abreviación tomando las primeras letras significativas (>= 3 letras)
    de las primeras dos palabras útiles, con un contador correlativo como AP-01."""
    
    # Eliminar números al final del nombre
    nombre_base_sin_numero = re.sub(r'\s*\d+$', '', nombre_base).strip()
    
    # Separar por palabras
    palabras = nombre_base_sin_numero.split()
    
    letras = ''
    for palabra in palabras:
        palabra_limpia = re.sub(r'\W+', '', palabra)  # quitar signos o símbolos
        if len(palabra_limpia) >= 3:  # solo contar palabras de al menos 3 letras
            letras += palabra_limpia[0].upper()
        if len(letras) == 2:  # detener al obtener 2 letras
            break
    
    # Asegurar formato del contador con ceros
    if contador < 10:
        return "{}-0{}".format(letras, contador)
    else:
        return "{}-{}".format(letras, contador)

def capitalizar_respetando_mayusculas(texto):
    palabras = texto.strip().split()
    resultado = []

    for i, palabra in enumerate(palabras):
        es_mayuscula_corta = palabra.isupper() and len(palabra) <= 3

        if i == 0:
            resultado.append(palabra.capitalize())
        elif es_mayuscula_corta:
            # Mantener si está al final o si la anterior o siguiente no son palabras largas
            es_ultima = i == len(palabras) - 1
            anterior_larga = len(palabras[i - 1]) >= 4
            siguiente_larga = len(palabras[i + 1]) >= 4 if i + 1 < len(palabras) else False

            if es_ultima or not (anterior_larga and siguiente_larga):
                resultado.append(palabra)  # mantener mayúscula como "I", "II", "A"
            else:
                resultado.append(palabra.lower())  # entre dos palabras largas: minúscula
        else:
            resultado.append(palabra.lower())

    return ' '.join(resultado)