# -*- coding: utf-8 -*-
import re
# import clr
# clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import Document

from pyrevit import revit

# Expresión regular corregida para diferenciar AR/DT según el número final
patron = re.compile(r"""
^200(04[5-7]|049|053|05[7-8]|06[013]|070)        # Código inicial
(-CSSP001)                                       # Parte fija
-(000|41[5-9]|42[0-5]|43[89]|44[78]|45[0-3]|459) # Rango intermedio
-ZZ-MD-(AR|ST|PL|EM|EE|ME|DT)                    # Especialidad
-(0021\d{2}|000001)$                             # Número final
""", re.VERBOSE)

def obtener_nombre_archivo():
    """Obtiene el nombre del modelo activo en Revit.

    Returns:
        str: Nombre del archivo sin la extensión `.rvt`.
    """
    doc = revit.doc
    return doc.Title  # Retorna solo el nombre sin la extensión .rvt

def validar_nombre(nombre_archivo):
    """Valida el nombre del archivo según el patrón permitido.
    Args:
        nombre_archivo (str): Nombre del archivo a validar.
    Returns:
        bool: ``True`` si cumple el patrón, ``False`` en caso contrario.
    """
    return bool(patron.match(nombre_archivo))
