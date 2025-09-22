# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import BuiltInParameter, Element

def getParameter(element, name):
    """Obtiene un parametro compartido si no es de solo lectura.
    type: (Element, str) -> Parameter"""
    if element is None or name is None:
        return None
    param = element.LookupParameter(name)
    if param and not param.IsReadOnly:
        return param
    return None

def SetParameter(parameter, new_value):
    """Setea un parámetro si no es de solo lectura.
    :param parameter: El parámetro de Revit a modificar.
    :param new_value: El nuevo valor como string (se adapta según tipo de parámetro).
    :return: True si se seteó correctamente, False si fue de solo lectura o inválido."""
    if parameter and not parameter.IsReadOnly:
        try:
            parameter.Set(new_value)
            return True
        
        except Exception as e:
            print("Error al setear el parametro: ", e)
    return False

def GetParameterAPI(element, built_param, es_editable = False):
    """
    Obtiene un parámetro usando BuiltInParameter.
    
    :param element: Elemento de Revit (instancia, tipo, etc.).
    :param built_param: Constante BuiltInParameter (por ejemplo, BuiltInParameter.ROOM_NAME).
    :param solo_si_es_editable: Si True, solo devuelve el parámetro si se puede editar.
    :return: Parámetro si existe (y editable si se pide), o None.
    """
    if element is None or built_param is None:
        return None
    try:
        param_api = element.get_Parameter(built_param)
        if param_api:
            if es_editable and param_api.IsReadOnly:
                return None
            return param_api
    except Exception as e:
        print("Error al obtener el parámetro por BuiltInParameter:", e)
    return None