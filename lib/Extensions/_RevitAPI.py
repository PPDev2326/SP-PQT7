# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import BuiltInParameter, Element, StorageType

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

def GetParameterAPI(element, built_param):
    """
    Obtiene un parámetro usando BuiltInParameter.
    
    :param element: Elemento de Revit (instancia, tipo, etc.).
    :param built_param: Constante BuiltInParameter (por ejemplo, BuiltInParameter.ROOM_NAME).
    :return: Parámetro si existe (y editable si se pide), o None.
    """
    if element is None or built_param is None:
        return None
    try:
        return element.get_Parameter(built_param)
    
    except Exception as e:
        print("Error al obtener el parámetro por BuiltInParameter:", e)
    return None

def get_param_value(param, default=None):
    """Devuelve el valor del parámetro según su tipo, o default si no existe."""
    if not param:
        return default
    try:
        if param.StorageType == StorageType.String:
            return (param.AsString() or "").strip()
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.ElementId:
            return param.AsElementId()
        return default
    except:
        return default
