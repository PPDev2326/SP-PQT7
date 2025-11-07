# -*- coding: utf-8 -*-
"""
Helper centralizado para obtener la especialidad del documento actual
Todos los scripts deben usar este módulo para obtener la especialidad
"""

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from DBRepositories.SpecialtiesRepository import SpecialtiesRepository

def get_current_specialty(document):
    """
    Obtiene el objeto Specialty del documento actual desde el parámetro de Información de Proyecto.
    Esta es la función centralizada que todos los scripts deben usar.
    
    Args:
        document: Documento activo de Revit
        
    Returns:
        Specialty: Objeto con toda la información de la especialidad, o None si no se encuentra
        
    Uso:
        from Helper._Specialty import get_current_specialty
        
        specialty_obj = get_current_specialty(doc)
        if specialty_obj:
            specialty_name = specialty_obj.name  # "ARQUITECTURA", "INSTALACIONES SANITARIAS", etc.
            specialty_code = specialty_obj.suffix  # "AR", "PL", etc.
            sp_accessibility = specialty_obj.accessibility_performance
            sp_code = specialty_obj.code_perfomance
            sp_sustainability = specialty_obj.sustainability
            sp_feature = specialty_obj.feature
    """
    if not document:
        return None
    
    try:
        # Obtener información del proyecto
        fec = FilteredElementCollector(document)
        project_info = fec.OfCategory(BuiltInCategory.OST_ProjectInformation).FirstElement()
        
        if not project_info:
            return None
        
        # Buscar el parámetro S&P_ESPECIALIDAD
        param = project_info.LookupParameter("S&P_ESPECIALIDAD")
        
        if not param or not param.HasValue:
            return None
        
        # Obtener el valor del parámetro
        specialty_value = param.AsString()
        
        if not specialty_value or specialty_value.strip() == "":
            return None
        
        # Usar el repositorio para obtener el objeto completo
        repo = SpecialtiesRepository()
        specialty_obj = repo.get_specialty_by_name(specialty_value.strip())
        
        return specialty_obj
        
    except Exception as e:
        print("Error obteniendo especialidad desde Información de Proyecto: {}".format(str(e)))
        return None


def get_specialty_name(document):
    """
    Obtiene solo el NOMBRE de la especialidad (forma rápida).
    
    Args:
        document: Documento activo de Revit
        
    Returns:
        str: Nombre de la especialidad o None
        
    Ejemplo: "ARQUITECTURA", "INSTALACIONES SANITARIAS", etc.
    """
    specialty_obj = get_current_specialty(document)
    return specialty_obj.name if specialty_obj else None


def get_specialty_suffix(document):
    """
    Obtiene solo el SUFIJO de la especialidad (forma rápida).
    
    Args:
        document: Documento activo de Revit
        
    Returns:
        str: Sufijo de la especialidad o None
        
    Ejemplo: "AR", "PL", "EE", etc.
    """
    specialty_obj = get_current_specialty(document)
    return specialty_obj.suffix if specialty_obj else None