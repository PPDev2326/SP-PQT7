# -*- coding: utf-8 -*-
from Models._Specialty import Specialty
from DBRepositories.ISpecialtyRepository import ISpecialtyRepository
from pyrevit import revit

class SpecialtiesRepository(ISpecialtyRepository):
    """
    Repositorio concreto para el manejo de especialidades
    Implementa ISpecialtyRepository y proporciona acceso a las especialidades predefinidas
    """
    
    def __init__(self):
        """
        Inicializa el repositorio con las especialidades predefinidas
        """
        self._specialty_repository = self._init_specialties()
    
    def get_specialty_by_document(self, document):
        """
        Obtiene el objeto de la especialidad desde el documento de Revit.
        
        Args:
            document (Document): Objeto Document de Revit
            
        Returns:
            Specialty: El objeto Specialty correspondiente, o None si no se encuentra
        """
        if not document:
            return None
        
        # Extraer el sufijo de especialidad del nombre del documento
        suffix = self._extract_specialty_suffix(document)
        
        if not suffix:
            return None
        
        return self._specialty_repository.get(suffix, None)
    
    def _extract_specialty_suffix(self, document):
        """
        Extrae el sufijo de especialidad del nombre del nombre del documento de Revit
        
        Args:
            document (Document): documento del tipo Document de Revit
            
        Returns:
            str: Sufijo de especialidad en mayúsculas, o None si no se encuentra
        """
        if not document:
            return None
        
        try:
            title = document.Title
            
            if not title or title.strip() == "":
                return None
            
            # Dividir por guiones
            parts = title.split("-")
            
            # Verificar que tenga suficientes partes para extraer el sufijo
            if len(parts) > 3:
                # Obtener el penúltimo elemento (sufijo de especialidad)
                suffix = parts[-2].strip()
                
                # Validar que sea un sufijo válido (2 letras)
                if len(suffix) == 2 and suffix.isalpha():
                    return suffix.upper()
            
            return None
            
        except (AttributeError, IndexError) as e:
            # Manejo de errores si document.Title no existe o hay problemas con el split
            return None
    
    def get_specialty_by_suffix(self, suffix):
        """
        Obtiene el objeto de la especialidad por su sufijo
        
        Args:
            suffix (str): El sufijo de la especialidad (por ejemplo, "AR" o "EE")
            
        Returns:
            Specialty: El objeto Specialty correspondiente al sufijo, o None si no se encuentra
        """
        if not suffix or suffix.strip() == "":
            return None
        
        suffix_upper = suffix.upper()
        return self._specialty_repository.get(suffix_upper, None)
    
    def get_specialty_by_name(self, name):
        """
        Obtiene una especialidad por su nombre
        
        Args:
            name (str): Nombre de la especialidad
            
        Returns:
            Specialty: Objeto especialidad o None si no se encuentra
        """
        if not name or name.strip() == "":
            return None
        
        name_lower = name.lower()
        for specialty in self._specialty_repository.values():
            if specialty.name.lower() == name_lower:
                return specialty
        
        return None
    
    def get_all_specialties(self):
        """
        Obtiene todas las especialidades disponibles
        
        Returns:
            list: Lista con todas las especialidades
        """
        return list(self._specialty_repository.values())
    
    def specialty_exists(self, suffix):
        """
        Verifica si existe una especialidad con el sufijo dado
        
        Args:
            suffix (str): Sufijo a verificar
            
        Returns:
            bool: True si existe la especialidad, False en caso contrario
        """
        if not suffix or suffix.strip() == "":
            return False
        
        return suffix.upper() in self._specialty_repository
    
    def _init_specialties(self):
        """
        Inicializa el diccionario de especialidades con los valores predefinidos
        
        Returns:
            dict: Diccionario con las especialidades indexadas por sufijo
        """
        specialties = {
            "AR": Specialty("ARQUITECTURA", "AR", "Durabilidad a largo plazo", "Reglamento Nacional de Edificaciones (RNE)", "Resistencia a largo plazo", "Resistente al desgaste"),
            "ST": Specialty("ESTRUCTURAS", "ST"),
            "PL": Specialty("INSTALACIONES SANITARIAS", "PL", "Resistencia al agua", "Norma Técnica de Edificación E.040", "Bajo consumo de agua", "Resistente al agua"),
            "EE": Specialty("INSTALACIONES ELECTRICAS", "EE", "Eficiencia energética", "Código Nacional de Electricidad (CNE)", "Eficiencia energetica", "Ahorro de energía"),
            "DT": Specialty("COMUNICACIONES", "DT", "Durabilidad a largo plazo", "Norma Técnica de Edificación E.090", "Resistencia a largo plazo", "Resistente al fuego"),
            "ME": Specialty("INSTALACIONES MECANICAS", "ME", "Eficiencia energética", "Reglamento Nacional de Edificaciones (RNE)", "Eficiencia energetica", "Durabilidad")
        }
        
        return specialties