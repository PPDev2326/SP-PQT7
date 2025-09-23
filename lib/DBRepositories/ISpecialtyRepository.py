# -*- coding: utf-8 -*-

class ISpecialtyRepository(object):
    """
    Interfaz abstracta para el repositorio de especialidades
    Define los metodos que debe implementar cualquier repositorio de especialidades
    """
    
    def get_specialty_by_suffix(self, suffix):
        """
        Obtiene una especialidad por su sufijo
        
        Args:
            suffix (str): Sufijo de la especialidad
            
        Returns:
            Specialty: Objeto especialidad o None si no se encuentra
        """
        raise NotImplementedError("Subclasses must implement this method")
    
    def get_specialty_by_name(self, name):
        """
        Obtiene una especialidad por su nombre
        
        Args:
            name (str): Nombre de la especialidad
            
        Returns:
            Specialty: Objeto especialidad o None si no se encuentra
        """
        raise NotImplementedError("Subclasses must implement this method")