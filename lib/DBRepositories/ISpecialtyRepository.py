# -*- coding: utf-8 -*-
from abc import ABC, abstractmethod

class ISpecialtyRepository(ABC):
    """
    Interfaz abstracta para el repositorio de especialidades
    Define los metodos que debe implementar cualquier repositorio de especialidades
    """
    
    @abstractmethod
    def get_specialty_by_suffix(self, suffix):
        """
        Obtiene una especialidad por su sufijo
        
        Args:
            suffix (str): Sufijo de la especialidad
            
        Returns:
            Specialty: Objeto especialidad o None si no se encuentra
        """
        pass
    
    @abstractmethod
    def get_specialty_by_name(self, name):
        """
        Obtiene una especialidad por su nombre
        
        Args:
            name (str): Nombre de la especialidad
            
        Returns:
            Specialty: Objeto especialidad o None si no se encuentra
        """
        pass