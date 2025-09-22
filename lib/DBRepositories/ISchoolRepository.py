# -*- coding: utf-8 -*-
# repositories/interfaces.py
from abc import ABC, abstractmethod

class ISchoolDBRepository(ABC):
    """
    Define los métodos que deben ser implementados por los repositorios.
    """
    
    @abstractmethod
    def codigo_colegio(self, codigo):
        """
        Obtiene un colegio por su código.
        
        Args:
            codigo (str): Código del colegio
            
        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        pass
    
    @abstractmethod
    def nombre_colegio(self, nombre):
        """
        Obtiene un colegio por su nombre.
        
        Args:
            nombre (str): Nombre del colegio
            
        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        pass
    
    @abstractmethod
    def propiedades_colegio(self):
        """
        Obtiene todos los colegios disponibles.
        
        Returns:
            list: Lista de todos los colegios
        """
        pass