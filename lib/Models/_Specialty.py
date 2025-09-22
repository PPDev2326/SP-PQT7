# -*- coding: utf-8 -*-
"""
Modulo de especialidades para pyRevit
Contiene las clases necesarias para manejar especialidades de construccion
"""

class Specialty(object):
    """
    Clase que representa una especialidad de construccion
    """
    
    def __init__(self, name=None, suffix=None):
        """
        Inicializa una nueva instancia de Specialty
        
        Args:
            name (str): Nombre de la especialidad
            suffix (str): Sufijo de la especialidad
        """
        self.name = name
        self.suffix = suffix
    
    def __str__(self):
        """
        Representacion en cadena de la especialidad
        
        Returns:
            str: Cadena en formato "SUFIJO - NOMBRE"
        """
        return "{0} - {1}".format(self.suffix, self.name)
    
    def __repr__(self):
        """
        Representacion para debugging
        
        Returns:
            str: Representacion de la instancia
        """
        return "Specialty(name='{0}', suffix='{1}')".format(self.name, self.suffix)