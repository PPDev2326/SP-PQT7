# -*- coding: utf-8 -*-
class ISchoolDBRepository(object):
    """
    Define los métodos que deben ser implementados por los repositorios.
    """

    def codigo_colegio(self, codigo):
        """
        Obtiene un colegio por su código.

        Args:
            codigo (str): Código del colegio

        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        raise NotImplementedError("Debe implementar 'codigo_colegio'")

    def nombre_colegio(self, nombre):
        """
        Obtiene un colegio por su nombre.

        Args:
            nombre (str): Nombre del colegio

        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        raise NotImplementedError("Debe implementar 'nombre_colegio'")

    def propiedades_colegio(self):
        """
        Obtiene todos los colegios disponibles.

        Returns:
            list: Lista de todos los colegios
        """
        raise NotImplementedError("Debe implementar 'propiedades_colegio'")