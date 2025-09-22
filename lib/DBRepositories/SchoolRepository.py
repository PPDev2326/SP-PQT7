# -*- coding: utf-8 -*-
# repositories/colegios_repository.py
from Models._Colegio import Colegio
from DBRepositories.ISchoolRepository import ISchoolDBRepository

class ColegiosRepository(ISchoolDBRepository):
    """
    Clase responsable de cargar y entregar datos a los servicios.
    """
    
    def __init__(self):
        # tipo: () -> None
        self._colegios = self._inicializar_colegios()
    
    def codigo_colegio(self, codigo):
        """
        Obtiene un colegio por su código.
        
        Args:
            codigo (str): Código del colegio
            
        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        
        return self._colegios.get(codigo)
    
    def nombre_colegio(self, nombre):
        """
        Obtiene un colegio por su nombre.
        
        Args:
            nombre (str): Nombre del colegio
            
        Returns:
            Colegio: El colegio si existe, None en caso contrario
        """
        for colegio in self._colegios.values():
            if colegio.name.lower() == nombre.lower():
                return colegio
        return None
    
    def propiedades_colegio(self):
        """
        Obtiene todos los colegios disponibles.
        
        Returns:
            list: Lista de todos los colegios
        """
        return list(self._colegios.values())
    
    def _inicializar_colegios(self):
        """
        Inicializa el diccionario de colegios con todos los datos.
        
        Returns:
            dict: Diccionario con código como clave y Colegio como valor
        """
        return {
            "200045": Colegio(
                codigo="200045",
                nombre="I.E MARIA AUXILIADORA",
                replacement_date="2023-12-21",
                om="200045-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="01-ARCH-P7-2779-IE  MARIA AUXILIADORA-430284",
                inst_code="200007",
                plot_code="P15054574"
            ),
            "200046": Colegio(
                codigo="200046",
                nombre="I.E 14654 - SALITRAL",
                replacement_date="2024-06-18",
                om="200046-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="02-ARCH-P7-2282-IE 14654-432184",
                inst_code="200008",
                plot_code="P15158777"
            ),
            "200047": Colegio(
                codigo="200047",
                nombre="I.E RICARDO PALMA",
                replacement_date="2024-02-26",
                om="200047-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="03-ARCH-P7-2775-IE  RICARDO PALMA-434927",
                inst_code="200009",
                plot_code="11206389"
            ),
            "200049": Colegio(
                codigo="200049",
                nombre="I.E 14037 SANTIAGO A. REQUENA CASTRO",
                replacement_date="2023-07-21",
                om="200049-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="04-ARCH-P7-2358-IE 14037-413350",
                inst_code="200001",
                plot_code="11103575"
            ),
            "200053": Colegio(
                codigo="200053",
                nombre="I.E VIRGEN DEL CARMEN",
                replacement_date="2024-04-23",
                om="200053-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="05-ARCH-P7-2806-IE VIRGEN DEL CARMEN-413307",
                inst_code="200001",
                plot_code="P15177460"
            ),
            "200057": Colegio(
                codigo="200057",
                nombre="I.E 14065 - LA UNION YAPATO",
                replacement_date="2023-12-28",
                om="200057-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="06-ARCH-P7-1680-IE 14065-414707",
                inst_code="200003",
                plot_code="11055485"
            ),
            "200058": Colegio(
                codigo="200058",
                nombre="I.E Nº 14062 NUESTRA SEÑORA DE LAS MERCEDES",
                replacement_date="2024-12-27",
                om="200058-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="07-ARCH-P7-2261-IE 14062-770240",
                inst_code="200003",
                plot_code="Ficha 039876 Asiento 17827"
            ),
            "200060": Colegio(
                codigo="200060",
                nombre="I.E SAN CRISTO",
                replacement_date="2025-08-22",
                om="200060-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="08-ARCH-P7-1784-IE SAN CRISTO-440688",
                inst_code="200004",
                plot_code="00022164"
            ),
            "200061": Colegio(
                codigo="200061",
                nombre="I.E 14010 MIGUEL F. CERRO",
                replacement_date="2024-06-25",
                om="200061-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="09-ARCH-P7-1675-IE 14010-440773",
                inst_code="200004",
                plot_code="03009653, 03008241"
            ),
            "200063": Colegio(
                codigo="200063",
                nombre="I.E 14858 - MIGUEL CHECA",
                replacement_date="2024-07-22",
                om="200063-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="10-ARCH-P7-1706-IE 14858-437775",
                inst_code="200010",
                plot_code="P11061164"
            ),
            "200070": Colegio(
                codigo="200070",
                nombre="I.E IGNACIO MERINO",
                replacement_date="2024-12-05",
                om="200070-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="11-ARCH-P7-2766-IE IGNACIO MERINO-438614",
                inst_code="200011",
                plot_code="11023138"
            ),
        }