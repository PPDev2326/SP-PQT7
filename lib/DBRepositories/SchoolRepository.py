# -*- coding: utf-8 -*-
# repositories/colegios_repository.py
from Autodesk.Revit.DB import Document
from Models._Colegio import Colegio, Location, Classification, COBie
from DBRepositories.ISchoolRepository import ISchoolDBRepository

class ColegiosRepository(ISchoolDBRepository):
    """
    Clase responsable de cargar y entregar datos a los servicios.
    """
    
    def __init__(self):
        # tipo: () -> None
        self._colegios = self._inicializar_colegios()
    
    def codigo_colegio(self, document):
        """
        Obtiene un objeto de tipo _Colegio mediante el documento del modelo.
        
        :param document: Tipo Document del modelo revit
        :type document: Document
        :return: Retorna un tipo Colegio del documento buscado
        :rtype: Colegio
        """
        if not document:
            return None
        
        title = document.Title
        parts = title.split("-")
        codigo= parts[0].strip()
        
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
                plot_code="P15054574",
                created_by="jtiburcio@syp.com.pe",
                warranty_desc="200045-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2779", code_cui="2513475", code_local="430284", province="Morropón", district="Chulucanas", populated_center="Chulucanas"),
                classification=Classification(facility_description="Secondary schools", facility_number="Co_25_10_77"),
                cobie=COBie(facility_description="Institución Educativa Secundaria",
                            project_description="Rehabilitación del local escolar maría auxiliadora con código local n°430284, distrito de Chulucanas, provincia de Morropón, departamento de Piura.", site_description="Calle Mz Z, Lote Nº10")
            ),
            "200046": Colegio(
                codigo="200046",
                nombre="I.E 14654 - SALITRAL",
                replacement_date="2024-06-18",
                om="200046-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="02-ARCH-P7-2282-IE 14654-432184",
                inst_code="200008",
                plot_code="P15158777",
                created_by="jtiburcio@syp.com.pe",
                warranty_desc="200046-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2282", code_cui="2513474", code_local="432184", province="Morropón", district="Salitral", populated_center="Salitral"),
                classification=Classification(facility_description="Primary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa Secundaria",
                            project_description="Rehabilitación del local escolar nº 14654 con código local nº432184.distrito de Salitral, provincia de Morropón, departamento de Piura.", site_description="Calle 29 DE MAYO 401")
            ),
            "200047": Colegio(
                codigo="200047",
                nombre="I.E RICARDO PALMA",
                replacement_date="2024-02-26",
                om="200047-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="03-ARCH-P7-2775-IE  RICARDO PALMA-434927",
                inst_code="200009",
                plot_code="11206389",
                created_by="ecabrera@syp.com.pe",
                warranty_desc="200047-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2775", code_cui="2428627", code_local="434927", province="Paita", district="Vichayal", populated_center="Miramar"),
                classification=Classification(facility_description="Secondary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa Secundaria",
                            project_description="Rehabilitación del local escolar Ricardo palma con código local nº 434927, distrito de Vichayal, provincia de Paita, departamento de Piura.", site_description="Avenida Buenos Aires 115")
            ),
            "200049": Colegio(
                codigo="200049",
                nombre="I.E 14037 SANTIAGO A. REQUENA CASTRO",
                replacement_date="2023-07-21",
                om="200049-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="04-ARCH-P7-2358-IE 14037-413350",
                inst_code="200001",
                plot_code="11103575",
                created_by="fgomez@syp.com.pe",
                warranty_desc="200049-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2358", code_cui="2428716", code_local="413350", province="Piura", district="Catacaos", populated_center="Catacaos"),
                classification=Classification(facility_description="Educational complexe", facility_number="Co_25_10"),
                cobie=COBie(facility_description="Institución Educativa Inicial -Primaria",
                            project_description="Rehabilitación del local escolar nº 14037 Santiago a. Requena Castro con código local nº 413350.distrito de Catacaos, provincia de Piura, departamento de Piura.", site_description="Av.Cayetano Heredia")
            ),
            "200053": Colegio(
                codigo="200053",
                nombre="I.E VIRGEN DEL CARMEN",
                replacement_date="2024-04-23",
                om="200053-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="05-ARCH-P7-2806-IE VIRGEN DEL CARMEN-413307",
                inst_code="200001",
                plot_code="P15177460",
                created_by="emendoza@syp.com.pe",
                warranty_desc="200053-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2806", code_cui="2428716", code_local="413307", province="Piura", district="Catacaos", populated_center="Catacaos"),
                classification=Classification(facility_description="Educational complexe", facility_number="Co_25_10"),
                cobie=COBie(facility_description="Institución Educativa Inicial -Primaria",
                            project_description="Rehabilitación del local escolar virgen del Carmen con código local nº 413307, distrito de Catacaos, provincia de Piura, departamento de Piura.", site_description="Calle San Francisco S/N")
            ),
            "200057": Colegio(
                codigo="200057",
                nombre="I.E 14065 - LA UNION YAPATO",
                replacement_date="2023-12-28",
                om="200057-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="06-ARCH-P7-1680-IE 14065-414707",
                inst_code="200003",
                plot_code="11055485",
                created_by="emendoza@syp.com.pe",
                warranty_desc="200057-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="1680", code_cui="2428581", code_local="414707", province="Piura", district="Unión", populated_center="Yapato"),
                classification=Classification(facility_description="Primary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa Primaria",
                            project_description="Rehabilitación del local escolar nº 14065 con código loca nº414707. Distrito de la unión, provincia de Piura, departamento de Piura.", site_description="Av.Juan Velasco")
            ),
            "200058": Colegio(
                codigo="200058",
                nombre="I.E Nº 14062 NUESTRA SEÑORA DE LAS MERCEDES",
                replacement_date="2024-12-27",
                om="200058-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="07-ARCH-P7-2261-IE 14062-770240",
                inst_code="200003",
                plot_code="Ficha 039876 Asiento 17827",
                created_by="mespinoza@syp.com.pe",
                warranty_desc="200058-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2261", code_cui="2428588", code_local="770240", province="Piura", district="La Unión", populated_center="Tablazo Norte"),
                classification=Classification(facility_description="Primary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa- Primaria",
                            project_description="Rehabilitación del local escolar nº 14062 Nuestra Señora de las Mercedes con código local nº 770240, distrito de La Unión, provincia de Piura, departamento de Piura.", site_description="Ca. San Martín Nº 401")
            ),
            "200060": Colegio(
                codigo="200060",
                nombre="I.E SAN CRISTO",
                replacement_date="2025-08-22",
                om="200060-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="08-ARCH-P7-1784-IE SAN CRISTO-440688",
                inst_code="200004",
                plot_code="00022164",
                created_by="nmadrid@syp.com.pe",
                warranty_desc="200060-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="1784", code_cui="2513470", code_local="440688", province="Sechura", district="Cristo nos valga", populated_center="San Cristo"),
                classification=Classification(facility_description="Educational complexes", facility_number="Co_25_10"),
                cobie=COBie(facility_description="Institución Educativa- Inicial-Primaria -Secundaria",
                            project_description="Rehabilitación del local escolar “San Cristo” con código local nº 440688, distrito de cristo nos valga, provincia de Sechura, departamento de Piura.", site_description="Calle Libertad-Sanchez Cerro S/N")
            ),
            "200061": Colegio(
                codigo="200061",
                nombre="I.E 14010 MIGUEL F. CERRO",
                replacement_date="2024-06-25",
                om="200061-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="09-ARCH-P7-1675-IE 14010-440773",
                inst_code="200004",
                plot_code="03009653, 03008241",
                created_by="mespinoza@syp.com.pe",
                warranty_desc="200061-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="1675", code_cui="2428585", code_local="440773", province="San Cristo", district="Vice", populated_center="Vice"),
                classification=Classification(facility_description="Primary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa Primaria",
                            project_description="Recuperación del local escolar nº 14010 Miguel F. Cerro-Vice con código local nº 440773, distrito de vice, provincia de Sechura, departamento de Piura.", site_description="Avenida Miguel F.Cerro S/N")
            ),
            "200063": Colegio(
                codigo="200063",
                nombre="I.E 14858 - MIGUEL CHECA",
                replacement_date="2024-07-22",
                om="200063-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="10-ARCH-P7-1706-IE 14858-437775",
                inst_code="200010",
                plot_code="P11061164",
                created_by="ecabrera@syp.com.pe",
                warranty_desc="200063-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="1706", code_cui="2428584", code_local="437775", province="Sullana", district="Miguel Checa", populated_center="Jibito"),
                classification=Classification(facility_description="Primary schools", facility_number="Co_25_10_66"),
                cobie=COBie(facility_description="Institución Educativa Primaria",
                            project_description="Rehabilitación del local escolar nº 14858 con código local nº 437775, distrito de miguel checa, provincia de Sullana, departamento de Piura.", site_description="Santa Julia")
            ),
            "200070": Colegio(
                codigo="200070",
                nombre="I.E IGNACIO MERINO",
                replacement_date="2024-12-05",
                om="200070-CSSP001-000-XX-OM-ZZ-000001",
                doc_ref="11-ARCH-P7-2766-IE IGNACIO MERINO-438614",
                inst_code="200011",
                plot_code="11023138",
                created_by="fgomez@syp.com.pe",
                warranty_desc="200070-CSSP001-000-XX-OM-ZZ-000001",
                location=Location(code_arcc="2766", code_cui="2428635", code_local="438614", province="Talara", district="Pariñas", populated_center="Talara"),
                classification=Classification(facility_description="Secondary schools", facility_number="Co_25_10_77"),
                cobie=COBie(facility_description="Institución Educativa Secundaria",
                            project_description="Rehabilitación del local escolar Ignacio merino con código local n°438614, distrito de Pariñas, provincia de talara, departamento de Piura.", site_description="Av. Ignacio Merino")
            ),
        }