# -*- coding: utf-8 -*-

class Location:
    """
    Representa la ubicación de un colegio.

    :param code_arcc: Código ARCC del colegio.
    :type code_arcc: str, optional
    :param code_cui: Código CUI del colegio.
    :type code_cui: str, optional
    :param code_local: Código local del colegio.
    :type code_local: str, optional
    :param province: Provincia donde se encuentra el colegio.
    :type province: str, optional
    :param district: Distrito donde se encuentra el colegio.
    :type district: str, optional
    :param populated_center: Centro poblado del colegio.
    :type populated_center: str, optional
    """
    def __init__(self, code_arcc=None, code_cui=None, code_local=None,
                province=None, district=None, populated_center=None):
        self.code_arcc = code_arcc
        self.code_cui = code_cui
        self.code_local = code_local
        self.province = province
        self.district = district
        self.populated_center = populated_center


class Classification:
    """
    Representa la clasificación del colegio.

    :param facility_description: Descripción de la instalación.
    :type facility_description: str, optional
    :param facility_number: Número de la instalación.
    :type facility_number: str, optional
    """
    def __init__(self, facility_description=None, facility_number=None):
        self.facility_description = facility_description
        self.facility_number = facility_number


class COBie:
    """
    Representa la información COBie del colegio.
    
    :param facility_description: Descripción de la instalación.
    :type facility_description: str, optional
    :param project_description: Descripción del proyecto.
    :type project_description: str, optional
    :param site_description: Descripción del sitio.
    :type site_description: str, optional
    """
    def __init__(self, facility_description=None,
                project_description=None, site_description=None):
        self.facility_description = facility_description
        self.project_description = project_description
        self.site_description = site_description

class COBieComponent:
    """
    Representa la información COBie del ccomponente.
    
    :param component_warranty: Inicio de la fecha de garantía.
    :type component_warranty: str, optional
    """
    def __init__(self, warranty_start = None):
        self.component_warranty = warranty_start

class Colegio:
    """
    Clase que representa un colegio con sus propiedades básicas.
    """
    
    def __init__(self, codigo, nombre, replacement_date, om, doc_ref, inst_code, plot_code, created_by, warranty_desc,
                location=None, classification=None, cobie=None, warranty_start=None):
        self.code = codigo
        self.name = nombre
        self.replacement_date = replacement_date
        self.operation_and_maintenance = om
        self.document_reference = doc_ref
        self.institution_code = inst_code
        self.plot_code = plot_code
        self.created_by = created_by
        self.warranty_description = warranty_desc
        
        # --- Nuevos parámetros agrupados ---
        self.location = location or Location()
        self.classification = classification or Classification()
        self.cobie = cobie or COBie()
        self.warranty_start_date = warranty_start or COBieComponent()

        
    
    def __str__(self):
        return "{} - {}".format(self.code, self.name)
    
    def __repr__(self):
        return "Colegio(codigo='{}', nombre='{}')".format(self.code, self.name)