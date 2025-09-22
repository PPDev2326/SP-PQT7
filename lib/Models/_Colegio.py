# -*- coding: utf-8 -*-

class Colegio:
    """
    Clase que representa un colegio con sus propiedades básicas.
    """
    
    def __init__(self, codigo, nombre, replacement_date, om, doc_ref, inst_code, plot_code):
        self.code = codigo
        self.name = nombre
        self.replacement_date = replacement_date
        self.operation_and_maintenance = om
        self.document_reference = doc_ref
        self.institution_code = inst_code
        self.plot_code = plot_code
    
    def __str__(self):
        return "{} - {}".format(self.code, self.name)
    
    def __repr__(self):
        return "Colegio(codigo='{}', nombre='{}')".format(self.code, self.name)