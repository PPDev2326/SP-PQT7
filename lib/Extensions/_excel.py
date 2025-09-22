# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms
from pyrevit.interop import xl

uniclass_dict = {}
master_dict = {}

class Excel:
    def get_Excel(self, nombre_sheet):
        ruta = forms.pick_excel_file()
        if not ruta:
            forms.alert("No se selecciono archivo excel.", exitscript=True)
            return []
        
        excel_file = xl.load(ruta, sheets=nombre_sheet, headers=True)
        filas = excel_file.get(nombre_sheet, {}).get("rows", [])
        
        encabezados = [h.strip() for h in excel_file[nombre_sheet]["headers"]]
        
        index_desc = encabezados.index("Descripción de elemento")
        index_sistema = encabezados.index("Sistema")
        index_espec = encabezados.index("Especialidad")
        index_sub = encabezados.index("Sub Especialidad")
        
        for fila in filas:
            descripcion = str(fila[index_desc]).strip()
            master_dict[descripcion] = {
                "Sistema" : fila[index_sistema],
                "Especialidad" : fila[index_espec],
                "Sub Especialidad" : fila[index_sub],
            }
        return master_dict
    
    def read_excel(self, nombre_sheet):
        ruta = forms.pick_excel_file()
        if not ruta:
            forms.alert("No se selecciono archivo excel.", exitscript=True)
        
        excel_file = xl.load(ruta, sheets=nombre_sheet, headers=True)
        filas = excel_file.get(nombre_sheet, {}).get("rows", [])
        
        return [excel_file, filas]

def get_excel_uniclass(nombre_sheet):
    ruta = forms.pick_excel_file()
    if not ruta:
        forms.alert("No se seleccionó ningún archivo uniclass.", exitscript=True)
        return []
    
    datos = xl.load(ruta, sheets=nombre_sheet, headers=True)
    filas = datos.get(nombre_sheet, {}).get("rows", [])
    
    encabezados = [h.strip() for h in datos[nombre_sheet]["headers"]]
    
    index_codigo = encabezados.index("Código")
    index_EF_number = encabezados.index("Classification.Uniclass.EF.Number")
    index_EF_desc = encabezados.index("Classification.Uniclass.EF.Description")
    index_Pr_number = encabezados.index("Classification.Uniclass.Pr.Number")
    index_Pr_desc = encabezados.index("Classification.Uniclass.Pr.Description")
    index_Ss_number = encabezados.index("Classification.Uniclass.Ss.Number")
    index_Ss_desc = encabezados.index("Classification.Uniclass.Ss.Description")
    
    for fila in filas:
        codigo = str(fila[index_codigo]).strip()
        uniclass_dict[codigo] = {
            "Classification.Uniclass.EF.Number": fila[index_EF_number],
            "Classification.Uniclass.EF.Description": fila[index_EF_desc],
            "Classification.Uniclass.Pr.Number": fila[index_Pr_number],
            "Classification.Uniclass.Pr.Description": fila[index_Pr_desc],
            "Classification.Uniclass.Ss.Number": fila[index_Ss_number],
            "Classification.Uniclass.Ss.Description": fila[index_Ss_desc]
        }
    return uniclass_dict