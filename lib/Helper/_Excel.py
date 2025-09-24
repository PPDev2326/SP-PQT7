# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms, revit
from pyrevit.interop import xl

class Excel:
    def read_excel(hoja, encabezados=False):
        """
        Lee un archivo Excel.
        
        Args:
            hoja (str): Nombre de la hoja. Si es None, selecciona automáticamente
            encabezados (bool): Si True, trata la primera fila como encabezados
        
        Returns:
            list: Filas del Excel
        """
        ruta = forms.pick_excel_file()
        if not ruta:
            forms.alert("No se seleccionó ningún archivo.", exitscript=True)
            return []
        
        datos = xl.load(ruta, sheets=hoja, headers=encabezados)
        filas = datos.get(hoja, {}).get('rows', [])
        
        return filas
    
    def get_headers(rows, required_column, start_row, max_row_send=2):
        """
        Busca los encabezados en las filas y retorna los índices de las columnas.
        
        Args:
            row (list): Lista de filas del Excel
            required_column (dict): Diccionario con nombres lógicos y posibles nombres de columnas
                Ejemplo: {
                    'item': ['Item', 'Elemento', 'Name'],
                    'specialty': ['Specialty', 'Especialidad', 'Type'], 
                    'cobie': ['COBie Requirement', 'COBie', 'Required']
                }
            start_row (int): Fila desde donde empezar a buscar encabezados (default: 0)
            max_row_send (int): Máximo número de filas a revisar (default: 2)
        
        Returns:
            dict: Diccionario con los índices de las columnas encontradas
                Ejemplo: {'item': 0, 'specialty': 1, 'cobie': 8, 'header_row': 4}
                Si no encuentra alguna columna, su valor será None
        """
        if required_column is None:
            required_column = {
                'item': ['Item', 'Elemento', 'Name', 'Nombre'],
                'specialty': ['Specialty', 'Especialidad', 'Type', 'Tipo'],
                'cobie': ['COBie Requirement', 'COBie', 'Required', 'Requerido']
            }
        if not rows:
            forms.alert("No hay filas para procesar.", exitscript=True)
            return {}
        
        # Buscar la fila de encabezados
        encabezados = None
        header_row_index = None
        
        max_filas = min(len(rows), start_row + max_row_send)
        
        for i in range(start_row, max_filas):
            fila = rows[i]
            if not fila:
                continue
        
        encabezados = rows[4]  # Fila 3 en Excel (índice 4)
        try:
            col_espe = encabezados.index("Specialty")               # Columna B
            col_descri = encabezados.index("COBie Requirement")     # Columna I
            col_item = encabezados.index("Item")                    # Columna A
        except ValueError:
            forms.alert("No se encontraron las columnas necesarias: 'Especialidad', 'COBie', 'NRM 1'.", exitscript=True)
            return []