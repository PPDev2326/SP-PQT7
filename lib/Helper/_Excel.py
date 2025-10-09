# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms, revit
from pyrevit.interop import xl

class Excel:
    def __init__(self):
        """Inicializa la clase Excel con ruta de archivo None."""
        self.ruta_archivo = None
    
    def read_excel(self, hoja, encabezados=False):
        """
        Lee un archivo Excel. Si ya se seleccionó un archivo anteriormente,
        usa ese mismo archivo para leer otras hojas.

        :param hoja: El nombre de la hoja de cálculo.
        :type hoja: str
        :param encabezados: Si es True, la primera fila se trata como encabezados.
        :type encabezados: bool, optional
        :return: Las filas del Excel.
        :rtype: list
        """
        # Solo pedir el archivo si no se ha seleccionado antes
        if self.ruta_archivo is None:
            self.ruta_archivo = forms.pick_excel_file()
            if not self.ruta_archivo:
                forms.alert("No se seleccionó ningún archivo.", exitscript=True)
                return []
        
        # Usar la ruta guardada
        datos = xl.load(self.ruta_archivo, sheets=str(hoja), headers=encabezados)
        filas = datos.get(hoja, {}).get('rows', [])
        
        return filas
    
    def reset_file(self):
        """Resetea la ruta del archivo para permitir seleccionar uno nuevo."""
        self.ruta_archivo = None
    
    def get_headers(self, rows, start_row = 0):
        """
        Busca los encabezados en las filas y retorna los índices de las columnas.
        
        :param rows: Lista de filas del Excel.
        :type rows: list
        :param start_row: Indice que decidira los encabezados.
        :type start_row: int, optional, default 0
        :return: Diccionario con los índices de las columnas encontradas con valor.
        :rtype: dict
        """
        if len(rows) <= start_row:
            forms.alert("El archivo no tiene suficientes filas para contener encabezados en la fila {}.".format(start_row + 1), exitscript=True)
            return {}

        headers = rows[start_row]
        headers_dict = {}
        
        for idx ,h in enumerate(headers):
            if h not in ("", None):         # => Ignoramos celdas vacias
                headers_dict[idx] = h
        return headers_dict
    
    def headers_required(self, headers, columns_name):
        """
        Filtra los encabezados encontrados y retorna solo los que están en columns_name.
        
        Args:
            headers (dict): Diccionario {indice: nombre_columna}
            columns_name (list): Lista de nombres de columnas requeridas
        
        Returns:
            dict: Diccionario existente {nombre_columna: indice}. Si alguna falta → None
        """
        
        found = {}
        for col in columns_name:
            idx = None
            for i, h in headers.items():
                if h == col:
                    idx = i
                    break
            found[col] = idx
        return found
    
    def get_data_by_headers_required(self, rows_data, columns_required, start_data=1):
        """
        Obtiene los datos del Excel basados en los encabezados requeridos.
        
        Args:
            rows_data (list): Filas del Excel
            columns_required (dict): Diccionario {nombre_columna: índice}
            start_data (int): Fila desde donde empiezan los datos (default: 1)
        
        Returns:
            list: Lista de dicts, cada fila con sus columnas requeridas
        """
        data = []
        for r in rows_data[start_data:]:
            row_dict = {}
            for col_name, idx in columns_required.items():
                if idx is not None and idx < len(r):
                    row_dict[col_name] = r[idx]
                else:
                    row_dict[col_name] = None
            data.append(row_dict)
        return data