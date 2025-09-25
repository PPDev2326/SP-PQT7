# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms, revit
from pyrevit.interop import xl

class Excel:
    def read_excel(self, hoja, encabezados=False):
        """
        Lee un archivo Excel con debug completo.
        """
        print("=== INICIANDO READ_EXCEL ===")
        print("Hoja solicitada: '{}'".format(hoja))
        print("Encabezados: {}".format(encabezados))
        
        ruta = forms.pick_excel_file()
        print("Ruta seleccionada: {}".format(ruta))
        
        if not ruta:
            print("ERROR: No se seleccionó archivo")
            forms.alert("No se seleccionó ningún archivo.", exitscript=True)
            return []
        
        try:
            print("Intentando cargar Excel...")
            datos = xl.load(ruta, headers=encabezados)
            print("Carga exitosa. Tipo de datos: {}".format(type(datos)))
            print("Datos cargados: {}".format(datos))
            
            if not datos:
                print("ERROR: datos está vacío")
                forms.alert("El archivo Excel no se pudo cargar o está vacío.")
                return []
                
            hojas_disponibles = list(datos.keys())
            print("Hojas disponibles: {}".format(hojas_disponibles))
            print("Número de hojas: {}".format(len(hojas_disponibles)))
            
            # Mostrar contenido de cada hoja
            for hoja_name in hojas_disponibles:
                hoja_data = datos[hoja_name]
                print("Hoja '{}': {}".format(hoja_name, type(hoja_data)))
                if isinstance(hoja_data, dict):
                    print("  Claves: {}".format(list(hoja_data.keys())))
                    if 'rows' in hoja_data:
                        rows = hoja_data['rows']
                        print("  Filas encontradas: {}".format(len(rows) if rows else 0))
                        if rows and len(rows) > 0:
                            print("  Primera fila: {}".format(rows[0]))
            
            forms.alert("Hojas encontradas:\n{}".format('\n'.join(hojas_disponibles)))
            
            # Buscar la hoja
            hoja_encontrada = None
            print("Buscando hoja: '{}'".format(hoja))
            
            # Coincidencia exacta
            if hoja in hojas_disponibles:
                hoja_encontrada = hoja
                print("Coincidencia exacta encontrada")
            else:
                print("No hay coincidencia exacta, buscando similar...")
                hoja_lower = hoja.lower()
                for h in hojas_disponibles:
                    print("  Comparando '{}' con '{}'".format(hoja_lower, h.lower()))
                    if hoja_lower in h.lower() or h.lower() in hoja_lower:
                        hoja_encontrada = h
                        print("  ¡Coincidencia parcial encontrada!: '{}'".format(h))
                        break
            
            if not hoja_encontrada:
                print("ERROR: No se encontró la hoja")
                forms.alert("No se encontró la hoja '{}'. Hojas disponibles:\n{}".format(
                    hoja, '\n'.join(hojas_disponibles)), exitscript=True)
                return []
            
            print("Usando hoja: '{}'".format(hoja_encontrada))
            hoja_data = datos.get(hoja_encontrada, {})
            print("Datos de la hoja: {}".format(hoja_data))
            
            if not isinstance(hoja_data, dict):
                print("ERROR: Los datos de la hoja no son un diccionario")
                return []
                
            if 'rows' not in hoja_data:
                print("ERROR: No se encontró la clave 'rows' en los datos")
                print("Claves disponibles: {}".format(list(hoja_data.keys())))
                return []
            
            filas = hoja_data['rows']
            print("Filas obtenidas: {}".format(type(filas)))
            print("Número de filas: {}".format(len(filas) if filas else 0))
            
            if not filas:
                forms.alert("La hoja '{}' no contiene datos o está vacía.".format(hoja_encontrada))
                return []
            
            print("=== ÉXITO: Retornando {} filas ===".format(len(filas)))
            return filas
            
        except Exception as e:
            print("EXCEPCIÓN: {}".format(str(e)))
            forms.alert("Error al cargar el Excel: {}".format(str(e)))
            return []
    
    def get_headers(rows, start_row = 0):
        """
        Busca los encabezados en las filas y retorna los índices de las columnas.
        
        Args:
            row (list): Lista de filas del Excel
            required_column (dict): Diccionario con nombres lógicos y posibles nombres de columnas
            start_row (int): Fila desde donde empezar a buscar encabezados (default: 0)
            max_row_send (int): Máximo número de filas a revisar (default: 2)
        
        Returns:
            dict: Diccionario con los índices de las columnas encontradas con valor.
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
    
    def headers_required(headers, columns_name):
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
    
    def get_data_by_headers_required(rows_data, columns_required, start_data=1):
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


# required_column = {
#                 "COBie.Type.Manufacturer": None,
#                 "COBie.Type.ModelNumber": None,
#                 "COBie.Type.WarrantyDurationParts": None,
#                 "COBie.Type.WarrantyDurationLabor": None,
#                 "COBie.Type.ReplacementCost": None,               # Varia de acuerdo al excel
#                 "COBie.Type.ExpectedLife": None,
#                 "COBie.Type.NominalLength": None,
#                 "COBie.Type.NominalWidth": None,
#                 "COBie.Type.NominalHeight": None,
#                 "COBie.Type.Color": None,
#                 "COBie.Type.Finish": None,
#                 "COBie.Type.Constituents": None
#             }