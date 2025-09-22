# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import forms, revit
from pyrevit.interop import xl

doc = revit.doc

def ignorar_categorias():
    # Lista de BuiltInCategory de las categorías a excluir
    categorias = [
        BuiltInCategory.OST_Conduit,       
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_FlexPipeCurves,
        BuiltInCategory.OST_StructuralFramingOpening,
        BuiltInCategory.OST_ColumnOpening
    ]
    # Obtener los nombres localizados de esas categorías
    nombres_excluir = [
        doc.Settings.Categories.get_Item(bic).Name
        for bic in categorias
    ]
    
    return nombres_excluir



def leer_excel_filtrado():
    ruta = forms.pick_excel_file()
    if not ruta:
        forms.alert("No se seleccionó ningún archivo.", exitscript=True)
        return []
    
    datos = xl.load(ruta, sheets="Datos", headers=False)
    filas = datos.get('Datos', {}).get('rows', [])
    
    if len(filas) < 9:
        forms.alert("El archivo no tiene suficientes filas para contener encabezados en la fila 9.", exitscript=True)
        return []
    
    encabezados = filas[8]  # Fila 9 en Excel (índice 8)
    try:
        col_espe = encabezados.index("Especialidad")            # Columna B
        col_descri = encabezados.index("COBie")                 # Columna I
        col_col1 = encabezados.index("NRM 1")                   # Columna A
    except ValueError:
        forms.alert("No se encontraron las columnas necesarias: 'Especialidad', 'COBie', 'NRM 1'.", exitscript=True)
        return []
    
    resultados_col3 = []
    
    for fila in filas[9:]:  # Desde la fila 10 (índice 9)
        if len(fila) <= max(col_espe, col_descri, col_col1):
            continue
        
        valor_espe = fila[col_espe]
        valor_descri = fila[col_descri]
        valor_col1 = fila[col_col1]
        
        # Convertir valores a cadena antes de aplicar comparaciones
        valor_espe_str = str(valor_espe).strip()
        valor_descri_str = str(valor_descri).strip()
        
        if valor_espe_str in ["OE.3 ARQUITECTURA", "OE.8 MOBILIARIO"] and valor_descri_str == "N":
            resultados_col3.append(valor_col1)
            
    if not resultados_col3:
        forms.alert("No se encontraron coincidencias con los filtros especificados.", exitscript=False)
    # else:
    #     # Mostrar la cantidad total de elementos encontrados sin usar f-string
    #     cantidad_resultados = len(resultados_col3)
    #     forms.alert("Se encontraron " + str(cantidad_resultados) + " elementos que cumplen con los criterios.", exitscript=False)
    
    return resultados_col3
























