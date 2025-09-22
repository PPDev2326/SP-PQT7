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
    """
    Lee un archivo Excel y filtra elementos según requerimientos COBie.
    
    Busca columnas 'Item', 'Specialty', 'COBie Requirement'.
    Clasifica elementos en dos listas según COBie Requirement (Y/N).
    
    Returns:
        list: [elementos_sin_cobie, elementos_con_cobie]
    """
    ruta = forms.pick_excel_file()
    if not ruta:
        forms.alert("No se seleccionó ningún archivo.", exitscript=True)
        return []
    
    datos = xl.load(ruta, sheets="Datos", headers=False)
    filas = datos.get('ELEMENTOS', {}).get('rows', [])
    
    if len(filas) < 8:
        forms.alert("El archivo no tiene suficientes filas para contener encabezados en la fila 9.", exitscript=True)
        return []
    
    encabezados = filas[8]  # Fila 9 en Excel (índice 8)
    try:
        col_espe = encabezados.index("Specialty")               # Columna B
        col_descri = encabezados.index("COBie Requirement")     # Columna I
        col_item = encabezados.index("Item")                    # Columna A
    except ValueError:
        forms.alert("No se encontraron las columnas necesarias: 'Especialidad', 'COBie', 'NRM 1'.", exitscript=True)
        return []
    
    resultados_cobie = []
    resultados_sin_cobie = []
    
    for fila in filas[9:]:  # Desde la fila 10 (índice 9)
        if len(fila) <= max(col_espe, col_descri, col_item):
            continue
        
        valor_espe = fila[col_espe]
        valor_descri = fila[col_descri]
        valor_col1 = fila[col_item]
        
        # Convertir valores a cadena antes de aplicar comparaciones
        valor_espe_str = str(valor_espe).strip()
        valor_descri_str = str(valor_descri).strip()
        
        if not valor_espe_str or valor_descri_str == "N":
            resultados_sin_cobie.append(valor_col1)
        if valor_espe_str and valor_descri_str == "Y":
            resultados_cobie.append(valor_col1)
            
    if not resultados_cobie:
        forms.alert("No se encontraron coincidencias con los filtros especificados.", exitscript=False)
    # else:
    #     # Mostrar la cantidad total de elementos encontrados sin usar f-string
    #     cantidad_resultados = len(resultados_col3)
    #     forms.alert("Se encontraron " + str(cantidad_resultados) + " elementos que cumplen con los criterios.", exitscript=False)
    
    return [resultados_sin_cobie, resultados_cobie]
























