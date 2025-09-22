# -*- coding: utf-8 -*-

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

# Obtener el documento activo
doc = __revit__.ActiveUIDocument.Document

nombre_archivo = doc.Title

fecha_mantenimiento = {
    '200080' : "2025-03-24T13:29:49",
    '200082' : "2025-04-15T13:29:49",
    '200091' : "2025-03-24T13:29:49",
    '200094' : "2025-03-24T13:29:49",
    '200096' : "2025-02-28T13:29:49",
    '200097' : "2025-03-10T13:29:49",
    '200102' : "2025-02-15T13:29:49",
    '200104' : "2025-03-24T13:29:49",
    '200108' : "2025-02-28T13:29:49",
    '200111' : "2025-03-10T13:29:49",
    '200112' : "2025-03-01T13:29:49",
    '200117' : "2025-04-07T13:29:49",
    '200127' : "2025-02-28T13:29:49",
    '200141' : "2025-02-28T13:29:49"
}

colegios = {
    '200080' : 'Virgen del Carmen',
    '200082' : '2028',
    '200091' : 'Peru Japon',
    '200094' : 'Santiago Antunez de Mayolo',
    '200096' : 'Virgen de Fatima',
    '200097' : '2051',
    '200102' : 'Simon Bolivar',
    '200104' : 'Juan Velasco Alvarado',
    '200108' : 'San Benito',
    '200111' : 'Imperio Tahuantinsuyo',
    '200112' : 'Vista Alegre',
    '200117' : 'William Fullbright',
    '200127' : 'Felipe Huaman Poma de Ayala',
    '200141' : 'General Prado'
}

especialidades = {
    'Arquitectura' : 'AR',
    'Estructuras': 'ES',
    'Instalaciones de comunicacion' : 'DT',
    'Instalaciones eléctricas' : 'EE',
    'Instalaciones mecánicas' : 'ME',
    'Instalaciones sanitarias' : 'SA',
    'Señaletica y evacuacion' : 'AR',
    'Alarma y deteccion' : 'DT',
    'Equipamiento y mobiliario' : 'EM'
}

series = {
    'Arquitectura' : '-CGC01-DO-CA-000013 | Dossier de Calidad - Arquitectura',
    'Instalaciones de comunicacion' : '-CGC01-DO-CA-000012 | Dossier de Calidad - Especialidades',
    'Instalaciones eléctricas' : '-CGC01-DO-CA-000012 | Dossier de Calidad - Especialidades',
    'Instalaciones mecánicas' : '-CGC01-DO-CA-000012 | Dossier de Calidad - Especialidades',
    'Instalaciones sanitarias' : '-CGC01-DO-CA-000012 | Dossier de Calidad - Especialidades',
    'Señaletica y evacuacion' : '-CGC01-DO-CA-000013 | Dossier de Calidad - Arquitectura',
    'Alarma y deteccion' : '-CGC01-DO-CA-000012 | Dossier de Calidad - Especialidades',
}

codigo_5D = {
    'Arquitectura' : 'OE.3',
    'Equipamiento y mobiliario' : 'OE.8',
    'Estructuras' : 'OE.2',
    'Instalaciones de comunicacion' : 'OE.6',
    'Instalaciones eléctricas' : 'OE.5',
    'Instalaciones mecánicas' : 'OE.5',
    'Instalaciones sanitarias' : 'OE.4',
    'Señaletica y evacuacion' : 'OE.3',
    'Alarma y deteccion' : 'OE.6',
}

contract = {
    'Virgen del Carmen' : '2457231',
    '2028' : '2459062',
    'Peru Japon' : '',
    'Santiago Antunez de Mayolo' : '2468424',
    'Virgen de Fatima' : '2433311',
    '2051' : '2471211',
    'Simon Bolivar' : '2475458',
    'Juan Velasco Alvarado' : '',
    'San Benito' : '2475484',
    'Imperio Tahuantinsuyo' : '2475363',
    'Vista Alegre' : '2475360',
    'William Fullbright' : '2481404',
    'Felipe Huaman Poma de Ayala' : '2485504',
    'General Prado' : '2489056'
}

institucion = {
    'Virgen del Carmen' : '301753',
    '2028' : '333099',
    'Peru Japon' : '',
    'Santiago Antunez de Mayolo' : '305911',
    'Virgen de Fatima' : '296803',
    '2051' : '296676',
    'Simon Bolivar' : '333103',
    'Juan Velasco Alvarado' : '',
    'San Benito' : '297077',
    'Imperio Tahuantinsuyo' : '305925',
    'Vista Alegre' : '319140',
    'William Fullbright' : '305671',
    'Felipe Huaman Poma de Ayala' : '313732',
    'General Prado' : '142561'
}

trama = {
    'Virgen del Carmen' : 'P01185446',
    '2028' : 'P01394499',
    'Peru Japon' : '',
    'Santiago Antunez de Mayolo' : 'P01392995',
    'Virgen de Fatima' : 'P01394175',
    '2051' : 'P01394551',
    'Simon Bolivar' : 'P44229293',
    'Juan Velasco Alvarado' : '',
    'San Benito' : 'P01254528',
    'Imperio Tahuantinsuyo' : 'P01190869',
    'Vista Alegre' : 'P01260012',
    'William Fullbright' : 'P01394473',
    'Felipe Huaman Poma de Ayala' : 'P02184887',
    'General Prado' : '70266546'
}

reglas = {
    "Instalaciones sanitarias": [
        "Resistente al agua", "Resistente al agua",
        "S300 | RNE IS.010-IS.020 | RNE OS.060 | ASTM D1785, D2466, D2241",
        "Bajo consumo de agua",
        "Aparato sanitario, tuberia, union de tuberia, equipo mecanico y accesorio de tuberia"
        ],
    "Instalaciones eléctricas": [
        "Ahorro de energia", "Eficiencia energética",
        "S300 | CNE-U | CNE-S | RNE EM.010 | IEC 60287 | IEEE-80  | ANSI | ASTM | NEC 250 | UNE-EN 12464-2 | NTP 370",
        "Eficiencia energetica",
        "Bandeja, caja de salida, equipo electrico, luminaria, placa frontal, rieles, tuberia flexible, tuberia para alimentadores, tuberia para iluminacion, union de bandeja y union de tubo"
        ],
    "Instalaciones mecánicas": [
        "Ahorro de energia", "Eficiencia energética",
        "S1300 | RNE | ASHRAE | SMACNA | ANSI | AHRI | EUROVENT | NEMA | UL | NFPA | CNE",
        "Eficiencia energetica",
        "Accesorio mecanico, conducto, union de conducto, tuberia, union de tuberia, terminal de aire y equipo mecanico"
        ],
    "Instalaciones de comunicacion": [
        "Ahorro de energia", "Durabilidad a largo plazo",
        "S300 | ANSI/TIA-568 | ISO/IEC 11801 | IEEE 802.3 | NFPA 72 | UL 2044 | EN 50132 | IEC 60268 | ISA 5.1.",
        "Resistencia a largo plazo",
        "Dispositivos de comunicación, bandejas de cables, uniones de bandeja de cables, equipos eléctricos, accesorios de tuberías, uniones de tubo, tubos y tuberías flexibles"
        ],
    "Equipamiento y mobiliario": [
        "Resistente al desgaste", "Reciclabilidad",
        "NTP 260 | RVM N 019 y 208 | ASTM A36/A36M | EN ISO 14184 | SCAQMD Norma 1113",
        "Resistencia a largo plazo",
        "Accesorios de montaje y manual de instrucciones"
        ],
    "Arquitectura": [
        "Resistente al desgaste", "Durabilidad a largo plazo",
        "999968-KO001-PR-AS-000001 | 999968-KO001-BD-AS-000001 | RVM N 010, 104, 208 | RNE",
        "Resistencia a largo plazo",
        "Accesorios de montaje y manual de instrucciones"
        ],
    "Señaletica y evacuacion": [
        "Resistente al desgaste", "Durabilidad a largo plazo",
        "S300 | D.S. N° 002-2018-PCM | NTP 350, 399",
        "Resistencia a largo plazo",
        "Accesorios de montaje y manual de instrucciones"
        ],
    "Alarma y deteccion" : [
        "Resistente al fuego", "Durabilidad a largo plazo",
        "S300 | RNE| CNE| NFPA 70| NFPA 72| ADA| UL 38| UL 268| UL 464| UL 521| UL 864| UL 1481| UL 1971",
        "Resistencia a largo plazo",
        "Dispositivos de comunicación, bandejas de cables, uniones de bandeja de cables, equipos eléctricos, accesorios de tuberías, uniones de tubo, tubos y tuberías flexibles"
        ]
}

def ObtenerCaracteristicas(especialidad):
    """
    Devuelve una lista de características asociadas a una especialidad.
    :param especialidad: (str) Nombre de la especialidad del modelo actual.
    :return: List[str | None] Lista con cinco elementos: [feature, performance, normas, sostenibilidad, Constituents].
    Si la especialidad no se encuentra, se devuelve una lista con un mensaje de error
    seguido de valores None.
    """
    valores = reglas.get(especialidad)
    
    if valores is None:
        return ["No se tiene la especialidad" + especialidad, None, None, None, None]
    return valores


def ObtenerCodigoContrato(doc):
    codigo_colegio = ObtenerCodigoColegio(doc)
    nombre_colegio = colegios.get(codigo_colegio)
    
    if not nombre_colegio:
        return "❌ Error: Colegio no encontrado para el codigo " + codigo_colegio
    
    return contract.get(nombre_colegio, "❌ Error: Contrado no encontrado para el colegio " + nombre_colegio)

def ObtenerCodigoInstitucion(doc):
    codigo_colegio = ObtenerCodigoColegio(doc)
    nombre_colegio = colegios.get(codigo_colegio)
    
    if not nombre_colegio:
        return None
    
    return institucion.get(nombre_colegio, "❌ Error: codigo de institucion no encontrada para el colegio " + nombre_colegio)

def ObtenerCodigoTrama(doc):
    codigo_colegio = ObtenerCodigoColegio(doc)
    nombre_colegio = colegios.get(codigo_colegio)
    
    if not nombre_colegio:
        return None
    
    return trama.get(nombre_colegio, "❌ Error: codigo de institucion no encontrada para el colegio " + nombre_colegio)

def obtener_codigo_5D():
    especialidad = obtener_especialidad()
    if especialidad == "Señaletica y evacuacion":
        especialidad = "Arquitectura"
    elif especialidad == "Alarma y deteccion":
        especialidad = "Instalaciones de comunicacion"
    
    valor = codigo_5D.get(especialidad)
    
    if valor:
        return especialidad, valor
    else:
        return None, "Codigo 5D no encontrado"

def obtener_especialidad():
    name_partes = nombre_archivo.split("-")[-2] # Obtenemos la especialidad del nombre del archivo
    name_codigo_final = nombre_archivo.split("-")[-1][-2:] #Obtenemos los ultimos 2 digitos del nomrbre del archivo
    for k, v in especialidades.items():
        if v in name_partes:
            if v == "AR":
                if name_codigo_final == "01":
                    return 'Arquitectura'
                elif name_codigo_final ==  "03":
                    return 'Señaletica y evacuacion'
            
            elif v == "DT":
                if name_codigo_final == "01":
                    return 'Instalaciones de comunicacion'
                elif name_codigo_final ==  "03":
                    return 'Alarma y deteccion'
            
            else:
                return k
    
    return "No se encontro ninguna especialidad en mi base de datos"

def obtener_codigo_colegio():
    """
    ⚠️ Método obsoleto: usar `ObtenerCodigoColegio(doc)` en su lugar.
    Este método ya no debe utilizarse y será eliminado en futuras versiones.
    """
    codigo_colegio = nombre_archivo.split("-")[0] # Obtenemos El codigo del colegio
    for k, v in colegios.items():
        if k in codigo_colegio:
            if k == codigo_colegio:
                return k
            else:
                return "Codigo no encontrado"

def ObtenerCodigoColegio(doc):
    titulo = doc.Title
    partes = titulo.split("-") #Obtenemos el codigo del colegio del modelo actual
    if len(partes) > 0:
        return partes[0]
    return "Formato de archivo no reconocido"

def obtener_mantenimiento():
    codigo_colegio = obtener_codigo_colegio()
    return fecha_mantenimiento.get(codigo_colegio, "No se encontro fecha de mantenimiento")

CODE_UNICLASS = "Pr.40.50"

def obtener_serial_number(activo, id):
    especialidad = obtener_especialidad()
    codigo_colegio = ObtenerCodigoColegio(doc)
    
    if "No se encontro" in especialidad or codigo_colegio == "Codigo no encontrado":
        return "Error: Datos insuficientes para construir el serial."
    
    if especialidad == 'Equipamiento y mobiliario':
        return ".".join([codigo_colegio, CODE_UNICLASS, activo, id])
    
    serie = series.get(especialidad)
    if not serie:
        return "Error: Serie no encontrada para la especialidad " + especialidad
    
    return codigo_colegio + serie

def obtener_model_number():
    especialidad = obtener_especialidad()
    codigo_colegio = ObtenerCodigoColegio(doc)
    
    if "No se encontro" in especialidad or codigo_colegio == "Codigo no encontrado":
        return "Error: Datos insuficientes para construir el serial."
    
    serie = series.get(especialidad)
    if not serie:
        return "Error: Serie no encontrada para la especialidad " + especialidad
    
    return codigo_colegio + serie