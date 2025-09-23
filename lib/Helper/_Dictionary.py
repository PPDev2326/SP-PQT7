# -*- coding: utf-8 -*-

# Mapeo de nombres de habitaciones a códigos COBie
ROOM_NAME_MAPPING = {
    "6 : DEPÓSITO": "001",
    "5 : SUM": "002", 
    "1 : COCINA": "003",
    "3 : Z.B": "004",
    "4 : DEPOSITO ALIMENTOS": "005",
    "2 : INSPECCIÓN ALIMENTOS": "006",
    "7 : PASILLO": "007",
    "9 : AULA MUNICIPIO ESCOLAR": "001",  # Repite 001
    "8 : SALA DE PROFESORES": "002",     # Repite 002
    "10 : PASILLO": "003",               # Repite 003
    "13 : DEPÓSITO GENERAL": "008",
    "12 : TÓPICO": "009",
    "11 : DEPÓSITO EDUCACIÓN FISICA": "010",
    "14 : PSICOLOGIA": "011",
    "16 : S.H DAMAS": "012",
    "15 : S.H VARONES": "013",
    "17 : PASILLO": "014",
    "25 : S.H DAMAS": "004",             # Repite 004
    "24 : S.H VARONES": "005",           # Repite 005
    "23 : SECRETARIA/SALA DE ESPERA": "006", # Repite 006
    "22 : ARCHIVO": "007",               # Repite 007
    "21 : DEPÓSITO DE MATERIAL EDU.SEC.": "008", # Repite 008
    "18 : APAFA": "009",                 # Repite 009
    "19 : ADMINISTRACIÓN": "010",        # Repite 010
    "20 : DIRECCIÓN": "011",             # Repite 011
    "26 : PASILLO": "012",               # Repite 012
    "27 : AULA": "015",
    "30 : PASILLO": "016",
    "29 : AULA": "013",
    "32 : PASILLO": "014",
    "31 : AULA": "001",                  # Repite 001
    "28 : PASILLO": "002",               # Repite 002
    "33 : AULA": "017",
    "36 : AULA": "018",
    "35 : PASILLO": "019",
    "37 : AULA": "015",                  # Repite 015
    "39 : AULA": "016",
    "38 : PASILLO": "017",               # Repite 017
    "40 : AULA": "003",                  # Repite 003
    "34 : AULA": "004",                  # Repite 004
    "41 : PASILLO": "005",               # Repite 005
    "52 : MODULO DE CONECTIVIDAD": "020",
    "51 : DEPÓSITO AIP": "021",
    "50 : AIP": "022",
    "43 : S.H DAMAS": "023",
    "44 : S.H DAMAS": "024",
    "49 : S.H DISCAP": "025",
    "45 : CTO DE LIMPIEZA": "026",
    "48 : DEPÓSITO": "027",
    "54 : DUC.": "028",
    "53 : DUC.": "029",
    "42 : PASILLO": "030",
    "47 : INGRESO S.H.D": "031",
    "46 : INGRESO S.H.D": "032",
    "55 : MODULO DE CONECTIVIDAD": "018", # Repite 018
    "65 : DEPÓSITO AIP": "019",           # Repite 019
    "64 : AIP": "020",                    # Repite 020
    "57 : S.H DAMAS": "021",              # Repite 021
    "59 : S.H DAMAS": "022",              # Repite 022
    "60 : S.H DISCAP": "023",             # Repite 023
    "58 : CTO DE LIMPIEZA": "024",        # Repite 024
    "61 : DEPÓSITO": "025",               # Repite 025
    "56 : PASILLO": "026",                # Repite 026
    "62 : INGRESO S.H.D": "027",          # Repite 027
    "63 : INGRESO S.H.D": "028",          # Repite 028
    "76 : MODULO DE CONECTIVIDAD": "006", # Repite 006
    "75 : DEPÓSITO AIP": "007",           # Repite 007
    "74 : AIP": "008",                    # Repite 008
    "68 : S.H DAMAS": "009",              # Repite 009
    "69 : S.H DAMAS": "010",              # Repite 010
    "67 : S.H DISCAP": "011",             # Repite 011
    "70 : CTO DE LIMPIEZA": "012",        # Repite 012
    "71 : DEPÓSITO": "013",               # Repite 013
    "66 : PASILLO": "014",                # Repite 014
    "72 : INGRESO S.H.D": "015",          # Repite 015
    "73 : INGRESO S.H.D": "016",          # Repite 016
    "78 : LABORATORIO": "033",
    "77 : DEPÓSITO": "034",
    "79 : PASILLO": "035",
    "81 : BIBLIOTECA": "029",
    "80 : DEPÓSITO DE BIBLIOTECA": "030",
    "82 : PASILLO": "031",
    "84 : TALLER DE MUSICA": "017",       # Repite 017
    "83 : DEPÓSITO DE TALLER DE MUSICA": "018", # Repite 018
    "85 : PASILLO": "019",                # Repite 019
    "86 : AULA": "036",
    "87 : AULA": "037",
    "88 : PASILLO": "038",
    "89 : AULA": "032",
    "90 : AULA": "033",                   # Repite 033
    "91 : PASILLO": "034",                # Repite 034
    "92 : AULA": "020",                   # Repite 020
    "93 : AULA": "021",                   # Repite 021
    "94 : PASILLO": "022",                # Repite 022
    "95 : AULA": "039",
    "96 : PASILLO": "040",
    "97 : AULA": "035",                   # Repite 035
    "98 : PASILLO": "036",                # Repite 036
    "99 : AULA": "023",                   # Repite 023
    "100 : PASILLO": "024",               # Repite 024
    "101 : PASILLO": "041",
    "102 : ESCALERA": "037",              # Repite 037
    "103 : PASILLO": "038",               # Repite 038
    "104 : PASILLO": "042",
    "105 : ESCALERA": "039",              # Repite 039
    "106 : PASILLO": "040",               # Repite 040
    "107 : PASILLO": "025",               # Repite 025
    "108 : PASILLO": "043",
    "109 : ESCALERA": "041",              # Repite 041
    "110 : PASILLO": "042",               # Repite 042
    "111 : PASILLO": "026",               # Repite 026
    "112 : PASILLO": "044",
    "113 : ESCALERA": "043",              # Repite 043
    "114 : PASILLO": "044",               # Repite 044
    "115 : PASILLO": "027",               # Repite 027
    "116 : CONTROL": "045",
    "117 : CTO. TABLERO": "046",
    "118 : CTO ELECTRICO": "047",
    "119 : CTO.BASURA": "048",
    "120 : CUARTO DE BOMBAS": "049",
    "121 : LOSA MULTIUSOS": "999",
    "122 : LOSA MULTIUSOS": "999",        # Repite 999
    "123 : PASILLO": "050",
    "124 : PASILLO": "045",               # Repite 045
    "125 : PASILLO": "028",               # Repite 028
    "126 : TECHO": "001",                 # Repite 001
    "127 : TECHO": "002",                 # Repite 002
    "128 : TECHO": "003",                 # Repite 003
    "129 : TECHO": "004",                 # Repite 004
    "130 : TECHO": "005",                 # Repite 005
    "131 : TECHO": "006",                 # Repite 006
    "132 : TECHO": "007",                 # Repite 007
    "133 : TECHO": "008",                 # Repite 008
    "134 : TECHO": "009",                 # Repite 009
    "135 : TECHO": "010",                 # Repite 010
    "136 : TECHO": "011",                 # Repite 011
    "137 : EXTERIOR": "999",              # Repite 999
    "138 : LOSA INTERMEDIA 1": "051",
    "139 : LOSA INTERMEDIA 2": "029"      # Repite 029
}

def find_mapped_number(room_name):
    """
    Busca el código COBie correspondiente al nombre de la habitación.
    
    Args:
        room_name (str): Nombre de la habitación a buscar
        
    Returns:
        str: Código COBie si encuentra coincidencia, None en caso contrario
    """
    if not room_name:
        return None
    name = room_name.strip().upper()
    return ROOM_NAME_MAPPING.get(name)

def get_formatted_string(number, name, separator=" : "):
    """
    Formatea strings con separador consistente.
    
    Args:
        number (str): Número o código
        name (str): Nombre del elemento
        separator (str): Separador entre elementos
        
    Returns:
        str: String formateado "número : nombre"
    """
    return "{}{}{}".format(number, separator, name)