import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os
import re
import unicodedata
import tempfile

# ==========================================
# COLUMNAS REQUERIDAS POR SAP
# ==========================================
COLUMNAS_PROTOCOLO = [
    'ResControlProd', 'GrupoPlanificador', 'Material', 'Descripcion', 
    'GrupoReceta', 'Version', 'Ancho', 'Perimetro', 'PorcEficiencia', 
    'NombreRecurso', 'EquipoXRecurso', 'Consecutivo', 'Velocidad', 
    'Alistamiento', 'OperarioXRecurso', '%Enlace', '%Recorte', 
    '%Sobrepeso', 'AfectaEnlace'
]

COLUMNAS_ESPECIFICACIONES = [
    'MATERIAL', 'Peso GTA cruda', 'Peso GTA Horneada', 'PESO GTA ROCEADA (gr)', 
    'Gtas/Sanduche', 'Base tiempo/min', 'Cant.crema sanduche', 'Galleta /paqueticos', 
    'PAQUETICOS UNIDAD', 'UNIDADES CORRUGADO', 'LARGO', 'ANCHO', 'CALIBRE', 
    'PESO N.IMPRESO UV(KG)', 'PESO PASTA CRUDA'
]

MAPEO_BASE = {
    'PorcEficiencia': ['porceficiencia', 'eficiencia', '% eficiencia'],
    'GrupoPlanificador': ['grupo planificador', 'grpplanif', 'planificador'],
    'NombreRecurso': ['pstotbjo', 'nombre recurso', 'recurso'],
    'Consecutivo': ['nrosecuencia', 'secuencia', 'operacion'],
    'Alistamiento': ['tiempominalistamientos', 'tiempo alistamiento'],
    'EquipoXRecurso': ['equiposxrecurso', 'equipos x recurso', 'equipo x recurso', 'cant equipos', 'cantequipos', '# equipos'],
    'OperarioXRecurso': ['totaloperarios', 'total operarios', 'mano de obra', 'operarios'],
    '%Enlace': ['porenlace', 'porcenlace', '% enlace', 'enlace'],
    '%Recorte': ['porcrecorte', '% recorte', 'recorte'],
    '%Sobrepeso': ['porcsobrepeso', '% sobrepeso', 'sobrepeso'],
    'AfectaEnlace': ['afectaenlace', 'afecta enlace'],
    'PESO GTA ROCEADA (gr)': ['pesogtaroceada', 'galleta roceada'],
    'Gtas/Sanduche': ['gtasanduche', 'galletas sanduche'],
    'Cant.crema sanduche': ['pesocremagta', 'cant crema sanduche'],
    'Galleta /paqueticos': ['gtaxpaquetico', 'galleta/paqueticos'],
    'PAQUETICOS UNIDAD': ['paqueticosxunidad', 'paqueticos unidad'],
    'UNIDADES CORRUGADO': ['unidadesxcorrugado', 'unidades corrugado'],
    'PESO N.IMPRESO UV(KG)': ['peso neto univenta especificaciones', 'peso n impreso']
}

def limpiar(texto):
    if pd.isna(texto): return ""
    t = str(texto)
    t = ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-zA-Z0-9]', '', t).lower()

def es_numero_real(val):
    if pd.isna(val) or str(val).strip() == "": return False
    try:
        float(str(val).replace(',', '.'))
        return True
    except ValueError:
        return False

def convertir_numero_normal(val):
    num = float(str(val).replace(',', '.'))
    return int(num) if num.is_integer() else num

def convertir_eficiencia(val):
    num = float(str(val).replace(',', '.'))
    if 0 < num <= 1.0: num = round(num * 100, 2)
    return int(num) if num.is_integer() else num

def resolver_operarios_y_equipos(recurso_str, dict_mo, dict_eq):
    rec = str(recurso_str).strip().upper()
    maquina = None
    if 'CAHO' in rec or 'CAM' in rec: maquina = 'CAMARA'
    elif 'EMBM' in rec or 'EMB' in rec: maquina = 'EMB'
    elif 'LAMI' in rec: maquina = 'LAMINA'
    elif 'EMPI' in rec or 'IND' in rec: maquina = 'IND'
    elif 'EMPM' in rec or 'MEP' in rec: maquina = 'MEP'
    elif 'ROCI' in rec: return 0.0, 1

    mo = dict_mo.get(maquina) if maquina else None
    eq = dict_eq.get(maquina) if maquina else None
    return mo, eq

def ajustar_secuencia_roceado(df_base):
    col_mat = next((c for c in df_base.columns if limpiar(c) == 'material'), None)
    col_rec = next((c for c in df_base.columns
