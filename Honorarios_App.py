import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
import plotly.express as px

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide", initial_sidebar_state="expanded")
st.title("📊 Panel de Control de Facturación y Honorarios")

# Ruta del archivo unificado en el servidor
ARCHIVO_LOCAL = "Honosrario NM.xlsx"

# --- INICIALIZACIÓN DE MEMORIA INTERNA (CATEGORÍAS POR DEFECTO) ---
if 'cat_agustin' not in st.session_state:
    st.session_state['cat_agustin'] = 'D'
if 'cat_laura' not in st.session_state:
    st.session_state['cat_laura'] = 'B'

# --- CONFIGURACIÓN DE ESCALAS DE MONOTRIBUTO ---
ESCALAS_MONOTRIBUTO = {
    'A': 6450000, 'B': 9450000, 'C': 13250000, 'D': 16450000,
    'E': 19350000, 'F': 24250000, 'G': 29000000, 'H': 44000000,
    'I': 49115000, 'J': 56400000, 'K': 68000000
}

def formato_abreviado(valor):
    if valor >= 1_000_000:
        return f"$ {valor / 1_000_000:.2f} M"
    elif valor >= 1_000:
        return f"$ {valor / 1_000:.1f} k"
    else:
        return f"$ {valor:.2f}"

# 2. Motor de carga y cálculo
@st.cache_data(ttl=300) # Refresca caché cada 5 minutos por si cambiaste el excel por fuera
def cargar_y_procesar_datos(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None, None, None, None, None, None, None, None, None

    df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
    df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
    df_indices = pd.read_excel(ruta_archivo, sheet_name="Indices")
    
    # Intentamos cargar la nueva hoja "Emisores"
    try:
        df_emisores = pd.read_excel(ruta_archivo, sheet_name="Emisores")
        df_emisores.columns = df_emisores.columns.str.strip()
        # Limpieza extrema del CUIT en la tabla Emisores para asegurar el cruce
        if 'Emisor' in df_emisores.columns:
            df_emisores['Emisor'] = df_emisores['Emisor'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True)
    except Exception:
        df_emisores = pd.DataFrame()
    
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns = df_indices.columns.str.strip()
    
    # --- CRUCE RELACIONAL DE CUITs ---
    if 'Emisor' in df_facturas.columns:
        # Limpiamos el CUIT de facturas de la misma manera
        df_facturas['Emisor'] = df_facturas['Emisor'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True)
        if not df_emisores.empty and 'Nombre Emisor' in df_emisores.columns:
            # Cruzamos Facturas con Emisores a través del CUIT
            df_facturas = pd.merge(df_facturas, df_emisores[['Emisor', 'Nombre Emisor']], on='Emisor', how='left')
    
    # Si por alguna razón la columna ya venía pre-cruzada en el Excel como 'Emisores.Nombre Emisor'
    if 'Nombre Emisor' not in df_facturas.columns and 'Emisores.Nombre Emisor' in df_facturas.columns:
        df_facturas['Nombre Emisor'] = df_facturas['Emisores.Nombre Emisor']
    
    # Aseguramos que no haya nulos
    if 'Nombre Emisor' in df_facturas.columns:
        df_facturas['Nombre Emisor'] = df_facturas['Nombre Emisor'].fillna('Agust
