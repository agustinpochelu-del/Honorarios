import streamlit as st
import pandas as pd

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide")
st.title("📊 Panel de Control de Facturación y Honorarios")

# 2. Motor de carga y cálculo de actualización
@st.cache_data
def cargar_y_procesar_datos(ruta_archivo):
    # Leemos las tres hojas del archivo XLSX
    df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
    df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
    df_indices = pd.read_excel(ruta_archivo, sheet_name="Indices")
    
    # Limpieza preventiva de espacios en los nombres de las columnas
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns = df_indices.columns.str.strip()
    
    # --- CORRECCIÓN DEL ERROR DE TEXTO VS NÚMERO ---
    # Forzamos a que las columnas de valores sean numéricas.
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Imp. Total'] = pd.to_numeric(df_facturas['Imp. Total'].astype(str).str.replace(',', '.'), errors='coerce')
    
    # --- MOTOR DE ACTUALIZACIÓN POR INFLACIÓN ---
    # Nos aseguramos de que las columnas de fecha sean de tipo datetime
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'])
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'])
    
    # Llevamos la fecha de la factura al día 1 del mes para cruzar con la tabla de índices
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()
    
    # Buscamos el último índice disponible (el más reciente de la tabla)
    df_indices_ordenado = df_indices.sort_values('MES_dt')
    ultimo_mes = df_indices_ordenado.iloc[-1]['MES_dt']
    ultimo_indice = df_indices_ordenado.iloc[-1]['IPC  IPIM']
    
    # Cruzamos las facturas con la tabla de índices
    df_res = pd.merge(
        df_facturas, 
        df_indices[['MES_dt', 'IPC  IPIM']], 
        left_on='Mes_Indice', 
        right_on='MES_dt', 
        how='left'
    )
    
    # Calculamos el coeficiente de actualización y el valor real a hoy
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Imp. Total Actualizado'] = df_res['Imp. Total'] * df_res['Coeficiente']
    
    # Cruzamos con el Maestro de Clientes usando el "Nro. Doc. Receptor"
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='
