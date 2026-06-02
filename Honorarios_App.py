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
    
    # Forzamos a que las columnas de valores sean numéricas
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Facturacion $'] = pd.to_numeric(df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce')
    
    # Conversión de Fechas internas para procesamiento técnico
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'], errors='coerce')
    
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
    
    # Calculamos el coeficiente de actualización y la facturación real a hoy
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    # Cruzamos con el Maestro de Clientes para traer la "Denominación Receptor"
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    # --- FILTRO Y ORDEN DE COLUMNAS PARA EL HISTORIAL VISUAL ---
    # Seleccionamos solo lo requerido, eliminando Moneda e Imp. Total
    columnas_historial = [
        'Fecha_dt', 'Tipo', 'Punto de Venta', 'Número Desde', 
        'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada'
    ]
    df_historial_visual = df_final[[col for col in columnas_historial if col in df_final.columns]].copy()
    
    # Formateamos la fecha del historial a DD/MM/YYYY
    df_historial_visual['Fecha'] = df_historial_visual['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_historial_visual = df_historial_visual.drop(columns=['Fecha_dt'])
    
    # Reordenamos para que 'Fecha' vuelva a estar al principio
    columnas_ordenadas = ['Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    df_historial_visual = df_historial_visual[columnas_ordenadas]
    
    # --- FORMATEO DE FECHAS EN CLIENTES E ÍNDICES (MM/YYYY) ---
    df_clientes_visual = df_clientes.copy()
    if 'Actualizacion' in df_clientes_visual.columns:
        df_clientes_visual['Actualizacion'] = pd.to_datetime(df_clientes_visual['Actualizacion'], errors='coerce').dt.strftime('%m/%Y')
        
    df_indices_visual = df_indices.copy()
    df_indices_visual['MES'] = df_indices_visual['MES_dt'].dt.strftime('%m/%Y')
    df_indices_visual = df_indices_visual.drop(columns=['MES_dt'])
    
    return df_historial_visual, df_clientes_visual, df_indices_visual, ultimo_mes, ultimo_indice

# Nombre de tu archivo de datos
archivo_excel = "Honosrario NM.xlsx"

try:
    # Procesamos la información con el nuevo esquema
    df_facturacion_completa, df_clientes, df_indices, ult_mes, ult_ind = cargar_y_procesar_datos(archivo_excel)
    
    # Cartel informativo del período base
    st.success(f"¡Datos procesados correctamente! Moneda homogénea calculada con base en el período: {ult_mes.strftime('%m/%Y')} (Índice: {ult_ind})")
    
    # Estructura de pestañas
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Tablero de Control", 
        "👥 Maestro de Clientes", 
        "🧾 Facturas Procesadas", 
        "📊 Índices Históricos"
    ])
    
    with tab1:
        st.subheader("Análisis Global de Facturación")
        st.info("Espacio listo para armar gráficos y las alertas de revisión de abonos.")
        
    with tab2:
        st.subheader("Maestro de Clientes")
        ver_solo_activos = st.checkbox("Mostrar solo clientes Activos", value=True)
        
        if ver_solo_activos:
            df_mostrar_clientes = df_clientes[df_clientes['Estado'] == 'Activo']
        else:
            df_mostrar_clientes = df_clientes
            
        st.dataframe(df_mostrar_clientes, use_container_width=True)
        
    with tab3:
        st.subheader("Historial de Facturación (Valores a Plata de Hoy)")
        st.dataframe(df_facturacion_completa, use_container_width=True)
        
    with tab4:
        st.subheader("Índices de Referencia (IPC / IPIM)")
        st.dataframe(df_indices, use_container_width=True)

except FileNotFoundError:
    st.error(f"No se encontró el archivo '{archivo_excel}'. Verificá que esté guardado en el mismo directorio.")
except Exception as e:
    st.error(f"Ocurrió un error al procesar los datos: {e}")
