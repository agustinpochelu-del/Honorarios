import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide")
st.title("📊 Panel de Control de Facturación y Honorarios")

# 2. CONFIGURACIÓN DEL ORIGEN DE DATOS (ONEDRIVE)
USAR_NUBE = True

# Tu link de OneDrive modificado para descarga directa
URL_NUBE = "https://1drv.ms/x/c/d157fed8b9ecc198/IQBjbEIKjpuyQ5t6EsSIuXVmAXAon-EyPNaN4Ae0qbskn2E?download=1"

# 3. Motor de carga y cálculo de actualización
@st.cache_data(ttl=600)  # Se refresca automáticamente cada 10 minutos
def cargar_y_procesar_datos(origen, es_nube):
    # Si viene de la nube, descargamos el archivo en memoria antes de pasárselo a Pandas
    if es_nube:
        try:
            respuesta = requests.get(origen)
            respuesta.raise_for_status() 
            archivo_final = BytesIO(respuesta.content)
        except Exception as e:
            raise Exception(f"Error al descargar desde OneDrive: {e}")
    else:
        archivo_final = origen

    # Leemos las tres hojas del archivo XLSX
    df_facturas = pd.read_excel(archivo_final, sheet_name="Facturas")
    df_clientes = pd.read_excel(archivo_final, sheet_name="Clientes")
    df_indices = pd.read_excel(archivo_final, sheet_name="Indices")
    
    # Limpieza preventiva de espacios en los nombres de las columnas
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns = df_indices.columns.str.strip()
    
    # Forzamos a que las columnas de valores sean numéricas
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Facturacion $'] = pd.to_numeric(df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce')
    if 'precio' in df_clientes.columns:
        df_clientes['precio'] = pd.to_numeric(df_clientes['precio'].astype(str).str.replace(',', '.'), errors='coerce')
    
    # Conversión de Fechas internas
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'], errors='coerce')
    
    # Llevamos la fecha de la factura al día 1 del mes para cruzar con la tabla de índices
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()
    
    # Buscamos el último índice disponible (el más reciente según calendario)
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
    
    # --- AJUSTE DE COEFICIENTE DE RESGUARDO ---
    # Calculamos el coeficiente nominal
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    
    # REGLA CONTABLE: Si el período de la factura es igual o posterior al último índice disponible
    # (o si no encuentra el índice en la tabla), forzamos a que el coeficiente sea 1.
    df_res['Coeficiente'] = df_res['Coeficiente'].fillna(1.0)
    
    # Multiplicamos la facturación original por el coeficiente definitivo
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    # Cruzamos con el Maestro de Clientes para traer la Denominación
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    # --- FILTRO Y ORDEN DE COLUMNAS PARA EL HISTORIAL VISUAL ---
    columnas_historial = [
        'Fecha_dt', 'Tipo', 'Punto de Venta', 'Número Desde', 
        'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada'
    ]
    df_historial_visual = df_final[[col for col in columnas_historial if col in df_final.columns]].copy()
    
    # Formateamos la fecha a DD/MM/YYYY
    df_historial_visual['Fecha'] = df_historial_visual['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_historial_visual = df_historial_visual.drop(columns=['Fecha_dt'])
    
    # Reordenamos columnas
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

# Determinamos el recurso a leer
recurso = URL_NUBE if USAR_NUBE else "Honosrario NM.xlsx"

try:
    # Procesamos la información
    df_facturacion_completa, df_clientes, df_indices, ult_mes, ult_ind = cargar_y_procesar_datos(recurso, USAR_NUBE)
    
    # Cartel informativo
    st.success(f"¡Conectado a OneDrive con éxito! Moneda homogénea base: {ult_mes.strftime('%m/%Y')} (Índice: {ult_ind})")
    
    # Botón manual para forzar recarga
    if st.button("🔄 Sincronizar cambios recientes de OneDrive"):
        st.cache_data.clear()
        st.rerun()
    
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
            
        st.dataframe(
            df_mostrar_clientes, 
            use_container_width=True,
            column_config={
                "precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f")
            }
        )
        
    with tab3:
        st.subheader("Historial de Facturación (Valores a Plata de Hoy)")
        st.dataframe(
            df_facturacion_completa, 
            use_container_width=True,
            column_config={
                "Facturacion $": st.column_config.NumberColumn("Facturación Original", format="$ %.2f"),
                "Facturacion $ Actualizada": st.column_config.NumberColumn("Facturación Actualizada", format="$ %.2f")
            }
        )
        
    with tab4:
        st.subheader("Índices de Referencia (IPC / IPIM)")
        st.dataframe(df_indices, use_container_width=True)

except Exception as e:
    st.error(f"Ocurrió un error al procesar los datos: {e}")
