import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime

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
    
    # Limpieza de columnas comentadas o vacías en índices
    df_indices = df_indices.loc[:, ~df_indices.columns.str.contains('^Unnamed')]
    columnas_indices_validas = ['MES', 'IPC  IPIM']
    df_indices = df_indices[[col for col in columnas_indices_validas if col in df_indices.columns]]
    
    # Forzamos a que las columnas de valores sean numéricas
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Facturacion $'] = pd.to_numeric(df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce')
    
    # Limpieza estricta de la columna precio en clientes
    if 'precio' in df_clientes.columns:
        df_clientes['precio'] = pd.to_numeric(df_clientes['precio'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
    
    # Conversión de Fechas internas
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'], errors='coerce')
    
    # Llevamos la fecha de la factura al día 1 del mes para cruzar con la tabla de índices
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()
    
    # Buscamos el último índice disponible
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
    
    # Ajuste de coeficiente de resguardo
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Coeficiente'] = df_res['Coeficiente'].fillna(1.0)
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    # Cruzamos con el Maestro de Clientes para traer la Denominación
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    # --- FILTRO Y ORDEN DE COLUMNAS PARA EL HISTORIAL VISUAL ---
    columnas_historial = [
        'Fecha_dt', 'Tipo', 'Punto de Venta', 'Número Desde', 
        'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada'
    ]
    df_historial_visual = df_final[[col for col in columnas_historial if col in df_final.columns]].copy()
    df_historial_visual['Fecha'] = df_historial_visual['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_historial_visual = df_historial_visual.drop(columns=['Fecha_dt'])
    
    columnas_ordenadas = ['Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    df_historial_visual = df_historial_visual[columnas_ordenadas]
    
    # --- PROCESAMIENTO LOGÍSTICO DE CLIENTES (ORDENAMIENTO Y ALERTAS) ---
    df_clientes_proc = df_clientes.copy()
    
    # Aseguramos que existan las columnas de control temporal
    df_clientes_proc['Actualizacion_dt'] = pd.to_datetime(df_clientes_proc['Actualizacion'], errors='coerce')
    df_clientes_proc['periodos'] = pd.to_numeric(df_clientes_proc['periodos'], errors='coerce').fillna(0)
    
    # Calculamos la alerta de vencimiento de honorario
    hoy = datetime.now()
    def verificar_vencimiento(row):
        if pd.isna(row['Actualizacion_dt']) or row['Actualiza'] != 'Si' or row['Estado'] != 'Activo':
            return "OK"
        # Diferencia aproximada en meses
        meses_transcurridos = (hoy.year - row['Actualizacion_dt'].year) * 12 + (hoy.month - row['Actualizacion_dt'].month)
        if meses_transcurridos >= row['periodos']:
            return "VENCIDO"
        return "OK"
    
    df_clientes_proc['Alerta_Revisión'] = df_clientes_proc.apply(verificar_vencimiento, axis=1)
    
    # Ordenamos: Activos primero, e Inactivos abajo. Dentro de cada grupo, el de mayor precio primero.
    # Convertimos temporalmente Estado a categoría o usamos ordenamiento booleano simulado (Activo > Inactivo)
    df_clientes_proc['Orden_Estado'] = df_clientes_proc['Estado'].apply(lambda x: 0 if x == 'Activo' else 1)
    df_clientes_proc = df_clientes_proc.sort_values(by=['Orden_Estado', 'precio'], ascending=[True, False])
    df_clientes_proc = df_clientes_proc.drop(columns=['Orden_Estado'])
    
    # Formateamos la fecha interna para la visualización definitiva
    df_clientes_proc['Actualizacion'] = df_clientes_proc['Actualizacion_dt'].dt.strftime('%m/%Y')
    df_clientes_proc = df_clientes_proc.drop(columns=['Actualizacion_dt'])
    
    # --- FORMATEO DE ÍNDICES ---
    df_indices_visual = df_indices.copy()
    df_indices_visual['MES'] = df_indices_visual['MES_dt'].dt.strftime('%m/%Y')
    df_indices_visual = df_indices_visual.drop(columns=['MES_dt'])
    
    return df_historial_visual, df_clientes_proc, df_indices_visual, ultimo_mes, ultimo_indice

# Funciones de Estilos Condicionales para las filas de clientes
def colorear_clientes(row):
    estilos = [''] * len(row)
    # Regla 1: Si está Inactivo, toda la fila se tiñe de un tono grisáceo/rojo tenue
    if row['Estado'] == 'Inactivo':
        return ['background-color: #fee2e2; color: #991b1b; opacity: 0.7;'] * len(row)
    
    # Regla 2: Si está Activo pero el abono está VENCIDO, resaltamos la alerta en amarillo/naranja contable
    if row['Alerta_Revisión'] == 'VENCIDO':
        # Buscamos el índice de la columna Alerta_Revisión para pintarla específicamente
        idx_alerta = row.index.get_loc('Alerta_Revisión')
        estilos[idx_alerta] = 'background-color: #fef08a; color: #854d0e; font-weight: bold;'
    return estilos

# Determinamos el recurso a leer
recurso = URL_NUBE if USAR_NUBE else "Honosrario NM.xlsx"

try:
    df_facturacion_completa, df_clientes, df_indices, ult_mes, ult_ind = cargar_y_procesar_datos(recurso, USAR_NUBE)
    
    st.success(f"¡Conectado a OneDrive con éxito! Moneda homogénea base: {ult_mes.strftime('%m/%Y')} (Índice: {ult_ind})")
    
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
        st.info("Espacio listo para armar gráficos contables sobre la facturación indexada.")
        
    with tab2:
        st.subheader("Maestro y Control de Abonos (Priorizado por Estado y Valor)")
        
        # Aplicamos el formateador de estilos condicionales de Pandas antes de renderizar
        df_estilado = df_clientes.style.apply(colorear_clientes, axis=1)
        
        st.dataframe(
            df_estilado, 
            use_container_width=True,
            column_config={
                "precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
                "Alerta_Revisión": st.column_config.TextColumn("Estado de Revisión")
            }
        )
        st.caption("💡 *Nota visual: Las filas rojas indican clientes Inactivos. Las celdas resaltadas en amarillo marcan que el plazo de revisión técnica (periodos) expiró respecto a la fecha de última actualización.*")
        
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
