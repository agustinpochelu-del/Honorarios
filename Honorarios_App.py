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
URL_NUBE = "https://1drv.ms/x/c/d157fed8b9ecc198/IQBjbEIKjpuyQ5t6EsSIuXVmAXAon-EyPNaN4Ae0qbskn2E?download=1"

# 3. Motor de carga y cálculo de actualización
@st.cache_data(ttl=600)
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
    
    # Ajuste de coeficiente de resguardo e indexación histórica
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Coeficiente'] = df_res['Coeficiente'].fillna(1.0)
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    # Cruzamos con el Maestro de Clientes para traer la Denominación
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    df_historial_visual = df_final[['Fecha_dt', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']].copy()
    df_historial_visual['Fecha'] = df_historial_visual['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_historial_visual = df_historial_visual.drop(columns=['Fecha_dt'])
    columnas_ordenadas = ['Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    df_historial_visual = df_historial_visual[columnas_ordenadas]
    
    # --- PROCESAMIENTO LOGÍSTICO DE CLIENTES ---
    df_clientes_proc = df_clientes.copy()
    df_clientes_proc['Actualizacion_dt'] = pd.to_datetime(df_clientes_proc['Actualizacion'], errors='coerce')
    df_clientes_proc['periodos'] = pd.to_numeric(df_clientes_proc['periodos'], errors='coerce').fillna(0).astype(int)
    
    # Mapeo del IPC correspondiente al mes de la última actualización de cada cliente
    df_clientes_proc['Mes_Actualizacion_dt'] = df_clientes_proc['Actualizacion_dt'].dt.to_period('M').dt.to_timestamp()
    df_clientes_proc = pd.merge(
        df_clientes_proc, 
        df_indices[['MES_dt', 'IPC  IPIM']], 
        left_on='Mes_Actualizacion_dt', 
        right_on='MES_dt', 
        how='left'
    ).rename(columns={'IPC  IPIM': 'IPC_Ult_Actualizacion'}).drop(columns=['MES_dt', 'Mes_Actualizacion_dt'])
    
    # Cálculo de antigüedad en meses y alertas
    hoy = datetime.now()
    def calcular_metricas_comerciales(row):
        if pd.isna(row['Actualizacion_dt']):
            return pd.Series([0, "OK", row['precio']])
        
        # Cantidad exacta de meses de antigüedad de la última actualización
        meses_antigüedad = (hoy.year - row['Actualizacion_dt'].year) * 12 + (hoy.month - row['Actualizacion_dt'].month)
        
        # Cálculo del honorario sugerido por inflación acumulada
        if pd.notna(row['IPC_Ult_Actualizacion']) and row['IPC_Ult_Actualizacion'] > 0:
            coef_inflacion = ultimo_indice / row['IPC_Ult_Actualizacion']
            sugerido = row['precio'] * coef_inflacion
        else:
            sugerido = row['precio']
            
        alerta = "VENCIDO" if (row['Actualiza'] == 'Si' and row['Estado'] == 'Activo' and meses_antigüedad >= row['periodos']) else "OK"
        return pd.Series([meses_antigüedad, alerta, sugerido])
    
    df_clientes_proc[['Meses Desactualizado', 'Alerta_Revisión', 'Honorario Sugerido']] = df_clientes_proc.apply(calcular_metricas_comerciales, axis=1)
    df_clientes_proc['Meses Desactualizado'] = df_clientes_proc['Meses Desactualizado'].astype(int)
    
    # Orden jerárquico contable
    df_clientes_proc['Orden_Estado'] = df_clientes_proc['Estado'].apply(lambda x: 0 if x == 'Activo' else 1)
    df_clientes_proc = df_clientes_proc.sort_values(by=['Orden_Estado', 'precio'], ascending=[True, False]).drop(columns=['Orden_Estado'])
    
    # Guardamos la fecha limpia para exposición visual
    df_clientes_proc['Actualizacion_Str'] = df_clientes_proc['Actualizacion_dt'].dt.strftime('%m/%Y')
    
    df_indices_visual = df_indices.copy()
    df_indices_visual['MES'] = df_indices_visual['MES_dt'].dt.strftime('%m/%Y')
    df_indices_visual = df_indices_visual.drop(columns=['MES_dt'])
    
    return df_historial_visual, df_clientes_proc, df_indices_visual, ultimo_mes, ultimo_indice, df_clientes

# Estilos condicionales visuales
def colorear_clientes(row):
    estilos = [''] * len(row)
    if row['Estado'] == 'Inactivo':
        return ['background-color: #fee2e2; color: #991b1b; opacity: 0.7;'] * len(row)
    if row['Alerta_Revisión'] == 'VENCIDO':
        idx_alerta = row.index.get_loc('Alerta_Revisión')
        estilos[idx_alerta] = 'background-color: #fef08a; color: #854d0e; font-weight: bold;'
    return estilos

try:
    df_facturacion_completa, df_clientes, df_indices, ult_mes, ult_ind, df_clientes_original = cargar_y_procesar_datos(URL_NUBE, USAR_NUBE)
    
    st.success(f"¡Conectado a OneDrive con éxito! Moneda homogénea base: {ult_mes.strftime('%m/%Y')} (Índice: {ult_ind})")
    
    if st.button("🔄 Sincronizar cambios recientes de OneDrive"):
        st.cache_data.clear()
        st.rerun()
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Simulador y Propuestas", 
        "👥 Maestro de Clientes", 
        "🧾 Facturas Procesadas", 
        "📊 Índices Históricos"
    ])
    
    with tab1:
        st.subheader("🛠️ Entorno Interactiva de Actualización de Abonos")
        st.write("A continuación se listan los clientes activos cuyos contratos requieren revisión técnica por inflación. Podés editar directamente los valores en la columna **'Nuevo Precio Pactado'** para simular o consolidar tu estrategia comercial:")
        
        # Filtramos solo los activos que requieren revisión o tienen precio para simular
        df_simulacion = df_clientes[df_clientes['Estado'] == 'Activo'].copy()
        df_simulacion['Nuevo Precio Pactado'] = df_simulacion['precio'] # Inicialmente igual al precio actual
        
        columnas_sim = ['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido', 'Nuevo Precio Pactado']
        
        # Permitimos la edición interactiva de la tabla en pantalla
        df_editado = st.data_editor(
            df_simulacion[columnas_sim],
            use_container_width=True,
            disabled=['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido'],
            column_config={
                "precio": st.column_config.NumberColumn("Precio Actual", format="$ %.2f"),
                "Honorario Sugerido": st.column_config.NumberColumn("Sugerido por IPC", format="$ %.2f"),
                "Nuevo Precio Pactado": st.column_config.NumberColumn("Nuevo Precio Pactado ✏️", format="$ %.2f"),
                "Meses Desactualizado": st.column_config.NumberColumn("Meses Inmóvil", format="%d")
            },
            key="editor_abonos"
        )
        
        # Lógica para armar el archivo de salida para OneDrive si el usuario quiere aplicar los cambios
        if st.button("💾 Generar Nueva Planilla de Clientes Actualizada"):
            df_maestro_nuevo = df_clientes_original.copy()
            
            # Reemplazamos los precios viejos por los nuevos pactados en la simulación
            for idx, row in df_editado.iterrows():
                cuit = row['Nro. Doc. Receptor']
                nuevo_val = row['Nuevo Precio Pactado']
                
                # Si el precio cambió, lo actualizamos en la tabla original y cambiamos la fecha de actualización a hoy
                if nuevo_val != row['precio']:
                    df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'precio'] = nuevo_val
                    df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'Actualizacion'] = datetime.today().strftime('%Y-%m-%d')
            
            # Convertimos a formato Excel en memoria para la descarga
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_maestro_nuevo.to_excel(writer, sheet_name='Clientes', index=False)
                # Nota técnica: Para no romper el Excel original, lo ideal es copiar esta solapa o descargar el bloque unificado.
            
            st.success("¡Estrategia procesada! Hacé clic abajo para descargar tu nueva solapa de clientes y pegarla en tu archivo de OneDrive:")
            st.download_button(
                label="📥 Descargar Clientes Actualizados.xlsx",
                data=output.getvalue(),
                file_name="Clientes_Actualizados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tab2:
        st.subheader("👥 Maestro de Clientes Completo")
        
        # Preparamos las columnas visuales limpias
        df_clientes_vista = df_clientes.copy()
        df_clientes_vista['Actualizacion'] = df_clientes_vista['Actualizacion_Str']
        columnas_maestro_vis = ['Nro. Doc. Receptor', 'Denominación Receptor', 'Formalidad', 'Periodicidad', 'precio', 'Estado', 'Actualiza', 'Actualizacion', 'periodos', 'Meses Desactualizado', 'Alerta_Revisión']
        
        df_estilado = df_clientes_vista[columnas_maestro_vis].style.apply(colorear_clientes, axis=1)
        
        st.dataframe(
            df_estilado, 
            use_container_width=True,
            column_config={
                "precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
                "periodos": st.column_config.NumberColumn("Período Revisión (Meses)", format="%d"),
                "Meses Desactualizado": st.column_config.NumberColumn("Meses Desactualizado", format="%d")
            }
        )
        
    with tab3:
        st.subheader("🧾 Historial de Facturación (Valores a Plata de Hoy)")
        st.dataframe(
            df_facturacion_completa, 
            use_container_width=True,
            column_config={
                "Facturacion $": st.column_config.NumberColumn("Facturación Original", format="$ %.2f"),
                "Facturacion $ Actualizada": st.column_config.NumberColumn("Facturación Actualizada", format="$ %.2f")
            }
        )
        
    with tab4:
        st.subheader("📊 Índices de Referencia (IPC / IPIM)")
        st.dataframe(df_indices, use_container_width=True)

except Exception as e:
    st.error(f"Ocurrió un error al procesar los datos: {e}")
