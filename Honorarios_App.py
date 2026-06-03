import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
import plotly.express as px

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide")
st.title("📊 Panel de Control de Facturación y Honorarios")

# Ruta del archivo unificado en el servidor
ARCHIVO_LOCAL = "Honosrario NM.xlsx"

# --- CONFIGURACIÓN DE ESCALAS DE MONOTRIBUTO (ACTUALIZAR SEGÚN AFIP) ---
# Estos son los topes de facturación anual NOMINAL por categoría (Ejemplo: Escala 2024 para Servicios)
ESCALAS_MONOTRIBUTO = {
    'A': 6450000,
    'B': 9450000,
    'C': 13250000,
    'D': 16450000,
    'E': 19350000,
    'F': 24250000,
    'G': 29000000,
    'H': 44000000,
    'I': 49115000,  # Límite superior para servicios suele ser menor al de ventas, ajustar según corresponda
    'J': 56400000,
    'K': 68000000
}

# Función auxiliar para abreviar números grandes (Millones y Miles)
def formato_abreviado(valor):
    if valor >= 1_000_000:
        return f"$ {valor / 1_000_000:.2f} M"
    elif valor >= 1_000:
        return f"$ {valor / 1_000:.1f} k"
    else:
        return f"$ {valor:.2f}"

# 2. Motor de carga y cálculo de actualización
def cargar_y_procesar_datos(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None, None, None, None, None, None, None, None, None

    # Leemos las tres hojas
    df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
    df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
    df_indices = pd.read_excel(ruta_archivo, sheet_name="Indices")
    
    # Limpieza preventiva
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns = df_indices.columns.str.strip()
    
    # Limpieza de índices
    df_indices = df_indices.loc[:, ~df_indices.columns.str.contains('^Unnamed')]
    columnas_indices_validas = ['MES', 'IPC  IPIM']
    df_indices = df_indices[[col for col in columnas_indices_validas if col in df_indices.columns]]
    
    # Forzamos numéricos
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Facturacion $'] = pd.to_numeric(df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce')
    if 'precio' in df_clientes.columns:
        df_clientes['precio'] = pd.to_numeric(df_clientes['precio'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
    
    # Fechas internas
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'], errors='coerce')
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()
    
    # Índices
    df_indices_ordenado = df_indices.sort_values('MES_dt')
    ultimo_mes = df_indices_ordenado.iloc[-1]['MES_dt']
    ultimo_indice = df_indices_ordenado.iloc[-1]['IPC  IPIM']
    
    # Cruce facturas vs índices
    df_res = pd.merge(df_facturas, df_indices[['MES_dt', 'IPC  IPIM']], left_on='Mes_Indice', right_on='MES_dt', how='left')
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Coeficiente'] = df_res['Coeficiente'].fillna(1.0)
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    # Cruce final con clientes
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    df_historial_visual = df_final[['Fecha_dt', 'Mes_Indice', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']].copy()
    
    # Clientes proc
    df_clientes_proc = df_clientes.copy()
    df_clientes_proc['Actualizacion_dt'] = pd.to_datetime(df_clientes_proc['Actualizacion'], errors='coerce')
    df_clientes_proc['periodos'] = pd.to_numeric(df_clientes_proc['periodos'], errors='coerce').fillna(0).astype(int)
    
    df_clientes_proc['Mes_Actualizacion_dt'] = df_clientes_proc['Actualizacion_dt'].dt.to_period('M').dt.to_timestamp()
    df_clientes_proc = pd.merge(
        df_clientes_proc, df_indices[['MES_dt', 'IPC  IPIM']], left_on='Mes_Actualizacion_dt', right_on='MES_dt', how='left'
    ).rename(columns={'IPC  IPIM': 'IPC_Ult_Actualizacion'}).drop(columns=['MES_dt', 'Mes_Actualizacion_dt'])
    
    hoy = datetime.now()
    def calcular_metricas_comerciales(row):
        if pd.isna(row['Actualizacion_dt']):
            return pd.Series([0, "OK", row['precio']])
        meses_antigüedad = (hoy.year - row['Actualizacion_dt'].year) * 12 + (hoy.month - row['Actualizacion_dt'].month)
        if pd.notna(row['IPC_Ult_Actualizacion']) and row['IPC_Ult_Actualizacion'] > 0:
            coef_inflacion = ultimo_indice / row['IPC_Ult_Actualizacion']
            sugerido = row['precio'] * coef_inflacion
        else:
            sugerido = row['precio']
        alerta = "VENCIDO" if (row['Actualiza'] == 'Si' and row['Estado'] == 'Activo' and meses_antigüedad >= row['periodos']) else "OK"
        return pd.Series([meses_antigüedad, alerta, sugerido])
    
    df_clientes_proc[['Meses Desactualizado', 'Alerta_Revisión', 'Honorario Sugerido']] = df_clientes_proc.apply(calcular_metricas_comerciales, axis=1)
    df_clientes_proc['Meses Desactualizado'] = df_clientes_proc['Meses Desactualizado'].astype(int)
    
    df_clientes_proc['Orden_Estado'] = df_clientes_proc['Estado'].apply(lambda x: 0 if x == 'Activo' else 1)
    df_clientes_proc = df_clientes_proc.sort_values(by=['Orden_Estado', 'precio'], ascending=[True, False]).drop(columns=['Orden_Estado'])
    df_clientes_proc['Actualizacion_Str'] = df_clientes_proc['Actualizacion_dt'].dt.strftime('%m/%Y')
    
    df_indices_visual = df_indices.copy()
    df_indices_visual['MES'] = df_indices_visual['MES_dt'].dt.strftime('%m/%Y')
    df_indices_visual = df_indices_visual.drop(columns=['MES_dt'])
    
    return df_historial_visual, df_clientes_proc, df_indices_visual, ultimo_mes, ultimo_indice, df_clientes, df_facturas, df_indices, df_final

def colorear_clientes(row):
    estilos = [''] * len(row)
    if row['Estado'] == 'Inactivo':
        return ['background-color: #fee2e2; color: #991b1b; opacity: 0.7;'] * len(row)
    if row['Alerta_Revisión'] == 'VENCIDO':
        idx_alerta = row.index.get_loc('Alerta_Revisión')
        estilos[idx_alerta] = 'background-color: #fef08a; color: #854d0e; font-weight: bold;'
    return estilos

# --- BLOQUE PRINCIPAL DE EJECUCIÓN ---
df_historial_base, df_clientes, df_indices_vis, ult_mes, ult_ind, df_clientes_orig, df_facturas_orig, df_indices_orig, df_motor_interno = cargar_y_procesar_datos(ARCHIVO_LOCAL)

# --- MENÚ LATERAL (BARRA LATERAL DE CONTROL) ---
with st.sidebar:
    st.header("📁 Control del Sistema")
    
    # Subida de archivo maestro
    archivo_subido = st.file_uploader("Actualizar Excel maestro", type=["xlsx"])
    if archivo_subido is not None:
        with open(ARCHIVO_LOCAL, "wb") as f:
            f.write(archivo_subido.getbuffer())
        st.success("¡Archivo maestro actualizado!")
        st.cache_data.clear()
        st.rerun()
        
    st.divider()
    
    # --- ALERTA DE HONORARIOS VENCIDOS ---
    if df_clientes is not None:
        clientes_vencidos = df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO']
        if not clientes_vencidos.empty:
            st.error(f"⚠️ {len(clientes_vencidos)} Honorarios a Renovar")
            with st.popover("Ver detalles de clientes"):
                st.write("**Abonos desactualizados:**")
                for _, row in clientes_vencidos.iterrows():
                    st.write(f"- 🔴 {row['Denominación Receptor']} *(Hace {row['Meses Desactualizado']} meses)*")
                st.info("👉 Entrá a la solapa **'🛠️ Simulador y Ajustes'** para actualizar los valores.")
        else:
            st.success("✅ Todos los honorarios están al día.")
            
    st.divider()

    # --- FILTRO TEMPORAL ---
    if df_historial_base is not None:
        st.header("⏳ Filtro Temporal")
        meses_disponibles = sorted(df_historial_base['Mes_Indice'].dropna().unique())
        if meses_disponibles:
            opciones_fechas = [m.strftime('%m/%Y') for m in meses_disponibles]
            rango_seleccionado = st.select_slider(
                "Seleccioná el rango de análisis:",
                options=opciones_fechas,
                value=(opciones_fechas[0], opciones_fechas[-1])
            )
            fecha_inicio_filtro = pd.to_datetime(rango_seleccionado[0], format='%m/%Y')
            fecha_fin_filtro = pd.to_datetime(rango_seleccionado[1], format='%m/%Y')
        else:
            fecha_inicio_filtro, fecha_fin_filtro = None, None
    else:
        fecha_inicio_filtro, fecha_fin_filtro = None, None

# --- RENDERIZADO CON FILTRADO APLICADO ---
if df_historial_base is not None:
    if fecha_inicio_filtro and fecha_fin_filtro:
        df_motor_filtrado = df_motor_interno[
            (df_motor_interno['Mes_Indice'] >= fecha_inicio_filtro) & 
            (df_motor_interno['Mes_Indice'] <= fecha_fin_filtro)
        ].copy()
        df_historial_filtrado = df_historial_base[
            (df_historial_base['Mes_Indice'] >= fecha_inicio_filtro) & 
            (df_historial_base['Mes_Indice'] <= fecha_fin_filtro)
        ].copy()
    else:
        df_motor_filtrado = df_motor_interno.copy()
        df_historial_filtrado = df_historial_base.copy()

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 Estadísticas",
        "🛠️ Simulador y Ajustes", 
        "👥 Clientes", 
        "🧾 Facturas", 
        "📊 Índices",
        "🏛️ Control Monotributo"  # NUEVA PESTAÑA
    ])
    
    with tab1:
        st.subheader("📊 Análisis Evolutivo Contable por Cliente y Período")
        if len(df_motor_filtrado) == 0:
            st.warning("No hay registros para el rango de tiempo seleccionado.")
        else:
            total_nominal = df_motor_filtrado['Facturacion $'].sum()
            total_actualizado = df_motor_filtrado['Facturacion $ Actualizada'].sum()
            mes_maximo_real = df_motor_interno['Mes_Indice'].max()
            mes_hace_12_real = mes_maximo_real - pd.DateOffset(months=11)
            
            fact_ult_mes = df_motor_interno[df_motor_interno['Mes_Indice'] == mes_maximo_real]['Facturacion $ Actualizada'].sum()
            fact_ult_12m = df_motor_interno[df_motor_interno['Mes_Indice'] >= mes_hace_12_real]['Facturacion $ Actualizada'].sum()
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Nominal en Rango", formato_abreviado(total_nominal))
            col2.metric("Real en Rango (A hoy)", formato_abreviado(total_actualizado))
            col3.metric("Últimos 12 Meses (Estudio Real)", formato_abreviado(fact_ult_12m))
            col4.metric(f"Último Mes Real ({mes_maximo_real.strftime('%m/%Y')})", formato_abreviado(fact_ult_mes))
            
            st.divider()
            
            # Gráfico Líneas
            df_motor_filtrado['Año-Mes'] = df_motor_filtrado['Mes_Indice'].dt.strftime('%Y-%m')
            df_evolucion_mensual = df_motor_filtrado.groupby('Año-Mes')[['Facturacion $', 'Facturacion $ Actualizada']].sum().reset_index()
            df_evolucion_mensual.rename(columns={'Facturacion $': 'Nominal Histórica', 'Facturacion $ Actualizada': 'Real Indexada'}, inplace=True)
            
            fig_linea = px.line(df_evolucion_mensual, x='Año-Mes', y=['Nominal Histórica', 'Real Indexada'],
                                labels={'value': 'Importe', 'variable': 'Tipo', 'Año-Mes': 'Mes'}, color_discrete_sequence=['#636EFA', '#00CC96'])
            fig_linea.update_layout(yaxis_tickformat="$.2s", hovermode="x unified")
            fig_linea.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_linea, use_container_width=True)
            
            st.divider()
            
            # Gráfico Barras
            df_ranking_clientes = df_motor_filtrado.groupby('Denominación Receptor')['Facturacion $ Actualizada'].sum().reset_index()
            df_ranking_clientes = df_ranking_clientes.sort_values(by='Facturacion $ Actualizada', ascending=True)
            fig_barras = px.bar(df_ranking_clientes, x='Facturacion $ Actualizada', y='Denominación Receptor', orientation='h',
                                labels={'Facturacion $ Actualizada': 'Total Facturado Real', 'Denominación Receptor': 'Cliente'}, color_discrete_sequence=['#AB63FA'])
            fig_barras.update_layout(xaxis_tickformat="$.2s", height=600)
            fig_barras.update_traces(hovertemplate="%{x:$,.2f}")
            st.plotly_chart(fig_barras, use_container_width=True)

    with tab2:
        st.subheader("🛠️ Entorno Interactivo de Actualización de Abonos")
        st.write("Modificá los valores en **'Nuevo Precio Pactado'**. Al guardar, se actualizará directamente la planilla base.")
        
        df_simulacion = df_clientes[df_clientes['Estado'] == 'Activo'].copy()
        df_simulacion['Nuevo Precio Pactado'] = df_simulacion['precio']
        columnas_sim = ['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido', 'Nuevo Precio Pactado']
        
        df_editado = st.data_editor(
            df_simulacion[columnas_sim], use_container_width=True,
            disabled=['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido'],
            column_config={
                "precio": st.column_config.NumberColumn("Precio Actual", format="$ %.2f"),
                "Honorario Sugerido": st.column_config.NumberColumn("Sugerido por IPC", format="$ %.2f"),
                "Nuevo Precio Pactado": st.column_config.NumberColumn("Nuevo Precio Pactado ✏️", format="$ %.2f"),
                "Meses Desactualizado": st.column_config.NumberColumn("Meses Inmóvil", format="%d")
            }, key="editor_abonos"
        )
        
        if st.button("💾 Guardar y Actualizar Base de Datos"):
            df_maestro_nuevo = df_clientes_orig.copy()
            cambios_realizados = 0
            for idx, row in df_editado.iterrows():
                cuit = row['Nro. Doc. Receptor']
                nuevo_val = row['Nuevo Precio Pactado']
                if nuevo_val != row['precio']:
                    df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'precio'] = nuevo_val
                    df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'Actualizacion'] = datetime.today().strftime('%Y-%m-%d')
                    cambios_realizados += 1
            if cambios_realizados > 0:
                with pd.ExcelWriter(ARCHIVO_LOCAL, engine='openpyxl') as writer:
                    df_maestro_nuevo.to_excel(writer, sheet_name='Clientes', index=False)
                    df_facturas_orig.to_excel(writer, sheet_name='Facturas', index=False)
                    df_indices_orig.to_excel(writer, sheet_name='Indices', index=False)
                st.success(f"¡Se actualizaron con éxito {cambios_realizados} clientes!")
                st.cache_data.clear()
                st.rerun()
            else:
                st.info("No se detectaron cambios.")

    with tab3:
        st.subheader("👥 Maestro de Clientes Completo")
        df_clientes_vista = df_clientes.copy()
        df_clientes_vista['Actualizacion'] = df_clientes_vista['Actualizacion_Str']
        columnas_maestro_vis = ['Nro. Doc. Receptor', 'Denominación Receptor', 'Formalidad', 'Periodicidad', 'precio', 'Estado', 'Actualiza', 'Actualizacion', 'periodos', 'Meses Desactualizado', 'Alerta_Revisión']
        df_estilado = df_clientes_vista[columnas_maestro_vis].style.apply(colorear_clientes, axis=1)
        st.dataframe(df_estilado, use_container_width=True, column_config={
            "precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
            "periodos": st.column_config.NumberColumn("Período Revisión", format="%d"),
            "Meses Desactualizado": st.column_config.NumberColumn("Meses Inmóvil", format="%d")
        })
        
    with tab4:
        st.subheader("🧾 Historial de Facturación Filtrado")
        df_historial_render = df_historial_filtrado.copy()
        df_historial_render['Fecha'] = df_historial_render['Fecha_dt'].dt.strftime('%d/%m/%Y')
        columnas_tabla = ['Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
        st.dataframe(df_historial_render[columnas_tabla], use_container_width=True, column_config={
            "Facturacion $": st.column_config.NumberColumn("Facturación Original", format="$ %.2f"),
            "Facturacion $ Actualizada": st.column_config.NumberColumn("Facturación Actualizada", format="$ %.2f")
        })
        
    with tab5:
        st.subheader("📊 Índices de Referencia (IPC / IPIM)")
        st.dataframe(df_indices_vis, use_container_width=True)
        
    with tab6:
        # --- NUEVA SECCIÓN DE CONTROL DE MONOTRIBUTO ---
        st.subheader("🏛️ Panel de Recategorización de Monotributo")
        st.write("AFIP controla la facturación **Nominal** (sin ajustar por inflación) de los últimos 12 meses móviles para determinar tu categoría.")
        
        # Obtenemos la facturación NOMINAL de los últimos 365 días reales registrados
        fecha_max_factura = df_motor_interno['Fecha_dt'].max()
        fecha_hace_un_año = fecha_max_factura - pd.DateOffset(years=1)
        
        # Filtramos facturas de los últimos 12 meses y sumamos la facturación original
        df_ultimos_12 = df_motor_interno[(df_motor_interno['Fecha_dt'] > fecha_hace_un_año) & (df_motor_interno['Fecha_dt'] <= fecha_max_factura)]
        facturacion_nominal_12m = df_ultimos_12['Facturacion $'].sum()
        
        # Determinamos la categoría teórica según la escala definida
        categoria_proyectada = "Excluido"
        tope_limite = 0
        for cat, limite in ESCALAS_MONOTRIBUTO.items():
            if facturacion_nominal_12m <= limite:
                categoria_proyectada = cat
                tope_limite = limite
                break
                
        # Interfaz del usuario
        col1, col2 = st.columns(2)
        
        with col1:
            st.info(f"📅 **Período evaluado:** {fecha_hace_un_año.strftime('%d/%m/%Y')} al {fecha_max_factura.strftime('%d/%m/%Y')}")
            st.metric("Facturación Nominal Últimos 12 Meses", f"$ {facturacion_nominal_12m:,.2f}")
            st.metric("Categoría Proyectada AFIP", f"Categoría {categoria_proyectada}")
            if categoria_proyectada != "Excluido":
                margen_restante = tope_limite - facturacion_nominal_12m
                st.caption(f"*Tenés un margen de $ {margen_restante:,.2f} antes de saltar a la próxima categoría.*")
        
        with col2:
            st.write("### Tu Situación Actual")
            mi_categoria = st.selectbox("¿En qué categoría estás inscripto hoy?", options=list(ESCALAS_MONOTRIBUTO.keys()))
            
            if categoria_proyectada == "Excluido":
                st.error("🚨 **¡ALERTA DE EXCLUSIÓN!** Has superado el tope máximo del Régimen Simplificado (Categoría K).")
            elif mi_categoria == categoria_proyectada:
                st.success("✅ **Bien categorizado.** En la próxima recategorización (Enero/Julio) deberías mantenerte en la misma letra.")
            elif list(ESCALAS_MONOTRIBUTO.keys()).index(categoria_proyectada) > list(ESCALAS_MONOTRIBUTO.keys()).index(mi_categoria):
                st.warning(f"⬆️ **Toca subir.** En la próxima recategorización vas a tener que subir a la **Categoría {categoria_proyectada}**.")
            else:
                st.info(f"⬇️ **Podés bajar.** Tu facturación bajó lo suficiente como para recategorizarte en la **Categoría {categoria_proyectada}** y pagar menos.")
                
        st.divider()
        st.caption("⚠️ *Aviso técnico:* Los topes de facturación de AFIP (Escalas) cambian por ley. Para mantener este panel preciso, acordate de actualizar los montos de la variable `ESCALAS_MONOTRIBUTO` en las primeras líneas del código fuente de la aplicación cada vez que AFIP publique las nuevas tablas.")

else:
    st.warning("⚠️ Todavía no hay ninguna base de datos activa. Usa el menú lateral para subir tu archivo 'Honosrario NM.xlsx'.")
