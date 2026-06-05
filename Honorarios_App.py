import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide", initial_sidebar_state="expanded")
st.title("📊 Panel de Control de Facturación y Honorarios")

ARCHIVO_LOCAL = "Honosrario NM.xlsx"

if 'cat_agustin' not in st.session_state:
    st.session_state['cat_agustin'] = 'D'
if 'cat_laura' not in st.session_state:
    st.session_state['cat_laura'] = 'B'

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

@st.cache_data(ttl=300)
def cargar_y_procesar_datos(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None, None, None, None, None, None, None, None, None

    df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
    df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
    df_indices = pd.read_excel(ruta_archivo, sheet_name="Indices")
    
    # --- LIMPIEZA ABSOLUTA DE ENCABEZADOS ---
    # Esto elimina caracteres fantasmas y espacios en todas las tablas sin renombrar a la fuerza
    df_facturas.columns = [str(c).replace('\ufeff', '').strip() for c in df_facturas.columns]
    df_clientes.columns = [str(c).replace('\ufeff', '').strip() for c in df_clientes.columns]
    df_indices.columns = [str(c).replace('\ufeff', '').strip() for c in df_indices.columns]
    
    # Eliminamos columnas repetidas si el Excel dejó alguna basura
    df_facturas = df_facturas.loc[:, ~df_facturas.columns.duplicated()]
    df_clientes = df_clientes.loc[:, ~df_clientes.columns.duplicated()]
    df_indices = df_indices.loc[:, ~df_indices.columns.duplicated()]
    
    df_indices = df_indices.loc[:, ~df_indices.columns.str.contains('^Unnamed')]
    columnas_indices_validas = ['MES', 'IPC  IPIM']
    df_indices = df_indices[[col for col in columnas_indices_validas if col in df_indices.columns]]
    
    df_indices['IPC  IPIM'] = pd.to_numeric(df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce')
    df_facturas['Facturacion $'] = pd.to_numeric(df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce')
    if 'precio' in df_clientes.columns:
        df_clientes['precio'] = pd.to_numeric(df_clientes['precio'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
    
    df_facturas['Fecha_dt'] = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt'] = pd.to_datetime(df_indices['MES'], errors='coerce')
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()
    
    df_indices_ordenado = df_indices.sort_values('MES_dt')
    ultimo_mes = df_indices_ordenado.iloc[-1]['MES_dt']
    ultimo_indice = df_indices_ordenado.iloc[-1]['IPC  IPIM']
    
    df_res = pd.merge(df_facturas, df_indices[['MES_dt', 'IPC  IPIM']], left_on='Mes_Indice', right_on='MES_dt', how='left')
    df_res['Coeficiente'] = ultimo_indice / df_res['IPC  IPIM']
    df_res['Coeficiente'] = df_res['Coeficiente'].fillna(1.0)
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']
    
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')
    
    # Tabla Historial Visual
    cols_hist = ['Fecha_dt', 'Mes_Indice', 'Emisor', 'Negocio', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    df_historial_visual = df_final[[col for col in cols_hist if col in df_final.columns]].copy()
    
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

if df_historial_base is None:
    st.error(f"⚠️ No se encontró el archivo '{ARCHIVO_LOCAL}'. Asegurate de guardarlo en la carpeta del sistema.")
    st.stop()

# --- PRE-CÁLCULO DE SEMÁFOROS ---
estado_honorarios = "🟢"
if df_clientes is not None and not df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO'].empty:
    estado_honorarios = "🔴"

estado_afip = "🟢"
if df_historial_base is not None:
    fecha_max = df_motor_interno['Fecha_dt'].max()
    fecha_12m = fecha_max - pd.DateOffset(years=1)
    fecha_3m = fecha_max - pd.DateOffset(months=3)
    
    def nivel_alerta_afip(nombre_emisor, cat_actual):
        df_emisor = df_motor_interno[df_motor_interno['Emisor'] == nombre_emisor]
        fact_12m = df_emisor[(df_emisor['Fecha_dt'] > fecha_12m) & (df_emisor['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
        fact_3m = df_emisor[(df_emisor['Fecha_dt'] > fecha_3m) & (df_emisor['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
        promedio = fact_3m / 3 if fact_3m > 0 else fact_12m / 12

        margen = ESCALAS_MONOTRIBUTO[cat_actual] - fact_12m
        if margen < 0: return 3
        elif margen <= promedio: return 2
        else: return 1

    nivel_ag = nivel_alerta_afip('Agustín', st.session_state['cat_agustin'])
    nivel_la = nivel_alerta_afip('Laura', st.session_state['cat_laura'])
    max_alerta = max(nivel_ag, nivel_la)
    
    if max_alerta == 3: estado_afip = "🔴"
    elif max_alerta == 2: estado_afip = "🟡"

# --- MENÚ LATERAL DE NAVEGACIÓN ---
with st.sidebar:
    st.header("Navegación del Sistema")
    MENU_ESTADISTICAS = "📈 Estudio: Estadísticas"
    MENU_CYGNUS = "🏠 Cygnus Home: KPIs"
    MENU_CLIENTES_ANALISIS = f"{estado_honorarios} Análisis de Clientes"
    MENU_AFIP = f"{estado_afip} Control Monotributo"
    MENU_FACTURAS = "🧾 Historial de Facturas"
    MENU_INDICES = "📊 Tabla de Índices"
    
    opciones_menu = [MENU_ESTADISTICAS, MENU_CYGNUS, MENU_CLIENTES_ANALISIS, MENU_AFIP, MENU_FACTURAS, MENU_INDICES]
    seleccion_pantalla = st.radio("Ir a sección:", opciones_menu, label_visibility="collapsed")
    
    st.divider()
    
    # Botón de refresco (Al pulsarlo, borra el caché y lee el archivo del disco en el acto)
    if st.button("🔄 Refrescar datos del archivo", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    st.header("⏳ Filtro Temporal")
    meses_disponibles = sorted(df_historial_base['Mes_Indice'].dropna().unique())
    if meses_disponibles:
        opciones_fechas = [m.strftime('%m/%Y') for m in meses_disponibles]
        rango_seleccionado = st.select_slider("Rango de análisis:", options=opciones_fechas, value=(opciones_fechas[0], opciones_fechas[-1]), label_visibility="collapsed")
        fecha_inicio_filtro = pd.to_datetime(rango_seleccionado[0], format='%m/%Y')
        fecha_fin_filtro = pd.to_datetime(rango_seleccionado[1], format='%m/%Y')
    else:
        fecha_inicio_filtro, fecha_fin_filtro = None, None

# --- APLICACIÓN DEL FILTRO TEMPORAL ---
if fecha_inicio_filtro and fecha_fin_filtro:
    df_motor_filtrado = df_motor_interno[(df_motor_interno['Mes_Indice'] >= fecha_inicio_filtro) & (df_motor_interno['Mes_Indice'] <= fecha_fin_filtro)].copy()
    df_historial_filtrado = df_historial_base[(df_historial_base['Mes_Indice'] >= fecha_inicio_filtro) & (df_historial_base['Mes_Indice'] <= fecha_fin_filtro)].copy()
else:
    df_motor_filtrado = df_motor_interno.copy()
    df_historial_filtrado = df_historial_base.copy()

# --- RENDERIZADO DE PANTALLAS ---

if seleccion_pantalla == MENU_ESTADISTICAS:
    st.subheader("📊 Estudio Contable: Análisis Evolutivo Real")
    st.caption("Esta vista procesa exclusivamente las facturas de 'Estudio'.")
    
    df_estudio_filtrado = df_motor_filtrado[df_motor_filtrado['Negocio'] == 'Estudio']
    df_estudio_interno = df_motor_interno[df_motor_interno['Negocio'] == 'Estudio']
    
    total_nominal = df_estudio_filtrado['Facturacion $'].sum()
    total_actualizado = df_estudio_filtrado['Facturacion $ Actualizada'].sum()
    mes_maximo_real = df_estudio_interno['Mes_Indice'].max()
    
    if pd.notna(mes_maximo_real):
        mes_hace_12_real = mes_maximo_real - pd.DateOffset(months=11)
        fact_ult_mes = df_estudio_interno[df_estudio_interno['Mes_Indice'] == mes_maximo_real]['Facturacion $ Actualizada'].sum()
        fact_ult_12m = df_estudio_interno[df_estudio_interno['Mes_Indice'] >= mes_hace_12_real]['Facturacion $ Actualizada'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Nominal en Rango", formato_abreviado(total_nominal))
        col2.metric("Real en Rango", formato_abreviado(total_actualizado))
        col3.metric("Últimos 12 Meses (Real)", formato_abreviado(fact_ult_12m))
        col4.metric(f"Último Mes ({mes_maximo_real.strftime('%m/%Y')})", formato_abreviado(fact_ult_mes))
        
        st.divider()
        if not df_estudio_filtrado.empty:
            df_estudio_filtrado['Año-Mes'] = df_estudio_filtrado['Mes_Indice'].dt.strftime('%Y-%m')
            df_evolucion_mensual = df_estudio_filtrado.groupby('Año-Mes')[['Facturacion $', 'Facturacion $ Actualizada']].sum().reset_index()
            df_evolucion_mensual.rename(columns={'Facturacion $': 'Nominal Histórica', 'Facturacion $ Actualizada': 'Real Indexada'}, inplace=True)
            fig_linea = px.line(df_evolucion_mensual, x='Año-Mes', y=['Nominal Histórica', 'Real Indexada'], color_discrete_sequence=['#636EFA', '#00CC96'])
            fig_linea.update_layout(yaxis_tickformat="$.2s", hovermode="x unified")
            fig_linea.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_linea, use_container_width=True)
            
            st.divider()
            df_ranking_clientes = df_estudio_filtrado.groupby('Denominación Receptor')['Facturacion $ Actualizada'].sum().reset_index().sort_values(by='Facturacion $ Actualizada', ascending=True)
            fig_barras = px.bar(df_ranking_clientes, x='Facturacion $ Actualizada', y='Denominación Receptor', orientation='h', color_discrete_sequence=['#AB63FA'])
            fig_barras.update_layout(xaxis_tickformat="$.2s", height=600)
            fig_barras.update_traces(hovertemplate="%{x:$,.2f}")
            st.plotly_chart(fig_barras, use_container_width=True)

elif seleccion_pantalla == MENU_CYGNUS:
    st.subheader("🏠 Cygnus Home: Rendimiento Comercial")
    
    df_cygnus_filtrado = df_motor_filtrado[df_motor_filtrado['Negocio'] == 'Cygnus Home']
    df_cygnus_interno = df_motor_interno[df_motor_interno['Negocio'] == 'Cygnus Home']
    
    if df_cygnus_interno.empty:
        st.info("Aún no hay facturas registradas en la base de datos para Cygnus Home.")
    else:
        mes_curso = df_cygnus_interno['Mes_Indice'].max()
        mes_anterior = mes_curso - pd.DateOffset(months=1)
        f_12m = mes_curso - pd.DateOffset(months=11)
        
        fact_mes_curso = df_cygnus_interno[df_cygnus_interno['Mes_Indice'] == mes_curso]['Facturacion $ Actualizada'].sum()
        fact_mes_anterior = df_cygnus_interno[df_cygnus_interno['Mes_Indice'] == mes_anterior]['Facturacion $ Actualizada'].sum()
        fact_ult_12m = df_cygnus_interno[df_cygnus_interno['Mes_Indice'] >= f_12m]['Facturacion $ Actualizada'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric(f"Mes en Curso ({mes_curso.strftime('%m/%Y')})", f"$ {fact_mes_curso:,.2f}")
        col2.metric(f"Mes Anterior ({mes_anterior.strftime('%m/%Y')})", f"$ {fact_mes_anterior:,.2f}")
        col3.metric("Últimos 12 Meses (Real)", formato_abreviado(fact_ult_12m))
        
        st.divider()
        if not df_cygnus_filtrado.empty:
            df_cygnus_filtrado['Año-Mes'] = df_cygnus_filtrado['Mes_Indice'].dt.strftime('%Y-%m')
            df_evo_cy = df_cygnus_filtrado.groupby('Año-Mes')[['Facturacion $ Actualizada']].sum().reset_index()
            fig_cy = px.line(df_evo_cy, x='Año-Mes', y='Facturacion $ Actualizada', color_discrete_sequence=['#FF7F0E'])
            fig_cy.update_layout(title="Evolución Real de Cygnus Home", yaxis_tickformat="$.2s", hovermode="x unified")
            fig_cy.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_cy, use_container_width=True)

elif seleccion_pantalla == MENU_CLIENTES_ANALISIS:
    st.subheader("👥 Análisis de Clientes y Ajustes")
    
    clientes_vencidos = df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO']
    if not clientes_vencidos.empty:
        st.error(f"⚠️ **Atención:** Hay **{len(clientes_vencidos)}** honorarios sugeridos para actualización.")
    else:
        st.success("✅ Todos los clientes activos están con sus honorarios al día.")
    
    st.write("#### 🛠️ Ajuste Rápido de Honorarios")
    df_simulacion = df_clientes[df_clientes['Estado'] == 'Activo'].copy()
    df_simulacion['Nuevo Precio Pactado'] = df_simulacion['precio']
    columnas_sim = ['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido', 'Nuevo Precio Pactado']
    
    df_editado = st.data_editor(
        df_simulacion[columnas_sim], use_container_width=True,
        disabled=['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido'],
        column_config={
            "precio": st.column_config.NumberColumn("Precio Actual", format="$ %.2f"),
            "Honorario Sugerido": st.column_config.NumberColumn("Sugerido por IPC", format="$ %.2f"),
            "Nuevo Precio Pactado": st.column_config.NumberColumn("Nuevo Precio Pactado ✏️", format="$ %.2f")
        }, key="editor_abonos"
    )
    
    if st.button("💾 Guardar y Actualizar Base de Datos", type="primary"):
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
            st.success(f"¡Se actualizaron con éxito {cambios_realizados} clientes! Refrescando...")
            st.cache_data.clear()
            st.rerun()

    st.divider()
    st.write("#### 📌 Vista General Maestra")
    df_clientes_vista = df_clientes.copy()
    df_clientes_vista['Actualizacion'] = df_clientes_vista['Actualizacion_Str']
    columnas_maestro_vis = ['Nro. Doc. Receptor', 'Denominación Receptor', 'Formalidad', 'Periodicidad', 'precio', 'Estado', 'Actualiza', 'Actualizacion', 'periodos', 'Meses Desactualizado', 'Alerta_Revisión']
    df_estilado = df_clientes_vista[columnas_maestro_vis].style.apply(colorear_clientes, axis=1)
    st.dataframe(df_estilado, use_container_width=True, column_config={"precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f")})

elif seleccion_pantalla == MENU_AFIP:
    st.subheader("🏛️ Panel de Recategorización e Inscripción Activa")
    
    mes_actual = datetime.now().month
    if mes_actual in [6, 12]:
        st.warning("⚠️ **Recordatorio de Agenda:** El mes que viene inicia el período de recategorización obligatoria de AFIP.")
    elif mes_actual in [1, 7]:
        st.error("🚨 **Período de Recategorización Activo:** Tenés tiempo hasta el día 20 de este mes para confirmar tu categoría en la web de AFIP.")
        
    def renderizar_afip(nombre_emisor, cat_actual, col):
        with col:
            st.write(f"### 👤 {nombre_emisor}")
            df_emisor = df_motor_interno[df_motor_interno['Emisor'] == nombre_emisor]
            if df_emisor.empty:
                st.info(f"Sin facturación registrada para {nombre_emisor}.")
            else:
                fecha_max_factura = df_motor_interno['Fecha_dt'].max()
                fecha_12m = fecha_max_factura - pd.DateOffset(years=1)
                facturacion_12m = df_emisor[(df_emisor['Fecha_dt'] > fecha_12m) & (df_emisor['Fecha_dt'] <= fecha_max_factura)]['Facturacion $'].sum()
                
                cat_sug = "Excluido"
                for c, lim in ESCALAS_MONOTRIBUTO.items():
                    if facturacion_12m <= lim:
                        cat_sug = c
                        break
                st.metric("Facturación Nominal (Últimos 12 Meses)", f"$ {facturacion_12m:,.2f}")
                st.info(f"📍 Categoría AFIP Sugerida: **{cat_sug}** (Actual en App: {cat_actual})")

    col_izq, col_der = st.columns(2)
    renderizar_afip('Agustín', st.session_state['cat_agustin'], col_izq)
    renderizar_afip('Laura', st.session_state['cat_laura'], col_der)
        
    st.divider()
    st.write("### 🛠️ Corrección de Categoría Actual")
    
    col_sel_ag, col_sel_la = st.columns(2)
    with col_sel_ag:
        cat_ag_sel = st.selectbox("Categoría Agustín:", options=list(ESCALAS_MONOTRIBUTO.keys()), index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_agustin']))
        if cat_ag_sel != st.session_state['cat_agustin']:
            st.session_state['cat_agustin'] = cat_ag_sel
            st.rerun()
    with col_sel_la:
        cat_la_sel = st.selectbox("Categoría Laura:", options=list(ESCALAS_MONOTRIBUTO.keys()), index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_laura']))
        if cat_la_sel != st.session_state['cat_laura']:
            st.session_state['cat_laura'] = cat_la_sel
            st.rerun()

elif seleccion_pantalla == MENU_FACTURAS:
    st.subheader("🧾 Historial de Facturación")
    
    vista_facturas = st.radio("Filtrar visualización de facturas:", ["🏢 Estudio Contable", "🏠 Cygnus Home", "🌎 Mostrar Todas las Facturas"], horizontal=True)
    
    if vista_facturas == "🏢 Estudio Contable":
        df_hist_render = df_historial_filtrado[df_historial_filtrado['Negocio'] == 'Estudio'].copy()
    elif vista_facturas == "🏠 Cygnus Home":
        df_hist_render = df_historial_filtrado[df_historial_filtrado['Negocio'] == 'Cygnus Home'].copy()
    else:
        df_hist_render = df_historial_filtrado.copy()
        
    df_hist_render['Fecha'] = df_hist_render['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_hist_render = df_hist_render.sort_values('Fecha_dt', ascending=False)
    
    columnas_tabla = ['Fecha', 'Emisor', 'Negocio', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    
    columnas_existentes = [c for c in columnas_tabla if c in df_hist_render.columns]
    
    st.dataframe(
        df_hist_render[columnas_existentes], 
        use_container_width=True, 
        column_config={
            "Facturacion $": st.column_config.NumberColumn("Original", format="$ %.2f"), 
            "Facturacion $ Actualizada": st.column_config.NumberColumn("A Plata de Hoy", format="$ %.2f")
        }
    )
    
elif seleccion_pantalla == MENU_INDICES:
    st.subheader("📊 Índices de Referencia")
    st.dataframe(df_indices_vis, use_container_width=True)
