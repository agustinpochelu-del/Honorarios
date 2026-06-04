import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
import plotly.express as px

# 1. Configuración de la página
st.set_page_config(page_title="Control de Honorarios", layout="wide", initial_sidebar_state="expanded")
st.title("📊 Panel de Control de Facturación y Honorarios")

ARCHIVO_LOCAL = "Honosrario NM.xlsx"

# Nombres Identificadores a prueba de balas (sin tildes y en mayúsculas)
EMISOR_AGUSTIN = "AGUSTIN"
EMISOR_LAURA = "LAURA"

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

# Función sin caché para que lea SIEMPRE la versión más nueva del Excel
def cargar_y_procesar_datos(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None, None, None, None, None, None, None, None, None

    df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
    df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
    df_indices = pd.read_excel(ruta_archivo, sheet_name="Indices")
    
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns = df_indices.columns.str.strip()
    
    df_indices = df_indices.loc[:, ~df_indices.columns.str.contains('^Unnamed')]
    df_indices = df_indices[['MES', 'IPC  IPIM']]
    
    # --- NUEVA LECTURA BLINDADA POR NOMBRE DE EMISOR ---
    if 'Nombre Emisor' in df_facturas.columns:
        # Forzamos texto, quitamos espacios, pasamos a mayúsculas y volamos la tilde de Agustín
        df_facturas['Nombre Emisor'] = df_facturas['Nombre Emisor'].astype(str).str.strip().str.upper()
        df_facturas['Nombre Emisor'] = df_facturas['Nombre Emisor'].str.replace('Í', 'I').str.replace('í', 'i')
        # Si quedó vacío o dice 'nan', lo asignamos al estudio (Agustín) por defecto
        df_facturas.loc[df_facturas['Nombre Emisor'].isin(['NAN', '', 'NONE', 'NULL']), 'Nombre Emisor'] = EMISOR_AGUSTIN
    else:
        df_facturas['Nombre Emisor'] = EMISOR_AGUSTIN 
        
    if 'Punto de Venta' in df_facturas.columns:
        df_facturas['Punto de Venta'] = df_facturas['Punto de Venta'].astype(str).str.replace(r'\D', '', regex=True)
        
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
    columnas_hist = ['Fecha_dt', 'Mes_Indice', 'Nombre Emisor', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    df_historial_visual = df_final[[c for c in columnas_hist if c in df_final.columns]].copy()
    
    df_clientes_proc = df_clientes.copy()
    df_clientes_proc['Actualizacion_dt'] = pd.to_datetime(df_clientes_proc['Actualizacion'], errors='coerce')
    df_clientes_proc['periodos'] = pd.to_numeric(df_clientes_proc['periodos'], errors='coerce').fillna(0).astype(int)
    
    df_clientes_proc['Mes_Actualizacion_dt'] = df_clientes_proc['Actualizacion_dt'].dt.to_period('M').dt.to_timestamp()
    df_clientes_proc = pd.merge(
        df_clientes_proc, df_indices[['MES_dt', 'IPC  IPIM']], left_on='Mes_Actualizacion_dt', right_on='MES_dt', how='left'
    ).rename(columns={'IPC  IPIM': 'IPC_Ult_Actualizacion'}).drop(columns=['MES_dt', 'Mes_Actualizacion_dt'])
    
    hoy = datetime.now()
    def calcular_metricas_comerciales(row):
        if pd.isna(row['Actualizacion_dt']): return pd.Series([0, "OK", row['precio']])
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
    if row['Estado'] == 'Inactivo': return ['background-color: #fee2e2; color: #991b1b; opacity: 0.7;'] * len(row)
    if row['Alerta_Revisión'] == 'VENCIDO':
        idx_alerta = row.index.get_loc('Alerta_Revisión')
        estilos[idx_alerta] = 'background-color: #fef08a; color: #854d0e; font-weight: bold;'
    return estilos

df_historial_base, df_clientes, df_indices_vis, ult_mes, ult_ind, df_clientes_orig, df_facturas_orig, df_indices_orig, df_motor_interno = cargar_y_procesar_datos(ARCHIVO_LOCAL)

if df_historial_base is None:
    st.error(f"⚠️ No se encontró '{ARCHIVO_LOCAL}'. Asegurate de cargar el archivo base en el servidor.")
    st.stop()

# --- CÁLCULO DE SEMÁFOROS PREVIOS ---
estado_honorarios = "🟢"
if df_clientes is not None and not df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO'].empty:
    estado_honorarios = "🔴"

estado_afip = "🟢"
if df_historial_base is not None:
    fecha_max = df_motor_interno['Fecha_dt'].max()
    fecha_12m = fecha_max - pd.DateOffset(years=1)
    fecha_3m = fecha_max - pd.DateOffset(months=3)
    
    def calc_alert_afip(nombre_emisor, cat_actual):
        df_emisor = df_motor_interno[df_motor_interno['Nombre Emisor'] == nombre_emisor]
        fact_12m = df_emisor[(df_emisor['Fecha_dt'] > fecha_12m) & (df_emisor['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
        fact_3m = df_emisor[(df_emisor['Fecha_dt'] > fecha_3m) & (df_emisor['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
        promedio = fact_3m / 3 if fact_3m > 0 else fact_12m / 12
        
        margen = ESCALAS_MONOTRIBUTO[cat_actual] - fact_12m
        if margen < 0: return 3
        elif margen <= promedio: return 2
        else: return 1

    nivel_ag = calc_alert_afip(EMISOR_AGUSTIN, st.session_state['cat_agustin'])
    nivel_la = calc_alert_afip(EMISOR_LAURA, st.session_state['cat_laura'])
    
    max_alerta = max(nivel_ag, nivel_la)
    if max_alerta == 3: estado_afip = "🔴"
    elif max_alerta == 2: estado_afip = "🟡"

# --- MENÚ LATERAL LIMPIO ---
with st.sidebar:
    st.header("Navegación del Sistema")
    MENU_ESTADISTICAS = "📈 Estudio: Estadísticas"
    MENU_CYGNUS = "🏠 Cygnus Home: KPIs"
    MENU_SIMULADOR = f"{estado_honorarios} Ajustes de Honorarios"
    MENU_AFIP = f"{estado_afip} Control Monotributo"
    MENU_CLIENTES = "👥 Maestro de Clientes"
    MENU_FACTURAS = "🧾 Historial de Facturas"
    MENU_INDICES = "📊 Tabla de Índices"
    
    opciones_menu = [MENU_ESTADISTICAS, MENU_CYGNUS, MENU_SIMULADOR, MENU_AFIP, MENU_CLIENTES, MENU_FACTURAS, MENU_INDICES]
    seleccion_pantalla = st.radio("Ir a:", opciones_menu, label_visibility="collapsed")
    
    st.divider()
    
    st.header("⏳ Filtro Temporal")
    meses_disponibles = sorted(df_historial_base['Mes_Indice'].dropna().unique())
    if meses_disponibles:
        ops = [m.strftime('%m/%Y') for m in meses_disponibles]
        rango_seleccionado = st.select_slider("Rango:", options=ops, value=(ops[0], ops[-1]), label_visibility="collapsed")
        fecha_ini = pd.to_datetime(rango_seleccionado[0], format='%m/%Y')
        fecha_fin = pd.to_datetime(rango_seleccionado[1], format='%m/%Y')
    else:
        fecha_ini, fecha_fin = None, None

    st.divider()
    # Carga manual por las dudas, pero como no hay caché, la app se refresca sola al tocar el menú
    archivo_subido = st.file_uploader("Actualizar Excel maestro", type=["xlsx"], label_visibility="collapsed")
    if archivo_subido is not None:
        with open(ARCHIVO_LOCAL, "wb") as f:
            f.write(archivo_subido.getbuffer())
        st.success("¡Archivo maestro actualizado!")
        st.rerun()

# --- APLICACIÓN DE FILTROS ---
if fecha_ini and fecha_fin:
    df_motor_filtrado = df_motor_interno[(df_motor_interno['Mes_Indice'] >= fecha_ini) & (df_motor_interno['Mes_Indice'] <= fecha_fin)].copy()
    df_hist_filtrado = df_historial_base[(df_historial_base['Mes_Indice'] >= fecha_ini) & (df_historial_base['Mes_Indice'] <= fecha_fin)].copy()
else:
    df_motor_filtrado = df_motor_interno.copy()
    df_hist_filtrado = df_historial_base.copy()

# Filtro estricto para diferenciar unidades de negocio basado en la nueva columna de Nombre
filtro_cygnus = (df_motor_filtrado['Nombre Emisor'] == EMISOR_LAURA) & (df_motor_filtrado['Punto de Venta'] == '2')

if seleccion_pantalla == MENU_ESTADISTICAS:
    st.subheader("📊 Estudio Contable: Evolución Real")
    st.caption("Esta vista excluye automáticamente la facturación correspondiente a Cygnus Home.")
    
    df_estudio = df_motor_filtrado[~filtro_cygnus]
    df_estudio_absoluto = df_motor_interno[~((df_motor_interno['Nombre Emisor'] == EMISOR_LAURA) & (df_motor_interno['Punto de Venta'] == '2'))]
    
    tot_nom = df_estudio['Facturacion $'].sum()
    tot_act = df_estudio['Facturacion $ Actualizada'].sum()
    
    fecha_max = df_estudio_absoluto['Mes_Indice'].max()
    if pd.notna(fecha_max):
        fecha_hace_12 = fecha_max - pd.DateOffset(months=11)
        fact_ult_mes = df_estudio_absoluto[df_estudio_absoluto['Mes_Indice'] == fecha_max]['Facturacion $ Actualizada'].sum()
        fact_ult_12m = df_estudio_absoluto[df_estudio_absoluto['Mes_Indice'] >= fecha_hace_12]['Facturacion $ Actualizada'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Nominal en Rango", formato_abreviado(tot_nom))
        col2.metric("Real en Rango", formato_abreviado(tot_act))
        col3.metric("Últimos 12 Meses (Real)", formato_abreviado(fact_ult_12m))
        col4.metric(f"Último Mes ({fecha_max.strftime('%m/%Y')})", formato_abreviado(fact_ult_mes))
        
        st.divider()
        df_estudio['Año-Mes'] = df_estudio['Mes_Indice'].dt.strftime('%Y-%m')
        df_evo = df_estudio.groupby('Año-Mes')[['Facturacion $', 'Facturacion $ Actualizada']].sum().reset_index()
        fig = px.line(df_evo, x='Año-Mes', y=['Facturacion $', 'Facturacion $ Actualizada'], labels={'value': 'Importe', 'variable': 'Tipo'}, color_discrete_sequence=['#636EFA', '#00CC96'])
        fig.update_layout(yaxis_tickformat="$.2s", hovermode="x unified")
        fig.update_traces(hovertemplate="%{y:$,.2f}")
        st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        df_rk = df_estudio.groupby('Denominación Receptor')['Facturacion $ Actualizada'].sum().reset_index().sort_values(by='Facturacion $ Actualizada', ascending=True)
        fig_bar = px.bar(df_rk, x='Facturacion $ Actualizada', y='Denominación Receptor', orientation='h', color_discrete_sequence=['#AB63FA'])
        fig_bar.update_layout(xaxis_tickformat="$.2s", height=600)
        fig_bar.update_traces(hovertemplate="%{x:$,.2f}")
        st.plotly_chart(fig_bar, use_container_width=True)

elif seleccion_pantalla == MENU_CYGNUS:
    st.subheader("🏠 Cygnus Home: Rendimiento Comercial")
    st.caption("Métricas filtradas exclusivamente para Laura - Punto de Venta 2.")
    
    df_cygnus = df_motor_filtrado[filtro_cygnus]
    df_cygnus_absoluto = df_motor_interno[(df_motor_interno['Nombre Emisor'] == EMISOR_LAURA) & (df_motor_interno['Punto de Venta'] == '2')]
    
    if df_cygnus_absoluto.empty:
        st.info("Aún no hay facturas registradas en la base de datos para Cygnus Home (PV 2 a nombre de Laura).")
    else:
        fecha_max = df_cygnus_absoluto['Mes_Indice'].max()
        f_12m = fecha_max - pd.DateOffset(months=11)
        f_6m = fecha_max - pd.DateOffset(months=5)
        
        fact_ult_mes = df_cygnus_absoluto[df_cygnus_absoluto['Mes_Indice'] == fecha_max]['Facturacion $ Actualizada'].sum()
        fact_ult_6m = df_cygnus_absoluto[df_cygnus_absoluto['Mes_Indice'] >= f_6m]['Facturacion $ Actualizada'].sum()
        fact_ult_12m = df_cygnus_absoluto[df_cygnus_absoluto['Mes_Indice'] >= f_12m]['Facturacion $ Actualizada'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Último Mes", f"$ {fact_ult_mes:,.2f}")
        col2.metric("Últimos 6 Meses", formato_abreviado(fact_ult_6m))
        col3.metric("Últimos 12 Meses", formato_abreviado(fact_ult_12m))
        
        st.divider()
        if not df_cygnus.empty:
            df_cygnus['Año-Mes'] = df_cygnus['Mes_Indice'].dt.strftime('%Y-%m')
            df_evo_cy = df_cygnus.groupby('Año-Mes')[['Facturacion $ Actualizada']].sum().reset_index()
            fig_cy = px.line(df_evo_cy, x='Año-Mes', y='Facturacion $ Actualizada', color_discrete_sequence=['#FF7F0E'])
            fig_cy.update_layout(title="Evolución Real de Cygnus Home", yaxis_tickformat="$.2s", hovermode="x unified")
            fig_cy.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_cy, use_container_width=True)

elif seleccion_pantalla == MENU_SIMULADOR:
    st.subheader("🛠️ Ajustes de Honorarios")
    cv = df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO']
    if not cv.empty: st.error(f"⚠️ Tenés {len(cv)} clientes atrasados.")
    else: st.success("✅ Honorarios al día.")
    
    df_sim = df_clientes[df_clientes['Estado'] == 'Activo'].copy()
    df_sim['Nuevo Precio Pactado'] = df_sim['precio']
    df_ed = st.data_editor(
        df_sim[['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido', 'Nuevo Precio Pactado']], 
        use_container_width=True, disabled=['Nro. Doc. Receptor', 'Denominación Receptor', 'precio', 'Meses Desactualizado', 'Honorario Sugerido'],
        column_config={"precio": st.column_config.NumberColumn("Actual", format="$ %.2f"), "Honorario Sugerido": st.column_config.NumberColumn("Sugerido", format="$ %.2f"), "Nuevo Precio Pactado": st.column_config.NumberColumn("Nuevo Precio ✏️", format="$ %.2f")}
    )
    if st.button("💾 Guardar Cambios en Excel", type="primary"):
        df_nuevo = df_clientes_orig.copy()
        cambios = 0
        for _, r in df_ed.iterrows():
            if r['Nuevo Precio Pactado'] != r['precio']:
                df_nuevo.loc[df_nuevo['Nro. Doc. Receptor'] == r['Nro. Doc. Receptor'], 'precio'] = r['Nuevo Precio Pactado']
                df_nuevo.loc[df_nuevo['Nro. Doc. Receptor'] == r['Nro. Doc. Receptor'], 'Actualizacion'] = datetime.today().strftime('%Y-%m-%d')
                cambios += 1
        if cambios > 0:
            with pd.ExcelWriter(ARCHIVO_LOCAL, engine='openpyxl') as w:
                df_nuevo.to_excel(w, sheet_name='Clientes', index=False)
                df_facturas_orig.to_excel(w, sheet_name='Facturas', index=False)
                df_indices_orig.to_excel(w, sheet_name='Indices', index=False)
            st.success(f"¡{cambios} actualizados!")
            st.rerun()

elif seleccion_pantalla == MENU_AFIP:
    st.subheader("🏛️ Panel Individual de Monotributo")
    
    col_ag, col_la = st.columns(2)
    
    def render_panel_afip(nombre, nombre_emisor, cat_actual, col):
        with col:
            st.write(f"### 👤 {nombre}")
            df_emisor = df_motor_interno[df_motor_interno['Nombre Emisor'] == nombre_emisor]
            fecha_max = df_motor_interno['Fecha_dt'].max()
            f_12m = fecha_max - pd.DateOffset(years=1)
            tot_12 = df_emisor[(df_emisor['Fecha_dt'] > f_12m) & (df_emisor['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
            
            cat_sug = "Excluido"
            for c, lim in ESCALAS_MONOTRIBUTO.items():
                if tot_12 <= lim: 
                    cat_sug = c
                    break
            
            st.metric("Total 12 Meses (Nominal)", f"$ {tot_12:,.2f}")
            st.info(f"📍 Categoría AFIP Sugerida: **{cat_sug}** (Actual: {cat_actual})")
            
    render_panel_afip("Agustín (Estudio)", EMISOR_AGUSTIN, st.session_state['cat_agustin'], col_ag)
    render_panel_afip("Laura (Estudio + Cygnus)", EMISOR_LAURA, st.session_state['cat_laura'], col_la)
    
    st.divider()
    st.write("### 🛠️ Sincronización Post-Recategorización")
    if st.checkbox("¿Hiciste recategorización en AFIP?"):
        n_ag = st.selectbox("Cat. Agustín:", list(ESCALAS_MONOTRIBUTO.keys()), index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_agustin']))
        n_la = st.selectbox("Cat. Laura:", list(ESCALAS_MONOTRIBUTO.keys()), index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_laura']))
        if st.button("Aplicar"):
            st.session_state['cat_agustin'] = n_ag
            st.session_state['cat_laura'] = n_la
            st.rerun()

elif seleccion_pantalla == MENU_CLIENTES:
    st.subheader("👥 Maestro de Clientes")
    df_v = df_clientes.copy()
    df_v['Actualizacion'] = df_v['Actualizacion_Str']
    cols = ['Nro. Doc. Receptor', 'Denominación Receptor', 'Formalidad', 'Periodicidad', 'precio', 'Estado', 'Actualiza', 'Actualizacion', 'periodos', 'Meses Desactualizado', 'Alerta_Revisión']
    st.dataframe(df_v[cols].style.apply(colorear_clientes, axis=1), use_container_width=True, column_config={"precio": st.column_config.NumberColumn("Precio", format="$ %.2f")})

elif seleccion_pantalla == MENU_FACTURAS:
    st.subheader("🧾 Historial de Facturas Filtrado")
    df_h = df_hist_filtrado.copy()
    df_h['Fecha'] = df_h['Fecha_dt'].dt.strftime('%d/%m/%Y')
    cols = ['Fecha', 'Nombre Emisor', 'Tipo', 'Punto de Venta', 'Nro. Doc. Receptor', 'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada']
    st.dataframe(df_h[cols], use_container_width=True, column_config={"Facturacion $": st.column_config.NumberColumn("Original", format="$ %.2f"), "Facturacion $ Actualizada": st.column_config.NumberColumn("Actualizada", format="$ %.2f")})

elif seleccion_pantalla == MENU_INDICES:
    st.subheader("📊 Índices IPC/IPIM")
    st.dataframe(df_indices_vis, use_container_width=True)
