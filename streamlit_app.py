import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
from fpdf import FPDF

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Indicadores FAMMA", 
    layout="wide", 
    page_icon="🏭"
)

st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 28px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CARGA DE DATOS ROBUSTA
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    try:
        try:
            url_base = st.secrets["connections"]["gsheets"]["spreadsheet"].strip()
        except Exception:
            st.error("⚠️ No se encontró la configuración de secretos (.streamlit/secrets.toml).")
            return [pd.DataFrame()] * 6

        gid_datos = "0"
        gid_oee_diario = "1767654796"
        gid_prod = "315437448"
        gid_operarios = "354131379"
        gid_oee_sem = "2079886194"
        gid_oee_men = "1696631148"
        
        base_export = url_base.split("/edit")[0] + "/export?format=csv&gid="
        
        def process_df(url):
            try:
                df = pd.read_csv(url)
            except Exception: return pd.DataFrame()
            
            cols_num = ['Tiempo (Min)', 'Buenas', 'Retrabajo', 'Observadas', 'OEE', 'Disponibilidad', 'Performance', 'Calidad', 'Eficiencia']
            for c in cols_num:
                matches = [col for col in df.columns if c.lower() in col.lower()]
                for match in matches:
                    df[match] = df[match].astype(str).str.replace(',', '.')
                    df[match] = df[match].str.replace('%', '')
                    df[match] = pd.to_numeric(df[match], errors='coerce').fillna(0.0)
            
            col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and 'inicio' not in c.lower() and 'fin' not in c.lower()), None)
            if col_fecha:
                df['Fecha_DT'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
                df['Fecha_Filtro'] = df['Fecha_DT'].dt.normalize()
                df = df.dropna(subset=['Fecha_Filtro'])
            
            cols_texto = ['Fábrica', 'Máquina', 'Evento', 'Código', 'Operador', 'Nivel Evento 3', 'Nivel Evento 4', 'Nivel Evento 6', 'Nombre', 'Inicio', 'Fin', 'Desde', 'Hasta', 'Semana', 'Mes']
            for c_txt in cols_texto:
                matches = [col for col in df.columns if c_txt.lower() in col.lower() and df[col].dtype == 'object']
                for match in matches:
                    df[match] = df[match].fillna('').astype(str)
            return df

        return (
            process_df(base_export + gid_datos), 
            process_df(base_export + gid_oee_diario), 
            process_df(base_export + gid_prod), 
            process_df(base_export + gid_operarios),
            process_df(base_export + gid_oee_sem),
            process_df(base_export + gid_oee_men)
        )
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return [pd.DataFrame()] * 6

df_raw, df_oee_diario, df_prod_raw, df_operarios_raw, df_oee_sem, df_oee_men = load_data()

if df_raw.empty:
    st.warning("No hay datos cargados en la base principal.")
    st.stop()

# ==========================================
# 3. INTERFAZ SUPERIOR: DASHBOARD
# ==========================================
st.title("🏭 INDICADORES FAMMA")

st.subheader("📊 1. Configuración del Dashboard")
col_d1, col_d2, col_d3 = st.columns([1, 1, 2])

with col_d1:
    tipo_informe = st.radio("Período a visualizar:", ["Diario", "Semanal", "Mensual"], horizontal=True, key="dash_tipo")

ini_filtro, fin_filtro = None, None
df_oee_target = pd.DataFrame()
label_periodo = ""

with col_d2:
    if tipo_informe == "Diario":
        min_d, max_d = df_raw['Fecha_Filtro'].min().date() if not df_raw.empty else pd.to_datetime("today").date(), df_raw['Fecha_Filtro'].max().date() if not df_raw.empty else pd.to_datetime("today").date()
        fecha_sel = st.date_input("Día a analizar:", value=max_d, min_value=min_d, max_value=max_d, key="dash_date")
        ini_filtro, fin_filtro = pd.to_datetime(fecha_sel), pd.to_datetime(fecha_sel)
        df_oee_target = df_oee_diario[df_oee_diario['Fecha_Filtro'] == ini_filtro]
        label_periodo = f"Día: {fecha_sel.strftime('%d-%m-%Y')}"

    elif tipo_informe == "Semanal":
        if not df_oee_sem.empty:
            col_sem = next((c for c in df_oee_sem.columns if 'semana' in c.lower()), df_oee_sem.columns[0])
            opciones_sem = [s for s in df_oee_sem[col_sem].unique() if s.strip() != ""]
            sem_sel = st.selectbox("Semana a analizar:", opciones_sem, key="dash_sem")
            df_oee_target = df_oee_sem[df_oee_sem[col_sem] == sem_sel]
            label_periodo = f"Semana: {sem_sel}"
            
            col_ini = next((c for c in df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin = next((c for c in df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini and col_fin and not df_oee_target.empty:
                ini_filtro = pd.to_datetime(df_oee_target.iloc[0][col_ini], dayfirst=True, errors='coerce')
                fin_filtro = pd.to_datetime(df_oee_target.iloc[0][col_fin], dayfirst=True, errors='coerce')
        else: st.warning("Datos semanales no disponibles.")

    elif tipo_informe == "Mensual":
        if not df_oee_men.empty:
            col_mes = next((c for c in df_oee_men.columns if 'mes' in c.lower()), df_oee_men.columns[0])
            opciones_mes = [m for m in df_oee_men[col_mes].unique() if m.strip() != ""]
            mes_sel = st.selectbox("Mes a analizar:", opciones_mes, key="dash_mes")
            df_oee_target = df_oee_men[df_oee_men[col_mes] == mes_sel]
            label_periodo = f"Mes: {mes_sel}"
            
            col_ini = next((c for c in df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin = next((c for c in df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini and col_fin and not df_oee_target.empty:
                ini_filtro = pd.to_datetime(df_oee_target.iloc[0][col_ini], dayfirst=True, errors='coerce')
                fin_filtro = pd.to_datetime(df_oee_target.iloc[0][col_fin], dayfirst=True, errors='coerce')
        else: st.warning("Datos mensuales no disponibles.")

with col_d3:
    opciones_fabricas = sorted(df_raw['Fábrica'].unique()) if not df_raw.empty else []
    fábricas = st.multiselect("Área / Fábrica:", opciones_fabricas, default=opciones_fabricas)
    opciones_maquinas = sorted(df_raw[df_raw['Fábrica'].isin(fábricas)]['Máquina'].unique()) if not df_raw.empty else []
    máquinas_globales = st.multiselect("Máquinas a incluir:", opciones_maquinas, default=opciones_maquinas)

st.divider()

# ==========================================
# 4. INTERFAZ SUPERIOR: PDF
# ==========================================
st.subheader("📄 2. Configurar y Exportar PDF")
col_p1, col_p2, col_p3 = st.columns([1, 1, 2])

with col_p1:
    pdf_tipo = st.radio("Período del PDF:", ["Diario", "Semanal", "Mensual"], horizontal=True, key="pdf_tipo")

pdf_ini, pdf_fin = None, None
pdf_df_oee_target = pd.DataFrame()
pdf_label = ""

with col_p2:
    if pdf_tipo == "Diario":
        pdf_fecha = st.date_input("Día para PDF:", value=max_d, min_value=min_d, max_value=max_d, key="pdf_date")
        pdf_ini, pdf_fin = pd.to_datetime(pdf_fecha), pd.to_datetime(pdf_fecha)
        pdf_df_oee_target = df_oee_diario[df_oee_diario['Fecha_Filtro'] == pdf_ini]
        pdf_label = f"Día {pdf_fecha.strftime('%d-%m-%Y')}"
        
    elif pdf_tipo == "Semanal":
        if not df_oee_sem.empty:
            pdf_sem = st.selectbox("Semana para PDF:", opciones_sem, key="pdf_sem_sel")
            pdf_df_oee_target = df_oee_sem[df_oee_sem[col_sem] == pdf_sem]
            pdf_label = f"Semana {pdf_sem}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
                
    elif pdf_tipo == "Mensual":
        if not df_oee_men.empty:
            pdf_mes = st.selectbox("Mes para PDF:", opciones_mes, key="pdf_mes_sel")
            pdf_df_oee_target = df_oee_men[df_oee_men[col_mes] == pdf_mes]
            pdf_label = f"Mes {pdf_mes}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')

st.divider()

# ==========================================
# 5. LÓGICA DE DATOS Y DASHBOARD
# ==========================================
if ini_filtro is not None and fin_filtro is not None:
    df_f = df_raw[(df_raw['Fecha_Filtro'] >= ini_filtro) & (df_raw['Fecha_Filtro'] <= fin_filtro)]
    df_prod_f = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_filtro) & (df_prod_raw['Fecha_Filtro'] <= fin_filtro)] if not df_prod_raw.empty else pd.DataFrame()
    df_op_f = df_operarios_raw[(df_operarios_raw['Fecha_Filtro'] >= ini_filtro) & (df_operarios_raw['Fecha_Filtro'] <= fin_filtro)] if not df_operarios_raw.empty else pd.DataFrame()
else:
    df_f, df_prod_f, df_op_f = df_raw.copy(), df_prod_raw.copy(), df_operarios_raw.copy()

df_f = df_f[df_f['Fábrica'].isin(fábricas) & df_f['Máquina'].isin(máquinas_globales)]

def get_metrics(name_filter, target_df):
    m = {'OEE': 0.0, 'DISP': 0.0, 'PERF': 0.0, 'CAL': 0.0}
    if target_df.empty: return m
    mask = target_df.apply(lambda row: row.astype(str).str.upper().str.contains(name_filter.upper()).any(), axis=1)
    datos = target_df[mask]
    if not datos.empty:
        for key, col_search in {'OEE':'OEE', 'DISP':'Disponibilidad', 'PERF':'Performance', 'CAL':'Calidad'}.items():
            actual_col = next((c for c in datos.columns if col_search.lower() in c.lower()), None)
            if actual_col:
                v = pd.to_numeric(datos[actual_col], errors='coerce').dropna().mean()
                if pd.notna(v): m[key] = float(v/100 if v > 1.1 else v)
    return m

def get_color_hex(val):
    if val < 0.85: return "#E02020"
    elif val <= 0.95: return "#D4A000"
    else: return "#21C354"

def render_metric_html(label, val):
    color = get_color_hex(val)
    return f"""
    <div style="line-height: 1.2; margin-bottom: 1rem;">
        <span style="font-size: 14px; color: gray;">{label}</span><br>
        <span style="font-size: 28px; font-weight: bold; color: {color};">{val:.1%}</span>
    </div>
    """

def show_metric_row(m):
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_html("OEE", m['OEE']), unsafe_allow_html=True)
    c2.markdown(render_metric_html("Disponibilidad", m['DISP']), unsafe_allow_html=True)
    c3.markdown(render_metric_html("Performance", m['PERF']), unsafe_allow_html=True)
    c4.markdown(render_metric_html("Calidad", m['CAL']), unsafe_allow_html=True)

# ---- RENDER DEL DASHBOARD ----
st.markdown(f"### Visualizando datos para: **{tipo_informe} - {label_periodo}**")

show_metric_row(get_metrics('GENERAL', df_oee_target))

t1, t2 = st.tabs(["Estampado", "Soldadura"])
with t1:
    show_metric_row(get_metrics('ESTAMPADO', df_oee_target))
    with st.expander("Ver Líneas"):
        for l in ['L1', 'L2', 'L3', 'L4']:
            st.markdown(f"**{l}**"); show_metric_row(get_metrics(l, df_oee_target)); st.markdown("---")
with t2:
    show_metric_row(get_metrics('SOLDADURA', df_oee_target))
    with st.expander("Ver Detalle"):
        st.markdown("**Celdas Robotizadas**"); show_metric_row(get_metrics('CELDA', df_oee_target)); st.markdown("---")
        st.markdown("**PRP**"); show_metric_row(get_metrics('PRP', df_oee_target))

# Gráficos Adicionales
col_graf1, col_graf2 = st.columns(2)
with col_graf1:
    st.subheader("Análisis de Tiempos")
    if not df_f.empty:
        df_f['Tipo'] = df_f['Evento'].apply(lambda x: 'Producción' if 'Producción' in str(x) else 'Parada')
        st.plotly_chart(px.pie(df_f, values='Tiempo (Min)', names='Tipo', hole=0.4), use_container_width=True)

with col_graf2:
    st.subheader("Balance Producción")
    if not df_prod_f.empty:
        c_maq = next((c for c in df_prod_f.columns if 'máquina' in c.lower() or 'maquina' in c.lower()), None)
        c_b = next((c for c in df_prod_f.columns if 'buenas' in c.lower()), 'Buenas')
        c_r = next((c for c in df_prod_f.columns if 'retrabajo' in c.lower()), 'Retrabajo')
        c_o = next((c for c in df_prod_f.columns if 'observadas' in c.lower()), 'Observadas')
        if c_maq:
            df_st = df_prod_f.groupby(c_maq)[[c_b, c_r, c_o]].sum().reset_index()
            st.plotly_chart(px.bar(df_st, x=c_maq, y=[c_b, c_r, c_o], barmode='stack'), use_container_width=True)

st.markdown("---")
st.subheader("Análisis de Fallas Top 15")
df_fallas = df_f[df_f['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)].copy()
if not df_fallas.empty:
    top_f = df_fallas.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(15)
    fig = px.bar(top_f, x='Tiempo (Min)', y='Nivel Evento 6', orientation='h', text='Tiempo (Min)', color='Tiempo (Min)', color_continuous_scale='Reds')
    fig.update_traces(texttemplate='%{text:.0f} min', textposition='outside')
    fig.update_layout(yaxis={'categoryorder':'total ascending'}, coloraxis_showscale=False, height=450)
    st.plotly_chart(fig, use_container_width=True)


# ==========================================
# 6. FUNCIONES DE PDF (FPDF)
# ==========================================
def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).encode('latin-1', 'replace').decode('latin-1')

def set_pdf_color(pdf, val):
    if val < 0.85: pdf.set_text_color(220, 20, 20)
    elif val <= 0.95: pdf.set_text_color(200, 150, 0)
    else: pdf.set_text_color(33, 195, 84)

def print_pdf_metric_row(pdf, prefix, m):
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.write(6, clean_text(f"{prefix} | OEE: "))
    set_pdf_color(pdf, m['OEE'])
    pdf.write(6, f"{m['OEE']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(6, clean_text(" | Disp: "))
    set_pdf_color(pdf, m['DISP'])
    pdf.write(6, f"{m['DISP']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(6, clean_text(" | Perf: "))
    set_pdf_color(pdf, m['PERF'])
    pdf.write(6, f"{m['PERF']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(6, clean_text(" | Calidad: "))
    set_pdf_color(pdf, m['CAL'])
    pdf.write(6, f"{m['CAL']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.ln(6)

def crear_pdf(area, label_reporte, oee_target_df, ini_date, fin_date):
    # Filtrar bases crudas si hay rango de fechas
    if ini_date is not None and fin_date is not None:
        df_pdf_raw = df_raw[(df_raw['Fecha_Filtro'] >= ini_date) & (df_raw['Fecha_Filtro'] <= fin_date)]
        df_prod_pdf_raw = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_date) & (df_prod_raw['Fecha_Filtro'] <= fin_date)] if not df_prod_raw.empty else pd.DataFrame()
    else:
        # Fallback si el reporte no tiene fechas
        df_pdf_raw = pd.DataFrame(columns=df_raw.columns)
        df_prod_pdf_raw = pd.DataFrame(columns=df_prod_raw.columns)

    df_pdf = df_pdf_raw[df_pdf_raw['Fábrica'].str.contains(area, case=False, na=False)].copy()
    
    df_prod_pdf = pd.DataFrame()
    if not df_prod_pdf_raw.empty:
        df_prod_pdf = df_prod_pdf_raw[(df_prod_pdf_raw['Máquina'].str.contains(area, case=False, na=False)) | 
                                      (df_prod_pdf_raw['Máquina'].isin(df_pdf['Máquina'].unique()))].copy()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, clean_text(f"Reporte de Indicadores - {area.upper()}"), ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, clean_text(f"Período del Reporte: {label_reporte}"), ln=True, align='C')
    pdf.ln(5)

    # 1. OEE DEL ÁREA Y MÁQUINAS
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("1. Resumen General y OEE"), ln=True)
    metrics_area = get_metrics(area, oee_target_df)
    print_pdf_metric_row(pdf, f"General {area.upper()}", metrics_area)
    
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Detalle OEE por Máquina/Línea:"), ln=True)
    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    for l in lineas:
        m_l = get_metrics(l, oee_target_df)
        print_pdf_metric_row(pdf, f"   -> {l} ", m_l)
    pdf.ln(5)

    # 2. ANÁLISIS DE FALLAS
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("2. Análisis de Fallas"), ln=True)
    df_fallas_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)]
    
    if not df_fallas_area.empty:
        top_fallas = df_fallas_area.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(10)
        fig_fallas = px.bar(top_fallas, x='Nivel Evento 6', y='Tiempo (Min)', title=f"Top 10 Fallas - {area}", color='Tiempo (Min)', color_continuous_scale='Reds', text='Tiempo (Min)')
        fig_fallas.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
        fig_fallas.update_layout(width=800, height=450, margin=dict(t=80, b=150, l=40, r=40))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
            fig_fallas.write_image(tmpfile.name, engine="kaleido")
            pdf.image(tmpfile.name, w=170)
            os.remove(tmpfile.name)
        pdf.ln(5)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos detallados de fallas para este período."), ln=True)

    # 3. PRODUCCIÓN VS PARADA
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("3. Relación Producción vs Parada"), ln=True)
    if not df_pdf.empty:
        df_pdf['Tipo'] = df_pdf['Evento'].apply(lambda x: 'Producción' if 'Producción' in str(x) else 'Parada')
        fig_pie = px.pie(df_pdf, values='Tiempo (Min)', names='Tipo', hole=0.4, color='Tipo', color_discrete_map={'Producción':'#2CA02C', 'Parada':'#D62728'})
        fig_pie.update_layout(width=500, height=350, margin=dict(t=30, b=20, l=20, r=20))
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile2:
            fig_pie.write_image(tmpfile2.name, engine="kaleido")
            pdf.image(tmpfile2.name, w=110)
            os.remove(tmpfile2.name)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos de tiempos para este período."), ln=True)
    pdf.ln(5)

    # 4. PRODUCCIÓN POR MÁQUINA
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("4. Producción por Máquina"), ln=True)
    if not df_prod_pdf.empty and 'Buenas' in df_prod_pdf.columns:
        prod_maq = df_prod_pdf.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
        fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=['#1F77B4', '#FF7F0E', '#d62728'], text_auto=True)
        fig_prod.update_layout(width=800, height=450, margin=dict(t=60, b=150, l=40, r=40))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
            fig_prod.write_image(tmpfile3.name, engine="kaleido")
            pdf.image(tmpfile3.name, w=170)
            os.remove(tmpfile3.name)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos de producción para este período."), ln=True)

    # --- NUEVAS SECCIONES: BAÑO Y REFRIGERIO ---
    def agregar_tabla_operarios(titulo, regex_keyword, numero_seccion):
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 10, clean_text(f"{numero_seccion}. {titulo}"), ln=True)
        
        try:
            # Apuntamos a las columnas exactas por su posición (0-indexed)
            # Aseguramos que el dataframe tenga al menos 17 columnas (0 a 16)
            if df_pdf_raw.shape[1] > 16:
                s_operario = df_pdf_raw.iloc[:, 0]
                s_tiempo = df_pdf_raw.iloc[:, 9]
                s_evento = df_pdf_raw.iloc[:, 16]
                
                # Aseguramos que el tiempo sea un número (reemplazando comas si hay)
                s_tiempo_num = pd.to_numeric(s_tiempo.astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                
                # Armamos un DataFrame temporal solo con lo que nos importa
                df_temp = pd.DataFrame({
                    'Operario': s_operario,
                    'Tiempo': s_tiempo_num,
                    'Evento': s_evento.astype(str)
                })
                
                # Filtramos por la palabra clave (BAÑO o REFRIGERIO)
                mask = df_temp['Evento'].str.contains(regex_keyword, case=False, na=False)
                df_filtrado = df_temp[mask]
                
                if not df_filtrado.empty:
                    # Agrupamos por el operario, sumamos el tiempo y contamos los eventos
                    resumen = df_filtrado.groupby('Operario').agg(
                        Total_Min=('Tiempo', 'sum'), 
                        Eventos=('Tiempo', 'count')
                    ).reset_index().sort_values('Total_Min', ascending=False)
                    
                    # Encabezados de tabla
                    pdf.set_font("Arial", 'B', 10)
                    pdf.cell(90, 8, clean_text("Operador"), border=1, align='C')
                    pdf.cell(30, 8, clean_text("Cant. Eventos"), border=1, align='C')
                    pdf.cell(30, 8, clean_text("Total (Min)"), border=1, align='C')
                    pdf.ln()
                    
                    # Filas de tabla
                    pdf.set_font("Arial", '', 10)
                    for _, r in resumen.iterrows():
                        op = clean_text(r['Operario']) if str(r['Operario']).strip() else "Desconocido"
                        pdf.cell(90, 8, op, border=1)
                        pdf.cell(30, 8, str(int(r['Eventos'])), border=1, align='C')
                        pdf.cell(30, 8, f"{r['Total_Min']:.1f}", border=1, align='C')
                        pdf.ln()
                else:
                    pdf.set_font("Arial", '', 10)
                    pdf.cell(0, 8, clean_text("No se registraron tiempos para este evento en este período."), ln=True)
            else:
                pdf.set_font("Arial", '', 10)
                pdf.cell(0, 8, clean_text("Error: La base de datos no tiene suficientes columnas (Faltan A, J o Q)."), ln=True)
                
        except Exception as e:
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 8, clean_text(f"Error procesando los datos: {str(e)}"), ln=True)

    # 5. TIEMPO DE BAÑO
    agregar_tabla_operarios("Tiempo de Baño por Operario", "BAÑO|BANO", 5)
    
    # 6. TIEMPO DE REFRIGERIO
    agregar_tabla_operarios("Tiempo de Refrigerio por Operario", "REFRIGERIO", 6)

    # FINALIZAR Y GUARDAR PDF
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f: pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 7. BOTONES DE EXPORTACIÓN PDF
# ==========================================
with col_p3:
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("🛠️ Preparar PDF Estampado", use_container_width=True):
            with st.spinner("Generando PDF..."):
                try:
                    pdf_data = crear_pdf("Estampado", pdf_label, pdf_df_oee_target, pdf_ini, pdf_fin)
                    st.download_button("⬇️ Guardar Estampado", data=pdf_data, file_name=f"Estampado_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    with col_btn2:
        if st.button("🛠️ Preparar PDF Soldadura", use_container_width=True):
            with st.spinner("Generando PDF..."):
                try:
                    pdf_data = crear_pdf("Soldadura", pdf_label, pdf_df_oee_target, pdf_ini, pdf_fin)
                    st.download_button("⬇️ Guardar Soldadura", data=pdf_data, file_name=f"Soldadura_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
