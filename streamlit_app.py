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
    page_title="Generador de Reportes FAMMA", 
    layout="centered", 
    page_icon="📄"
)

st.markdown("""
<style>
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; }
    .stButton>button { height: 3rem; font-size: 16px; font-weight: bold; }
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
# 3. INTERFAZ MINIMALISTA: CONFIGURACIÓN PDF
# ==========================================
st.title("📄 Exportación de Reportes FAMMA")
st.write("Seleccione los parámetros para generar y descargar los reportes consolidados en formato PDF.")
st.divider()

col_p1, col_p2 = st.columns([1, 1])

with col_p1:
    pdf_tipo = st.radio("1. Seleccione el Tipo de Reporte:", ["Diario", "Semanal", "Mensual"], horizontal=True)

pdf_ini, pdf_fin = None, None
pdf_df_oee_target = pd.DataFrame()
pdf_label = ""

with col_p2:
    st.write("2. Seleccione el Período:")
    if pdf_tipo == "Diario":
        min_d = df_raw['Fecha_Filtro'].min().date() if not df_raw.empty else pd.to_datetime("today").date()
        max_d = df_raw['Fecha_Filtro'].max().date() if not df_raw.empty else pd.to_datetime("today").date()
        pdf_fecha = st.date_input("Día para PDF:", value=max_d, min_value=min_d, max_value=max_d)
        pdf_ini, pdf_fin = pd.to_datetime(pdf_fecha), pd.to_datetime(pdf_fecha)
        pdf_df_oee_target = df_oee_diario[df_oee_diario['Fecha_Filtro'] == pdf_ini]
        pdf_label = f"Día {pdf_fecha.strftime('%d-%m-%Y')}"
        
    elif pdf_tipo == "Semanal":
        if not df_oee_sem.empty:
            col_sem = next((c for c in df_oee_sem.columns if 'semana' in c.lower()), df_oee_sem.columns[0])
            opciones_sem = [s for s in df_oee_sem[col_sem].unique() if s.strip() != ""]
            pdf_sem = st.selectbox("Semana para PDF:", opciones_sem)
            pdf_df_oee_target = df_oee_sem[df_oee_sem[col_sem] == pdf_sem]
            pdf_label = f"Semana {pdf_sem}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos semanales.")
                
    elif pdf_tipo == "Mensual":
        if not df_oee_men.empty:
            col_mes = next((c for c in df_oee_men.columns if 'mes' in c.lower()), df_oee_men.columns[0])
            opciones_mes = [m for m in df_oee_men[col_mes].unique() if m.strip() != ""]
            pdf_mes = st.selectbox("Mes para PDF:", opciones_mes)
            pdf_df_oee_target = df_oee_men[df_oee_men[col_mes] == pdf_mes]
            pdf_label = f"Mes {pdf_mes}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos mensuales.")

st.divider()

# ==========================================
# 4. FUNCIONES DE AYUDA PARA DATOS Y PDF
# ==========================================
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

# ==========================================
# 5. MOTOR GENERADOR DEL PDF
# ==========================================
def crear_pdf(area, label_reporte, oee_target_df, ini_date, fin_date, tipo_rep):
    if ini_date is not None and fin_date is not None:
        df_pdf_raw = df_raw[(df_raw['Fecha_Filtro'] >= ini_date) & (df_raw['Fecha_Filtro'] <= fin_date)]
        df_prod_pdf_raw = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_date) & (df_prod_raw['Fecha_Filtro'] <= fin_date)] if not df_prod_raw.empty else pd.DataFrame()
        df_op_pdf_raw = df_operarios_raw[(df_operarios_raw['Fecha_Filtro'] >= ini_date) & (df_operarios_raw['Fecha_Filtro'] <= fin_date)] if not df_operarios_raw.empty else pd.DataFrame()
    else:
        df_pdf_raw = pd.DataFrame(columns=df_raw.columns)
        df_prod_pdf_raw = pd.DataFrame(columns=df_prod_raw.columns)
        df_op_pdf_raw = pd.DataFrame(columns=df_operarios_raw.columns)

    df_pdf = df_pdf_raw[df_pdf_raw['Fábrica'].str.contains(area, case=False, na=False)].copy()
    
    df_prod_pdf = pd.DataFrame()
    if not df_prod_pdf_raw.empty:
        df_prod_pdf = df_prod_pdf_raw[(df_prod_pdf_raw['Máquina'].str.contains(area, case=False, na=False)) | 
                                      (df_prod_pdf_raw['Máquina'].isin(df_pdf['Máquina'].unique()))].copy()

    # Iniciar PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Encabezado
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, clean_text(f"Reporte de Indicadores - {area.upper()}"), ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, clean_text(f"Período del Reporte: {label_reporte} ({tipo_rep})"), ln=True, align='C')
    pdf.ln(5)

    # 1. OEE
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

    # 2. Análisis de Fallas
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("2. Análisis de Fallas Top 10"), ln=True)
    df_fallas_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)]
    if not df_fallas_area.empty:
        top_fallas = df_fallas_area.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(10)
        fig_fallas = px.bar(top_fallas, x='Nivel Evento 6', y='Tiempo (Min)', title=f"Top Fallas - {area}", color='Tiempo (Min)', color_continuous_scale='Reds', text='Tiempo (Min)')
        fig_fallas.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
        fig_fallas.update_layout(width=700, height=400, margin=dict(t=50, b=120, l=40, r=40))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
            fig_fallas.write_image(tmpfile.name, engine="kaleido")
            pdf.image(tmpfile.name, w=160)
            os.remove(tmpfile.name)
        pdf.ln(5)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos detallados de fallas para este período."), ln=True)

    # 3. Producción vs Parada
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("3. Relación Producción vs Parada"), ln=True)
    if not df_pdf.empty:
        df_pdf['Tipo'] = df_pdf['Evento'].apply(lambda x: 'Producción' if 'Producción' in str(x) else 'Parada')
        fig_pie = px.pie(df_pdf, values='Tiempo (Min)', names='Tipo', hole=0.4, color='Tipo', color_discrete_map={'Producción':'#2CA02C', 'Parada':'#D62728'})
        fig_pie.update_layout(width=400, height=300, margin=dict(t=30, b=20, l=20, r=20))
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile2:
            fig_pie.write_image(tmpfile2.name, engine="kaleido")
            pdf.image(tmpfile2.name, w=100)
            os.remove(tmpfile2.name)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos de tiempos para este período."), ln=True)
    pdf.ln(5)

    # 4. Rendimiento de Operadores (CON GRÁFICOS DIARIOS EN SEMANAL/MENSUAL)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, clean_text("4. Rendimiento de Operadores (Performance)"), ln=True)
    
    if not df_op_pdf_raw.empty:
        col_op = next((c for c in df_op_pdf_raw.columns if any(x in c.lower() for x in ['operador', 'nombre'])), 'Operador')
        df_op_pdf_raw['Perf_Num'] = pd.to_numeric(df_op_pdf_raw['Performance'], errors='coerce').fillna(0)
        if df_op_pdf_raw['Perf_Num'].mean() <= 1.5 and df_op_pdf_raw['Perf_Num'].mean() > 0:
            df_op_pdf_raw['Perf_Num'] = df_op_pdf_raw['Perf_Num'] * 100
            
        # Tabla resumen
        df_resumen = df_op_pdf_raw.groupby(col_op).agg(
            Días=('Fecha_Filtro', 'nunique'),
            Perf_Promedio=('Perf_Num', 'mean')
        ).reset_index().sort_values('Perf_Promedio', ascending=False)
        
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(100, 8, clean_text("Operador"), border=1)
        pdf.cell(40, 8, clean_text("Días Trabajados"), border=1, align='C')
        pdf.cell(40, 8, clean_text("Perf. Promedio"), border=1, align='C', ln=True)
        
        pdf.set_font("Arial", '', 10)
        for _, row in df_resumen.iterrows():
            pdf.cell(100, 8, clean_text(str(row[col_op])[:50]), border=1)
            pdf.cell(40, 8, str(row['Días']), border=1, align='C')
            perf_val = row['Perf_Promedio']
            
            # Cambiar color según rendimiento en la tabla
            if perf_val >= 90: pdf.set_text_color(33, 195, 84)
            elif perf_val >= 80: pdf.set_text_color(200, 150, 0)
            else: pdf.set_text_color(220, 20, 20)
            
            pdf.cell(40, 8, f"{perf_val:.1f}%", border=1, align='C', ln=True)
            pdf.set_text_color(0, 0, 0) # Reset color
            
        # Gráficos de Barras Horizontales por Operador (Solo Semanal/Mensual)
        if tipo_rep in ["Semanal", "Mensual"]:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, clean_text("Evolución Diaria por Operador vs Objetivo (90%)"), ln=True)
            pdf.ln(5)
            
            operadores = df_resumen[col_op].tolist()
            
            for op in operadores:
                df_ind = df_op_pdf_raw[df_op_pdf_raw[col_op] == op].sort_values('Fecha_Filtro')
                df_ind['Fecha_Str'] = df_ind['Fecha_Filtro'].dt.strftime('%d-%b')
                
                # Crear gráfico de Plotly
                fig_op = px.bar(
                    df_ind, x='Perf_Num', y='Fecha_Str', orientation='h', text='Perf_Num',
                    title=f"Rendimiento de: {op}",
                    labels={'Perf_Num': 'Performance (%)', 'Fecha_Str': ''}
                )
                
                colores_barras = ['#21C354' if val >= 90 else '#E02020' for val in df_ind['Perf_Num']]
                fig_op.update_traces(marker_color=colores_barras, texttemplate='%{text:.1f}%', textposition='outside')
                fig_op.add_vline(x=90, line_dash="dash", line_color="black", annotation_text=" Meta 90%")
                
                # Altura compacta para el PDF
                altura_graf = 180 + len(df_ind) * 25
                max_x = max(100, df_ind['Perf_Num'].max() * 1.15 if not df_ind.empty else 100)
                
                fig_op.update_layout(
                    xaxis=dict(range=[0, max_x]),
                    yaxis={'categoryorder':'category descending'},
                    height=altura_graf,
                    margin=dict(l=10, r=20, t=40, b=20)
                )
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile_op:
                    fig_op.write_image(tmpfile_op.name, engine="kaleido")
                    
                    # Salto de página automático si no entra el gráfico
                    altura_estimada_pdf = (altura_graf / 4) # Escala aprox de pixeles a mm
                    if pdf.get_y() + altura_estimada_pdf > 270:
                        pdf.add_page()
                        
                    pdf.image(tmpfile_op.name, w=160)
                    os.remove(tmpfile_op.name)
                
                pdf.ln(5)

    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos de operadores para este período."), ln=True)

    # FINALIZAR Y GUARDAR PDF
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f: pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 6. BOTONES DE EXPORTACIÓN EN PANTALLA PRINCIPAL
# ==========================================
st.write("3. Generar y Descargar:")
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("🛠️ Preparar Reporte: ESTAMPADO", use_container_width=True):
        with st.spinner("Construyendo documento PDF con gráficos (puede demorar unos segundos)..."):
            try:
                pdf_data = crear_pdf("Estampado", pdf_label, pdf_df_oee_target, pdf_ini, pdf_fin, pdf_tipo)
                st.download_button("⬇️ Descargar PDF Estampado", data=pdf_data, file_name=f"Estampado_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                st.success("¡Documento de Estampado listo!")
            except Exception as e:
                st.error(f"Error generando PDF: {e}")
                
with col_btn2:
    if st.button("🛠️ Preparar Reporte: SOLDADURA", use_container_width=True):
        with st.spinner("Construyendo documento PDF con gráficos (puede demorar unos segundos)..."):
            try:
                pdf_data = crear_pdf("Soldadura", pdf_label, pdf_df_oee_target, pdf_ini, pdf_fin, pdf_tipo)
                st.download_button("⬇️ Descargar PDF Soldadura", data=pdf_data, file_name=f"Soldadura_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                st.success("¡Documento de Soldadura listo!")
            except Exception as e:
                st.error(f"Error generando PDF: {e}")
