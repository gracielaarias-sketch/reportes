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
    page_title="Exportación de Reportes FAMMA", 
    layout="wide", 
    page_icon="📄"
)

st.markdown("""
<style>
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; }
    .stButton>button { height: 3rem; font-size: 16px; font-weight: bold; }
    .header-style { font-size: 24px; font-weight: bold; margin-bottom: 10px; }
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
            st.error("Atención: No se encontró la configuración de secretos (.streamlit/secrets.toml).")
            return [pd.DataFrame()] * 8

        gid_datos = "0"
        gid_oee_diario = "1767654796"
        gid_prod = "315437448"
        gid_op_diario = "354131379"
        gid_oee_sem = "2079886194"
        gid_oee_men = "1696631148"
        gid_op_sem = "2038636509"
        gid_op_men = "1171574188"
        
        base_export = url_base.split("/edit")[0] + "/export?format=csv&gid="
        
        def process_df(url, is_daily=False):
            try:
                df = pd.read_csv(url)
            except Exception: return pd.DataFrame()
            
            # Limpieza de numéricos
            cols_num = ['Tiempo (Min)', 'Buenas', 'Retrabajo', 'Observadas', 'OEE', 'Disponibilidad', 'Performance', 'Calidad', 'Eficiencia']
            for c in cols_num:
                matches = [col for col in df.columns if c.lower() in col.lower()]
                for match in matches:
                    df[match] = df[match].astype(str).str.replace(',', '.')
                    df[match] = df[match].str.replace('%', '')
                    df[match] = pd.to_numeric(df[match], errors='coerce').fillna(0.0)
            
            # Limpieza de fechas
            col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and 'inicio' not in c.lower() and 'fin' not in c.lower()), None)
            if col_fecha:
                df['Fecha_DT'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
                df['Fecha_Filtro'] = df['Fecha_DT'].dt.normalize()
                if is_daily:
                    df = df.dropna(subset=['Fecha_Filtro'])
            
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].fillna('').astype(str).str.strip()
            return df

        return (
            process_df(base_export + gid_datos, is_daily=True), 
            process_df(base_export + gid_oee_diario, is_daily=True), 
            process_df(base_export + gid_prod, is_daily=True), 
            process_df(base_export + gid_op_diario, is_daily=True),
            process_df(base_export + gid_oee_sem, is_daily=False),
            process_df(base_export + gid_oee_men, is_daily=False),
            process_df(base_export + gid_op_sem, is_daily=False),
            process_df(base_export + gid_op_men, is_daily=False)
        )
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return [pd.DataFrame()] * 8

df_raw, df_oee_diario, df_prod_raw, df_op_diario_raw, df_oee_sem, df_oee_men, df_op_sem_raw, df_op_men_raw = load_data()

if df_raw.empty:
    st.warning("No hay datos cargados en la base principal.")
    st.stop()

# ==========================================
# 3. INTERFAZ: CONFIGURACIÓN PDF
# ==========================================
st.markdown('<div class="header-style">Exportación de Reportes FAMMA</div>', unsafe_allow_html=True)
st.write("Seleccione los parámetros para generar y descargar los reportes consolidados en formato PDF.")
st.divider()

col_p1, col_p2, col_p3 = st.columns([1, 1, 1.5])

with col_p1:
    st.write("**1. Tipo de Reporte:**")
    pdf_tipo = st.radio("Período:", ["Diario", "Semanal", "Mensual"], horizontal=True, label_visibility="collapsed")

pdf_ini, pdf_fin = None, None
pdf_df_oee_target = pd.DataFrame()
pdf_df_op_target = pd.DataFrame()
pdf_label = ""

with col_p2:
    st.write("**2. Seleccione el Período:**")
    if pdf_tipo == "Diario":
        min_d = df_raw['Fecha_Filtro'].min().date() if not df_raw.empty else pd.to_datetime("today").date()
        max_d = df_raw['Fecha_Filtro'].max().date() if not df_raw.empty else pd.to_datetime("today").date()
        pdf_fecha = st.date_input("Día para PDF:", value=max_d, min_value=min_d, max_value=max_d, label_visibility="collapsed")
        
        pdf_ini, pdf_fin = pd.to_datetime(pdf_fecha), pd.to_datetime(pdf_fecha)
        pdf_df_oee_target = df_oee_diario[df_oee_diario['Fecha_Filtro'] == pdf_ini]
        pdf_df_op_target = df_op_diario_raw[df_op_diario_raw['Fecha_Filtro'] == pdf_ini]
        pdf_label = f"Día {pdf_fecha.strftime('%d-%m-%Y')}"
        
    elif pdf_tipo == "Semanal":
        if not df_oee_sem.empty:
            col_sem = df_oee_sem.columns[0]
            opciones_sem = [s for s in df_oee_sem[col_sem].unique() if s.strip() != "" and str(s).lower() != "nan"]
            pdf_sem = st.selectbox("Semana para PDF:", opciones_sem, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_sem[df_oee_sem[col_sem].astype(str) == str(pdf_sem)]
            
            col_sem_op = df_op_sem_raw.columns[0] if not df_op_sem_raw.empty else None
            if col_sem_op:
                pdf_df_op_target = df_op_sem_raw[df_op_sem_raw[col_sem_op].astype(str) == str(pdf_sem)]
                
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
            col_mes = df_oee_men.columns[0]
            opciones_mes = [m for m in df_oee_men[col_mes].unique() if m.strip() != "" and str(m).lower() != "nan"]
            pdf_mes = st.selectbox("Mes para PDF:", opciones_mes, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_men[df_oee_men[col_mes].astype(str) == str(pdf_mes)]
            
            col_mes_op = df_op_men_raw.columns[0] if not df_op_men_raw.empty else None
            if col_mes_op:
                pdf_df_op_target = df_op_men_raw[df_op_men_raw[col_mes_op].astype(str) == str(pdf_mes)]
                
            pdf_label = f"Mes {pdf_mes}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos mensuales.")

with col_p3:
    st.write("**3. Generar y Descargar:**")
    col_btn1, col_btn2 = st.columns(2)

st.divider()

# ==========================================
# 4. FUNCIONES DE AYUDA PARA DATOS Y PDF
# ==========================================
def get_metrics_direct(name_filter, target_df):
    m = {'OEE': 0.0, 'DISP': 0.0, 'PERF': 0.0, 'CAL': 0.0}
    if target_df.empty: return m
    
    mask = target_df.apply(lambda row: row.astype(str).str.upper().str.contains(name_filter.upper()), axis=1)
    datos = target_df[mask.any(axis=1)]
    
    if not datos.empty:
        fila = datos.iloc[0] 
        for key, col_search in {'OEE':['OEE'], 'DISP':['DISPONIBILIDAD', 'DISP'], 'PERF':['PERFORMANCE', 'PERFO'], 'CAL':['CALIDAD', 'CAL']}.items():
            actual_col = next((c for c in datos.columns if any(x in c.upper() for x in col_search)), None)
            if actual_col:
                val_str = str(fila[actual_col]).replace('%', '').replace(',', '.').strip()
                v = pd.to_numeric(val_str, errors='coerce')
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
def crear_pdf(area, label_reporte, oee_target_df, op_target_df, ini_date, fin_date, p_tipo):
    if ini_date is not None and fin_date is not None:
        df_pdf_raw = df_raw[(df_raw['Fecha_Filtro'] >= ini_date) & (df_raw['Fecha_Filtro'] <= fin_date)]
        df_prod_pdf_raw = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_date) & (df_prod_raw['Fecha_Filtro'] <= fin_date)] if not df_prod_raw.empty else pd.DataFrame()
    else:
        df_pdf_raw = pd.DataFrame(columns=df_raw.columns)
        df_prod_pdf_raw = pd.DataFrame(columns=df_prod_raw.columns)

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
    pdf.cell(0, 10, clean_text(f"Periodo del Reporte: {label_reporte}"), ln=True, align='C')
    pdf.ln(5)

    # 1. OEE
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("1. Resumen General y OEE"), ln=True)
    metrics_area = get_metrics_direct(area, oee_target_df)
    print_pdf_metric_row(pdf, f"General {area.upper()}", metrics_area)
    
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Detalle OEE por Maquina/Linea:"), ln=True)
    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    for l in lineas:
        m_l = get_metrics_direct(l, oee_target_df)
        print_pdf_metric_row(pdf, f"   -> {l} ", m_l)
    pdf.ln(5)

    # 2. Análisis de Fallas
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("2. Analisis de Fallas"), ln=True)
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
        
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(0, 8, clean_text("Detalle de Fallas por Maquina:"), ln=True)
        col_inicio = next((c for c in df_pdf.columns if 'inicio' in c.lower() or 'desde' in c.lower()), None)
        col_fin = next((c for c in df_pdf.columns if 'fin' in c.lower() or 'hasta' in c.lower()), None)

        maquinas_con_fallas = sorted(df_fallas_area['Máquina'].unique())
        for maq in maquinas_con_fallas:
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(0, 8, clean_text(f"-> Maquina: {maq}"), ln=True)
            
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(15, 8, clean_text("Inicio"), border=1, align='C')
            pdf.cell(15, 8, clean_text("Fin"), border=1, align='C')
            pdf.cell(90, 8, clean_text("Falla"), border=1)
            pdf.cell(15, 8, clean_text("Min"), border=1, align='C')
            pdf.cell(45, 8, clean_text("Levanto la falla"), border=1, ln=True)
            
            pdf.set_font("Arial", '', 8)
            df_maq = df_fallas_area[df_fallas_area['Máquina'] == maq]
            
            cols_dup = [c for c in [col_inicio, col_fin, 'Nivel Evento 6', 'Operador'] if c is not None]
            if cols_dup: df_maq = df_maq.drop_duplicates(subset=cols_dup)
            df_maq = df_maq.sort_values('Tiempo (Min)', ascending=False)
            
            for _, row in df_maq.iterrows():
                val_inicio = str(row[col_inicio])[:5] if col_inicio and str(row[col_inicio]) != 'nan' else "-"
                val_fin = str(row[col_fin])[:5] if col_fin and str(row[col_fin]) != 'nan' else "-"
                
                pdf.cell(15, 8, clean_text(val_inicio), border=1, align='C')
                pdf.cell(15, 8, clean_text(val_fin), border=1, align='C')
                pdf.cell(90, 8, clean_text(str(row['Nivel Evento 6'])[:60]), border=1)
                pdf.cell(15, 8, clean_text(f"{row['Tiempo (Min)']:.1f}"), border=1, align='C')
                pdf.cell(45, 8, clean_text(str(row['Operador'])[:30]), border=1, ln=True)
            pdf.ln(3) 
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos detallados de fallas para este periodo."), ln=True)

    pdf.ln(5)
    
    # 3. PRODUCCIÓN VS PARADA
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("3. Relacion Produccion vs Parada"), ln=True)
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
        pdf.cell(0, 8, clean_text("No hay datos de tiempos para este periodo."), ln=True)
    pdf.ln(5)
    
    # 4. PRODUCCIÓN POR MÁQUINA
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, clean_text("4. Produccion por Maquina"), ln=True)
    if not df_prod_pdf.empty and 'Buenas' in df_prod_pdf.columns:
        prod_maq = df_prod_pdf.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
        fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=['#1F77B4', '#FF7F0E', '#d62728'], text_auto=True)
        fig_prod.update_layout(width=800, height=450, margin=dict(t=60, b=150, l=40, r=40))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
            fig_prod.write_image(tmpfile3.name, engine="kaleido")
            pdf.image(tmpfile3.name, w=170)
            os.remove(tmpfile3.name)
            
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 8, clean_text("Desglose por Codigo de Producto:"), ln=True)
        pdf.set_font("Arial", 'B', 8)
        
        pdf.cell(40, 8, clean_text("Maquina"), border=1)
        pdf.cell(60, 8, clean_text("Codigo de Producto"), border=1)
        pdf.cell(25, 8, clean_text("Buenas"), border=1, align='C')
        pdf.cell(25, 8, clean_text("Retrabajo"), border=1, align='C')
        pdf.cell(30, 8, clean_text("Observadas"), border=1, align='C', ln=True)
        
        pdf.set_font("Arial", '', 8)
        c_cod = next((c for c in df_prod_pdf.columns if 'código' in c.lower() or 'codigo' in c.lower()), 'Código')
        
        df_prod_group = df_prod_pdf.groupby(['Máquina', c_cod])[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index().sort_values('Máquina')
        for _, row in df_prod_group.iterrows():
            pdf.cell(40, 8, clean_text(str(row['Máquina'])[:25]), border=1)
            pdf.cell(60, 8, clean_text(str(row[c_cod])[:40]), border=1) 
            pdf.cell(25, 8, clean_text(str(int(row['Buenas']))), border=1, align='C')
            pdf.cell(25, 8, clean_text(str(int(row['Retrabajo']))), border=1, align='C')
            pdf.cell(30, 8, clean_text(str(int(row['Observadas']))), border=1, align='C', ln=True)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay datos de produccion para este periodo."), ln=True)
    pdf.ln(5)

    # =========================================================
    # 5. PERFORMANCE DE OPERARIOS (AMBAS ÁREAS + MÁQUINAS)
    # =========================================================
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, clean_text(f"5. Performance de Operarios ({label_reporte})"), ln=True)
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 6, clean_text("Los siguientes cuadros resumen el desempeno y maquinas operadas en ambos sectores."), ln=True)
    pdf.ln(5)
    
    if not op_target_df.empty:
        col_op = next((c for c in op_target_df.columns if 'operador' in c.lower() or 'nombre' in c.lower()), op_target_df.columns[1] if len(op_target_df.columns)>1 else op_target_df.columns[0])
        
        if p_tipo == "Diario":
            col_perf = op_target_df.columns[5] if len(op_target_df.columns) > 5 else None
            col_area = op_target_df.columns[14] if len(op_target_df.columns) > 14 else None
            col_maq = next((c for c in op_target_df.columns if 'máquina' in c.lower() or 'maquina' in c.lower()), None)
        else:
            col_perf = op_target_df.columns[7] if len(op_target_df.columns) > 7 else None
            col_area = op_target_df.columns[1] if len(op_target_df.columns) > 1 else None
            col_maq = None
        
        if col_perf and col_area:
            op_target_df['Perf_Clean'] = pd.to_numeric(op_target_df[col_perf].astype(str).str.replace('%', '').str.replace(',', '.'), errors='coerce').fillna(0)
            if op_target_df['Perf_Clean'].mean() <= 1.5 and op_target_df['Perf_Clean'].mean() > 0:
                op_target_df['Perf_Clean'] = op_target_df['Perf_Clean'] * 100
            op_target_df['Perf_Int'] = op_target_df['Perf_Clean'].round().astype(int)
            
            if p_tipo == "Diario":
                if col_maq:
                    df_grouped = op_target_df.groupby([col_op, col_area]).agg(
                        Perf_Int=('Perf_Int', 'mean'),
                        Maquinas=(col_maq, lambda x: ', '.join(sorted(set([str(i).strip() for i in x.dropna() if str(i).strip() != '']))))
                    ).reset_index()
                else:
                    df_grouped = op_target_df.groupby([col_op, col_area]).agg(Perf_Int=('Perf_Int', 'mean')).reset_index()
                    df_grouped['Maquinas'] = "-"
            else:
                df_grouped = op_target_df.copy()
                if not df_pdf_raw.empty:
                    col_maq_raw = next((c for c in df_pdf_raw.columns if 'máquina' in c.lower() or 'maquina' in c.lower()), None)
                    col_op_raw = next((c for c in df_pdf_raw.columns if 'operador' in c.lower() or 'nombre' in c.lower()), 'Operador')
                    if col_maq_raw and col_op_raw:
                        maq_dict = df_pdf_raw.groupby(col_op_raw)[col_maq_raw].apply(lambda x: ', '.join(sorted(set([str(i).strip() for i in x.dropna() if str(i).strip() != ''])))).to_dict()
                        df_grouped['Maquinas'] = df_grouped[col_op].map(maq_dict).fillna('-')
                    else:
                        df_grouped['Maquinas'] = "-"
                else:
                    df_grouped['Maquinas'] = "-"

            df_grouped['Perf_Int'] = df_grouped['Perf_Int'].round().astype(int)
            
            df_est = df_grouped[df_grouped[col_area].astype(str).str.contains('ESTAMPADO', case=False, na=False)].sort_values('Perf_Int', ascending=False)
            df_sol = df_grouped[df_grouped[col_area].astype(str).str.contains('SOLDADURA', case=False, na=False)].sort_values('Perf_Int', ascending=False)
            
            def imprimir_cuadro_perfo(titulo, df_seccion):
                pdf.set_font("Arial", 'B', 12)
                pdf.cell(0, 8, clean_text(titulo), ln=True)
                
                pdf.set_font("Arial", 'B', 10)
                pdf.cell(70, 8, clean_text("Operador"), border=1)
                pdf.cell(70, 8, clean_text("Maquinas Asignadas"), border=1)
                pdf.cell(40, 8, clean_text("Performance (%)"), border=1, align='C', ln=True)
                
                pdf.set_font("Arial", '', 10)
                if df_seccion.empty:
                    pdf.cell(180, 8, clean_text("Sin registros para esta area."), border=1, align='C', ln=True)
                else:
                    for _, row in df_seccion.iterrows():
                        perf_val = row['Perf_Int']
                        
                        pdf.cell(70, 8, clean_text(str(row[col_op])[:35]), border=1)
                        pdf.cell(70, 8, clean_text(str(row.get('Maquinas', '-'))[:35]), border=1)
                        
                        if perf_val >= 90: pdf.set_text_color(33, 195, 84)
                        elif perf_val >= 80: pdf.set_text_color(200, 150, 0)
                        else: pdf.set_text_color(220, 20, 20)
                        
                        pdf.cell(40, 8, clean_text(str(perf_val) + "%"), border=1, align='C', ln=True)
                        pdf.set_text_color(0, 0, 0)
                pdf.ln(5)
                
            imprimir_cuadro_perfo("Operarios ESTAMPADO", df_est)
            imprimir_cuadro_perfo("Operarios SOLDADURA", df_sol)
            
        else:
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 8, clean_text("No se encontraron las columnas necesarias en la base de datos para generar este cuadro."), ln=True)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No hay registros de performance de operarios para el periodo seleccionado."), ln=True)

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f: pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 6. BOTONES DE EXPORTACIÓN EN PANTALLA
# ==========================================
with col_p3:
    with col_btn1:
        if st.button("Preparar Reporte: ESTAMPADO", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Estampado", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Estampado", data=pdf_data, file_name=f"Estampado_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    with col_btn2:
        if st.button("Preparar Reporte: SOLDADURA", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Soldadura", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Soldadura", data=pdf_data, file_name=f"Soldadura_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
