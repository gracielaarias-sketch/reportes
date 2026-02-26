import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import tempfile
import os
from fpdf import FPDF

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Generador de Reportes PDF - FAMMA", 
    layout="centered", 
    page_icon="📄"
)

# --- BOTÓN DE ACTUALIZAR DATOS ---
col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("📄 Reportes PDF - FAMMA")
with col_btn:
    st.write("") # Espaciador
    if st.button("Actualizar Datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown("Seleccione un periodo de tiempo y descargue el informe gerencial.")

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
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        gid_datos, gid_oee, gid_prod, gid_operarios = "0", "1767654796", "315437448", "354131379"
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
            
            col_fecha = next((c for c in df.columns if 'fecha' in c.lower()), None)
            if col_fecha:
                df['Fecha_DT'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
                df['Fecha_Filtro'] = df['Fecha_DT'].dt.normalize()
                df = df.dropna(subset=['Fecha_Filtro'])
            
            cols_texto = ['Fábrica', 'Máquina', 'Evento', 'Código', 'Operador', 'Nivel Evento 3', 'Nivel Evento 4', 'Nivel Evento 6', 'Nombre', 'Inicio', 'Fin', 'Desde', 'Hasta']
            for c_txt in cols_texto:
                matches = [col for col in df.columns if c_txt.lower() in col.lower()]
                for match in matches:
                    df[match] = df[match].fillna('').astype(str)
            return df

        return process_df(base_export + gid_datos), process_df(base_export + gid_oee), \
               process_df(base_export + gid_prod), process_df(base_export + gid_operarios)
    except Exception as e:
        st.error(f"Error: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_raw, df_oee_raw, df_prod_raw, df_operarios_raw = load_data()

if df_raw.empty:
    st.warning("No hay datos cargados. Verifique la conexión con Google Sheets.")
    st.stop()

# ==========================================
# 3. INTERFAZ DE EXPORTACIÓN (PERIODO)
# ==========================================
min_d, max_d = df_raw['Fecha_Filtro'].min().date(), df_raw['Fecha_Filtro'].max().date()

st.markdown("---")
rango_fechas = st.date_input(
    "Seleccionar periodo para el reporte:", 
    value=(max_d, max_d), 
    min_value=min_d, 
    max_value=max_d
)
st.markdown("---")

if not isinstance(rango_fechas, tuple) or len(rango_fechas) != 2:
    st.info("Por favor, selecciona la fecha de inicio y fin del periodo.")
    st.stop()

fecha_ini = pd.to_datetime(rango_fechas[0]).normalize()
fecha_fin = pd.to_datetime(rango_fechas[1]).normalize()

# ==========================================
# 4. CLASE PDF PERSONALIZADA
# ==========================================
class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area
        self.fecha_str = fecha_str
        self.theme_color = theme_color

    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 8, 30)
        
        self.set_font("Times", 'B', 16)
        self.set_text_color(*self.theme_color)
        self.cell(0, 10, clean_text(f"REPORTE GERENCIAL - {self.area.upper()}"), ln=True, align='R')
        
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 6, clean_text(f"Periodo: {self.fecha_str}"), ln=True, align='R')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

# ==========================================
# 5. FUNCIONES AUXILIARES PARA PDF
# ==========================================
def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).encode('latin-1', 'replace').decode('latin-1')

def check_space(pdf, required_height):
    if pdf.get_y() + required_height > (pdf.h - 15):
        pdf.add_page()

def get_metrics_pdf(name_filter, df_oee_target):
    m = {'OEE': 0.0, 'DISP': 0.0, 'PERF': 0.0, 'CAL': 0.0}
    if df_oee_target.empty: return m
    mask = df_oee_target.apply(lambda row: row.astype(str).str.upper().str.contains(name_filter.upper()).any(), axis=1)
    datos = df_oee_target[mask]
    if not datos.empty:
        for key, col_search in {'OEE':'OEE', 'DISP':'Disponibilidad', 'PERF':'Performance', 'CAL':'Calidad'}.items():
            actual_col = next((c for c in datos.columns if col_search.lower() in c.lower()), None)
            if actual_col:
                vals = pd.to_numeric(datos[actual_col], errors='coerce').dropna()
                if not vals.empty:
                    v = vals.mean() # Promedio ponderado
                    m[key] = float(v/100 if v > 1.1 else v)
    return m

def set_pdf_color(pdf, val):
    if val < 0.85: pdf.set_text_color(220, 20, 20)
    elif val <= 0.95: pdf.set_text_color(200, 150, 0)
    else: pdf.set_text_color(33, 195, 84)

def print_pdf_metric_row(pdf, prefix, m):
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.write(6, clean_text(f"{prefix} | OEE: "))
    set_pdf_color(pdf, m['OEE']); pdf.write(6, f"{m['OEE']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(6, clean_text(" | Disp: "))
    set_pdf_color(pdf, m['DISP']); pdf.write(6, f"{m['DISP']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(6, clean_text(" | Perf: "))
    set_pdf_color(pdf, m['PERF']); pdf.write(6, f"{m['PERF']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(6, clean_text(" | Cal: "))
    set_pdf_color(pdf, m['CAL']); pdf.write(6, f"{m['CAL']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.ln(6)

def print_section_title(pdf, title, theme_color):
    pdf.ln(4)
    pdf.set_font("Times", 'B', 14)
    pdf.set_text_color(*theme_color)
    pdf.cell(0, 8, clean_text(title), ln=True)
    
    x, y = pdf.get_x(), pdf.get_y()
    pdf.set_draw_color(*theme_color)
    pdf.set_line_width(0.5)
    pdf.line(x, y, x + 190, y)
    
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(4)

def setup_table_header(pdf, theme_color):
    pdf.set_fill_color(*theme_color)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(*theme_color)

def setup_table_row(pdf):
    pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(50, 50, 50)
    pdf.set_draw_color(200, 200, 200)

def redactar_resumen_ejecutivo(pdf, area, df_pdf, df_oee_pdf):
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    if df_pdf.empty and df_oee_pdf.empty:
        pdf.multi_cell(0, 6, clean_text("No hay suficientes datos registrados en este periodo para generar un resumen ejecutivo."))
        return

    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    mejores_oee = {}
    for l in lineas:
        m = get_metrics_pdf(l, df_oee_pdf)
        if m['OEE'] > 0: mejores_oee[l] = m['OEE']
    
    texto_oee = ""
    if mejores_oee:
        mejor_maq = max(mejores_oee, key=mejores_oee.get)
        texto_oee = f"Durante este periodo, la línea/máquina con mejor rendimiento general fue la {mejor_maq} con un OEE promedio de {mejores_oee[mejor_maq]:.1%}. "

    df_fallas = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)]
    texto_fallas = "No se registraron tiempos muertos por fallas significativos. "
    if not df_fallas.empty:
        total_falla_min = df_fallas['Tiempo (Min)'].sum()
        falla_top = df_fallas.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().idxmax()
        texto_fallas = f"Se registró un total de {total_falla_min:.1f} minutos de tiempo de parada por fallas. La causa principal de impacto en el periodo fue '{falla_top}'. "

    resumen = f"Resumen Ejecutivo: {texto_oee}{texto_fallas}Este informe abarca todos los registros documentados en el periodo seleccionado."
    
    pdf.set_fill_color(245, 245, 245)
    pdf.multi_cell(0, 7, clean_text(resumen), border=0, fill=True)
    pdf.ln(5)

# ==========================================
# 6. GENERACIÓN DEL PDF
# ==========================================
def crear_pdf(area, f_ini, f_fin):
    if f_ini == f_fin: fecha_str = f_ini.strftime('%d/%m/%Y')
    else: fecha_str = f"{f_ini.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}"
    
    if area.upper() == "ESTAMPADO":
        theme_color = (41, 128, 185)
        chart_bars = ['#1F77B4', '#AEC7E8', '#FF7F0E']
    else:
        theme_color = (211, 84, 0)
        chart_bars = ['#E67E22', '#FAD7A1', '#d62728']
        
    hex_theme = '#%02x%02x%02x' % theme_color

    df_pdf = df_raw[(df_raw['Fecha_Filtro'] >= f_ini) & (df_raw['Fecha_Filtro'] <= f_fin) & (df_raw['Fábrica'].str.contains(area, case=False))].copy()
    df_oee_pdf = df_oee_raw[(df_oee_raw['Fecha_Filtro'] >= f_ini) & (df_oee_raw['Fecha_Filtro'] <= f_fin)].copy()
    
    df_prod_pdf = pd.DataFrame()
    if not df_prod_raw.empty:
        df_prod_pdf = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= f_ini) & (df_prod_raw['Fecha_Filtro'] <= f_fin) & 
                                  (df_prod_raw['Máquina'].str.contains(area, case=False) | df_prod_raw['Máquina'].isin(df_pdf['Máquina'].unique()))].copy()
    
    df_op_pdf = pd.DataFrame()
    if not df_operarios_raw.empty:
        df_op_pdf = df_operarios_raw[(df_operarios_raw['Fecha_Filtro'] >= f_ini) & (df_operarios_raw['Fecha_Filtro'] <= f_fin)].copy()

    pdf = ReportePDF(area, fecha_str, theme_color)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # 0. RESUMEN EJECUTIVO
    redactar_resumen_ejecutivo(pdf, area, df_pdf, df_oee_pdf)

    # 1. OEE DEL ÁREA Y MÁQUINAS
    check_space(pdf, 60)
    print_section_title(pdf, "1. Resumen General y OEE", theme_color)
    
    metrics_area = get_metrics_pdf(area, df_oee_pdf)
    print_pdf_metric_row(pdf, f"General {area.upper()}", metrics_area)
    
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Tiempos Promedio (por registro en el periodo):"), ln=True)
    pdf.set_font("Arial", '', 10)
    
    if not df_pdf.empty:
        avg_bano = df_pdf[df_pdf['Nivel Evento 4'].astype(str).str.contains('Baño', case=False, na=False)]['Tiempo (Min)'].mean()
        avg_refr = df_pdf[df_pdf['Nivel Evento 4'].astype(str).str.contains('Refrigerio', case=False, na=False)]['Tiempo (Min)'].mean()
        
        pdf.cell(0, 6, clean_text(f"   • Promedio Baño: {avg_bano:.1f} min" if not pd.isna(avg_bano) else "   • Promedio Baño: Sin registros"), ln=True)
        pdf.cell(0, 6, clean_text(f"   • Promedio Refrigerio: {avg_refr:.1f} min" if not pd.isna(avg_refr) else "   • Promedio Refrigerio: Sin registros"), ln=True)
    else:
         pdf.cell(0, 6, clean_text("   • Sin datos de tiempos para el área y periodo."), ln=True)
    pdf.ln(3)

    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Detalle OEE por Línea / Célula:"), ln=True)
    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    for l in lineas:
        m_l = get_metrics_pdf(l, df_oee_pdf)
        print_pdf_metric_row(pdf, f"   ➤ {l} ", m_l)
        
    # ASTERISCO DE PROMEDIO PONDERADO SI ES UN RANGO DE FECHAS
    if f_ini != f_fin:
        pdf.ln(1)
        pdf.set_font("Arial", 'I', 8)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 5, clean_text("* Los porcentajes mostrados corresponden a un promedio ponderado de los registros en el periodo seleccionado."), ln=True)
        pdf.set_text_color(0, 0, 0)
    pdf.ln(2)

    # 2. ANÁLISIS DE FALLAS (GRÁFICO PARETO)
    df_fallas_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)]
    if not df_fallas_area.empty:
        check_space(pdf, 110) 
        print_section_title(pdf, "2. Análisis de Fallas (Diagrama Pareto)", theme_color)
        
        top_fallas = df_fallas_area.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(10)
        top_fallas['% Acumulado'] = (top_fallas['Tiempo (Min)'].cumsum() / top_fallas['Tiempo (Min)'].sum()) * 100
        
        fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
        fig_pareto.add_trace(
            go.Bar(x=top_fallas['Nivel Evento 6'], y=top_fallas['Tiempo (Min)'], name='Tiempo Perdido', marker_color=hex_theme, text=top_fallas['Tiempo (Min)'].round(1), textposition='outside'),
            secondary_y=False,
        )
        fig_pareto.add_trace(
            go.Scatter(x=top_fallas['Nivel Evento 6'], y=top_fallas['% Acumulado'], name='% Acumulado', mode='lines+markers', line=dict(color='red', width=3)),
            secondary_y=True,
        )
        fig_pareto.update_layout(title_text=f"Impacto de Fallas - {area}", width=800, height=450, margin=dict(t=50, b=120, l=40, r=40), plot_bgcolor='rgba(0,0,0,0)', showlegend=False)
        fig_pareto.update_yaxes(title_text="Minutos", secondary_y=False)
        fig_pareto.update_yaxes(title_text="% Acumulado", range=[0, 105], secondary_y=True)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
            fig_pareto.write_image(tmpfile.name, engine="kaleido")
            pdf.image(tmpfile.name, w=170)
            os.remove(tmpfile.name)
        
        pdf.ln(3)
        check_space(pdf, 30)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(*theme_color)
        pdf.cell(0, 8, clean_text("Registro Detallado de Fallas:"), ln=True)
        pdf.ln(1)
        
        col_inicio = next((c for c in df_pdf.columns if 'inicio' in c.lower() or 'desde' in c.lower()), None)
        col_fin = next((c for c in df_pdf.columns if 'fin' in c.lower() or 'hasta' in c.lower()), None)

        maquinas_con_fallas = sorted(df_fallas_area['Máquina'].unique())
        for maq in maquinas_con_fallas:
            check_space(pdf, 30)
            pdf.set_font("Arial", 'B', 9)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(0, 8, clean_text(f"Máquina: {maq}"), ln=True)
            
            setup_table_header(pdf, theme_color)
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(20, 7, clean_text("Fecha"), border=1, align='C', fill=True)
            pdf.cell(15, 7, clean_text("Inicio"), border=1, align='C', fill=True)
            pdf.cell(15, 7, clean_text("Fin"), border=1, align='C', fill=True)
            pdf.cell(80, 7, clean_text("Falla"), border=1, fill=True)
            pdf.cell(15, 7, clean_text("Min"), border=1, align='C', fill=True)
            pdf.cell(45, 7, clean_text("Operador"), border=1, ln=True, fill=True)
            
            setup_table_row(pdf)
            pdf.set_font("Arial", '', 8)
            df_maq = df_fallas_area[df_fallas_area['Máquina'] == maq]
            
            cols_dup = [c for c in [col_inicio, col_fin, 'Nivel Evento 6', 'Operador'] if c is not None]
            if cols_dup: df_maq = df_maq.drop_duplicates(subset=cols_dup)
            df_maq = df_maq.sort_values(['Fecha_Filtro', 'Tiempo (Min)'], ascending=[False, False])
            
            for _, row in df_maq.iterrows():
                val_fecha = pd.to_datetime(row['Fecha_Filtro']).strftime('%d/%m') if pd.notna(row['Fecha_Filtro']) else "-"
                val_inicio = str(row[col_inicio])[:5] if col_inicio and str(row[col_inicio]) != 'nan' else "-"
                val_fin = str(row[col_fin])[:5] if col_fin and str(row[col_fin]) != 'nan' else "-"
                
                pdf.cell(20, 7, clean_text(val_fecha), border='B', align='C')
                pdf.cell(15, 7, clean_text(val_inicio), border='B', align='C')
                pdf.cell(15, 7, clean_text(val_fin), border='B', align='C')
                pdf.cell(80, 7, clean_text(str(row['Nivel Evento 6'])[:55]), border='B')
                pdf.cell(15, 7, clean_text(f"{row['Tiempo (Min)']:.1f}"), border='B', align='C')
                pdf.cell(45, 7, clean_text(str(row['Operador'])[:25]), border='B', ln=True)
            pdf.ln(4) 
    else:
        check_space(pdf, 30)
        print_section_title(pdf, "2. Análisis de Fallas", theme_color)
        pdf.set_font("Arial", 'I', 9)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, clean_text("No se registraron fallas en este periodo."), ln=True)

    # 3. HORARIOS DE OPERACIÓN POR MÁQUINA
    check_space(pdf, 40)
    print_section_title(pdf, "3. Horarios Extremos de Operación en el Periodo", theme_color)

    if not df_pdf.empty:
        col_inicio = next((c for c in df_pdf.columns if 'inicio' in c.lower() or 'desde' in c.lower()), None)
        col_fin = next((c for c in df_pdf.columns if 'fin' in c.lower() or 'hasta' in c.lower()), None)

        if col_inicio and col_fin:
            df_times = df_pdf[['Máquina', col_inicio, col_fin]].copy()
            df_times['dt_inicio'] = pd.to_datetime(df_times[col_inicio], format='%H:%M:%S', errors='coerce').fillna(pd.to_datetime(df_times[col_inicio], format='%H:%M', errors='coerce'))
            df_times['dt_fin'] = pd.to_datetime(df_times[col_fin], format='%H:%M:%S', errors='coerce').fillna(pd.to_datetime(df_times[col_fin], format='%H:%M', errors='coerce'))

            setup_table_header(pdf, theme_color)
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(80, 8, clean_text("Máquina"), border=1, fill=True)
            pdf.cell(50, 8, clean_text("Registro Más Temprano"), border=1, align='C', fill=True)
            pdf.cell(50, 8, clean_text("Registro Más Tardío"), border=1, align='C', ln=True, fill=True)

            setup_table_row(pdf)
            pdf.set_font("Arial", '', 9)
            for maq in sorted(df_times['Máquina'].unique()):
                maq_data = df_times[df_times['Máquina'] == maq]
                
                min_time = maq_data['dt_inicio'].min()
                max_time = maq_data['dt_fin'].max()

                str_min = min_time.strftime('%H:%M') if pd.notnull(min_time) else "-"
                str_max = max_time.strftime('%H:%M') if pd.notnull(max_time) else "-"

                pdf.cell(80, 8, clean_text(str(maq)[:50]), border='B')
                pdf.cell(50, 8, clean_text(str_min), border='B', align='C')
                pdf.cell(50, 8, clean_text(str_max), border='B', align='C', ln=True)
        else:
            pdf.set_font("Arial", 'I', 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, clean_text("No se encontraron columnas de horario."), ln=True)
    pdf.ln(5)

    # 4. PRODUCCIÓN VS PARADA
    if not df_pdf.empty:
        check_space(pdf, 100)
        print_section_title(pdf, "4. Relación Producción vs Parada", theme_color)
        df_pdf['Tipo'] = df_pdf['Evento'].apply(lambda x: 'Producción' if 'Producción' in str(x) else 'Parada')
        color_parada = '#D62728' 
        
        fig_pie = px.pie(df_pdf, values='Tiempo (Min)', names='Tipo', hole=0.4, color='Tipo', color_discrete_map={'Producción':hex_theme, 'Parada':color_parada})
        fig_pie.update_layout(width=500, height=350, margin=dict(t=30, b=20, l=20, r=20), plot_bgcolor='rgba(0,0,0,0)')
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile2:
            fig_pie.write_image(tmpfile2.name, engine="kaleido")
            pdf.image(tmpfile2.name, w=110)
            os.remove(tmpfile2.name)
        pdf.ln(5)

    # 5. PRODUCCIÓN POR MÁQUINA
    if not df_prod_pdf.empty and 'Buenas' in df_prod_pdf.columns:
        check_space(pdf, 110)
        print_section_title(pdf, "5. Producción por Máquina (Acumulado)", theme_color)
        prod_maq = df_prod_pdf.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
        
        fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=chart_bars, text_auto=True)
        fig_prod.update_layout(width=800, height=450, margin=dict(t=60, b=150, l=40, r=40), plot_bgcolor='rgba(0,0,0,0)')
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
            fig_prod.write_image(tmpfile3.name, engine="kaleido")
            pdf.image(tmpfile3.name, w=170)
            os.remove(tmpfile3.name)
            
        pdf.ln(3)
        check_space(pdf, 35)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(*theme_color)
        pdf.cell(0, 8, clean_text("Desglose por Código de Producto:"), ln=True)
        
        setup_table_header(pdf, theme_color)
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(40, 7, clean_text("Máquina"), border=1, fill=True)
        pdf.cell(60, 7, clean_text("Código de Producto"), border=1, fill=True)
        pdf.cell(25, 7, clean_text("Buenas"), border=1, align='C', fill=True)
        pdf.cell(25, 7, clean_text("Retrabajo"), border=1, align='C', fill=True)
        pdf.cell(30, 7, clean_text("Observadas"), border=1, align='C', ln=True, fill=True)
        
        setup_table_row(pdf)
        pdf.set_font("Arial", '', 8)
        c_cod = next((c for c in df_prod_pdf.columns if 'código' in c.lower() or 'codigo' in c.lower()), 'Código')
        
        df_prod_group = df_prod_pdf.groupby(['Máquina', c_cod])[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index().sort_values('Máquina')
        for _, row in df_prod_group.iterrows():
            pdf.cell(40, 7, clean_text(str(row['Máquina'])[:25]), border='B')
            pdf.cell(60, 7, clean_text(str(row[c_cod])[:40]), border='B') 
            pdf.cell(25, 7, clean_text(str(int(row['Buenas']))), border='B', align='C')
            pdf.cell(25, 7, clean_text(str(int(row['Retrabajo']))), border='B', align='C')
            pdf.cell(30, 7, clean_text(str(int(row['Observadas']))), border='B', align='C', ln=True)
        pdf.ln(5)

    # 6. TIEMPOS POR OPERARIO
    if not df_pdf.empty:
        check_space(pdf, 110)
        print_section_title(pdf, "6. Tiempos Totales por Operario", theme_color)
        op_tiempos = df_pdf.groupby('Operador')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False)
        
        fig_op = px.bar(op_tiempos, x='Operador', y='Tiempo (Min)', color='Tiempo (Min)', color_continuous_scale='Blues' if area.upper()=="ESTAMPADO" else 'Oranges', text='Tiempo (Min)')
        fig_op.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
        fig_op.update_layout(width=800, height=450, margin=dict(t=80, b=150, l=40, r=40), plot_bgcolor='rgba(0,0,0,0)', coloraxis_showscale=False)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile4:
            fig_op.write_image(tmpfile4.name, engine="kaleido")
            pdf.image(tmpfile4.name, w=170)
            os.remove(tmpfile4.name)
        pdf.ln(5)

    # 7. PERFORMANCE DE OPERARIOS
    check_space(pdf, 40)
    print_section_title(pdf, "7. Performance de Operarios", theme_color)
    
    if not df_op_pdf.empty:
        c_op_name = next((c for c in df_op_pdf.columns if 'operador' in c.lower() or 'nombre' in c.lower()), None)
        c_perf = next((c for c in df_op_pdf.columns if 'performance' in c.lower()), None)
        
        if c_op_name and c_perf:
            # Promedio ponderado (media de sus registros en el rango)
            df_op_pdf[c_perf] = pd.to_numeric(df_op_pdf[c_perf], errors='coerce')
            df_op_print = df_op_pdf.groupby(c_op_name)[c_perf].mean().reset_index().sort_values(c_op_name)
            
            setup_table_header(pdf, theme_color)
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(120, 8, clean_text("Operador"), border=1, fill=True)
            
            # Cambiar el título de la tabla si es un rango de fechas
            header_perf = "Performance Promedio (%) *" if f_ini != f_fin else "Performance (%)"
            pdf.cell(60, 8, clean_text(header_perf), border=1, align='C', ln=True, fill=True)
            
            setup_table_row(pdf)
            pdf.set_font("Arial", '', 9)
            for _, row in df_op_print.iterrows():
                perf_val = row[c_perf]
                perf_str = f"{perf_val:.2f}" if pd.notna(perf_val) else "-"
                
                pdf.cell(120, 8, clean_text(str(row[c_op_name])[:60]), border='B')
                pdf.cell(60, 8, clean_text(perf_str), border='B', align='C', ln=True)
                
            # Advertencia de promedio ponderado al final de la tabla
            if f_ini != f_fin:
                pdf.ln(2)
                pdf.set_font("Arial", 'I', 8)
                pdf.set_text_color(100, 100, 100)
                pdf.cell(0, 5, clean_text("* La performance mostrada corresponde a un promedio ponderado de los registros en el periodo seleccionado."), ln=True)
                pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_font("Arial", 'I', 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, clean_text("No se encontraron columnas de Operador/Performance en la base de datos."), ln=True)
    else:
        pdf.set_font("Arial", 'I', 9)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, clean_text("No hay registros de performance para el periodo seleccionado."), ln=True)

    # FINALIZAR Y GUARDAR PDF
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f:
        pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 6. BOTONES EN LA INTERFAZ PRINCIPAL
# ==========================================
col1, col2 = st.columns(2)

with col1:
    if st.button("Generar PDF Estampado", use_container_width=True):
        with st.spinner(f"Generando PDF de Estampado..."):
            try:
                pdf_data = crear_pdf("Estampado", fecha_ini, fecha_fin)
                file_n = f"Reporte_Estampado_{fecha_ini.strftime('%d%m')}" if fecha_ini == fecha_fin else f"Reporte_Estampado_{fecha_ini.strftime('%d%m')}_al_{fecha_fin.strftime('%d%m')}"
                st.download_button(
                    label="Descargar PDF Estampado", 
                    data=pdf_data, 
                    file_name=f"{file_n}.pdf", 
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("¡PDF listo para descargar!")
            except Exception as e:
                st.error(f"Error generando el PDF: {e}")

with col2:
    if st.button("Generar PDF Soldadura", use_container_width=True):
        with st.spinner(f"Generando PDF de Soldadura..."):
            try:
                pdf_data = crear_pdf("Soldadura", fecha_ini, fecha_fin)
                file_n = f"Reporte_Soldadura_{fecha_ini.strftime('%d%m')}" if fecha_ini == fecha_fin else f"Reporte_Soldadura_{fecha_ini.strftime('%d%m')}_al_{fecha_fin.strftime('%d%m')}"
                st.download_button(
                    label="Descargar PDF Soldadura", 
                    data=pdf_data, 
                    file_name=f"{file_n}.pdf", 
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("¡PDF listo para descargar!")
            except Exception as e:
                st.error(f"Error generando el PDF: {e}")
