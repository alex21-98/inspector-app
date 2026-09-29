import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from adjustText import adjust_text
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
import io
import datetime

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Reporte MP-Calidad", layout="wide", page_icon="📊")

st.title("📊🖥️ Generador de Reportes de MP-Calidad Gerencial")
st.markdown("Complejo Agroindustrial Beta - Sistema de Visualización de Defectos")

# ==========================================
# 1. CARGA DE DATOS Y PREPARACIÓN
# ==========================================
MAPEO_COLUMNAS = {
    'Fecha': ['fecha', 'fec', 'date'], 'Mes': ['mes', 'month'],
    'Año': ['año', 'anio', 'year'], 'Semana': ['semana','SEMANA', 'sem', 'week', 'no sem'],
    'Fundo': ['Fundos', 'campo', 'fundo'], 'Lote': ['lote', 'lot', 'id lote'],
    'Variedad': ['variedad', 'var']
}

@st.cache_data
def cargar_hoja_cruda(file, sheet, freq):
    """Carga y estandariza una hoja genérica retornando la data cruda para promedios simples"""
    df_raw = pd.read_excel(file, sheet_name=sheet)
    df_proc = df_raw.copy()
    nuevos_nombres = {}
    
    for col_excel in df_proc.columns:
        col_limpia = str(col_excel).strip().lower()
        for nombre_estandar, sinonimos in MAPEO_COLUMNAS.items():
            if col_limpia in [s.lower() for s in sinonimos] or col_limpia == nombre_estandar.lower():
                nuevos_nombres[col_excel] = nombre_estandar
                break

    df_proc = df_proc.rename(columns=nuevos_nombres)
    df_proc = df_proc.loc[:, ~df_proc.columns.duplicated()]

    if 'Fecha' in df_proc.columns:
        df_proc['Fecha'] = pd.to_datetime(df_proc['Fecha'], errors='coerce', dayfirst=True)
        df_proc = df_proc.dropna(subset=['Fecha'])

        # Lógica de Periodos
        if freq == 'D':
            df_proc['Periodo'] = df_proc['Fecha'].dt.strftime('%d/%m/%Y')
            df_proc['Orden_Periodo'] = df_proc['Fecha'].dt.normalize()
        elif freq == 'W':
            df_proc['Orden_Periodo'] = df_proc['Fecha'] - pd.to_timedelta(df_proc['Fecha'].dt.dayofweek, unit='d')
            if 'Semana' in df_proc.columns:
                df_proc['Periodo'] = 'Sem ' + df_proc['Semana'].fillna(0).astype(int).astype(str)
            else:
                df_proc['Periodo'] = df_proc['Orden_Periodo'].dt.strftime('Sem %V')
        elif freq == 'M':
            df_proc['Periodo'] = df_proc['Fecha'].dt.strftime('%m/%Y')
            df_proc['Orden_Periodo'] = df_proc['Fecha'] - pd.to_timedelta(df_proc['Fecha'].dt.day - 1, unit='d')
        elif freq == 'Y':
            df_proc['Periodo'] = df_proc['Fecha'].dt.strftime('%Y')
            df_proc['Orden_Periodo'] = pd.to_datetime(df_proc['Fecha'].dt.year, format='%Y')

    columnas_requeridas = ['Periodo', 'Orden_Periodo', 'Fundo', 'Lote', 'Variedad']
    if all(col in df_proc.columns for col in columnas_requeridas):
        df_proc['Variedad'] = df_proc['Variedad'].astype(str)
        df_proc['Fundo'] = df_proc['Fundo'].astype(str)
        df_proc['Lote'] = df_proc['Lote'].astype(str)
        
    return df_proc

def filtrar_dataframe(df, fecha_ini, fecha_fin, fundos_sel, variedades_sel, lotes_sel):
    """Aplica los filtros de UI consistentes para cualquier DataFrame cargado"""
    if df.empty or 'Fecha' not in df.columns: return df
    
    mask_fechas = (df['Fecha'].dt.date >= fecha_ini) & (df['Fecha'].dt.date <= fecha_fin)
    df_filt = df[mask_fechas].copy()
    
    if fundos_sel and 'Fundo' in df_filt.columns: 
        df_filt = df_filt[df_filt['Fundo'].isin(fundos_sel)]
    if variedades_sel and 'Variedad' in df_filt.columns: 
        df_filt = df_filt[df_filt['Variedad'].isin(variedades_sel)]
    
    if 'Fundo' in df_filt.columns and 'Lote' in df_filt.columns:
        df_filt['Etiqueta_Lote'] = df_filt['Fundo'] + " - " + df_filt['Lote']
        if lotes_sel: 
            df_filt = df_filt[df_filt['Etiqueta_Lote'].isin(lotes_sel)]
            
    return df_filt

# ==========================================
# LÓGICA DE GRÁFICOS REUTILIZABLE
# ==========================================
def generar_grafico_metricas(data_var, col_agrupacion, periodo_col, orden_col, defecto, 
                             titulo_personalizado, texto_fundos, var, tolerancias_defectos, 
                             es_escala_decimal, sufijo="%", multiplicador=1.0):
                             
    fig, ax = plt.subplots(figsize=(14, 8), dpi=150)
    
    # Asegurarse que Fundo está en groupby si agrupan por otra cosa
    cols_groupby = [col_agrupacion, periodo_col, orden_col]
    if 'Fundo' not in cols_groupby and 'Fundo' in data_var.columns:
        cols_groupby.append('Fundo')
        
    # PROMEDIO SIMPLE: Agrupamos directamente la data cruda, dándole peso igual a cada muestra original
    data_agg = data_var.groupby(cols_groupby, dropna=False)[defecto].mean().reset_index()
    
    periodos_ordenados = data_agg.sort_values(orden_col)[periodo_col].unique()
    entidades_presentes = data_agg[col_agrupacion].unique()
    textos_a_ajustar = []

    num_lineas = len(entidades_presentes)
    fs_dinamico = 14 if num_lineas <= 5 else (13 if num_lineas <= 8 else 12)
    pad_box = 0.3 if num_lineas <= 5 else (0.25 if num_lineas <= 8 else 0.15)
    
    colores_fuertes = [
        '#1976D2', '#388E3C', '#FBC02D', '#8E24AA', '#F57C00', '#0097A7', '#689F38', '#C2185B', '#111111', '#455A64',
        '#3F51B5', '#00796B', '#AFB42B', '#512DA8', '#0288D1', '#F50057', '#C0CA33', '#8D6E63', '#26A69A', '#00ACC1'
    ]

    for i, entidad in enumerate(entidades_presentes):
        data_entidad = data_agg[data_agg[col_agrupacion] == entidad]
        color_asignado = colores_fuertes[i % len(colores_fuertes)]

        valores_entidad_alineados = []
        for per in periodos_ordenados:
            valor_per = data_entidad.loc[data_entidad[periodo_col] == per, defecto]
            if not valor_per.empty and pd.notna(valor_per.iloc[0]):
                valores_entidad_alineados.append(valor_per.iloc[0])
            else:
                valores_entidad_alineados.append(np.nan)

        ax.plot(periodos_ordenados, valores_entidad_alineados, marker='o', label=entidad, 
                linewidth=4, color=color_asignado, markersize=8, markeredgecolor='white', 
                markeredgewidth=1.5, zorder=5)

        for x_val, p in zip(periodos_ordenados, valores_entidad_alineados):
            if pd.notna(p):
                val_etq = p * multiplicador if es_escala_decimal else p
                t = ax.text(x_val, p, f'{val_etq:.1f}{sufijo}', fontsize=fs_dinamico, fontweight='bold', color='white', 
                            ha='center', va='center', zorder=10, 
                            bbox=dict(facecolor=color_asignado, alpha=0.9, edgecolor='white', boxstyle=f'square,pad={pad_box}'))
                textos_a_ajustar.append(t)

    if textos_a_ajustar:
        adjust_text(textos_a_ajustar, ax=ax, only_move={'points': 'xy', 'text': 'xy'}, 
                    expand_points=(1.15, 1.25), expand_text=(1.15, 1.25), 
                    force_text=(1.0, 1.5), force_points=(0.6, 1.4), 
                    arrowprops=dict(arrowstyle='-', color='#78909C', lw=1.2, alpha=0.8, zorder=2), 
                    max_move=40)

    if defecto in tolerancias_defectos:
        valor_tol = tolerancias_defectos[defecto]
        valor_tol_grafico = valor_tol / multiplicador if es_escala_decimal else valor_tol
        ax.axhline(y=valor_tol_grafico, color='#D32F2F', linestyle='--', linewidth=2.5, label=f'Límite ({valor_tol}{sufijo})', zorder=3)
        ax.text(0, valor_tol_grafico, f' T: {valor_tol}{sufijo} ', color='white', fontsize=12, fontweight='bold',
                ha='left', va='bottom', zorder=15, bbox=dict(facecolor='#D32F2F', edgecolor='white', alpha=0.9, boxstyle='round,pad=0.3'))

    columnas_leyenda = min(len(entidades_presentes) + 1, 5)
    total_items = len(entidades_presentes) + 1
    filas_leyenda = (total_items + columnas_leyenda - 1) // columnas_leyenda
    pad_dinamico = 10 + (filas_leyenda * 15)

    ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.02), ncol=columnas_leyenda, frameon=False, fontsize=11)
    
    color_texto_principal = '#45605A'
    color_borde_grafico = '#B0BEC5'
    
    ax.set_title(f"{titulo_personalizado}: {defecto} - {texto_fundos}\n".upper(), fontsize=18, fontweight='bold', color=color_texto_principal, pad=pad_dinamico)
    ax.set_xlabel(f"\nVariedad: {str(var).upper()}", fontsize=14, fontweight='bold', color=color_texto_principal)
    ax.set_xticklabels(periodos_ordenados, rotation=45, ha='right', fontsize=12)
    ax.get_yaxis().set_visible(False)
    ax.margins(y=0.25)
    
    for spine in ax.spines.values():
        spine.set_visible(True)
        spine.set_color(color_borde_grafico)
        spine.set_linewidth(2)

    plt.figtext(0.99, 0.01, 'Complejo Agroindustrial Beta', horizontalalignment='right', fontsize=10, color='gray', style='italic')
    plt.tight_layout()
    
    return fig

# --- PANEL LATERAL ---
with st.sidebar:
    st.header("📁 1. Carga de Datos")
    uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        excel_obj = pd.ExcelFile(uploaded_file)
        hojas_disponibles = excel_obj.sheet_names
        hoja_final = st.selectbox("👉 Selecciona la hoja principal (Defectos):", hojas_disponibles)
        st.header("⚙️ 2. Periodicidad")
        opcion_freq = st.selectbox("Agrupar reporte por:", ["Semanal", "Diario", "Mensual", "Anual"])
        mapa_freq = {"Diario": 'D', "Semanal": 'W', "Mensual": 'M', "Anual": 'Y'}
        freq_elegida = mapa_freq[opcion_freq]

# --- MANEJO DE ESTADO PARA DESCARGA (Evita pérdida de datos) ---
if 'pptx_export' not in st.session_state:
    st.session_state.pptx_export = None
    st.session_state.pptx_filename = ""

# --- LÓGICA PRINCIPAL ---
if uploaded_file is not None:
    df_raw_proc = cargar_hoja_cruda(uploaded_file, hoja_final, freq_elegida)
    lista_defectos = [col for col in df_raw_proc.columns if '%' in str(col)]
    
    if not df_raw_proc.empty:
        st.divider()
        col_f1, col_f2 = st.columns(2)
        
        fecha_min = df_raw_proc['Fecha'].min().date()
        fecha_max = df_raw_proc['Fecha'].max().date()
        
        fecha_ini = col_f1.date_input("Fecha Inicio", value=fecha_min, min_value=fecha_min, max_value=fecha_max)
        fecha_fin = col_f2.date_input("Fecha Fin", value=fecha_max, min_value=fecha_min, max_value=fecha_max)
        
        # Filtros Iniciales para extraer opciones únicas
        df_temp = filtrar_dataframe(df_raw_proc, fecha_ini, fecha_fin, [], [], [])
        
        col1, col2, col3 = st.columns(3)
        fundos_sel = col1.multiselect("Fundos", df_temp['Fundo'].dropna().unique() if 'Fundo' in df_temp.columns else [])
        variedades_sel = col2.multiselect("Variedades", df_temp['Variedad'].dropna().unique() if 'Variedad' in df_temp.columns else [])
        if 'Fundo' in df_temp.columns and 'Lote' in df_temp.columns:
            df_temp['Etiqueta_Lote'] = df_temp['Fundo'] + " - " + df_temp['Lote']
            lotes_sel = col3.multiselect("Lotes", df_temp['Etiqueta_Lote'].dropna().unique())
        else:
            lotes_sel = []
        
        defectos_sel = st.multiselect("Defectos a graficar", lista_defectos)

        # ==========================================
        # FILTRO PARA NIVEL DE DETALLE DEL GRÁFICO
        # ==========================================
        st.divider()
        st.subheader("📈 3. Configuración Visual de Gráficos")
        col_visual1, col_visual2 = st.columns(2)
        
        with col_visual1:
            nivel_agrupacion = st.radio(
                "Selecciona cómo deseas agrupar las líneas en el gráfico:",
                ["Por Lote", "Por Fundo"],
                horizontal=True
            )
            col_agrupacion = 'Etiqueta_Lote' if nivel_agrupacion == "Por Lote" else 'Fundo'
            
        with col_visual2:
            titulo_personalizado = st.text_input(
                "Título base del gráfico:", 
                value="Evaluación de PT",
                help="Cambia este texto para modificar el inicio del título en gráficos y diapositivas"
            )

        tolerancias_defectos = {}
        if defectos_sel:
            st.divider()
            st.subheader("📏 4. Configuración de Tolerancias (%)")
            cols = st.columns(min(len(defectos_sel), 4))
            for i, defecto in enumerate(defectos_sel):
                with cols[i % 4]:
                    tol_val = st.number_input(f"{defecto}", min_value=0.0, max_value=100.0, value=0.0, step=0.5, key=f"tol_{defecto}")
                    if tol_val > 0:
                        tolerancias_defectos[defecto] = tol_val

        st.divider()
        
        # ==========================================
        # GENERACIÓN Y VISUALIZACIÓN
        # ==========================================
        if st.button("🚀 Generar Reporte Gerencial", type="primary", use_container_width=True):
            if not fundos_sel or not variedades_sel or not defectos_sel:
                st.warning("⚠️ Debes seleccionar al menos un Fundo, Variedad y Defecto para generar el reporte.")
            else:
                prs = Presentation()
                diapositivas_creadas = 0
                graficos_expander = st.expander("Ver Previsualización de Gráficos", expanded=True)
                
                # 1. GENERACIÓN DE HOJA PRINCIPAL (DEFECTOS)
                df_plot = filtrar_dataframe(df_raw_proc, fecha_ini, fecha_fin, fundos_sel, variedades_sel, lotes_sel)
                
                if not df_plot.empty:
                    texto_fundos = " y ".join(df_plot['Fundo'].unique())
                    for defecto in defectos_sel:
                        for var in variedades_sel:
                            data_var = df_plot[df_plot['Variedad'] == var]
                            if data_var[defecto].isnull().all() or data_var.empty: continue
                            
                            max_val_test = data_var[defecto].max()
                            es_escala_decimal = max_val_test < 1.0 if pd.notna(max_val_test) else False
                            
                            fig = generar_grafico_metricas(
                                data_var, col_agrupacion, 'Periodo', 'Orden_Periodo', defecto, 
                                titulo_personalizado, texto_fundos, var, tolerancias_defectos, 
                                es_escala_decimal, sufijo="%", multiplicador=100.0
                            )
                            graficos_expander.pyplot(fig)
                            
                            image_stream = io.BytesIO()
                            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight', transparent=False)
                            plt.close(fig)

                            slide = prs.slides.add_slide(prs.slide_layouts[5])
                            title_shape = slide.shapes.title
                            title_shape.text = f"{titulo_personalizado}: {defecto}"
                            title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(69, 96, 90)
                            title_shape.text_frame.paragraphs[0].font.bold = True
                            image_stream.seek(0)
                            slide.shapes.add_picture(image_stream, Inches(0.25), Inches(1.8), width=Inches(9.5))
                            diapositivas_creadas += 1

                # 2. INTEGRACIÓN AUTOMÁTICA DE HOJAS EXTRAS (BRIX Y ACIDEZ)
                # Hoja BRIX
                if "BRIX" in hojas_disponibles:
                    df_brix_crudo = cargar_hoja_cruda(uploaded_file, "BRIX", freq_elegida)
                    df_brix_plot = filtrar_dataframe(df_brix_crudo, fecha_ini, fecha_fin, fundos_sel, variedades_sel, lotes_sel)
                    
                    if not df_brix_plot.empty and 'Brix Promedio' in df_brix_plot.columns:
                        for var in variedades_sel:
                            data_brix_var = df_brix_plot[df_brix_plot['Variedad'] == var]
                            if data_brix_var['Brix Promedio'].isnull().all() or data_brix_var.empty: continue
                            
                            fig = generar_grafico_metricas(
                                data_brix_var, col_agrupacion, 'Periodo', 'Orden_Periodo', 'Brix Promedio', 
                                "Análisis de BRIX", texto_fundos, var, {}, 
                                False, sufijo="°Bx", multiplicador=1.0  # Sin % porque Brix es un grado absoluto
                            )
                            graficos_expander.pyplot(fig)
                            
                            image_stream = io.BytesIO()
                            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight', transparent=False)
                            plt.close(fig)
                            slide = prs.slides.add_slide(prs.slide_layouts[5])
                            title_shape = slide.shapes.title
                            title_shape.text = f"Análisis de BRIX: Brix Promedio"
                            title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(69, 96, 90)
                            title_shape.text_frame.paragraphs[0].font.bold = True
                            image_stream.seek(0)
                            slide.shapes.add_picture(image_stream, Inches(0.25), Inches(1.8), width=Inches(9.5))
                            diapositivas_creadas += 1

                # Hoja ACIDEZ
                if "ACIDEZ" in hojas_disponibles:
                    df_acidez_crudo = cargar_hoja_cruda(uploaded_file, "ACIDEZ", freq_elegida)
                    df_acidez_plot = filtrar_dataframe(df_acidez_crudo, fecha_ini, fecha_fin, fundos_sel, variedades_sel, lotes_sel)
                    
                    if not df_acidez_plot.empty and '% Acidez' in df_acidez_plot.columns:
                        for var in variedades_sel:
                            data_acidez_var = df_acidez_plot[df_acidez_plot['Variedad'] == var]
                            if data_acidez_var['% Acidez'].isnull().all() or data_acidez_var.empty: continue
                            
                            max_acidez = data_acidez_var['% Acidez'].max()
                            es_dec_acidez = max_acidez < 1.0 if pd.notna(max_acidez) else False
                            
                            fig = generar_grafico_metricas(
                                data_acidez_var, col_agrupacion, 'Periodo', 'Orden_Periodo', '% Acidez', 
                                "Análisis de Acidez", texto_fundos, var, {}, 
                                es_dec_acidez, sufijo="%", multiplicador=100.0
                            )
                            graficos_expander.pyplot(fig)
                            
                            image_stream = io.BytesIO()
                            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight', transparent=False)
                            plt.close(fig)
                            slide = prs.slides.add_slide(prs.slide_layouts[5])
                            title_shape = slide.shapes.title
                            title_shape.text = f"Análisis de Acidez: % Acidez"
                            title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(69, 96, 90)
                            title_shape.text_frame.paragraphs[0].font.bold = True
                            image_stream.seek(0)
                            slide.shapes.add_picture(image_stream, Inches(0.25), Inches(1.8), width=Inches(9.5))
                            diapositivas_creadas += 1

                if diapositivas_creadas > 0:
                    pptx_stream = io.BytesIO()
                    prs.save(pptx_stream)
                    pptx_stream.seek(0)
                    
                    # GUARDAR EN SESSION STATE para que el botón de descarga no desaparezca
                    st.session_state.pptx_export = pptx_stream.getvalue()
                    st.session_state.pptx_filename = f"Reporte_Calidad_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    st.success(f"✅ Reporte generado exitosamente con {diapositivas_creadas} diapositivas.")

        # BOTÓN DE DESCARGA RENDERIZADO FUERA DEL st.button("Generar")
        if st.session_state.pptx_export:
            st.download_button(
                label="📥 Descargar Presentación Generada (PPTX)",
                data=st.session_state.pptx_export,
                file_name=st.session_state.pptx_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )
else:
    st.info("👋 Sube un archivo Excel en el menú lateral para comenzar a configurar el reporte.")
