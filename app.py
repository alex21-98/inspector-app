import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from adjustText import adjust_text
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import datetime
import base64
from io import BytesIO

# ==========================================
# CONFIGURACIÓN INICIAL DE STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Reporte de Calidad - Beta",
    page_icon="📊",
    layout="wide"
)

# Estilo CSS personalizado
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #45605A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        background-color: #45605A;
        color: white;
        font-weight: bold;
        padding: 0.5rem 2rem;
        font-size: 1.2rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# TÍTULO PRINCIPAL
# ==========================================
st.markdown('<h1 class="main-header">📊 Reporte Gerencial de Calidad<br>Complejo Agroindustrial Beta</h1>', 
            unsafe_allow_html=True)

# ==========================================
# 1. CARGA DE DATOS
# ==========================================
st.header("📁 1. Carga de Datos")

# Inicializar variables de sesión
if 'df_crudo' not in st.session_state:
    st.session_state.df_crudo = None
if 'df_final' not in st.session_state:
    st.session_state.df_final = None
if 'lista_defectos' not in st.session_state:
    st.session_state.lista_defectos = None

archivo_subido = st.file_uploader(
    "Sube tu archivo Excel (.xlsx):",
    type=['xlsx'],
    help="Arrastra y suelta tu archivo Excel aquí"
)

if archivo_subido is not None:
    try:
        # Leer el archivo Excel
        excel_obj = pd.ExcelFile(archivo_subido)
        
        # Selección de hoja
        hojas = excel_obj.sheet_names
        hoja_seleccionada = st.selectbox(
            "📋 Selecciona la hoja de trabajo:",
            options=hojas,
            index=0
        )
        
        # Cargar datos
        df_crudo = pd.read_excel(archivo_subido, sheet_name=hoja_seleccionada)
        st.session_state.df_crudo = df_crudo
        
        st.success(f"✅ Hoja '{hoja_seleccionada}' cargada exitosamente - {df_crudo.shape[0]} filas x {df_crudo.shape[1]} columnas")
        
        # Mostrar vista previa
        with st.expander("🔍 Vista previa de los datos"):
            st.dataframe(df_crudo.head(10))
            
    except Exception as e:
        st.error(f"❌ Error al cargar el archivo: {str(e)}")

# ==========================================
# 2. NORMALIZACIÓN Y AGRUPACIÓN
# ==========================================
if st.session_state.df_crudo is not None:
    st.header("⚙️ 2. Procesamiento de Datos")
    
    MAPEO_COLUMNAS = {
        'Fecha': ['fecha','Fecha','FECHA', 'fec', 'date'],
        'Mes': ['mes','Mes','MES', 'month'],
        'Año': ['año','Año','AÑO', 'anio', 'year'],
        'Semana': ['semana','Semana', 'SEMANA', 'sem', 'week', 'no sem'],
        'Fundo': ['Fundos','FUNDOS', 'campo', 'fundo','FUNDO','Fundo'],
        'Lote': ['lote','LOTE' 'lot', 'id lote'],
        'Variedad': ['variedad','VARIEDAD','VARIEDADES',variedades','Variedad', 'var']
    }

    def preparar_datos(df_raw):
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
        
        # Identificar columnas de defectos (con %)
        columnas_defectos = [col for col in df_proc.columns if '%' in str(col) or 'defecto' in str(col).lower()]
        
        # Si no encuentra columnas con %, busca columnas numéricas que no sean las estándar
        if not columnas_defectos:
            columnas_estandar = ['Fecha', 'Fundo', 'Lote', 'Variedad', 'Mes', 'Año', 'Semana']
            columnas_defectos = [col for col in df_proc.columns if col not in columnas_estandar 
                               and df_proc[col].dtype in ['float64', 'int64']]

        if 'Fecha' in df_proc.columns:
            df_proc['Fecha'] = pd.to_datetime(df_proc['Fecha'], errors='coerce', dayfirst=True)
            df_proc = df_proc.dropna(subset=['Fecha'])

        columnas_requeridas = ['Fecha', 'Fundo', 'Lote', 'Variedad']

        if all(col in df_proc.columns for col in columnas_requeridas):
            df_proc['Variedad'] = df_proc['Variedad'].astype(str)
            df_proc['Fundo'] = df_proc['Fundo'].astype(str)
            df_proc['Lote'] = df_proc['Lote'].astype(str)

            columnas_agrupar = columnas_requeridas + columnas_defectos
            df_diario = df_proc[columnas_agrupar].groupby(columnas_requeridas).mean().reset_index()
            df_diario = df_diario.sort_values(by=columnas_requeridas, ascending=[True, True, True, True])
            return df_diario, columnas_defectos
        else:
            st.error("❌ Faltan columnas críticas en el Excel (Asegúrate de tener Fecha, Fundo, Lote y Variedad).")
            return None, None

    with st.spinner("Procesando datos..."):
        df_final, lista_defectos = preparar_datos(st.session_state.df_crudo)
        
    if df_final is not None:
        st.session_state.df_final = df_final
        st.session_state.lista_defectos = lista_defectos
        st.success(f"✅ Datos procesados: {len(lista_defectos)} defectos encontrados, {df_final.shape[0]} registros diarios")
        
        # Mostrar defectos encontrados
        st.info(f"📋 Defectos detectados: {', '.join(lista_defectos[:10])}{'...' if len(lista_defectos) > 10 else ''}")

# ==========================================
# 3. FILTROS INTERACTIVOS
# ==========================================
if st.session_state.df_final is not None:
    st.header("🎯 3. Configuración del Reporte")
    
    df_plot = st.session_state.df_final.copy()
    
    # Filtro de fechas
    col1, col2 = st.columns(2)
    with col1:
        fecha_min = df_plot['Fecha'].min().date()
        fecha_max = df_plot['Fecha'].max().date()
        
        fecha_ini = st.date_input(
            "📅 Fecha de inicio:",
            value=fecha_min,
            min_value=fecha_min,
            max_value=fecha_max
        )
    
    with col2:
        fecha_fin = st.date_input(
            "📅 Fecha de fin:",
            value=fecha_max,
            min_value=fecha_min,
            max_value=fecha_max
        )
    
    # Aplicar filtro de fechas
    df_plot = df_plot[(df_plot['Fecha'].dt.date >= fecha_ini) & 
                      (df_plot['Fecha'].dt.date <= fecha_fin)]
    
    if df_plot.empty:
        st.warning("⚠️ No hay datos en el rango de fechas seleccionado")
    else:
        # Filtros en cascada
        st.subheader("Filtros en Cascada")
        
        # 1. Fundos
        fundos_unicos = sorted(df_plot['Fundo'].unique())
        fundos_sel = st.multiselect(
            "🏢 Fundos:",
            options=fundos_unicos,
            default=fundos_unicos,
            help="Selecciona uno o más fundos"
        )
        
        if fundos_sel:
            df_plot = df_plot[df_plot['Fundo'].isin(fundos_sel)]
            
            # 2. Variedades (depende de fundos)
            variedades_unicas = sorted(df_plot['Variedad'].unique())
            variedades_sel = st.multiselect(
                "🌾 Variedades:",
                options=variedades_unicas,
                default=variedades_unicas,
                help="Selecciona una o más variedades"
            )
            
            if variedades_sel:
                df_plot = df_plot[df_plot['Variedad'].isin(variedades_sel)]
                
                # 3. Lotes (depende de fundos y variedades)
                df_plot['Etiqueta_Lote'] = df_plot['Fundo'] + " - " + df_plot['Lote']
                lotes_unicos = sorted(df_plot['Etiqueta_Lote'].unique())
                lotes_sel = st.multiselect(
                    "📍 Lotes:",
                    options=lotes_unicos,
                    default=lotes_unicos,
                    help="Selecciona uno o más lotes"
                )
                
                if lotes_sel:
                    df_plot = df_plot[df_plot['Etiqueta_Lote'].isin(lotes_sel)]
                    
                    # 4. Defectos
                    defectos_unicos = sorted(st.session_state.lista_defectos)
                    defectos_sel = st.multiselect(
                        "⚠️ Defectos a incluir en el reporte:",
                        options=defectos_unicos,
                        default=defectos_unicos[:min(5, len(defectos_unicos))],
                        help="Selecciona los defectos para el reporte"
                    )
                    
                    if defectos_sel:
                        # Preparar datos para gráficos
                        df_plot['Fecha_Str'] = df_plot['Fecha'].dt.strftime('%d/%m/%Y')
                        
                        # Mostrar resumen
                        st.success(f"✅ Configuración lista: {len(fundos_sel)} fundos, {len(variedades_sel)} variedades, "
                                 f"{len(lotes_sel)} lotes, {len(defectos_sel)} defectos")
                        
                        # Vista previa de datos filtrados
                        with st.expander("🔍 Ver datos filtrados"):
                            st.dataframe(df_plot.head(20))
                        
                        # ==========================================
                        # 4. PREVISUALIZACIÓN DE GRÁFICOS
                        # ==========================================
                        st.header("👁️ 4. Previsualización de Gráficos")
                        
                        # Mostrar primer gráfico como preview
                        if st.checkbox("Mostrar preview del primer gráfico", value=True):
                            defecto_ejemplo = defectos_sel[0]
                            variedad_ejemplo = variedades_sel[0]
                            
                            data_var = df_plot[df_plot['Variedad'] == variedad_ejemplo]
                            
                            if not data_var.empty:
                                fig, ax = plt.subplots(figsize=(12, 6), dpi=150)
                                
                                fechas_unicas_dt = data_var.sort_values('Fecha')['Fecha'].unique()
                                fechas_str_ordenadas = [pd.to_datetime(f).strftime('%d/%m/%Y') for f in fechas_unicas_dt]
                                
                                lotes_presentes = data_var['Etiqueta_Lote'].unique()
                                colores = ['#D32F2F', '#1976D2', '#388E3C', '#FBC02D', '#8E24AA',
                                         '#F57C00', '#0097A7', '#689F38', '#C2185B', '#111111']
                                
                                for i, lote in enumerate(lotes_presentes):
                                    data_lote = data_var[data_var['Etiqueta_Lote'] == lote]
                                    valores_lote_alineados = []
                                    
                                    for f_str in fechas_str_ordenadas:
                                        valor_dia = data_lote.loc[data_lote['Fecha_Str'] == f_str, defecto_ejemplo]
                                        if not valor_dia.empty and pd.notna(valor_dia.iloc[0]):
                                            valores_lote_alineados.append(valor_dia.iloc[0])
                                        else:
                                            valores_lote_alineados.append(np.nan)
                                    
                                    color = colores[i % len(colores)]
                                    ax.plot(fechas_str_ordenadas, valores_lote_alineados,
                                           marker='o', label=lote, linewidth=2, color=color,
                                           markersize=8)
                                
                                ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.15), ncol=3)
                                ax.set_title(f"Preview: {defecto_ejemplo} - Variedad {variedad_ejemplo}", 
                                           fontsize=14, fontweight='bold')
                                plt.xticks(rotation=45, ha='right')
                                ax.get_yaxis().set_visible(False)
                                
                                st.pyplot(fig)
                                st.caption("Este es un preview. Los gráficos finales tendrán mayor resolución y etiquetas ajustadas.")
                        
                        # ==========================================
                        # 5. GENERAR POWERPOINT
                        # ==========================================
                        st.header("📊 5. Generar Presentación PowerPoint")
                        
                        st.info(f"Se generarán gráficos para {len(defectos_sel)} defectos × {len(variedades_sel)} variedades = "
                               f"{len(defectos_sel) * len(variedades_sel)} diapositivas aproximadamente")
                        
                        if st.button("🚀 GENERAR REPORTE POWERPOINT", type="primary"):
                            with st.spinner("Generando gráficos y presentación... Esto puede tomar unos minutos..."):
                                
                                # Crear presentación
                                prs = Presentation()
                                diapositivas_creadas = 0
                                
                                # Barra de progreso
                                progress_bar = st.progress(0)
                                total_graficos = len(defectos_sel) * len(variedades_sel)
                                grafico_actual = 0
                                
                                color_texto_principal = '#45605A'
                                color_borde_grafico = '#B0BEC5'
                                
                                colores_fuertes = [
                                    '#D32F2F', '#1976D2', '#388E3C', '#FBC02D', '#8E24AA',
                                    '#F57C00', '#0097A7', '#689F38', '#C2185B', '#111111',
                                    '#E64A19', '#512DA8', '#0288D1', '#AFB42B', '#5D4037',
                                    '#00796B', '#303F9F', '#C0CA33', '#455A64', '#D81B60'
                                ]
                                
                                for defecto in defectos_sel:
                                    if defecto not in df_plot.columns:
                                        continue
                                    
                                    for var in variedades_sel:
                                        data_var = df_plot[df_plot['Variedad'] == var]
                                        if data_var[defecto].isnull().all() or data_var.empty:
                                            grafico_actual += 1
                                            progress_bar.progress(grafico_actual / total_graficos)
                                            continue
                                        
                                        # Crear gráfico
                                        plt.figure(figsize=(14, 8), dpi=300)
                                        
                                        fechas_unicas_dt = data_var.sort_values('Fecha')['Fecha'].unique()
                                        fechas_str_ordenadas = [pd.to_datetime(f).strftime('%d/%m/%Y') for f in fechas_unicas_dt]
                                        
                                        lotes_presentes = data_var['Etiqueta_Lote'].unique()
                                        textos_a_ajustar = []
                                        
                                        for i, lote in enumerate(lotes_presentes):
                                            data_lote = data_var[data_var['Etiqueta_Lote'] == lote]
                                            color_asignado = colores_fuertes[i % len(colores_fuertes)]
                                            
                                            valores_lote_alineados = []
                                            for f_str in fechas_str_ordenadas:
                                                valor_dia = data_lote.loc[data_lote['Fecha_Str'] == f_str, defecto]
                                                if not valor_dia.empty and pd.notna(valor_dia.iloc[0]):
                                                    valores_lote_alineados.append(valor_dia.iloc[0])
                                                else:
                                                    valores_lote_alineados.append(np.nan)
                                            
                                            plt.plot(fechas_str_ordenadas, valores_lote_alineados,
                                                    marker='o', label=lote, linewidth=4, color=color_asignado,
                                                    markersize=10, markeredgecolor='white', markeredgewidth=1.5, zorder=5)
                                            
                                            for x_val, p in zip(fechas_str_ordenadas, valores_lote_alineados):
                                                if pd.notna(p):
                                                    val_etq = p * 100 if p < 1 else p
                                                    t = plt.text(x_val, p, f'{val_etq:.1f}%',
                                                               fontsize=14, fontweight='bold', color='white', zorder=10,
                                                               bbox=dict(facecolor=color_asignado, alpha=0.9,
                                                                        edgecolor='white', boxstyle='square,pad=0.3'))
                                                    textos_a_ajustar.append(t)
                                        
                                        if textos_a_ajustar:
                                            try:
                                                adjust_text(textos_a_ajustar,
                                                          expand_points=(1.5, 2.5),
                                                          expand_text=(2.0, 3.0),
                                                          force_text=(2.0, 4.0),
                                                          force_points=(0.5, 1.0),
                                                          arrowprops=dict(arrowstyle='-', color='#78909C', lw=1.5, alpha=0.9, zorder=2),
                                                          max_move=50)
                                            except:
                                                pass
                                        
                                        plt.legend(loc='upper center', bbox_to_anchor=(0.5, 1.15),
                                                 ncol=min(len(lotes_presentes), 5), frameon=False, fontsize=12)
                                        
                                        plt.xlabel(f"\nVariedad: {str(var).upper()}", fontsize=14, fontweight='bold', 
                                                 color=color_texto_principal)
                                        plt.xticks(rotation=45, ha='right', fontsize=12)
                                        
                                        plt.gca().get_yaxis().set_visible(False)
                                        
                                        for spine in plt.gca().spines.values():
                                            spine.set_visible(True)
                                            spine.set_color(color_borde_grafico)
                                            spine.set_linewidth(2)
                                        
                                        fecha_hoy = datetime.datetime.now().strftime("%d/%m/%Y")
                                        plt.figtext(0.99, 0.01, 'Complejo Agroindustrial Beta',
                                                  horizontalalignment='right', fontsize=10, color='gray', style='italic')
                                        
                                        plt.tight_layout()
                                        
                                        # Guardar gráfico en memoria
                                        image_stream = io.BytesIO()
                                        plt.savefig(image_stream, format='png', dpi=300, bbox_inches='tight', transparent=False)
                                        plt.close()
                                        
                                        # Agregar diapositiva
                                        slide_layout = prs.slide_layouts[5]
                                        slide = prs.slides.add_slide(slide_layout)
                                        
                                        title_shape = slide.shapes.title
                                        title_shape.text = f"Evaluación Diaria de Calidad: {defecto}"
                                        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(69, 96, 90)
                                        title_shape.text_frame.paragraphs[0].font.bold = True
                                        
                                        image_stream.seek(0)
                                        slide.shapes.add_picture(image_stream, Inches(0.25), Inches(1.8), width=Inches(9.5))
                                        diapositivas_creadas += 1
                                        
                                        # Actualizar progreso
                                        grafico_actual += 1
                                        progress_bar.progress(grafico_actual / total_graficos)
                                
                                if diapositivas_creadas > 0:
                                    # Guardar presentación
                                    pptx_buffer = io.BytesIO()
                                    prs.save(pptx_buffer)
                                    pptx_buffer.seek(0)
                                    
                                    # Completar barra de progreso
                                    progress_bar.progress(1.0)
                                    
                                    st.success(f"✅ ¡Reporte generado exitosamente! {diapositivas_creadas} diapositivas creadas.")
                                    
                                    # Botón de descarga
                                    st.download_button(
                                        label="📥 DESCARGAR PRESENTACIÓN POWERPOINT",
                                        data=pptx_buffer,
                                        file_name=f"Reporte_Gerencia_Calidad_Beta_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                        use_container_width=True
                                    )
                                else:
                                    st.warning("⚠️ No se generó ninguna diapositiva. Revisa si hay datos con los filtros aplicados.")
