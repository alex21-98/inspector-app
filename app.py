import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from adjustText import adjust_text
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
import io
import datetime
import math

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Reporte MP-Calidad", layout="wide", page_icon="📊")

st.title("📊🖥️ Generador de Reportes de MP-Calidad Gerencial")
st.markdown("Complejo Agroindustrial Beta - Sistema de Visualización de Defectos")

# ==========================================
# FUNCIÓN DE LIMPIEZA DE ESTADO (NUEVO)
# ==========================================
# Esta función borra el PPTX generado si el usuario cambia cualquier filtro
def limpiar_estado():
    if 'pptx_export' in st.session_state:
        st.session_state.pptx_export = None

# ==========================================
# 1. CARGA DE DATOS Y PREPARACIÓN
# ==========================================
MAPEO_COLUMNAS = {
    'Fecha': ['fecha', 'fec', 'date'], 'Mes': ['mes', 'month'],
    'Año': ['año', 'anio', 'year'], 'Semana': ['semana','SEMANA', 'sem', 'week', 'no sem'],
    'Fundo': ['Fundos', 'campo', 'fundo'], 'Lote': ['lote', 'lot', 'id lote'],
    'Variedad': ['variedad', 'var']
}

# Modificado para recibir "file_bytes" y evitar que el archivo se agote y pida resubir
@st.cache_data
def cargar_hoja_cruda(file_bytes, sheet, freq):
    """Carga y estandariza una hoja genérica retornando la data cruda para promedios simples"""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
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
# FUNCIONES VISUALES
# ==========================================
def calcular_ylim(valores, tol_num=None, es_porcentaje=False):
    vals = [float(v) for v in valores if pd.notna(v)]
    if tol_num is not None and pd.notna(tol_num): vals.append(float(tol_num))
    vals = [v for v in vals if not math.isnan(v)]
    
    if not vals: return 0, 1
    d_min, d_max = min(vals), max(vals)
    rango = d_max - d_min

    if rango == 0:
        margen = max(abs(d_max) * 0.1, 1)
        return (0 if es_porcentaje else d_min - margen), d_max + margen

    if es_porcentaje:
        y_min = 0
        y_max = d_max + (rango * 0.25)
    else:
        y_min = max(0, d_min - (rango * 0.5))
        y_max = d_max + (rango * 0.5)

    if tol_num is not None and y_max < tol_num * 1.15:
        y_max = tol_num * 1.15

    return y_min, y_max

def configurar_eje_y_sutil(ax, es_porcentaje):
    ax.get_yaxis().set_visible(True)
    ax.yaxis.set_major_locator(mtick.MaxNLocator(nbins=5, prune='lower'))
    ax.tick_params(axis='y', labelsize=10, colors='#9E9E9E', length=0, pad=5)
    
    if es_porcentaje:
        ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100 if es_porcentaje else 1, decimals=0))
        
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_color('#CFD8DC')
    ax.spines['bottom'].set_linewidth(1.5)
    ax.grid(axis='y', color='#F0F4F8', linewidth=1.2, alpha=1.0, zorder=0)

def reordenar_etiquetas_verticalmente(textos, ax, sep_frac=0.04):
    if not textos: return
    y_min, y_max = ax.get_ylim()
    rango_y = y_max - y_min
    if rango_y <= 0: return
    sep_min = rango_y * sep_frac
    
    grupos = {}
    for t in textos:
        x_val = round(t.get_position()[0], 3)
        if x_val not in grupos: grupos[x_val] = []
        grupos[x_val].append(t)
        
    for x_val, items in grupos.items():
        if len(items) <= 1: continue
        items = sorted(items, key=lambda x: x.get_position()[1])
        for _ in range(15):
            hubo_cambio = False
            for i in range(len(items) - 1):
                y1 = items[i].get_position()[1]
                y2 = items[i+1].get_position()[1]
                if (y2 - y1) < sep_min:
                    diff = sep_min - (y2 - y1)
                    items[i].set_position((x_val, y1 - diff / 2.0))
                    items[i+1].set_position((x_val, y2 + diff / 2.0))
                    hubo_cambio = True
            if not hubo_cambio: break
        
        margen = rango_y * 0.02
        for t in items:
            y_curr = t.get_position()[1]
            if y_curr > y_max - margen: t.set_position((x_val, y_max - margen))
            if y_curr < y_min + margen: t.set_position((x_val, y_min + margen))

def anclar_x_original(textos):
    for t in textos:
        if hasattr(t, '_data_x'):
            _, y_curr = t.get_position()
            t.set_position((t._data_x, y_curr))

# ==========================================
# LÓGICA DE GRÁFICOS REUTILIZABLE
# ==========================================
def generar_grafico_metricas(data_var, col_agrupacion, periodo_col, orden_col, defecto, 
                             titulo_personalizado, var, tolerancias_defectos, 
                             es_escala_decimal, sufijo="%", multiplicador=1.0):
                             
    cols_groupby = [col_agrupacion, periodo_col, orden_col]
    if 'Fundo' not in cols_groupby and 'Fundo' in data_var.columns:
        cols_groupby.append('Fundo')
        
    data_agg = data_var.groupby(cols_groupby, dropna=False)[defecto].mean().reset_index()
    data_agg = data_agg.dropna(subset=[periodo_col, orden_col]).sort_values(orden_col)
    
    periodos_ordenados = data_agg[periodo_col].unique().tolist()
    entidades_presentes = data_agg[col_agrupacion].unique()

    fig, ax = plt.subplots(figsize=(16.5, 8.5), dpi=130)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    fig.subplots_adjust(top=0.70, bottom=0.1, left=0.06, right=0.95)
    titulo_completo = f"{titulo_personalizado}: {defecto}".upper()
    fig.suptitle(f"{titulo_completo}\nVARIEDAD: {str(var).upper()}", 
                 fontsize=18, fontweight='bold', color='#1A237E', y=0.95, va='top')

    textos_a_ajustar = []
    colores_fuertes = [
        '#2874A6',  # 1. Azul acero oscuro (Jayanca I)
        '#239B56',  # 2. Verde bosque/esmeralda oscuro (Olmos)
        '#D68910',  # 3. Dorado/Ámbar oscuro (Olmos II)
        '#76448A',  # 4. Púrpura apagado/Ciruela (Pacora)
        '#B03A2E',  # 5. Rojo terracota
        '#117A65',  # 6. Teal/Verde azulado oscuro
        '#935116',  # 7. Marrón roble
        '#2C3E50',  # 8. Gris naval/Antracita oscuro
        '#1F618D',  # 9. Azul cobalto oscuro
        '#9A7D0A',  # 10. Ocre / Oliva corporativo
        '#AF601A',  # 11. Cobre
        '#5B2C6F',  # 12. Berenjena
        '#148F77',  # 13. Cian oscuro apagado
        '#78281F',  # 14. Rojo vino/Burdeos
        '#34495E'   # 15. Gris pizarra
    ]

    num_entidades = len(entidades_presentes)
    fs_label = 13 if num_entidades <= 6 else 11
    ultimo_periodo_grafico = periodos_ordenados[-1] if periodos_ordenados else None
    
    max_per_period = {per: data_agg[data_agg[periodo_col] == per][defecto].max() for per in periodos_ordenados}
    min_per_period = {per: data_agg[data_agg[periodo_col] == per][defecto].min() for per in periodos_ordenados}
    es_porcentaje = sufijo == "%"

    for i, entidad in enumerate(entidades_presentes):
        d_entidad = data_agg[data_agg[col_agrupacion] == entidad]
        color_asignado = colores_fuertes[i % len(colores_fuertes)]

        vals_raw = [d_entidad.loc[d_entidad[periodo_col] == per, defecto].iloc[0] if not d_entidad.loc[d_entidad[periodo_col] == per, defecto].empty else np.nan for per in periodos_ordenados]
        vals_plot = [(v * multiplicador if (pd.notna(v) and es_escala_decimal) else v) for v in vals_raw]

        ax.plot(periodos_ordenados, vals_plot, marker='o', label=entidad, 
                linewidth=3.0, color=color_asignado, markersize=8, markeredgecolor='white', 
                markeredgewidth=1.5, zorder=5, alpha=0.85)

        for x_val, v_raw, v_plot in zip(periodos_ordenados, vals_raw, vals_plot):
            if pd.isna(v_plot): continue
            
            debe_etiquetar = True if num_entidades <= 6 else (v_raw == max_per_period[x_val] or v_raw == min_per_period[x_val] or x_val == ultimo_periodo_grafico)
            
            if debe_etiquetar:
                val_etq = v_raw * multiplicador if es_escala_decimal else v_plot
                txt = f'{val_etq:.1f}{sufijo}'
                t = ax.text(x_val, v_plot, txt, fontsize=fs_label, fontweight='bold', color='white', 
                            ha='center', va='center', zorder=10, 
                            bbox=dict(facecolor=color_asignado, alpha=0.9, edgecolor='none', boxstyle='square,pad=0.6'))
                t._data_value = v_plot
                t._data_x = t.get_position()[0]
                textos_a_ajustar.append(t)

    tol_num = tolerancias_defectos.get(defecto)
    if tol_num is not None:
        valor_tol_grafico = tol_num / multiplicador if es_escala_decimal and not es_porcentaje else tol_num
        ax.axhline(y=valor_tol_grafico, color='#D32F2F', linestyle='--', linewidth=2.5, alpha=0.8, label=f'Límite ({tol_num}{sufijo})', zorder=3)

    valores_ylim = [v * multiplicador if (es_escala_decimal) else v for v in data_agg[defecto].dropna().tolist()]
    y_min_plot, y_max_plot = calcular_ylim(valores_ylim, tol_num, es_porcentaje=es_porcentaje)
    ax.set_ylim(y_min_plot, y_max_plot)
    configurar_eje_y_sutil(ax, es_porcentaje=es_porcentaje)

    ncol_dinamico = min(len(entidades_presentes) + (1 if tol_num is not None else 0), 5)
    ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=ncol_dinamico, frameon=False, fontsize=12)

    if textos_a_ajustar:
        try:
            adjust_text(textos_a_ajustar, ax=ax, only_move={'text': 'y'}, force_text=(0.0, 0.5), lim=50)
            reordenar_etiquetas_verticalmente(textos_a_ajustar, ax, sep_frac=0.05)
            anclar_x_original(textos_a_ajustar)
        except Exception:
            pass

    ax.set_xticks(range(len(periodos_ordenados)))
    rotacion = 0 if len(periodos_ordenados) <= 10 else 45
    ax.set_xticklabels(periodos_ordenados, rotation=rotacion, ha='center' if rotacion == 0 else 'right', fontsize=12)
    ax.tick_params(axis='x', pad=5)
    
    plt.figtext(0.99, 0.01, 'Complejo Agroindustrial Beta', horizontalalignment='right', fontsize=10, color='gray', style='italic')
    plt.tight_layout()
    
    return fig

# --- PANEL LATERAL ---
with st.sidebar:
    st.header("📁 1. Carga de Datos")
    # Limpiamos el estado si el usuario decide subir un archivo nuevo
    uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"], on_change=limpiar_estado)
    if uploaded_file is not None:
        # CONVERSIÓN A BYTES PARA EVITAR EL "AGOTAMIENTO" DEL ARCHIVO
        file_bytes = uploaded_file.getvalue() 
        excel_obj = pd.ExcelFile(io.BytesIO(file_bytes))
        hojas_disponibles = excel_obj.sheet_names
        
        # Enlazados a limpiar_estado
        hoja_final = st.selectbox("👉 Selecciona la hoja principal (Defectos):", hojas_disponibles, on_change=limpiar_estado)
        st.header("⚙️ 2. Periodicidad")
        opcion_freq = st.selectbox("Agrupar reporte por:", ["Semanal", "Diario", "Mensual", "Anual"], on_change=limpiar_estado)
        mapa_freq = {"Diario": 'D', "Semanal": 'W', "Mensual": 'M', "Anual": 'Y'}
        freq_elegida = mapa_freq[opcion_freq]

# --- MANEJO DE ESTADO PARA DESCARGA ---
if 'pptx_export' not in st.session_state:
    st.session_state.pptx_export = None
    st.session_state.pptx_filename = ""

# --- LÓGICA PRINCIPAL ---
if uploaded_file is not None:
    # Se pasa file_bytes en lugar de file_uploader
    df_raw_proc = cargar_hoja_cruda(file_bytes, hoja_final, freq_elegida)
    lista_defectos = [col for col in df_raw_proc.columns if '%' in str(col)]
    
    if not df_raw_proc.empty:
        st.divider()
        col_f1, col_f2 = st.columns(2)
        
        fecha_min = df_raw_proc['Fecha'].min().date()
        fecha_max = df_raw_proc['Fecha'].max().date()
        
        fecha_ini = col_f1.date_input("Fecha Inicio", value=fecha_min, min_value=fecha_min, max_value=fecha_max, on_change=limpiar_estado)
        fecha_fin = col_f2.date_input("Fecha Fin", value=fecha_max, min_value=fecha_min, max_value=fecha_max, on_change=limpiar_estado)
        
        # ==========================================
        # FILTROS EN CASCADA INTELIGENTES
        # ==========================================
        df_base = filtrar_dataframe(df_raw_proc, fecha_ini, fecha_fin, [], [], [])
        col1, col2, col3 = st.columns(3)
        
        # 1. Fundos
        fundos_disp = df_base['Fundo'].dropna().unique() if 'Fundo' in df_base.columns else []
        fundos_sel = col1.multiselect("Fundos", sorted(list(fundos_disp)), on_change=limpiar_estado)
        
        # 2. Variedades (Se actualiza automáticamente en base al fundo seleccionado)
        df_var = filtrar_dataframe(df_base, fecha_ini, fecha_fin, fundos_sel, [], [])
        vars_disp = df_var['Variedad'].dropna().unique() if 'Variedad' in df_var.columns else []
        variedades_sel = col2.multiselect("Variedades", sorted(list(vars_disp)), on_change=limpiar_estado)
        
        # 3. Lotes (Se actualiza automáticamente en base al fundo y variedad seleccionada)
        df_lote = filtrar_dataframe(df_base, fecha_ini, fecha_fin, fundos_sel, variedades_sel, [])
        if 'Fundo' in df_lote.columns and 'Lote' in df_lote.columns:
            df_lote['Etiqueta_Lote'] = df_lote['Fundo'] + " - " + df_lote['Lote']
            lotes_disp = df_lote['Etiqueta_Lote'].dropna().unique()
            lotes_sel = col3.multiselect("Lotes", sorted(list(lotes_disp)), on_change=limpiar_estado)
        else:
            lotes_sel = []
        
        defectos_sel = st.multiselect("Defectos a graficar", lista_defectos, on_change=limpiar_estado)

        st.divider()
        st.subheader("📈 3. Configuración Visual de Gráficos")
        col_visual1, col_visual2 = st.columns(2)
        
        with col_visual1:
            nivel_agrupacion = st.radio(
                "Selecciona cómo deseas agrupar las líneas en el gráfico:",
                ["Por Lote", "Por Fundo"],
                horizontal=True,
                on_change=limpiar_estado
            )
            col_agrupacion = 'Etiqueta_Lote' if nivel_agrupacion == "Por Lote" else 'Fundo'
            
        with col_visual2:
            titulo_personalizado = st.text_input(
                "Título base del gráfico:", 
                value="Evaluación de PT",
                help="Cambia este texto para modificar el inicio del título en gráficos y diapositivas",
                on_change=limpiar_estado
            )

        tolerancias_defectos = {}
        if defectos_sel:
            st.divider()
            st.subheader("📏 4. Configuración de Tolerancias (%)")
            cols = st.columns(min(len(defectos_sel), 4))
            for i, defecto in enumerate(defectos_sel):
                with cols[i % 4]:
                    tol_val = st.number_input(f"{defecto}", min_value=0.0, max_value=100.0, value=0.0, step=0.5, key=f"tol_{defecto}", on_change=limpiar_estado)
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
                
                df_plot = filtrar_dataframe(df_raw_proc, fecha_ini, fecha_fin, fundos_sel, variedades_sel, lotes_sel)
                
                if not df_plot.empty:
                    for defecto in defectos_sel:
                        for var in variedades_sel:
                            data_var = df_plot[df_plot['Variedad'] == var]
                            if data_var[defecto].isnull().all() or data_var.empty: continue
                            
                            max_val_test = data_var[defecto].max()
                            es_escala_decimal = max_val_test < 1.0 if pd.notna(max_val_test) else False
                            
                            fig = generar_grafico_metricas(
                                data_var, col_agrupacion, 'Periodo', 'Orden_Periodo', defecto, 
                                titulo_personalizado, var, tolerancias_defectos, 
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

                if diapositivas_creadas > 0:
                    pptx_stream = io.BytesIO()
                    prs.save(pptx_stream)
                    pptx_stream.seek(0)
                    
                    st.session_state.pptx_export = pptx_stream.getvalue()
                    st.session_state.pptx_filename = f"Reporte_Calidad_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    st.success(f"✅ Reporte generado exitosamente con {diapositivas_creadas} diapositivas.")
                else:
                    st.warning("⚠️ La combinación de filtros seleccionada no devolvió datos para graficar. Intenta con otra combinación.")

        # BOTÓN DE DESCARGA
        # Ahora, si el usuario cambia algo arriba, esta variable se limpia automáticamente y el botón desaparece
        # evitando que descargue el PPTX de la selección anterior
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
