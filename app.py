import streamlit as st
import pandas as pd
import plotly.express as px

# Configuracion de pagina sin emojis
st.set_page_config(
    page_title="Dashboard de Monitoreo de Servicios",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS
st.markdown("""
<style>
    .stApp {
        background-color: #f8f9fa;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    .main-title {
        color: #003366;
        text-align: center;
        padding: 20px 0;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    h2, h3 {
        color: #003366;
        font-weight: 600;
    }
    .kpi-container {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        border-left: 6px solid #FFD700;
        margin-bottom: 24px;
        text-align: center;
    }
    .kpi-container h3 {
        color: #003366;
        margin: 0;
        font-size: 1.1rem;
        font-weight: 600;
        text-transform: uppercase;
    }
    .kpi-container p {
        color: #000000;
        margin: 10px 0 0 0;
        font-size: 2.5rem;
        font-weight: 700;
    }
    hr {
        border-color: #003366;
        margin-top: 5px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

SISTEMAS_PERMITIDOS = [
    "ALFA", "APP MOVIL", "APPS EXTERNAS", "APPS INTRANET", "APPS PORTAL", 
    "GESTOR DOKUS", "HOMINIS", "IGA", "INSAP", "ITA", "NUEVA SEDE ELECTRONICA", 
    "PORTAL WEB E INTRANET", "SIAF", "SIGDEA PORTAL EMPLEADO", 
    "SIGDEA SEDE ELECTRONICA", "SIM", "SIRI", "STRATEGOS", "X-ROAD"
]

PALABRAS_CLAVE_ALARMA = ["rojo", "error", "caido", "fallo", "caído"]

def normalize_column_names(df):
    if df.empty:
        return df
    df.columns = df.columns.astype(str).str.strip().str.lower()
    col_mapping = {}
    for col in df.columns:
        if "fecha" in col: col_mapping[col] = "fecha"
        elif "horario" in col or "control" in col: col_mapping[col] = "horario_control"
        elif "exacta" in col or ("hora" in col and "control" not in col): col_mapping[col] = "hora_exacta"
        elif "aplicativ" in col or "sistema" in col: col_mapping[col] = "aplicativo"
        elif "inconveniente" in col or "problema" in col: col_mapping[col] = "inconvenientes"
        elif "comentario" in col or "admin" in col: col_mapping[col] = "comentario_admin"
    return df.rename(columns=col_mapping)

@st.cache_data
def process_data(file):
    try:
        # Leer TODAS las hojas, pd.read_excel devuelve un diccionario
        hojas = pd.read_excel(file, sheet_name=None)
        lista_dfs = []
        
        # Recorrer cada hoja
        for nombre_hoja, df in hojas.items():
            if df.empty:
                continue
                
            df = normalize_column_names(df)
            
            if "aplicativo" not in df.columns or "inconvenientes" not in df.columns:
                continue
                
            # Agregar columna Mes con el nombre de la hoja
            df["Mes"] = str(nombre_hoja).strip().upper()
            lista_dfs.append(df)
            
        if not lista_dfs:
            return pd.DataFrame()
            
        # Unir todos los DataFrames
        full_df = pd.concat(lista_dfs, ignore_index=True)
        
        # Convertir nombres de sistemas a MAYUSCULAS y eliminar espacios
        full_df["aplicativo"] = full_df["aplicativo"].fillna("").astype(str).str.strip().str.upper()
        
        # Validar que esten en la lista (corregir o filtrar)
        full_df = full_df[full_df["aplicativo"].isin(SISTEMAS_PERMITIDOS)]
        
        full_df["inconvenientes"] = full_df["inconvenientes"].fillna("").astype(str).str.strip().str.lower()
        
        if "comentario_admin" in full_df.columns:
            full_df["comentario_admin"] = full_df["comentario_admin"].fillna("").astype(str).str.strip()
        else:
            full_df["comentario_admin"] = "Sin comentario"
            
        if "hora_exacta" not in full_df.columns:
            full_df["hora_exacta"] = "No registrada"
            
        if "fecha" not in full_df.columns:
            full_df["fecha"] = "No registra"
        else:
            full_df['fecha'] = pd.to_datetime(full_df['fecha'], errors='coerce')
        
        # Detectar alarma y crear columna booleana
        pattern = "|".join(PALABRAS_CLAVE_ALARMA)
        full_df["Alarma"] = full_df["inconvenientes"].str.contains(pattern, case=False, na=False)
        
        return full_df
        
    except Exception as e:
        st.error(f"Error procesando los datos: {str(e)}")
        return pd.DataFrame()

# Interfaz Principal
st.markdown("<h1 class='main-title'>Dashboard de Monitoreo de Servicios</h1>", unsafe_allow_html=True)
st.markdown("Plataforma de analisis de disponibilidad de sistemas e incidentes reportados.")

uploaded_file = st.file_uploader("Seleccionar archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Procesando y agrupando todas las hojas de datos..."):
        df = process_data(uploaded_file)
        
    if df.empty:
        st.info("No se encontraron registros validos en el archivo cargado.")
    else:
        # Validaciones para asegurar lectura multinivel
        meses_unicos = sorted(df['Mes'].unique().tolist())
        sistemas_unicos = sorted(df['aplicativo'].unique().tolist())
        
        # Ocurren error si solo aparece un mes o sistema
        if len(meses_unicos) <= 1:
            st.error("Error: El archivo analizado cuenta con datos de un solo mes. Es obligatorio un archivo con multiples meses.")
            st.stop()
            
        if len(sistemas_unicos) <= 1:
            st.error("Error: El archivo analizado unicamente leyo un sistema. Es obligatorio procesar multiples sistemas.")
            st.stop()
            
        # Resumen de mejora extra
        st.success(f"Se analizaron {len(meses_unicos)} meses y {len(sistemas_unicos)} sistemas en total.")
        
        st.sidebar.markdown("<h2>Filtros</h2>", unsafe_allow_html=True)
        st.sidebar.markdown("<hr/>", unsafe_allow_html=True)
        
        # Todos los datos se visualizan por defecto porque default es igual a todos los disponibles
        filtro_mes = st.sidebar.multiselect("Seleccionar Mes", options=meses_unicos, default=meses_unicos)
        filtro_sistema = st.sidebar.multiselect("Seleccionar Sistema", options=sistemas_unicos, default=sistemas_unicos)
        
        df_filtered = df.copy()
        if filtro_mes:
            df_filtered = df_filtered[df_filtered['Mes'].isin(filtro_mes)]
        if filtro_sistema:
            df_filtered = df_filtered[df_filtered['aplicativo'].isin(filtro_sistema)]
            
        if df_filtered.empty:
            st.warning("Los filtros seleccionados no dejaron datos disponibles.")
        else:
            # Seleccionar solo alarmas para los calculos y graficos detallados de fallas
            df_alarmas_filtradas = df_filtered[df_filtered["Alarma"] == True]
            
            # Tarjetas de Indicadores Clave (KPIs) globales basados en ALARMAS
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"<div class='kpi-container'><h3>Total de Alarmas</h3><p>{len(df_alarmas_filtradas)}</p></div>", unsafe_allow_html=True)
                
            with col2:
                sistema_max = df_alarmas_filtradas['aplicativo'].value_counts().idxmax() if not df_alarmas_filtradas.empty else "Ninguno"
                st.markdown(f"<div class='kpi-container'><h3>Sistema con mas Alarmas</h3><p>{sistema_max}</p></div>", unsafe_allow_html=True)
                
            with col3:
                mes_max = df_alarmas_filtradas['Mes'].value_counts().idxmax() if not df_alarmas_filtradas.empty else "Ninguno"
                st.markdown(f"<div class='kpi-container'><h3>Mes con mas Alarmas</h3><p>{mes_max}</p></div>", unsafe_allow_html=True)
                
            st.markdown("<hr/>", unsafe_allow_html=True)
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.markdown("<h3>Alarmas por Sistema</h3>", unsafe_allow_html=True)
                if not df_alarmas_filtradas.empty:
                    conteo_sistemas = df_alarmas_filtradas['aplicativo'].value_counts().reset_index()
                    conteo_sistemas.columns = ['Sistema', 'Total Alarmas']
                    fig1 = px.bar(
                        conteo_sistemas, 
                        x='Sistema', 
                        y='Total Alarmas',
                        color_discrete_sequence=["#003366"]
                    )
                    fig1.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig1, use_container_width=True)
                else:
                    st.info("No hay registros de alarmas para graficar sistemas.")
                    
            with col_chart2:
                st.markdown("<h3>Alarmas por Mes</h3>", unsafe_allow_html=True)
                if not df_alarmas_filtradas.empty:
                    orden_meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
                    conteo_mes = df_alarmas_filtradas['Mes'].value_counts().reset_index()
                    conteo_mes.columns = ['Mes', 'Total Alarmas']
                    
                    categorias_presentes = [m for m in orden_meses if m in conteo_mes['Mes'].values]
                    if categorias_presentes:
                        conteo_mes['Mes'] = pd.Categorical(conteo_mes['Mes'], categories=categorias_presentes, ordered=True)
                        conteo_mes = conteo_mes.sort_values('Mes')
                        
                    fig2 = px.line(
                        conteo_mes, 
                        x='Mes', 
                        y='Total Alarmas', 
                        markers=True,
                        color_discrete_sequence=["#FFD700"]
                    )
                    fig2.update_traces(line=dict(width=4), marker=dict(size=12, color="#003366"))
                    fig2.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("No hay registros de alarmas para graficar meses.")
            
            # Tabla de Alarmas
            st.markdown("<hr/>", unsafe_allow_html=True)
            st.markdown("<h3>Registro Detallado de Incidentes y Alarmas</h3>", unsafe_allow_html=True)
            
            if not df_alarmas_filtradas.empty:
                df_mostrar = df_alarmas_filtradas.copy()
                if pd.api.types.is_datetime64_any_dtype(df_mostrar['fecha']):
                    df_mostrar['fecha'] = df_mostrar['fecha'].dt.strftime('%Y-%m-%d').fillna('No registra')
                else:
                    df_mostrar['fecha'] = df_mostrar['fecha'].astype(str)
                    
                columnas_mostrar = ['fecha', 'hora_exacta', 'Mes', 'aplicativo', 'inconvenientes', 'comentario_admin']
                columnas_finales = [c for c in columnas_mostrar if c in df_mostrar.columns]
                
                st.dataframe(
                    df_mostrar[columnas_finales], 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "fecha": "Fecha",
                        "hora_exacta": "Hora",
                        "Mes": "Mes Registro",
                        "aplicativo": "Sistema Afectado",
                        "inconvenientes": "Motivo",
                        "comentario_admin": "Comentario"
                    }
                )
            else:
                st.info("En el filtrado actual no hay registros clasificados como ALARMA.")

