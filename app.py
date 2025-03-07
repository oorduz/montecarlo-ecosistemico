import streamlit as st
import pandas as pd
import numpy as np
from scipy.special import inv_boxcox
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from data.cices_data import cices
import os
import base64
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="ECO-VPN: Simulación Monte Carlo para Servicios Ecosistémicos",
    page_icon="🌳",
    layout="wide"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E6091;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #1E6091;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .info-box {
        background-color: #E8F4F8;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
    .stat-box {
        background-color: #E8F4F8;
        padding: 1rem;
        border-radius: 5px;
        display: inline-block;
        width: 23%;
        margin: 0.5%;
        text-align: center;
        position: relative;
        padding-top: 2.5rem;  /* Espacio arriba para la etiqueta */
        min-height: 120px;  
    }
  .stat-label {
        font-size: 1rem;
        color: #333;
        font-weight: bold;
        position: absolute;
        top: 0.5rem;
        left: 0;
        right: 0;
        text-align: center;
        background-color: #E8F4F8;
        padding: 0.3rem 0;
        border-bottom: 1px solid #ccc;
    }
    .stat-value {
         font-size: 1.5rem;
        font-weight: bold;
        color: #1E6091;
        position: relative;
        z-index: 2;  /* Asegurar que esté por encima del fondo */
        text-shadow: 
            -1px -1px 0 white,  
            1px -1px 0 white,
            -1px 1px 0 white,
            1px 1px 0 white; 
    }
    .alert-success {
        background-color: #D4EDDA;
        color: #155724;
        padding: 0.75rem;
        border-radius: 0.25rem;
        margin-bottom: 1rem;
    }
    .alert-warning {
        background-color: #FFF3CD;
        color: #856404;
        padding: 0.75rem;
        border-radius: 0.25rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Título de la aplicación
st.markdown('<h1 class="main-header">ECO-VPN: Simulación Monte Carlo para Servicios Ecosistémicos</h1>', unsafe_allow_html=True)
st.markdown('<p>Evaluación económica de servicios ecosistémicos mediante simulación estocástica</p>', unsafe_allow_html=True)

# Crear DataFrame con los datos de CICES precargados
boxcox_gmm_cice_df = pd.DataFrame(cices)

# Función para agregar CICES_Clase
def prepare_cices_data(df):
    df['CICES_Clase'] = df['CICES'] + ': ' + df['Clase']
    return df

# Función para generar el enlace de descarga de Excel
def get_excel_download_link(df, filename="datos_simulacion.xlsx", text="Descargar Excel"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Simulación', index=False)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">{text}</a>'
    return href

# Función para realizar la simulación Monte Carlo
def run_monte_carlo_simulation(annos_cices, beneficios, cantidad_anios, boxcox_gmm_cice_df, num_datos=1000):
    # Datos de la tasa de descuento según Res. 1092 /2022 DNP
    if cantidad_anios <= 5:
        tasa_descuento = 0.095  # 9.5%
    elif cantidad_anios <= 25:
        tasa_descuento = 0.064  # 6.4%
    else:
        tasa_descuento = 0.035  # 3.5%
    
    num_datos_generados = int(num_datos * 1.10)
    
    # Diccionario para almacenar los datos de la simulación
    simulacion_data = {}
    
    # Simulación para cada "Cice"
    with st.spinner('Generando simulación para servicios ecosistémicos...'):
        for cice, annos in annos_cices.items():
            codigo = cice.split(':')[0]
            # Filtrar el DataFrame para el CICES actual
            cice_row = boxcox_gmm_cice_df[boxcox_gmm_cice_df['CICES'] == codigo].iloc[0]
            
            vr_lambda = cice_row['lambda']
            varianzas = cice_row['varianzas']
            desviacion_estandar = np.sqrt(varianzas)
            medias = cice_row['medias']
            pesos = cice_row['pesos']
            num_comp = cice_row['num_comp']
            
            datos_sim = []
            valores = np.array([])
            valores_bx = np.array([])
            
            for i in range(int(num_comp)):
                # Generamos valores aleatorios para la simulación según el peso
                size = int(np.round(pesos[i] * num_datos_generados))
                datos_sim_pesos = np.random.normal(loc=medias[i], scale=desviacion_estandar[i], size=size)
                datos_sim.append(datos_sim_pesos)
            
            valores_bx = np.concatenate(datos_sim)
            
            # Aplicar inversa de BoxCox
            valores = inv_boxcox(valores_bx, vr_lambda)
            
            # Filtrar valores válidos
            valores = valores[~np.isnan(valores)]
            valores = valores[valores > 0]
            
            valores = valores[:num_datos]  # Solo escogemos los valores deseados
            valores_bx = ((valores**vr_lambda) - 1) / vr_lambda
            
            # Iteramos sobre cada año
            for i in annos:
                columna = f'Cice: {codigo} (año {i})'
                columna_bx = f'Cice: {codigo} (año {i}) bx'
                if i == 1:
                    simulacion_data[columna] = valores
                    simulacion_data[columna_bx] = valores_bx
                else:
                    simulacion_data[columna] = [v / (1 + tasa_descuento) ** (i - 1) for v in valores]
                    simulacion_data[columna_bx] = [(v**vr_lambda) / vr_lambda for v in simulacion_data[columna]]
    
    # Simulación para cada "Beneficio"
    with st.spinner('Generando simulación para beneficios...'):
        for key, value in beneficios.items():
            valor, annos = beneficios[key]
            lista_valores_beneficios = [valor] * num_datos_generados
            lista_valores_beneficios = lista_valores_beneficios[:num_datos]
            
            x = 0
            for i in annos:
                anno = annos[x]
                columna = f'Beneficio: {key} (año {anno})'
                valores_anno = [v / (1 + tasa_descuento) ** (i - 1) for v in lista_valores_beneficios]
                simulacion_data[columna] = valores_anno
                x += 1
    
    df_simulacion = pd.DataFrame(simulacion_data)
    
    # Calcular VPN para cada año
    with st.spinner('Calculando VPN por año...'):
        for year in range(1, cantidad_anios + 1):
            column_name = f'vpn (año {year})'
            vpn_values = []
            
            for index, row in df_simulacion.iterrows():
                cice_sum = 0
                beneficio_sum = 0
                for column in df_simulacion.columns:
                    if f'(año {year})' in column:
                        if 'Cice:' in column and 'bx' not in column:
                            cice_sum += row[column]
                        elif 'Beneficio:' in column:
                            beneficio_sum += row[column]
                
                vpn_value = beneficio_sum - cice_sum
                vpn_values.append(vpn_value)
            
            df_simulacion[column_name] = vpn_values
    
    # Reorganizar las columnas
    columnas_ordenadas = []
    for i in range(1, cantidad_anios + 1):
        for columna in df_simulacion.columns:
            if f'(año {i})' in columna:
                columnas_ordenadas.append(columna)
    
    # Crear una copia explícita al reorganizar las columnas
    df_simulacion = df_simulacion[columnas_ordenadas].copy()
    
    # Asigna la suma de columnas que empiezan con "vpn" a una variable temporal
    vpn_suma = df_simulacion.filter(regex='^vpn').sum(axis=1)
    
    # Crear la columna 'VPN TOTAL'
    df_simulacion['VPN TOTAL'] = vpn_suma
    
    return df_simulacion, tasa_descuento

# Inicializar variables de sesión
if 'cices_seleccionados' not in st.session_state:
    st.session_state.cices_seleccionados = []
if 'annos_cices' not in st.session_state:
    st.session_state.annos_cices = {}
if 'beneficios' not in st.session_state:
    st.session_state.beneficios = {}
if 'cantidad_anios' not in st.session_state:
    st.session_state.cantidad_anios = 10
if 'num_simulaciones' not in st.session_state:
    st.session_state.num_simulaciones = 1000
if 'resultado_simulacion' not in st.session_state:
    st.session_state.resultado_simulacion = None
if 'tasa_descuento' not in st.session_state:
    st.session_state.tasa_descuento = None

# Crear pestañas
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1️⃣ Configuración", 
    "2️⃣ Servicios Ecosistémicos", 
    "3️⃣ Beneficios", 
    "4️⃣ Simulación", 
    "5️⃣ Resultados"
])

# TAB 1: Configuración
with tab1:
    st.markdown('<h2 class="sub-header">Configuración del Proyecto</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.text_input("Nombre del Proyecto", value="Evaluación Económica de Servicios Ecosistémicos", key="nombre_proyecto")
        
    with col2:
        st.session_state.cantidad_anios = st.slider(
            "Duración del Proyecto (años)", 
            min_value=1, 
            max_value=50, 
            value=st.session_state.cantidad_anios,
            step=1
        )
        
    st.session_state.num_simulaciones = st.select_slider(
        "Número de Simulaciones",
        options=[100, 500, 1000, 5000, 10000],
        value=st.session_state.num_simulaciones
    )
    
    st.markdown('<div class="info-box">La tasa de descuento se calcula automáticamente según la Resolución 1092/2022 del DNP:<br>'
                '• Para proyectos ≤ 5 años: 9.5%<br>'
                '• Para proyectos entre 6 y 25 años: 6.4%<br>'
                '• Para proyectos > 25 años: 3.5%</div>', unsafe_allow_html=True)
    
    if st.button("Continuar con Servicios Ecosistémicos →", type="primary"):
        st.markdown(f'<div class="alert-success">Configuración guardada correctamente. Duración del proyecto: {st.session_state.cantidad_anios} años</div>', unsafe_allow_html=True)

# TAB 2: Servicios Ecosistémicos
with tab2:
    st.markdown('<h2 class="sub-header">Selección de Servicios Ecosistémicos</h2>', unsafe_allow_html=True)
    
    # Preparar los datos de CICES
    boxcox_gmm_cice_df = prepare_cices_data(boxcox_gmm_cice_df)
    
    # Lista de opciones de CICES
    unique_cices_clase = boxcox_gmm_cice_df['CICES_Clase'].unique().tolist()
    unique_cices_clase.sort()
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        cice_seleccionado = st.selectbox(
            "Seleccionar Servicio Ecosistémico", 
            unique_cices_clase,
            index=0
        )
    
    with col2:
        if st.button("Agregar Servicio", type="primary"):
            if cice_seleccionado not in st.session_state.cices_seleccionados:
                st.session_state.cices_seleccionados.append(cice_seleccionado)
                st.session_state.annos_cices[cice_seleccionado] = []
    
    # Mostrar lista de CICES seleccionados
    if st.session_state.cices_seleccionados:
        st.markdown('<h3>Servicios Ecosistémicos Seleccionados</h3>', unsafe_allow_html=True)
        
        for i, cice in enumerate(st.session_state.cices_seleccionados):
            col1, col2, col3 = st.columns([3, 2, 1])
            
            with col1:
                st.write(f"{i+1}. {cice}")
            
            with col2:
                # Entrada para años
                annos_input = st.text_input(
                    f"Años para {cice.split(':')[0]}",
                    placeholder="Ej: 1,3,5,7 o 2-7",
                    key=f"annos_{i}"
                )
                
                if annos_input:
                    try:
                        if "-" in annos_input:
                            inicio, fin = map(int, annos_input.split("-"))
                            annos = list(range(inicio, fin + 1))
                        else:
                            annos = [int(x) for x in annos_input.replace(" ", "").split(",")]
                        
                        if any(year > st.session_state.cantidad_anios for year in annos):
                            st.error(f"Error: Hay años mayores a la duración del proyecto ({st.session_state.cantidad_anios} años)")
                        elif any(year < 1 for year in annos):
                            st.error("Error: Los años deben ser mayores o iguales a 1")
                        else:
                            st.session_state.annos_cices[cice] = annos
                    except ValueError:
                        st.error("Formato de años incorrecto. Use ej: 1,3,5,7 o 2-7")
            
            with col3:
                if st.button("Eliminar", key=f"delete_{i}"):
                    st.session_state.cices_seleccionados.remove(cice)
                    if cice in st.session_state.annos_cices:
                        del st.session_state.annos_cices[cice]
                    st.rerun()
    
    if st.button("Continuar con Beneficios →", type="primary"):
        # Validar que todos los CICES tengan años asignados
        if all(len(st.session_state.annos_cices.get(cice, [])) > 0 for cice in st.session_state.cices_seleccionados):
            st.markdown('<div class="alert-success">Servicios Ecosistémicos guardados correctamente.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="alert-warning">Error: Todos los servicios ecosistémicos deben tener años asignados.</div>', unsafe_allow_html=True)

# TAB 3: Beneficios
with tab3:
    st.markdown('<h2 class="sub-header">Definición de Beneficios</h2>', unsafe_allow_html=True)
    
    with st.form("beneficio_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            nombre_beneficio = st.text_input("Nombre del Beneficio", placeholder="Ej: Beneficio Turístico")
        
        with col2:
            valor_beneficio = st.number_input("Valor del Beneficio (COP)", min_value=0.0, step=1000000.0, format="%f")
        
        with col3:
            annos_beneficio = st.text_input(
                "Años en que se recibe el beneficio",
                placeholder="Ej: 1,3,5,7 o 2-7"
            )
        
        submitted = st.form_submit_button("Agregar Beneficio", type="primary")
        
        if submitted:
            if nombre_beneficio and valor_beneficio > 0 and annos_beneficio:
                try:
                    if "-" in annos_beneficio:
                        inicio, fin = map(int, annos_beneficio.split("-"))
                        annos = list(range(inicio, fin + 1))
                    else:
                        annos = [int(x) for x in annos_beneficio.replace(" ", "").split(",")]
                    
                    if any(year > st.session_state.cantidad_anios for year in annos):
                        st.error(f"Error: Hay años mayores a la duración del proyecto ({st.session_state.cantidad_anios} años)")
                    elif any(year < 1 for year in annos):
                        st.error("Error: Los años deben ser mayores o iguales a 1")
                    else:
                        st.session_state.beneficios[nombre_beneficio] = [valor_beneficio, annos]
                        st.success(f"Beneficio '{nombre_beneficio}' agregado correctamente.")
                except ValueError:
                    st.error("Formato de años incorrecto. Use ej: 1,3,5,7 o 2-7")
    
    # Mostrar lista de beneficios
    if st.session_state.beneficios:
        st.markdown('<h3>Beneficios Definidos</h3>', unsafe_allow_html=True)
        
        beneficios_df = pd.DataFrame(
            [[key, value[0], str(value[1])] for key, value in st.session_state.beneficios.items()],
            columns=["Nombre", "Valor (COP)", "Años"]
        )
        
        # Format values as currency
        beneficios_df["Valor (COP)"] = beneficios_df["Valor (COP)"].apply(lambda x: f"${x:,.2f}")
        
        st.dataframe(beneficios_df, use_container_width=True)
        
        if st.button("Eliminar Todos los Beneficios"):
            st.session_state.beneficios = {}
            st.rerun()
    
    if st.button("Continuar con Simulación →", type="primary"):
        if st.session_state.beneficios:
            st.markdown('<div class="alert-success">Beneficios guardados correctamente.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="alert-warning">Advertencia: No se han definido beneficios.</div>', unsafe_allow_html=True)

# TAB 4: Simulación
with tab4:
    st.markdown('<h2 class="sub-header">Ejecutar Simulación Monte Carlo</h2>', unsafe_allow_html=True)
    
    # Mostrar resumen de configuración
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.markdown('<h3>Resumen de configuración</h3>', unsafe_allow_html=True)
    
    # Calcular tasa de descuento
    if st.session_state.cantidad_anios <= 5:
        tasa_descuento = 0.095  # 9.5%
    elif st.session_state.cantidad_anios <= 25:
        tasa_descuento = 0.064  # 6.4%
    else:
        tasa_descuento = 0.035  # 3.5%
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"**Duración del proyecto:** {st.session_state.cantidad_anios} años")
        st.markdown(f"**Tasa de descuento:** {tasa_descuento * 100:.1f}%")
        st.markdown(f"**Número de simulaciones:** {st.session_state.num_simulaciones:,}")
    
    with col2:
        st.markdown(f"**Servicios ecosistémicos:** {len(st.session_state.cices_seleccionados)} seleccionados")
        st.markdown(f"**Beneficios:** {len(st.session_state.beneficios)} definidos")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Botón para ejecutar la simulación
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⚡ EJECUTAR SIMULACIÓN", type="primary", use_container_width=True):
            if not st.session_state.cices_seleccionados:
                st.error("Error: Debe seleccionar al menos un servicio ecosistémico.")
            elif not all(len(st.session_state.annos_cices.get(cice, [])) > 0 for cice in st.session_state.cices_seleccionados):
                st.error("Error: Todos los servicios ecosistémicos deben tener años asignados.")
            elif not st.session_state.beneficios:
                st.warning("Advertencia: No ha definido beneficios. El VPN podría ser negativo.")
                
                # Confirmar si desea continuar sin beneficios
                if st.button("Continuar sin beneficios"):
                    st.session_state.resultado_simulacion, st.session_state.tasa_descuento = run_monte_carlo_simulation(
                        st.session_state.annos_cices,
                        st.session_state.beneficios,
                        st.session_state.cantidad_anios,
                        boxcox_gmm_cice_df,
                        st.session_state.num_simulaciones
                    )
                    st.success("¡Simulación completada con éxito!")
                    st.rerun()
            else:
                st.session_state.resultado_simulacion, st.session_state.tasa_descuento = run_monte_carlo_simulation(
                    st.session_state.annos_cices,
                    st.session_state.beneficios,
                    st.session_state.cantidad_anios,
                    boxcox_gmm_cice_df,
                    st.session_state.num_simulaciones
                )
                st.success("¡Simulación completada con éxito!")
                st.rerun()
    
    if st.session_state.resultado_simulacion is not None:
        st.button("Ver Resultados →", type="primary", on_click=lambda: None)

# TAB 5: Resultados
with tab5:
    if st.session_state.resultado_simulacion is not None:
        st.markdown('<h2 class="sub-header">Resultados de la Simulación</h2>', unsafe_allow_html=True)
        
        df_simulacion = st.session_state.resultado_simulacion
        
        # Calcular estadísticas para VPN TOTAL
        vpn_total_stats = df_simulacion['VPN TOTAL'].agg(['mean', 'std', 'median'])
        mode_vpn = stats.mode(df_simulacion['VPN TOTAL'])[0]
        skewness_vpn = stats.skew(df_simulacion['VPN TOTAL'])
        kurtosis_vpn = stats.kurtosis(df_simulacion['VPN TOTAL'])
        range_vpn = df_simulacion['VPN TOTAL'].max() - df_simulacion['VPN TOTAL'].min()
        min_vpn = df_simulacion['VPN TOTAL'].min()
        max_vpn = df_simulacion['VPN TOTAL'].max()
        confidence_level = 0.95
        confidence_interval = stats.t.interval(confidence_level, len(df_simulacion['VPN TOTAL']) - 1, 
                                              loc=np.mean(df_simulacion['VPN TOTAL']), 
                                              scale=stats.sem(df_simulacion['VPN TOTAL']))
        
        # Probabilidad de VPN positivo
        prob_vpn_positive = (df_simulacion['VPN TOTAL'] > 0).mean() * 100
        
        # Mostrar métricas principales
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown('<h3>Estadísticas del VPN Total</h3>', unsafe_allow_html=True)
            
            col_a, col_b, col_c, col_d = st.columns(4)
            with col_a:
                st.markdown(f'<div class="stat-box"><div class="stat-label">Media</div><div class="stat-value">${vpn_total_stats["mean"]:,.2f}</div></div>', unsafe_allow_html=True)
            with col_b:
                st.markdown(f'<div class="stat-box"><div class="stat-label">Mediana</div><div class="stat-value">${vpn_total_stats["median"]:,.2f}</div></div>', unsafe_allow_html=True)
            with col_c:
                st.markdown(f'<div class="stat-box"><div class="stat-label">Desv. Estándar</div><div class="stat-value">${vpn_total_stats["std"]:,.2f}</div></div>', unsafe_allow_html=True)
            with col_d:
                st.markdown(f'<div class="stat-box"><div class="stat-label">Rango</div><div class="stat-value">${min_vpn:,.0f} - ${max_vpn:,.0f}</div></div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<h3>Probabilidad de éxito</h3>', unsafe_allow_html=True)
            
            # Crear gráfico circular para mostrar probabilidad
            fig = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = prob_vpn_positive,
                domain = {'x': [0, 1], 'y': [0, 1]},
                title = {'text': "VPN Positivo"},
                gauge = {
                    'axis': {'range': [0, 100]},
                    'bar': {'color': "#2E7D32" if prob_vpn_positive >= 70 else "#FFC107"},
                    'steps': [
                        {'range': [0, 50], 'color': "#FFCDD2"},
                        {'range': [50, 70], 'color': "#FFE082"},
                        {'range': [70, 100], 'color': "#C8E6C9"}
                    ],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 50
                    }
                }
            ))
            
            fig.update_layout(height=250, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig, use_container_width=True)
        
        # Histograma y distribución del VPN Total
        st.markdown('<h3>Distribución del VPN Total</h3>', unsafe_allow_html=True)
        
        # Crear histograma con Plotly
        fig = px.histogram(
            df_simulacion, 
            x='VPN TOTAL',
            nbins=30,
            marginal="box",
            title="Histograma y Diagrama de Caja del VPN Total",
            labels={"VPN TOTAL": "VPN Total (COP)"},
            color_discrete_sequence=['#1E6091']
        )
        
       # Añadir línea vertical para el VPN = 0
        fig.add_vline(x=0, line_dash="dash", line_color="red")
        
        # Añadir línea vertical para la media
        fig.add_vline(x=vpn_total_stats["mean"], line_dash="dash", line_color="green")
        
        # Personalizar diseño
        fig.update_layout(
            height=500,
            bargap=0.1,
            xaxis_title="VPN Total (COP)",
            yaxis_title="Frecuencia",
            showlegend=False
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # VPN por Año
        st.markdown('<h3>VPN por Año</h3>', unsafe_allow_html=True)
        
        # Extraer datos de VPN por año
        vpn_by_year = pd.DataFrame({
            'Año': range(1, st.session_state.cantidad_anios + 1),
            'VPN': [df_simulacion[f'vpn (año {year})'].mean() for year in range(1, st.session_state.cantidad_anios + 1)]
        })
        
        # Crear gráfico de línea para VPN por año
        fig = px.line(
            vpn_by_year,
            x='Año',
            y='VPN',
            markers=True,
            title="Valor Presente Neto promedio por año",
            labels={"VPN": "VPN (COP)"},
            color_discrete_sequence=['#1E6091']
        )
        
        # Añadir línea horizontal para VPN = 0
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        
        # Personalizar diseño
        fig.update_layout(
            height=400,
            xaxis=dict(dtick=1),
            yaxis_tickformat="$,.0f"
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Estadísticas detalladas
        with st.expander("Ver estadísticas detalladas"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### Estadísticas descriptivas")
                stats_text = f"""
                - **Media:** ${vpn_total_stats['mean']:,.2f}
                - **Error típico:** ${stats.sem(df_simulacion['VPN TOTAL']):,.2f}
                - **Mediana:** ${vpn_total_stats['median']:,.2f}
                - **Moda:** ${mode_vpn:,.2f}
                - **Desviación estándar:** ${vpn_total_stats['std']:,.2f}
                - **Varianza:** ${np.var(df_simulacion['VPN TOTAL'], ddof=1):,.2f}
                """
                st.markdown(stats_text)
            
            with col2:
                st.markdown("### Estadísticas de forma y rango")
                more_stats = f"""
                - **Curtosis:** {kurtosis_vpn:.4f}
                - **Coeficiente de asimetría:** {skewness_vpn:.4f}
                - **Rango:** ${range_vpn:,.2f}
                - **Mínimo:** ${min_vpn:,.2f}
                - **Máximo:** ${max_vpn:,.2f}
                - **Intervalo de confianza (95%):** ${confidence_interval[0]:,.2f} a ${confidence_interval[1]:,.2f}
                """
                st.markdown(more_stats)
        
        # Descargar resultados
        st.markdown('<h3>Descargar Resultados</h3>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Crear un DataFrame resumen para exportar
            df_resumen = pd.DataFrame({
                'Métrica': ['Media', 'Mediana', 'Desviación Estándar', 'Mínimo', 'Máximo', 
                           'Probabilidad VPN Positivo (%)', 'Curtosis', 'Asimetría',
                           'Límite Inferior IC 95%', 'Límite Superior IC 95%'],
                'Valor': [vpn_total_stats['mean'], vpn_total_stats['median'], vpn_total_stats['std'],
                         min_vpn, max_vpn, prob_vpn_positive, kurtosis_vpn, skewness_vpn,
                         confidence_interval[0], confidence_interval[1]]
            })
            
            # Crear buffer para el Excel
            output_resumen = BytesIO()
            with pd.ExcelWriter(output_resumen, engine='openpyxl') as writer:
                df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
                vpn_by_year.to_excel(writer, sheet_name='VPN_por_año', index=False)
            
            resumen_data = output_resumen.getvalue()
            b64_resumen = base64.b64encode(resumen_data).decode()
            
            st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_resumen}" download="resumen_simulacion.xlsx" class="btn-download">📊 Descargar Resumen</a>', unsafe_allow_html=True)
        
        with col2:
            # Crear enlace para descargar los datos completos
            excel_link = get_excel_download_link(df_simulacion, "simulacion_completa.xlsx", "📈 Descargar Datos Completos")
            st.markdown(excel_link, unsafe_allow_html=True)
    
    else:
        st.info("No hay resultados disponibles. Por favor, ejecute la simulación en la pestaña anterior.")