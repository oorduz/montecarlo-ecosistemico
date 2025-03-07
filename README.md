# ECO-VPN: Simulación Monte Carlo para Servicios Ecosistémicos

Esta aplicación web implementa un modelo de simulación Monte Carlo para la evaluación económica de servicios ecosistémicos, calculando el Valor Presente Neto (VPN) en diferentes escenarios.

## Características

- Configuración personalizada de proyectos con duración variable
- Selección de servicios ecosistémicos desde una base de datos CICES precargada
- Definición de beneficios económicos asociados al proyecto
- Simulación Monte Carlo con hasta 10,000 iteraciones
- Visualización interactiva de resultados con histogramas, gráficos de línea y estadísticas detalladas
- Exportación de resultados en formato Excel

## Requisitos

- Python 3.7 o superior
- Streamlit
- Pandas
- NumPy
- SciPy
- Matplotlib
- Seaborn
- Plotly

## Instalación

1. Clonar el repositorio:
```bash
git clone https://github.com/oorduz/montecarlo-ecosistemico.git
cd montecarlo-ecosistemico
```

2. Instalar las dependencias:
```bash
pip install -r requirements.txt
```

3. Ejecutar la aplicación:
```bash
streamlit run app.py
```

## Uso

1. **Configuración**: Define el nombre del proyecto, la duración y el número de simulaciones.
2. **Servicios Ecosistémicos**: Selecciona los servicios ecosistémicos de la base de datos CICES y especifica los años en que ocurren.
3. **Beneficios**: Agrega los beneficios económicos del proyecto y los años en que se obtienen.
4. **Simulación**: Ejecuta la simulación Monte Carlo con los parámetros definidos.
5. **Resultados**: Analiza los resultados a través de visualizaciones y estadísticas, y descarga los datos para un análisis más detallado.

## Metodología

La aplicación implementa el método de simulación Monte Carlo para estimar la distribución probabilística del VPN del proyecto, considerando:

- Modelos de mezcla gaussiana (GMM) para representar la variabilidad de los servicios ecosistémicos
- Transformaciones Box-Cox para manejar distribuciones sesgadas
- Tasas de descuento según la Resolución 1092/2022 del DNP
- Cálculo del VPN como la diferencia entre beneficios y costos por servicios ecosistémicos

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue para discutir los cambios que te gustaría realizar.
